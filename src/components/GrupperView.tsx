import { Fragment, useEffect, useMemo, useState } from 'react';
import type { StandardField, StudentAssignmentChange } from '../utils/excelUtils';
import { removeSubjectAssignmentsForStudents } from '../utils/excelUtils';
import {
  BLOKK_LABELS,
  DEFAULT_MAX_PER_SUBJECT,
  getBlokkNumber,
  getResolvedGroupsByTarget,
  getSettingsForSubject,
  makeGroup,
  shouldShowGroup,
  type BlokkLabel,
  type SubjectGroup,
  type SubjectSettingsByName,
  type StudentIdsByBlokk,
} from '../utils/subjectGroups';
import type { ClassBlockRestrictions } from '../utils/progressiveHybridBalance';
import styles from './GrupperView.module.css';

interface GrupperViewProps {
  data: StandardField[];
  blokkCount: number;
  subjectOptions: string[];
  subjectSettingsByName: SubjectSettingsByName;
  classBlockRestrictions: ClassBlockRestrictions;
  changeLog: StudentAssignmentChange[];
  onSaveSubjectSettingsByName: (values: SubjectSettingsByName) => void;
  onStudentDataUpdate: (updatedData: StandardField[], changes: StudentAssignmentChange[]) => void;
  onOpenStudentCard?: (studentId: string) => void;
}

type BlokkField = `blokk${1 | 2 | 3 | 4 | 5 | 6 | 7 | 8}`;
type SortKey = 'blokk' | 'subject' | 'group' | 'count' | 'max';
type SortDirection = 'asc' | 'desc';
type StudentMovement = 'current' | 'moved-in' | 'moved-out';

interface GroupRow {
  key: string;
  subject: string;
  blokk: BlokkLabel;
  blokkNumber: number;
  groupId: string;
  groupLabel: string;
  count: number;
  max: number;
  studentIds: string[];
  recentInCount: number;
  recentOutCount: number;
}

interface StudentLookupEntry {
  student: StandardField;
  index: number;
  studentId: string;
}

interface VisibleStudentRow {
  studentId: string;
  navn: string;
  klasse: string;
  movement: StudentMovement;
  statusLabel: string;
  changedAt: string;
  reason: string;
}

interface SubjectBlockChoice {
  blokkNumber: number;
  studentCount: number;
}

const parseSubjects = (value: string | null): string[] => {
  if (!value) {
    return [];
  }

  return value
    .split(/[,;]/)
    .map((subject) => subject.trim())
    .filter((subject) => subject.length > 0);
};

const isSameSubject = (left: string, right: string): boolean => {
  return left.localeCompare(right, 'nb', { sensitivity: 'base' }) === 0;
};

const compareText = (left: string, right: string): number => {
  return left.localeCompare(right, 'nb', { sensitivity: 'base', numeric: true });
};

const normalizeSubjectKey = (value: string): string => {
  return value.trim().toLocaleLowerCase('nb');
};

const inferClassLevels = (student: StandardField): string[] => {
  if (student.fjerdearsElev) {
    return ['VG2', 'VG3'];
  }

  const classGroup = student.klasse;
  const normalized = (classGroup || '').trim().toUpperCase();
  if (!normalized) {
    return [];
  }

  const match = normalized.match(/^(\d)/);
  if (!match) {
    return ['VG1', 'VG2', 'VG3'].includes(normalized) ? [normalized] : [];
  }

  const year = Number.parseInt(match[1], 10);
  if (year === 1) {
    return ['VG1'];
  }
  if (year === 2) {
    return ['VG2'];
  }
  if (year >= 3) {
    return ['VG3'];
  }

  return [];
};

const isBlockAllowedForStudent = (
  student: StandardField,
  blokkNumber: number,
  restrictions: ClassBlockRestrictions
): boolean => {
  const classLevels = inferClassLevels(student);
  if (classLevels.length === 0) {
    return true;
  }

  return classLevels.some((classLevel) => restrictions[classLevel]?.[blokkNumber as 1 | 2 | 3 | 4] ?? true);
};

const getStudentId = (student: StandardField, index: number): string => {
  return student.studentId || `${student.navn || 'ukjent'}:${student.klasse || 'ukjent'}:${index}`;
};

const getBlokkKey = (blokkNumber: number): BlokkField => {
  return `blokk${blokkNumber}` as BlokkField;
};

const formatSubjects = (subjects: string[]): string | null => {
  return subjects.length > 0 ? subjects.join(', ') : null;
};

const getStudentSubjectsByBlokkLines = (student: StandardField, visibleBlokkCount: number): string[] => {
  const assignments: string[] = [];

  for (let blokkNumber = 1; blokkNumber <= visibleBlokkCount; blokkNumber += 1) {
    const subjects = parseSubjects(student[getBlokkKey(blokkNumber)] as string | null);
    if (subjects.length === 0) {
      continue;
    }

    assignments.push(`Blokk ${blokkNumber}: ${subjects.join(', ')}`);
  }

  return assignments.length > 0 ? assignments : ['Ingen fag tildelt'];
};

const buildChangeKey = (studentId: string, subject: string): string => {
  return `${studentId}::${subject.trim().toLocaleLowerCase('nb')}`;
};

const getMovementStatus = (
  change: StudentAssignmentChange | undefined,
  blokkNumber: number,
  currentlyVisible: boolean
): { movement: StudentMovement; statusLabel: string } => {
  if (!change) {
    return {
      movement: currentlyVisible ? 'current' : 'moved-out',
      statusLabel: currentlyVisible ? '' : 'Ut',
    };
  }

  if (currentlyVisible) {
    if (change.toBlokk === blokkNumber) {
      if (change.fromBlokk > 0 && change.fromBlokk !== blokkNumber) {
        return { movement: 'moved-in', statusLabel: `Inn fra Blokk ${change.fromBlokk}` };
      }

      return {
        movement: 'moved-in',
        statusLabel: change.fromBlokk === 0 ? 'Lagt til' : 'Oppdatert',
      };
    }

    return { movement: 'current', statusLabel: '' };
  }

  if (change.toBlokk > 0 && change.toBlokk !== blokkNumber) {
    return { movement: 'moved-out', statusLabel: `Ut til Blokk ${change.toBlokk}` };
  }

  return { movement: 'moved-out', statusLabel: 'Fjernet' };
};

const sortStudents = (rows: VisibleStudentRow[]): VisibleStudentRow[] => {
  return [...rows].sort((left, right) => {
    const nameCompare = compareText(left.navn, right.navn);
    if (nameCompare !== 0) {
      return nameCompare;
    }

    return compareText(left.klasse, right.klasse);
  });
};

export const GrupperView = ({
  data,
  blokkCount,
  subjectOptions,
  subjectSettingsByName,
  classBlockRestrictions,
  changeLog,
  onSaveSubjectSettingsByName,
  onStudentDataUpdate,
  onOpenStudentCard,
}: GrupperViewProps) => {
  const visibleBlokkCount = Math.min(blokkCount, 4);
  const [sortKey, setSortKey] = useState<SortKey>('blokk');
  const [sortDirection, setSortDirection] = useState<SortDirection>('asc');
  const [expandedGroupKey, setExpandedGroupKey] = useState<string | null>(null);
  const [selectedStudentIds, setSelectedStudentIds] = useState<string[]>([]);
  const [bulkTargetSubject, setBulkTargetSubject] = useState('');
  const [bulkTargetBlokk, setBulkTargetBlokk] = useState('1');
  const [showMassUpdateDialog, setShowMassUpdateDialog] = useState(false);
  const [pendingFjern, setPendingFjern] = useState(false);
  const [statusMessage, setStatusMessage] = useState('');

  const getStudentHoverDetails = (studentId: string): string[] => {
    const lookup = studentById.get(studentId);
    if (!lookup) {
      return [];
    }

    return getStudentSubjectsByBlokkLines(lookup.student, visibleBlokkCount);
  };

  const studentById = useMemo(() => {
    const map = new Map<string, StudentLookupEntry>();
    data.forEach((student, index) => {
      const studentId = getStudentId(student, index);
      map.set(studentId, {
        student,
        index,
        studentId,
      });
    });
    return map;
  }, [data]);

  const latestChangeByStudentSubject = useMemo(() => {
    const map = new Map<string, StudentAssignmentChange>();

    changeLog.forEach((entry) => {
      const key = buildChangeKey(entry.studentId, entry.subject);
      const existing = map.get(key);
      if (!existing) {
        map.set(key, entry);
        return;
      }

      if (new Date(entry.changedAt).getTime() >= new Date(existing.changedAt).getTime()) {
        map.set(key, entry);
      }
    });

    return map;
  }, [changeLog]);

  const latestChangesBySubject = useMemo(() => {
    const grouped = new Map<string, StudentAssignmentChange[]>();

    latestChangeByStudentSubject.forEach((change) => {
      const key = normalizeSubjectKey(change.subject);
      const list = grouped.get(key);
      if (list) {
        list.push(change);
        return;
      }

      grouped.set(key, [change]);
    });

    return grouped;
  }, [latestChangeByStudentSubject]);

  const subjectsInData = useMemo(() => {
    const subjectSet = new Set<string>(subjectOptions);

    data.forEach((student) => {
      for (let blokkNumber = 1; blokkNumber <= visibleBlokkCount; blokkNumber += 1) {
        parseSubjects(student[getBlokkKey(blokkNumber)] as string | null).forEach((subject) => {
          subjectSet.add(subject);
        });
      }
    });

    Object.keys(subjectSettingsByName).forEach((subject) => subjectSet.add(subject));
    return Array.from(subjectSet).sort(compareText);
  }, [data, subjectOptions, subjectSettingsByName, visibleBlokkCount]);

  const groupRows = useMemo(() => {
    const rows: GroupRow[] = [];

    subjectsInData.forEach((subject) => {
        const breakdown: Record<BlokkLabel, number> = {
          'Blokk 1': 0,
          'Blokk 2': 0,
          'Blokk 3': 0,
          'Blokk 4': 0,
        };

        const studentIdsByBlokk: StudentIdsByBlokk = {
          'Blokk 1': [],
          'Blokk 2': [],
          'Blokk 3': [],
          'Blokk 4': [],
        };

        data.forEach((student, index) => {
          const studentId = getStudentId(student, index);

          BLOKK_LABELS.slice(0, visibleBlokkCount).forEach((blokkLabel) => {
            const blokkNumber = getBlokkNumber(blokkLabel);
            const subjectsInBlokk = parseSubjects(student[getBlokkKey(blokkNumber)] as string | null);
            if (subjectsInBlokk.some((value) => isSameSubject(value, subject))) {
              breakdown[blokkLabel] += 1;
              studentIdsByBlokk[blokkLabel].push(studentId);
            }
          });
        });

        const settings = getSettingsForSubject(subjectSettingsByName, subject, breakdown);
        const groupsByTarget = getResolvedGroupsByTarget(
          settings.groups || [],
          studentIdsByBlokk,
          settings.groupStudentAssignments || {}
        );

        BLOKK_LABELS.slice(0, visibleBlokkCount).forEach((blokkLabel) => {
          groupsByTarget[blokkLabel]
            .filter(shouldShowGroup)
            .forEach((group) => {
              const activeStudentIds = new Set(group.allocatedStudentIds);
              const recentInCount = group.allocatedStudentIds.filter((studentId) => {
                const change = latestChangeByStudentSubject.get(buildChangeKey(studentId, subject));
                return !!change && change.toBlokk === getBlokkNumber(blokkLabel);
              }).length;

              const changesForSubject = latestChangesBySubject.get(normalizeSubjectKey(subject)) || [];

              const recentOutCount = changesForSubject.filter((entry) => {
                return isSameSubject(entry.subject, subject)
                  && entry.fromBlokk === getBlokkNumber(blokkLabel)
                  && entry.toBlokk !== getBlokkNumber(blokkLabel)
                  && !activeStudentIds.has(entry.studentId);
              }).length;

              rows.push({
                key: `${subject}::${group.id}`,
                subject,
                blokk: blokkLabel,
                blokkNumber: getBlokkNumber(blokkLabel),
                groupId: group.id,
                groupLabel: group.label,
                count: group.allocatedCount,
                max: group.max,
                studentIds: group.allocatedStudentIds,
                recentInCount,
                recentOutCount,
              });
            });
        });
    });

    return rows;
  }, [data, latestChangeByStudentSubject, latestChangesBySubject, subjectSettingsByName, subjectsInData, visibleBlokkCount]);

  const sortedGroupRows = useMemo(() => {
    const rows = [...groupRows];

    const compareRows = (left: GroupRow, right: GroupRow, key: SortKey): number => {
      if (key === 'blokk') {
        return left.blokkNumber - right.blokkNumber;
      }
      if (key === 'subject') {
        return compareText(left.subject, right.subject);
      }
      if (key === 'group') {
        return compareText(left.groupLabel, right.groupLabel);
      }
      if (key === 'count') {
        return left.count - right.count;
      }
      return left.max - right.max;
    };

    rows.sort((left, right) => {
      const primary = compareRows(left, right, sortKey);
      if (primary !== 0) {
        return sortDirection === 'asc' ? primary : -primary;
      }

      const fallbackBlokk = left.blokkNumber - right.blokkNumber;
      if (fallbackBlokk !== 0) {
        return fallbackBlokk;
      }

      const fallbackSubject = compareText(left.subject, right.subject);
      if (fallbackSubject !== 0) {
        return fallbackSubject;
      }

      return compareText(left.groupLabel, right.groupLabel);
    });

    return rows;
  }, [groupRows, sortDirection, sortKey]);

  const expandedGroup = useMemo(() => {
    return sortedGroupRows.find((row) => row.key === expandedGroupKey) || null;
  }, [expandedGroupKey, sortedGroupRows]);

  useEffect(() => {
    if (!expandedGroupKey) {
      return;
    }

    const stillExists = sortedGroupRows.some((row) => row.key === expandedGroupKey);
    if (!stillExists) {
      setExpandedGroupKey(null);
      setSelectedStudentIds([]);
    }
  }, [expandedGroupKey, sortedGroupRows]);

  useEffect(() => {
    if (!expandedGroup) {
      setShowMassUpdateDialog(false);
      return;
    }

    setSelectedStudentIds([]);
    setBulkTargetSubject(expandedGroup.subject);
    setBulkTargetBlokk(String(expandedGroup.blokkNumber));
  }, [expandedGroup]);

  const selectedStudentEntries = useMemo(() => {
    return selectedStudentIds
      .map((studentId) => studentById.get(studentId))
      .filter((entry): entry is StudentLookupEntry => !!entry);
  }, [selectedStudentIds, studentById]);

  const availableSubjectsForModal = useMemo(() => {
    return Array.from(new Set(groupRows.map((row) => row.subject))).sort(compareText);
  }, [groupRows]);

  const availableBlocksForModal = useMemo(() => {
    const targetSubject = bulkTargetSubject.trim();
    if (!targetSubject) {
      return [] as SubjectBlockChoice[];
    }

    const byBlock = new Map<number, number>();
    groupRows.forEach((row) => {
      if (!isSameSubject(row.subject, targetSubject)) {
        return;
      }

      byBlock.set(row.blokkNumber, (byBlock.get(row.blokkNumber) || 0) + row.count);
    });

    return Array.from(byBlock.entries())
      .filter(([blokkNumber]) => {
        return selectedStudentEntries.every((entry) => {
          return isBlockAllowedForStudent(entry.student, blokkNumber, classBlockRestrictions);
        });
      })
      .sort((left, right) => left[0] - right[0])
      .map(([blokkNumber, studentCount]) => ({ blokkNumber, studentCount }));
  }, [bulkTargetSubject, classBlockRestrictions, groupRows, selectedStudentEntries]);

  const targetGroupOptions = useMemo(() => {
    const targetBlokkNumber = Number.parseInt(bulkTargetBlokk, 10);
    if (!bulkTargetSubject.trim() || Number.isNaN(targetBlokkNumber)) {
      return [] as GroupRow[];
    }

    return groupRows
      .filter((row) => isSameSubject(row.subject, bulkTargetSubject.trim()) && row.blokkNumber === targetBlokkNumber)
      .sort((left, right) => compareText(left.groupLabel, right.groupLabel));
  }, [bulkTargetBlokk, bulkTargetSubject, groupRows]);

  useEffect(() => {
    if (availableBlocksForModal.length === 0) {
      setBulkTargetBlokk('');
      return;
    }

    const current = Number.parseInt(bulkTargetBlokk, 10);
    const exists = availableBlocksForModal.some((choice) => choice.blokkNumber === current);
    if (!exists) {
      setBulkTargetBlokk(String(availableBlocksForModal[0].blokkNumber));
    }
  }, [availableBlocksForModal, bulkTargetBlokk]);

  const visibleActiveStudents = useMemo(() => {
    if (!expandedGroup) {
      return [] as VisibleStudentRow[];
    }

    const rows = expandedGroup.studentIds.reduce<VisibleStudentRow[]>((acc, studentId) => {
      const lookup = studentById.get(studentId);
      if (!lookup) {
        return acc;
      }

      const change = latestChangeByStudentSubject.get(buildChangeKey(studentId, expandedGroup.subject));
      const status = getMovementStatus(change, expandedGroup.blokkNumber, true);

      acc.push({
        studentId,
        navn: lookup.student.navn || 'Ukjent',
        klasse: lookup.student.klasse || 'Ingen klasse',
        movement: status.movement,
        statusLabel: status.statusLabel,
        changedAt: change?.changedAt || '',
        reason: change?.reason || '',
      });

      return acc;
    }, []);

    return sortStudents(rows);
  }, [expandedGroup, latestChangeByStudentSubject, studentById]);

  const visibleMovedOutStudents = useMemo(() => {
    if (!expandedGroup) {
      return [] as VisibleStudentRow[];
    }

    const activeStudentIds = new Set(expandedGroup.studentIds);
    const rows: VisibleStudentRow[] = [];

    latestChangeByStudentSubject.forEach((change) => {
      if (!isSameSubject(change.subject, expandedGroup.subject)) {
        return;
      }

      if (change.fromBlokk !== expandedGroup.blokkNumber || change.toBlokk === expandedGroup.blokkNumber) {
        return;
      }

      if (activeStudentIds.has(change.studentId)) {
        return;
      }

      const lookup = studentById.get(change.studentId);
      const status = getMovementStatus(change, expandedGroup.blokkNumber, false);
      rows.push({
        studentId: change.studentId,
        navn: lookup?.student.navn || change.navn || 'Ukjent',
        klasse: lookup?.student.klasse || change.klasse || 'Ingen klasse',
        movement: status.movement,
        statusLabel: status.statusLabel,
        changedAt: change.changedAt,
        reason: change.reason,
      });
    });

    return sortStudents(rows);
  }, [expandedGroup, latestChangeByStudentSubject, studentById]);

  const selectedActiveCount = selectedStudentIds.length;

  const allExpandedStudentsSelected = visibleActiveStudents.length > 0
    && visibleActiveStudents.every((row) => selectedStudentIds.includes(row.studentId));

  const applyStatusMessage = (message: string) => {
    setStatusMessage(message);
    window.setTimeout(() => {
      setStatusMessage('');
    }, 2800);
  };

  const toggleSort = (key: SortKey) => {
    if (sortKey === key) {
      setSortDirection((prev) => (prev === 'asc' ? 'desc' : 'asc'));
      return;
    }

    setSortKey(key);
    setSortDirection('asc');
  };

  const ensureGroupForSubjectBlokk = (
    currentSettings: SubjectSettingsByName,
    subject: string,
    blokkNumber: number,
    preferredGroupId?: string
  ): { nextSettings: SubjectSettingsByName; groupId: string | null } => {
    const blokkLabel = `Blokk ${blokkNumber}` as BlokkLabel;
    const subjectSettings = currentSettings[subject] || {
      defaultMax: DEFAULT_MAX_PER_SUBJECT,
      groups: [],
      groupStudentAssignments: {},
    };
    const currentGroups = ((subjectSettings.groups || []) as SubjectGroup[]).slice();
    const existingGroups = currentGroups.filter((group) => group.blokk === blokkLabel && group.enabled !== false);

    if (preferredGroupId && existingGroups.some((group) => group.id === preferredGroupId)) {
      return {
        nextSettings: currentSettings,
        groupId: preferredGroupId,
      };
    }

    if (existingGroups.length > 0) {
      return {
        nextSettings: currentSettings,
        groupId: existingGroups[0].id,
      };
    }

    const nextGroup = makeGroup(blokkLabel, blokkLabel, subjectSettings.defaultMax || DEFAULT_MAX_PER_SUBJECT, true);
    return {
      nextSettings: {
        ...currentSettings,
        [subject]: {
          ...subjectSettings,
          groups: [...currentGroups, nextGroup],
          groupStudentAssignments: { ...(subjectSettings.groupStudentAssignments || {}) },
        },
      },
      groupId: nextGroup.id,
    };
  };

  const handleToggleAllExpanded = () => {
    if (allExpandedStudentsSelected) {
      setSelectedStudentIds([]);
      return;
    }

    setSelectedStudentIds(visibleActiveStudents.map((row) => row.studentId));
  };

  const handleToggleStudent = (studentId: string) => {
    setSelectedStudentIds((prev) => {
      if (prev.includes(studentId)) {
        return prev.filter((value) => value !== studentId);
      }

      return [...prev, studentId];
    });
  };

  const closeMassUpdateDialog = () => {
    setShowMassUpdateDialog(false);
    setPendingFjern(false);
    if (expandedGroup) {
      setBulkTargetSubject(expandedGroup.subject);
      setBulkTargetBlokk(String(expandedGroup.blokkNumber));
    }
  };

  const openMassUpdateDialog = () => {
    if (!expandedGroup || selectedStudentIds.length === 0) {
      applyStatusMessage('Velg minst en elev først');
      return;
    }

    setBulkTargetSubject(expandedGroup.subject);
    setBulkTargetBlokk(String(expandedGroup.blokkNumber));
    setShowMassUpdateDialog(true);
  };

  const handleRemoveSelected = (): boolean => {
    if (!expandedGroup || selectedStudentIds.length === 0) {
      applyStatusMessage('Velg minst en elev først');
      return false;
    }

    const reason = `Grupper: fjernet valgte elever fra ${expandedGroup.subject} ${expandedGroup.groupLabel}`;
    const result = removeSubjectAssignmentsForStudents(data, expandedGroup.subject, selectedStudentIds, reason);

    if (result.changes.length === 0) {
      applyStatusMessage('Ingen elever ble fjernet fra faget');
      return false;
    }

    const updatedAssignments = Object.fromEntries(
      Object.entries(subjectSettingsByName[expandedGroup.subject]?.groupStudentAssignments || {}).filter(
        ([studentId]) => !selectedStudentIds.includes(studentId)
      )
    );

    onSaveSubjectSettingsByName({
      ...subjectSettingsByName,
      [expandedGroup.subject]: {
        ...(subjectSettingsByName[expandedGroup.subject] || {
          defaultMax: DEFAULT_MAX_PER_SUBJECT,
          groups: [],
        }),
        groupStudentAssignments: updatedAssignments,
      },
    });
    onStudentDataUpdate(result.updatedData, result.changes);
    setSelectedStudentIds([]);
    setShowMassUpdateDialog(false);
    applyStatusMessage(`${result.changes.length} fagvalg fjernet`);
    return true;
  };

  const handleMoveSelected = (): boolean => {
    if (!expandedGroup) {
      return false;
    }

    if (selectedStudentIds.length === 0) {
      applyStatusMessage('Velg minst en elev først');
      return false;
    }

    const targetSubject = bulkTargetSubject.trim();
    const targetBlokkNumber = Number.parseInt(bulkTargetBlokk, 10);

    if (!targetSubject) {
      applyStatusMessage('Velg et målfag');
      return false;
    }

    if (Number.isNaN(targetBlokkNumber) || targetBlokkNumber < 1 || targetBlokkNumber > visibleBlokkCount) {
      applyStatusMessage('Velg en gyldig målblokk');
      return false;
    }

    const blockAllowedForSelection = selectedStudentEntries.every((entry) => {
      return isBlockAllowedForStudent(entry.student, targetBlokkNumber, classBlockRestrictions);
    });

    if (!blockAllowedForSelection) {
      applyStatusMessage('Valgt blokk er ikke tillatt for alle valgte elever');
      return false;
    }

    const preferredGroupId = targetGroupOptions.length > 0
      ? [...targetGroupOptions]
        .sort((left, right) => {
          if (left.count !== right.count) {
            return left.count - right.count;
          }

          return compareText(left.groupLabel, right.groupLabel);
        })[0]?.groupId
      : undefined;

    let nextSettings = { ...subjectSettingsByName };
    const ensuredGroup = ensureGroupForSubjectBlokk(nextSettings, targetSubject, targetBlokkNumber, preferredGroupId);
    nextSettings = ensuredGroup.nextSettings;
    const targetGroupId = ensuredGroup.groupId;

    const selectedSet = new Set(selectedStudentIds);
    const changedStudentIds: string[] = [];
    const skippedStudentIds: string[] = [];
    const changes: StudentAssignmentChange[] = [];

    const nextData = data.map((student, index) => {
      const studentId = getStudentId(student, index);
      if (!selectedSet.has(studentId)) {
        return student;
      }

      const sourceKey = getBlokkKey(expandedGroup.blokkNumber);
      const sourceSubjects = parseSubjects(student[sourceKey] as string | null);
      const hasSourceSubject = sourceSubjects.some((value) => isSameSubject(value, expandedGroup.subject));

      if (!hasSourceSubject) {
        skippedStudentIds.push(studentId);
        return student;
      }

      const sameSubject = isSameSubject(expandedGroup.subject, targetSubject);
      const sameBlokk = expandedGroup.blokkNumber === targetBlokkNumber;
      const sameGroup = targetGroupId === expandedGroup.groupId;

      if (sameSubject && sameBlokk && sameGroup) {
        skippedStudentIds.push(studentId);
        return student;
      }

      const hasTargetElsewhere = Array.from({ length: visibleBlokkCount }, (_, offset) => offset + 1).some((blokkNumber) => {
        const value = parseSubjects(student[getBlokkKey(blokkNumber)] as string | null);
        return value.some((subject) => {
          if (!isSameSubject(subject, targetSubject)) {
            return false;
          }

          if (sameSubject && blokkNumber === expandedGroup.blokkNumber) {
            return false;
          }

          return true;
        });
      });

      if (hasTargetElsewhere) {
        skippedStudentIds.push(studentId);
        return student;
      }

      const navn = student.navn || 'Ukjent';
      const klasse = student.klasse || 'Ingen klasse';
      const changedAt = new Date().toISOString();

      if (sameSubject && sameBlokk) {
        changes.push({
          studentId,
          navn,
          klasse,
          subject: targetSubject,
          fromBlokk: expandedGroup.blokkNumber,
          toBlokk: targetBlokkNumber,
          reason: `Grupper: flyttet ${targetSubject} fra gruppe ${expandedGroup.groupLabel} til ny gruppe i Blokk ${targetBlokkNumber}`,
          changedAt,
        });
        changedStudentIds.push(studentId);
        return student;
      }

      const nextStudent: StandardField = { ...student };
      const nextSourceSubjects = sourceSubjects.filter((value) => !isSameSubject(value, expandedGroup.subject));
      const targetKey = getBlokkKey(targetBlokkNumber);

      nextStudent[sourceKey] = formatSubjects(nextSourceSubjects);

      const targetSubjectsAfterSourceUpdate = parseSubjects(nextStudent[targetKey] as string | null);
      const nextTargetSubjects = targetSubjectsAfterSourceUpdate.some((value) => isSameSubject(value, targetSubject))
        ? targetSubjectsAfterSourceUpdate
        : [...targetSubjectsAfterSourceUpdate, targetSubject];

      if (!sameSubject) {
        changes.push({
          studentId,
          navn,
          klasse,
          subject: expandedGroup.subject,
          fromBlokk: expandedGroup.blokkNumber,
          toBlokk: 0,
          reason: `Grupper: fjernet ${expandedGroup.subject} ${expandedGroup.groupLabel} før bytte til ${targetSubject} i Blokk ${targetBlokkNumber}`,
          changedAt,
        });

        changes.push({
          studentId,
          navn,
          klasse,
          subject: targetSubject,
          fromBlokk: 0,
          toBlokk: targetBlokkNumber,
          reason: `Grupper: la til ${targetSubject} i Blokk ${targetBlokkNumber} fra ${expandedGroup.subject} ${expandedGroup.groupLabel}`,
          changedAt,
        });
      } else {
        changes.push({
          studentId,
          navn,
          klasse,
          subject: targetSubject,
          fromBlokk: expandedGroup.blokkNumber,
          toBlokk: targetBlokkNumber,
          reason: `Grupper: flyttet ${targetSubject} fra ${expandedGroup.groupLabel} til Blokk ${targetBlokkNumber}`,
          changedAt,
        });
      }
      changedStudentIds.push(studentId);

      nextStudent[targetKey] = formatSubjects(nextTargetSubjects);
      return nextStudent;
    });

    if (changes.length === 0) {
      applyStatusMessage(skippedStudentIds.length > 0 ? 'Ingen valgte elever kunne oppdateres' : 'Ingen endringer å lagre');
      return false;
    }

    const changedSet = new Set(changedStudentIds);
    const sourceSettings = nextSettings[expandedGroup.subject] || {
      defaultMax: DEFAULT_MAX_PER_SUBJECT,
      groups: [],
      groupStudentAssignments: {},
    };
    const sourceAssignments = { ...(sourceSettings.groupStudentAssignments || {}) };

    if (!isSameSubject(expandedGroup.subject, targetSubject)) {
      changedStudentIds.forEach((studentId) => {
        delete sourceAssignments[studentId];
      });

      nextSettings[expandedGroup.subject] = {
        ...sourceSettings,
        groupStudentAssignments: sourceAssignments,
      };
    }

    const targetSettings = nextSettings[targetSubject] || {
      defaultMax: DEFAULT_MAX_PER_SUBJECT,
      groups: [],
      groupStudentAssignments: {},
    };
    const targetAssignments = { ...(targetSettings.groupStudentAssignments || {}) };
    changedStudentIds.forEach((studentId) => {
      if (targetGroupId) {
        targetAssignments[studentId] = targetGroupId;
      }
    });

    nextSettings[targetSubject] = {
      ...targetSettings,
      groupStudentAssignments: targetAssignments,
    };

    if (isSameSubject(expandedGroup.subject, targetSubject)) {
      Object.keys(sourceAssignments).forEach((studentId) => {
        if (changedSet.has(studentId) && targetGroupId) {
          sourceAssignments[studentId] = targetGroupId;
        }
      });

      nextSettings[targetSubject] = {
        ...nextSettings[targetSubject],
        groupStudentAssignments: sourceAssignments,
      };
    }

    onSaveSubjectSettingsByName(nextSettings);
    onStudentDataUpdate(nextData, changes);
    setSelectedStudentIds([]);
    setShowMassUpdateDialog(false);
    applyStatusMessage(
      skippedStudentIds.length > 0
        ? `${changes.length} elever oppdatert, ${skippedStudentIds.length} hoppet over`
        : `${changes.length} elever oppdatert`
    );
    return true;
  };

  if (groupRows.length === 0) {
    return <div className={styles.empty}>Ingen grupper tilgjengelig.</div>;
  }

  const renderSortLabel = (label: string, key: SortKey) => {
    const isActive = sortKey === key;
    const suffix = isActive ? (sortDirection === 'asc' ? ' ▲' : ' ▼') : '';
    return `${label}${suffix}`;
  };

  return (
    <div className={styles.wrapper}>
      <div className={styles.headerRow}>
        <div>
          <h3 className={styles.title}>Grupper</h3>
          <p className={styles.subtitle}>Sorter på blokk eller fag, åpne en gruppe og masseoppdater elever direkte.</p>
        </div>
        <div className={styles.summaryBadge}>{groupRows.length} grupper</div>
      </div>

      {statusMessage && <div className={styles.statusMessage}>{statusMessage}</div>}

      <table className={styles.table}>
        <thead>
          <tr>
            <th>
              <button type="button" className={styles.sortButton} onClick={() => toggleSort('blokk')}>
                {renderSortLabel('Blokk', 'blokk')}
              </button>
            </th>
            <th>
              <button type="button" className={styles.sortButton} onClick={() => toggleSort('subject')}>
                {renderSortLabel('Fag', 'subject')}
              </button>
            </th>
            <th>Endringer</th>
            <th>
              <button type="button" className={styles.sortButton} onClick={() => toggleSort('count')}>
                {renderSortLabel('Elever', 'count')}
              </button>
            </th>
          </tr>
        </thead>
        <tbody>
          {sortedGroupRows.map((row) => {
            const isExpanded = expandedGroupKey === row.key;

            return (
              <Fragment key={row.key}>
                <tr
                  key={row.key}
                  className={`${styles.groupRow} ${isExpanded ? styles.groupRowExpanded : ''}`.trim()}
                  onClick={() => setExpandedGroupKey((prev) => (prev === row.key ? null : row.key))}
                >
                  <td>{row.blokk}</td>
                  <td className={styles.subjectCell}>{row.subject}</td>
                  <td>
                    {(() => {
                      const net = row.recentInCount - row.recentOutCount;
                      if (net === 0 && row.recentInCount === 0) {
                        return <span className={styles.changeNone}>—</span>;
                      }
                      return (
                        <span className={net > 0 ? styles.changeAdded : net < 0 ? styles.changeRemoved : styles.changeNone}>
                          {net > 0 ? '+' : ''}{net}
                          {' '}
                          <span className={styles.changeDetail}>
                            (+{row.recentInCount} | -{row.recentOutCount})
                          </span>
                        </span>
                      );
                    })()}
                  </td>
                  <td>{row.count}</td>
                </tr>
                {isExpanded && expandedGroup && expandedGroup.key === row.key && (
                  <tr className={styles.detailRow}>
                    <td colSpan={4}>
                      <div className={styles.detailPanel}>
                        <div className={styles.studentSectionHeader}>
                          <div>
                            <h4 className={styles.studentSectionTitle}>Elever i valgt faggruppe</h4>
                            <p className={styles.studentSectionMeta}>{visibleActiveStudents.length} aktive i gruppen.</p>
                          </div>
                          <label className={styles.selectAllToggle}>
                            <input
                              type="checkbox"
                              checked={allExpandedStudentsSelected}
                              onChange={handleToggleAllExpanded}
                            />
                            <span>Velg alle</span>
                          </label>
                        </div>

                        <div className={styles.selectionActions}>
                          <button
                            type="button"
                            className={styles.primaryButton}
                            onClick={(event) => {
                              event.stopPropagation();
                              openMassUpdateDialog();
                            }}
                          >
                            Masseoppdater ({selectedActiveCount} valgt)
                          </button>
                          <button
                            type="button"
                            className={styles.linkActionButton}
                            onClick={(event) => {
                              event.stopPropagation();
                              setSelectedStudentIds([]);
                            }}
                          >
                            Tøm valg
                          </button>
                        </div>

                        <div className={styles.studentCardsGrid}>
                          {visibleActiveStudents.map((student) => {
                            const isSelected = selectedStudentIds.includes(student.studentId);
                            const hoverDetails = getStudentHoverDetails(student.studentId);

                            return (
                              <button
                                key={student.studentId}
                                type="button"
                                className={[
                                  styles.studentCard,
                                  isSelected ? styles.studentCardSelected : '',
                                  student.movement === 'moved-in' ? styles.studentCardMovedIn : '',
                                ].filter(Boolean).join(' ')}
                                onClick={() => handleToggleStudent(student.studentId)}
                                onDoubleClick={() => onOpenStudentCard?.(student.studentId)}
                              >
                                <span className={styles.studentCardCheckboxWrap}>
                                  <input
                                    type="checkbox"
                                    checked={isSelected}
                                    readOnly
                                    tabIndex={-1}
                                    className={styles.studentCardCheckbox}
                                    aria-hidden="true"
                                  />
                                </span>
                                <span className={styles.studentCardMain}>
                                  <span className={styles.studentNameWithTooltip}>
                                    <span className={styles.studentCardName}>{student.navn}</span>
                                    {hoverDetails.length > 0 && (
                                      <span className={styles.studentNameTooltip} role="tooltip" aria-hidden="true">
                                        {hoverDetails.map((line, index) => (
                                          <span
                                            key={`${student.studentId}-tooltip-${index}`}
                                            className={`${styles.studentNameTooltipLine} ${index % 2 === 1 ? styles.studentNameTooltipLineBold : ''}`.trim()}
                                          >
                                            {line}
                                          </span>
                                        ))}
                                      </span>
                                    )}
                                  </span>
                                  <span className={styles.studentCardMeta}>{student.klasse}</span>
                                </span>
                              </button>
                            );
                          })}
                        </div>

                        {visibleMovedOutStudents.length > 0 && (
                          <div className={styles.movedOutSection}>
                            <h5 className={styles.movedOutTitle}>Flyttet ut av gruppen</h5>
                            <div className={styles.studentCardsGrid}>
                              {visibleMovedOutStudents.map((student) => (
                                (() => {
                                  const hoverDetails = getStudentHoverDetails(student.studentId);
                                  return (
                                    <div
                                      key={student.studentId}
                                      className={`${styles.studentCard} ${styles.studentCardMovedOut}`.trim()}
                                      onDoubleClick={() => onOpenStudentCard?.(student.studentId)}
                                    >
                                      <span className={styles.studentCardMain}>
                                        <span className={styles.studentNameWithTooltip}>
                                          <span className={styles.studentCardName}>{student.navn}</span>
                                          {hoverDetails.length > 0 && (
                                            <span className={styles.studentNameTooltip} role="tooltip" aria-hidden="true">
                                              {hoverDetails.map((line, index) => (
                                                <span
                                                  key={`${student.studentId}-moved-tooltip-${index}`}
                                                  className={`${styles.studentNameTooltipLine} ${index % 2 === 1 ? styles.studentNameTooltipLineBold : ''}`.trim()}
                                                >
                                                  {line}
                                                </span>
                                              ))}
                                            </span>
                                          )}
                                        </span>
                                        <span className={styles.studentCardMeta}>{student.klasse}</span>
                                      </span>
                                    </div>
                                  );
                                })()
                              ))}
                            </div>
                          </div>
                        )}

                        {visibleActiveStudents.length === 0 && visibleMovedOutStudents.length === 0 && (
                          <div className={styles.emptyStateCell}>Ingen elever i denne gruppen.</div>
                        )}
                      </div>
                    </td>
                  </tr>
                )}
              </Fragment>
            );
          })}
        </tbody>
      </table>

      <datalist id="grupper-subject-options">
        {availableSubjectsForModal.map((subject) => (
          <option key={subject} value={subject} />
        ))}
      </datalist>

      {showMassUpdateDialog && expandedGroup && (
        <div className={styles.modalOverlay} onClick={closeMassUpdateDialog}>
          <div className={styles.modalContent} onClick={(event) => event.stopPropagation()}>
            <h3 className={styles.modalTitle}>Massoppdater elever</h3>
            <p className={styles.modalMeta}>{selectedStudentIds.length} elever valgt</p>

            <div className={styles.modalFields}>
              <label className={styles.modalField}>
                <strong>Nytt fag:</strong>
                <select
                  className={styles.modalSelect}
                  value={bulkTargetSubject}
                  onChange={(event) => setBulkTargetSubject(event.target.value)}
                >
                  <option value="">Velg fag...</option>
                  {availableSubjectsForModal.map((subject) => (
                    <option key={subject} value={subject}>
                      {subject}
                    </option>
                  ))}
                </select>
              </label>

              <label className={styles.modalField}>
                <strong>Ny blokk:</strong>
                <select
                  className={styles.modalSelect}
                  value={bulkTargetBlokk}
                  onChange={(event) => setBulkTargetBlokk(event.target.value)}
                >
                  <option value="">Velg blokk...</option>
                  {availableBlocksForModal.map((choice) => (
                    <option key={choice.blokkNumber} value={choice.blokkNumber}>
                      Blokk {choice.blokkNumber} ({choice.studentCount} elever)
                    </option>
                  ))}
                </select>
              </label>
            </div>

            <div className={styles.modalActions}>
              <button type="button" className={styles.primaryButton} onClick={() => { void handleMoveSelected(); }}>
                Endre
              </button>
              {!pendingFjern ? (
                <button type="button" className={styles.secondaryButton} onClick={() => { setPendingFjern(true); }}>
                  Fjern
                </button>
              ) : (
                <>
                  <button type="button" className={styles.secondaryButton} onClick={() => { void handleRemoveSelected(); }}>
                    Bekreft ({selectedStudentIds.length})
                  </button>
                  <button type="button" className={styles.linkActionButton} onClick={() => { setPendingFjern(false); }}>
                    Angre
                  </button>
                </>
              )}
              <button type="button" className={styles.linkActionButton} onClick={closeMassUpdateDialog}>
                Avbryt
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};