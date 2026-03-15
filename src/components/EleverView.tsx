import { useEffect, useMemo, useRef, useState } from 'react';
import type { StandardField, StudentAssignmentChange } from '../utils/excelUtils';
import type { SubjectSettingsByName } from './SubjectTally';
import styles from './EleverView.module.css';

type StudentFilter = 'all' | 'missing' | 'overloaded' | 'collisions' | 'duplicates';
type WarningType = 'missing' | 'overloaded';

interface WarningIgnoreEntry {
  comment: string;
  ignoredAt: string;
}

interface EleverViewProps {
  data: StandardField[];
  blokkCount: number;
  subjectOptions: string[];
  subjectSettingsByName: SubjectSettingsByName;
  onSaveSubjectSettingsByName: (values: SubjectSettingsByName) => void;
  warningIgnoresByStudentAndType: Record<string, Partial<Record<WarningType, WarningIgnoreEntry>>>;
  onSaveWarningIgnore: (studentId: string, type: WarningType, comment: string) => void;
  onRemoveWarningIgnore: (studentId: string, type: WarningType) => void;
  changeLog: StudentAssignmentChange[];
  onStudentDataUpdate: (updatedData: StandardField[], changes: StudentAssignmentChange[]) => void;
  externallySelectedStudentId?: string;
}

interface AssignmentEntry {
  subject: string;
  blokkNumber: number;
}

type BlokkLabel = 'Blokk 1' | 'Blokk 2' | 'Blokk 3' | 'Blokk 4';

interface SubjectGroupSetting {
  id: string;
  blokk: BlokkLabel;
  sourceBlokk: BlokkLabel;
  enabled: boolean;
  max: number;
  createdAt: string;
}

interface SubjectGroupOption {
  id: string;
  label: string;
  count: number;
  max: number;
}

interface SubjectGroupMetrics {
  optionsByBlokk: Record<number, SubjectGroupOption[]>;
  totalsByBlokk: Record<number, { students: number; spaces: number }>;
}

interface EditAssignmentState {
  rowKey: string;
  fromSubject: string;
  fromBlokkNumber: number;
  selectedSubject: string;
  selectedBlokk: string;
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

const getBlokkKey = (blokkNumber: number): keyof StandardField => {
  return `blokk${blokkNumber}` as keyof StandardField;
};

const getStudentId = (student: StandardField, index: number): string => {
  return student.studentId || `${student.navn || 'ukjent'}:${student.klasse || 'ukjent'}:${index}`;
};

const extractAssignments = (student: StandardField, blokkCount: number): AssignmentEntry[] => {
  const entries: AssignmentEntry[] = [];

  for (let blokkNumber = 1; blokkNumber <= blokkCount; blokkNumber += 1) {
    const field = getBlokkKey(blokkNumber);
    const value = student[field];
    const subjects = parseSubjects(typeof value === 'string' ? value : null);

    subjects.forEach((subject) => {
      entries.push({ subject, blokkNumber });
    });
  }

  return entries;
};

const hasMissingSubjects = (student: StandardField, blokkCount: number): boolean => {
  const assignments = extractAssignments(student, blokkCount).filter((entry) => entry.blokkNumber <= 4);
  return assignments.length < 3;
};

const hasTooManySubjects = (student: StandardField, blokkCount: number): boolean => {
  const assignments = extractAssignments(student, blokkCount).filter((entry) => entry.blokkNumber <= 4);
  return assignments.length >= 4;
};

const hasBlokkCollisions = (student: StandardField, blokkCount: number): boolean => {
  const assignments = extractAssignments(student, blokkCount).filter((entry) => entry.blokkNumber <= 4);
  const byBlokk = new Map<number, number>();

  assignments.forEach((entry) => {
    byBlokk.set(entry.blokkNumber, (byBlokk.get(entry.blokkNumber) || 0) + 1);
  });

  return Array.from(byBlokk.values()).some((count) => count > 1);
};

const hasDuplicateSubjects = (student: StandardField, blokkCount: number): boolean => {
  const assignments = extractAssignments(student, blokkCount);
  const seen = new Set<string>();

  for (const assignment of assignments) {
    const normalized = assignment.subject.toLocaleLowerCase('nb');
    if (seen.has(normalized)) {
      return true;
    }
    seen.add(normalized);
  }

  return false;
};

const matchesSearch = (student: StandardField, query: string, index: number): boolean => {
  const trimmedQuery = query.trim().toLocaleLowerCase('nb');
  if (!trimmedQuery) {
    return true;
  }

  const id = getStudentId(student, index).toLocaleLowerCase('nb');
  const navn = (student.navn || '').toLocaleLowerCase('nb');
  const klasse = (student.klasse || '').toLocaleLowerCase('nb');

  return navn.includes(trimmedQuery) || klasse.includes(trimmedQuery) || id.includes(trimmedQuery);
};

const formatTimestamp = (iso: string): string => {
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) {
    return iso;
  }

  return new Intl.DateTimeFormat('nb-NO', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  }).format(date);
};

const formatChangeLabel = (change: { reason: string; fromBlokk: number; toBlokk: number }): string => {
  const reason = change.reason || '';
  if (reason.toUpperCase().includes('BALANSERING')) {
    const from = change.fromBlokk > 0 ? `Blokk ${change.fromBlokk}` : 'ingen blokk';
    const to = change.toBlokk > 0 ? `Blokk ${change.toBlokk}` : 'ingen blokk';
    return `Balansering: ${from} \u2192 ${to}`;
  }
  return reason;
};

const makeGroupId = () => {
  return `group-${Math.random().toString(36).slice(2, 11)}`;
};

export const EleverView = ({
  data,
  blokkCount,
  subjectOptions,
  subjectSettingsByName,
  onSaveSubjectSettingsByName,
  warningIgnoresByStudentAndType,
  onSaveWarningIgnore,
  onRemoveWarningIgnore,
  changeLog,
  onStudentDataUpdate,
  externallySelectedStudentId,
}: EleverViewProps) => {
  const studentRowRefs = useRef<Record<string, HTMLButtonElement | null>>({});
  const [studentQuery, setStudentQuery] = useState('');
  const [activeFilter, setActiveFilter] = useState<StudentFilter>('all');
  const [selectedStudentId, setSelectedStudentId] = useState('');
  const [subjectToAdd, setSubjectToAdd] = useState('');
  const [blokkToAdd, setBlokkToAdd] = useState('');
  const [warningIgnoreDraftByType, setWarningIgnoreDraftByType] = useState<Partial<Record<WarningType, string>>>({});
  const [statusMessage, setStatusMessage] = useState('');
  const [editAssignment, setEditAssignment] = useState<EditAssignmentState | null>(null);
  const [showAddSubjectModal, setShowAddSubjectModal] = useState(false);
  const [logExpanded, setLogExpanded] = useState(false);

  const sortedSubjectOptions = useMemo(() => {
    return subjectOptions.slice().sort((a, b) => a.localeCompare(b, 'nb', { sensitivity: 'base' }));
  }, [subjectOptions]);

  const studentSummaries = useMemo(() => {
    return data.map((student, index) => {
      const assignments = extractAssignments(student, blokkCount);
      const studentId = getStudentId(student, index);
      const missing = hasMissingSubjects(student, blokkCount);
      const overloaded = hasTooManySubjects(student, blokkCount);
      return {
        student,
        index,
        studentId,
        assignments,
        missing,
        overloaded,
        missingIgnored: missing && !!warningIgnoresByStudentAndType[studentId]?.missing,
        overloadedIgnored: overloaded && !!warningIgnoresByStudentAndType[studentId]?.overloaded,
        collisions: hasBlokkCollisions(student, blokkCount),
        duplicates: hasDuplicateSubjects(student, blokkCount),
      };
    });
  }, [data, blokkCount, warningIgnoresByStudentAndType]);

  const counts = useMemo(() => {
    return {
      missing: studentSummaries.filter((entry) => entry.missing && !entry.missingIgnored).length,
      overloaded: studentSummaries.filter((entry) => entry.overloaded && !entry.overloadedIgnored).length,
      collisions: studentSummaries.filter((entry) => entry.collisions).length,
      duplicates: studentSummaries.filter((entry) => entry.duplicates).length,
    };
  }, [studentSummaries]);

  const filteredStudents = useMemo(() => {
    return studentSummaries.filter((entry) => {
      if (!matchesSearch(entry.student, studentQuery, entry.index)) {
        return false;
      }

      if (activeFilter === 'missing') {
        return entry.missing && !entry.missingIgnored;
      }
      if (activeFilter === 'overloaded') {
        return entry.overloaded && !entry.overloadedIgnored;
      }
      if (activeFilter === 'collisions') {
        return entry.collisions;
      }
      if (activeFilter === 'duplicates') {
        return entry.duplicates;
      }

      return true;
    });
  }, [activeFilter, studentQuery, studentSummaries]);

  useEffect(() => {
    if (filteredStudents.length === 0) {
      setSelectedStudentId('');
      return;
    }

    const stillVisible = filteredStudents.some((entry) => entry.studentId === selectedStudentId);
    if (!stillVisible) {
      setSelectedStudentId(filteredStudents[0].studentId);
    }
  }, [filteredStudents, selectedStudentId]);

  useEffect(() => {
    setEditAssignment(null);
    setShowAddSubjectModal(false);
    setSubjectToAdd('');
    setBlokkToAdd('');
    setWarningIgnoreDraftByType({});
    setLogExpanded(false);
  }, [selectedStudentId]);

  useEffect(() => {
    if (!externallySelectedStudentId) {
      return;
    }

    setActiveFilter('all');
    setStudentQuery('');
    setSelectedStudentId(externallySelectedStudentId);

    const scrollToSelectedRow = () => {
      window.requestAnimationFrame(() => {
        window.requestAnimationFrame(() => {
          const row = studentRowRefs.current[externallySelectedStudentId];
          row?.scrollIntoView({
            behavior: 'smooth',
            block: 'center',
            inline: 'nearest',
          });
        });
      });
    };

    scrollToSelectedRow();
  }, [externallySelectedStudentId]);

  const getSubjectGroupMetrics = (subject: string): SubjectGroupMetrics => {
    const result: SubjectGroupMetrics = {
      optionsByBlokk: {},
      totalsByBlokk: {},
    };

    const subjectSettings = subjectSettingsByName[subject];
    const groups = ((subjectSettings?.groups || []) as SubjectGroupSetting[]).slice();
    const assignments = subjectSettings?.groupStudentAssignments || {};
    const defaultMax = typeof subjectSettings?.defaultMax === 'number' ? subjectSettings.defaultMax : 30;

    const subjectCountsByBlokk: Record<number, number> = {};
    for (let blokkNumber = 1; blokkNumber <= Math.min(blokkCount, 4); blokkNumber += 1) {
      let count = 0;
      data.forEach((student) => {
        const blokkValue = student[getBlokkKey(blokkNumber)] as string | null;
        const subjectsInBlokk = parseSubjects(blokkValue);
        if (subjectsInBlokk.some((entry) => isSameSubject(entry, subject))) {
          count += 1;
        }
      });
      subjectCountsByBlokk[blokkNumber] = count;
    }

    const enabledByBlokk = new Map<number, SubjectGroupSetting[]>();

    groups
      .filter((group) => group.enabled !== false)
      .forEach((group) => {
        const blokkNumber = Number.parseInt(group.blokk.replace('Blokk ', ''), 10);
        if (Number.isNaN(blokkNumber)) {
          return;
        }

        const current = enabledByBlokk.get(blokkNumber) || [];
        current.push(group);
        enabledByBlokk.set(blokkNumber, current);
      });

    // Fallback: if a subject exists in a blokk but no persisted group is present, infer one group.
    Object.entries(subjectCountsByBlokk).forEach(([blokkKey, count]) => {
      const blokkNumber = Number.parseInt(blokkKey, 10);
      if (!Number.isFinite(blokkNumber) || count <= 0) {
        return;
      }

      if (enabledByBlokk.has(blokkNumber)) {
        return;
      }

      const inferredGroup: SubjectGroupSetting = {
        id: `inferred-${subject}-${blokkNumber}`,
        blokk: `Blokk ${blokkNumber}` as BlokkLabel,
        sourceBlokk: `Blokk ${blokkNumber}` as BlokkLabel,
        enabled: true,
        max: defaultMax,
        createdAt: '',
      };

      enabledByBlokk.set(blokkNumber, [inferredGroup]);
    });

    enabledByBlokk.forEach((list, blokkNumber) => {
      list.sort((left, right) => {
        if ((left.createdAt || '') !== (right.createdAt || '')) {
          return (left.createdAt || '').localeCompare(right.createdAt || '');
        }
        return (left.id || '').localeCompare(right.id || '');
      });

      const countsByGroupId: Record<string, number> = {};
      list.forEach((group) => {
        countsByGroupId[group.id] = 0;
      });

      const studentIdsInBlokk: string[] = [];

      data.forEach((student, index) => {
        const studentId = getStudentId(student, index);
        const blokkValue = student[getBlokkKey(blokkNumber)] as string | null;
        const subjectsInBlokk = parseSubjects(blokkValue);
        if (subjectsInBlokk.some((entry) => isSameSubject(entry, subject))) {
          studentIdsInBlokk.push(studentId);
        }
      });

      let unassignedCount = 0;
      studentIdsInBlokk.forEach((studentId) => {
        const assignedGroupId = assignments[studentId];
        if (assignedGroupId && countsByGroupId[assignedGroupId] !== undefined) {
          countsByGroupId[assignedGroupId] += 1;
          return;
        }
        unassignedCount += 1;
      });

      if (unassignedCount > 0 && list.length > 0) {
        const base = Math.floor(unassignedCount / list.length);
        const remainder = unassignedCount % list.length;

        list.forEach((group, index) => {
          countsByGroupId[group.id] += base + (index < remainder ? 1 : 0);
        });
      }

      const options = list.map((group, index) => ({
        id: group.id,
        label: `${blokkNumber}-${index + 1}`,
        count: countsByGroupId[group.id] || 0,
        max: typeof group.max === 'number' ? Math.max(0, group.max) : 0,
      }));

      result.optionsByBlokk[blokkNumber] = options;
      result.totalsByBlokk[blokkNumber] = {
        students: options.reduce((sum, option) => sum + option.count, 0),
        spaces: options.reduce((sum, option) => sum + option.max, 0),
      };
    });

    return result;
  };

  const groupMetricsBySubject = useMemo(() => {
    const metrics: Record<string, SubjectGroupMetrics> = {};
    const subjectsInData = new Set<string>();

    data.forEach((student) => {
      for (let blokkNumber = 1; blokkNumber <= Math.min(blokkCount, 4); blokkNumber += 1) {
        parseSubjects(student[getBlokkKey(blokkNumber)] as string | null).forEach((subject) => {
          subjectsInData.add(subject);
        });
      }
    });

    Object.keys(subjectSettingsByName).forEach((subject) => subjectsInData.add(subject));

    subjectsInData.forEach((subject) => {
      metrics[subject] = getSubjectGroupMetrics(subject);
    });

    return metrics;
  }, [data, blokkCount, subjectSettingsByName]);

  const addBlokkOptions = useMemo(() => {
    const normalizedSubject = subjectToAdd.trim();

    if (!normalizedSubject) {
      return [];
    }

    return getBlokkOptionsForSubject(normalizedSubject);
  }, [subjectToAdd, blokkCount, subjectSettingsByName, groupMetricsBySubject]);

  useEffect(() => {
    if (addBlokkOptions.length === 0) {
      setBlokkToAdd('');
      return;
    }

    const hasCurrent = addBlokkOptions.some((blokk) => String(blokk) === blokkToAdd);
    if (!hasCurrent) {
      setBlokkToAdd(String(addBlokkOptions[0]));
    }
  }, [addBlokkOptions, blokkToAdd]);

  function getSubjectGroupOptionsForBlokk(subject: string, blokkNumber: number): SubjectGroupOption[] {
    const metrics = groupMetricsBySubject[subject];
    return metrics?.optionsByBlokk[blokkNumber] || [];
  }

  const getSubjectTotalsForBlokk = (subject: string, blokkNumber: number): { students: number; spaces: number } | null => {
    const metrics = groupMetricsBySubject[subject];
    if (!metrics) {
      return null;
    }

    return metrics.totalsByBlokk[blokkNumber] || null;
  };

  const ensureGroupForSubjectBlokk = (
    subject: string,
    blokkNumber: number,
    preferredGroupId?: string
  ): string | null => {
    const blokkLabel = `Blokk ${blokkNumber}` as BlokkLabel;
    const currentSubjectSettings = subjectSettingsByName[subject] || { defaultMax: 30, groups: [] };
    const currentGroups = ((currentSubjectSettings.groups || []) as SubjectGroupSetting[]).slice();

    const enabledInBlokk = currentGroups.filter((group) => group.blokk === blokkLabel && group.enabled !== false);

    if (preferredGroupId && enabledInBlokk.some((group) => group.id === preferredGroupId)) {
      return preferredGroupId;
    }

    if (enabledInBlokk.length > 0) {
      return enabledInBlokk[0].id;
    }

    const newGroupId = makeGroupId();
    const defaultMax = typeof currentSubjectSettings.defaultMax === 'number'
      ? currentSubjectSettings.defaultMax
      : 30;

    const newGroup = {
      id: newGroupId,
      blokk: blokkLabel,
      sourceBlokk: blokkLabel,
      enabled: true,
      max: defaultMax,
      createdAt: new Date().toISOString(),
    };

    onSaveSubjectSettingsByName({
      ...subjectSettingsByName,
      [subject]: {
        ...currentSubjectSettings,
        groups: [...currentGroups, newGroup],
      },
    });

    return newGroupId;
  };

  const saveSubjectGroupAssignment = (subject: string, studentId: string, groupId: string | null) => {
    const currentSubjectSettings = subjectSettingsByName[subject] || { defaultMax: 30, groups: [] };
    const currentAssignments = {
      ...(currentSubjectSettings.groupStudentAssignments || {}),
    };

    if (groupId) {
      currentAssignments[studentId] = groupId;
    } else {
      delete currentAssignments[studentId];
    }

    onSaveSubjectSettingsByName({
      ...subjectSettingsByName,
      [subject]: {
        ...currentSubjectSettings,
        groupStudentAssignments: currentAssignments,
      },
    });
  };

  function getBlokkOptionsForSubject(subject: string, includeBlokk?: number): number[] {
    const options = Array.from({ length: blokkCount }, (_, index) => index + 1)
      .filter((blokkNumber) => getSubjectGroupOptionsForBlokk(subject, blokkNumber).length > 0);

    if (includeBlokk && !options.includes(includeBlokk)) {
      options.push(includeBlokk);
    }

    return options.sort((a, b) => a - b);
  }

  function chooseLeastPopulatedGroup(
    subject: string,
    blokkNumber: number,
    fallbackGroupId: string | null
  ): string | null {
    const options = getSubjectGroupOptionsForBlokk(subject, blokkNumber);

    if (options.length === 0) {
      return fallbackGroupId;
    }

    const minCount = options.reduce((min, option) => Math.min(min, option.count), Infinity);
    const candidates = options.filter((option) => option.count === minCount);

    if (candidates.length === 0) {
      return fallbackGroupId;
    }

    const randomIndex = Math.floor(Math.random() * candidates.length);
    return candidates[randomIndex].id;
  }

  const selectedStudentEntry = useMemo(() => {
    return filteredStudents.find((entry) => entry.studentId === selectedStudentId) || null;
  }, [filteredStudents, selectedStudentId]);

  const selectedWarningRows = useMemo(() => {
    if (!selectedStudentEntry) {
      return [] as Array<{ type: WarningType; title: string; ignored: WarningIgnoreEntry | null }>;
    }

    const rows: Array<{ type: WarningType; title: string; ignored: WarningIgnoreEntry | null }> = [];

    if (selectedStudentEntry.missing) {
      rows.push({
        type: 'missing',
        title: 'Færre enn 3 blokkfag',
        ignored: warningIgnoresByStudentAndType[selectedStudentEntry.studentId]?.missing || null,
      });
    }

    if (selectedStudentEntry.overloaded) {
      rows.push({
        type: 'overloaded',
        title: '4+ blokkfag',
        ignored: warningIgnoresByStudentAndType[selectedStudentEntry.studentId]?.overloaded || null,
      });
    }

    return rows;
  }, [selectedStudentEntry, warningIgnoresByStudentAndType]);

  const selectedStudentChanges = useMemo(() => {
    if (!selectedStudentEntry) {
      return [];
    }

    return changeLog
      .filter((change) => change.studentId === selectedStudentEntry.studentId)
      .slice()
      .reverse();
  }, [changeLog, selectedStudentEntry]);

  const visibleStudentChanges = useMemo(() => {
    return logExpanded ? selectedStudentChanges : [];
  }, [logExpanded, selectedStudentChanges]);

  const applyStatusMessage = (message: string) => {
    setStatusMessage(message);
    window.setTimeout(() => {
      setStatusMessage('');
    }, 2500);
  };

  const handleSaveWarningIgnore = (type: WarningType) => {
    if (!selectedStudentEntry) {
      return;
    }

    const comment = (warningIgnoreDraftByType[type] || '').trim();
    onSaveWarningIgnore(selectedStudentEntry.studentId, type, comment);
    setWarningIgnoreDraftByType((prev) => ({
      ...prev,
      [type]: '',
    }));
  };

  const handleRemoveAssignment = (subject: string, blokkNumber: number) => {
    if (!selectedStudentEntry) {
      return;
    }

    const nextData = data.map((student, index) => {
      const currentId = getStudentId(student, index);
      if (currentId !== selectedStudentEntry.studentId) {
        return student;
      }

      const blokkKey = getBlokkKey(blokkNumber);
      const existingSubjects = parseSubjects(student[blokkKey] as string | null);
      let removed = false;
      const remainingSubjects = existingSubjects.filter((item) => {
        if (!removed && isSameSubject(item, subject)) {
          removed = true;
          return false;
        }
        return true;
      });

      return {
        ...student,
        [blokkKey]: remainingSubjects.length > 0 ? remainingSubjects.join(', ') : null,
      };
    });

    const change: StudentAssignmentChange = {
      studentId: selectedStudentEntry.studentId,
      navn: selectedStudentEntry.student.navn || 'Ukjent',
      klasse: selectedStudentEntry.student.klasse || 'Ingen klasse',
      subject,
      fromBlokk: blokkNumber,
      toBlokk: 0,
      reason: `Elever: fjernet ${subject} fra Blokk ${blokkNumber}`,
      changedAt: new Date().toISOString(),
    };

    onStudentDataUpdate(nextData, [change]);
    saveSubjectGroupAssignment(subject, selectedStudentEntry.studentId, null);
    applyStatusMessage(`${subject} fjernet fra Blokk ${blokkNumber}`);
  };

  const handleAddSubject = () => {
    if (!selectedStudentEntry) {
      return;
    }

    const normalizedSubject = subjectToAdd.trim();
    const blokkNumber = Number.parseInt(blokkToAdd, 10);

    if (!normalizedSubject) {
      applyStatusMessage('Skriv inn eller velg et fag først');
      return;
    }

    if (Number.isNaN(blokkNumber) || blokkNumber < 1 || blokkNumber > blokkCount) {
      applyStatusMessage('Velg en gyldig blokk');
      return;
    }

    const validBlokker = getBlokkOptionsForSubject(normalizedSubject);
    if (!validBlokker.includes(blokkNumber)) {
      applyStatusMessage('Valgt blokk finnes ikke for dette faget');
      return;
    }

    const seededGroupId = ensureGroupForSubjectBlokk(normalizedSubject, blokkNumber, undefined);
    const selectedGroupId = chooseLeastPopulatedGroup(normalizedSubject, blokkNumber, seededGroupId);

    const alreadyAssigned = selectedStudentEntry.assignments.some((assignment) =>
      isSameSubject(assignment.subject, normalizedSubject)
    );

    if (alreadyAssigned) {
      applyStatusMessage('Eleven har allerede dette faget');
      return;
    }

    const nextData = data.map((student, index) => {
      const currentId = getStudentId(student, index);
      if (currentId !== selectedStudentEntry.studentId) {
        return student;
      }

      const blokkKey = getBlokkKey(blokkNumber);
      const existingSubjects = parseSubjects(student[blokkKey] as string | null);

      return {
        ...student,
        [blokkKey]: [...existingSubjects, normalizedSubject].join(', '),
      };
    });

    const change: StudentAssignmentChange = {
      studentId: selectedStudentEntry.studentId,
      navn: selectedStudentEntry.student.navn || 'Ukjent',
      klasse: selectedStudentEntry.student.klasse || 'Ingen klasse',
      subject: normalizedSubject,
      fromBlokk: 0,
      toBlokk: blokkNumber,
      reason: `Elever: la til ${normalizedSubject} i Blokk ${blokkNumber}`,
      changedAt: new Date().toISOString(),
    };

    onStudentDataUpdate(nextData, [change]);
    saveSubjectGroupAssignment(normalizedSubject, selectedStudentEntry.studentId, selectedGroupId);
    setShowAddSubjectModal(false);
    setSubjectToAdd('');
    setBlokkToAdd('');
    applyStatusMessage(`${normalizedSubject} lagt til i Blokk ${blokkNumber}`);
  };

  const openAddSubjectModal = () => {
    if (!selectedStudentEntry) {
      return;
    }

    if (!subjectToAdd && sortedSubjectOptions.length > 0) {
      setSubjectToAdd(sortedSubjectOptions[0]);
    }

    setShowAddSubjectModal(true);
  };

  const openEditAssignment = (assignment: AssignmentEntry, rowKey: string) => {
    const subject = assignment.subject;
    const blokkNumber = assignment.blokkNumber;
    const blokkOptions = getBlokkOptionsForSubject(subject, blokkNumber);
    const initialBlokk = blokkOptions.includes(blokkNumber)
      ? blokkNumber
      : (blokkOptions[0] ?? blokkNumber);

    setEditAssignment({
      rowKey,
      fromSubject: subject,
      fromBlokkNumber: blokkNumber,
      selectedSubject: subject,
      selectedBlokk: String(initialBlokk),
    });
  };

  const handleSaveEditAssignment = () => {
    if (!selectedStudentEntry || !editAssignment) {
      return;
    }

    const fromSubject = editAssignment.fromSubject;
    const fromBlokk = editAssignment.fromBlokkNumber;
    const toSubject = editAssignment.selectedSubject.trim();
    const toBlokk = Number.parseInt(editAssignment.selectedBlokk, 10);

    if (!toSubject) {
      applyStatusMessage('Velg et fag');
      return;
    }

    if (Number.isNaN(toBlokk) || toBlokk < 1 || toBlokk > blokkCount) {
      applyStatusMessage('Velg en gyldig blokk');
      return;
    }

    const duplicateElsewhere = selectedStudentEntry.assignments.some((assignment) => {
      if (!isSameSubject(assignment.subject, toSubject)) {
        return false;
      }

      return !(isSameSubject(assignment.subject, fromSubject) && assignment.blokkNumber === fromBlokk);
    });

    if (duplicateElsewhere) {
      applyStatusMessage('Eleven har allerede dette faget et annet sted');
      return;
    }

    const seededGroupId = ensureGroupForSubjectBlokk(toSubject, toBlokk, undefined);
    const targetGroupId = chooseLeastPopulatedGroup(toSubject, toBlokk, seededGroupId);

    const nextData = data.map((student, index) => {
      const currentId = getStudentId(student, index);
      if (currentId !== selectedStudentEntry.studentId) {
        return student;
      }

      const fromKey = getBlokkKey(fromBlokk);
      const toKey = getBlokkKey(toBlokk);
      const fromSubjects = parseSubjects(student[fromKey] as string | null);
      const toSubjects = parseSubjects(student[toKey] as string | null);

      let removed = false;
      const remainingSourceSubjects = fromSubjects.filter((item) => {
        if (!removed && isSameSubject(item, fromSubject)) {
          removed = true;
          return false;
        }
        return true;
      });

      if (!removed) {
        return student;
      }

      const nextTargetSubjects = toSubjects.some((item) => isSameSubject(item, toSubject))
        ? toSubjects
        : [...toSubjects, toSubject];

      return {
        ...student,
        [fromKey]: remainingSourceSubjects.length > 0 ? remainingSourceSubjects.join(', ') : null,
        [toKey]: nextTargetSubjects.join(', '),
      };
    });

    const change: StudentAssignmentChange = {
      studentId: selectedStudentEntry.studentId,
      navn: selectedStudentEntry.student.navn || 'Ukjent',
      klasse: selectedStudentEntry.student.klasse || 'Ingen klasse',
      subject: toSubject,
      fromBlokk,
      toBlokk,
      reason: `Elever: endret ${fromSubject} (Blokk ${fromBlokk}) til ${toSubject} (Blokk ${toBlokk})`,
      changedAt: new Date().toISOString(),
    };

    onStudentDataUpdate(nextData, [change]);
    if (!isSameSubject(fromSubject, toSubject)) {
      saveSubjectGroupAssignment(fromSubject, selectedStudentEntry.studentId, null);
    }
    saveSubjectGroupAssignment(toSubject, selectedStudentEntry.studentId, targetGroupId);
    setEditAssignment(null);
    applyStatusMessage(`Oppdatert til ${toSubject} i Blokk ${toBlokk}`);
  };

  const modalSubjectOptions = useMemo(() => {
    const set = new Set(subjectOptions);
    if (editAssignment?.fromSubject) {
      set.add(editAssignment.fromSubject);
    }

    return Array.from(set).sort((a, b) => a.localeCompare(b, 'nb', { sensitivity: 'base' }));
  }, [subjectOptions, editAssignment]);

  const modalBlokkOptions = useMemo(() => {
    if (!editAssignment) {
      return [] as number[];
    }

    const includeCurrent = isSameSubject(editAssignment.selectedSubject, editAssignment.fromSubject)
      ? editAssignment.fromBlokkNumber
      : undefined;

    return getBlokkOptionsForSubject(editAssignment.selectedSubject, includeCurrent);
  }, [editAssignment, groupMetricsBySubject, blokkCount]);

  const handleModalSubjectChange = (value: string) => {
    if (!editAssignment) {
      return;
    }

    const includeCurrent = isSameSubject(value, editAssignment.fromSubject)
      ? editAssignment.fromBlokkNumber
      : undefined;
    const blokkOptions = getBlokkOptionsForSubject(value, includeCurrent);
    const nextBlokk = blokkOptions[0] ?? editAssignment.fromBlokkNumber;

    setEditAssignment({
      ...editAssignment,
      selectedSubject: value,
      selectedBlokk: String(nextBlokk),
    });
  };

  if (data.length === 0) {
    return <p className={styles.empty}>Ingen elevdata tilgjengelig.</p>;
  }

  return (
    <div className={styles.wrapper}>
      <div className={styles.filterGroup}>
        <div className={styles.filterLabel}>Elevfilter</div>
        <div className={styles.filters}>
          <button
            type="button"
            className={`${styles.filterButton} ${activeFilter === 'all' ? styles.filterButtonActive : ''}`.trim()}
            onClick={() => setActiveFilter('all')}
          >
            Alle ({data.length})
          </button>
          <button
            type="button"
            className={`${styles.filterButton} ${activeFilter === 'missing' ? styles.filterButtonActive : ''}`.trim()}
            onClick={() => setActiveFilter('missing')}
            disabled={counts.missing === 0}
          >
            Mangler fag ({counts.missing})
          </button>
          <button
            type="button"
            className={`${styles.filterButton} ${activeFilter === 'overloaded' ? styles.filterButtonActive : ''}`.trim()}
            onClick={() => setActiveFilter('overloaded')}
            disabled={counts.overloaded === 0}
          >
            For mange fag ({counts.overloaded})
          </button>
          <button
            type="button"
            className={`${styles.filterButton} ${activeFilter === 'collisions' ? styles.filterButtonActive : ''}`.trim()}
            onClick={() => setActiveFilter('collisions')}
            disabled={counts.collisions === 0}
          >
            Blokk-kollisjoner ({counts.collisions})
          </button>
          <button
            type="button"
            className={`${styles.filterButton} ${activeFilter === 'duplicates' ? styles.filterButtonActive : ''}`.trim()}
            onClick={() => setActiveFilter('duplicates')}
            disabled={counts.duplicates === 0}
          >
            Duplikater ({counts.duplicates})
          </button>
        </div>
      </div>

      <div className={styles.viewerGrid}>
        <aside className={styles.listPanel}>
          <input
            type="search"
            className={styles.searchInput}
            value={studentQuery}
            onChange={(event) => setStudentQuery(event.target.value)}
            placeholder="Sok elevnavn, klasse eller id"
          />

          <div className={styles.studentList}>
            {filteredStudents.map((entry) => (
              <button
                key={entry.studentId}
                type="button"
                ref={(element) => {
                  studentRowRefs.current[entry.studentId] = element;
                }}
                className={`${styles.studentRow} ${entry.studentId === selectedStudentId ? styles.studentRowActive : ''}`.trim()}
                onClick={() => setSelectedStudentId(entry.studentId)}
              >
                <span className={styles.studentName}>{entry.student.navn || 'Ukjent elev'}</span>
                <small className={styles.studentMeta}>
                  {entry.student.klasse || 'Ingen klasse'} | {entry.assignments.length} fag
                </small>
              </button>
            ))}
          </div>
        </aside>

        <section className={styles.detailPanel}>
          {!selectedStudentEntry ? (
            <p className={styles.empty}>Ingen elever matcher filteret.</p>
          ) : (
            <>
              <div className={styles.studentHeader}>
                <div>
                  <h3>{selectedStudentEntry.student.navn || 'Ukjent elev'}</h3>
                  <p>{selectedStudentEntry.student.klasse || 'Ingen klasse'}</p>
                </div>
              </div>

              {selectedWarningRows.length > 0 && (
                <div className={styles.warningIgnorePanel}>
                  {selectedWarningRows.map((warning) => (
                    <div key={`warning-${warning.type}`} className={styles.warningIgnoreRow}>
                      <div className={styles.warningIgnoreTitle}>{warning.title}</div>
                      {warning.ignored ? (
                        <div className={styles.warningIgnoreControls}>
                          <span className={styles.warningIgnoredLabel}>Ignorert: {warning.ignored.comment || 'Ingen kommentar'}</span>
                          <button
                            type="button"
                            className={styles.warningIgnoreButton}
                            onClick={() => onRemoveWarningIgnore(selectedStudentEntry.studentId, warning.type)}
                          >
                            Fjern ignorering
                          </button>
                        </div>
                      ) : (
                        <div className={styles.warningIgnoreControls}>
                          <input
                            type="text"
                            maxLength={140}
                            className={styles.warningIgnoreInput}
                            placeholder="Kommentar (valgfritt)"
                            value={warningIgnoreDraftByType[warning.type] || ''}
                            onChange={(event) => {
                              const value = event.target.value;
                              setWarningIgnoreDraftByType((prev) => ({
                                ...prev,
                                [warning.type]: value,
                              }));
                            }}
                          />
                          <button
                            type="button"
                            className={styles.warningIgnoreButton}
                            onClick={() => handleSaveWarningIgnore(warning.type)}
                          >
                            Ignorer
                          </button>
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              )}

              {statusMessage && <p className={styles.statusMessage}>{statusMessage}</p>}

              <table className={styles.assignmentTable}>
                <thead>
                  <tr>
                    <th>Fag</th>
                    <th>Blokk</th>
                    <th>Handling</th>
                  </tr>
                </thead>
                <tbody>
                  {selectedStudentEntry.assignments.length === 0 ? (
                    <tr>
                      <td colSpan={3} className={styles.emptyCell}>Ingen fag registrert</td>
                    </tr>
                  ) : (
                    selectedStudentEntry.assignments
                      .slice()
                      .sort((a, b) => {
                        if (a.blokkNumber !== b.blokkNumber) {
                          return a.blokkNumber - b.blokkNumber;
                        }
                        return a.subject.localeCompare(b.subject, 'nb', { sensitivity: 'base' });
                      })
                      .map((assignment, index) => {
                        const rowKey = `${assignment.subject}-${assignment.blokkNumber}-${index}`;

                        return (
                        <tr key={rowKey}>
                          <td>{assignment.subject}</td>
                          <td>Blokk {assignment.blokkNumber}</td>
                          <td className={styles.actionsCell}>
                            <div className={styles.actionsButtons}>
                              <button
                                type="button"
                                className={styles.moveButton}
                                onClick={() => openEditAssignment(assignment, rowKey)}
                                title={`Endre ${assignment.subject}`}
                              >
                                Endre
                              </button>
                              <button
                                type="button"
                                className={styles.removeButton}
                                onClick={() => handleRemoveAssignment(assignment.subject, assignment.blokkNumber)}
                              >
                                Fjern
                              </button>
                            </div>
                          </td>
                        </tr>
                      )})
                  )}
                </tbody>
              </table>

              <div className={styles.addBar}>
                <button type="button" className={styles.addButton} onClick={openAddSubjectModal}>
                  Legg til fag
                </button>
              </div>

              {editAssignment && (
                <div className={styles.modalOverlay} onClick={() => setEditAssignment(null)}>
                  <div className={styles.modal} onClick={(event) => event.stopPropagation()}>
                    <h4>Endre fagvalg</h4>
                    <div className={styles.modalRow}>
                      <label className={styles.modalLabel} htmlFor="edit-subject-select">Fag</label>
                      <select
                        id="edit-subject-select"
                        className={styles.moveSelect}
                        value={editAssignment.selectedSubject}
                        onChange={(event) => handleModalSubjectChange(event.target.value)}
                      >
                        {modalSubjectOptions.map((subject) => (
                          <option key={`edit-subject-${subject}`} value={subject}>
                            {subject}
                          </option>
                        ))}
                      </select>
                    </div>
                    <div className={styles.modalRow}>
                      <label className={styles.modalLabel} htmlFor="edit-blokk-select">Blokk</label>
                      <select
                        id="edit-blokk-select"
                        className={styles.moveSelect}
                        value={editAssignment.selectedBlokk}
                        onChange={(event) => {
                          setEditAssignment({
                            ...editAssignment,
                            selectedBlokk: event.target.value,
                          });
                        }}
                      >
                        {modalBlokkOptions.map((blokk) => {
                          const totals = getSubjectTotalsForBlokk(editAssignment.selectedSubject, blokk);
                          const students = totals?.students ?? 0;
                          const spaces = totals?.spaces ?? 0;

                          return (
                            <option key={`edit-blokk-${blokk}`} value={String(blokk)}>
                              Blokk {blokk} ({students} / {spaces})
                            </option>
                          );
                        })}
                      </select>
                    </div>
                    <div className={styles.modalActions}>
                      <button
                        type="button"
                        className={styles.removeButton}
                        onClick={() => setEditAssignment(null)}
                      >
                        Avbryt
                      </button>
                      <button
                        type="button"
                        className={styles.moveButton}
                        onClick={handleSaveEditAssignment}
                      >
                        Lagre
                      </button>
                    </div>
                  </div>
                </div>
              )}

              {showAddSubjectModal && (
                <div className={styles.modalOverlay} onClick={() => setShowAddSubjectModal(false)}>
                  <div className={styles.modal} onClick={(event) => event.stopPropagation()}>
                    <h4>Legg til fag</h4>
                    <div className={styles.modalRow}>
                      <label className={styles.modalLabel} htmlFor="add-subject-select">Fag</label>
                      <select
                        id="add-subject-select"
                        className={styles.moveSelect}
                        value={subjectToAdd}
                        onChange={(event) => setSubjectToAdd(event.target.value)}
                      >
                        {sortedSubjectOptions.map((subject) => (
                          <option key={`add-subject-${subject}`} value={subject}>
                            {subject}
                          </option>
                        ))}
                      </select>
                    </div>
                    <div className={styles.modalRow}>
                      <label className={styles.modalLabel} htmlFor="add-blokk-select">Blokk</label>
                      <select
                        id="add-blokk-select"
                        className={styles.moveSelect}
                        value={blokkToAdd}
                        onChange={(event) => setBlokkToAdd(event.target.value)}
                        disabled={addBlokkOptions.length === 0}
                      >
                        {addBlokkOptions.length === 0 ? (
                          <option value="">Ingen blokker for valgt fag</option>
                        ) : (
                          addBlokkOptions.map((blokk) => {
                            const totals = getSubjectTotalsForBlokk(subjectToAdd.trim(), blokk);
                            const students = totals?.students ?? 0;
                            const spaces = totals?.spaces ?? 0;

                            return (
                              <option key={`add-blokk-${blokk}`} value={String(blokk)}>
                                Blokk {blokk} ({students} / {spaces})
                              </option>
                            );
                          })
                        )}
                      </select>
                    </div>
                    <div className={styles.modalActions}>
                      <button
                        type="button"
                        className={styles.removeButton}
                        onClick={() => setShowAddSubjectModal(false)}
                      >
                        Avbryt
                      </button>
                      <button
                        type="button"
                        className={styles.moveButton}
                        onClick={handleAddSubject}
                      >
                        Legg til
                      </button>
                    </div>
                  </div>
                </div>
              )}

              <div className={styles.logPanel}>
                <button
                  type="button"
                  className={styles.logHeaderButton}
                  onClick={() => setLogExpanded((prev) => !prev)}
                >
                  <span>Endringslogg ({selectedStudentChanges.length})</span>
                  <span className={styles.logHeaderChevron}>{logExpanded ? '−' : '+'}</span>
                </button>
                {logExpanded && (
                  selectedStudentChanges.length === 0 ? (
                    <p className={styles.logEmpty}>Ingen endringer registrert for denne eleven ennå.</p>
                  ) : (
                    <ul className={styles.logList}>
                      {visibleStudentChanges.map((change, index) => (
                        <li key={`${change.changedAt}-${index}`}>
                          <span title={change.reason}>{change.subject ? <strong>{change.subject}</strong> : null}{change.subject ? ': ' : ''}{formatChangeLabel(change)}</span>
                          <small>{formatTimestamp(change.changedAt)}</small>
                        </li>
                      ))}
                    </ul>
                  )
                )}
              </div>
            </>
          )}
        </section>
      </div>
    </div>
  );
};
