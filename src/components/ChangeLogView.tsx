import { useEffect, useMemo, useRef, useState } from 'react';
import type { StandardField, StudentAssignmentChange, GroupMoveLogEntry } from '../utils/excelUtils';
import { loadXlsx } from '../utils/excelUtils';
import {
  BLOKK_LABELS,
  getBlokkNumber,
  getResolvedGroupsByTarget,
  getSettingsForSubject,
  type BlokkLabel,
  type SubjectSettingsByName,
} from '../utils/subjectGroups';
import styles from './ChangeLogView.module.css';

interface ChangeLogViewProps {
  changeLog: StudentAssignmentChange[];
  groupMoveLog: GroupMoveLogEntry[];
  currentStudents: StandardField[];
  subjectSettingsByName: SubjectSettingsByName;
  excludedSubjects: string[];
  onOpenStudentCard?: (studentId: string) => void;
}

interface GroupedStudentChange {
  studentId: string;
  navn: string;
  klasse: string;
  changes: StudentAssignmentChange[];
}

interface StudentStatusLogEntry {
  action: 'added' | 'removed' | 'readded';
  changedAt: string;
  reason: string;
}

interface FlattenedStudentStatusEntry extends StudentStatusLogEntry {
  studentId: string;
  navn: string;
  klasse: string;
}

interface BalancingRoundOption {
  id: number;
  timestamp: string;
}

interface SummaryEntry {
  subject: string;
  fromBlokk: number;
  toBlokk: number;
  lastChangedAt: string;
}

interface UnresolvedBalanceWarning {
  key: string;
  studentId: string;
  navn: string;
  klasse: string;
  reason: string;
}

type LogMode = 'detailed' | 'summary';
type ChangeType = 'added' | 'removed' | 'moved';
type BlokkField = 'blokk1' | 'blokk2' | 'blokk3' | 'blokk4' | 'blokk5' | 'blokk6' | 'blokk7' | 'blokk8';

const BLOKK_FIELDS: BlokkField[] = ['blokk1', 'blokk2', 'blokk3', 'blokk4', 'blokk5', 'blokk6', 'blokk7', 'blokk8'];

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

const formatDateForFilename = (value: Date): string => {
  const year = value.getFullYear();
  const month = `${value.getMonth() + 1}`.padStart(2, '0');
  const day = `${value.getDate()}`.padStart(2, '0');
  return `${year}-${month}-${day}`;
};

const escapeHtml = (value: string): string => {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
};

const parseSubjects = (value: string | null | undefined): string[] => {
  if (!value) {
    return [];
  }

  return value
    .split(/[,;]/)
    .map((subject) => subject.trim())
    .filter((subject) => subject.length > 0);
};

const normalizeSearchText = (value: string): string => {
  return value.trim().toLocaleLowerCase('nb');
};

const parseSearchTokens = (value: string): string[] => {
  return normalizeSearchText(value)
    .split(/\s+/)
    .map((token) => token.trim())
    .filter((token) => token.length > 0);
};

const isSameSubject = (left: string, right: string): boolean => {
  return left.localeCompare(right, 'nb', { sensitivity: 'base' }) === 0;
};

const getStudentId = (student: StandardField, index: number): string => {
  return student.studentId || `${student.navn || 'ukjent'}:${student.klasse || 'ukjent'}:${index}`;
};

const isManualStudentId = (studentId: string): boolean => {
  return studentId.startsWith('manual:');
};

const getBlokkKey = (blokkNumber: number): BlokkField => {
  return `blokk${blokkNumber}` as BlokkField;
};

const formatFinalAllocation = (student: StandardField | undefined): string => {
  if (!student) {
    return '';
  }

  return BLOKK_FIELDS
    .map((blokkField, index) => {
      const subjects = parseSubjects(student[blokkField] as string | null | undefined);
      if (subjects.length === 0) {
        return '';
      }

      return `B${index + 1}: ${subjects.join(', ')}`;
    })
    .filter((value) => value.length > 0)
    .join(' | ');
};

const getChangeType = (fromBlokk: number, toBlokk: number): ChangeType => {
  if (fromBlokk <= 0 && toBlokk > 0) {
    return 'added';
  }

  if (fromBlokk > 0 && toBlokk <= 0) {
    return 'removed';
  }

  return 'moved';
};

const getChangeItemClass = (changeType: ChangeType, stylesMap: Record<string, string>): string => {
  if (changeType === 'added') {
    return stylesMap.changeItemAdded;
  }

  if (changeType === 'removed') {
    return stylesMap.changeItemRemoved;
  }

  return stylesMap.changeItemMoved;
};

const formatDetailedEntryLabel = (entry: StudentAssignmentChange): string => {
  const reason = entry.reason || '';
  if (reason.toUpperCase().includes('BALANSERING')) {
    const fromLabel = entry.fromBlokk > 0 ? `Blokk ${entry.fromBlokk}` : 'ingen blokk';
    const toLabel = entry.toBlokk > 0 ? `Blokk ${entry.toBlokk}` : 'ingen blokk';
    return `Balansering: ${fromLabel} \u2192 ${toLabel}`;
  }
  return reason;
};

const getDetailedEntryClass = (entry: StudentAssignmentChange, stylesMap: Record<string, string>): string => {
  const reasonUpper = (entry.reason || '').toUpperCase();
  if (reasonUpper.includes('ADVARSEL')) {
    return stylesMap.changeItemWarning;
  }

  if (reasonUpper.includes('BALANSERING')) {
    return stylesMap.changeItemNotice;
  }

  return getChangeItemClass(getChangeType(entry.fromBlokk, entry.toBlokk), stylesMap);
};

const getWordLineStyle = (changeType: ChangeType): string => {
  if (changeType === 'added') {
    return 'background:#eaf9f0;border-left:3px solid #2f8f5b;color:#1f5d3d;';
  }

  if (changeType === 'removed') {
    return 'background:#fff1f1;border-left:3px solid #c45555;color:#7a2c2c;';
  }

  return 'background:#eef5ff;border-left:3px solid #2a63b7;color:#1f3f6c;';
};

const isStudentStatusChange = (change: StudentAssignmentChange): boolean => {
  if (change.changeCategory === 'student-status') {
    return true;
  }

  if (change.studentStatusAction) {
    return true;
  }

  const reasonLower = (change.reason || '').toLocaleLowerCase('nb');
  return reasonLower.includes('la til elev')
    || reasonLower.includes('fjernet elev fra elevlisten')
    || reasonLower.includes('gjenla til elev i elevlisten');
};

const resolveStatusAction = (change: StudentAssignmentChange): StudentStatusLogEntry['action'] | null => {
  if (change.studentStatusAction) {
    return change.studentStatusAction;
  }

  const reasonLower = (change.reason || '').toLocaleLowerCase('nb');
  if (reasonLower.includes('gjenla til elev i elevlisten')) {
    return 'readded';
  }
  if (reasonLower.includes('fjernet elev fra elevlisten')) {
    return 'removed';
  }
  if (reasonLower.includes('la til elev')) {
    return 'added';
  }

  return null;
};

const formatStudentStatusLabel = (entry: StudentStatusLogEntry): string => {
  if (entry.action === 'added') {
    return 'Elev lagt til';
  }
  if (entry.action === 'removed') {
    return 'Elev fjernet fra elevliste';
  }
  return 'Elev lagt tilbake i elevliste';
};

export const ChangeLogView = ({
  changeLog,
  groupMoveLog,
  currentStudents,
  subjectSettingsByName,
  excludedSubjects,
  onOpenStudentCard,
}: ChangeLogViewProps) => {
  const [mode, setMode] = useState<LogMode>('summary');
  const [sortBy, setSortBy] = useState<'name' | 'programomrade'>('name');
  const [warningsExpanded, setWarningsExpanded] = useState(false);
  const [groupMovesExpanded, setGroupMovesExpanded] = useState(true);
  const [studentStatusExpanded, setStudentStatusExpanded] = useState(true);
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedRoundId, setSelectedRoundId] = useState<'final' | number>('final');
  const [exportDropdownOpen, setExportDropdownOpen] = useState(false);
  const exportDropdownRef = useRef<HTMLDivElement>(null);

  // Close export dropdown on outside click
  useEffect(() => {
    if (!exportDropdownOpen) return;
    const handleClick = (e: MouseEvent) => {
      if (exportDropdownRef.current && !exportDropdownRef.current.contains(e.target as Node)) {
        setExportDropdownOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClick);
    return () => document.removeEventListener('mousedown', handleClick);
  }, [exportDropdownOpen]);

  const balancingRoundOptions = useMemo(() => {
    const byRoundId = new Map<number, string>();

    changeLog.forEach((entry) => {
      if (typeof entry.balancingRoundId !== 'number') {
        return;
      }

      const existing = byRoundId.get(entry.balancingRoundId);
      if (!existing || new Date(entry.changedAt).getTime() > new Date(existing).getTime()) {
        byRoundId.set(entry.balancingRoundId, entry.changedAt);
      }
    });

    return Array.from(byRoundId.entries())
      .map(([id, timestamp]) => ({ id, timestamp } as BalancingRoundOption))
      .sort((left, right) => {
        return new Date(right.timestamp).getTime() - new Date(left.timestamp).getTime();
      });
  }, [changeLog]);

  useEffect(() => {
    if (balancingRoundOptions.length === 0) {
      if (selectedRoundId !== 'final') {
        setSelectedRoundId('final');
      }
      return;
    }

    if (selectedRoundId === 'final') {
      return;
    }

    const stillExists = balancingRoundOptions.some((round) => round.id === selectedRoundId);
    if (!stillExists) {
      setSelectedRoundId('final');
    }
  }, [balancingRoundOptions, selectedRoundId]);

  const effectiveChangeLog = useMemo(() => {
    if (selectedRoundId === 'final') {
      return changeLog;
    }

    const selectedRound = balancingRoundOptions.find((round) => round.id === selectedRoundId);
    if (!selectedRound) {
      return changeLog;
    }

    const cutoff = new Date(selectedRound.timestamp).getTime();
    return changeLog.filter((entry) => new Date(entry.changedAt).getTime() <= cutoff);
  }, [balancingRoundOptions, changeLog, selectedRoundId]);

  const isRoundPreviewMode = selectedRoundId !== 'final';

  const studentsById = useMemo(() => {
    const map = new Map<string, StandardField>();
    currentStudents.forEach((student, index) => {
      const inferredId = getStudentId(student, index);
      map.set(inferredId, student);
      if (student.studentId && student.studentId.trim().length > 0) {
        map.set(student.studentId, student);
      }
    });
    return map;
  }, [currentStudents]);

  const studentsByNameClass = useMemo(() => {
    const map = new Map<string, StandardField>();
    currentStudents.forEach((student) => {
      const key = `${(student.navn || '').trim().toLocaleLowerCase('nb')}|${(student.klasse || '').trim().toLocaleLowerCase('nb')}`;
      if (!map.has(key)) {
        map.set(key, student);
      }
    });
    return map;
  }, [currentStudents]);

  const groupedChanges = useMemo(() => {
    const byStudentId = new Map<string, GroupedStudentChange>();

    effectiveChangeLog.forEach((entry) => {
      const existing = byStudentId.get(entry.studentId);
      if (existing) {
        existing.changes.push(entry);
        return;
      }

      byStudentId.set(entry.studentId, {
        studentId: entry.studentId,
        navn: entry.navn || 'Ukjent',
        klasse: entry.klasse || 'Ingen klasse',
        changes: [entry],
      });
    });

    return Array.from(byStudentId.values())
      .map((group) => ({
        ...group,
        changes: [...group.changes].sort((left, right) => {
          return new Date(right.changedAt).getTime() - new Date(left.changedAt).getTime();
        }),
      }))
      .sort((left, right) => {
        const nameCompare = left.navn.localeCompare(right.navn, 'nb', { sensitivity: 'base' });
        if (nameCompare !== 0) {
          return nameCompare;
        }

        return left.klasse.localeCompare(right.klasse, 'nb', { sensitivity: 'base' });
      });
  }, [effectiveChangeLog]);

  const groupedSummaries = useMemo(() => {
    return groupedChanges.map((group) => {
      const oldestFirst = [...group.changes].sort((left, right) => {
        return new Date(left.changedAt).getTime() - new Date(right.changedAt).getTime();
      });

      const assignmentChanges = oldestFirst.filter((entry) => !isStudentStatusChange(entry));
      const studentStatusDetailed = oldestFirst
        .filter((entry) => isStudentStatusChange(entry))
        .map((entry) => {
          const action = resolveStatusAction(entry);
          if (!action) {
            return null;
          }

          return {
            action,
            changedAt: entry.changedAt,
            reason: entry.reason,
          } as StudentStatusLogEntry;
        })
        .filter((entry): entry is StudentStatusLogEntry => entry !== null);

      const summaryStatusEntries = (() => {
        const summaryEntriesRaw: Array<StudentStatusLogEntry | null> = [];
        const openRemovedIndexes: number[] = [];

        studentStatusDetailed.forEach((entry) => {
          if (entry.action === 'removed') {
            summaryEntriesRaw.push(entry);
            openRemovedIndexes.push(summaryEntriesRaw.length - 1);
            return;
          }

          if (entry.action === 'readded') {
            if (openRemovedIndexes.length > 0) {
              const cancelledIndex = openRemovedIndexes.pop();
              if (typeof cancelledIndex === 'number') {
                summaryEntriesRaw[cancelledIndex] = null;
              }
              return;
            }

            summaryEntriesRaw.push(entry);
            return;
          }

          summaryEntriesRaw.push(entry);
        });

        return summaryEntriesRaw.filter((entry): entry is StudentStatusLogEntry => entry !== null);
      })();

      const bySubject = new Map<string, SummaryEntry>();

      assignmentChanges.forEach((entry) => {
        const key = entry.subject.trim().toLocaleLowerCase('nb');
        if (!key) {
          return;
        }

        const existing = bySubject.get(key);
        if (!existing) {
          bySubject.set(key, {
            subject: entry.subject,
            fromBlokk: entry.fromBlokk,
            toBlokk: entry.toBlokk,
            lastChangedAt: entry.changedAt,
          });
          return;
        }

        existing.toBlokk = entry.toBlokk;
        existing.lastChangedAt = entry.changedAt;
      });

      const summaryEntries = Array.from(bySubject.values()).sort((left, right) => {
        return left.subject.localeCompare(right.subject, 'nb', { sensitivity: 'base' });
      }).filter((entry) => entry.fromBlokk !== entry.toBlokk);

      return {
        ...group,
        summaryEntries,
        detailedStatusEntries: studentStatusDetailed,
        summaryStatusEntries,
        detailedAssignmentChanges: [...assignmentChanges].sort((left, right) => {
          return new Date(right.changedAt).getTime() - new Date(left.changedAt).getTime();
        }),
      };
    });
  }, [groupedChanges]);

  const visibleChangeGroups = useMemo(() => {
    if (mode === 'summary') {
      return groupedSummaries.filter((group) => group.summaryEntries.length > 0);
    }

    return groupedSummaries.filter((group) => group.detailedAssignmentChanges.length > 0);
  }, [groupedSummaries, mode]);

  const filteredVisibleGroups = useMemo(() => {
    const tokens = parseSearchTokens(searchQuery);
    let result = tokens.length === 0
      ? visibleChangeGroups
      : visibleChangeGroups.filter((group) => {
        const searchableName = `${group.navn} ${group.klasse}`.toLocaleLowerCase('nb');
        const subjectSet = new Set<string>([
          ...group.detailedAssignmentChanges.map((entry) => entry.subject),
          ...group.summaryEntries.map((entry) => entry.subject),
        ]);
        const searchableSubjects = Array.from(subjectSet)
          .join(' ')
          .toLocaleLowerCase('nb');

        return tokens.every((token) => {
          return searchableName.includes(token) || searchableSubjects.includes(token);
        });
      });

    if (sortBy === 'programomrade') {
      const getProgramomrade = (studentId: string, navn: string, klasse: string): string => {
        const student = studentsById.get(studentId)
          || studentsByNameClass.get(`${navn.trim().toLocaleLowerCase('nb')}|${klasse.trim().toLocaleLowerCase('nb')}`);
        return student?.programomrade || '';
      };

      result = [...result].sort((a, b) => {
        const progA = getProgramomrade(a.studentId, a.navn, a.klasse);
        const progB = getProgramomrade(b.studentId, b.navn, b.klasse);
        // Extract leading number for numeric sort (2 before 3)
        const numA = parseInt(progA.match(/(\d+)/)?.[1] || '999', 10);
        const numB = parseInt(progB.match(/(\d+)/)?.[1] || '999', 10);
        if (numA !== numB) return numA - numB;
        const progCompare = progA.localeCompare(progB, 'nb', { sensitivity: 'base' });
        if (progCompare !== 0) return progCompare;
        const nameCompare = a.navn.localeCompare(b.navn, 'nb', { sensitivity: 'base' });
        if (nameCompare !== 0) return nameCompare;
        return a.klasse.localeCompare(b.klasse, 'nb', { sensitivity: 'base' });
      });
    }

    return result;
  }, [searchQuery, sortBy, studentsById, studentsByNameClass, visibleChangeGroups]);

  const filteredStudentStatusEntries = useMemo(() => {
    const tokens = parseSearchTokens(searchQuery);
    const visibleStatusGroups = groupedSummaries.filter((group) => {
      const entries = mode === 'detailed' ? group.detailedStatusEntries : group.summaryStatusEntries;
      if (entries.length === 0) {
        return false;
      }

      if (tokens.length === 0) {
        return true;
      }

      const searchableName = `${group.navn} ${group.klasse}`.toLocaleLowerCase('nb');
      return tokens.every((token) => searchableName.includes(token));
    });

    return visibleStatusGroups.flatMap((group) => {
      const entries = mode === 'detailed' ? group.detailedStatusEntries : group.summaryStatusEntries;
      return entries.map((entry) => ({
        ...entry,
        studentId: group.studentId,
        navn: group.navn,
        klasse: group.klasse,
      } as FlattenedStudentStatusEntry));
    });
  }, [groupedSummaries, mode, searchQuery]);

  const filteredGroupMoves = useMemo(() => {
    if (groupMoveLog.length === 0) {
      return [] as GroupMoveLogEntry[];
    }

    const tokens = parseSearchTokens(searchQuery);

    const relevant = groupMoveLog.filter((entry) => {
      if (tokens.length === 0) {
        return true;
      }
      const searchable = `${entry.subject} ${entry.groupLabel}`.toLocaleLowerCase('nb');
      return tokens.every((token) => searchable.includes(token));
    });

    if (mode === 'detailed') {
      return [...relevant].sort((a, b) =>
        new Date(b.changedAt).getTime() - new Date(a.changedAt).getTime()
      );
    }

    // Summary mode: compute net change per subject+groupLabel
    const netByKey = new Map<string, { subject: string; groupLabel: string; fromBlokk: number; toBlokk: number; lastChangedAt: string; action?: GroupMoveLogEntry['action'] }>();

    // Subject-add/remove entries pass through directly in summary mode
    const subjectActionEntries: GroupMoveLogEntry[] = [];

    // Process in chronological order for correct net computation
    const chronological = [...relevant].sort((a, b) =>
      new Date(a.changedAt).getTime() - new Date(b.changedAt).getTime()
    );

    chronological.forEach((entry) => {
      if (entry.action === 'subject-added' || entry.action === 'subject-removed') {
        subjectActionEntries.push(entry);
        return;
      }
      const key = `${entry.subject.trim().toLocaleLowerCase('nb')}::${entry.groupLabel.trim().toLocaleLowerCase('nb')}`;
      const existing = netByKey.get(key);
      if (!existing) {
        netByKey.set(key, {
          subject: entry.subject,
          groupLabel: entry.groupLabel,
          fromBlokk: entry.fromBlokk,
          toBlokk: entry.toBlokk,
          lastChangedAt: entry.changedAt,
        });
      } else {
        existing.toBlokk = entry.toBlokk;
        existing.lastChangedAt = entry.changedAt;
      }
    });

    const moveEntries = Array.from(netByKey.values())
      .filter((entry) => entry.fromBlokk !== entry.toBlokk)
      .sort((a, b) => a.subject.localeCompare(b.subject, 'nb', { sensitivity: 'base' }))
      .map((entry) => ({
        subject: entry.subject,
        groupLabel: entry.groupLabel,
        fromBlokk: entry.fromBlokk,
        toBlokk: entry.toBlokk,
        changedAt: entry.lastChangedAt,
      } as GroupMoveLogEntry));

    return [
      ...subjectActionEntries.sort((a, b) => new Date(b.changedAt).getTime() - new Date(a.changedAt).getTime()),
      ...moveEntries,
    ];
  }, [groupMoveLog, mode, searchQuery]);

  const excludedSubjectsSet = useMemo(() => {
    return new Set(
      excludedSubjects
        .map((subject) => subject.trim().toLocaleLowerCase('nb'))
        .filter((subject) => subject.length > 0)
    );
  }, [excludedSubjects]);

  const balancingWarnings = useMemo(() => {
    const unresolved: UnresolvedBalanceWarning[] = [];
    const subjectsInData = new Set<string>();
    const candidateBlocksBySubject = new Map<string, number[]>();

    currentStudents.forEach((student) => {
      BLOKK_FIELDS.slice(0, 4).forEach((blokkField) => {
        parseSubjects(student[blokkField] as string | null | undefined).forEach((subject) => {
          if (!excludedSubjectsSet.has(subject.trim().toLocaleLowerCase('nb'))) {
            subjectsInData.add(subject);
          }
        });
      });
    });

    Object.keys(subjectSettingsByName).forEach((subject) => {
      if (!excludedSubjectsSet.has(subject.trim().toLocaleLowerCase('nb'))) {
        subjectsInData.add(subject);
      }
    });

    Array.from(subjectsInData).forEach((subject) => {
      const breakdown: Record<BlokkLabel, number> = {
        'Blokk 1': 0,
        'Blokk 2': 0,
        'Blokk 3': 0,
        'Blokk 4': 0,
      };
      const studentIdsByBlokk: Record<BlokkLabel, string[]> = {
        'Blokk 1': [],
        'Blokk 2': [],
        'Blokk 3': [],
        'Blokk 4': [],
      };

      currentStudents.forEach((student, index) => {
        const studentId = getStudentId(student, index);
        BLOKK_LABELS.forEach((blokkLabel) => {
          const blokkNumber = getBlokkNumber(blokkLabel);
          const subjectsInBlokk = parseSubjects(student[getBlokkKey(blokkNumber)] as string | null | undefined);
          if (subjectsInBlokk.some((value) => isSameSubject(value, subject))) {
            breakdown[blokkLabel] += 1;
            studentIdsByBlokk[blokkLabel].push(studentId);
          }
        });
      });

      const settings = getSettingsForSubject(subjectSettingsByName, subject, breakdown);
      const resolvedByTarget = getResolvedGroupsByTarget(
        settings.groups || [],
        studentIdsByBlokk,
        settings.groupStudentAssignments || {}
      );

      const candidateBlocks = BLOKK_LABELS
        .filter((blokkLabel) => resolvedByTarget[blokkLabel].some((group) => group.enabled))
        .map((blokkLabel) => getBlokkNumber(blokkLabel));
      candidateBlocksBySubject.set(subject.trim().toLocaleLowerCase('nb'), candidateBlocks);
    });

    currentStudents.forEach((student, index) => {
      const studentId = getStudentId(student, index);
      const selectedSubjects = Array.from(
        new Set(
          BLOKK_FIELDS.slice(0, 4)
            .flatMap((blokkField) => parseSubjects(student[blokkField] as string | null | undefined))
            .filter((subject) => !excludedSubjectsSet.has(subject.trim().toLocaleLowerCase('nb')))
        )
      );

      const fixedByBlock = new Map<number, string[]>();

      selectedSubjects.forEach((subject) => {
        const candidates = candidateBlocksBySubject.get(subject.trim().toLocaleLowerCase('nb')) || [];
        if (candidates.length !== 1) {
          return;
        }

        const onlyBlock = candidates[0];
        fixedByBlock.set(onlyBlock, [...(fixedByBlock.get(onlyBlock) || []), subject]);
      });

      const collisionDetails = Array.from(fixedByBlock.entries())
        .filter(([, subjects]) => subjects.length > 1)
        .map(([block, subjects]) => `Blokk ${block}: ${subjects.join(', ')}`);

      if (collisionDetails.length === 0) {
        return;
      }

      unresolved.push({
        key: `${studentId}::final-unresolved`,
        studentId,
        navn: student.navn || 'Ukjent',
        klasse: student.klasse || 'Ingen klasse',
        reason: `Kan ikke balansere uten kollisjon med dagens grupper: ${collisionDetails.join(' | ')}`,
      });
    });

    return {
      unresolved: unresolved.sort((left, right) => {
        const nameCompare = left.navn.localeCompare(right.navn, 'nb', { sensitivity: 'base' });
        if (nameCompare !== 0) {
          return nameCompare;
        }

        return left.klasse.localeCompare(right.klasse, 'nb', { sensitivity: 'base' });
      }),
    };
  }, [currentStudents, excludedSubjectsSet, studentsById, subjectSettingsByName]);

  const hasBalancingWarnings = balancingWarnings.unresolved.length > 0;

  const finalAllocationByStudentId = useMemo(() => {
    const map = new Map<string, string>();

    groupedSummaries.forEach((group) => {
      const studentKey = `${group.navn.trim().toLocaleLowerCase('nb')}|${group.klasse.trim().toLocaleLowerCase('nb')}`;
      const currentStudent = studentsById.get(group.studentId) || studentsByNameClass.get(studentKey);
      map.set(group.studentId, formatFinalAllocation(currentStudent));
    });

    return map;
  }, [groupedSummaries, studentsById, studentsByNameClass]);

  const formatBlokk = (value: number) => {
    if (value <= 0) {
      return 'ingen blokk';
    }

    return `Blokk ${value}`;
  };

  const renderSummaryHtml = (entry: SummaryEntry): string => {
    if (entry.fromBlokk <= 0 && entry.toBlokk > 0) {
      return `<strong>${escapeHtml(entry.subject)}</strong>: lagt til i ${escapeHtml(formatBlokk(entry.toBlokk))}`;
    }

    if (entry.fromBlokk > 0 && entry.toBlokk <= 0) {
      return `<strong>${escapeHtml(entry.subject)}</strong>: fjernet fra ${escapeHtml(formatBlokk(entry.fromBlokk))}`;
    }

    if (entry.fromBlokk === entry.toBlokk) {
      return `<strong>${escapeHtml(entry.subject)}</strong>: ingen netto endring (${escapeHtml(formatBlokk(entry.toBlokk))})`;
    }

    return `<strong>${escapeHtml(entry.subject)}</strong>: flyttet fra ${escapeHtml(formatBlokk(entry.fromBlokk))} til ${escapeHtml(formatBlokk(entry.toBlokk))}`;
  };

  const renderSummaryContent = (entry: SummaryEntry) => {
    if (entry.fromBlokk <= 0 && entry.toBlokk > 0) {
      return (
        <>
          <strong>{entry.subject}</strong>: lagt til i {formatBlokk(entry.toBlokk)}
        </>
      );
    }

    if (entry.fromBlokk > 0 && entry.toBlokk <= 0) {
      return (
        <>
          <strong>{entry.subject}</strong>: fjernet fra {formatBlokk(entry.fromBlokk)}
        </>
      );
    }

    if (entry.fromBlokk === entry.toBlokk) {
      return (
        <>
          <strong>{entry.subject}</strong>: ingen netto endring ({formatBlokk(entry.toBlokk)})
        </>
      );
    }

    return (
      <>
        <strong>{entry.subject}</strong>: flyttet fra {formatBlokk(entry.fromBlokk)} til {formatBlokk(entry.toBlokk)}
      </>
    );
  };

  const handleExportToExcel = async () => {
    // Only export students with actual summary changes (fromBlokk !== toBlokk)
    const studentsWithChanges = groupedSummaries.filter((group) => group.summaryEntries.length > 0);
    if (studentsWithChanges.length === 0) {
      return;
    }

    try {
      const XLSX = await loadXlsx();

      // Determine max blokk count from the data (use 4 as standard)
      const maxBlokk = 4;

      // Build header row: Navn | Programområde | Blokk 1 | Blokk 2 | ...
      const headerRow: string[] = ['Navn', 'Programområde'];
      for (let b = 1; b <= maxBlokk; b++) {
        headerRow.push(`Blokk ${b}`);
      }

      const dataRows: string[][] = [headerRow];

      studentsWithChanges.forEach((group) => {
        const studentKey = `${group.navn.trim().toLocaleLowerCase('nb')}|${group.klasse.trim().toLocaleLowerCase('nb')}`;
        const currentStudent = studentsById.get(group.studentId) || studentsByNameClass.get(studentKey);
        const programomrade = currentStudent?.programomrade || '';

        // Build a map: for each blokk, what subject is there now and what changed
        const blokkSubjects: string[] = new Array(maxBlokk).fill('');
        const blokkChanges: string[] = new Array(maxBlokk).fill('');

        // Fill final subjects from current student data
        for (let b = 1; b <= maxBlokk; b++) {
          const blokkField = `blokk${b}` as BlokkField;
          const subjects = currentStudent ? parseSubjects(currentStudent[blokkField] as string | null | undefined) : [];
          blokkSubjects[b - 1] = subjects.join(', ');
        }

        // Fill change info from summary entries
        group.summaryEntries.forEach((entry) => {
          if (entry.toBlokk > 0 && entry.toBlokk <= maxBlokk) {
            if (entry.fromBlokk > 0) {
              const changeText = `(${entry.fromBlokk} til ${entry.toBlokk})`;
              blokkChanges[entry.toBlokk - 1] = blokkChanges[entry.toBlokk - 1]
                ? `${blokkChanges[entry.toBlokk - 1]} ${changeText}`
                : changeText;
            } else {
              const changeText = '(lagt til)';
              blokkChanges[entry.toBlokk - 1] = blokkChanges[entry.toBlokk - 1]
                ? `${blokkChanges[entry.toBlokk - 1]} ${changeText}`
                : changeText;
            }
          } else if (entry.fromBlokk > 0 && entry.toBlokk <= 0 && entry.fromBlokk <= maxBlokk) {
            const changeText = '(fjernet)';
            blokkChanges[entry.fromBlokk - 1] = blokkChanges[entry.fromBlokk - 1]
              ? `${blokkChanges[entry.fromBlokk - 1]} ${changeText}`
              : changeText;
          }
        });

        const row: string[] = [group.navn, programomrade];
        for (let b = 0; b < maxBlokk; b++) {
          const cell = blokkChanges[b]
            ? `${blokkSubjects[b]} ${blokkChanges[b]}`
            : blokkSubjects[b];
          row.push(cell);
        }
        dataRows.push(row);
      });

      const worksheet = XLSX.utils.aoa_to_sheet(dataRows);

      // Auto-size columns
      const colWidths = headerRow.map((_, colIndex) => {
        let maxLen = 0;
        dataRows.forEach((row) => {
          const cellLen = (row[colIndex] || '').length;
          if (cellLen > maxLen) maxLen = cellLen;
        });
        return { wch: Math.max(maxLen + 2, 10) };
      });
      worksheet['!cols'] = colWidths;

      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Logg');
      XLSX.writeFile(workbook, `logg-endringer-${formatDateForFilename(new Date())}.xlsx`);
    } catch (error) {
      console.error('Kunne ikke eksportere logg til Excel:', error);
      window.alert('Kunne ikke eksportere logg til Excel. Prøv igjen.');
    }
  };

  const handleExportToWord = () => {
    if (filteredVisibleGroups.length === 0 && filteredStudentStatusEntries.length === 0 && filteredGroupMoves.length === 0) {
      return;
    }

    try {
      const generatedAt = new Date();

      const buildStudentHtml = (group: typeof filteredVisibleGroups[number]) => {
        const studentKey = `${group.navn.trim().toLocaleLowerCase('nb')}|${group.klasse.trim().toLocaleLowerCase('nb')}`;
        const currentStudent = studentsById.get(group.studentId) || studentsByNameClass.get(studentKey);
        const finalSelection = formatFinalAllocation(currentStudent);
        const prog = currentStudent?.programomrade || '';

        const changeLines = mode === 'detailed'
          ? group.detailedAssignmentChanges
            .map((entry) => {
              const changeType = getChangeType(entry.fromBlokk, entry.toBlokk);
              return `<tr><td style="${getWordLineStyle(changeType)}padding:5px 8px;border-radius:6px 0 0 6px;"><strong>${escapeHtml(entry.subject)}</strong>: ${escapeHtml(entry.reason)}</td><td style="${getWordLineStyle(changeType)}padding:5px 8px;border-radius:0 6px 6px 0;text-align:right;white-space:nowrap;width:120px;color:#687c98;font-size:8.5pt;">${escapeHtml(formatTimestamp(entry.changedAt))}</td></tr>`;
            })
            .join('')
          : group.summaryEntries
            .map((entry) => {
              const changeType = getChangeType(entry.fromBlokk, entry.toBlokk);
              return `<tr><td style="${getWordLineStyle(changeType)}padding:5px 8px;border-radius:6px 0 0 6px;">${renderSummaryHtml(entry)}</td><td style="${getWordLineStyle(changeType)}padding:5px 8px;border-radius:0 6px 6px 0;text-align:right;white-space:nowrap;width:120px;color:#687c98;font-size:8.5pt;">${escapeHtml(formatTimestamp(entry.lastChangedAt))}</td></tr>`;
            })
            .join('');

        const progLabel = prog ? ` <span style="color:#687c98;font-weight:normal;font-size:9pt;">\u2014 ${escapeHtml(prog)}</span>` : '';

        return `<section class="student"><h3>${escapeHtml(group.navn)} (${escapeHtml(group.klasse)})${progLabel}</h3><p class="student-meta">${escapeHtml(finalSelection || 'Ingen aktive fagvalg registrert')}</p><div class="student-change-block"><table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:separate;border-spacing:0 5px;">${changeLines}</table></div><div class="student-spacer">&nbsp;</div></section>`;
      };

      let studentRows: string;
      if (sortBy === 'programomrade') {
        // Group students by programområde with section headers
        const groupsByProg = new Map<string, typeof filteredVisibleGroups>();
        filteredVisibleGroups.forEach((group) => {
          const studentKey = `${group.navn.trim().toLocaleLowerCase('nb')}|${group.klasse.trim().toLocaleLowerCase('nb')}`;
          const currentStudent = studentsById.get(group.studentId) || studentsByNameClass.get(studentKey);
          const prog = currentStudent?.programomrade || 'Ukjent programområde';
          const existing = groupsByProg.get(prog);
          if (existing) {
            existing.push(group);
          } else {
            groupsByProg.set(prog, [group]);
          }
        });
        const progSections: string[] = [];
        groupsByProg.forEach((groups, prog) => {
          const studentsHtml = groups.map(buildStudentHtml).join('');
          progSections.push(`<h2 style="font-size:14pt;margin:20px 0 8px;padding-bottom:4px;border-bottom:2px solid #c0cfe0;color:#2c3e50;">${escapeHtml(prog)}</h2>${studentsHtml}`);
        });
        studentRows = progSections.join('');
      } else {
        studentRows = filteredVisibleGroups.map(buildStudentHtml).join('');
      }

      const statusRows = filteredStudentStatusEntries
        .map((entry) => {
          const style = entry.action === 'removed'
            ? getWordLineStyle('removed')
            : entry.action === 'added'
              ? getWordLineStyle('added')
              : getWordLineStyle('moved');

          return `<tr><td style="${style}padding:5px 8px;border-radius:6px 0 0 6px;"><strong>${escapeHtml(entry.navn)} (${escapeHtml(entry.klasse)})</strong>: ${escapeHtml(formatStudentStatusLabel(entry))}</td><td style="${style}padding:5px 8px;border-radius:0 6px 6px 0;text-align:right;white-space:nowrap;width:120px;color:#687c98;font-size:8.5pt;">${escapeHtml(formatTimestamp(entry.changedAt))}</td></tr>`;
        })
        .join('');

      const statusSection = statusRows.length > 0
        ? `<section><h2>Elever lagt til / fjernet</h2><table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:separate;border-spacing:0 5px;">${statusRows}</table></section>`
        : '';

      const groupMoveRows = filteredGroupMoves
        .map((entry) => {
          if (entry.action === 'subject-added') {
            const style = getWordLineStyle('added');
            return `<tr><td style="${style}padding:5px 8px;border-radius:6px 0 0 6px;"><strong>${escapeHtml(entry.subject)}</strong>: Fag lagt til</td><td style="${style}padding:5px 8px;border-radius:0 6px 6px 0;text-align:right;white-space:nowrap;width:120px;color:#687c98;font-size:8.5pt;">${escapeHtml(formatTimestamp(entry.changedAt))}</td></tr>`;
          }
          if (entry.action === 'subject-removed') {
            const style = getWordLineStyle('removed');
            return `<tr><td style="${style}padding:5px 8px;border-radius:6px 0 0 6px;"><strong>${escapeHtml(entry.subject)}</strong>: Fag fjernet</td><td style="${style}padding:5px 8px;border-radius:0 6px 6px 0;text-align:right;white-space:nowrap;width:120px;color:#687c98;font-size:8.5pt;">${escapeHtml(formatTimestamp(entry.changedAt))}</td></tr>`;
          }
          const style = getWordLineStyle('moved');
          return `<tr><td style="${style}padding:5px 8px;border-radius:6px 0 0 6px;"><strong>${escapeHtml(entry.subject)}</strong> (${escapeHtml(entry.groupLabel)}): Blokk ${entry.fromBlokk} \u2192 Blokk ${entry.toBlokk}</td><td style="${style}padding:5px 8px;border-radius:0 6px 6px 0;text-align:right;white-space:nowrap;width:120px;color:#687c98;font-size:8.5pt;">${escapeHtml(formatTimestamp(entry.changedAt))}</td></tr>`;
        })
        .join('');

      const groupMoveSection = groupMoveRows.length > 0
        ? `<section><h2>Gruppeendringer</h2><table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:separate;border-spacing:0 5px;">${groupMoveRows}</table></section>`
        : '';

      const title = mode === 'detailed' ? 'Logg (detaljert)' : 'Logg (oppsummert)';
      const studentSectionHeader = sortBy === 'programomrade' ? '' : '<h2>Elevendringer</h2>';
      const htmlDocument = `<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtml(title)}</title><style>body{font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#1f2b3d;margin:24px;}h1{font-size:18pt;margin:0 0 14px;}h2{font-size:13pt;margin:18px 0 8px;padding-bottom:4px;border-bottom:1px solid #d9e3f0;}.intro{margin:0 0 10px;color:#5b6d86;}h3{font-size:13pt;margin:0 0 6px;}.student{padding:8px 0;}.student-meta{margin:0 0 8px;color:#5b6d86;}.student-change-block{margin-top:8px;}.student-spacer{height:24pt;line-height:24pt;font-size:1pt;}</style></head><body><h1>${escapeHtml(title)}</h1><p class="intro">Generert: ${escapeHtml(formatTimestamp(generatedAt.toISOString()))}</p>${groupMoveSection}${statusSection}<section>${studentSectionHeader}${studentRows}</section></body></html>`;

      const blob = new Blob(['\ufeff', htmlDocument], { type: 'application/msword;charset=utf-8' });
      const filename = `logg-${mode}-${formatDateForFilename(generatedAt)}.doc`;
      const url = window.URL.createObjectURL(blob);
      const link = window.document.createElement('a');
      link.href = url;
      link.download = filename;
      window.document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);
    } catch (error) {
      console.error('Kunne ikke eksportere logg til Word:', error);
      window.alert('Kunne ikke eksportere logg til Word. Proev igjen.');
    }
  };

  if (changeLog.length === 0 && groupMoveLog.length === 0 && !hasBalancingWarnings) {
    return <div className={styles.empty}>Ingen endringer registrert enda.</div>;
  }

  return (
    <div className={styles.wrapper}>
      {hasBalancingWarnings && (
        <section className={`${styles.warningPanel} ${isRoundPreviewMode ? styles.previewOutline : ''}`.trim()}>
          <button
            type="button"
            className={styles.warningToggleBtn}
            onClick={() => setWarningsExpanded((prev) => !prev)}
            aria-expanded={warningsExpanded}
          >
            <span>{warningsExpanded ? '▼' : '▶'}</span>
            <span className={styles.warningTitle}>
              Balanseringsvarsler ({balancingWarnings.unresolved.length})
            </span>
          </button>

          {warningsExpanded && (
            <>
              {balancingWarnings.unresolved.length > 0 && (
                <div className={styles.warningSection}>
                  <h5>Elever med uloselige fagkombinasjoner ({balancingWarnings.unresolved.length})</h5>
                  <ul>
                    {balancingWarnings.unresolved.map((warning) => (
                      <li key={warning.key}>
                        {onOpenStudentCard ? (
                          <button
                            type="button"
                            className={styles.warningStudentButton}
                            onClick={() => onOpenStudentCard(warning.studentId)}
                          >
                            {warning.navn} ({warning.klasse})
                          </button>
                        ) : (
                          <strong>{warning.navn} ({warning.klasse})</strong>
                        )}
                        {' '}
                        - {warning.reason}
                      </li>
                    ))}
                  </ul>
                </div>
              )}
            </>
          )}
        </section>
      )}

      <div className={styles.topBar}>
        <div className={styles.topBarControls}>
          <div className={styles.modeToggle}>
            <button
              type="button"
              className={`${styles.modeBtn} ${mode === 'summary' ? styles.modeBtnActive : ''}`.trim()}
              onClick={() => setMode('summary')}
            >
              Oppsummert logg
            </button>
            <button
              type="button"
              className={`${styles.modeBtn} ${mode === 'detailed' ? styles.modeBtnActive : ''}`.trim()}
              onClick={() => setMode('detailed')}
            >
              Detaljert logg
            </button>
          </div>
          <div className={styles.topBarDivider} />
          <div className={styles.modeToggle}>
            <button
              type="button"
              className={`${styles.modeBtn} ${sortBy === 'name' ? styles.modeBtnActive : ''}`.trim()}
              onClick={() => setSortBy('name')}
            >
              Sorter: Navn
            </button>
            <button
              type="button"
              className={`${styles.modeBtn} ${sortBy === 'programomrade' ? styles.modeBtnActive : ''}`.trim()}
              onClick={() => setSortBy('programomrade')}
            >
              Sorter: Programområde
            </button>
          </div>
        </div>
        <div className={styles.topBarActions}>
          <div className={styles.roundPickerWrap}>
            <label htmlFor="log-round-select" className={styles.roundPickerLabel}>Balans.runde:</label>
            <select
              id="log-round-select"
              className={styles.roundPickerSelect}
              value={String(selectedRoundId)}
              onChange={(event) => {
                const value = event.target.value;
                setSelectedRoundId(value === 'final' ? 'final' : Number.parseInt(value, 10));
              }}
            >
              <option value="final">Final (siste resultat)</option>
              {balancingRoundOptions.map((round) => (
                <option key={`round-${round.id}`} value={String(round.id)}>
                  Runde {round.id} ({formatTimestamp(round.timestamp)})
                </option>
              ))}
            </select>
            {isRoundPreviewMode && (
              <span className={styles.roundWarningIcon} title="Viser ikke siste/finale balanseringsrunde" aria-label="Viser ikke siste/finale balanseringsrunde">
                ▲
              </span>
            )}
          </div>
          <div className={styles.exportDropdownWrap} ref={exportDropdownRef}>
            <button
              type="button"
              className={styles.exportBtn}
              onClick={() => setExportDropdownOpen((prev) => !prev)}
            >
              Eksporter ▾
            </button>
            {exportDropdownOpen && (
              <div className={styles.exportDropdownMenu}>
                <button type="button" className={styles.exportDropdownItem} onClick={() => { setExportDropdownOpen(false); handleExportToExcel(); }}>
                  Eksporter til Excel
                </button>
                <button type="button" className={styles.exportDropdownItem} onClick={() => { setExportDropdownOpen(false); handleExportToWord(); }}>
                  Eksporter til Word
                </button>
              </div>
            )}
          </div>
        </div>
      </div>
      <div className={styles.searchRow}>
        <input
          type="search"
          className={styles.searchInput}
          value={searchQuery}
          onChange={(event) => setSearchQuery(event.target.value)}
          placeholder="Søk elevnavn eller fag"
          aria-label="Sok i logg"
        />
      </div>

      <div className={isRoundPreviewMode ? styles.previewOutline : ''}>
      {filteredGroupMoves.length > 0 && (
        <section className={styles.groupMovePanel}>
          <button
            type="button"
            className={styles.statusToggleButton}
            onClick={() => setGroupMovesExpanded((prev) => !prev)}
            aria-expanded={groupMovesExpanded}
          >
            <span>{groupMovesExpanded ? '▼' : '▶'}</span>
            <span className={styles.groupMovePanelTitle}>
              Gruppeendringer ({filteredGroupMoves.length})
            </span>
          </button>

          {groupMovesExpanded && (
            <ul className={styles.statusList}>
              {filteredGroupMoves.map((entry, index) => {
                if (entry.action === 'subject-added') {
                  return (
                    <li
                      key={`gm-sa-${entry.subject}-${entry.changedAt}-${index}`}
                      className={`${styles.changeItem} ${styles.changeItemAdded}`.trim()}
                    >
                      <span className={styles.changeTime}>{formatTimestamp(entry.changedAt)}</span>
                      <span><strong>{entry.subject}</strong>: Fag lagt til</span>
                    </li>
                  );
                }
                if (entry.action === 'subject-removed') {
                  return (
                    <li
                      key={`gm-sr-${entry.subject}-${entry.changedAt}-${index}`}
                      className={`${styles.changeItem} ${styles.changeItemRemoved}`.trim()}
                    >
                      <span className={styles.changeTime}>{formatTimestamp(entry.changedAt)}</span>
                      <span><strong>{entry.subject}</strong>: Fag fjernet</span>
                    </li>
                  );
                }
                return (
                  <li
                    key={`gm-${entry.subject}-${entry.groupLabel}-${entry.changedAt}-${index}`}
                    className={`${styles.changeItem} ${styles.changeItemGroupMove}`.trim()}
                  >
                    <span className={styles.changeTime}>{formatTimestamp(entry.changedAt)}</span>
                    <span>
                      <strong>{entry.subject}</strong> ({entry.groupLabel}): {formatBlokk(entry.fromBlokk)} → {formatBlokk(entry.toBlokk)}
                    </span>
                  </li>
                );
              })}
            </ul>
          )}
        </section>
      )}

      {filteredStudentStatusEntries.length > 0 && (
        <section className={styles.statusPanel}>
          <button
            type="button"
            className={styles.statusToggleButton}
            onClick={() => setStudentStatusExpanded((prev) => !prev)}
            aria-expanded={studentStatusExpanded}
          >
            <span>{studentStatusExpanded ? '▼' : '▶'}</span>
            <span className={styles.statusPanelTitle}>
              Elever lagt til / fjernet ({filteredStudentStatusEntries.length})
            </span>
          </button>

          {studentStatusExpanded && (
            <ul className={styles.statusList}>
              {filteredStudentStatusEntries.map((entry, index) => {
                const itemClass = entry.action === 'removed'
                  ? styles.changeItemRemoved
                  : entry.action === 'added'
                    ? styles.changeItemAdded
                    : styles.changeItemMoved;

                return (
                  <li
                    key={`${entry.studentId}-status-${entry.changedAt}-${index}`}
                    className={`${styles.changeItem} ${itemClass}`.trim()}
                    title={entry.reason}
                  >
                    <span className={styles.changeTime}>{formatTimestamp(entry.changedAt)}</span>
                    <span>
                      {onOpenStudentCard ? (
                        <button
                          type="button"
                          className={styles.studentHeaderBtn}
                          onClick={() => onOpenStudentCard(entry.studentId)}
                        >
                          <strong>{entry.navn} ({entry.klasse})</strong>
                        </button>
                      ) : (
                        <strong>{entry.navn} ({entry.klasse})</strong>
                      )}
                      {': '}
                      {formatStudentStatusLabel(entry)}
                    </span>
                  </li>
                );
              })}
            </ul>
          )}
        </section>
      )}

      {filteredVisibleGroups.length === 0 ? (
        <div className={styles.empty}>
          {filteredStudentStatusEntries.length > 0
            ? 'Ingen fagendringer matcher søket.'
            : 'Ingen elever matcher søket.'}
        </div>
      ) : (
        filteredVisibleGroups.map((group, groupIndex) => {
          const studentKey = `${group.navn.trim().toLocaleLowerCase('nb')}|${group.klasse.trim().toLocaleLowerCase('nb')}`;
          const currentStudent = studentsById.get(group.studentId) || studentsByNameClass.get(studentKey);
          const programomrade = currentStudent?.programomrade || '';

          // Show programområde group heading when sorted by programområde
          let showProgramHeading = false;
          if (sortBy === 'programomrade') {
            if (groupIndex === 0) {
              showProgramHeading = true;
            } else {
              const prevGroup = filteredVisibleGroups[groupIndex - 1];
              const prevKey = `${prevGroup.navn.trim().toLocaleLowerCase('nb')}|${prevGroup.klasse.trim().toLocaleLowerCase('nb')}`;
              const prevStudent = studentsById.get(prevGroup.studentId) || studentsByNameClass.get(prevKey);
              const prevProg = prevStudent?.programomrade || '';
              showProgramHeading = prevProg !== programomrade;
            }
          }

          return (
            <div key={group.studentId}>
              {showProgramHeading && (
                <h3 className={styles.programHeading}>{programomrade || 'Ukjent programområde'}</h3>
              )}
              <section className={styles.studentBlock}>
            <h4 className={styles.studentHeader}>
              {onOpenStudentCard ? (
                <button
                  type="button"
                  className={styles.studentHeaderBtn}
                  onClick={() => onOpenStudentCard(group.studentId)}
                >
                  {group.navn}
                  {isManualStudentId(group.studentId) ? (
                    <span className={styles.manualBadge} title="Manuelt lagt til" aria-label="Manuelt lagt til">
                      +
                    </span>
                  ) : null}
                  {' '}({group.klasse})
                  {programomrade && <span className={styles.programLabel}>{programomrade}</span>}
                  {' '}- {mode === 'detailed' ? group.detailedAssignmentChanges.length : group.summaryEntries.length} endringer
                </button>
              ) : (
                <>
                  {group.navn}
                  {isManualStudentId(group.studentId) ? (
                    <span className={styles.manualBadge} title="Manuelt lagt til" aria-label="Manuelt lagt til">
                      +
                    </span>
                  ) : null}
                  {' '}({group.klasse})
                  {programomrade && <span className={styles.programLabel}>{programomrade}</span>}
                  {' '}- {mode === 'detailed' ? group.detailedAssignmentChanges.length : group.summaryEntries.length} endringer
                </>
              )}
            </h4>

            {mode === 'summary' && (
              <p className={styles.studentAllocation}>
                {finalAllocationByStudentId.get(group.studentId) || 'Ingen aktive fagvalg registrert'}
              </p>
            )}
            <ul className={styles.changeList}>
              {mode === 'detailed'
                ? group.detailedAssignmentChanges.map((entry, index) => {
                  return (
                    <li
                      key={`${group.studentId}-${entry.changedAt}-${index}`}
                      className={`${styles.changeItem} ${getDetailedEntryClass(entry, styles)}`.trim()}
                    >
                      <span className={styles.changeTime}>{formatTimestamp(entry.changedAt)}</span>
                      <span title={entry.reason}>
                        <strong>{entry.subject}</strong>: {formatDetailedEntryLabel(entry)}
                      </span>
                    </li>
                  );
                })
                : group.summaryEntries.map((entry, index) => {
                  const changeType = getChangeType(entry.fromBlokk, entry.toBlokk);
                  return (
                    <li
                      key={`${group.studentId}-${entry.subject}-${index}`}
                      className={`${styles.changeItem} ${getChangeItemClass(changeType, styles)}`.trim()}
                    >
                      <span className={styles.changeTime}>{formatTimestamp(entry.lastChangedAt)}</span>
                      <span>{renderSummaryContent(entry)}</span>
                    </li>
                  );
                })}
            </ul>
          </section>
            </div>
          );
        })
      )}
      </div>
    </div>
  );
};
