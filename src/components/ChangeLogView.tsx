import { useMemo, useState } from 'react';
import type { StandardField, StudentAssignmentChange } from '../utils/excelUtils';
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
  currentStudents,
  subjectSettingsByName,
  excludedSubjects,
  onOpenStudentCard,
}: ChangeLogViewProps) => {
  const [mode, setMode] = useState<LogMode>('summary');
  const [warningsExpanded, setWarningsExpanded] = useState(false);
  const [studentStatusExpanded, setStudentStatusExpanded] = useState(true);
  const [searchQuery, setSearchQuery] = useState('');

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

    changeLog.forEach((entry) => {
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
  }, [changeLog]);

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
    if (tokens.length === 0) {
      return visibleChangeGroups;
    }

    return visibleChangeGroups.filter((group) => {
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
  }, [searchQuery, visibleChangeGroups]);

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

  const handleExportToWord = () => {
    if (filteredVisibleGroups.length === 0) {
      return;
    }

    try {
      const generatedAt = new Date();

      const studentRows = filteredVisibleGroups.map((group) => {
        const studentKey = `${group.navn.trim().toLocaleLowerCase('nb')}|${group.klasse.trim().toLocaleLowerCase('nb')}`;
        const currentStudent = studentsById.get(group.studentId) || studentsByNameClass.get(studentKey);
        const finalSelection = formatFinalAllocation(currentStudent);

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

        return `<section class="student"><h3>${escapeHtml(group.navn)} (${escapeHtml(group.klasse)})</h3><p class="student-meta">${escapeHtml(finalSelection || 'Ingen aktive fagvalg registrert')}</p><div class="student-change-block"><table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:separate;border-spacing:0 5px;">${changeLines}</table></div><div class="student-spacer">&nbsp;</div></section>`;
      }).join('');

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

      const title = mode === 'detailed' ? 'Logg (detaljert)' : 'Logg (oppsummert)';
      const htmlDocument = `<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtml(title)}</title><style>body{font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#1f2b3d;margin:24px;}h1{font-size:18pt;margin:0 0 14px;}h2{font-size:13pt;margin:18px 0 8px;padding-bottom:4px;border-bottom:1px solid #d9e3f0;}.intro{margin:0 0 10px;color:#5b6d86;}h3{font-size:13pt;margin:0 0 6px;}.student{padding:8px 0;}.student-meta{margin:0 0 8px;color:#5b6d86;}.student-change-block{margin-top:8px;}.student-spacer{height:24pt;line-height:24pt;font-size:1pt;}</style></head><body><h1>${escapeHtml(title)}</h1><p class="intro">Generert: ${escapeHtml(formatTimestamp(generatedAt.toISOString()))}</p>${statusSection}<section><h2>Elevendringer</h2>${studentRows}</section></body></html>`;

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

  if (groupedChanges.length === 0 && !hasBalancingWarnings) {
    return <div className={styles.empty}>Ingen endringer registrert enda.</div>;
  }

  return (
    <div className={styles.wrapper}>
      {hasBalancingWarnings && (
        <section className={styles.warningPanel}>
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
          <input
            type="search"
            className={styles.searchInput}
            value={searchQuery}
            onChange={(event) => setSearchQuery(event.target.value)}
            placeholder="Søk elevnavn eller fag"
            aria-label="Sok i logg"
          />
        </div>
        <button type="button" className={styles.exportBtn} onClick={handleExportToWord}>
          Eksporter til Word
        </button>
      </div>

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
        filteredVisibleGroups.map((group) => (
          <section key={group.studentId} className={styles.studentBlock}>
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
                  {' '}({group.klasse}) - {mode === 'detailed' ? group.detailedAssignmentChanges.length : group.summaryEntries.length} endringer
                </button>
              ) : (
                <>
                  {group.navn}
                  {isManualStudentId(group.studentId) ? (
                    <span className={styles.manualBadge} title="Manuelt lagt til" aria-label="Manuelt lagt til">
                      +
                    </span>
                  ) : null}
                  {' '}({group.klasse}) - {mode === 'detailed' ? group.detailedAssignmentChanges.length : group.summaryEntries.length} endringer
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
        ))
      )}
    </div>
  );
};
