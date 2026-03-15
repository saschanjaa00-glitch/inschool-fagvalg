import { useMemo, useState } from 'react';
import type { StandardField, StudentAssignmentChange } from '../utils/excelUtils';
import styles from './ChangeLogView.module.css';

interface ChangeLogViewProps {
  changeLog: StudentAssignmentChange[];
  currentStudents: StandardField[];
  onOpenStudentCard?: (studentId: string) => void;
}

interface GroupedStudentChange {
  studentId: string;
  navn: string;
  klasse: string;
  changes: StudentAssignmentChange[];
}

interface SummaryEntry {
  subject: string;
  fromBlokk: number;
  toBlokk: number;
  lastChangedAt: string;
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

export const ChangeLogView = ({ changeLog, currentStudents, onOpenStudentCard }: ChangeLogViewProps) => {
  const [mode, setMode] = useState<LogMode>('summary');

  const studentsById = useMemo(() => {
    const map = new Map<string, StandardField>();
    currentStudents.forEach((student) => {
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

      const bySubject = new Map<string, SummaryEntry>();

      oldestFirst.forEach((entry) => {
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
      };
    });
  }, [groupedChanges]);

  const visibleGroups = useMemo(() => {
    if (mode === 'summary') {
      return groupedSummaries.filter((group) => group.summaryEntries.length > 0);
    }

    return groupedSummaries;
  }, [groupedSummaries, mode]);

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
    if (visibleGroups.length === 0) {
      return;
    }

    try {
      const generatedAt = new Date();

      const studentRows = visibleGroups.map((group) => {
        const studentKey = `${group.navn.trim().toLocaleLowerCase('nb')}|${group.klasse.trim().toLocaleLowerCase('nb')}`;
        const currentStudent = studentsById.get(group.studentId) || studentsByNameClass.get(studentKey);
        const finalSelection = formatFinalAllocation(currentStudent);

        const changeLines = mode === 'detailed'
          ? group.changes
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

      const title = mode === 'detailed' ? 'Logg (detaljert)' : 'Logg (oppsummert)';
      const htmlDocument = `<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtml(title)}</title><style>body{font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#1f2b3d;margin:24px;}h1{font-size:18pt;margin:0 0 14px;}h2{font-size:13pt;margin:18px 0 8px;padding-bottom:4px;border-bottom:1px solid #d9e3f0;}.intro{margin:0 0 10px;color:#5b6d86;}h3{font-size:13pt;margin:0 0 6px;}.student{padding:8px 0;}.student-meta{margin:0 0 8px;color:#5b6d86;}.student-change-block{margin-top:8px;}.student-spacer{height:24pt;line-height:24pt;font-size:1pt;}</style></head><body><h1>${escapeHtml(title)}</h1><p class="intro">Generert: ${escapeHtml(formatTimestamp(generatedAt.toISOString()))}</p><section><h2>Elevendringer</h2>${studentRows}</section></body></html>`;

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

  if (groupedChanges.length === 0) {
    return <div className={styles.empty}>Ingen endringer registrert enda.</div>;
  }

  if (visibleGroups.length === 0) {
    return <div className={styles.empty}>Ingen oppsummerte endringer registrert.</div>;
  }

  return (
    <div className={styles.wrapper}>
      <div className={styles.topBar}>
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
        <button type="button" className={styles.exportBtn} onClick={handleExportToWord}>
          Eksporter til Word
        </button>
      </div>

      {visibleGroups.map((group) => (
        <section key={group.studentId} className={styles.studentBlock}>
          <h4 className={styles.studentHeader}>
            {onOpenStudentCard ? (
              <button
                type="button"
                className={styles.studentHeaderBtn}
                onClick={() => onOpenStudentCard(group.studentId)}
              >
                {group.navn} ({group.klasse}) - {mode === 'detailed' ? group.changes.length : group.summaryEntries.length} endringer
              </button>
            ) : (
              `${group.navn} (${group.klasse}) - ${mode === 'detailed' ? group.changes.length : group.summaryEntries.length} endringer`
            )}
          </h4>
          {mode === 'summary' && (
            <p className={styles.studentAllocation}>
              {finalAllocationByStudentId.get(group.studentId) || 'Ingen aktive fagvalg registrert'}
            </p>
          )}
          <ul className={styles.changeList}>
            {mode === 'detailed'
              ? group.changes.map((entry, index) => {
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
      ))}
    </div>
  );
};
