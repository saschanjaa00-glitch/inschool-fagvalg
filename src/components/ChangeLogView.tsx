import { useMemo, useState } from 'react';
import type { StudentAssignmentChange } from '../utils/excelUtils';
import styles from './ChangeLogView.module.css';

interface ChangeLogViewProps {
  changeLog: StudentAssignmentChange[];
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

export const ChangeLogView = ({ changeLog, onOpenStudentCard }: ChangeLogViewProps) => {
  const [mode, setMode] = useState<LogMode>('summary');

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
      });

      return {
        ...group,
        summaryEntries,
      };
    });
  }, [groupedChanges]);

  const formatBlokk = (value: number) => {
    if (value <= 0) {
      return 'ingen blokk';
    }

    return `Blokk ${value}`;
  };

  const renderSummaryText = (entry: SummaryEntry): string => {
    if (entry.fromBlokk <= 0 && entry.toBlokk > 0) {
      return `${entry.subject}: lagt til i ${formatBlokk(entry.toBlokk)}`;
    }

    if (entry.fromBlokk > 0 && entry.toBlokk <= 0) {
      return `${entry.subject}: fjernet fra ${formatBlokk(entry.fromBlokk)}`;
    }

    if (entry.fromBlokk === entry.toBlokk) {
      return `${entry.subject}: ingen netto endring (${formatBlokk(entry.toBlokk)})`;
    }

    return `${entry.subject}: flyttet fra ${formatBlokk(entry.fromBlokk)} til ${formatBlokk(entry.toBlokk)}`;
  };

  const handleExportToWord = async () => {
    if (groupedSummaries.length === 0) {
      return;
    }

    try {
      const { Document, Packer, Paragraph, TextRun, HeadingLevel } = await import('docx');
      const generatedAt = new Date();
      const children: InstanceType<typeof Paragraph>[] = [];

      children.push(
        new Paragraph({
          text: mode === 'detailed' ? 'Endringslogg (detaljert)' : 'Endringslogg (oppsummert)',
          heading: HeadingLevel.HEADING_1,
        })
      );
      children.push(
        new Paragraph({
          text: `Generert: ${formatTimestamp(generatedAt.toISOString())}`,
        })
      );
      children.push(new Paragraph({ text: '' }));

      groupedSummaries.forEach((group) => {
        const changeCount = mode === 'detailed' ? group.changes.length : group.summaryEntries.length;

        children.push(
          new Paragraph({
            text: `${group.navn} (${group.klasse}) - ${changeCount} endringer`,
            heading: HeadingLevel.HEADING_2,
          })
        );

        if (mode === 'detailed') {
          group.changes.forEach((entry) => {
            children.push(
              new Paragraph({
                bullet: { level: 0 },
                children: [
                  new TextRun({ text: `${formatTimestamp(entry.changedAt)}: `, bold: true }),
                  new TextRun(entry.reason),
                ],
              })
            );
          });
        } else {
          group.summaryEntries.forEach((entry) => {
            children.push(
              new Paragraph({
                bullet: { level: 0 },
                children: [
                  new TextRun({ text: `${formatTimestamp(entry.lastChangedAt)}: `, bold: true }),
                  new TextRun(renderSummaryText(entry)),
                ],
              })
            );
          });
        }

        children.push(new Paragraph({ text: '' }));
      });

      const doc = new Document({
        sections: [{ children }],
      });

      const blob = await Packer.toBlob(doc);
      const filename = `endringslogg-${mode}-${formatDateForFilename(generatedAt)}.docx`;
      const url = window.URL.createObjectURL(blob);
      const link = window.document.createElement('a');
      link.href = url;
      link.download = filename;
      window.document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);
    } catch (error) {
      console.error('Kunne ikke eksportere endringslogg til Word:', error);
      window.alert('Kunne ikke eksportere endringslogg til Word. Proev igjen.');
    }
  };

  if (groupedChanges.length === 0) {
    return <div className={styles.empty}>Ingen endringer registrert enda.</div>;
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

      {groupedSummaries.map((group) => (
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
          <ul className={styles.changeList}>
            {mode === 'detailed'
              ? group.changes.map((entry, index) => (
                <li key={`${group.studentId}-${entry.changedAt}-${index}`} className={styles.changeItem}>
                  <span className={styles.changeTime}>{formatTimestamp(entry.changedAt)}</span>
                  <span>{entry.reason}</span>
                </li>
              ))
              : group.summaryEntries.map((entry, index) => (
                <li key={`${group.studentId}-${entry.subject}-${index}`} className={styles.changeItem}>
                  <span className={styles.changeTime}>{formatTimestamp(entry.lastChangedAt)}</span>
                  <span>{renderSummaryText(entry)}</span>
                </li>
              ))}
          </ul>
        </section>
      ))}
    </div>
  );
};
