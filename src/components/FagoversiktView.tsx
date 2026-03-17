import { Fragment, useMemo, useState } from 'react';
import type { StandardField } from '../utils/excelUtils';
import {
  BLOKK_LABELS,
  getSettingsForSubject,
  type BlokkLabel,
  type SubjectSettingsByName,
} from '../utils/subjectGroups';
import styles from './FagoversiktView.module.css';

interface FagoversiktViewProps {
  data: StandardField[];
  blokkCount: number;
  subjectSettingsByName: SubjectSettingsByName;
  onOpenStudentCard?: (studentId: string) => void;
}

type BlokkField = `blokk${1 | 2 | 3 | 4 | 5 | 6 | 7 | 8}`;

interface SubjectStudentRow {
  studentId: string;
  navn: string;
  klasse: string;
}

interface SubjectOverviewRow {
  subjectKey: string;
  subject: string;
  blokkNumbers: number[];
  students: SubjectStudentRow[];
  totalMax: number | null;
  overTotalLimit: boolean;
}

const compareText = (left: string, right: string): number => {
  return left.localeCompare(right, 'nb', { sensitivity: 'base', numeric: true });
};

const normalizeSubjectKey = (value: string): string => {
  return value.trim().toLocaleLowerCase('nb');
};

const parseSubjects = (value: string | null): string[] => {
  if (!value) {
    return [];
  }

  return value
    .split(/[,;]/)
    .map((subject) => subject.trim())
    .filter((subject) => subject.length > 0);
};

const getStudentId = (student: StandardField, index: number): string => {
  return student.studentId || `${student.navn || 'ukjent'}:${student.klasse || 'ukjent'}:${index}`;
};

const getBlokkKey = (blokkNumber: number): BlokkField => {
  return `blokk${blokkNumber}` as BlokkField;
};

export const FagoversiktView = ({
  data,
  blokkCount,
  subjectSettingsByName,
  onOpenStudentCard,
}: FagoversiktViewProps) => {
  const [expandedSubjectKey, setExpandedSubjectKey] = useState<string | null>(null);

  const visibleBlokkCount = Math.min(blokkCount, 8);

  const rows = useMemo(() => {
    const subjects = new Map<
      string,
      {
        displaySubject: string;
        blokkNumbers: Set<number>;
        studentsById: Map<string, SubjectStudentRow>;
      }
    >();

    data.forEach((student, index) => {
      const studentId = getStudentId(student, index);
      const studentRow: SubjectStudentRow = {
        studentId,
        navn: student.navn || 'Ukjent',
        klasse: student.klasse || 'Ingen klasse',
      };

      for (let blokkNumber = 1; blokkNumber <= visibleBlokkCount; blokkNumber += 1) {
        const subjectsInBlokk = parseSubjects(student[getBlokkKey(blokkNumber)] as string | null);

        subjectsInBlokk.forEach((subject) => {
          const subjectKey = normalizeSubjectKey(subject);
          const existing = subjects.get(subjectKey);

          if (!existing) {
            subjects.set(subjectKey, {
              displaySubject: subject,
              blokkNumbers: new Set([blokkNumber]),
              studentsById: new Map([[studentId, studentRow]]),
            });
            return;
          }

          existing.blokkNumbers.add(blokkNumber);
          existing.studentsById.set(studentId, studentRow);
        });
      }
    });

    return Array.from(subjects.entries())
      .map(([subjectKey, value]) => {
        const breakdown: Record<BlokkLabel, number> = {
          'Blokk 1': 0,
          'Blokk 2': 0,
          'Blokk 3': 0,
          'Blokk 4': 0,
        };

        value.blokkNumbers.forEach((blokkNumber) => {
          const label = `Blokk ${blokkNumber}` as BlokkLabel;
          if (BLOKK_LABELS.includes(label)) {
            breakdown[label] = 1;
          }
        });

        const settings = getSettingsForSubject(subjectSettingsByName, value.displaySubject, breakdown);
        const enabledGroups = (settings.groups || []).filter((group) => group.enabled !== false);
        const totalMax = enabledGroups.length > 0
          ? enabledGroups.reduce((sum, group) => sum + Math.max(0, group.max), 0)
          : null;

        const students = Array.from(value.studentsById.values()).sort((left, right) => {
          const byName = compareText(left.navn, right.navn);
          if (byName !== 0) {
            return byName;
          }

          return compareText(left.klasse, right.klasse);
        });

        return {
          subjectKey,
          subject: value.displaySubject,
          blokkNumbers: Array.from(value.blokkNumbers).sort((left, right) => left - right),
          students,
          totalMax,
          overTotalLimit: totalMax !== null && students.length > totalMax,
        } as SubjectOverviewRow;
      })
      .sort((left, right) => compareText(left.subject, right.subject));
  }, [data, subjectSettingsByName, visibleBlokkCount]);

  if (rows.length === 0) {
    return <div className={styles.empty}>Ingen fag tilgjengelig.</div>;
  }

  return (
    <div className={styles.wrapper}>
      <div className={styles.headerRow}>
        <div>
          <h3 className={styles.title}>Fagoversikt</h3>
          <p className={styles.subtitle}>Fag sortert alfabetisk med blokker og elevliste per fag.</p>
        </div>
        <div className={styles.summaryBadge}>{rows.length} fag</div>
      </div>

      <table className={styles.table}>
        <thead>
          <tr>
            <th>Fag</th>
            <th>Blokker</th>
            <th>Elever</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => {
            const isExpanded = expandedSubjectKey === row.subjectKey;

            return (
              <Fragment key={row.subjectKey}>
                <tr
                  className={[
                    styles.subjectRow,
                    isExpanded ? styles.subjectRowExpanded : '',
                    row.overTotalLimit ? styles.subjectRowOverLimit : '',
                  ].filter(Boolean).join(' ')}
                  onClick={() => setExpandedSubjectKey((prev) => (prev === row.subjectKey ? null : row.subjectKey))}
                >
                  <td className={styles.subjectCell}>
                    <span>{row.subject}</span>
                    {row.overTotalLimit && <span className={styles.overLimitBadge}>Over maks</span>}
                  </td>
                  <td>{row.blokkNumbers.map((blokk) => `Blokk ${blokk}`).join(', ')}</td>
                  <td>
                    {row.students.length}
                    {row.totalMax !== null && <span className={styles.limitMeta}> / {row.totalMax}</span>}
                  </td>
                </tr>
                {isExpanded && (
                  <tr className={styles.detailRow}>
                    <td colSpan={3}>
                      <div className={styles.detailPanel}>
                        <h4 className={styles.studentTitle}>Elever i {row.subject}</h4>
                        <ul className={styles.studentList}>
                          {row.students.map((student) => (
                            <li key={student.studentId}>
                              <button
                                type="button"
                                className={styles.studentButton}
                                onClick={(event) => {
                                  event.stopPropagation();
                                  onOpenStudentCard?.(student.studentId);
                                }}
                              >
                                <span className={styles.studentName}>{student.navn}</span>
                                <span className={styles.studentClass}>{student.klasse}</span>
                              </button>
                            </li>
                          ))}
                        </ul>
                      </div>
                    </td>
                  </tr>
                )}
              </Fragment>
            );
          })}
        </tbody>
      </table>
    </div>
  );
};
