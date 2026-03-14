import { useEffect, useState } from 'react';
import type { ParsedFile, ColumnMapping, StandardField, SubjectCount, StudentAssignmentChange } from './utils/excelUtils';
import {
  mergeFiles,
  tallySubjects,
  autoDetectMapping,
  exportToExcel,
  exportToTabText,
  loadXlsx,
  moveSubjectAssignmentsBetweenBlokker,
  swapSubjectAssignmentsBetweenBlokker,
  exportToExcelDetailed,
  removeSubjectAssignmentsForStudents,
} from './utils/excelUtils';
import './App.css';
import { FileUploader } from './components/FileUploader';
import { ColumnMapper } from './components/ColumnMapper';
import { MergedDataView } from './components/MergedDataView';
import { SubjectTally } from './components/SubjectTally';
import { EleverView } from './components/EleverView';
import type { SubjectSettingsByName } from './components/SubjectTally';

const LOCAL_STORAGE_KEY = 'fagvalg-opptelling-state-v1';

interface PersistedAppState {
  parsedFiles: ParsedFile[];
  mappings: Record<string, ColumnMapping>;
  mergedData: StandardField[];
  subjects: SubjectCount[];
  studentAssignmentChanges?: StudentAssignmentChange[];
  subjectSettingsByName: SubjectSettingsByName;
  warningIgnoresByStudentId?: Record<string, { comment: string; ignoredAt: string }>;
  warningIgnoresByStudentAndType?: Record<string, Partial<Record<WarningType, WarningIgnoreEntry>>>;
  subjectMaxByName?: Record<string, number>;
  blokkCount: number;
  selectedMergedSubject: string;
}

type WarningType = 'missing' | 'overloaded';

interface WarningIgnoreEntry {
  comment: string;
  ignoredAt: string;
}

function App() {
  const [parsedFiles, setParsedFiles] = useState<ParsedFile[]>([]);
  const [mappings, setMappings] = useState<Map<string, ColumnMapping>>(new Map());
  const [mergedData, setMergedData] = useState<StandardField[]>([]);
  const [subjects, setSubjects] = useState<SubjectCount[]>([]);
  const [studentAssignmentChanges, setStudentAssignmentChanges] = useState<StudentAssignmentChange[]>([]);
  const [subjectSettingsByName, setSubjectSettingsByName] = useState<SubjectSettingsByName>({});
  const [blokkCount, setBlokkCount] = useState(4);
  
  const [columnMapperExpanded, setColumnMapperExpanded] = useState(false);
  const [activeDataTab, setActiveDataTab] = useState<'import' | 'subjects' | 'students' | 'elever'>('import');
  const [warningExpanded, setWarningExpanded] = useState(false);
  const [warningBlokkCollisionExpanded, setWarningBlokkCollisionExpanded] = useState(false);
  const [warningFewSubjectsExpanded, setWarningFewSubjectsExpanded] = useState(false);
  const [warningFourSubjectsExpanded, setWarningFourSubjectsExpanded] = useState(false);
  const [warningCopyStatus, setWarningCopyStatus] = useState('');
  const [warningIgnoresByStudentAndType, setWarningIgnoresByStudentAndType] = useState<
    Record<string, Partial<Record<WarningType, WarningIgnoreEntry>>>
  >({});
  const [warningIgnoreDraftByStudentId, setWarningIgnoreDraftByStudentId] = useState<Record<string, string>>({});
  const [selectedEleverStudentId, setSelectedEleverStudentId] = useState('');
  const [selectedMergedSubject, setSelectedMergedSubject] = useState('');
  const [isHydratedFromStorage, setIsHydratedFromStorage] = useState(false);

  useEffect(() => {
    try {
      const savedState = localStorage.getItem(LOCAL_STORAGE_KEY);
      if (!savedState) {
        return;
      }

      const parsedState = JSON.parse(savedState) as Partial<PersistedAppState>;

      if (Array.isArray(parsedState.parsedFiles)) {
        setParsedFiles(parsedState.parsedFiles);
      }

      if (parsedState.mappings && typeof parsedState.mappings === 'object') {
        setMappings(new Map(Object.entries(parsedState.mappings)));
      }

      if (Array.isArray(parsedState.mergedData)) {
        setMergedData(parsedState.mergedData);
      }

      if (Array.isArray(parsedState.subjects)) {
        setSubjects(parsedState.subjects);
      }

      if (Array.isArray(parsedState.studentAssignmentChanges)) {
        setStudentAssignmentChanges(parsedState.studentAssignmentChanges);
      }

      if (parsedState.subjectSettingsByName && typeof parsedState.subjectSettingsByName === 'object') {
        setSubjectSettingsByName(parsedState.subjectSettingsByName);
      } else if (parsedState.subjectMaxByName && typeof parsedState.subjectMaxByName === 'object') {
        // Backward compatibility for persisted v1 shape.
        const migrated: SubjectSettingsByName = Object.fromEntries(
          Object.entries(parsedState.subjectMaxByName).map(([subject, max]) => [
            subject,
            {
              defaultMax: typeof max === 'number' ? max : 30,
              blokkMaxOverrides: {},
              blokkEnabled: {
                'Blokk 1': true,
                'Blokk 2': true,
                'Blokk 3': true,
                'Blokk 4': true,
              },
              blokkOrder: ['Blokk 1', 'Blokk 2', 'Blokk 3', 'Blokk 4'],
            },
          ])
        ) as SubjectSettingsByName;
        setSubjectSettingsByName(migrated);
      }

      if (parsedState.warningIgnoresByStudentAndType && typeof parsedState.warningIgnoresByStudentAndType === 'object') {
        setWarningIgnoresByStudentAndType(parsedState.warningIgnoresByStudentAndType);
      } else if (parsedState.warningIgnoresByStudentId && typeof parsedState.warningIgnoresByStudentId === 'object') {
        // Backward compatibility for warning ignores stored without type.
        const migrated = Object.fromEntries(
          Object.entries(parsedState.warningIgnoresByStudentId).map(([studentId, value]) => [
            studentId,
            {
              missing: value,
            },
          ])
        ) as Record<string, Partial<Record<WarningType, WarningIgnoreEntry>>>;
        setWarningIgnoresByStudentAndType(migrated);
      }

      if (typeof parsedState.blokkCount === 'number') {
        setBlokkCount(parsedState.blokkCount);
      }

      if (typeof parsedState.selectedMergedSubject === 'string') {
        setSelectedMergedSubject(parsedState.selectedMergedSubject);
      }
    } catch {
      // Ignore malformed localStorage data and continue with fresh state.
    } finally {
      setIsHydratedFromStorage(true);
    }
  }, []);

  useEffect(() => {
    if (!isHydratedFromStorage) {
      return;
    }

    const persistedState: PersistedAppState = {
      parsedFiles,
      mappings: Object.fromEntries(mappings.entries()),
      mergedData,
      subjects,
      studentAssignmentChanges,
      subjectSettingsByName,
      warningIgnoresByStudentAndType,
      blokkCount,
      selectedMergedSubject,
    };

    localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(persistedState));
  }, [
    isHydratedFromStorage,
    parsedFiles,
    mappings,
    mergedData,
    subjects,
    studentAssignmentChanges,
    subjectSettingsByName,
    warningIgnoresByStudentAndType,
    blokkCount,
    selectedMergedSubject,
  ]);

  const handleFilesAdded = (files: ParsedFile[]) => {
    setParsedFiles((prev) => [...prev, ...files]);
    
    // Auto-detect and apply mappings for new files
    const newMappings = new Map(mappings);
    files.forEach((file) => {
      const autoMapping = autoDetectMapping(file.columns, blokkCount, file.data);
      newMappings.set(file.id, autoMapping);
    });
    setMappings(newMappings);
  };

  const handleMappingChange = (fileId: string, mapping: ColumnMapping) => {
    const newMappings = new Map(mappings);
    newMappings.set(fileId, mapping);
    setMappings(newMappings);
  };

  const handleMerge = () => {
    const merged = mergeFiles(parsedFiles, mappings).sort((a, b) => {
      const classA = (a.klasse || '').trim();
      const classB = (b.klasse || '').trim();
      const classCompare = classA.localeCompare(classB, 'nb', { sensitivity: 'base' });

      if (classCompare !== 0) {
        return classCompare;
      }

      const nameA = (a.navn || '').trim();
      const nameB = (b.navn || '').trim();
      return nameA.localeCompare(nameB, 'nb', { sensitivity: 'base' });
    });

    setMergedData(merged);
    setSubjects(tallySubjects(merged));
    setStudentAssignmentChanges([]);
    setActiveDataTab('subjects');
  };

  const handleReset = () => {
    setParsedFiles([]);
    setMappings(new Map());
    setMergedData([]);
    setSubjects([]);
    setStudentAssignmentChanges([]);
    setSubjectSettingsByName({});
    setWarningIgnoresByStudentAndType({});
    setWarningIgnoreDraftByStudentId({});
    setSelectedEleverStudentId('');
    setSelectedMergedSubject('');
  };

  const handleApplySubjectBlockMoves = (
    subject: string,
    operations: Array<
      | { type: 'move'; fromBlokk: number; toBlokk: number; reason: string }
      | { type: 'swap'; blokkA: number; blokkB: number; reason: string }
    >
  ) => {
    if (operations.length === 0) {
      return;
    }

    let workingData = mergedData;
    let allChanges: StudentAssignmentChange[] = [];

    operations.forEach((operation) => {
      const result = operation.type === 'swap'
        ? swapSubjectAssignmentsBetweenBlokker(
          workingData,
          subject,
          operation.blokkA,
          operation.blokkB,
          operation.reason
        )
        : moveSubjectAssignmentsBetweenBlokker(
          workingData,
          subject,
          operation.fromBlokk,
          operation.toBlokk,
          operation.reason
        );

      workingData = result.updatedData;
      allChanges = [...allChanges, ...result.changes];
    });

    if (allChanges.length === 0) {
      return;
    }

    setMergedData(workingData);
    setSubjects(tallySubjects(workingData));
    setStudentAssignmentChanges((prev) => [...prev, ...allChanges]);
  };

  const handleClearStoredData = () => {
    localStorage.removeItem(LOCAL_STORAGE_KEY);
    handleReset();
    setBlokkCount(4);
  };

  const handleRemoveFile = (fileId: string) => {
    setParsedFiles((prev) => prev.filter((f) => f.id !== fileId));
    const newMappings = new Map(mappings);
    newMappings.delete(fileId);
    setMappings(newMappings);
  };

  const handleExport = async () => {
    await exportToExcel(mergedData, 'merged_students.xlsx');
  };

  const handleExportDetailed = async () => {
    await exportToExcelDetailed(mergedData, blokkCount, 'merged_students_full.xlsx');
  };

  const handleExportText = () => {
    exportToTabText(mergedData, 'merged_students.txt');
  };

  const handleStudentAssignmentsUpdated = (
    updatedData: StandardField[],
    changes: StudentAssignmentChange[]
  ) => {
    if (changes.length === 0) {
      return;
    }

    setMergedData(updatedData);
    setSubjects(tallySubjects(updatedData));
    setStudentAssignmentChanges((prev) => [...prev, ...changes]);
  };

  const handleRemoveStudentsFromSubject = (
    subject: string,
    studentIds: string[],
    reason: string
  ) => {
    const result = removeSubjectAssignmentsForStudents(mergedData, subject, studentIds, reason);
    if (result.changes.length === 0) {
      return;
    }

    setMergedData(result.updatedData);
    setSubjects(tallySubjects(result.updatedData));
    setStudentAssignmentChanges((prev) => [...prev, ...result.changes]);
  };

  const getWarningStudentId = (student: StandardField, indexHint?: number): string => {
    if (student.studentId && student.studentId.trim().length > 0) {
      return student.studentId;
    }

    if (typeof indexHint === 'number') {
      return `${student.navn || 'ukjent'}:${student.klasse || 'ukjent'}:${indexHint}`;
    }

    const index = mergedData.indexOf(student);
    if (index >= 0) {
      return `${student.navn || 'ukjent'}:${student.klasse || 'ukjent'}:${index}`;
    }

    return `${student.navn || 'ukjent'}:${student.klasse || 'ukjent'}`;
  };

  const isWarningIgnored = (studentId: string, type: WarningType): boolean => {
    return !!warningIgnoresByStudentAndType[studentId]?.[type];
  };

  const parseWarningSubjects = (value: string | null): string[] => {
    if (!value) {
      return [];
    }

    return value
      .split(/[,;]/)
      .map((subject) => subject.trim())
      .filter((subject) => subject.length > 0);
  };

  const getWarningSubjects = (student: StandardField): string[] => {
    return [student.blokk1, student.blokk2, student.blokk3, student.blokk4]
      .flatMap((blokk) => parseWarningSubjects(blokk));
  };

  const getWarningSubjectsInBlokk = (student: StandardField, blokkNumber: number): string[] => {
    const blokkKey = `blokk${blokkNumber}` as keyof StandardField;
    const value = student[blokkKey];
    return parseWarningSubjects(typeof value === 'string' ? value : null);
  };

  const warningEntries = mergedData.map((student, index) => {
    const subjectCount = getWarningSubjects(student).length;
    const collisionBlokker = [1, 2, 3, 4].filter((blokkNumber) => getWarningSubjectsInBlokk(student, blokkNumber).length > 1);

    return {
      student,
      studentId: getWarningStudentId(student, index),
      subjectCount,
      hasBlokkCollision: collisionBlokker.length > 0,
      collisionDetails: collisionBlokker.map((blokkNumber) => {
        const subjects = getWarningSubjectsInBlokk(student, blokkNumber);
        return `Blokk ${blokkNumber}: ${subjects.join(', ')}`;
      }),
    };
  });

  // Get students with less than 3 blokkfag
  const getStudentsWithFewSubjects = () => {
    return warningEntries
      .filter((entry) => {
        return entry.subjectCount < 3;
      })
      .sort((a, b) => (a.student.navn || '').localeCompare(b.student.navn || '', 'nb', { sensitivity: 'base' }));
  };

  const studentsWithFewSubjects = getStudentsWithFewSubjects();

  const activeStudentsWithFewSubjects = studentsWithFewSubjects.filter(
    (entry) => !isWarningIgnored(entry.studentId, 'missing')
  );

  const ignoredStudentsWithFewSubjects = studentsWithFewSubjects.filter(
    (entry) => isWarningIgnored(entry.studentId, 'missing')
  );

  const getStudentsWithFourSubjects = () => {
    return warningEntries.filter((entry) => {
      return entry.subjectCount >= 4;
    })
      .sort((a, b) => (a.student.navn || '').localeCompare(b.student.navn || '', 'nb', { sensitivity: 'base' }));
  };

  const studentsWithFourSubjects = getStudentsWithFourSubjects();

  const activeStudentsWithFourSubjects = studentsWithFourSubjects.filter(
    (entry) => !isWarningIgnored(entry.studentId, 'overloaded')
  );

  const ignoredStudentsWithFourSubjects = studentsWithFourSubjects.filter(
    (entry) => isWarningIgnored(entry.studentId, 'overloaded')
  );

  const getStudentsWithBlokkCollisions = () => {
    return warningEntries
      .filter((entry) => entry.hasBlokkCollision)
      .sort((a, b) => (a.student.navn || '').localeCompare(b.student.navn || '', 'nb', { sensitivity: 'base' }));
  };

  const studentsWithBlokkCollisions = getStudentsWithBlokkCollisions();
  const activeStudentsWithBlokkCollisions = studentsWithBlokkCollisions;

  const fewSubjectsIgnoredCount = ignoredStudentsWithFewSubjects.length;

  const fourSubjectsIgnoredCount = ignoredStudentsWithFourSubjects.length;
  const hasActiveWarnings = activeStudentsWithFewSubjects.length > 0
    || activeStudentsWithFourSubjects.length > 0
    || activeStudentsWithBlokkCollisions.length > 0;

  const warningStudentIds = new Set<string>([
    ...studentsWithFewSubjects.map((entry) => entry.studentId),
    ...studentsWithFourSubjects.map((entry) => entry.studentId),
  ]);

  useEffect(() => {
    setWarningIgnoresByStudentAndType((prev) => {
      const next = Object.fromEntries(
        Object.entries(prev)
          .filter(([studentId]) => warningStudentIds.has(studentId))
          .map(([studentId, value]) => {
            const filtered: Partial<Record<WarningType, WarningIgnoreEntry>> = {};
            if (value.missing && studentsWithFewSubjects.some((entry) => entry.studentId === studentId)) {
              filtered.missing = value.missing;
            }
            if (value.overloaded && studentsWithFourSubjects.some((entry) => entry.studentId === studentId)) {
              filtered.overloaded = value.overloaded;
            }
            return [studentId, filtered];
          })
          .filter(([, value]) => Object.keys(value).length > 0)
      ) as Record<string, Partial<Record<WarningType, WarningIgnoreEntry>>>;

      return Object.keys(next).length === Object.keys(prev).length ? prev : next;
    });

    setWarningIgnoreDraftByStudentId((prev) => {
      const next = Object.fromEntries(
        Object.entries(prev).filter(([studentId]) => warningStudentIds.has(studentId))
      ) as Record<string, string>;

      return Object.keys(next).length === Object.keys(prev).length ? prev : next;
    });
  }, [studentsWithFewSubjects, studentsWithFourSubjects]);

  const saveWarningIgnore = (studentId: string, type: WarningType, explicitComment?: string) => {
    const rawComment = explicitComment ?? warningIgnoreDraftByStudentId[studentId] ?? '';
    const comment = rawComment.trim();

    setWarningIgnoresByStudentAndType((prev) => ({
      ...prev,
      [studentId]: {
        ...(prev[studentId] || {}),
        [type]: {
          comment,
          ignoredAt: new Date().toISOString(),
        },
      },
    }));
  };

  const removeWarningIgnore = (studentId: string, type: WarningType) => {
    setWarningIgnoresByStudentAndType((prev) => {
      const next = { ...prev };
      const current = { ...(next[studentId] || {}) };
      delete current[type];
      if (Object.keys(current).length === 0) {
        delete next[studentId];
      } else {
        next[studentId] = current;
      }
      return next;
    });
  };

  const matchesSelectedSubject = (value: string | null, selectedSubject: string) => {
    if (!value || !selectedSubject) {
      return false;
    }

    return value
      .split(/[,;]/)
      .map((subject) => subject.trim())
      .filter(Boolean)
      .some((subject) => subject.localeCompare(selectedSubject, 'nb', { sensitivity: 'base' }) === 0);
  };

  const filteredMergedData = selectedMergedSubject
    ? mergedData.filter((student) => {
      return (
        matchesSelectedSubject(student.blokk1, selectedMergedSubject)
        || matchesSelectedSubject(student.blokk2, selectedMergedSubject)
        || matchesSelectedSubject(student.blokk3, selectedMergedSubject)
        || matchesSelectedSubject(student.blokk4, selectedMergedSubject)
      );
    })
    : mergedData;

  const handleWarningExport = async () => {
    const XLSX = await loadXlsx();
    const warningRows = [
      ...activeStudentsWithFewSubjects.map((entry) => {
        const subjects = getWarningSubjects(entry.student);
        return {
          Kategori: 'Under 3 fag',
          Navn: entry.student.navn || 'Ukjent',
          Klasse: entry.student.klasse || 'Ingen klasse',
          AntallFag: subjects.length,
          Fag: subjects.join(', '),
          Ignorert: 'Nei',
          Kommentar: '',
        };
      }),
      ...activeStudentsWithFourSubjects.map((entry) => {
        const subjects = getWarningSubjects(entry.student);
        return {
          Kategori: '4 fag',
          Navn: entry.student.navn || 'Ukjent',
          Klasse: entry.student.klasse || 'Ingen klasse',
          AntallFag: subjects.length,
          Fag: subjects.join(', '),
          Ignorert: 'Nei',
          Kommentar: '',
        };
      }),
      ...activeStudentsWithBlokkCollisions.map((entry) => {
        const subjects = getWarningSubjects(entry.student);
        return {
          Kategori: 'Blokk-kollisjon',
          Navn: entry.student.navn || 'Ukjent',
          Klasse: entry.student.klasse || 'Ingen klasse',
          AntallFag: subjects.length,
          Fag: subjects.join(', '),
          Ignorert: 'Nei',
          Kommentar: entry.collisionDetails.join(' | '),
        };
      })
    ];

    const worksheet = XLSX.utils.json_to_sheet(warningRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Advarsler');
    XLSX.writeFile(workbook, 'warning_students.xlsx');
  };

  const handleWarningCopy = async () => {
    const fewSubjectsText = activeStudentsWithFewSubjects.map((entry) => {
      const subjects = getWarningSubjects(entry.student);
      return `${entry.student.navn || 'Ukjent'} (${entry.student.klasse || 'Ingen klasse'}) - ${subjects.length} fag: ${subjects.join(', ') || 'Ingen'}`;
    });

    const fourSubjectsText = activeStudentsWithFourSubjects.map((entry) => {
      const subjects = getWarningSubjects(entry.student);
      return `${entry.student.navn || 'Ukjent'} (${entry.student.klasse || 'Ingen klasse'}) - 4 fag: ${subjects.join(', ')}`;
    });

    const blokkCollisionText = activeStudentsWithBlokkCollisions.map((entry) => {
      return `${entry.student.navn || 'Ukjent'} (${entry.student.klasse || 'Ingen klasse'}) - ${entry.collisionDetails.join(' | ')}`;
    });

    const clipboardText = [
      `Advarselsliste`,
      ``,
      `Under 3 fag (${activeStudentsWithFewSubjects.length})`,
      ...fewSubjectsText,
      ``,
      `4+ fag (${activeStudentsWithFourSubjects.length})`,
      ...fourSubjectsText,
      ``,
      `Blokk-kollisjon (${activeStudentsWithBlokkCollisions.length})`,
      ...blokkCollisionText
    ].join('\n');

    try {
      await navigator.clipboard.writeText(clipboardText);
      setWarningCopyStatus('Kopiert til utklippstavle');
    } catch {
      setWarningCopyStatus('Kopiering mislyktes');
    }

    window.setTimeout(() => {
      setWarningCopyStatus('');
    }, 2000);
  };

  const handleOpenStudentInElever = (studentId: string) => {
    setSelectedEleverStudentId(studentId);
    setActiveDataTab('elever');
  };

  const hasDataTabs = parsedFiles.length > 0 || mergedData.length > 0;

  return (
    <div className="app">
      <main className="main">
        <div className="content-container">
          <header className="header">
            <h1>Fagvalg - Oversikt</h1>
            <p>Slå sammen fagvalg fra flere programområder og trinn</p>
          </header>

        {hasDataTabs && (
          <>
              {mergedData.length > 0 && (
              <div className={`warning-box ${hasActiveWarnings ? '' : 'warning-box-clear'}`.trim()}>
                <h3 
                  className={`collapsible-header warning-header ${hasActiveWarnings ? '' : 'warning-header-clear'}`.trim()}
                  onClick={() => setWarningExpanded(!warningExpanded)}
                >
                  <span className="chevron">{warningExpanded ? '▼' : '▶'}</span>
                  {hasActiveWarnings
                    ? `⚠️ Advarsel: ${activeStudentsWithBlokkCollisions.length} blokk-kollisjon, ${activeStudentsWithFewSubjects.length} under 3 fag, ${activeStudentsWithFourSubjects.length} med 4+ fag`
                    : '✅ Ingen aktive advarsler'}
                </h3>
                {warningExpanded && (
                  <div className="warning-content">
                    {!hasActiveWarnings && studentsWithFewSubjects.length === 0 && studentsWithFourSubjects.length === 0 && studentsWithBlokkCollisions.length === 0 && (
                      <p className="warning-clear-message">Alle elever har gyldig antall blokkfag.</p>
                    )}
                    <div className="warning-actions">
                      <button type="button" className="warning-action-btn" onClick={handleWarningExport}>
                        Eksporter til Excel
                      </button>
                      <button type="button" className="warning-action-btn" onClick={handleWarningCopy}>
                        Kopier til utklippstavle
                      </button>
                      {warningCopyStatus && <span className="warning-copy-status">{warningCopyStatus}</span>}
                    </div>

                    <button
                      type="button"
                      className="warning-subtitle-toggle"
                      onClick={() => setWarningBlokkCollisionExpanded((prev) => !prev)}
                    >
                      <span className="chevron">{warningBlokkCollisionExpanded ? '▼' : '▶'}</span>
                      <span className="warning-subtitle">Elever med blokk-kollisjon ({activeStudentsWithBlokkCollisions.length})</span>
                    </button>
                    {warningBlokkCollisionExpanded && (
                      <ul>
                        {studentsWithBlokkCollisions.map((entry, idx) => {
                          const student = entry.student;
                          return (
                            <li key={`collision-${entry.studentId}-${idx}`}>
                              <div className="warning-line">
                                <span>
                                    <button
                                      type="button"
                                      className="warning-student-link"
                                      onClick={() => handleOpenStudentInElever(entry.studentId)}
                                    >
                                      <strong>{student.navn || 'Ukjent'}</strong>
                                    </button>{' '}
                                    ({student.klasse || 'Ingen klasse'}) - {entry.collisionDetails.join(' | ')}
                                </span>
                              </div>
                            </li>
                          );
                        })}
                      </ul>
                    )}

                    <hr className="warning-divider" />

                    <button
                      type="button"
                      className="warning-subtitle-toggle"
                      onClick={() => setWarningFewSubjectsExpanded((prev) => !prev)}
                    >
                      <span className="chevron">{warningFewSubjectsExpanded ? '▼' : '▶'}</span>
                      <span className="warning-subtitle">Elever med færre enn 3 blokkfag ({activeStudentsWithFewSubjects.length}, ignorert: {fewSubjectsIgnoredCount})</span>
                    </button>
                    {warningFewSubjectsExpanded && (
                      <ul>
                        {studentsWithFewSubjects.map((entry, idx) => {
                          const student = entry.student;
                          const studentId = entry.studentId;
                          const ignored = warningIgnoresByStudentAndType[studentId]?.missing;
                          const subjects = getWarningSubjects(student);
                          return (
                            <li key={`few-${studentId}-${idx}`} className={ignored ? 'warning-ignored-item' : ''}>
                              <div className="warning-line">
                                <span>
                                    <button
                                      type="button"
                                      className="warning-student-link"
                                      onClick={() => handleOpenStudentInElever(studentId)}
                                    >
                                      <strong>{student.navn || 'Ukjent'}</strong>
                                    </button>{' '}
                                    ({student.klasse || 'Ingen klasse'}) - {subjects.length} fag: {subjects.join(', ') || 'Ingen'}
                                </span>
                                {ignored && <span className="warning-ignore-badge">Ignorert</span>}
                              </div>
                              {ignored ? (
                                <div className="warning-ignore-row">
                                  <span className="warning-ignore-comment">Kommentar: {ignored.comment || 'Ingen kommentar'}</span>
                                  <button
                                    type="button"
                                    className="warning-inline-btn"
                                    onClick={() => removeWarningIgnore(studentId, 'missing')}
                                  >
                                    Fjern ignorering
                                  </button>
                                </div>
                              ) : (
                                <div className="warning-ignore-row">
                                  <input
                                    type="text"
                                    maxLength={140}
                                    className="warning-ignore-input"
                                    placeholder="Kommentar (valgfritt)"
                                    value={warningIgnoreDraftByStudentId[studentId] || ''}
                                    onChange={(event) => {
                                      const value = event.target.value;
                                      setWarningIgnoreDraftByStudentId((prev) => ({
                                        ...prev,
                                        [studentId]: value,
                                      }));
                                    }}
                                  />
                                  <button
                                    type="button"
                                    className="warning-inline-btn"
                                    onClick={() => saveWarningIgnore(studentId, 'missing')}
                                  >
                                    Ignorer
                                  </button>
                                </div>
                              )}
                            </li>
                          );
                        })}
                      </ul>
                    )}

                    <hr className="warning-divider" />

                    <button
                      type="button"
                      className="warning-subtitle-toggle"
                      onClick={() => setWarningFourSubjectsExpanded((prev) => !prev)}
                    >
                      <span className="chevron">{warningFourSubjectsExpanded ? '▼' : '▶'}</span>
                      <span className="warning-subtitle">Elever med 4+ blokkfag ({activeStudentsWithFourSubjects.length}, ignorert: {fourSubjectsIgnoredCount})</span>
                    </button>
                    {warningFourSubjectsExpanded && (
                      <ul>
                        {studentsWithFourSubjects.map((entry, idx) => {
                          const student = entry.student;
                          const studentId = entry.studentId;
                          const ignored = warningIgnoresByStudentAndType[studentId]?.overloaded;
                          const subjects = getWarningSubjects(student);
                          return (
                            <li key={`four-${studentId}-${idx}`} className={ignored ? 'warning-ignored-item' : ''}>
                              <div className="warning-line">
                                <span>
                                    <button
                                      type="button"
                                      className="warning-student-link"
                                      onClick={() => handleOpenStudentInElever(studentId)}
                                    >
                                      <strong>{student.navn || 'Ukjent'}</strong>
                                    </button>{' '}
                                    ({student.klasse || 'Ingen klasse'}) - 4 fag: {subjects.join(', ')}
                                </span>
                                {ignored && <span className="warning-ignore-badge">Ignorert</span>}
                              </div>
                              {ignored ? (
                                <div className="warning-ignore-row">
                                  <span className="warning-ignore-comment">Kommentar: {ignored.comment || 'Ingen kommentar'}</span>
                                  <button
                                    type="button"
                                    className="warning-inline-btn"
                                    onClick={() => removeWarningIgnore(studentId, 'overloaded')}
                                  >
                                    Fjern ignorering
                                  </button>
                                </div>
                              ) : (
                                <div className="warning-ignore-row">
                                  <input
                                    type="text"
                                    maxLength={140}
                                    className="warning-ignore-input"
                                    placeholder="Kommentar (valgfritt)"
                                    value={warningIgnoreDraftByStudentId[studentId] || ''}
                                    onChange={(event) => {
                                      const value = event.target.value;
                                      setWarningIgnoreDraftByStudentId((prev) => ({
                                        ...prev,
                                        [studentId]: value,
                                      }));
                                    }}
                                  />
                                  <button
                                    type="button"
                                    className="warning-inline-btn"
                                    onClick={() => saveWarningIgnore(studentId, 'overloaded')}
                                  >
                                    Ignorer
                                  </button>
                                </div>
                              )}
                            </li>
                          );
                        })}
                      </ul>
                    )}
                  </div>
                )}
              </div>
              )}
            
            <div className="control-row-group">
              <div className="control-row-label">Visning</div>
              <div className="data-tabs" role="tablist" aria-label="Data visning">
                <button
                  type="button"
                  role="tab"
                  aria-selected={activeDataTab === 'import'}
                  className={`data-tab ${activeDataTab === 'import' ? 'data-tab-active' : ''}`.trim()}
                  onClick={() => setActiveDataTab('import')}
                >
                  Last inn data
                </button>
                <button
                  type="button"
                  role="tab"
                  aria-selected={activeDataTab === 'subjects'}
                  className={`data-tab ${activeDataTab === 'subjects' ? 'data-tab-active' : ''}`.trim()}
                  onClick={() => setActiveDataTab('subjects')}
                >
                  Fagoversikt ({subjects.length} fag)
                </button>
                <button
                  type="button"
                  role="tab"
                  aria-selected={activeDataTab === 'elever'}
                  className={`data-tab ${activeDataTab === 'elever' ? 'data-tab-active' : ''}`.trim()}
                  onClick={() => setActiveDataTab('elever')}
                >
                  Elever ({mergedData.length})
                </button>
                <button
                  type="button"
                  role="tab"
                  aria-selected={activeDataTab === 'students'}
                  className={`data-tab ${activeDataTab === 'students' ? 'data-tab-active' : ''}`.trim()}
                  onClick={() => setActiveDataTab('students')}
                >
                  Elevtabell ({mergedData.length} elever)
                </button>
              </div>
            </div>

            <div className="data-tab-panel">
              {activeDataTab === 'import' ? (
                <>
                  <FileUploader onFilesAdded={handleFilesAdded} />

                  {parsedFiles.length > 0 && (
                    <>
                      <div className="uploaded-files">
                        <h3 className="uploaded-files-title">Opplastede filer ({parsedFiles.length})</h3>
                        <ul>
                          {parsedFiles.map((file) => (
                            <li key={file.id}>
                              <span>{file.filename}</span>
                              <button
                                onClick={() => handleRemoveFile(file.id)}
                                className="remove-btn"
                              >
                                Fjern
                              </button>
                            </li>
                          ))}
                        </ul>
                      </div>

                      <div className="column-mapper-section">
                        <h3
                          className="collapsible-header"
                          onClick={() => setColumnMapperExpanded(!columnMapperExpanded)}
                        >
                          <span className="chevron">{columnMapperExpanded ? '▼' : '▶'}</span>
                          Oppsett
                        </h3>
                        {columnMapperExpanded && (
                          <ColumnMapper
                            files={parsedFiles}
                            onMappingChange={handleMappingChange}
                            currentMappings={mappings}
                            blokkCount={blokkCount}
                            onBlokkCountChange={setBlokkCount}
                          />
                        )}
                      </div>
                    </>
                  )}

                  <div className="action-buttons">
                    <button onClick={handleMerge} className="merge-btn" disabled={parsedFiles.length === 0}>
                      Slå sammen data
                    </button>
                    <button onClick={handleClearStoredData} className="clear-storage-btn">
                      Tøm data
                    </button>
                    <button
                      onClick={handleExport}
                      className="export-btn"
                      disabled={mergedData.length === 0}
                      title={mergedData.length === 0 ? 'Slå sammen data først' : 'Eksporter sammenslått data'}
                    >
                      Eksporter til Novaschem
                    </button>
                    <button
                      onClick={handleExportDetailed}
                      className="export-btn"
                      disabled={mergedData.length === 0}
                      title={mergedData.length === 0 ? 'Slå sammen data først' : 'Eksporter med separate blokk-kolonner og fullstendige fagnavn'}
                    >
                      Eksporter til Excel (full)
                    </button>
                    <button
                      onClick={handleExportText}
                      className="export-btn"
                      disabled={mergedData.length === 0}
                      title={mergedData.length === 0 ? 'Slå sammen data først' : 'Eksporter som tabulatorseparert tekstfil med fagnummer'}
                    >
                      Eksporter til TXT
                    </button>
                  </div>
                </>
              ) : activeDataTab === 'subjects' ? (
                <SubjectTally
                  subjects={subjects}
                  mergedData={mergedData}
                  subjectSettingsByName={subjectSettingsByName}
                  onSaveSubjectSettingsByName={setSubjectSettingsByName}
                  onApplySubjectBlockMoves={handleApplySubjectBlockMoves}
                  onRemoveStudentsFromSubject={handleRemoveStudentsFromSubject}
                />
              ) : activeDataTab === 'students' ? (
                <MergedDataView
                  data={filteredMergedData}
                  totalDataCount={mergedData.length}
                  selectedSubject={selectedMergedSubject}
                  onSubjectFilterChange={setSelectedMergedSubject}
                  subjectOptions={subjects.map((subject) => subject.subject)}
                  blokkCount={blokkCount}
                />
              ) : (
                <EleverView
                  data={mergedData}
                  blokkCount={blokkCount}
                  subjectOptions={subjects.map((subject) => subject.subject)}
                  subjectSettingsByName={subjectSettingsByName}
                  onSaveSubjectSettingsByName={setSubjectSettingsByName}
                  warningIgnoresByStudentAndType={warningIgnoresByStudentAndType}
                  onSaveWarningIgnore={(studentId, type, comment) => saveWarningIgnore(studentId, type, comment)}
                  onRemoveWarningIgnore={removeWarningIgnore}
                  changeLog={studentAssignmentChanges}
                  onStudentDataUpdate={handleStudentAssignmentsUpdated}
                  externallySelectedStudentId={selectedEleverStudentId}
                />
              )}
            </div>
          </>
        )}

        {!hasDataTabs && (
          <div className="data-tab-panel">
            <FileUploader onFilesAdded={handleFilesAdded} />
          </div>
        )}
        </div>
      </main>
    </div>
  );
}

export default App;
