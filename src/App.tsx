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
  subjectMaxByName?: Record<string, number>;
  blokkCount: number;
  selectedMergedSubject: string;
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
  const [activeDataTab, setActiveDataTab] = useState<'subjects' | 'students' | 'elever'>('subjects');
  const [warningExpanded, setWarningExpanded] = useState(false);
  const [warningCopyStatus, setWarningCopyStatus] = useState('');
  const [warningIgnoresByStudentId, setWarningIgnoresByStudentId] = useState<Record<string, { comment: string; ignoredAt: string }>>({});
  const [warningIgnoreDraftByStudentId, setWarningIgnoreDraftByStudentId] = useState<Record<string, string>>({});
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

      if (parsedState.warningIgnoresByStudentId && typeof parsedState.warningIgnoresByStudentId === 'object') {
        setWarningIgnoresByStudentId(parsedState.warningIgnoresByStudentId);
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
      warningIgnoresByStudentId,
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
    warningIgnoresByStudentId,
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
  };

  const handleReset = () => {
    setParsedFiles([]);
    setMappings(new Map());
    setMergedData([]);
    setSubjects([]);
    setStudentAssignmentChanges([]);
    setSubjectSettingsByName({});
    setWarningIgnoresByStudentId({});
    setWarningIgnoreDraftByStudentId({});
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

  const getWarningStudentId = (student: StandardField): string => {
    if (student.studentId && student.studentId.trim().length > 0) {
      return student.studentId;
    }

    return `${student.navn || 'ukjent'}|${student.klasse || 'ukjent'}`;
  };

  // Get students with less than 3 blokkfag
  const getStudentsWithFewSubjects = () => {
    return mergedData
      .filter((student) => {
        const blokkCount = [
          student.blokk1,
          student.blokk2,
          student.blokk3,
          student.blokk4
        ].filter((blokk) => blokk && blokk.trim() !== '').length;
        return blokkCount < 3;
      })
      .sort((a, b) => (a.navn || '').localeCompare(b.navn || '', 'nb', { sensitivity: 'base' }));
  };

  const studentsWithFewSubjects = getStudentsWithFewSubjects();

  const getStudentsWithFourSubjects = () => {
    return mergedData.filter((student) => {
      const blokkCount = [
        student.blokk1,
        student.blokk2,
        student.blokk3,
        student.blokk4
      ].filter((blokk) => blokk && blokk.trim() !== '').length;
      return blokkCount >= 4;
    })
      .sort((a, b) => (a.navn || '').localeCompare(b.navn || '', 'nb', { sensitivity: 'base' }));
  };

  const studentsWithFourSubjects = getStudentsWithFourSubjects();

  const fewSubjectsIgnoredCount = studentsWithFewSubjects.filter(
    (student) => !!warningIgnoresByStudentId[getWarningStudentId(student)]
  ).length;

  const fourSubjectsIgnoredCount = studentsWithFourSubjects.filter(
    (student) => !!warningIgnoresByStudentId[getWarningStudentId(student)]
  ).length;

  const warningStudentIds = new Set<string>([
    ...studentsWithFewSubjects.map((student) => getWarningStudentId(student)),
    ...studentsWithFourSubjects.map((student) => getWarningStudentId(student)),
  ]);

  useEffect(() => {
    setWarningIgnoresByStudentId((prev) => {
      const next = Object.fromEntries(
        Object.entries(prev).filter(([studentId]) => warningStudentIds.has(studentId))
      ) as Record<string, { comment: string; ignoredAt: string }>;

      return Object.keys(next).length === Object.keys(prev).length ? prev : next;
    });

    setWarningIgnoreDraftByStudentId((prev) => {
      const next = Object.fromEntries(
        Object.entries(prev).filter(([studentId]) => warningStudentIds.has(studentId))
      ) as Record<string, string>;

      return Object.keys(next).length === Object.keys(prev).length ? prev : next;
    });
  }, [studentsWithFewSubjects, studentsWithFourSubjects]);

  const saveWarningIgnore = (studentId: string) => {
    const rawComment = warningIgnoreDraftByStudentId[studentId] || '';
    const comment = rawComment.trim();
    if (!comment) {
      return;
    }

    setWarningIgnoresByStudentId((prev) => ({
      ...prev,
      [studentId]: {
        comment,
        ignoredAt: new Date().toISOString(),
      },
    }));
  };

  const removeWarningIgnore = (studentId: string) => {
    setWarningIgnoresByStudentId((prev) => {
      const next = { ...prev };
      delete next[studentId];
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
      ...studentsWithFewSubjects.map((student) => {
        const studentId = getWarningStudentId(student);
        const ignored = warningIgnoresByStudentId[studentId];
        const subjects = [student.blokk1, student.blokk2, student.blokk3, student.blokk4]
          .filter((blokk) => blokk && blokk.trim() !== '');
        return {
          Kategori: 'Under 3 fag',
          Navn: student.navn || 'Ukjent',
          Klasse: student.klasse || 'Ingen klasse',
          AntallFag: subjects.length,
          Fag: subjects.join(', '),
          Ignorert: ignored ? 'Ja' : 'Nei',
          Kommentar: ignored?.comment || '',
        };
      }),
      ...studentsWithFourSubjects.map((student) => {
        const studentId = getWarningStudentId(student);
        const ignored = warningIgnoresByStudentId[studentId];
        const subjects = [student.blokk1, student.blokk2, student.blokk3, student.blokk4]
          .filter((blokk) => blokk && blokk.trim() !== '');
        return {
          Kategori: '4 fag',
          Navn: student.navn || 'Ukjent',
          Klasse: student.klasse || 'Ingen klasse',
          AntallFag: subjects.length,
          Fag: subjects.join(', '),
          Ignorert: ignored ? 'Ja' : 'Nei',
          Kommentar: ignored?.comment || '',
        };
      })
    ];

    const worksheet = XLSX.utils.json_to_sheet(warningRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Advarsler');
    XLSX.writeFile(workbook, 'warning_students.xlsx');
  };

  const handleWarningCopy = async () => {
    const fewSubjectsText = studentsWithFewSubjects.map((student) => {
      const studentId = getWarningStudentId(student);
      const ignored = warningIgnoresByStudentId[studentId];
      const subjects = [student.blokk1, student.blokk2, student.blokk3, student.blokk4]
        .filter((blokk) => blokk && blokk.trim() !== '');
      const ignoreText = ignored ? ` [IGNORERT: ${ignored.comment}]` : '';
      return `${student.navn || 'Ukjent'} (${student.klasse || 'Ingen klasse'}) - ${subjects.length} fag: ${subjects.join(', ') || 'Ingen'}${ignoreText}`;
    });

    const fourSubjectsText = studentsWithFourSubjects.map((student) => {
      const studentId = getWarningStudentId(student);
      const ignored = warningIgnoresByStudentId[studentId];
      const subjects = [student.blokk1, student.blokk2, student.blokk3, student.blokk4]
        .filter((blokk) => blokk && blokk.trim() !== '');
      const ignoreText = ignored ? ` [IGNORERT: ${ignored.comment}]` : '';
      return `${student.navn || 'Ukjent'} (${student.klasse || 'Ingen klasse'}) - 4 fag: ${subjects.join(', ')}${ignoreText}`;
    });

    const clipboardText = [
      `Advarselsliste`,
      ``,
      `Under 3 fag (${studentsWithFewSubjects.length})`,
      ...fewSubjectsText,
      ``,
      `4+ fag (${studentsWithFourSubjects.length})`,
      ...fourSubjectsText
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

  return (
    <div className="app">
      <main className="main">
        <div className="content-container">
          <header className="header">
            <h1>Fagvalg - Oversikt</h1>
            <p>Slå sammen fagvalg fra flere programområder og trinn</p>
          </header>

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

            <div className="action-buttons">
              <button onClick={handleMerge} className="merge-btn">
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
        )}

        {mergedData.length > 0 && (
          <>
            {(studentsWithFewSubjects.length > 0 || studentsWithFourSubjects.length > 0) && (
              <div className="warning-box">
                <h3 
                  className="collapsible-header warning-header" 
                  onClick={() => setWarningExpanded(!warningExpanded)}
                >
                  <span className="chevron">{warningExpanded ? '▼' : '▶'}</span>
                  ⚠️ Advarsel: {studentsWithFewSubjects.length} under 3 fag, {studentsWithFourSubjects.length} med 4+ fag
                </h3>
                {warningExpanded && (
                  <div className="warning-content">
                    <div className="warning-actions">
                      <button type="button" className="warning-action-btn" onClick={handleWarningExport}>
                        Eksporter til Excel
                      </button>
                      <button type="button" className="warning-action-btn" onClick={handleWarningCopy}>
                        Kopier til utklippstavle
                      </button>
                      {warningCopyStatus && <span className="warning-copy-status">{warningCopyStatus}</span>}
                    </div>

                    <h4 className="warning-subtitle">Elever med færre enn 3 blokkfag ({studentsWithFewSubjects.length}, ignorert: {fewSubjectsIgnoredCount})</h4>
                    <ul>
                      {studentsWithFewSubjects.map((student, idx) => {
                        const studentId = getWarningStudentId(student);
                        const ignored = warningIgnoresByStudentId[studentId];
                        const blokkCount = [
                          student.blokk1,
                          student.blokk2,
                          student.blokk3,
                          student.blokk4
                        ].filter((blokk) => blokk && blokk.trim() !== '').length;
                        const subjects = [
                          student.blokk1,
                          student.blokk2,
                          student.blokk3,
                          student.blokk4
                        ].filter((blokk) => blokk && blokk.trim() !== '');
                        return (
                          <li key={`few-${studentId}-${idx}`} className={ignored ? 'warning-ignored-item' : ''}>
                            <div className="warning-line">
                              <span>
                                <strong>{student.navn || 'Ukjent'}</strong> ({student.klasse || 'Ingen klasse'}) - {blokkCount} fag: {subjects.join(', ') || 'Ingen'}
                              </span>
                              {ignored && <span className="warning-ignore-badge">Ignorert</span>}
                            </div>
                            {ignored ? (
                              <div className="warning-ignore-row">
                                <span className="warning-ignore-comment">Kommentar: {ignored.comment}</span>
                                <button
                                  type="button"
                                  className="warning-inline-btn"
                                  onClick={() => removeWarningIgnore(studentId)}
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
                                  placeholder="Kort kommentar for ignorering"
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
                                  onClick={() => saveWarningIgnore(studentId)}
                                  disabled={!(warningIgnoreDraftByStudentId[studentId] || '').trim()}
                                >
                                  Ignorer
                                </button>
                              </div>
                            )}
                          </li>
                        );
                      })}
                    </ul>

                    <hr className="warning-divider" />

                    <h4 className="warning-subtitle">Elever med 4+ blokkfag ({studentsWithFourSubjects.length}, ignorert: {fourSubjectsIgnoredCount})</h4>
                    <ul>
                      {studentsWithFourSubjects.map((student, idx) => {
                        const studentId = getWarningStudentId(student);
                        const ignored = warningIgnoresByStudentId[studentId];
                        const subjects = [
                          student.blokk1,
                          student.blokk2,
                          student.blokk3,
                          student.blokk4
                        ].filter((blokk) => blokk && blokk.trim() !== '');
                        return (
                          <li key={`four-${studentId}-${idx}`} className={ignored ? 'warning-ignored-item' : ''}>
                            <div className="warning-line">
                              <span>
                                <strong>{student.navn || 'Ukjent'}</strong> ({student.klasse || 'Ingen klasse'}) - 4 fag: {subjects.join(', ')}
                              </span>
                              {ignored && <span className="warning-ignore-badge">Ignorert</span>}
                            </div>
                            {ignored ? (
                              <div className="warning-ignore-row">
                                <span className="warning-ignore-comment">Kommentar: {ignored.comment}</span>
                                <button
                                  type="button"
                                  className="warning-inline-btn"
                                  onClick={() => removeWarningIgnore(studentId)}
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
                                  placeholder="Kort kommentar for ignorering"
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
                                  onClick={() => saveWarningIgnore(studentId)}
                                  disabled={!(warningIgnoreDraftByStudentId[studentId] || '').trim()}
                                >
                                  Ignorer
                                </button>
                              </div>
                            )}
                          </li>
                        );
                      })}
                    </ul>
                  </div>
                )}
              </div>
            )}
            
            <div className="data-tabs" role="tablist" aria-label="Data visning">
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

            <div className="data-tab-panel">
              {activeDataTab === 'subjects' ? (
                <SubjectTally
                  subjects={subjects}
                  mergedData={mergedData}
                  subjectSettingsByName={subjectSettingsByName}
                  onSaveSubjectSettingsByName={setSubjectSettingsByName}
                  onApplySubjectBlockMoves={handleApplySubjectBlockMoves}
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
                  changeLog={studentAssignmentChanges}
                  onStudentDataUpdate={handleStudentAssignmentsUpdated}
                />
              )}
            </div>
          </>
        )}
        </div>
      </main>
    </div>
  );
}

export default App;
