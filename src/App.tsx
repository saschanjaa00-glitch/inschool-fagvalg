import { useEffect, useMemo, useRef, useState, type ChangeEvent } from 'react';
import type { ParsedFile, ColumnMapping, StandardField, SubjectCount, StudentAssignmentChange, GroupMoveLogEntry } from './utils/excelUtils';
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
import { anonymizeStudents } from './utils/anonymize';
import './App.css';
import { FileUploader } from './components/FileUploader';
import { ColumnMapper } from './components/ColumnMapper';
import { MergedDataView } from './components/MergedDataView';
import { SubjectTally } from './components/SubjectTally';
import { GrupperView } from './components/GrupperView';
import { FagoversiktView } from './components/FagoversiktView';
import { EleverView } from './components/EleverView';
import { ChangeLogView } from './components/ChangeLogView';
import { BalanseringView } from './components/BalanseringView';
import type { SubjectSettingsByName } from './components/SubjectTally';
import {
  DEFAULT_BALANCING_CONFIG,
  DEFAULT_CLASS_BLOCK_RESTRICTIONS,
  type ClassBlockRestrictions,
  type ProgressiveHybridBalanceResult,
} from './utils/progressiveHybridBalance';

const LOCAL_STORAGE_KEY = 'fagvalg-opptelling-state-v1';
const JSON_TRANSFER_FORMAT = 'inschool-balansering-state';
const JSON_TRANSFER_VERSION = 1;
const MAX_HISTORY_STATES = 20;

interface PersistedAppState {
  parsedFiles: ParsedFile[];
  mappings: Record<string, ColumnMapping>;
  mergedData: StandardField[];
  baselineMergedData?: StandardField[];
  subjects: SubjectCount[];
  studentAssignmentChanges?: StudentAssignmentChange[];
  groupMoveLog?: GroupMoveLogEntry[];
  subjectSettingsByName: SubjectSettingsByName;
  warningIgnoresByStudentId?: Record<string, { comment: string; ignoredAt: string }>;
  warningIgnoresByStudentAndType?: Record<string, Partial<Record<WarningType, WarningIgnoreEntry>>>;
  subjectMaxByName?: Record<string, number>;
  blokkCount: number;
  selectedMergedSubject: string;
  classBlockRestrictions?: ClassBlockRestrictions;
  balancingExcludedSubjects?: string[];
  balancingExcludedStudentIds?: string[];
  nextBalancingRoundId?: number;
}

interface ExportedAppStateFile {
  format: string;
  version: number;
  exportedAt: string;
  appState: PersistedAppState;
}

type WarningType = 'missing' | 'overloaded';

interface WarningIgnoreEntry {
  comment: string;
  ignoredAt: string;
}

type GroupSubview = 'subjects' | 'groups' | 'fagoversikt';
type StudentSubview = 'elever' | 'students';

const clonePersistedState = (state: PersistedAppState): PersistedAppState => {
  return JSON.parse(JSON.stringify(state)) as PersistedAppState;
};

const arePersistedStatesEqual = (left: PersistedAppState, right: PersistedAppState): boolean => {
  return JSON.stringify(left) === JSON.stringify(right);
};

const normalizeSubject = (value: string): string => {
  return value.trim().toLocaleLowerCase('nb');
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

const resolveStatusAction = (change: StudentAssignmentChange): 'added' | 'removed' | 'readded' | null => {
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

const getSummarizedChangeLogCounts = (
  changeLog: StudentAssignmentChange[]
): { students: number; changes: number } => {
  if (changeLog.length === 0) {
    return { students: 0, changes: 0 };
  }

  const changesByStudentId = new Map<string, StudentAssignmentChange[]>();

  changeLog.forEach((entry) => {
    const existing = changesByStudentId.get(entry.studentId) || [];
    existing.push(entry);
    changesByStudentId.set(entry.studentId, existing);
  });

  let studentCount = 0;
  let changeCount = 0;

  changesByStudentId.forEach((studentChanges) => {
    const oldestFirst = [...studentChanges].sort((left, right) => {
      return new Date(left.changedAt).getTime() - new Date(right.changedAt).getTime();
    });

    const assignmentChanges = oldestFirst.filter((entry) => !isStudentStatusChange(entry));
    const assignmentSummaryBySubject = new Map<
      string,
      { fromBlokk: number; toBlokk: number }
    >();

    assignmentChanges.forEach((entry) => {
      const subjectKey = normalizeSubject(entry.subject || '');
      if (!subjectKey) {
        return;
      }

      const existing = assignmentSummaryBySubject.get(subjectKey);
      if (!existing) {
        assignmentSummaryBySubject.set(subjectKey, {
          fromBlokk: entry.fromBlokk,
          toBlokk: entry.toBlokk,
        });
        return;
      }

      existing.toBlokk = entry.toBlokk;
    });

    const summarizedAssignmentCount = Array.from(assignmentSummaryBySubject.values()).filter((entry) => {
      return entry.fromBlokk !== entry.toBlokk;
    }).length;

    const statusEntriesRaw: Array<'added' | 'removed' | 'readded' | null> = [];
    const openRemovedIndexes: number[] = [];

    oldestFirst
      .filter((entry) => isStudentStatusChange(entry))
      .forEach((entry) => {
        const action = resolveStatusAction(entry);
        if (!action) {
          return;
        }

        if (action === 'removed') {
          statusEntriesRaw.push(action);
          openRemovedIndexes.push(statusEntriesRaw.length - 1);
          return;
        }

        if (action === 'readded') {
          if (openRemovedIndexes.length > 0) {
            const cancelledIndex = openRemovedIndexes.pop();
            if (typeof cancelledIndex === 'number') {
              statusEntriesRaw[cancelledIndex] = null;
            }
            return;
          }

          statusEntriesRaw.push(action);
          return;
        }

        statusEntriesRaw.push(action);
      });

    const summarizedStatusCount = statusEntriesRaw.filter((entry) => entry !== null).length;
    const summarizedTotal = summarizedAssignmentCount + summarizedStatusCount;

    if (summarizedTotal > 0) {
      studentCount += 1;
      changeCount += summarizedTotal;
    }
  });

  return {
    students: studentCount,
    changes: changeCount,
  };
};

function App() {
  const jsonImportInputRef = useRef<HTMLInputElement | null>(null);
  const studentExportMenuRef = useRef<HTMLDivElement | null>(null);
  const [parsedFiles, setParsedFiles] = useState<ParsedFile[]>([]);
  const [anonymize, setAnonymize] = useState(false);
  const [mappings, setMappings] = useState<Map<string, ColumnMapping>>(new Map());
  const [mergedData, setMergedData] = useState<StandardField[]>([]);
  const [baselineMergedData, setBaselineMergedData] = useState<StandardField[]>([]);
  const [subjects, setSubjects] = useState<SubjectCount[]>([]);
  const [studentAssignmentChanges, setStudentAssignmentChanges] = useState<StudentAssignmentChange[]>([]);
  const [groupMoveLog, setGroupMoveLog] = useState<GroupMoveLogEntry[]>([]);
  const [subjectSettingsByName, setSubjectSettingsByName] = useState<SubjectSettingsByName>({});
  const [blokkCount, setBlokkCount] = useState(4);
  const [nextBalancingRoundId, setNextBalancingRoundId] = useState(1);
  
  const [columnMapperExpanded, setColumnMapperExpanded] = useState(false);
  const [activeDataTab, setActiveDataTab] = useState<
    'import' | 'subjects' | 'groups' | 'fagoversikt' | 'students' | 'elever' | 'balancing' | 'changelog'
  >('import');
  const [activeGroupTab, setActiveGroupTab] = useState<GroupSubview>('subjects');
  const [activeStudentTab, setActiveStudentTab] = useState<StudentSubview>('elever');
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
  const [eleverViewActivationToken, setEleverViewActivationToken] = useState(0);
  const [selectedMergedSubject, setSelectedMergedSubject] = useState('');
  const [subjectSettingsAutoOpenToken, setSubjectSettingsAutoOpenToken] = useState(0);
  const [subjectSettingsAutoOpenHandledToken, setSubjectSettingsAutoOpenHandledToken] = useState(0);
  const [undoHistory, setUndoHistory] = useState<PersistedAppState[]>([]);
  const [redoHistory, setRedoHistory] = useState<PersistedAppState[]>([]);
  const [classBlockRestrictions, setClassBlockRestrictions] = useState<ClassBlockRestrictions>(
    DEFAULT_CLASS_BLOCK_RESTRICTIONS
  );
  const [balancingExcludedSubjects, setBalancingExcludedSubjects] = useState<string[]>(
    DEFAULT_BALANCING_CONFIG.excludedSubjects
  );
  const [balancingExcludedStudentIds, setBalancingExcludedStudentIds] = useState<string[]>([]);
  const [isHydratedFromStorage, setIsHydratedFromStorage] = useState(false);
  const [showReloadConfirmModal, setShowReloadConfirmModal] = useState(false);
  const [isReloadConfirmArmed, setIsReloadConfirmArmed] = useState(false);
  const [showHelpModal, setShowHelpModal] = useState(false);
  const [jsonTransferStatus, setJsonTransferStatus] = useState('');
  const [isStudentExportMenuOpen, setIsStudentExportMenuOpen] = useState(false);
  const persistTimeoutRef = useRef<number | null>(null);

  const showJsonTransferStatus = (message: string) => {
    setJsonTransferStatus(message);

    window.setTimeout(() => {
      setJsonTransferStatus('');
    }, 3000);
  };

  const buildPersistedState = (): PersistedAppState => ({
    parsedFiles,
    mappings: Object.fromEntries(mappings.entries()),
    mergedData,
    baselineMergedData,
    subjects,
    studentAssignmentChanges,
    groupMoveLog,
    subjectSettingsByName,
    warningIgnoresByStudentAndType,
    blokkCount,
    selectedMergedSubject,
    classBlockRestrictions,
    balancingExcludedSubjects,
    balancingExcludedStudentIds,
    nextBalancingRoundId,
  });

  const pushUndoSnapshot = (snapshot: PersistedAppState) => {
    const clonedSnapshot = clonePersistedState(snapshot);

    setUndoHistory((prev) => {
      const lastSnapshot = prev[prev.length - 1];
      if (lastSnapshot && arePersistedStatesEqual(lastSnapshot, clonedSnapshot)) {
        return prev;
      }

      const next = [...prev, clonedSnapshot];
      return next.slice(-MAX_HISTORY_STATES);
    });

    setRedoHistory([]);
  };

  const captureUndoSnapshot = () => {
    pushUndoSnapshot(buildPersistedState());
  };

  const applyPersistedState = (
    parsedState: Partial<PersistedAppState>,
    options?: { preserveActiveView?: boolean }
  ) => {
    const importedParsedFiles = Array.isArray(parsedState.parsedFiles) ? parsedState.parsedFiles : [];
    const importedMergedData = Array.isArray(parsedState.mergedData) ? parsedState.mergedData : [];
    const preserveActiveView = options?.preserveActiveView ?? false;

    setParsedFiles(importedParsedFiles);
    setMappings(
      parsedState.mappings && typeof parsedState.mappings === 'object'
        ? new Map(Object.entries(parsedState.mappings))
        : new Map()
    );
    setMergedData(importedMergedData);
    setBaselineMergedData(
      Array.isArray(parsedState.baselineMergedData) ? parsedState.baselineMergedData : importedMergedData
    );
    setSubjects(
      Array.isArray(parsedState.subjects)
        ? parsedState.subjects
        : importedMergedData.length > 0
          ? tallySubjects(importedMergedData)
          : []
    );
    setStudentAssignmentChanges(
      Array.isArray(parsedState.studentAssignmentChanges) ? parsedState.studentAssignmentChanges : []
    );
    setGroupMoveLog(
      Array.isArray(parsedState.groupMoveLog) ? parsedState.groupMoveLog : []
    );
    const maxRoundInChanges = Array.isArray(parsedState.studentAssignmentChanges)
      ? parsedState.studentAssignmentChanges.reduce((maxRound, change) => {
        return Math.max(maxRound, typeof change.balancingRoundId === 'number' ? change.balancingRoundId : 0);
      }, 0)
      : 0;

    if (parsedState.subjectSettingsByName && typeof parsedState.subjectSettingsByName === 'object') {
      setSubjectSettingsByName(parsedState.subjectSettingsByName);
    } else if (parsedState.subjectMaxByName && typeof parsedState.subjectMaxByName === 'object') {
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
    } else {
      setSubjectSettingsByName({});
    }

    if (parsedState.warningIgnoresByStudentAndType && typeof parsedState.warningIgnoresByStudentAndType === 'object') {
      setWarningIgnoresByStudentAndType(parsedState.warningIgnoresByStudentAndType);
    } else if (parsedState.warningIgnoresByStudentId && typeof parsedState.warningIgnoresByStudentId === 'object') {
      const migrated = Object.fromEntries(
        Object.entries(parsedState.warningIgnoresByStudentId).map(([studentId, value]) => [
          studentId,
          {
            missing: value,
          },
        ])
      ) as Record<string, Partial<Record<WarningType, WarningIgnoreEntry>>>;
      setWarningIgnoresByStudentAndType(migrated);
    } else {
      setWarningIgnoresByStudentAndType({});
    }

    setBlokkCount(typeof parsedState.blokkCount === 'number' ? parsedState.blokkCount : 4);
    setSelectedMergedSubject(typeof parsedState.selectedMergedSubject === 'string' ? parsedState.selectedMergedSubject : '');
    setClassBlockRestrictions(
      parsedState.classBlockRestrictions && typeof parsedState.classBlockRestrictions === 'object'
        ? parsedState.classBlockRestrictions
        : DEFAULT_CLASS_BLOCK_RESTRICTIONS
    );
    setBalancingExcludedSubjects(
      Array.isArray(parsedState.balancingExcludedSubjects)
        ? parsedState.balancingExcludedSubjects
        : DEFAULT_BALANCING_CONFIG.excludedSubjects
    );
    setBalancingExcludedStudentIds(
      Array.isArray(parsedState.balancingExcludedStudentIds)
        ? parsedState.balancingExcludedStudentIds
        : []
    );
    setNextBalancingRoundId(
      typeof parsedState.nextBalancingRoundId === 'number'
        ? parsedState.nextBalancingRoundId
        : Math.max(1, maxRoundInChanges + 1)
    );
    setWarningIgnoreDraftByStudentId({});
    setSelectedEleverStudentId('');
    setColumnMapperExpanded(importedParsedFiles.length > 0 && importedMergedData.length === 0);
    setShowReloadConfirmModal(false);
    setIsReloadConfirmArmed(false);
    setIsStudentExportMenuOpen(false);

    if (preserveActiveView && importedMergedData.length > 0) {
      setActiveDataTab((current) => (current === 'import' ? 'subjects' : current));
      return;
    }

    setActiveGroupTab('subjects');
    setActiveStudentTab('elever');
    setActiveDataTab(importedMergedData.length > 0 ? 'subjects' : 'import');
  };

  const handleUndo = () => {
    if (undoHistory.length === 0) {
      return;
    }

    const currentSnapshot = clonePersistedState(buildPersistedState());
    const previousSnapshot = undoHistory[undoHistory.length - 1];

    setUndoHistory((prev) => prev.slice(0, -1));
    setRedoHistory((prev) => [...prev, currentSnapshot].slice(-MAX_HISTORY_STATES));
    applyPersistedState(previousSnapshot, { preserveActiveView: true });
  };

  const handleRedo = () => {
    if (redoHistory.length === 0) {
      return;
    }

    const currentSnapshot = clonePersistedState(buildPersistedState());
    const nextSnapshot = redoHistory[redoHistory.length - 1];

    setRedoHistory((prev) => prev.slice(0, -1));
    setUndoHistory((prev) => [...prev, currentSnapshot].slice(-MAX_HISTORY_STATES));
    applyPersistedState(nextSnapshot, { preserveActiveView: true });
  };

  const handleSaveSubjectSettingsByName = (nextSettings: SubjectSettingsByName) => {
    if (JSON.stringify(nextSettings) === JSON.stringify(subjectSettingsByName)) {
      return;
    }

    captureUndoSnapshot();
    setSubjectSettingsByName(nextSettings);
  };

  const handleClassBlockRestrictionsChange = (nextRestrictions: ClassBlockRestrictions) => {
    if (JSON.stringify(nextRestrictions) === JSON.stringify(classBlockRestrictions)) {
      return;
    }

    captureUndoSnapshot();
    setClassBlockRestrictions(nextRestrictions);
  };

  const handleBlokkCountChange = (nextBlokkCount: number) => {
    if (nextBlokkCount === blokkCount) {
      return;
    }

    captureUndoSnapshot();
    setBlokkCount(nextBlokkCount);
  };

  useEffect(() => {
    if (activeDataTab === 'subjects' || activeDataTab === 'groups' || activeDataTab === 'fagoversikt') {
      setActiveGroupTab(activeDataTab);
    }
  }, [activeDataTab]);

  useEffect(() => {
    if (activeDataTab === 'elever' || activeDataTab === 'students') {
      setActiveStudentTab(activeDataTab);
    }
  }, [activeDataTab]);

  useEffect(() => {
    if (activeDataTab !== 'elever' || selectedEleverStudentId) {
      return;
    }

    setEleverViewActivationToken((prev) => prev + 1);
  }, [activeDataTab, selectedEleverStudentId]);

  useEffect(() => {
    if (!isStudentExportMenuOpen) {
      return;
    }

    const handleDocumentClick = (event: MouseEvent) => {
      if (!studentExportMenuRef.current?.contains(event.target as Node)) {
        setIsStudentExportMenuOpen(false);
      }
    };

    document.addEventListener('mousedown', handleDocumentClick);

    return () => {
      document.removeEventListener('mousedown', handleDocumentClick);
    };
  }, [isStudentExportMenuOpen]);

  useEffect(() => {
    if (!showHelpModal) {
      return;
    }

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        setShowHelpModal(false);
      }
    };

    document.addEventListener('keydown', handleKeyDown);

    return () => {
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [showHelpModal]);

  useEffect(() => {
    try {
      const savedState = localStorage.getItem(LOCAL_STORAGE_KEY);
      if (!savedState) {
        return;
      }

      const parsedState = JSON.parse(savedState) as Partial<PersistedAppState>;
      applyPersistedState(parsedState);
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

    const persistedState = buildPersistedState();

    if (persistTimeoutRef.current !== null) {
      window.clearTimeout(persistTimeoutRef.current);
    }

    persistTimeoutRef.current = window.setTimeout(() => {
      localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(persistedState));
      persistTimeoutRef.current = null;
    }, 250);
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
    classBlockRestrictions,
    balancingExcludedSubjects,
    balancingExcludedStudentIds,
    nextBalancingRoundId,
  ]);

  useEffect(() => {
    return () => {
      if (persistTimeoutRef.current !== null) {
        window.clearTimeout(persistTimeoutRef.current);
      }
    };
  }, []);

  const handleFilesAdded = (files: ParsedFile[]) => {
    captureUndoSnapshot();
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
    captureUndoSnapshot();
    const newMappings = new Map(mappings);
    newMappings.set(fileId, mapping);
    setMappings(newMappings);
  };

  const handleMerge = () => {
    captureUndoSnapshot();
    let merged = mergeFiles(parsedFiles, mappings).sort((a, b) => {
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

    if (anonymize) {
      merged = anonymizeStudents(merged);
    }

    setMergedData(merged);
    setBaselineMergedData(merged);
    setSubjects(tallySubjects(merged));
    setStudentAssignmentChanges([]);
    setGroupMoveLog([]);
    setNextBalancingRoundId(1);
    setActiveDataTab('subjects');
    setSubjectSettingsAutoOpenToken((prev) => prev + 1);
  };

  const closeReloadConfirmModal = () => {
    setShowReloadConfirmModal(false);
    setIsReloadConfirmArmed(false);
    setNextBalancingRoundId(1);
  };

  const handleLoadDataClick = () => {
    if (parsedFiles.length === 0) {
      return;
    }

    if (mergedData.length > 0) {
      setShowReloadConfirmModal(true);
      setIsReloadConfirmArmed(false);
      return;
    }

    handleMerge();
  };

  const handleReset = () => {
    captureUndoSnapshot();
    setParsedFiles([]);
    setMappings(new Map());
    setMergedData([]);
    setBaselineMergedData([]);
    setSubjects([]);
    setStudentAssignmentChanges([]);
    setGroupMoveLog([]);
    setSubjectSettingsByName({});
    setWarningIgnoresByStudentAndType({});
    setWarningIgnoreDraftByStudentId({});
    setSelectedEleverStudentId('');
    setSelectedMergedSubject('');
    setClassBlockRestrictions(DEFAULT_CLASS_BLOCK_RESTRICTIONS);
    setBalancingExcludedSubjects(DEFAULT_BALANCING_CONFIG.excludedSubjects);
    setBalancingExcludedStudentIds([]);
    setColumnMapperExpanded(false);
    setActiveDataTab('import');
    setActiveGroupTab('subjects');
    setActiveStudentTab('elever');
    setShowReloadConfirmModal(false);
    setIsReloadConfirmArmed(false);
  };

  const handleApplyBalancingResult = (result: ProgressiveHybridBalanceResult) => {
    if (result.moveRecords.length === 0 && result.diagnostics.unresolvedCollisions.length === 0) {
      return;
    }

    captureUndoSnapshot();

    const nowIso = new Date().toISOString();
    const currentRoundId = nextBalancingRoundId;
    const studentById = new Map<string, StandardField>();
    mergedData.forEach((student, index) => {
      const inferredId = student.studentId || `${student.navn || 'ukjent'}:${student.klasse || 'ukjent'}:${index}`;
      studentById.set(inferredId, student);
    });

    const changes: StudentAssignmentChange[] = result.moveRecords.map((move) => ({
      studentId: move.studentId,
      navn: move.studentName,
      klasse: studentById.get(move.studentId)?.klasse || 'Ingen klasse',
      subject: move.subjectName,
      fromBlokk: move.fromBlock,
      toBlokk: move.toBlock,
      reason: `Balansering (${move.reason}): ${move.fromGroupCode} -> ${move.toGroupCode}, scoreDelta ${move.scoreDelta.toFixed(2)}`,
      changedAt: nowIso,
      balancingRoundId: currentRoundId,
    }));

    const unresolvedWarningChanges: StudentAssignmentChange[] = result.diagnostics.unresolvedCollisions.map((entry) => {
      const separatorIndex = entry.lastIndexOf(':');
      const studentId = separatorIndex >= 0 ? entry.slice(0, separatorIndex) : entry;
      const subjectCode = separatorIndex >= 0 ? entry.slice(separatorIndex + 1) : 'UKJENT';
      const student = studentById.get(studentId);

      return {
        studentId,
        navn: student?.navn || 'Ukjent',
        klasse: student?.klasse || 'Ingen klasse',
        subject: subjectCode,
        fromBlokk: 0,
        toBlokk: 0,
        reason: `ADVARSEL: Kunne ikke plassere elev uten kollisjon for fagkode ${subjectCode}`,
        changedAt: nowIso,
        balancingRoundId: currentRoundId,
      };
    });

    setMergedData(result.updatedData);
    setSubjects(tallySubjects(result.updatedData));
    setSubjectSettingsByName(result.updatedSubjectSettingsByName as SubjectSettingsByName);
    setStudentAssignmentChanges((prev) => [...prev, ...changes, ...unresolvedWarningChanges]);
    setNextBalancingRoundId((prev) => prev + 1);
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

    captureUndoSnapshot();

    setMergedData(workingData);
    setSubjects(tallySubjects(workingData));
    setStudentAssignmentChanges((prev) => [...prev, ...allChanges]);
  };

  const handleGroupMoved = (entry: GroupMoveLogEntry) => {
    setGroupMoveLog((prev) => [...prev, entry]);
  };

  const handleClearStoredData = () => {
    localStorage.removeItem(LOCAL_STORAGE_KEY);
    handleReset();
    setBlokkCount(4);
  };

  const handleRemoveFile = (fileId: string) => {
    captureUndoSnapshot();
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

  const handleStudentListExport = (exportType: 'novaschem' | 'excel-full' | 'txt') => {
    setIsStudentExportMenuOpen(false);

    if (exportType === 'novaschem') {
      void handleExport();
      return;
    }

    if (exportType === 'excel-full') {
      void handleExportDetailed();
      return;
    }

    handleExportText();
  };

  const handleExportJson = () => {
    const exportedState: ExportedAppStateFile = {
      format: JSON_TRANSFER_FORMAT,
      version: JSON_TRANSFER_VERSION,
      exportedAt: new Date().toISOString(),
      appState: buildPersistedState(),
    };

    const blob = new Blob([JSON.stringify(exportedState, null, 2)], {
      type: 'application/json;charset=utf-8',
    });
    const downloadUrl = URL.createObjectURL(blob);
    const downloadLink = document.createElement('a');
    const timestamp = new Date().toISOString().replace(/[.:]/g, '-');

    downloadLink.href = downloadUrl;
    downloadLink.download = `inschool-balansering-${timestamp}.json`;
    downloadLink.click();

    URL.revokeObjectURL(downloadUrl);
    showJsonTransferStatus('JSON eksportert');
  };

  const handleImportJsonClick = () => {
    jsonImportInputRef.current?.click();
  };

  const handleImportJson = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    event.target.value = '';

    if (!file) {
      return;
    }

    if ((parsedFiles.length > 0 || mergedData.length > 0) && !window.confirm('Dette vil erstatte gjeldende arbeidsdata. Fortsette?')) {
      return;
    }

    try {
      const fileText = await file.text();
      const parsedJson = JSON.parse(fileText) as Partial<ExportedAppStateFile & PersistedAppState>;
      const importedState = parsedJson
        && typeof parsedJson === 'object'
        && parsedJson.format === JSON_TRANSFER_FORMAT
        && parsedJson.version === JSON_TRANSFER_VERSION
        && parsedJson.appState
        && typeof parsedJson.appState === 'object'
          ? parsedJson.appState
          : parsedJson;

      if (!importedState || typeof importedState !== 'object') {
        throw new Error('Invalid JSON state');
      }

      captureUndoSnapshot();
      applyPersistedState(importedState as Partial<PersistedAppState>);
      setIsHydratedFromStorage(true);
      showJsonTransferStatus(`Importerte ${file.name}`);
    } catch {
      showJsonTransferStatus('Kunne ikke importere JSON-filen');
    }
  };

  const handleStudentAssignmentsUpdated = (
    updatedData: StandardField[],
    changes: StudentAssignmentChange[]
  ) => {
    if (changes.length === 0) {
      return;
    }

    captureUndoSnapshot();

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

    captureUndoSnapshot();

    setMergedData(result.updatedData);
    setSubjects(tallySubjects(result.updatedData));
    setStudentAssignmentChanges((prev) => [...prev, ...result.changes]);
  };

  const handleAddSubject = (subjectName: string) => {
    const key = subjectName.trim().toLocaleLowerCase('nb');
    const alreadyExists = Object.keys(subjectSettingsByName).some(
      (name) => name.trim().toLocaleLowerCase('nb') === key
    );
    if (alreadyExists) return;

    captureUndoSnapshot();
    setSubjectSettingsByName((prev) => ({
      ...prev,
      [subjectName.trim()]: { defaultMax: 30 },
    }));
    setGroupMoveLog((prev) => [...prev, {
      subject: subjectName.trim(),
      groupLabel: '',
      fromBlokk: 0,
      toBlokk: 0,
      changedAt: new Date().toISOString(),
      action: 'subject-added',
    }]);
  };

  const handleRemoveSubject = (subjectName: string) => {
    // Remove from all students
    const allStudentIds = mergedData.map((student, index) =>
      student.studentId || `${student.navn || 'ukjent'}:${student.klasse || 'ukjent'}:${index}`
    );
    const result = removeSubjectAssignmentsForStudents(
      mergedData,
      subjectName,
      allStudentIds,
      `Fag fjernet: ${subjectName}`
    );

    captureUndoSnapshot();

    if (result.changes.length > 0) {
      setMergedData(result.updatedData);
      setSubjects(tallySubjects(result.updatedData));
      setStudentAssignmentChanges((prev) => [...prev, ...result.changes]);
    }

    // Remove from subject settings
    setSubjectSettingsByName((prev) => {
      const next = { ...prev };
      const keyLower = subjectName.trim().toLocaleLowerCase('nb');
      for (const name of Object.keys(next)) {
        if (name.trim().toLocaleLowerCase('nb') === keyLower) {
          delete next[name];
        }
      }
      return next;
    });
    setGroupMoveLog((prev) => [...prev, {
      subject: subjectName.trim(),
      groupLabel: '',
      fromBlokk: 0,
      toBlokk: 0,
      changedAt: new Date().toISOString(),
      action: 'subject-removed',
    }]);
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

  const warningEntries = useMemo(() => {
    return mergedData
      .map((student, index) => ({ student, index }))
      .filter(({ student }) => !student.removedFromElevlist)
      .map(({ student, index }) => {
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
  }, [mergedData]);

  const studentsWithFewSubjects = useMemo(() => {
    return warningEntries
      .filter((entry) => entry.subjectCount < 3)
      .sort((a, b) => (a.student.navn || '').localeCompare(b.student.navn || '', 'nb', { sensitivity: 'base' }));
  }, [warningEntries]);

  const studentsWithFourSubjects = useMemo(() => {
    return warningEntries
      .filter((entry) => entry.subjectCount >= 4)
      .sort((a, b) => (a.student.navn || '').localeCompare(b.student.navn || '', 'nb', { sensitivity: 'base' }));
  }, [warningEntries]);

  const studentsWithBlokkCollisions = useMemo(() => {
    return warningEntries
      .filter((entry) => entry.hasBlokkCollision)
      .sort((a, b) => (a.student.navn || '').localeCompare(b.student.navn || '', 'nb', { sensitivity: 'base' }));
  }, [warningEntries]);

  const activeStudentsWithFewSubjects = useMemo(() => {
    return studentsWithFewSubjects.filter((entry) => !isWarningIgnored(entry.studentId, 'missing'));
  }, [studentsWithFewSubjects, warningIgnoresByStudentAndType]);

  const ignoredStudentsWithFewSubjects = useMemo(() => {
    return studentsWithFewSubjects.filter((entry) => isWarningIgnored(entry.studentId, 'missing'));
  }, [studentsWithFewSubjects, warningIgnoresByStudentAndType]);

  const activeStudentsWithFourSubjects = useMemo(() => {
    return studentsWithFourSubjects.filter((entry) => !isWarningIgnored(entry.studentId, 'overloaded'));
  }, [studentsWithFourSubjects, warningIgnoresByStudentAndType]);

  const ignoredStudentsWithFourSubjects = useMemo(() => {
    return studentsWithFourSubjects.filter((entry) => isWarningIgnored(entry.studentId, 'overloaded'));
  }, [studentsWithFourSubjects, warningIgnoresByStudentAndType]);

  const activeStudentsWithBlokkCollisions = studentsWithBlokkCollisions;

  const fewSubjectsIgnoredCount = ignoredStudentsWithFewSubjects.length;

  const fourSubjectsIgnoredCount = ignoredStudentsWithFourSubjects.length;
  const hasActiveWarnings = activeStudentsWithFewSubjects.length > 0
    || activeStudentsWithFourSubjects.length > 0
    || activeStudentsWithBlokkCollisions.length > 0;

  const warningStudentIds = useMemo(() => {
    return new Set<string>([
      ...studentsWithFewSubjects.map((entry) => entry.studentId),
      ...studentsWithFourSubjects.map((entry) => entry.studentId),
    ]);
  }, [studentsWithFewSubjects, studentsWithFourSubjects]);

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

    if (warningIgnoresByStudentAndType[studentId]?.[type]?.comment === comment) {
      return;
    }

    captureUndoSnapshot();

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
    if (!warningIgnoresByStudentAndType[studentId]?.[type]) {
      return;
    }

    captureUndoSnapshot();

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

  const filteredMergedData = useMemo(() => {
    if (!selectedMergedSubject) {
      return mergedData;
    }

    return mergedData.filter((student) => {
      return (
        matchesSelectedSubject(student.blokk1, selectedMergedSubject)
        || matchesSelectedSubject(student.blokk2, selectedMergedSubject)
        || matchesSelectedSubject(student.blokk3, selectedMergedSubject)
        || matchesSelectedSubject(student.blokk4, selectedMergedSubject)
      );
    });
  }, [mergedData, selectedMergedSubject]);

  const subjectOptions = useMemo(() => {
    return subjects.map((subject) => subject.subject);
  }, [subjects]);

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

  const hasImportSession = parsedFiles.length > 0 || mergedData.length > 0;
  const hasLoadedData = mergedData.length > 0;
  const summarizedChangeLogCounts = useMemo(() => {
    return getSummarizedChangeLogCounts(studentAssignmentChanges);
  }, [studentAssignmentChanges]);

  return (
    <div className="app">
      <main className="main">
        <div className="content-container">
          <header className="header">
            <div className="header-title-row">
              <h1>Fagvalg - Oversikt og balanseringsverktøy</h1>
              <button
                type="button"
                className="header-help-btn"
                onClick={() => setShowHelpModal(true)}
                aria-haspopup="dialog"
                aria-expanded={showHelpModal}
              >
                Hjelp
              </button>
            </div>
            <p>Laget av Sascha Njaa Tjelta</p>
          </header>

          <>
            {hasLoadedData && (
              <div className="warning-section-group">
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
                {hasActiveWarnings && <hr className="warning-section-divider" />}
              </div>
              )}
            
            <div className="control-row-group">
              <div className="control-row-label">Visning</div>
              <div className="data-tab-bar">
                <div className="data-tabs" role="tablist" aria-label="Data visning">
                  <button
                    type="button"
                    role="tab"
                    aria-selected={activeDataTab === 'import'}
                    className={`data-tab ${activeDataTab === 'import' ? 'data-tab-active' : ''}`.trim()}
                    onClick={() => setActiveDataTab('import')}
                  >
                    Data
                  </button>
                  {hasLoadedData && (
                    <>
                      <button
                        type="button"
                        role="tab"
                        aria-selected={activeDataTab === 'subjects' || activeDataTab === 'groups' || activeDataTab === 'fagoversikt'}
                        className={`data-tab ${activeDataTab === 'subjects' || activeDataTab === 'groups' || activeDataTab === 'fagoversikt' ? 'data-tab-active' : ''}`.trim()}
                        onClick={() => setActiveDataTab(activeGroupTab)}
                      >
                        Grupper
                      </button>
                      <button
                        type="button"
                        role="tab"
                        aria-selected={activeDataTab === 'elever' || activeDataTab === 'students'}
                        className={`data-tab ${activeDataTab === 'elever' || activeDataTab === 'students' ? 'data-tab-active' : ''}`.trim()}
                        onClick={() => setActiveDataTab(activeStudentTab)}
                      >
                        Elever
                      </button>
                      <button
                        type="button"
                        role="tab"
                        aria-selected={activeDataTab === 'balancing'}
                        className={`data-tab ${activeDataTab === 'balancing' ? 'data-tab-active' : ''}`.trim()}
                        onClick={() => setActiveDataTab('balancing')}
                      >
                        Balansering
                      </button>
                      <button
                        type="button"
                        role="tab"
                        aria-selected={activeDataTab === 'changelog'}
                        className={`data-tab ${activeDataTab === 'changelog' ? 'data-tab-active' : ''}`.trim()}
                        onClick={() => setActiveDataTab('changelog')}
                      >
                        Logg ({summarizedChangeLogCounts.students} elever, {summarizedChangeLogCounts.changes} endringer)
                      </button>
                    </>
                  )}
                </div>
                <div className="data-tab-actions">
                  <button
                    type="button"
                    className="export-btn tabline-action-btn history-btn"
                    disabled={undoHistory.length === 0}
                    onClick={handleUndo}
                    aria-label="Angre"
                    title={undoHistory.length > 0 ? `Angre siste endring (${undoHistory.length}/${MAX_HISTORY_STATES})` : 'Ingen endringer å angre'}
                  >
                    ↶
                  </button>
                  <button
                    type="button"
                    className="export-btn tabline-action-btn history-btn"
                    disabled={redoHistory.length === 0}
                    onClick={handleRedo}
                    aria-label="Gjør om"
                    title={redoHistory.length > 0 ? `Gjør om siste angrede endring (${redoHistory.length}/${MAX_HISTORY_STATES})` : 'Ingen endringer å gjøre om'}
                  >
                    ↷
                  </button>
                  <div className="tabline-menu" ref={studentExportMenuRef}>
                    <button
                      type="button"
                      className="export-btn tabline-action-btn"
                      disabled={!hasLoadedData}
                      onClick={() => setIsStudentExportMenuOpen((prev) => !prev)}
                      aria-expanded={isStudentExportMenuOpen}
                      aria-haspopup="menu"
                      title={hasLoadedData ? 'Eksporter elevliste i valgt format' : 'Last inn data først'}
                    >
                      Eksporter elevliste {isStudentExportMenuOpen ? '▲' : '▼'}
                    </button>
                    {isStudentExportMenuOpen && hasLoadedData && (
                      <div className="tabline-menu-popover" role="menu" aria-label="Eksporter elevliste">
                        <button
                          type="button"
                          className="tabline-menu-item"
                          role="menuitem"
                          onClick={() => handleStudentListExport('excel-full')}
                        >
                          Eksporter til Excel (full)
                        </button>
                        <button
                          type="button"
                          className="tabline-menu-item"
                          role="menuitem"
                          onClick={() => handleStudentListExport('novaschem')}
                        >
                          Eksporter til Novaschem (.xlsx)
                        </button>
                        <button
                          type="button"
                          className="tabline-menu-item"
                          role="menuitem"
                          onClick={() => handleStudentListExport('txt')}
                        >
                          Eksporter til Novaschem (.txt)
                        </button>
                      </div>
                    )}
                  </div>
                  <button
                    type="button"
                    className="export-btn tabline-action-btn"
                    disabled={!hasImportSession}
                    onClick={handleExportJson}
                    title={hasImportSession ? 'Eksporter arbeidsdata til JSON' : 'Last inn eller importer data først'}
                  >
                    Eksporter .JSON
                  </button>
                  <button
                    type="button"
                    className="export-btn tabline-action-btn"
                    onClick={handleImportJsonClick}
                    title="Importer arbeidsdata fra JSON"
                  >
                    Importer .JSON
                  </button>
                  {jsonTransferStatus && <span className="tabline-status">{jsonTransferStatus}</span>}
                  <input
                    ref={jsonImportInputRef}
                    type="file"
                    accept=".json,application/json"
                    className="visually-hidden"
                    onChange={handleImportJson}
                  />
                </div>
              </div>
              {hasLoadedData && (activeDataTab === 'subjects' || activeDataTab === 'groups' || activeDataTab === 'fagoversikt') && (
                  <div className="data-tabs data-subtabs" role="tablist" aria-label="Grupper visning">
                    <button
                      type="button"
                      role="tab"
                      aria-selected={activeDataTab === 'subjects'}
                      className={`data-tab ${activeDataTab === 'subjects' ? 'data-tab-active' : ''}`.trim()}
                      onClick={() => {
                        setActiveGroupTab('subjects');
                        setActiveDataTab('subjects');
                      }}
                    >
                      Blokkoversikt
                    </button>
                    <button
                      type="button"
                      role="tab"
                      aria-selected={activeDataTab === 'groups'}
                      className={`data-tab ${activeDataTab === 'groups' ? 'data-tab-active' : ''}`.trim()}
                      onClick={() => {
                        setActiveGroupTab('groups');
                        setActiveDataTab('groups');
                      }}
                    >
                      Gruppeoversikt
                    </button>
                    <button
                      type="button"
                      role="tab"
                      aria-selected={activeDataTab === 'fagoversikt'}
                      className={`data-tab ${activeDataTab === 'fagoversikt' ? 'data-tab-active' : ''}`.trim()}
                      onClick={() => {
                        setActiveGroupTab('fagoversikt');
                        setActiveDataTab('fagoversikt');
                      }}
                    >
                      Fagoversikt
                    </button>
                  </div>
                )}
              {hasLoadedData && (activeDataTab === 'elever' || activeDataTab === 'students') && (
                  <div className="data-tabs data-subtabs" role="tablist" aria-label="Elever visning">
                    <button
                      type="button"
                      role="tab"
                      aria-selected={activeDataTab === 'elever'}
                      className={`data-tab ${activeDataTab === 'elever' ? 'data-tab-active' : ''}`.trim()}
                      onClick={() => {
                        setActiveStudentTab('elever');
                        setActiveDataTab('elever');
                      }}
                    >
                      Elevoversikt
                    </button>
                    <button
                      type="button"
                      role="tab"
                      aria-selected={activeDataTab === 'students'}
                      className={`data-tab ${activeDataTab === 'students' ? 'data-tab-active' : ''}`.trim()}
                      onClick={() => {
                        setActiveStudentTab('students');
                        setActiveDataTab('students');
                      }}
                    >
                      Elevtabell
                    </button>
                  </div>
                )}
            </div>

            <div className="data-tab-panel">
              {!hasLoadedData || activeDataTab === 'import' ? (
                <>
                  <FileUploader onFilesAdded={handleFilesAdded} anonymize={anonymize} onAnonymizeChange={setAnonymize} />

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
                            onBlokkCountChange={handleBlokkCountChange}
                          />
                        )}
                      </div>
                    </>
                  )}

                  <div className="action-buttons">
                    <button onClick={handleLoadDataClick} className="load-btn" disabled={parsedFiles.length === 0}>
                      Last inn data
                    </button>
                    <button onClick={handleClearStoredData} className="clear-storage-btn">
                      Tøm data
                    </button>
                  </div>
                </>
              ) : activeDataTab === 'subjects' ? (
                <SubjectTally
                  subjects={subjects}
                  mergedData={mergedData}
                  blokkCount={blokkCount}
                  subjectSettingsByName={subjectSettingsByName}
                  autoOpenSettingsToken={
                    subjectSettingsAutoOpenToken > subjectSettingsAutoOpenHandledToken
                      ? subjectSettingsAutoOpenToken
                      : undefined
                  }
                  onAutoOpenSettingsHandled={(token) => {
                    setSubjectSettingsAutoOpenHandledToken((prev) => (token > prev ? token : prev));
                  }}
                  onSaveSubjectSettingsByName={handleSaveSubjectSettingsByName}
                  onApplySubjectBlockMoves={handleApplySubjectBlockMoves}
                  onGroupMoved={handleGroupMoved}
                  onRemoveStudentsFromSubject={handleRemoveStudentsFromSubject}
                  onOpenStudentCard={handleOpenStudentInElever}
                  onAddSubject={handleAddSubject}
                  onRemoveSubject={handleRemoveSubject}
                />
              ) : activeDataTab === 'groups' ? (
                <GrupperView
                  data={mergedData}
                  blokkCount={blokkCount}
                  subjectOptions={subjectOptions}
                  subjectSettingsByName={subjectSettingsByName}
                  classBlockRestrictions={classBlockRestrictions}
                  changeLog={studentAssignmentChanges}
                  onSaveSubjectSettingsByName={handleSaveSubjectSettingsByName}
                  onStudentDataUpdate={handleStudentAssignmentsUpdated}
                  onOpenStudentCard={handleOpenStudentInElever}
                />
              ) : activeDataTab === 'fagoversikt' ? (
                <FagoversiktView
                  data={mergedData}
                  blokkCount={blokkCount}
                  subjectSettingsByName={subjectSettingsByName}
                  onOpenStudentCard={handleOpenStudentInElever}
                />
              ) : activeDataTab === 'changelog' ? (
                <ChangeLogView
                  changeLog={studentAssignmentChanges}
                  groupMoveLog={groupMoveLog}
                  currentStudents={mergedData}
                  subjectSettingsByName={subjectSettingsByName}
                  excludedSubjects={balancingExcludedSubjects}
                  onOpenStudentCard={handleOpenStudentInElever}
                />
              ) : activeDataTab === 'balancing' ? (
                <BalanseringView
                  mergedData={mergedData}
                  blokkCount={blokkCount}
                  subjectSettingsByName={subjectSettingsByName}
                  restrictions={classBlockRestrictions}
                  excludedSubjects={balancingExcludedSubjects}
                  excludedStudentIds={balancingExcludedStudentIds}
                  onExcludedSubjectsChange={setBalancingExcludedSubjects}
                  onExcludedStudentIdsChange={setBalancingExcludedStudentIds}
                  onRestrictionsChange={handleClassBlockRestrictionsChange}
                  onApplyResult={handleApplyBalancingResult}
                />
              ) : activeDataTab === 'students' ? (
                <MergedDataView
                  data={filteredMergedData}
                  totalDataCount={mergedData.length}
                  selectedSubject={selectedMergedSubject}
                  onSubjectFilterChange={setSelectedMergedSubject}
                  subjectOptions={subjectOptions}
                  blokkCount={blokkCount}
                />
              ) : (
                <EleverView
                  data={mergedData}
                  blokkCount={blokkCount}
                  subjectOptions={subjectOptions}
                  subjectSettingsByName={subjectSettingsByName}
                  onSaveSubjectSettingsByName={handleSaveSubjectSettingsByName}
                  warningIgnoresByStudentAndType={warningIgnoresByStudentAndType}
                  onSaveWarningIgnore={(studentId, type, comment) => saveWarningIgnore(studentId, type, comment)}
                  onRemoveWarningIgnore={removeWarningIgnore}
                  changeLog={studentAssignmentChanges}
                  onStudentDataUpdate={handleStudentAssignmentsUpdated}
                  externallySelectedStudentId={selectedEleverStudentId}
                  onExternalSelectionHandled={() => setSelectedEleverStudentId('')}
                  activationToken={eleverViewActivationToken}
                />
              )}
            </div>

            {showHelpModal && (
              <div className="app-modal-overlay" onClick={() => setShowHelpModal(false)}>
                <div
                  className="app-help-modal"
                  role="dialog"
                  aria-modal="true"
                  aria-labelledby="help-modal-title"
                  onClick={(event) => event.stopPropagation()}
                >
                  <div className="app-help-header">
                    <h3 id="help-modal-title">Hjelp - Slik bruker du verktøyet</h3>
                    <button
                      type="button"
                      className="app-help-close"
                      onClick={() => setShowHelpModal(false)}
                      aria-label="Lukk hjelp"
                    >
                      Lukk
                    </button>
                  </div>

                  <div className="app-help-content">
                    <section>
                      <h4>1. Start med data</h4>
                      <p>
                        Gå til fanen <strong>Data</strong>. Last opp en eller flere filer. I <strong>Oppsett</strong> kan du sjekke at
                        kolonner er riktig mappet (navn, klasse og blokker). Når alt ser riktig ut, trykk <strong>Last inn data</strong>.
                      </p>
                    </section>

                    <section>
                      <h4>2. Se advarsler tidlig</h4>
                      <p>
                        Etter innlasting vises advarsler øverst. Her ser du elever med blokk-kollisjon, færre enn 3 fag,
                        eller 4+ fag. Du kan eksportere advarsler, kopiere dem, og hoppe direkte til eleven i elevvisningen.
                      </p>
                    </section>

                    <section>
                      <h4>3. Sett opp grupper i lineær flyt</h4>
                      <p>
                        Gå til <strong>Grupper</strong> og jobb først i <strong>Blokkoversikt</strong>: juster kapasitet, opprett eller fjern grupper,
                        og kontroller fordeling per blokk. Bytt deretter til <strong>Gruppeoversikt</strong> for å finjustere flytting av elever
                        mellom grupper med full oversikt over konsekvenser.
                      </p>
                    </section>

                    <section>
                      <h4>4. Korriger elevdata</h4>
                      <p>
                        I <strong>Elever</strong> kan du redigere fag per elev, legge til nye elever, markere <strong>Fjerdeårs-elev</strong>,
                        og følge elevens endringslogg. Manuelt lagt til elev fungerer videre i samme pipeline som importerte elever.
                      </p>
                    </section>

                    <section>
                      <h4>5. Kjør balansering</h4>
                      <p>
                        I <strong>Balansering</strong> setter du tillatte blokkregler per trinn og eventuelle fag som skal holdes utenfor.
                        Kjør balansering når gruppestrukturen er klar. Bruk diagnosefeltet for å se hva som ble forbedret,
                        og kjør flere runder ved behov.
                      </p>
                    </section>

                    <section>
                      <h4>6. Kontroller resultat i Logg</h4>
                      <p>
                        I <strong>Logg</strong> ser du detaljert eller oppsummert historikk per elev. Du kan søke på elevnavn og fag,
                        og eksportere den filtrerte loggen til Word.
                      </p>
                    </section>

                    <section>
                      <h4>7. Eksporter og del</h4>
                      <p>
                        Bruk knappen <strong>Eksporter elevliste</strong> for Excel/TXT. Bruk <strong>Eksporter .JSON</strong> for å lagre hele arbeidsøkten,
                        slik at du eller kollegaer kan importere videre senere.
                      </p>
                    </section>

                    <section>
                      <h4>Tips for trygg arbeidsflyt</h4>
                      <p>
                        Jobb i rekkefolgen Data, Grupper, Elever, Balansering, Logg og Eksport. Bruk angre/gjør om
                        i topplinjen hvis du vil teste alternative valg uten å miste kontroll på prosessen.
                      </p>
                    </section>
                  </div>
                </div>
              </div>
            )}

            {showReloadConfirmModal && (
              <div
                className="app-modal-overlay"
                onClick={() => {
                  if (isReloadConfirmArmed) {
                    setIsReloadConfirmArmed(false);
                    return;
                  }

                  closeReloadConfirmModal();
                }}
              >
                <div
                  className="app-confirm-modal"
                  onClick={(event) => {
                    event.stopPropagation();

                    if (isReloadConfirmArmed) {
                      setIsReloadConfirmArmed(false);
                    }
                  }}
                >
                  <h4>Last inn data på nytt?</h4>
                  <p className="app-confirm-message">
                    Dette vil overstyre alle endringer som er gjort.
                  </p>
                  <div className="app-confirm-actions">
                    <button
                      type="button"
                      className="app-confirm-btn app-confirm-secondary"
                      onClick={(event) => {
                        event.stopPropagation();
                        closeReloadConfirmModal();
                      }}
                    >
                      Nei
                    </button>
                    <button
                      type="button"
                      className={`app-confirm-btn app-confirm-primary ${
                        isReloadConfirmArmed ? 'app-confirm-armed' : ''
                      }`.trim()}
                      onClick={(event) => {
                        event.stopPropagation();

                        if (isReloadConfirmArmed) {
                          handleMerge();
                          closeReloadConfirmModal();
                          return;
                        }

                        setIsReloadConfirmArmed(true);
                      }}
                    >
                      {isReloadConfirmArmed ? 'Bekreft' : 'Ja'}
                    </button>
                  </div>
                </div>
              </div>
            )}
          </>
        </div>
      </main>
    </div>
  );
}

export default App;
