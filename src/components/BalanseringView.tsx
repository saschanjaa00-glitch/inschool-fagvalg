import { useEffect, useMemo, useRef, useState } from 'react';
import type { StandardField } from '../utils/excelUtils';
import {
  DEFAULT_BALANCING_CONFIG,
  DEFAULT_CLASS_BLOCK_RESTRICTIONS,
  type BalancingConfig,
  type BalancingWeights,
  type BlockNumber,
  type ClassBlockRestrictions,
  type ProgressiveHybridBalanceResult,
  type SubjectSettingsByNameLike,
} from '../utils/progressiveHybridBalance';
import { mapSubjectToCode } from '../utils/subjectCodeMapping';
import type { BalancingWorkerInbound, BalancingWorkerOutbound } from '../workers/progressiveHybridBalance.worker.types';
import styles from './BalanseringView.module.css';

interface BalanseringViewProps {
  mergedData: StandardField[];
  subjectSettingsByName: SubjectSettingsByNameLike;
  restrictions: ClassBlockRestrictions;
  excludedSubjects: string[];
  excludedStudentIds: string[];
  onExcludedSubjectsChange: (subjects: string[]) => void;
  onExcludedStudentIdsChange: (studentIds: string[]) => void;
  onRestrictionsChange: (value: ClassBlockRestrictions) => void;
  onApplyResult: (result: ProgressiveHybridBalanceResult) => void;
}

const DEFAULT_CLASS_LEVELS = ['VG1', 'VG2', 'VG3'];
const FIXED_COLLISION_WEIGHT = DEFAULT_BALANCING_CONFIG.weights.collisionD;
const SETTING_DESCRIPTIONS = {
  overcapA: 'Straffer grupper som ligger over maks kapasitet. Høyere verdi gjør at overfylte grupper prioriteres hardere.',
  imbalanceB: 'Straffer skjev fordeling mellom grupper i samme fag. Høyere verdi gir jevnere gruppestørrelser.',
  peakC: 'Straffer spesielt store toppgrupper. Høyere verdi presser ned de største gruppene raskere.',
  collisionD: 'Straff for blokkollisjon. Denne er låst svært høyt slik at kollisjoner i praksis forbys.',
  movesE: 'Straffer antall flytt som faktisk endrer blokk. Omfordeling mellom grupper i samme blokk er gratis.',
  repeatF: 'Straffer å flytte samme elev flere ganger når flyttene endrer blokk. Omfordeling innen samme blokk er gratis.',
  alpha: 'Øker hvor mye store grupper påvirker ubalansestraffen.',
  beta: 'Øker hvor hardt de største gruppene teller i topptrykk-beregningen.',
  maxRelaxation: 'Starter balanseringen så mange plasser under virkelig maks før den gradvis nærmer seg faktisk maks.',
  maxFlowIterationsPerOffset: 'Maks antall flyt-iterasjoner som kjores per kapasitetsniva (deterministisk stoppkriterium).',
  maxLookaheadAttempts: 'Hvor mange lookahead-forsøk motoren kan bruke for å finne kjeder av flytt.',
  maxDepth2Chains: 'Maks antall dypere to-stegs kjeder som prøves i lookahead.',
  restrictions: 'Klassebegrensninger styrer hvilke blokker hvert trinn har lov til å bruke.',
  excludedSubjects: 'Utelukkede fag blir ikke flyttet og teller ikke i balanseringskostnader, men de opptar fortsatt blokker og vises ellers i appen.',
  excludedStudents: 'Valgte klasser eller elever blir laast og kan ikke flyttes av balanseringen.',
} as const;

const formatNumber = (value: number): string => {
  return Number.isFinite(value) ? value.toFixed(2) : String(value);
};

const parseInputNumber = (value: string, fallback: number): number => {
  const parsed = Number.parseFloat(value);
  return Number.isNaN(parsed) ? fallback : parsed;
};

const parseSubjects = (value: string | null): string[] => {
  if (!value) {
    return [];
  }

  return value
    .split(/[,;]/)
    .map((part) => part.trim())
    .filter((part) => part.length > 0);
};

const inferStudentId = (row: StandardField, index: number): string => {
  return row.studentId || `${row.navn || 'ukjent'}:${row.klasse || 'ukjent'}:${index}`;
};

const ALL_BLOCKS: BlockNumber[] = [1, 2, 3, 4];

const normalizeRestrictions = (input: ClassBlockRestrictions): ClassBlockRestrictions => {
  const next: ClassBlockRestrictions = {
    VG1: { 1: false, 2: false, 3: false, 4: false },
    VG2: { 1: false, 2: false, 3: false, 4: false },
    VG3: { 1: false, 2: false, 3: false, 4: false },
  };

  Object.entries(DEFAULT_CLASS_BLOCK_RESTRICTIONS).forEach(([classKey, map]) => {
    next[classKey] = {
      ...(next[classKey] || {}),
      ...(map || {}),
    };
  });

  Object.entries(input).forEach(([classKey, map]) => {
    next[classKey] = {
      ...(next[classKey] || {}),
      ...(map || {}),
    };
  });
  return next;
};

const inferClassLevels = (row: StandardField): string[] => {
  if (row.fjerdearsElev) {
    return ['VG2', 'VG3'];
  }

  const classGroup = row.klasse;
  const normalized = (classGroup || '').trim().toUpperCase();
  if (normalized.length === 0) {
    return [];
  }

  const match = normalized.match(/^(\d)/);
  if (!match) {
    return DEFAULT_CLASS_LEVELS.includes(normalized) ? [normalized] : [];
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

type BalancePresetMode = 'even' | 'underMax' | 'advanced';

const EVEN_PRESET_MAX_RELAXATION = 20;
const UNDER_MAX_PRESET_MAX_RELAXATION = 6;
const EVEN_BALANCE_OFFSETS: number[] = [20, 15, 10, 8, 6, 4, 2, 0];

export const BalanseringView = ({
  mergedData,
  subjectSettingsByName,
  restrictions,
  excludedSubjects,
  excludedStudentIds,
  onExcludedSubjectsChange,
  onExcludedStudentIdsChange,
  onRestrictionsChange,
  onApplyResult,
}: BalanseringViewProps) => {
  const [weights, setWeights] = useState<BalancingWeights>(DEFAULT_BALANCING_CONFIG.weights);
  const [maxRelaxation, setMaxRelaxation] = useState(String(EVEN_PRESET_MAX_RELAXATION));
  const [maxFlowIterationsPerOffset, setMaxFlowIterationsPerOffset] = useState(
    String(DEFAULT_BALANCING_CONFIG.maxFlowIterationsPerOffset)
  );
  const [maxLookaheadAttempts, setMaxLookaheadAttempts] = useState(String(DEFAULT_BALANCING_CONFIG.maxLookaheadAttempts));
  const [maxDepth2Chains, setMaxDepth2Chains] = useState(String(DEFAULT_BALANCING_CONFIG.maxDepth2Chains));
  const [presetMode, setPresetMode] = useState<BalancePresetMode>('even');
  const [parametersExpanded, setParametersExpanded] = useState(false);
  const [excludedSubjectsExpanded, setExcludedSubjectsExpanded] = useState(false);
  const [excludedStudentsExpanded, setExcludedStudentsExpanded] = useState(false);
  const [expandedClassGroups, setExpandedClassGroups] = useState<Set<string>>(new Set());
  const [diagnosticsExpanded, setDiagnosticsExpanded] = useState(false);
  const [statusMessage, setStatusMessage] = useState('');
  const [lastResult, setLastResult] = useState<ProgressiveHybridBalanceResult | null>(null);
  const [isBalancing, setIsBalancing] = useState(false);
  const workerRef = useRef<Worker | null>(null);
  const activeRequestIdRef = useRef(0);

  const effectiveRestrictions = useMemo(() => normalizeRestrictions(restrictions), [restrictions]);
  const visibleClassLevels = useMemo(() => {
    const levels = new Set<string>();

    mergedData.forEach((row) => {
      inferClassLevels(row).forEach((classLevel) => levels.add(classLevel));
    });

    const orderedLevels = DEFAULT_CLASS_LEVELS.filter((level) => levels.has(level));
    return orderedLevels.length > 0 ? orderedLevels : DEFAULT_CLASS_LEVELS.filter((level) => level !== 'VG1');
  }, [mergedData]);

  const availableSubjects = useMemo(() => {
    const subjects = new Set<string>();

    mergedData.forEach((row) => {
      [row.blokk1, row.blokk2, row.blokk3, row.blokk4].forEach((value) => {
        parseSubjects(value).forEach((subject) => subjects.add(subject));
      });
    });

    Object.keys(subjectSettingsByName).forEach((subject) => subjects.add(subject));

    return Array.from(subjects).sort((left, right) => left.localeCompare(right, 'nb', { sensitivity: 'base' }));
  }, [mergedData, subjectSettingsByName]);

  const studentsByClass = useMemo(() => {
    const byClass = new Map<string, Array<{ studentId: string; name: string }>>();

    mergedData.forEach((row, index) => {
      if (row.removedFromElevlist) {
        return;
      }

      const classGroup = (row.klasse || '').trim().toUpperCase();
      if (!classGroup) {
        return;
      }

      const studentId = inferStudentId(row, index);
      const name = (row.navn || 'Ukjent').trim() || 'Ukjent';
      const current = byClass.get(classGroup) || [];
      current.push({ studentId, name });
      byClass.set(classGroup, current);
    });

    return Array.from(byClass.entries())
      .map(([classGroup, students]) => ({
        classGroup,
        students: students.sort((left, right) => left.name.localeCompare(right.name, 'nb', { sensitivity: 'base' })),
      }))
      .sort((left, right) => left.classGroup.localeCompare(right.classGroup, 'nb', { sensitivity: 'base' }));
  }, [mergedData]);

  const lockableStudentIds = useMemo(() => {
    const ids = new Set<string>();
    studentsByClass.forEach((entry) => {
      entry.students.forEach((student) => ids.add(student.studentId));
    });
    return ids;
  }, [studentsByClass]);

  useEffect(() => {
    const normalizedStudentIds = excludedStudentIds.filter((studentId) => lockableStudentIds.has(studentId));
    if (normalizedStudentIds.length !== excludedStudentIds.length) {
      onExcludedStudentIdsChange(normalizedStudentIds);
    }
  }, [excludedStudentIds, lockableStudentIds, onExcludedStudentIdsChange]);

  const hasAnyAllowedRestriction = useMemo(() => {
    return visibleClassLevels.some((classKey) => {
      return ALL_BLOCKS.some((block) => effectiveRestrictions[classKey]?.[block] ?? false);
    });
  }, [effectiveRestrictions, visibleClassLevels]);

  useEffect(() => {
    if (!hasAnyAllowedRestriction) {
      setParametersExpanded(false);
      setExcludedSubjectsExpanded(false);
    }
  }, [hasAnyAllowedRestriction]);

  useEffect(() => {
    const worker = new Worker(new URL('../workers/progressiveHybridBalance.worker.ts', import.meta.url), {
      type: 'module',
    });

    workerRef.current = worker;

    const handleMessage = (event: MessageEvent<BalancingWorkerOutbound>) => {
      const message = event.data;
      if (!message || message.requestId !== activeRequestIdRef.current) {
        return;
      }

      if (message.type === 'error') {
        setStatusMessage(`Balansering feilet: ${message.message}`);
        setIsBalancing(false);
        return;
      }

      const result = message.result;
      setLastResult(result);
      onApplyResult(result);

      const unresolvedCollisionCount = result.diagnostics.unresolvedCollisions.length;
      setStatusMessage(
        `Kjort ferdig: ${result.diagnostics.moveCount} flytt, ${result.diagnostics.uniqueStudentsMoved} elever, score ${formatNumber(
          result.diagnostics.beforeScore.total
        )} -> ${formatNumber(result.diagnostics.afterScore.total)}${
          unresolvedCollisionCount > 0
            ? `, ADVARSEL: ${unresolvedCollisionCount} elevfag kan ikke plasseres uten kollisjon (se logg)`
            : ''
        }`
      );
      setIsBalancing(false);
    };

    const handleError = () => {
      setStatusMessage('Balanseringsarbeideren krasjet. Prover igjen kan hjelpe.');
      setIsBalancing(false);
    };

    worker.addEventListener('message', handleMessage);
    worker.addEventListener('error', handleError);

    return () => {
      worker.removeEventListener('message', handleMessage);
      worker.removeEventListener('error', handleError);
      worker.terminate();
      workerRef.current = null;
    };
  }, [onApplyResult]);

  const runBalancing = () => {
    if (isBalancing) {
      return;
    }

    if (mergedData.length === 0) {
      setStatusMessage('Ingen elevdata a balansere. Last inn data forst.');
      return;
    }

    if (!workerRef.current) {
      setStatusMessage('Kunne ikke starte balanseringsarbeider. Last siden pa nytt.');
      return;
    }

    const parsedMaxRelaxation = Math.max(
      0,
      Math.floor(parseInputNumber(maxRelaxation, DEFAULT_BALANCING_CONFIG.maxRelaxation))
    );

    const effectiveMaxRelaxation =
      presetMode === 'even'
        ? EVEN_PRESET_MAX_RELAXATION
        : presetMode === 'underMax'
          ? UNDER_MAX_PRESET_MAX_RELAXATION
          : parsedMaxRelaxation;

    const capacityOffsets = presetMode === 'even' ? EVEN_BALANCE_OFFSETS : undefined;

    const selectedStudentIds = new Set(excludedStudentIds);
    const lockKeys = new Set<string>();

    mergedData.forEach((row, index) => {
      if (row.removedFromElevlist) {
        return;
      }

      const studentId = inferStudentId(row, index);
      const isExcludedByStudent = selectedStudentIds.has(studentId);
      if (!isExcludedByStudent) {
        return;
      }

      const subjects = new Set<string>([
        ...parseSubjects(row.blokk1),
        ...parseSubjects(row.blokk2),
        ...parseSubjects(row.blokk3),
        ...parseSubjects(row.blokk4),
      ]);

      subjects.forEach((subjectName) => {
        const subjectCode = mapSubjectToCode(subjectName);
        lockKeys.add(`${studentId}|${subjectCode}`);
      });
    });

    const config: Partial<BalancingConfig> = {
      weights: {
        ...weights,
        collisionD: FIXED_COLLISION_WEIGHT,
      },
      maxRelaxation: effectiveMaxRelaxation,
      capacityOffsets,
      maxFlowIterationsPerOffset: Math.max(
        1,
        Math.floor(
          parseInputNumber(
            maxFlowIterationsPerOffset,
            DEFAULT_BALANCING_CONFIG.maxFlowIterationsPerOffset
          )
        )
      ),
      maxLookaheadAttempts: Math.max(
        0,
        Math.floor(parseInputNumber(maxLookaheadAttempts, DEFAULT_BALANCING_CONFIG.maxLookaheadAttempts))
      ),
      maxDepth2Chains: Math.max(0, Math.floor(parseInputNumber(maxDepth2Chains, DEFAULT_BALANCING_CONFIG.maxDepth2Chains))),
      classBlockRestrictions: effectiveRestrictions,
      excludedSubjects,
      lockedAssignmentKeys: Array.from(lockKeys),
    };

    const requestId = activeRequestIdRef.current + 1;
    activeRequestIdRef.current = requestId;
    setIsBalancing(true);
    setStatusMessage('Balanserer...');

    const message: BalancingWorkerInbound = {
      type: 'run',
      requestId,
      payload: {
        rows: mergedData,
        subjectSettingsByName,
        config,
      },
    };

    workerRef.current.postMessage(message);
  };

  const updateRestriction = (classKey: string, block: BlockNumber, allowed: boolean) => {
    const next = {
      ...effectiveRestrictions,
      [classKey]: {
        ...(effectiveRestrictions[classKey] || {}),
        [block]: allowed,
      },
    };

    onRestrictionsChange(next);
  };

  const allowAll = () => {
    const next = visibleClassLevels.reduce<ClassBlockRestrictions>((acc, classKey) => {
      acc[classKey] = { 1: true, 2: true, 3: true, 4: true };
      return acc;
    }, {});

    onRestrictionsChange(next);
  };

  const toggleExcludedSubject = (subject: string, excluded: boolean) => {
    if (excluded) {
      onExcludedSubjectsChange(
        [...excludedSubjects, subject]
          .filter((value, index, arr) => arr.indexOf(value) === index)
          .sort((left, right) => left.localeCompare(right, 'nb', { sensitivity: 'base' }))
      );
      return;
    }

    onExcludedSubjectsChange(excludedSubjects.filter((item) => item !== subject));
  };

  const clearExcludedSubjects = () => {
    onExcludedSubjectsChange([]);
  };

  const toggleExcludedClassGroup = (classGroup: string, excluded: boolean) => {
    const classStudentIds = studentsByClass
      .find((entry) => entry.classGroup === classGroup)
      ?.students.map((student) => student.studentId) || [];

    if (excluded) {
      onExcludedStudentIdsChange(
        [...excludedStudentIds, ...classStudentIds]
          .filter((value, index, arr) => arr.indexOf(value) === index)
      );
      return;
    }

    const classStudentIdSet = new Set(classStudentIds);
    onExcludedStudentIdsChange(excludedStudentIds.filter((value) => !classStudentIdSet.has(value)));
  };

  const toggleExcludedStudent = (studentId: string, excluded: boolean) => {
    if (excluded) {
      onExcludedStudentIdsChange([...excludedStudentIds, studentId].filter((value, index, arr) => arr.indexOf(value) === index));
      return;
    }

    onExcludedStudentIdsChange(excludedStudentIds.filter((value) => value !== studentId));
  };

  const clearExcludedStudents = () => {
    onExcludedStudentIdsChange([]);
  };

  const toggleClassExpanded = (classGroup: string) => {
    setExpandedClassGroups((prev) => {
      const next = new Set(prev);
      if (next.has(classGroup)) {
        next.delete(classGroup);
      } else {
        next.add(classGroup);
      }
      return next;
    });
  };

  const selectPreset = (mode: BalancePresetMode) => {
    setPresetMode(mode);

    if (mode === 'even') {
      setMaxRelaxation(String(EVEN_PRESET_MAX_RELAXATION));
      setParametersExpanded(false);
      return;
    }

    if (mode === 'underMax') {
      setMaxRelaxation(String(UNDER_MAX_PRESET_MAX_RELAXATION));
      setParametersExpanded(false);
      return;
    }

    setParametersExpanded(true);
  };

  return (
    <div className={styles.wrapper}>
      {isBalancing && (
        <div className={styles.balanceOverlay} role="status" aria-live="polite" aria-busy="true">
          <div className={styles.balanceOverlayCard}>
            <h4>Balanserer data...</h4>
            <div className={styles.juggleTrack}>
              <span className={`${styles.juggleBall} ${styles.juggleBallOne}`.trim()} />
              <span className={`${styles.juggleBall} ${styles.juggleBallTwo}`.trim()} />
              <span className={`${styles.juggleBall} ${styles.juggleBallThree}`.trim()} />
            </div>
            <p>Jobber med fordeling av elever i grupper.</p>
          </div>
        </div>
      )}

      <section className={styles.card}>
        <h3>Hybrid balansering</h3>
        <p className={styles.description}>
          Blokk-kollisjoner repareres forst, for appen balanserer gruppene. Logg per elev finner du pa Logg
          etterpa, denne kan brukes for a gjore endringene i InSchool. Forst, sjekk lassebegreninger, sett opp hvilke
          blokker hvert trinn skal kunne bruke. Velg sa type balansering. Få antall under maks tar minst tid, men kan gi mer
          skjevfordelte grupper enn Balanser mest mulig jevnt. Er det fag som ikke skal balanseres, kan disse utelukkes
          fra balansering.
        </p>

        <div className={styles.constraintsBox}>
          <h4>Klassebegrensninger per blokk</h4>
          <p title={SETTING_DESCRIPTIONS.restrictions}>
            Standard: VG2 kan ikke i Blokk 4, VG3 kan ikke i Blokk 1. Kryss av hva som er tillatt.
          </p>
          <table className={styles.restrictionTable}>
            <thead>
              <tr>
                <th>Trinn</th>
                <th>Blokk 1</th>
                <th>Blokk 2</th>
                <th>Blokk 3</th>
                <th>Blokk 4</th>
              </tr>
            </thead>
            <tbody>
              {visibleClassLevels.map((classKey) => {
                return (
                  <tr key={classKey}>
                    <td>{classKey}</td>
                    {([1, 2, 3, 4] as BlockNumber[]).map((block) => {
                      const allowed = effectiveRestrictions[classKey]?.[block] ?? false;
                      return (
                        <td key={`${classKey}-${block}`}>
                          <button
                            type="button"
                            className={`${styles.restrictionToggle} ${allowed ? styles.restrictionToggleOn : styles.restrictionToggleOff}`.trim()}
                            onClick={() => updateRestriction(classKey, block, !allowed)}
                          >
                            {allowed ? 'Tillatt' : 'Ikke tillatt'}
                          </button>
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
            </tbody>
          </table>
          <button type="button" className={styles.secondaryBtn} onClick={allowAll}>
            Tillatt alle
          </button>
        </div>

        <div className={`${styles.presetRow} ${!hasAnyAllowedRestriction ? styles.disabledSection : ''}`.trim()}>
          <button
            type="button"
            className={`${styles.presetBtn} ${presetMode === 'even' ? styles.presetBtnActive : ''}`.trim()}
            onClick={() => selectPreset('even')}
            disabled={!hasAnyAllowedRestriction}
          >
            Balanser mest mulig jevnt
          </button>
          <button
            type="button"
            className={`${styles.presetBtn} ${presetMode === 'underMax' ? styles.presetBtnActive : ''}`.trim()}
            onClick={() => selectPreset('underMax')}
            disabled={!hasAnyAllowedRestriction}
          >
            Få antall under maks
          </button>
          <button
            type="button"
            className={`${styles.presetBtn} ${presetMode === 'advanced' ? styles.presetBtnActive : ''}`.trim()}
            onClick={() => selectPreset('advanced')}
            disabled={!hasAnyAllowedRestriction}
          >
            Avansert
          </button>
        </div>

        <div className={`${styles.constraintsBox} ${!hasAnyAllowedRestriction ? styles.disabledSection : ''}`.trim()}>
          <button
            type="button"
            className={styles.collapsibleHeaderBtn}
            onClick={() => setParametersExpanded((prev) => !prev)}
            aria-expanded={parametersExpanded}
            disabled={!hasAnyAllowedRestriction}
          >
            <span className={styles.chevron}>{parametersExpanded ? '▼' : '▶'}</span>
            <span>Parametere</span>
          </button>

          {parametersExpanded && (
            <fieldset className={styles.parametersFieldset} disabled={!hasAnyAllowedRestriction || presetMode !== 'advanced'}>
              {presetMode !== 'advanced' && (
                <p className={styles.parametersHint}>Velg Avansert for a redigere parameterne.</p>
              )}
              <div className={styles.weightsGrid}>
              <label>
                Overkapasitet (A)
                <input
                  type="number"
                  value={weights.overcapA}
                  step="0.1"
                  title={SETTING_DESCRIPTIONS.overcapA}
                  onChange={(event) =>
                    setWeights((prev) => ({ ...prev, overcapA: parseInputNumber(event.target.value, prev.overcapA) }))
                  }
                />
              </label>
              <label>
                Ubalanse (B)
                <input
                  type="number"
                  value={weights.imbalanceB}
                  step="0.1"
                  title={SETTING_DESCRIPTIONS.imbalanceB}
                  onChange={(event) =>
                    setWeights((prev) => ({ ...prev, imbalanceB: parseInputNumber(event.target.value, prev.imbalanceB) }))
                  }
                />
              </label>
              <label>
                Topptrykk (C)
                <input
                  type="number"
                  value={weights.peakC}
                  step="0.1"
                  title={SETTING_DESCRIPTIONS.peakC}
                  onChange={(event) =>
                    setWeights((prev) => ({ ...prev, peakC: parseInputNumber(event.target.value, prev.peakC) }))
                  }
                />
              </label>
              <label>
                Kollisjon (D)
                <input
                  type="number"
                  value={FIXED_COLLISION_WEIGHT}
                  step="1000"
                  disabled
                  title={SETTING_DESCRIPTIONS.collisionD}
                />
              </label>
              <label>
                Flyttkost (E)
                <input
                  type="number"
                  value={weights.movesE}
                  step="0.1"
                  title={SETTING_DESCRIPTIONS.movesE}
                  onChange={(event) =>
                    setWeights((prev) => ({ ...prev, movesE: parseInputNumber(event.target.value, prev.movesE) }))
                  }
                />
              </label>
              <label>
                Repeatkost (F)
                <input
                  type="number"
                  value={weights.repeatF}
                  step="0.1"
                  title={SETTING_DESCRIPTIONS.repeatF}
                  onChange={(event) =>
                    setWeights((prev) => ({ ...prev, repeatF: parseInputNumber(event.target.value, prev.repeatF) }))
                  }
                />
              </label>
              <label>
                Storgruppe-faktor (alpha)
                <input
                  type="number"
                  value={weights.alpha}
                  step="0.01"
                  title={SETTING_DESCRIPTIONS.alpha}
                  onChange={(event) =>
                    setWeights((prev) => ({ ...prev, alpha: parseInputNumber(event.target.value, prev.alpha) }))
                  }
                />
              </label>
              <label>
                Peak-faktor (beta)
                <input
                  type="number"
                  value={weights.beta}
                  step="0.1"
                  title={SETTING_DESCRIPTIONS.beta}
                  onChange={(event) =>
                    setWeights((prev) => ({ ...prev, beta: parseInputNumber(event.target.value, prev.beta) }))
                  }
                />
              </label>
              <label>
                Start under maks
                <input
                  type="number"
                  value={maxRelaxation}
                  title={SETTING_DESCRIPTIONS.maxRelaxation}
                  onChange={(event) => setMaxRelaxation(event.target.value)}
                />
              </label>
              <label>
                Maks flyt-iterasjoner per pass
                <input
                  type="number"
                  value={maxFlowIterationsPerOffset}
                  title={SETTING_DESCRIPTIONS.maxFlowIterationsPerOffset}
                  onChange={(event) => setMaxFlowIterationsPerOffset(event.target.value)}
                />
              </label>
              <label>
                Lookahead-forsok
                <input
                  type="number"
                  value={maxLookaheadAttempts}
                  title={SETTING_DESCRIPTIONS.maxLookaheadAttempts}
                  onChange={(event) => setMaxLookaheadAttempts(event.target.value)}
                />
              </label>
              <label>
                Max depth-2 kjeder
                <input
                  type="number"
                  value={maxDepth2Chains}
                  title={SETTING_DESCRIPTIONS.maxDepth2Chains}
                  onChange={(event) => setMaxDepth2Chains(event.target.value)}
                />
              </label>
              </div>
            </fieldset>
          )}
        </div>

        {!hasAnyAllowedRestriction && (
          <p className={styles.restrictionNotice}>Velg minst én blokk som Tillatt for å aktivere balanseringsinnstillinger.</p>
        )}

        {availableSubjects.length > 0 && (
          <div className={`${styles.constraintsBox} ${!hasAnyAllowedRestriction ? styles.disabledSection : ''}`.trim()}>
            <button
              type="button"
              className={styles.collapsibleHeaderBtn}
              onClick={() => setExcludedSubjectsExpanded((prev) => !prev)}
              aria-expanded={excludedSubjectsExpanded}
              disabled={!hasAnyAllowedRestriction}
            >
              <span className={styles.chevron}>{excludedSubjectsExpanded ? '▼' : '▶'}</span>
              <span>Utelukk fag fra balansering</span>
            </button>

            {excludedSubjectsExpanded && (
              <>
                <div className={styles.subjectExclusionHeader}>
                  <p title={SETTING_DESCRIPTIONS.excludedSubjects}>
                    Utelukkede fag blir ikke flyttet og teller ikke med i score, overkapasitet eller fagmetrikker.
                  </p>
                  <button type="button" className={styles.secondaryBtn} onClick={clearExcludedSubjects} disabled={!hasAnyAllowedRestriction}>
                    Nullstill fagvalg
                  </button>
                </div>

                <div className={styles.subjectExclusionList}>
                  {availableSubjects.map((subject) => {
                    const isExcluded = excludedSubjects.includes(subject);
                    return (
                      <label key={subject} className={styles.subjectToggle}>
                        <input
                          type="checkbox"
                          checked={isExcluded}
                          title={SETTING_DESCRIPTIONS.excludedSubjects}
                          onChange={(event) => toggleExcludedSubject(subject, event.target.checked)}
                          disabled={!hasAnyAllowedRestriction}
                        />
                        <span>{subject}</span>
                      </label>
                    );
                  })}
                </div>
              </>
            )}
          </div>
        )}

        {studentsByClass.length > 0 && (
          <div className={`${styles.constraintsBox} ${!hasAnyAllowedRestriction ? styles.disabledSection : ''}`.trim()}>
            <button
              type="button"
              className={styles.collapsibleHeaderBtn}
              onClick={() => setExcludedStudentsExpanded((prev) => !prev)}
              aria-expanded={excludedStudentsExpanded}
              disabled={!hasAnyAllowedRestriction}
            >
              <span className={styles.chevron}>{excludedStudentsExpanded ? '▼' : '▶'}</span>
              <span>Unnta klasser/elever fra balansering</span>
            </button>

            {excludedStudentsExpanded && (
              <>
                <div className={styles.subjectExclusionHeader}>
                  <p title={SETTING_DESCRIPTIONS.excludedStudents}>
                    Valgte klasser/elever blir laast og kan ikke flyttes under balansering.
                  </p>
                  <button type="button" className={styles.secondaryBtn} onClick={clearExcludedStudents} disabled={!hasAnyAllowedRestriction}>
                    Nullstill
                  </button>
                </div>

                <div className={styles.classExclusionList}>
                  {studentsByClass.map((entry) => {
                    const selectedCount = entry.students.filter((student) => excludedStudentIds.includes(student.studentId)).length;
                    const classSelected = selectedCount === entry.students.length && entry.students.length > 0;
                    const classPartiallySelected = selectedCount > 0 && selectedCount < entry.students.length;
                    const isExpanded = expandedClassGroups.has(entry.classGroup);

                    return (
                      <div key={entry.classGroup} className={styles.classExclusionItem}>
                        <div className={styles.classRow}>
                          <label className={styles.classToggle}>
                            <input
                              type="checkbox"
                              checked={classSelected}
                              onChange={(event) => toggleExcludedClassGroup(entry.classGroup, event.target.checked)}
                              disabled={!hasAnyAllowedRestriction}
                            />
                            <span className={styles.className}>{entry.classGroup}</span>
                            <span className={styles.classMeta}>{entry.students.length} elever</span>
                            {classPartiallySelected && (
                              <span className={styles.classMeta}>{selectedCount} valgt</span>
                            )}
                          </label>
                          <button
                            type="button"
                            className={styles.classExpandBtn}
                            onClick={() => toggleClassExpanded(entry.classGroup)}
                            disabled={!hasAnyAllowedRestriction}
                          >
                            {isExpanded ? 'Skjul elever' : 'Vis elever'}
                          </button>
                        </div>

                        {isExpanded && (
                          <div className={styles.classStudentList}>
                            {entry.students.map((student) => {
                              const checked = excludedStudentIds.includes(student.studentId);
                              return (
                                <label key={student.studentId} className={styles.studentToggle}>
                                  <input
                                    type="checkbox"
                                    checked={checked}
                                    disabled={!hasAnyAllowedRestriction}
                                    onChange={(event) => toggleExcludedStudent(student.studentId, event.target.checked)}
                                  />
                                  <span>{student.name}</span>
                                </label>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </>
            )}
          </div>
        )}

        <div className={`${styles.actionRow} ${!hasAnyAllowedRestriction ? styles.disabledSection : ''}`.trim()}>
          <button
            type="button"
            className={styles.primaryBtn}
            onClick={runBalancing}
            disabled={!hasAnyAllowedRestriction || isBalancing}
            title={!hasAnyAllowedRestriction ? 'Aktiver minst én blokk først' : undefined}
          >
            Kjør balansering
          </button>
          {statusMessage && <span className={styles.status}>{statusMessage}</span>}
        </div>

        {lastResult && (
          <div className={styles.runStatusBox}>
            <div className={styles.runStatusTitle}>Siste kjoring</div>
            <div className={styles.runStatusGrid}>
              <div>Pass: {lastResult.diagnostics.passesRun}</div>
              <div>Flytt: {lastResult.diagnostics.moveCount}</div>
              <div>Over maks før: {lastResult.diagnostics.beforeOvercapSeatCount}</div>
              <div>Over maks etter: {lastResult.diagnostics.afterOvercapSeatCount}</div>
              <div>Unike elever: {lastResult.diagnostics.uniqueStudentsMoved}</div>
              <div>Lookahead forsok: {lastResult.diagnostics.lookaheadAttempts}</div>
              <div>Lookahead suksess: {lastResult.diagnostics.lookaheadSuccess}</div>
              <div>Lookahead rollback: {lastResult.diagnostics.lookaheadRollback}</div>
              <div>Repeterte flytt: {lastResult.diagnostics.repeatedMoveCount}</div>
              <div>Uloselige kollisjoner: {lastResult.diagnostics.unresolvedCollisions.length}</div>
            </div>
          </div>
        )}
      </section>

      {lastResult && (
        <section className={styles.card}>
          <button
            type="button"
            className={styles.collapsibleHeaderBtn}
            onClick={() => setDiagnosticsExpanded((prev) => !prev)}
            aria-expanded={diagnosticsExpanded}
          >
            <span className={styles.chevron}>{diagnosticsExpanded ? '▼' : '▶'}</span>
            <span>Diagnostikk</span>
          </button>

          {diagnosticsExpanded && (
            <>
              <div className={styles.diagnosticsGrid}>
                <div>Score for: {formatNumber(lastResult.diagnostics.beforeScore.total)}</div>
                <div>Score etter: {formatNumber(lastResult.diagnostics.afterScore.total)}</div>
                <div>Over maks før: {lastResult.diagnostics.beforeOvercapSeatCount}</div>
                <div>Over maks etter: {lastResult.diagnostics.afterOvercapSeatCount}</div>
                <div>Flytt: {lastResult.diagnostics.moveCount}</div>
                <div>Unike elever: {lastResult.diagnostics.uniqueStudentsMoved}</div>
                <div>Repeterte flytt: {lastResult.diagnostics.repeatedMoveCount}</div>
                <div>Pass: {lastResult.diagnostics.passesRun}</div>
                <div>Lookahead forsok: {lastResult.diagnostics.lookaheadAttempts}</div>
                <div>Lookahead suksess: {lastResult.diagnostics.lookaheadSuccess}</div>
                <div>Lookahead rollback: {lastResult.diagnostics.lookaheadRollback}</div>
                <div>Uloselige kollisjoner: {lastResult.diagnostics.unresolvedCollisions.length}</div>
              </div>

              <div className={styles.subSection}>
                <h5>Siste flytt</h5>
                <div className={styles.movesList}>
                  {lastResult.moveRecords.length === 0 ? (
                    <div>Ingen flytt i denne kjoringen.</div>
                  ) : (
                    lastResult.moveRecords.slice(-50).reverse().map((move, index) => (
                      <div key={`${move.studentId}-${move.subjectCode}-${index}`} className={styles.moveRow}>
                        <strong>{move.studentName}</strong>
                        <span>
                          {move.subjectName}: {move.fromGroupCode}/B{move.fromBlock} {'->'} {move.toGroupCode}/B{move.toBlock}
                        </span>
                        <span>
                          {move.reason}, delta {formatNumber(move.scoreDelta)}
                        </span>
                      </div>
                    ))
                  )}
                </div>
              </div>
            </>
          )}
        </section>
      )}
    </div>
  );
};
