import { useMemo, useState } from 'react';
import type { StandardField } from '../utils/excelUtils';
import {
  DEFAULT_BALANCING_CONFIG,
  DEFAULT_CLASS_BLOCK_RESTRICTIONS,
  progressiveHybridBalance,
  type BalancingConfig,
  type BalancingWeights,
  type BlockNumber,
  type ClassBlockRestrictions,
  type ProgressiveHybridBalanceResult,
  type SubjectSettingsByNameLike,
} from '../utils/progressiveHybridBalance';
import styles from './BalanseringView.module.css';

interface BalanseringViewProps {
  mergedData: StandardField[];
  subjectSettingsByName: SubjectSettingsByNameLike;
  restrictions: ClassBlockRestrictions;
  onRestrictionsChange: (value: ClassBlockRestrictions) => void;
  onApplyResult: (result: ProgressiveHybridBalanceResult) => void;
}

const DEFAULT_CLASS_LEVELS = ['VG1', 'VG2', 'VG3'];
const FIXED_COLLISION_WEIGHT = DEFAULT_BALANCING_CONFIG.weights.collisionD;

const formatNumber = (value: number): string => {
  return Number.isFinite(value) ? value.toFixed(2) : String(value);
};

const parseInputNumber = (value: string, fallback: number): number => {
  const parsed = Number.parseFloat(value);
  return Number.isNaN(parsed) ? fallback : parsed;
};

const normalizeRestrictions = (input: ClassBlockRestrictions): ClassBlockRestrictions => {
  const next: ClassBlockRestrictions = { ...DEFAULT_CLASS_BLOCK_RESTRICTIONS };
  Object.entries(input).forEach(([classKey, map]) => {
    next[classKey] = {
      ...(next[classKey] || {}),
      ...(map || {}),
    };
  });
  return next;
};

export const BalanseringView = ({
  mergedData,
  subjectSettingsByName,
  restrictions,
  onRestrictionsChange,
  onApplyResult,
}: BalanseringViewProps) => {
  const [weights, setWeights] = useState<BalancingWeights>(DEFAULT_BALANCING_CONFIG.weights);
  const [maxRelaxation, setMaxRelaxation] = useState(String(DEFAULT_BALANCING_CONFIG.maxRelaxation));
  const [maxPassMillis, setMaxPassMillis] = useState(String(DEFAULT_BALANCING_CONFIG.maxPassMillis));
  const [maxLookaheadAttempts, setMaxLookaheadAttempts] = useState(String(DEFAULT_BALANCING_CONFIG.maxLookaheadAttempts));
  const [maxDepth2Chains, setMaxDepth2Chains] = useState(String(DEFAULT_BALANCING_CONFIG.maxDepth2Chains));
  const [statusMessage, setStatusMessage] = useState('');
  const [lastResult, setLastResult] = useState<ProgressiveHybridBalanceResult | null>(null);

  const effectiveRestrictions = useMemo(() => normalizeRestrictions(restrictions), [restrictions]);

  const runBalancing = () => {
    if (mergedData.length === 0) {
      setStatusMessage('Ingen elevdata a balansere. Last inn data forst.');
      return;
    }

    const config: Partial<BalancingConfig> = {
      weights: {
        ...weights,
        collisionD: FIXED_COLLISION_WEIGHT,
      },
      maxRelaxation: Math.max(0, Math.floor(parseInputNumber(maxRelaxation, DEFAULT_BALANCING_CONFIG.maxRelaxation))),
      maxPassMillis: Math.max(200, Math.floor(parseInputNumber(maxPassMillis, DEFAULT_BALANCING_CONFIG.maxPassMillis))),
      maxLookaheadAttempts: Math.max(
        0,
        Math.floor(parseInputNumber(maxLookaheadAttempts, DEFAULT_BALANCING_CONFIG.maxLookaheadAttempts))
      ),
      maxDepth2Chains: Math.max(0, Math.floor(parseInputNumber(maxDepth2Chains, DEFAULT_BALANCING_CONFIG.maxDepth2Chains))),
      classBlockRestrictions: effectiveRestrictions,
    };

    const result = progressiveHybridBalance(mergedData, subjectSettingsByName, config);
    setLastResult(result);
    onApplyResult(result);

    const unresolvedCollisionCount = result.diagnostics.unresolvedCollisions.length;
    setStatusMessage(
      `Kjort ferdig: ${result.diagnostics.moveCount} flytt, ${result.diagnostics.uniqueStudentsMoved} elever, score ${formatNumber(
        result.diagnostics.beforeScore.total
      )} -> ${formatNumber(result.diagnostics.afterScore.total)}${
        unresolvedCollisionCount > 0
          ? `, ADVARSEL: ${unresolvedCollisionCount} elevfag kan ikke plasseres uten kollisjon (se endringslogg)`
          : ''
      }`
    );
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

  const resetSchoolDefaults = () => {
    onRestrictionsChange({ ...DEFAULT_CLASS_BLOCK_RESTRICTIONS });
  };

  return (
    <div className={styles.wrapper}>
      <section className={styles.card}>
        <h3>Hybrid balansering</h3>
        <p className={styles.description}>
          Min-cost flyt-lignende global pass med lokal lookahead-reparasjon. Motormodulene er deterministiske med
          stabil tie-breaker.
        </p>

        <div className={styles.constraintsBox}>
          <h4>Klassebegrensninger per blokk</h4>
          <p>Standard: VG2 kan ikke i Blokk 4, VG3 kan ikke i Blokk 1. Kryss av hva som er tillatt.</p>
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
              {DEFAULT_CLASS_LEVELS.map((classKey) => {
                return (
                  <tr key={classKey}>
                    <td>{classKey}</td>
                    {([1, 2, 3, 4] as BlockNumber[]).map((block) => {
                      const allowed = effectiveRestrictions[classKey]?.[block] ?? true;
                      return (
                        <td key={`${classKey}-${block}`}>
                          <label className={styles.checkboxWrap}>
                            <input
                              type="checkbox"
                              checked={allowed}
                              onChange={(event) => updateRestriction(classKey, block, event.target.checked)}
                            />
                            <span>{allowed ? 'Tillatt' : 'Ikke tillatt'}</span>
                          </label>
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
            </tbody>
          </table>
          <button type="button" className={styles.secondaryBtn} onClick={resetSchoolDefaults}>
            Gjenopprett skolestandard
          </button>
        </div>

        <div className={styles.weightsGrid}>
          <label>
            Overkapasitet (A)
            <input
              type="number"
              value={weights.overcapA}
              step="0.1"
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
              title="Låst: kollisjon er hardt forbudt i balanseringen"
            />
          </label>
          <label>
            Flyttkost (E)
            <input
              type="number"
              value={weights.movesE}
              step="0.1"
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
              onChange={(event) =>
                setWeights((prev) => ({ ...prev, beta: parseInputNumber(event.target.value, prev.beta) }))
              }
            />
          </label>
          <label>
            Kapasitets-relaksasjon
            <input type="number" value={maxRelaxation} onChange={(event) => setMaxRelaxation(event.target.value)} />
          </label>
          <label>
            Maks tid per pass (ms)
            <input type="number" value={maxPassMillis} onChange={(event) => setMaxPassMillis(event.target.value)} />
          </label>
          <label>
            Lookahead-forsok
            <input
              type="number"
              value={maxLookaheadAttempts}
              onChange={(event) => setMaxLookaheadAttempts(event.target.value)}
            />
          </label>
          <label>
            Max depth-2 kjeder
            <input type="number" value={maxDepth2Chains} onChange={(event) => setMaxDepth2Chains(event.target.value)} />
          </label>
        </div>

        <div className={styles.actionRow}>
          <button type="button" className={styles.primaryBtn} onClick={runBalancing}>
            Kjor balansering
          </button>
          {statusMessage && <span className={styles.status}>{statusMessage}</span>}
        </div>
      </section>

      {lastResult && (
        <section className={styles.card}>
          <h4>Diagnostikk</h4>
          <div className={styles.diagnosticsGrid}>
            <div>Score for: {formatNumber(lastResult.diagnostics.beforeScore.total)}</div>
            <div>Score etter: {formatNumber(lastResult.diagnostics.afterScore.total)}</div>
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
        </section>
      )}
    </div>
  );
};
