import { useMemo, useRef, useState } from 'react';
import type { StandardField } from '../utils/excelUtils';
import { getBlokkFields } from '../utils/excelUtils';
import {
  BLOKK_LABELS,
  getActiveTotal,
  getResolvedGroupsByTarget,
  getSettingsForSubject,
  shouldShowGroup,
  type BlokkLabel,
  type ResolvedGroup,
  type SubjectGroup,
  type SubjectSettingsByName,
  type StudentIdsByBlokk,
} from '../utils/subjectGroups';
import {
  type BalancingConfig,
  type BlockNumber,
  type ClassBlockRestrictions,
  type ProgressiveHybridBalanceResult,
  type ScoreBreakdown,
  type SubjectSettingsByNameLike,
} from '../utils/progressiveHybridBalance';
import type {
  BalancingWorkerInbound,
  BalancingWorkerOutbound,
} from '../workers/progressiveHybridBalance.worker.types';
import styles from './SimulateView.module.css';

interface SimulateViewProps {
  mergedData: StandardField[];
  blokkCount: number;
  subjectSettingsByName: SubjectSettingsByName;
  restrictions: ClassBlockRestrictions;
  excludedSubjects: string[];
  excludedStudentIds: string[];
}

interface SimulationResult {
  rank: number;
  placements: Record<string, BlokkLabel[]>;
  score: ScoreBreakdown;
  moveCount: number;
  uniqueStudentsMoved: number;
}

const normalizeSubjectKey = (value: string): string => {
  return value.trim().toLocaleLowerCase('nb');
};

const parseSubjects = (value: string | null): string[] => {
  if (!value) return [];
  return value
    .split(/[,;]/)
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
};

const getStudentId = (student: StandardField, index: number): string => {
  return student.studentId || `${student.navn || 'ukjent'}:${student.klasse || 'ukjent'}:${index}`;
};

/** Infer VG level strings from a class group string. */
const inferClassLevels = (classGroup: string, isFourthYear: boolean): string[] => {
  if (isFourthYear) return ['VG4'];
  const match = classGroup.trim().toUpperCase().match(/^(\d)/);
  if (!match) {
    const normalized = classGroup.trim().toUpperCase();
    return normalized ? [normalized] : [];
  }
  const year = Number.parseInt(match[1], 10);
  if (year === 1) return ['VG1'];
  if (year === 2) return ['VG2'];
  if (year >= 3) return ['VG3'];
  return [];
};

/**
 * Check if a block is allowed for a given set of VG levels, per the restrictions.
 * If ANY VG level in the set is forbidden from the block, return false.
 */
const isBlockAllowedForLevels = (
  vgLevels: Set<string>,
  block: BlockNumber,
  restrictions: ClassBlockRestrictions
): boolean => {
  for (const level of vgLevels) {
    const rules = restrictions[level];
    if (!rules) continue;
    const allowed = rules[block];
    if (typeof allowed === 'boolean' && !allowed) return false;
  }
  return true;
};

/**
 * For a subject, compute which blocks are allowed based on the VG levels
 * of students taking that subject and the class-block restrictions.
 */
const getAllowedBlocksForSubject = (
  vgLevels: Set<string>,
  availableBlocks: BlokkLabel[],
  restrictions: ClassBlockRestrictions
): BlokkLabel[] => {
  if (vgLevels.size === 0) return availableBlocks;
  return availableBlocks.filter((label) => {
    const blockNum = Number.parseInt(label.replace('Blokk ', ''), 10) as BlockNumber;
    return isBlockAllowedForLevels(vgLevels, blockNum, restrictions);
  });
};

/**
 * Re-assign all groups of a subject to specific target blocks.
 * Returns a cloned SubjectSettingsByName with the groups moved.
 */
const applySubjectPlacement = (
  settings: SubjectSettingsByName,
  subject: string,
  targetBlocks: BlokkLabel[]
): SubjectSettingsByName => {
  const current = settings[subject];
  if (!current || !current.groups) return settings;

  const enabledGroups = current.groups.filter((g) => g.enabled);
  // Collect all groups per their current block
  const groupsByCurrentBlock = new Map<string, SubjectGroup[]>();
  enabledGroups.forEach((g) => {
    const list = groupsByCurrentBlock.get(g.blokk) || [];
    list.push(g);
    groupsByCurrentBlock.set(g.blokk, list);
  });

  // Get the unique blocks currently in use (in order)
  const currentBlocks = Array.from(new Set(enabledGroups.map((g) => g.blokk)));
  currentBlocks.sort((a, b) => a.localeCompare(b, 'nb'));

  // Map each current block to a target block
  const blockMapping = new Map<string, string>();
  currentBlocks.forEach((block, i) => {
    blockMapping.set(block, targetBlocks[i % targetBlocks.length]);
  });

  const newGroups = current.groups.map((g) => {
    if (!g.enabled) return g;
    const newBlokk = blockMapping.get(g.blokk) || g.blokk;
    return { ...g, blokk: newBlokk, sourceBlokk: newBlokk };
  });

  return {
    ...settings,
    [subject]: {
      ...current,
      groups: newGroups,
      // Clear student-to-group assignments since groups are moving
      groupStudentAssignments: {},
    },
  };
};

/**
 * Get all unique block labels where a subject currently has enabled groups.
 */
const getSubjectBlocks = (settings: SubjectSettingsByName, subject: string): BlokkLabel[] => {
  const current = settings[subject];
  if (!current || !current.groups) return [];
  const blocks = new Set<BlokkLabel>();
  current.groups.forEach((g) => {
    if (g.enabled) blocks.add(g.blokk);
  });
  return Array.from(blocks).sort((a, b) => a.localeCompare(b, 'nb'));
};

/**
 * Generate permutations of block placements for a single subject.
 * Given the subject occupies N blocks, generate all combinations of N blocks from available.
 */
const generateBlockPermutations = (
  blockCount: number,
  availableBlocks: BlokkLabel[]
): BlokkLabel[][] => {
  if (blockCount === 0 || availableBlocks.length === 0) return [[]];
  if (blockCount === 1) return availableBlocks.map((b) => [b]);

  // Generate combinations (not permutations) with order since blocks are ordered
  const results: BlokkLabel[][] = [];

  const pick = (remaining: number, start: number, current: BlokkLabel[]) => {
    if (remaining === 0) {
      results.push([...current]);
      return;
    }
    for (let i = start; i < availableBlocks.length; i++) {
      current.push(availableBlocks[i]);
      pick(remaining - 1, i, current);
      current.pop();
    }
  };

  pick(blockCount, 0, []);
  return results;
};

/**
 * Generate all placement combinations across multiple subjects.
 * Caps total combinations to avoid exponential blowup.
 */
const MAX_COMBINATIONS = 400;

const generateAllCombinations = (
  subjects: string[],
  subjectBlockCounts: Map<string, number>,
  availableBlocks: BlokkLabel[],
  allowedBlocksPerSubject?: Map<string, BlokkLabel[]>
): Array<Record<string, BlokkLabel[]>> => {
  if (subjects.length === 0) return [{}];

  // Get permutations per subject, using per-subject allowed blocks when available
  const perSubject = subjects.map((subject) => {
    const blockCount = subjectBlockCounts.get(subject) || 1;
    const subjectAvailable = allowedBlocksPerSubject?.get(subject) ?? availableBlocks;
    return {
      subject,
      permutations: generateBlockPermutations(blockCount, subjectAvailable),
    };
  });

  // Calculate total combinations
  const totalCombinations = perSubject.reduce(
    (product, entry) => product * Math.max(1, entry.permutations.length),
    1
  );

  if (totalCombinations > MAX_COMBINATIONS) {
    return []; // Signal too many - handled in UI
  }

  // Generate cartesian product
  const results: Array<Record<string, BlokkLabel[]>> = [{}];

  for (const entry of perSubject) {
    const newResults: Array<Record<string, BlokkLabel[]>> = [];
    for (const existing of results) {
      for (const perm of entry.permutations) {
        newResults.push({ ...existing, [entry.subject]: perm });
      }
    }
    results.length = 0;
    results.push(...newResults);
  }

  // Deduplicate combinations that produce the same block configuration
  const seen = new Set<string>();
  return results.filter((combo) => {
    const key = subjects.map((s) => `${s}:${(combo[s] || []).join(',')}`).join('|');
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
};

export const SimulateView = ({
  mergedData,
  blokkCount,
  subjectSettingsByName,
  restrictions,
  excludedSubjects,
  excludedStudentIds,
}: SimulateViewProps) => {
  const [selectedSubjects, setSelectedSubjects] = useState<Set<string>>(new Set());
  const [isRunning, setIsRunning] = useState(false);
  const [runningProgress, setRunningProgress] = useState('');
  const [results, setResults] = useState<SimulationResult[] | null>(null);
  const [showResults, setShowResults] = useState(false);
  const [baselineScore, setBaselineScore] = useState<ScoreBreakdown | null>(null);
  const [tooManyCombinations, setTooManyCombinations] = useState(false);
  const abortRef = useRef(false);
  const workersRef = useRef<Worker[]>([]);

  const activeBlokklabels = useMemo(
    () => BLOKK_LABELS.slice(0, Math.min(blokkCount, 4)),
    [blokkCount]
  );

  // Build subject stats from merged data
  const subjectStatsByKey = useMemo(() => {
    const stats = new Map<
      string,
      {
        subject: string;
        breakdown: Record<BlokkLabel, number>;
        idsByBlokk: StudentIdsByBlokk;
      }
    >();

    const ensureStats = (subject: string) => {
      const key = normalizeSubjectKey(subject);
      if (!stats.has(key)) {
        stats.set(key, {
          subject,
          breakdown: Object.fromEntries(activeBlokklabels.map((b) => [b, 0])),
          idsByBlokk: Object.fromEntries(activeBlokklabels.map((b) => [b, []])),
        });
      }
      return stats.get(key)!;
    };

    Object.keys(subjectSettingsByName).forEach((subject) => ensureStats(subject));

    mergedData.forEach((student, index) => {
      const studentId = getStudentId(student, index);
      getBlokkFields(blokkCount)
        .map((field, i) => ({
          label: `Blokk ${i + 1}` as BlokkLabel,
          value: (student as unknown as Record<string, string | null>)[field] ?? null,
        }))
        .forEach(({ label, value }) => {
          parseSubjects(value).forEach((subject) => {
            const entry = ensureStats(subject);
            entry.breakdown[label] = (entry.breakdown[label] || 0) + 1;
            entry.idsByBlokk[label] = [...(entry.idsByBlokk[label] || []), studentId];
          });
        });
    });

    return stats;
  }, [mergedData, subjectSettingsByName, activeBlokklabels, blokkCount]);

  // Build a map of VG levels per subject for restriction-based filtering
  const vgLevelsBySubject = useMemo(() => {
    const result = new Map<string, Set<string>>();
    mergedData.forEach((student) => {
      const klasse = student.klasse || '';
      const isFourthYear = student.fjerdearsElev === true;
      const levels = inferClassLevels(klasse, isFourthYear);
      if (levels.length === 0) return;

      getBlokkFields(blokkCount)
        .map((field) => (student as unknown as Record<string, string | null>)[field] ?? null)
        .forEach((value) => {
          parseSubjects(value).forEach((subject) => {
            const key = normalizeSubjectKey(subject);
            if (!result.has(key)) result.set(key, new Set());
            levels.forEach((l) => result.get(key)!.add(l));
          });
        });
    });
    return result;
  }, [mergedData, blokkCount]);

  // For each student, collect which subjects they take (normalized keys)
  const studentSubjectList = useMemo(() => {
    const result: Array<{ studentId: string; subjects: string[] }> = [];
    mergedData.forEach((student, index) => {
      const studentId = getStudentId(student, index);
      const subjects: string[] = [];
      getBlokkFields(blokkCount)
        .map((field) => (student as unknown as Record<string, string | null>)[field] ?? null)
        .forEach((value) => {
          parseSubjects(value).forEach((subject) => {
            subjects.push(normalizeSubjectKey(subject));
          });
        });
      if (subjects.length > 0) {
        result.push({ studentId, subjects });
      }
    });
    return result;
  }, [mergedData, blokkCount]);

  const getResolvedForSubject = (subject: string) => {
    const entry = subjectStatsByKey.get(normalizeSubjectKey(subject));
    const breakdown: Record<BlokkLabel, number> = entry?.breakdown
      ?? Object.fromEntries(activeBlokklabels.map((b) => [b, 0]));
    const studentIdsByBlokk: StudentIdsByBlokk = entry?.idsByBlokk
      ?? Object.fromEntries(activeBlokklabels.map((b) => [b, []]));

    const settings = getSettingsForSubject(subjectSettingsByName, subject, breakdown, activeBlokklabels);
    const groups = settings.groups || [];
    const groupsByTarget = getResolvedGroupsByTarget(
      groups,
      studentIdsByBlokk,
      settings.groupStudentAssignments || {},
      activeBlokklabels
    );

    return {
      settings,
      groups,
      groupsByTarget,
      activeTotal: getActiveTotal(groupsByTarget),
    };
  };

  // Build subject rows for the table
  const subjectRows = useMemo(() => {
    const allSubjectKeys = new Set<string>();
    subjectStatsByKey.forEach((_, key) => allSubjectKeys.add(key));

    return Array.from(allSubjectKeys)
      .map((key) => {
        const entry = subjectStatsByKey.get(key)!;
        const resolved = getResolvedForSubject(entry.subject);
        return {
          subject: entry.subject,
          subjectKey: key,
          ...resolved,
        };
      })
      .filter((row) => row.activeTotal > 0)
      .sort((a, b) => a.subject.localeCompare(b.subject, 'nb', { sensitivity: 'base', numeric: true }));
  }, [subjectStatsByKey, subjectSettingsByName, activeBlokklabels]);

  const toggleSubject = (subject: string) => {
    setSelectedSubjects((prev) => {
      const next = new Set(prev);
      if (next.has(subject)) {
        next.delete(subject);
      } else {
        next.add(subject);
      }
      return next;
    });
    setTooManyCombinations(false);
  };

  const selectAll = () => {
    setSelectedSubjects(new Set(subjectRows.map((r) => r.subject)));
    setTooManyCombinations(false);
  };

  const selectNone = () => {
    setSelectedSubjects(new Set());
    setTooManyCombinations(false);
  };

  // Estimate the number of combinations for the current selection
  const estimatedCombinations = useMemo(() => {
    if (selectedSubjects.size === 0) return 0;
    const selected = Array.from(selectedSubjects);
    return selected.reduce((product, subject) => {
      const key = normalizeSubjectKey(subject);
      const blocks = getSubjectBlocks(subjectSettingsByName, subject);
      const blockCount = blocks.length || 1;
      const vgLevels = vgLevelsBySubject.get(key) ?? new Set<string>();
      const allowed = getAllowedBlocksForSubject(vgLevels, activeBlokklabels, restrictions);
      // C(allowed + blockCount - 1, blockCount) — combinations with repetition
      let combos = 1;
      for (let i = 0; i < blockCount; i++) {
        combos = combos * (allowed.length + blockCount - 1 - i) / (i + 1);
      }
      return product * Math.max(1, Math.floor(combos));
    }, 1);
  }, [selectedSubjects, subjectSettingsByName, vgLevelsBySubject, activeBlokklabels, restrictions]);

  const stopSimulation = () => {
    abortRef.current = true;
    workersRef.current.forEach((w) => {
      w.postMessage({ type: 'stop' } satisfies BalancingWorkerInbound);
    });
  };

  const runSimulation = async () => {
    if (selectedSubjects.size === 0) return;

    const selected = Array.from(selectedSubjects);

    // Determine how many blocks each subject occupies
    const subjectBlockCounts = new Map<string, number>();
    selected.forEach((subject) => {
      const blocks = getSubjectBlocks(subjectSettingsByName, subject);
      subjectBlockCounts.set(subject, blocks.length || 1);
    });

    // Build per-subject allowed blocks based on VG restrictions
    const allowedBlocksPerSubject = new Map<string, BlokkLabel[]>();
    selected.forEach((subject) => {
      const key = normalizeSubjectKey(subject);
      const vgLevels = vgLevelsBySubject.get(key) ?? new Set<string>();
      allowedBlocksPerSubject.set(
        subject,
        getAllowedBlocksForSubject(vgLevels, activeBlokklabels, restrictions)
      );
    });

    // Generate all combinations (pre-filtered by VG restrictions)
    const combinations = generateAllCombinations(
      selected, subjectBlockCounts, activeBlokklabels, allowedBlocksPerSubject
    );

    if (combinations.length === 0) {
      setTooManyCombinations(true);
      return;
    }

    setTooManyCombinations(false);
    setIsRunning(true);
    setRunningProgress(`Starter simulering med ${combinations.length} kombinasjoner...`);
    setResults(null);
    setBaselineScore(null);
    abortRef.current = false;
    workersRef.current = [];

    // Helper: count unavoidable single-group collisions for a given settings snapshot.
    // A collision is unavoidable when a student has two subjects that each only exist
    // in one block, and those blocks are the same.
    const countSingleGroupCollisions = (settings: SubjectSettingsByName): number => {
      // Build: normalized subject key → set of blocks it has enabled groups in
      const subjectBlockMap = new Map<string, Set<string>>();
      for (const [name, s] of Object.entries(settings)) {
        if (!s.groups) continue;
        const blocks = new Set<string>();
        s.groups.forEach((g) => { if (g.enabled) blocks.add(g.blokk); });
        if (blocks.size > 0) subjectBlockMap.set(normalizeSubjectKey(name), blocks);
      }

      let collisions = 0;
      for (const { subjects } of studentSubjectList) {
        // Only look at subjects with exactly one block (single-group / unavoidable)
        const fixedBlockSubjects: string[] = [];
        for (const subKey of subjects) {
          const blocks = subjectBlockMap.get(subKey);
          if (blocks && blocks.size === 1) {
            fixedBlockSubjects.push(blocks.values().next().value!);
          }
        }
        if (fixedBlockSubjects.length < 2) continue;
        // Count how many share the same block
        const counts = new Map<string, number>();
        for (const b of fixedBlockSubjects) {
          counts.set(b, (counts.get(b) || 0) + 1);
        }
        for (const c of counts.values()) {
          if (c > 1) collisions += c - 1;
        }
      }
      return collisions;
    };

    const baselineCollisions = countSingleGroupCollisions(subjectSettingsByName);

    const balancingConfig: Partial<BalancingConfig> = {
      classBlockRestrictions: restrictions,
      excludedSubjects,
      lockedAssignmentKeys: excludedStudentIds.map((id) => `lock:${id}`),
      blockCount: blokkCount,
      maxLookaheadAttempts: 30,
      maxDepth2Chains: 50,
      capacityOffsets: [10, 2, 0],
    };

    // Run baseline balancing with current placement to get the reference score
    try {
      setRunningProgress('Beregner nåværende score...');
      const baseResult = await runSingleBalancing(
        mergedData,
        subjectSettingsByName as SubjectSettingsByNameLike,
        balancingConfig
      );
      setBaselineScore(baseResult.diagnostics.afterScore);
    } catch {
      // If baseline fails, continue without it
    }

    const allResults: SimulationResult[] = [];

    // Pre-filter combinations and build their modified settings
    const validCombos: Array<{ combination: Record<string, BlokkLabel[]>; settings: SubjectSettingsByName }> = [];
    for (const combination of combinations) {
      let modifiedSettings: SubjectSettingsByName = { ...subjectSettingsByName };
      selected.forEach((subject) => {
        const targetBlocks = combination[subject];
        if (targetBlocks && targetBlocks.length > 0) {
          modifiedSettings = applySubjectPlacement(modifiedSettings, subject, targetBlocks);
        }
      });

      const comboCollisions = countSingleGroupCollisions(modifiedSettings);
      if (comboCollisions > baselineCollisions) continue;

      validCombos.push({ combination, settings: modifiedSettings });
    }

    // Run combinations in parallel batches of WORKER_POOL_SIZE
    const WORKER_POOL_SIZE = 8;
    let completed = 0;

    for (let batchStart = 0; batchStart < validCombos.length; batchStart += WORKER_POOL_SIZE) {
      if (abortRef.current) break;

      const batch = validCombos.slice(batchStart, batchStart + WORKER_POOL_SIZE);
      setRunningProgress(
        `Simulerer ${Math.min(batchStart + WORKER_POOL_SIZE, validCombos.length)} / ${validCombos.length} kombinasjoner...`
      );

      const batchPromises = batch.map(async ({ combination, settings: modifiedSettings }) => {
        try {
          const result = await runSingleBalancing(
            mergedData,
            modifiedSettings as SubjectSettingsByNameLike,
            balancingConfig
          );
          return {
            rank: 0,
            placements: combination,
            score: result.diagnostics.afterScore,
            moveCount: result.diagnostics.moveCount,
            uniqueStudentsMoved: result.diagnostics.uniqueStudentsMoved,
          } as SimulationResult;
        } catch {
          return null;
        }
      });

      const batchResults = await Promise.all(batchPromises);
      for (const r of batchResults) {
        if (r) allResults.push(r);
      }
      completed += batch.length;
      setRunningProgress(
        `Fullført ${completed} / ${validCombos.length} kombinasjoner...`
      );
    }

    // Sort by total score (lower is better) and assign ranks
    allResults.sort((a, b) => a.score.total - b.score.total);
    allResults.forEach((r, i) => {
      r.rank = i + 1;
    });

    setResults(allResults);
    setIsRunning(false);
    setShowResults(true);
  };

  const runSingleBalancing = (
    rows: StandardField[],
    settings: SubjectSettingsByNameLike,
    config: Partial<BalancingConfig>
  ): Promise<ProgressiveHybridBalanceResult> => {
    return new Promise((resolve, reject) => {
      const worker = new Worker(
        new URL('../workers/progressiveHybridBalance.worker.ts', import.meta.url),
        { type: 'module' }
      );
      workersRef.current.push(worker);

      const requestId = Date.now() + Math.random();

      worker.onmessage = (event: MessageEvent<BalancingWorkerOutbound>) => {
        const msg = event.data;
        if (msg.type === 'success') {
          worker.terminate();
          resolve(msg.result);
        } else if (msg.type === 'error') {
          worker.terminate();
          reject(new Error(msg.message));
        }
        // Ignore progress messages for simulation
      };

      worker.onerror = (error) => {
        worker.terminate();
        reject(error);
      };

      worker.postMessage({
        type: 'run',
        requestId,
        payload: { rows, subjectSettingsByName: settings, config },
      } satisfies BalancingWorkerInbound);
    });
  };

  const renderGroupCell = (groupsByTarget: Record<BlokkLabel, ResolvedGroup[]>, targetBlokk: BlokkLabel) => {
    const entries = (groupsByTarget[targetBlokk] || []).filter(shouldShowGroup);
    const gridClass =
      entries.length <= 1
        ? styles.groupCardsGridOne
        : entries.length === 2 || entries.length === 4
          ? styles.groupCardsGridTwo
          : styles.groupCardsGridThree;

    return (
      <td key={targetBlokk}>
        <div className={styles.groupStack}>
          <div className={`${styles.groupCardsGrid} ${gridClass}`.trim()}>
            {entries.map((entry) => {
              const isAtMax = entry.enabled && !entry.overfilled && entry.allocatedCount === entry.max;
              return (
                <div
                  key={entry.id}
                  className={[
                    styles.groupCard,
                    entry.enabled ? styles.groupCardActive : styles.groupCardInactive,
                    isAtMax ? styles.groupCardAtMax : '',
                    entry.overfilled ? styles.groupCardOverfilled : '',
                  ]
                    .filter(Boolean)
                    .join(' ')}
                  title={`${entry.label} (${entry.allocatedCount} / ${entry.max})`}
                >
                  <span className={styles.groupCount}>{entry.allocatedCount}</span>
                </div>
              );
            })}
            {entries.length === 0 && <div className={styles.groupEmptySlot}>Tom</div>}
          </div>
        </div>
      </td>
    );
  };

  const formatScore = (value: number): string => {
    return Number.isFinite(value) ? value.toFixed(1) : String(value);
  };

  const openBestInPopup = () => {
    if (!results || results.length === 0) return;
    const best = results[0];
    const baseline = baselineScore;

    const html = `<!DOCTYPE html>
<html lang="nb">
<head>
<meta charset="utf-8">
<title>Beste simuleringsresultat</title>
<style>
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; margin: 20px; background: #f8fafe; color: #1f3040; }
  h2 { margin: 0 0 12px; font-size: 18px; }
  .baseline { background: #eef5fc; border: 1px solid #c7d7e8; border-radius: 6px; padding: 10px 14px; margin-bottom: 16px; font-size: 13px; }
  .baseline strong { color: #2a6cb8; }
  .card { background: #fff; border: 2px solid #2a6cb8; border-radius: 8px; padding: 16px; }
  .rank { font-size: 15px; font-weight: 700; color: #2a6cb8; margin-bottom: 10px; }
  .metrics { display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 14px; }
  .metric { text-align: center; min-width: 70px; }
  .metric-value { display: block; font-size: 20px; font-weight: 700; color: #173d63; }
  .metric-label { font-size: 10px; color: #57708a; text-transform: uppercase; letter-spacing: 0.05em; }
  .placements { border-top: 1px solid #e0e8f0; padding-top: 10px; }
  .placement { font-size: 13px; margin: 3px 0; }
  .placement-subject { font-weight: 600; }
</style>
</head>
<body>
<h2>Beste simuleringsresultat</h2>
${baseline ? `<div class="baseline">Nåværende: score <strong>${Number.isFinite(baseline.total) ? baseline.total.toFixed(1) : baseline.total}</strong> · kollisjoner <strong>${Number.isFinite(baseline.collision) ? baseline.collision.toFixed(1) : baseline.collision}</strong></div>` : ''}
<div class="card">
  <div class="rank">#${best.rank} Beste resultat</div>
  <div class="metrics">
    <div class="metric"><span class="metric-value">${Number.isFinite(best.score.total) ? best.score.total.toFixed(1) : best.score.total}</span><span class="metric-label">Total</span></div>
    <div class="metric"><span class="metric-value">${Number.isFinite(best.score.overcap) ? best.score.overcap.toFixed(1) : best.score.overcap}</span><span class="metric-label">Overkapasitet</span></div>
    <div class="metric"><span class="metric-value">${Number.isFinite(best.score.imbalance) ? best.score.imbalance.toFixed(1) : best.score.imbalance}</span><span class="metric-label">Ubalanse</span></div>
    <div class="metric"><span class="metric-value">${Number.isFinite(best.score.peak) ? best.score.peak.toFixed(1) : best.score.peak}</span><span class="metric-label">Topptrykk</span></div>
    <div class="metric"><span class="metric-value">${Number.isFinite(best.score.collision) ? best.score.collision.toFixed(1) : best.score.collision}</span><span class="metric-label">Kollisjoner</span></div>
    <div class="metric"><span class="metric-value">${best.moveCount}</span><span class="metric-label">Flytt</span></div>
    <div class="metric"><span class="metric-value">${best.uniqueStudentsMoved}</span><span class="metric-label">Elever flyttet</span></div>
  </div>
  <div class="placements">
    ${Object.entries(best.placements).map(([subject, blocks]) =>
      `<div class="placement"><span class="placement-subject">${subject}</span> → ${blocks.join(', ')}</div>`
    ).join('\n    ')}
  </div>
</div>
</body>
</html>`;

    const popup = window.open('', '_blank', 'width=480,height=420,resizable=yes,scrollbars=yes');
    if (popup) {
      popup.document.write(html);
      popup.document.close();
    }
  };

  if (subjectRows.length === 0) {
    return <div className={styles.empty}>Ingen fag tilgjengelig for simulering.</div>;
  }

  return (
    <div className={styles.wrapper}>
      <div className={styles.headerRow}>
        <div>
          <h3 className={styles.title}>Simulering</h3>
          <p className={styles.subtitle}>
            Velg fag og simuler ulike blokkplasseringer. Ingen faktiske endringer gjøres.
          </p>
        </div>
      </div>

      <div className={styles.toolbar}>
        <button
          type="button"
          className={styles.simulateBtn}
          disabled={selectedSubjects.size === 0 || isRunning}
          onClick={runSimulation}
        >
          Simuler
        </button>
        {results && !showResults && (
          <button
            type="button"
            className={styles.showResultsBtn}
            onClick={() => setShowResults(true)}
          >
            Vis siste resultat ({results.length})
          </button>
        )}
        {selectedSubjects.size > 0 && (
          <span className={estimatedCombinations > MAX_COMBINATIONS ? styles.comboBadgeOver : styles.comboBadge}>
            {estimatedCombinations} / {MAX_COMBINATIONS}
          </span>
        )}
        <span className={styles.selectedBadge}>
          {selectedSubjects.size} fag valgt
        </span>
        <div className={styles.selectActions}>
          <button type="button" className={styles.selectActionBtn} onClick={selectAll}>
            Velg alle
          </button>
          <button type="button" className={styles.selectActionBtn} onClick={selectNone}>
            Fjern alle
          </button>
        </div>
      </div>

      {tooManyCombinations && (
        <div className={styles.warningText}>
          For mange kombinasjoner (maks {MAX_COMBINATIONS}). Velg færre fag for simulering.
        </div>
      )}

      <table className={styles.table}>
        <colgroup>
          <col style={{ width: '30px' }} />
          <col />
          {activeBlokklabels.map((b) => (
            <col key={b} style={{ width: '80px' }} />
          ))}
          <col style={{ width: '50px' }} />
        </colgroup>
        <thead>
          <tr>
            <th></th>
            <th style={{ textAlign: 'left' }}>Fag</th>
            {activeBlokklabels.map((b) => (
              <th key={b}>{b}</th>
            ))}
            <th>Totalt</th>
          </tr>
        </thead>
        <tbody>
          {subjectRows.map((row) => {
            const isSelected = selectedSubjects.has(row.subject);
            return (
              <tr
                key={row.subjectKey}
                className={[styles.subjectRow, isSelected ? styles.subjectRowSelected : '']
                  .filter(Boolean)
                  .join(' ')}
                onClick={() => toggleSubject(row.subject)}
              >
                <td>
                  <input
                    type="checkbox"
                    className={styles.subjectCheckbox}
                    checked={isSelected}
                    onChange={() => toggleSubject(row.subject)}
                    onClick={(e) => e.stopPropagation()}
                  />
                </td>
                <td className={styles.subjectCell}>
                  <span>{row.subject}</span>
                </td>
                {activeBlokklabels.map((b) => renderGroupCell(row.groupsByTarget, b))}
                <td className={styles.totalCell}>{row.activeTotal}</td>
              </tr>
            );
          })}
        </tbody>
      </table>

      {isRunning && (
        <div className={styles.runningOverlay}>
          <div className={styles.runningCard}>
            <h4>Simulerer...</h4>
            <p className={styles.runningProgress}>{runningProgress}</p>
            <button type="button" className={styles.stopBtn} onClick={stopSimulation}>
              Stopp
            </button>
          </div>
        </div>
      )}

      {showResults && results && (
        <div className={styles.overlay} onClick={() => setShowResults(false)}>
          <div
            className={styles.resultsModal}
            role="dialog"
            aria-modal="true"
            onClick={(e) => e.stopPropagation()}
          >
            <div className={styles.resultsHeader}>
              <h3>
                Simuleringsresultater ({results.length} kombinasjoner)
              </h3>
              <div className={styles.resultsHeaderActions}>
                {results.length > 0 && (
                  <button type="button" className={styles.popupBtn} onClick={openBestInPopup}>
                    Åpne beste i eget vindu
                  </button>
                )}
                <button type="button" className={styles.closeBtn} onClick={() => setShowResults(false)}>
                  Lukk
                </button>
              </div>
            </div>
            <div className={styles.resultsBody}>
              {baselineScore && (
                <div className={styles.baselineInfo}>
                  Nåværende: score{' '}
                  <span className={styles.baselineScore}>{formatScore(baselineScore.total)}</span>
                  {' · kollisjoner '}
                  <span className={styles.baselineScore}>{formatScore(baselineScore.collision)}</span>
                </div>
              )}
              {results.length === 0 ? (
                <div className={styles.empty}>Ingen resultater tilgjengelig.</div>
              ) : (
                <div className={styles.resultsList}>
                  {results.map((result, i) => (
                    <div
                      key={i}
                      className={[styles.resultCard, i === 0 ? styles.resultCardBest : '']
                        .filter(Boolean)
                        .join(' ')}
                    >
                      <div className={styles.resultRank}>
                        #{result.rank}
                        {i === 0 && <span className={styles.resultBestLabel}>Beste resultat</span>}
                      </div>
                      <div className={styles.resultMetrics}>
                        <div className={styles.resultMetric}>
                          <span className={styles.resultMetricValue}>{formatScore(result.score.total)}</span>
                          <span className={styles.resultMetricLabel}>Total</span>
                        </div>
                        <div className={styles.resultMetric}>
                          <span className={styles.resultMetricValue}>{formatScore(result.score.overcap)}</span>
                          <span className={styles.resultMetricLabel}>Overkapasitet</span>
                        </div>
                        <div className={styles.resultMetric}>
                          <span className={styles.resultMetricValue}>{formatScore(result.score.imbalance)}</span>
                          <span className={styles.resultMetricLabel}>Ubalanse</span>
                        </div>
                        <div className={styles.resultMetric}>
                          <span className={styles.resultMetricValue}>{formatScore(result.score.peak)}</span>
                          <span className={styles.resultMetricLabel}>Topptrykk</span>
                        </div>
                        <div className={styles.resultMetric}>
                          <span className={styles.resultMetricValue}>{formatScore(result.score.collision)}</span>
                          <span className={styles.resultMetricLabel}>Kollisjoner</span>
                        </div>
                        <div className={styles.resultMetric}>
                          <span className={styles.resultMetricValue}>{result.moveCount}</span>
                          <span className={styles.resultMetricLabel}>Flytt</span>
                        </div>
                        <div className={styles.resultMetric}>
                          <span className={styles.resultMetricValue}>{result.uniqueStudentsMoved}</span>
                          <span className={styles.resultMetricLabel}>Elever flyttet</span>
                        </div>
                      </div>
                      <div className={styles.resultPlacements}>
                        {Object.entries(result.placements).map(([subject, blocks]) => (
                          <div key={subject} className={styles.resultPlacementItem}>
                            <span className={styles.resultPlacementSubject}>{subject}</span>
                            {' → '}
                            {blocks.join(', ')}
                          </div>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
