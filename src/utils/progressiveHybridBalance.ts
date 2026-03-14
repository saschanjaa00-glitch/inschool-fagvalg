import type { StandardField } from './excelUtils';
import { mapSubjectToCode } from './subjectCodeMapping';

export type BlockNumber = 1 | 2 | 3 | 4;

export interface SubjectGroupLike {
  id: string;
  blokk: `Blokk ${BlockNumber}`;
  sourceBlokk: `Blokk ${BlockNumber}`;
  enabled: boolean;
  max: number;
  createdAt: string;
}

export interface SubjectSettingsLike {
  defaultMax: number;
  groups?: SubjectGroupLike[];
  groupStudentAssignments?: Record<string, string>;
}

export type SubjectSettingsByNameLike = Record<string, SubjectSettingsLike>;

export interface AssignmentNode {
  subjectCode: string;
  subjectName: string;
  groupCode: string;
  groupId: string;
  block: BlockNumber;
  locked: boolean;
}

export interface StudentNode {
  id: string;
  fullName: string;
  classGroup: string;
  assignments: AssignmentNode[];
}

export interface GroupNode {
  key: string;
  subjectCode: string;
  subjectName: string;
  groupCode: string;
  groupId: string;
  block: BlockNumber;
  capacity: number;
  size: number;
  studentIds: Set<string>;
}

export interface MoveRecord {
  studentId: string;
  studentName: string;
  subjectCode: string;
  subjectName: string;
  fromGroupCode: string;
  fromBlock: BlockNumber;
  toGroupCode: string;
  toBlock: BlockNumber;
  reason: 'overcap' | 'equalize' | 'peak-reduction' | 'collision-fix' | 'lookahead-chain';
  scoreDelta: number;
  chainStep?: number;
}

export interface ScoreBreakdown {
  total: number;
  overcap: number;
  imbalance: number;
  peak: number;
  collision: number;
  moves: number;
  repeat: number;
}

export interface SubjectMetrics {
  subjectName: string;
  maxGroup: number;
  imbalance: number;
}

export interface BalanceDiagnostics {
  beforeScore: ScoreBreakdown;
  afterScore: ScoreBreakdown;
  subjectMetricsBefore: SubjectMetrics[];
  subjectMetricsAfter: SubjectMetrics[];
  moveCount: number;
  uniqueStudentsMoved: number;
  repeatedMoveCount: number;
  lookaheadAttempts: number;
  lookaheadSuccess: number;
  lookaheadRollback: number;
  unresolvedCollisions: string[];
  passesRun: number;
}

export interface ClassBlockRestrictions {
  // Key examples: VG1, VG2, VG3
  [classKey: string]: Partial<Record<BlockNumber, boolean>>;
}

export interface BalancingWeights {
  overcapA: number;
  imbalanceB: number;
  peakC: number;
  collisionD: number;
  movesE: number;
  repeatF: number;
  alpha: number;
  beta: number;
}

export interface BalancingConfig {
  weights: BalancingWeights;
  epsilon: number;
  maxRelaxation: number;
  maxPassMillis: number;
  maxLookaheadAttempts: number;
  maxDepth2Chains: number;
  classBlockRestrictions: ClassBlockRestrictions;
  lockedAssignmentKeys: string[];
}

interface InternalState {
  students: Map<string, StudentNode>;
  groups: Map<string, GroupNode>;
  groupsBySubject: Map<string, GroupNode[]>;
  originalRows: StandardField[];
  subjectSettingsByName: SubjectSettingsByNameLike;
  studentMoveCounts: Map<string, number>;
  history: MoveRecord[];
}

interface CandidateMove {
  studentId: string;
  subjectCode: string;
  subjectName: string;
  fromGroupId: string;
  fromGroupCode: string;
  fromBlock: BlockNumber;
  toGroupId: string;
  toGroupCode: string;
  toBlock: BlockNumber;
  estimatedScoreDelta: number;
  priorityTag: 'overcap' | 'equalize' | 'peak-reduction';
}

interface FlowNetwork {
  candidates: CandidateMove[];
  offset: number;
}

interface SolvedFlow {
  selected: CandidateMove[];
}

interface ApplyMovesResult {
  applied: MoveRecord[];
  skipped: CandidateMove[];
}

export interface ProgressiveHybridBalanceResult {
  moveRecords: MoveRecord[];
  updatedData: StandardField[];
  updatedSubjectSettingsByName: SubjectSettingsByNameLike;
  diagnostics: BalanceDiagnostics;
}

const BIG_COLLISION_PENALTY = 1_000_000;

const DEFAULT_WEIGHTS: BalancingWeights = {
  overcapA: 8,
  imbalanceB: 1,
  peakC: 1.8,
  collisionD: BIG_COLLISION_PENALTY,
  movesE: 0.7,
  repeatF: 1.6,
  alpha: 0.06,
  beta: 1,
};

export const DEFAULT_CLASS_BLOCK_RESTRICTIONS: ClassBlockRestrictions = {
  VG2: {
    4: false,
  },
  VG3: {
    1: false,
  },
};

export const DEFAULT_BALANCING_CONFIG: BalancingConfig = {
  weights: DEFAULT_WEIGHTS,
  epsilon: 0.0001,
  maxRelaxation: 2,
  maxPassMillis: 1400,
  maxLookaheadAttempts: 180,
  maxDepth2Chains: 50,
  classBlockRestrictions: DEFAULT_CLASS_BLOCK_RESTRICTIONS,
  lockedAssignmentKeys: [],
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

type BlokkKey = `blokk${BlockNumber}`;

const getBlokkKey = (block: BlockNumber): BlokkKey => {
  return `blokk${block}` as BlokkKey;
};

const normalizeSubject = (value: string): string => value.trim().toLocaleLowerCase('nb');

const inferStudentId = (row: StandardField, index: number): string => {
  return row.studentId || `${row.navn || 'ukjent'}:${row.klasse || 'ukjent'}:${index}`;
};

const inferClassLevel = (classGroup: string): string => {
  const match = classGroup.trim().toUpperCase().match(/^(\d)/);
  if (!match) {
    return classGroup.trim().toUpperCase();
  }

  const year = Number.parseInt(match[1], 10);
  if (year === 1) {
    return 'VG1';
  }
  if (year === 2) {
    return 'VG2';
  }
  if (year >= 3) {
    return 'VG3';
  }
  return classGroup.trim().toUpperCase();
};

const blockFromLabel = (label: `Blokk ${BlockNumber}`): BlockNumber => {
  return Number.parseInt(label.replace('Blokk ', ''), 10) as BlockNumber;
};

const getEffectiveCapacity = (capacity: number, offset: number): number => {
  return Math.max(0, capacity - offset);
};

const cloneState = (state: InternalState): InternalState => {
  const students = new Map<string, StudentNode>();
  state.students.forEach((student, id) => {
    students.set(id, {
      ...student,
      assignments: student.assignments.map((assignment) => ({ ...assignment })),
    });
  });

  const groups = new Map<string, GroupNode>();
  state.groups.forEach((group, key) => {
    groups.set(key, {
      ...group,
      studentIds: new Set(group.studentIds),
    });
  });

  const groupsBySubject = new Map<string, GroupNode[]>();
  state.groupsBySubject.forEach((subjectGroups, subjectCode) => {
    groupsBySubject.set(
      subjectCode,
      subjectGroups
        .map((group) => groups.get(group.key))
        .filter((group): group is GroupNode => !!group)
    );
  });

  return {
    students,
    groups,
    groupsBySubject,
    originalRows: state.originalRows,
    subjectSettingsByName: state.subjectSettingsByName,
    studentMoveCounts: new Map(state.studentMoveCounts),
    history: [...state.history],
  };
};

const isClassAllowedInBlock = (
  classGroup: string,
  block: BlockNumber,
  restrictions: ClassBlockRestrictions
): boolean => {
  const classLevel = inferClassLevel(classGroup);
  const levelRules = restrictions[classLevel];
  if (!levelRules) {
    return true;
  }

  const allowed = levelRules[block];
  if (typeof allowed === 'boolean') {
    return allowed;
  }

  return true;
};

const assignmentLockKey = (studentId: string, subjectCode: string): string => `${studentId}|${subjectCode}`;

const compareCandidate = (left: CandidateMove, right: CandidateMove): number => {
  if (left.estimatedScoreDelta !== right.estimatedScoreDelta) {
    return left.estimatedScoreDelta - right.estimatedScoreDelta;
  }

  const overloadLeft = Math.abs(left.fromBlock - left.toBlock);
  const overloadRight = Math.abs(right.fromBlock - right.toBlock);
  if (overloadLeft !== overloadRight) {
    return overloadRight - overloadLeft;
  }

  if (left.studentId !== right.studentId) {
    return left.studentId.localeCompare(right.studentId, 'nb', { sensitivity: 'base' });
  }

  if (left.subjectCode !== right.subjectCode) {
    return left.subjectCode.localeCompare(right.subjectCode, 'nb', { sensitivity: 'base' });
  }

  if (left.toBlock !== right.toBlock) {
    return left.toBlock - right.toBlock;
  }

  return left.toGroupCode.localeCompare(right.toGroupCode, 'nb', { sensitivity: 'base' });
};

const sortGroupsStable = (groups: SubjectGroupLike[]): SubjectGroupLike[] => {
  return [...groups].sort((left, right) => {
    const leftBlock = blockFromLabel(left.blokk);
    const rightBlock = blockFromLabel(right.blokk);
    if (leftBlock !== rightBlock) {
      return leftBlock - rightBlock;
    }

    if (left.createdAt !== right.createdAt) {
      return left.createdAt.localeCompare(right.createdAt);
    }

    return left.id.localeCompare(right.id, 'nb', { sensitivity: 'base' });
  });
};

const groupCodeByIndex = (block: BlockNumber, index: number): string => `${block}-${index + 1}`;

const ensureSubjectGroups = (
  subjectName: string,
  subjectSettingsByName: SubjectSettingsByNameLike,
  occupancyByBlock: Record<BlockNumber, number>
): SubjectGroupLike[] => {
  const raw = subjectSettingsByName[subjectName];
  const defaultMax = raw?.defaultMax ?? 30;

  if (raw?.groups && raw.groups.length > 0) {
    return sortGroupsStable(raw.groups);
  }

  const autoGroups: SubjectGroupLike[] = [];
  ([1, 2, 3, 4] as BlockNumber[]).forEach((block) => {
    if (occupancyByBlock[block] > 0) {
      autoGroups.push({
        id: `auto-${mapSubjectToCode(subjectName)}-${block}-1`,
        blokk: `Blokk ${block}`,
        sourceBlokk: `Blokk ${block}`,
        enabled: true,
        max: defaultMax,
        createdAt: '1970-01-01T00:00:00.000Z',
      });
    }
  });

  return sortGroupsStable(autoGroups);
};

export const buildState = (
  rows: StandardField[],
  subjectSettingsByName: SubjectSettingsByNameLike,
  config?: Partial<BalancingConfig>
): InternalState => {
  const mergedConfig: BalancingConfig = {
    ...DEFAULT_BALANCING_CONFIG,
    ...config,
    weights: {
      ...DEFAULT_BALANCING_CONFIG.weights,
      ...(config?.weights || {}),
    },
    classBlockRestrictions: {
      ...DEFAULT_BALANCING_CONFIG.classBlockRestrictions,
      ...(config?.classBlockRestrictions || {}),
    },
    lockedAssignmentKeys: config?.lockedAssignmentKeys || DEFAULT_BALANCING_CONFIG.lockedAssignmentKeys,
  };

  const students = new Map<string, StudentNode>();
  const occupancyBySubject = new Map<string, Record<BlockNumber, number>>();

  rows.forEach((row, index) => {
    const studentId = inferStudentId(row, index);
    const fullName = row.navn || 'Ukjent';
    const classGroup = row.klasse || '';

    const assignments: AssignmentNode[] = [];

    ([1, 2, 3, 4] as BlockNumber[]).forEach((block) => {
      const subjectsInBlock = parseSubjects(row[getBlokkKey(block)] as string | null);

      subjectsInBlock.forEach((subjectName) => {
        const normalized = normalizeSubject(subjectName);
        const existing = occupancyBySubject.get(normalized) || { 1: 0, 2: 0, 3: 0, 4: 0 };
        existing[block] = (existing[block] || 0) + 1;
        occupancyBySubject.set(normalized, existing);

        assignments.push({
          subjectCode: mapSubjectToCode(subjectName),
          subjectName,
          groupCode: '',
          groupId: '',
          block,
          locked: mergedConfig.lockedAssignmentKeys.includes(assignmentLockKey(studentId, mapSubjectToCode(subjectName))),
        });
      });
    });

    students.set(studentId, {
      id: studentId,
      fullName,
      classGroup,
      assignments,
    });
  });

  const groups = new Map<string, GroupNode>();
  const groupsBySubject = new Map<string, GroupNode[]>();

  const subjectNameLookup = new Map<string, string>();
  students.forEach((student) => {
    student.assignments.forEach((assignment) => {
      if (!subjectNameLookup.has(assignment.subjectCode)) {
        subjectNameLookup.set(assignment.subjectCode, assignment.subjectName);
      }
    });
  });

  subjectNameLookup.forEach((subjectName, subjectCode) => {
    const occupancy = occupancyBySubject.get(normalizeSubject(subjectName)) || { 1: 0, 2: 0, 3: 0, 4: 0 };
    const sourceGroups = ensureSubjectGroups(subjectName, subjectSettingsByName, occupancy);

    const indexedByBlock = new Map<BlockNumber, SubjectGroupLike[]>();
    ([1, 2, 3, 4] as BlockNumber[]).forEach((block) => indexedByBlock.set(block, []));

    sourceGroups.forEach((group) => {
      if (!group.enabled) {
        return;
      }
      const block = blockFromLabel(group.blokk);
      indexedByBlock.get(block)?.push(group);
    });

    const builtGroups: GroupNode[] = [];

    ([1, 2, 3, 4] as BlockNumber[]).forEach((block) => {
      const blockGroups = sortGroupsStable(indexedByBlock.get(block) || []);
      blockGroups.forEach((group, index) => {
        const groupCode = groupCodeByIndex(block, index);
        const key = `${subjectCode}|${groupCode}|${block}`;
        const node: GroupNode = {
          key,
          subjectCode,
          subjectName,
          groupCode,
          groupId: group.id,
          block,
          capacity: Math.max(0, Math.floor(group.max)),
          size: 0,
          studentIds: new Set<string>(),
        };

        groups.set(key, node);
        builtGroups.push(node);
      });
    });

    groupsBySubject.set(subjectCode, builtGroups);
  });

  students.forEach((student) => {
    student.assignments.forEach((assignment) => {
      const subjectGroups = groupsBySubject.get(assignment.subjectCode) || [];
      const targetGroupsInBlock = subjectGroups.filter((group) => group.block === assignment.block);

      if (targetGroupsInBlock.length === 0) {
        const fallbackKey = `${assignment.subjectCode}|${assignment.block}-1|${assignment.block}`;
        const fallbackNode: GroupNode = {
          key: fallbackKey,
          subjectCode: assignment.subjectCode,
          subjectName: assignment.subjectName,
          groupCode: `${assignment.block}-1`,
          groupId: `auto-${assignment.subjectCode}-${assignment.block}-fallback`,
          block: assignment.block,
          capacity: 30,
          size: 0,
          studentIds: new Set<string>(),
        };

        groups.set(fallbackKey, fallbackNode);
        const existing = groupsBySubject.get(assignment.subjectCode) || [];
        groupsBySubject.set(assignment.subjectCode, [...existing, fallbackNode]);
      }

      const reloadedGroups = (groupsBySubject.get(assignment.subjectCode) || []).filter(
        (group) => group.block === assignment.block
      );

      const orderedGroups = [...reloadedGroups].sort((left, right) => {
        if (left.groupCode !== right.groupCode) {
          return left.groupCode.localeCompare(right.groupCode, 'nb', { sensitivity: 'base' });
        }
        return left.groupId.localeCompare(right.groupId, 'nb', { sensitivity: 'base' });
      });

      const perSubjectSettings = subjectSettingsByName[assignment.subjectName];
      const explicitGroupId = perSubjectSettings?.groupStudentAssignments?.[student.id];
      const explicitGroup = orderedGroups.find((group) => group.groupId === explicitGroupId);
      const fallbackIndex = Math.abs(
        [...student.id].reduce((sum, ch) => sum + ch.charCodeAt(0), 0)
      ) % orderedGroups.length;
      const targetGroup = explicitGroup || orderedGroups[fallbackIndex];

      assignment.groupCode = targetGroup.groupCode;
      assignment.groupId = targetGroup.groupId;

      targetGroup.size += 1;
      targetGroup.studentIds.add(student.id);
    });
  });

  return {
    students,
    groups,
    groupsBySubject,
    originalRows: rows,
    subjectSettingsByName,
    studentMoveCounts: new Map<string, number>(),
    history: [],
  };
};

const getCollisionCount = (student: StudentNode): number => {
  const byBlock = new Map<BlockNumber, number>();
  student.assignments.forEach((assignment) => {
    byBlock.set(assignment.block, (byBlock.get(assignment.block) || 0) + 1);
  });

  let collisions = 0;
  byBlock.forEach((count) => {
    if (count > 1) {
      collisions += count - 1;
    }
  });

  return collisions;
};

const getSubjectSizes = (state: InternalState): Map<string, number[]> => {
  const result = new Map<string, number[]>();
  state.groupsBySubject.forEach((groups, subjectCode) => {
    const sizes = groups
      .map((group) => group.size)
      .sort((left, right) => left - right);
    result.set(subjectCode, sizes);
  });
  return result;
};

const computeSubjectImbalanceRaw = (sizes: number[]): number => {
  let penalty = 0;
  for (let i = 0; i < sizes.length; i += 1) {
    for (let j = i + 1; j < sizes.length; j += 1) {
      const diff = sizes[i] - sizes[j];
      penalty += diff * diff;
    }
  }
  return penalty;
};

export const computeScore = (
  state: InternalState,
  weights: BalancingWeights,
  moveRecords: MoveRecord[]
): ScoreBreakdown => {
  let overcapRaw = 0;
  state.groups.forEach((group) => {
    const excess = Math.max(0, group.size - group.capacity);
    overcapRaw += excess * excess;
  });

  const sizesBySubject = getSubjectSizes(state);
  let imbalanceRaw = 0;
  let peakRaw = 0;

  sizesBySubject.forEach((sizes) => {
    if (sizes.length === 0) {
      return;
    }

    const peak = Math.max(...sizes);
    const wSubject = 1 + weights.alpha * peak;
    imbalanceRaw += wSubject * computeSubjectImbalanceRaw(sizes);
    peakRaw += weights.beta * peak * peak;
  });

  const collisionRaw = Array.from(state.students.values())
    .reduce((sum, student) => sum + getCollisionCount(student), 0);

  const movesRaw = moveRecords.length;
  const repeatedRaw = Array.from(state.studentMoveCounts.values())
    .reduce((sum, value) => sum + Math.max(0, value - 1), 0);

  const overcap = weights.overcapA * overcapRaw;
  const imbalance = weights.imbalanceB * imbalanceRaw;
  const peak = weights.peakC * peakRaw;
  const collision = weights.collisionD * collisionRaw;
  const moves = weights.movesE * movesRaw;
  const repeat = weights.repeatF * repeatedRaw;

  return {
    total: overcap + imbalance + peak + collision + moves + repeat,
    overcap,
    imbalance,
    peak,
    collision,
    moves,
    repeat,
  };
};

const moveIsFeasible = (
  state: InternalState,
  move: CandidateMove,
  offset: number,
  restrictions: ClassBlockRestrictions
): boolean => {
  const student = state.students.get(move.studentId);
  if (!student) {
    return false;
  }

  const assignment = student.assignments.find(
    (item) => item.subjectCode === move.subjectCode && item.block === move.fromBlock
  );

  if (!assignment || assignment.locked) {
    return false;
  }

  if (!isClassAllowedInBlock(student.classGroup, move.toBlock, restrictions)) {
    return false;
  }

  const hasOtherSubjectInTarget = student.assignments.some((item) => {
    return item.block === move.toBlock && item.subjectCode !== move.subjectCode;
  });

  if (hasOtherSubjectInTarget) {
    return false;
  }

  const targetGroup = Array.from(state.groups.values()).find((group) => group.groupId === move.toGroupId);
  if (!targetGroup) {
    return false;
  }

  // Offset<0 allows temporary extra headroom in relaxed passes. Offset=0 is strict.
  const effectiveCapacity = getEffectiveCapacity(targetGroup.capacity, offset);
  if (targetGroup.size + 1 > effectiveCapacity) {
    return false;
  }

  return true;
};

const applyCandidateMoveToState = (
  state: InternalState,
  move: CandidateMove,
  reason: MoveRecord['reason'],
  scoreDelta: number,
  offset: number,
  restrictions: ClassBlockRestrictions,
  chainStep?: number
): boolean => {
  if (!moveIsFeasible(state, move, offset, restrictions)) {
    return false;
  }

  const student = state.students.get(move.studentId);
  if (!student) {
    return false;
  }

  const assignment = student.assignments.find((item) => {
    return item.subjectCode === move.subjectCode && item.block === move.fromBlock;
  });

  if (!assignment || assignment.locked) {
    return false;
  }

  const fromGroup = Array.from(state.groups.values()).find((group) => group.groupId === move.fromGroupId);
  const toGroup = Array.from(state.groups.values()).find((group) => group.groupId === move.toGroupId);

  if (!fromGroup || !toGroup) {
    return false;
  }

  if (!fromGroup.studentIds.has(student.id)) {
    return false;
  }

  fromGroup.studentIds.delete(student.id);
  fromGroup.size = Math.max(0, fromGroup.size - 1);
  toGroup.studentIds.add(student.id);
  toGroup.size += 1;

  assignment.block = move.toBlock;
  assignment.groupCode = move.toGroupCode;
  assignment.groupId = move.toGroupId;

  state.studentMoveCounts.set(student.id, (state.studentMoveCounts.get(student.id) || 0) + 1);

  state.history.push({
    studentId: student.id,
    studentName: student.fullName,
    subjectCode: move.subjectCode,
    subjectName: move.subjectName,
    fromGroupCode: move.fromGroupCode,
    fromBlock: move.fromBlock,
    toGroupCode: move.toGroupCode,
    toBlock: move.toBlock,
    reason,
    scoreDelta,
    chainStep,
  });

  return true;
};

const estimateMoveDelta = (
  state: InternalState,
  move: CandidateMove,
  config: BalancingConfig,
  reason: MoveRecord['reason'],
  offset: number
): number => {
  if (!moveIsFeasible(state, move, offset, config.classBlockRestrictions)) {
    return Number.POSITIVE_INFINITY;
  }

  const before = computeScore(state, config.weights, state.history).total;
  const cloned = cloneState(state);
  const applied = applyCandidateMoveToState(cloned, move, reason, 0, offset, config.classBlockRestrictions);
  if (!applied) {
    return Number.POSITIVE_INFINITY;
  }

  const after = computeScore(cloned, config.weights, cloned.history).total;
  return after - before;
};

export const buildFlowNetwork = (
  state: InternalState,
  config: BalancingConfig,
  offset: number
): FlowNetwork => {
  const candidates: CandidateMove[] = [];

  state.students.forEach((student) => {
    student.assignments.forEach((assignment) => {
      if (assignment.locked) {
        return;
      }

      const subjectGroups = state.groupsBySubject.get(assignment.subjectCode) || [];
      const sourceGroup = subjectGroups.find((group) => group.groupId === assignment.groupId);
      if (!sourceGroup) {
        return;
      }

      subjectGroups.forEach((targetGroup) => {
        if (targetGroup.groupId === sourceGroup.groupId) {
          return;
        }

        const baseMove: CandidateMove = {
          studentId: student.id,
          subjectCode: assignment.subjectCode,
          subjectName: assignment.subjectName,
          fromGroupId: sourceGroup.groupId,
          fromGroupCode: sourceGroup.groupCode,
          fromBlock: sourceGroup.block,
          toGroupId: targetGroup.groupId,
          toGroupCode: targetGroup.groupCode,
          toBlock: targetGroup.block,
          estimatedScoreDelta: 0,
          priorityTag: 'equalize',
        };

        const sourceExcess = Math.max(0, sourceGroup.size - getEffectiveCapacity(sourceGroup.capacity, offset));
        const targetSlack = getEffectiveCapacity(targetGroup.capacity, offset) - targetGroup.size;

        let priorityTag: CandidateMove['priorityTag'] = 'equalize';
        if (sourceExcess > 0 && targetSlack > 0) {
          priorityTag = 'overcap';
        } else {
          const subjectGroupsSorted = [...subjectGroups].sort((left, right) => right.size - left.size);
          const isFromPeak = subjectGroupsSorted[0]?.groupId === sourceGroup.groupId;
          if (isFromPeak && targetGroup.size < sourceGroup.size) {
            priorityTag = 'peak-reduction';
          }
        }

        baseMove.priorityTag = priorityTag;
        baseMove.estimatedScoreDelta = estimateMoveDelta(state, baseMove, config, priorityTag, offset);

        if (Number.isFinite(baseMove.estimatedScoreDelta)) {
          candidates.push(baseMove);
        }
      });
    });
  });

  candidates.sort(compareCandidate);

  return {
    candidates,
    offset,
  };
};

export const solveFlow = (network: FlowNetwork): SolvedFlow => {
  // Deterministic successive selection works as a min-cost flow approximation for this UI budget.
  const selected: CandidateMove[] = [];
  const usedAssignmentKeys = new Set<string>();

  network.candidates.forEach((candidate) => {
    if (candidate.estimatedScoreDelta >= 0) {
      return;
    }

    const assignmentKey = `${candidate.studentId}|${candidate.subjectCode}`;
    if (usedAssignmentKeys.has(assignmentKey)) {
      return;
    }

    usedAssignmentKeys.add(assignmentKey);
    selected.push(candidate);
  });

  return { selected };
};

export const extractMovesFromFlow = (solved: SolvedFlow): CandidateMove[] => {
  return [...solved.selected].sort(compareCandidate);
};

export const applyMoves = (
  state: InternalState,
  moves: CandidateMove[],
  config: BalancingConfig,
  offset: number
): ApplyMovesResult => {
  const applied: MoveRecord[] = [];
  const skipped: CandidateMove[] = [];

  moves.forEach((move) => {
    if (!moveIsFeasible(state, move, offset, config.classBlockRestrictions)) {
      skipped.push(move);
      return;
    }

    const scoreDelta = estimateMoveDelta(state, move, config, move.priorityTag, offset);
    if (!Number.isFinite(scoreDelta) || scoreDelta >= -config.epsilon) {
      skipped.push(move);
      return;
    }

    const success = applyCandidateMoveToState(
      state,
      move,
      move.priorityTag,
      scoreDelta,
      offset,
      config.classBlockRestrictions
    );
    if (!success) {
      skipped.push(move);
      return;
    }

    applied.push({
      studentId: move.studentId,
      studentName: state.students.get(move.studentId)?.fullName || 'Ukjent',
      subjectCode: move.subjectCode,
      subjectName: move.subjectName,
      fromGroupCode: move.fromGroupCode,
      fromBlock: move.fromBlock,
      toGroupCode: move.toGroupCode,
      toBlock: move.toBlock,
      reason: move.priorityTag,
      scoreDelta,
    });
  });

  return { applied, skipped };
};

const generateAllCandidateMoves = (
  state: InternalState,
  config: BalancingConfig,
  offset: number
): CandidateMove[] => {
  return buildFlowNetwork(state, config, offset).candidates;
};

const trySingleMoveImprove = (state: InternalState, config: BalancingConfig, offset: number): boolean => {
  const candidates = generateAllCandidateMoves(state, config, offset)
    .filter((candidate) => candidate.estimatedScoreDelta < -config.epsilon);
  if (candidates.length === 0) {
    return false;
  }

  const best = candidates[0];
  return applyCandidateMoveToState(
    state,
    best,
    best.priorityTag,
    best.estimatedScoreDelta,
    offset,
    config.classBlockRestrictions
  );
};

export const tryLookaheadChain = (
  state: InternalState,
  config: BalancingConfig,
  depth: 1 | 2,
  maxAttempts: number,
  offset: number
): { success: boolean; attempts: number; rollbacks: number } => {
  let attempts = 0;
  let rollbacks = 0;

  const firstLevel = generateAllCandidateMoves(state, config, offset);

  for (const first of firstLevel) {
    if (attempts >= maxAttempts) {
      break;
    }

    attempts += 1;
    const baseline = computeScore(state, config.weights, state.history).total;
    const snapshot = cloneState(state);

    if (!moveIsFeasible(snapshot, first, offset, config.classBlockRestrictions)) {
      continue;
    }

    const firstDelta = estimateMoveDelta(snapshot, first, config, 'lookahead-chain', offset);
    if (!Number.isFinite(firstDelta)) {
      continue;
    }

    const firstApplied = applyCandidateMoveToState(
      snapshot,
      first,
      'lookahead-chain',
      firstDelta,
      offset,
      config.classBlockRestrictions,
      1
    );
    if (!firstApplied) {
      continue;
    }

    let success = false;

    const secondLevel = generateAllCandidateMoves(snapshot, config, offset)
      .filter((candidate) => `${candidate.studentId}|${candidate.subjectCode}` !== `${first.studentId}|${first.subjectCode}`);

    if (depth === 1) {
      const after = computeScore(snapshot, config.weights, snapshot.history).total;
      if (after <= baseline - config.epsilon) {
        state.students = snapshot.students;
        state.groups = snapshot.groups;
        state.groupsBySubject = snapshot.groupsBySubject;
        state.studentMoveCounts = snapshot.studentMoveCounts;
        state.history = snapshot.history;
        success = true;
      }
    } else {
      for (const second of secondLevel.slice(0, config.maxDepth2Chains)) {
        const secondDelta = estimateMoveDelta(snapshot, second, config, 'lookahead-chain', offset);
        if (!Number.isFinite(secondDelta)) {
          continue;
        }

        const secondApplied = applyCandidateMoveToState(
          snapshot,
          second,
          'lookahead-chain',
          secondDelta,
          offset,
          config.classBlockRestrictions,
          2
        );
        if (!secondApplied) {
          continue;
        }

        const after = computeScore(snapshot, config.weights, snapshot.history).total;
        if (after <= baseline - config.epsilon) {
          state.students = snapshot.students;
          state.groups = snapshot.groups;
          state.groupsBySubject = snapshot.groupsBySubject;
          state.studentMoveCounts = snapshot.studentMoveCounts;
          state.history = snapshot.history;
          success = true;
          break;
        }
      }
    }

    if (success) {
      return { success: true, attempts, rollbacks };
    }

    rollbacks += 1;
  }

  return { success: false, attempts, rollbacks };
};

export const localSearchImprove = (
  state: InternalState,
  config: BalancingConfig,
  offset: number
): { lookaheadAttempts: number; lookaheadSuccess: number; lookaheadRollback: number } => {
  let lookaheadAttempts = 0;
  let lookaheadSuccess = 0;
  let lookaheadRollback = 0;

  let improved = true;
  while (improved) {
    improved = trySingleMoveImprove(state, config, offset);
    if (!improved) {
      break;
    }
  }

  const depth1 = tryLookaheadChain(state, config, 1, config.maxLookaheadAttempts, offset);
  lookaheadAttempts += depth1.attempts;
  lookaheadRollback += depth1.rollbacks;
  if (depth1.success) {
    lookaheadSuccess += 1;
    return { lookaheadAttempts, lookaheadSuccess, lookaheadRollback };
  }

  const remainingAttempts = Math.max(0, config.maxLookaheadAttempts - lookaheadAttempts);
  if (remainingAttempts > 0) {
    const depth2 = tryLookaheadChain(state, config, 2, remainingAttempts, offset);
    lookaheadAttempts += depth2.attempts;
    lookaheadRollback += depth2.rollbacks;
    if (depth2.success) {
      lookaheadSuccess += 1;
    }
  }

  return { lookaheadAttempts, lookaheadSuccess, lookaheadRollback };
};

const scoreFallbackCollisionMove = (
  state: InternalState,
  config: BalancingConfig,
  candidate: CandidateMove
): number => {
  return estimateMoveDelta(state, candidate, config, 'collision-fix', 0);
};

export const repairCollisions = (
  state: InternalState,
  config: BalancingConfig
): { unresolved: string[]; applied: MoveRecord[] } => {
  const unresolved: string[] = [];
  const applied: MoveRecord[] = [];

  const students = Array.from(state.students.values()).sort((left, right) => left.id.localeCompare(right.id));

  students.forEach((student) => {
    const byBlock = new Map<BlockNumber, AssignmentNode[]>();
    student.assignments.forEach((assignment) => {
      const entries = byBlock.get(assignment.block) || [];
      entries.push(assignment);
      byBlock.set(assignment.block, entries);
    });

    byBlock.forEach((assignmentsInBlock) => {
      if (assignmentsInBlock.length <= 1) {
        return;
      }

      const toMove = assignmentsInBlock.slice(1);

      toMove.forEach((assignment) => {
        const subjectGroups = state.groupsBySubject.get(assignment.subjectCode) || [];
        const sourceGroup = subjectGroups.find((group) => group.groupId === assignment.groupId);
        if (!sourceGroup) {
          unresolved.push(`${student.id}:${assignment.subjectCode}`);
          return;
        }

        const strictCandidates = subjectGroups
          .filter((group) => group.groupId !== sourceGroup.groupId)
          .map((target) => ({
            studentId: student.id,
            subjectCode: assignment.subjectCode,
            subjectName: assignment.subjectName,
            fromGroupId: sourceGroup.groupId,
            fromGroupCode: sourceGroup.groupCode,
            fromBlock: sourceGroup.block,
            toGroupId: target.groupId,
            toGroupCode: target.groupCode,
            toBlock: target.block,
            estimatedScoreDelta: 0,
            priorityTag: 'equalize' as const,
          }))
          .filter((candidate) => moveIsFeasible(state, candidate, 0, config.classBlockRestrictions));

        const pickStrict = [...strictCandidates]
          .map((candidate) => ({
            candidate,
            score: scoreFallbackCollisionMove(state, config, candidate),
          }))
          .filter((entry) => Number.isFinite(entry.score))
          .sort((left, right) => left.score - right.score)[0];

        if (pickStrict && pickStrict.score <= 0) {
          const success = applyCandidateMoveToState(
            state,
            pickStrict.candidate,
            'collision-fix',
            pickStrict.score,
            0,
            config.classBlockRestrictions
          );
          if (success) {
            applied.push(state.history[state.history.length - 1]);
            return;
          }
        }

        unresolved.push(`${student.id}:${assignment.subjectCode}`);
      });
    });
  });

  return { unresolved, applied };
};

const applyMovesToRows = (rows: StandardField[], moves: MoveRecord[]): StandardField[] => {
  const nextRows = rows.map((row) => ({ ...row }));
  const rowIndexByStudentId = new Map<string, number>();

  nextRows.forEach((row, index) => {
    rowIndexByStudentId.set(inferStudentId(row, index), index);
  });

  moves.forEach((move) => {
    const rowIndex = rowIndexByStudentId.get(move.studentId);
    if (rowIndex === undefined) {
      return;
    }

    const row = nextRows[rowIndex];
    const fromKey = getBlokkKey(move.fromBlock);
    const toKey = getBlokkKey(move.toBlock);

    const fromSubjects = parseSubjects(row[fromKey] as string | null);
    const toSubjects = parseSubjects(row[toKey] as string | null);

    const keepFrom = fromSubjects.filter((value) => normalizeSubject(value) !== normalizeSubject(move.subjectName));
    const hasInTo = toSubjects.some((value) => normalizeSubject(value) === normalizeSubject(move.subjectName));
    const nextTo = hasInTo ? toSubjects : [...toSubjects, move.subjectName];

    row[fromKey] = keepFrom.length > 0 ? keepFrom.join(', ') : null;
    row[toKey] = nextTo.length > 0 ? nextTo.join(', ') : null;
  });

  return nextRows;
};

const updateSubjectSettingsAssignments = (
  subjectSettingsByName: SubjectSettingsByNameLike,
  moveRecords: MoveRecord[],
  state: InternalState
): SubjectSettingsByNameLike => {
  const nextSettings: SubjectSettingsByNameLike = {};

  Object.entries(subjectSettingsByName).forEach(([subject, settings]) => {
    nextSettings[subject] = {
      ...settings,
      groups: settings.groups ? [...settings.groups] : settings.groups,
      groupStudentAssignments: { ...(settings.groupStudentAssignments || {}) },
    };
  });

  moveRecords.forEach((move) => {
    const subjectName = move.subjectName;
    const source = nextSettings[subjectName] || {
      defaultMax: 30,
      groups: [],
      groupStudentAssignments: {},
    };

    const groupNode = (state.groupsBySubject.get(move.subjectCode) || []).find((group) => {
      return group.groupCode === move.toGroupCode && group.block === move.toBlock;
    });

    if (!groupNode) {
      return;
    }

    nextSettings[subjectName] = {
      ...source,
      groupStudentAssignments: {
        ...(source.groupStudentAssignments || {}),
        [move.studentId]: groupNode.groupId,
      },
    };
  });

  return nextSettings;
};

const collectSubjectMetrics = (state: InternalState): SubjectMetrics[] => {
  const metrics: SubjectMetrics[] = [];

  state.groupsBySubject.forEach((groups) => {
    if (groups.length === 0) {
      return;
    }

    const sizes = groups.map((group) => group.size);
    const maxGroup = Math.max(...sizes);
    const imbalance = computeSubjectImbalanceRaw(sizes);

    metrics.push({
      subjectName: groups[0].subjectName,
      maxGroup,
      imbalance,
    });
  });

  metrics.sort((left, right) => left.subjectName.localeCompare(right.subjectName, 'nb', { sensitivity: 'base' }));
  return metrics;
};

const mergeConfig = (config?: Partial<BalancingConfig>): BalancingConfig => {
  return {
    ...DEFAULT_BALANCING_CONFIG,
    ...config,
    weights: {
      ...DEFAULT_BALANCING_CONFIG.weights,
      ...(config?.weights || {}),
    },
    classBlockRestrictions: {
      ...DEFAULT_BALANCING_CONFIG.classBlockRestrictions,
      ...(config?.classBlockRestrictions || {}),
    },
    lockedAssignmentKeys: config?.lockedAssignmentKeys || DEFAULT_BALANCING_CONFIG.lockedAssignmentKeys,
  };
};

export const progressiveHybridBalance = (
  rows: StandardField[],
  subjectSettingsByName: SubjectSettingsByNameLike,
  config?: Partial<BalancingConfig>
): ProgressiveHybridBalanceResult => {
  const mergedConfig = mergeConfig(config);
  let state = buildState(rows, subjectSettingsByName, mergedConfig);
  let bestStrictState = cloneState(state);

  const beforeScore = computeScore(state, mergedConfig.weights, state.history);
  const subjectMetricsBefore = collectSubjectMetrics(state);

  let passesRun = 0;
  let lookaheadAttempts = 0;
  let lookaheadSuccess = 0;
  let lookaheadRollback = 0;

  const startTime = Date.now();

  for (let offset = -mergedConfig.maxRelaxation; offset <= 0; offset += 1) {
    if (Date.now() - startTime > mergedConfig.maxPassMillis * (mergedConfig.maxRelaxation + 1)) {
      break;
    }

    passesRun += 1;
    let improving = true;

    while (improving) {
      if (Date.now() - startTime > mergedConfig.maxPassMillis * (mergedConfig.maxRelaxation + 1)) {
        improving = false;
        break;
      }

      const flow = buildFlowNetwork(state, mergedConfig, offset);
      const solved = solveFlow(flow);
      const extracted = extractMovesFromFlow(solved);
      const { applied } = applyMoves(state, extracted, mergedConfig, offset);

      if (applied.length === 0) {
        improving = false;
        break;
      }
    }

    const local = localSearchImprove(state, mergedConfig, offset);
    lookaheadAttempts += local.lookaheadAttempts;
    lookaheadSuccess += local.lookaheadSuccess;
    lookaheadRollback += local.lookaheadRollback;

    // Keep only states that satisfy strict capacity constraints as committable output.
    if (offset === 0) {
      bestStrictState = cloneState(state);
    }
  }

  state = cloneState(bestStrictState);

  const collisionRepair = repairCollisions(state, mergedConfig);
  const afterScore = computeScore(state, mergedConfig.weights, state.history);
  const subjectMetricsAfter = collectSubjectMetrics(state);

  const updatedData = applyMovesToRows(rows, state.history);
  const updatedSubjectSettingsByName = updateSubjectSettingsAssignments(subjectSettingsByName, state.history, state);

  const uniqueStudentsMoved = new Set(state.history.map((entry) => entry.studentId));
  const repeatedMoveCount = Array.from(state.studentMoveCounts.values()).reduce((sum, value) => sum + Math.max(0, value - 1), 0);

  const diagnostics: BalanceDiagnostics = {
    beforeScore,
    afterScore,
    subjectMetricsBefore,
    subjectMetricsAfter,
    moveCount: state.history.length,
    uniqueStudentsMoved: uniqueStudentsMoved.size,
    repeatedMoveCount,
    lookaheadAttempts,
    lookaheadSuccess,
    lookaheadRollback,
    unresolvedCollisions: collisionRepair.unresolved,
    passesRun,
  };

  return {
    moveRecords: state.history,
    updatedData,
    updatedSubjectSettingsByName,
    diagnostics,
  };
};
