import type { StandardField } from './excelUtils';
import { mapSubjectToCode } from './subjectCodeMapping';

export type BlockNumber = 1 | 2 | 3 | 4;

export interface SubjectGroupLike {
  id: string;
  blokk: string;
  sourceBlokk: string;
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
  isFourthYear: boolean;
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
  reason: 'overcap' | 'equalize' | 'peak-reduction' | 'collision-fix' | 'lookahead-chain' | 'student-rotation';
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

export interface BalancingProgress {
  phase: string;
  movesApplied: number;
  pass: number;
  totalPasses: number;
}

export interface BalanceDiagnostics {
  beforeScore: ScoreBreakdown;
  afterScore: ScoreBreakdown;
  beforeOvercapSeatCount: number;
  afterOvercapSeatCount: number;
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
  capacityOffsets?: number[];
  maxFlowIterationsPerOffset: number;
  maxLookaheadAttempts: number;
  maxDepth2Chains: number;
  classBlockRestrictions: ClassBlockRestrictions;
  excludedSubjects: string[];
  lockedAssignmentKeys: string[];
  blockCount?: number;
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

interface RotationSpec {
  subjectCode: string;
  subjectName: string;
  fromBlock: BlockNumber;
  fromGroupId: string;
  fromGroupCode: string;
  toBlock: BlockNumber;
}

interface RotationGroupTarget {
  assignment: RotationSpec;
  targetGroupId: string;
  targetGroupCode: string;
}

interface RotationCandidate {
  studentId: string;
  steps: RotationSpec[];
  estimatedScoreDelta: number;
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
const DEFAULT_PROGRESSIVE_CAPACITY_STEP = 2;
const ENABLE_CROSS_BLOCK_ROTATIONS = true;

const DEFAULT_WEIGHTS: BalancingWeights = {
  overcapA: 10,
  imbalanceB: 1,
  peakC: 1.8,
  collisionD: BIG_COLLISION_PENALTY,
  movesE: 0.7,
  repeatF: 1.6,
  alpha: 0.06,
  beta: 1,
};

export const DEFAULT_CLASS_BLOCK_RESTRICTIONS: ClassBlockRestrictions = {};

export const DEFAULT_BALANCING_CONFIG: BalancingConfig = {
  weights: DEFAULT_WEIGHTS,
  epsilon: 0.0001,
  maxRelaxation: 10,
  maxFlowIterationsPerOffset: 250,
  maxLookaheadAttempts: 300,
  maxDepth2Chains: 150,
  classBlockRestrictions: DEFAULT_CLASS_BLOCK_RESTRICTIONS,
  excludedSubjects: [],
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

type BlokkKey = string;

const getBlokkKey = (block: number): BlokkKey => {
  return `blokk${block}`;
};

const normalizeSubject = (value: string): string => value.trim().toLocaleLowerCase('nb');

const toExcludedSubjectSet = (subjects: string[]): Set<string> => {
  return new Set(subjects.map((subject) => normalizeSubject(subject)).filter((subject) => subject.length > 0));
};

const isExcludedSubjectName = (subjectName: string, excludedSubjects: Set<string>): boolean => {
  return excludedSubjects.has(normalizeSubject(subjectName));
};

const isExcludedSubjectCode = (state: InternalState, subjectCode: string, excludedSubjects: Set<string>): boolean => {
  const subjectName = state.groupsBySubject.get(subjectCode)?.[0]?.subjectName
    || Array.from(state.students.values())
      .flatMap((student) => student.assignments)
      .find((assignment) => assignment.subjectCode === subjectCode)?.subjectName
    || subjectCode;

  return isExcludedSubjectName(subjectName, excludedSubjects);
};

const isFreeRedistributionMove = (move: MoveRecord): boolean => {
  return move.fromBlock === move.toBlock;
};

const inferStudentId = (row: StandardField, index: number): string => {
  return row.studentId || `${row.navn || 'ukjent'}:${row.klasse || 'ukjent'}:${index}`;
};

const inferClassLevels = (classGroup: string, isFourthYear: boolean): string[] => {
  if (isFourthYear) {
    return ['VG4'];
  }

  const match = classGroup.trim().toUpperCase().match(/^(\d)/);
  if (!match) {
    const normalized = classGroup.trim().toUpperCase();
    return normalized ? [normalized] : [];
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

const blockFromLabel = (label: string): BlockNumber => {
  return Number.parseInt(label.replace('Blokk ', ''), 10) as BlockNumber;
};

const getEffectiveCapacity = (capacity: number, offset: number): number => {
  return Math.max(0, capacity - offset);
};

const buildCapacityOffsetSchedule = (maxRelaxation: number, step: number = DEFAULT_PROGRESSIVE_CAPACITY_STEP): number[] => {
  const normalizedRelaxation = Math.max(0, Math.floor(maxRelaxation));
  const normalizedStep = Math.max(1, Math.floor(step));

  if (normalizedRelaxation === 0) {
    return [0];
  }

  const offsets: number[] = [];
  for (let offset = normalizedRelaxation; offset > 0; offset -= normalizedStep) {
    offsets.push(offset);
  }
  offsets.push(0);

  return offsets;
};

const normalizeCapacityOffsets = (capacityOffsets: number[] | undefined, maxRelaxation: number): number[] => {
  if (!Array.isArray(capacityOffsets) || capacityOffsets.length === 0) {
    return buildCapacityOffsetSchedule(maxRelaxation);
  }

  const sanitized = Array.from(
    new Set(
      capacityOffsets
        .map((value) => Math.max(0, Math.floor(value)))
        .filter((value) => Number.isFinite(value))
    )
  ).sort((left, right) => right - left);

  if (sanitized.length === 0) {
    return buildCapacityOffsetSchedule(maxRelaxation);
  }

  if (sanitized[sanitized.length - 1] !== 0) {
    sanitized.push(0);
  }

  return sanitized;
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
  studentId: string,
  classGroup: string,
  block: BlockNumber,
  restrictions: ClassBlockRestrictions,
  isFourthYear: boolean
): boolean => {
  // Per-student override for fourth-year students takes precedence over VG4 class-level setting
  if (isFourthYear && restrictions[studentId]) {
    const studentRules = restrictions[studentId];
    const allowed = studentRules[block];
    return typeof allowed === 'boolean' ? allowed : true;
  }

  const classLevels = inferClassLevels(classGroup, isFourthYear);
  if (classLevels.length === 0) {
    return true;
  }

  return classLevels.some((classLevel) => {
    const levelRules = restrictions[classLevel];
    if (!levelRules) {
      return true;
    }

    const allowed = levelRules[block];
    if (typeof allowed === 'boolean') {
      return allowed;
    }

    return true;
  });
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
  occupancyByBlock: Record<number, number>,
  blockNumbers: BlockNumber[] = [1, 2, 3, 4]
): SubjectGroupLike[] => {
  const raw = subjectSettingsByName[subjectName];
  const defaultMax = raw?.defaultMax ?? 30;

  if (raw?.groups && raw.groups.length > 0) {
    return sortGroupsStable(raw.groups);
  }

  const autoGroups: SubjectGroupLike[] = [];
  blockNumbers.forEach((block) => {
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
  const occupancyBySubject = new Map<string, Record<number, number>>();
  const blockNumbers = Array.from(
    { length: Math.max(1, Math.min(8, mergedConfig.blockCount ?? 4)) },
    (_, i) => (i + 1) as BlockNumber
  );

  rows.forEach((row, index) => {
    const studentId = inferStudentId(row, index);
    const fullName = row.navn || 'Ukjent';
    const classGroup = row.klasse || '';
    const isFourthYear = row.fjerdearsElev === true;

    const assignments: AssignmentNode[] = [];

    blockNumbers.forEach((block) => {
      const subjectsInBlock = parseSubjects(row[getBlokkKey(block) as keyof StandardField] as string | null);

      subjectsInBlock.forEach((subjectName) => {
        const normalized = normalizeSubject(subjectName);
        const existing = occupancyBySubject.get(normalized) || {};
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
      isFourthYear,
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
    const occupancy = occupancyBySubject.get(normalizeSubject(subjectName)) || {};
    const sourceGroups = ensureSubjectGroups(subjectName, subjectSettingsByName, occupancy, blockNumbers);

    const indexedByBlock = new Map<number, SubjectGroupLike[]>();
    blockNumbers.forEach((block) => indexedByBlock.set(block, []));
    // Also index any blocks that exist in the groups but are outside blockNumbers
    sourceGroups.forEach((group) => {
      const block = blockFromLabel(group.blokk);
      if (!indexedByBlock.has(block)) {
        indexedByBlock.set(block, []);
      }
    });

    sourceGroups.forEach((group) => {
      if (!group.enabled) {
        return;
      }
      const block = blockFromLabel(group.blokk);
      indexedByBlock.get(block)?.push(group);
    });

    const builtGroups: GroupNode[] = [];

    Array.from(indexedByBlock.keys()).sort((a, b) => a - b).forEach((block) => {
      const blockNum = block as BlockNumber;
      const blockGroups = sortGroupsStable(indexedByBlock.get(block) || []);
      blockGroups.forEach((group, index) => {
        const groupCode = groupCodeByIndex(blockNum, index);
        const key = `${subjectCode}|${groupCode}|${block}`;
        const node: GroupNode = {
          key,
          subjectCode,
          subjectName,
          groupCode,
          groupId: group.id,
          block: blockNum,
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

const computeOvercapSeatCountForExcluded = (state: InternalState, excludedSubjects: Set<string>): number => {
  let seatCount = 0;
  state.groups.forEach((group) => {
    if (isExcludedSubjectName(group.subjectName, excludedSubjects)) {
      return;
    }
    seatCount += Math.max(0, group.size - group.capacity);
  });
  return seatCount;
};

const computeOvercapRaw = (state: InternalState, excludedSubjects: Set<string>): number => {
  let overcapRaw = 0;
  state.groups.forEach((group) => {
    if (isExcludedSubjectName(group.subjectName, excludedSubjects)) {
      return;
    }
    const excess = Math.max(0, group.size - group.capacity);
    overcapRaw += excess * excess;
  });
  return overcapRaw;
};

export const computeScore = (
  state: InternalState,
  weights: BalancingWeights,
  moveRecords: MoveRecord[],
  excludedSubjects: Set<string> = new Set<string>()
): ScoreBreakdown => {
  const overcapRaw = computeOvercapRaw(state, excludedSubjects);

  const sizesBySubject = getSubjectSizes(state);
  let imbalanceRaw = 0;
  let peakRaw = 0;

  sizesBySubject.forEach((sizes, subjectCode) => {
    if (isExcludedSubjectCode(state, subjectCode, excludedSubjects)) {
      return;
    }

    if (sizes.length === 0) {
      return;
    }

    const peak = Math.max(...sizes);
    const wSubject = 1 + weights.alpha * peak;
    imbalanceRaw += wSubject * computeSubjectImbalanceRaw(sizes);
    peakRaw += weights.beta * peak * peak;
  });

  const collisionRaw = Array.from(state.students.values()).reduce((sum, student) => {
    const includedAssignments = student.assignments.filter((assignment) => {
      return !isExcludedSubjectName(assignment.subjectName, excludedSubjects);
    });

    if (includedAssignments.length === student.assignments.length) {
      return sum + getCollisionCount(student);
    }

    const byBlock = new Map<BlockNumber, number>();
    includedAssignments.forEach((assignment) => {
      byBlock.set(assignment.block, (byBlock.get(assignment.block) || 0) + 1);
    });

    let collisions = 0;
    byBlock.forEach((count) => {
      if (count > 1) {
        collisions += count - 1;
      }
    });

    return sum + collisions;
  }, 0);

  const includedMoveRecords = moveRecords.filter((move) => {
    if (isExcludedSubjectName(move.subjectName, excludedSubjects)) {
      return false;
    }

    return !isFreeRedistributionMove(move);
  });
  const movesRaw = includedMoveRecords.length;
  const repeatCounts = new Map<string, number>();
  includedMoveRecords.forEach((move) => {
    repeatCounts.set(move.studentId, (repeatCounts.get(move.studentId) || 0) + 1);
  });
  const repeatedRaw = Array.from(repeatCounts.values()).reduce((sum, value) => sum + Math.max(0, value - 1), 0);

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
  restrictions: ClassBlockRestrictions,
  ignoreCapacity: boolean = false
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

  if (!isClassAllowedInBlock(student.id, student.classGroup, move.toBlock, restrictions, student.isFourthYear)) {
    return false;
  }

  if (move.toBlock !== move.fromBlock) {
    const hasOtherSubjectInTarget = student.assignments.some((item) => {
      return item.block === move.toBlock && item.subjectCode !== move.subjectCode;
    });

    if (hasOtherSubjectInTarget) {
      return false;
    }
  }

  const targetGroup = Array.from(state.groups.values()).find((group) => group.groupId === move.toGroupId);
  if (!targetGroup) {
    return false;
  }

  if (!ignoreCapacity) {
    // Offset>0 temporarily lowers usable capacity. Offset=0 is the real group max.
    const effectiveCapacity = getEffectiveCapacity(targetGroup.capacity, offset);
    if (targetGroup.size + 1 > effectiveCapacity) {
      return false;
    }
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
  chainStep?: number,
  ignoreCapacity: boolean = false
): boolean => {
  if (!moveIsFeasible(state, move, offset, restrictions, ignoreCapacity)) {
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

const findAssignmentIndex = (student: StudentNode, spec: RotationSpec): number => {
  return student.assignments.findIndex((assignment) => {
    return assignment.subjectCode === spec.subjectCode
      && assignment.block === spec.fromBlock
      && assignment.groupId === spec.fromGroupId;
  });
};

const pickRotationTargets = (
  state: InternalState,
  student: StudentNode,
  steps: RotationSpec[],
  offset: number,
  restrictions: ClassBlockRestrictions
): RotationGroupTarget[] | null => {
  const provisionalGroupDelta = new Map<string, number>();
  const outgoingBlocks = new Set(steps.map((step) => step.fromBlock));
  const incomingBlocks = new Set(steps.map((step) => step.toBlock));

  const occupiedBlocksAfterOutgoing = new Set(
    student.assignments
      .filter((assignment) => !outgoingBlocks.has(assignment.block))
      .map((assignment) => assignment.block)
  );

  for (const block of incomingBlocks) {
    if (occupiedBlocksAfterOutgoing.has(block)) {
      return null;
    }
  }

  steps.forEach((step) => {
    const sourceGroup = Array.from(state.groups.values()).find((group) => group.groupId === step.fromGroupId);
    if (sourceGroup) {
      provisionalGroupDelta.set(sourceGroup.key, (provisionalGroupDelta.get(sourceGroup.key) || 0) - 1);
    }
  });

  const targets: RotationGroupTarget[] = [];

  for (const step of steps) {
    if (!isClassAllowedInBlock(student.id, student.classGroup, step.toBlock, restrictions, student.isFourthYear)) {
      return null;
    }

    const targetGroups = (state.groupsBySubject.get(step.subjectCode) || [])
      .filter((group) => group.block === step.toBlock)
      .sort((left, right) => {
        const leftLoad = (left.size + (provisionalGroupDelta.get(left.key) || 0)) / Math.max(1, left.capacity);
        const rightLoad = (right.size + (provisionalGroupDelta.get(right.key) || 0)) / Math.max(1, right.capacity);
        if (leftLoad !== rightLoad) {
          return leftLoad - rightLoad;
        }
        if (left.groupCode !== right.groupCode) {
          return left.groupCode.localeCompare(right.groupCode, 'nb', { sensitivity: 'base' });
        }
        return left.groupId.localeCompare(right.groupId, 'nb', { sensitivity: 'base' });
      });

    const chosen = targetGroups.find((group) => {
      const effectiveCapacity = getEffectiveCapacity(group.capacity, offset);
      return group.size + (provisionalGroupDelta.get(group.key) || 0) + 1 <= effectiveCapacity;
    });

    if (!chosen) {
      return null;
    }

    provisionalGroupDelta.set(chosen.key, (provisionalGroupDelta.get(chosen.key) || 0) + 1);
    targets.push({
      assignment: step,
      targetGroupId: chosen.groupId,
      targetGroupCode: chosen.groupCode,
    });
  }

  return targets;
};

const applyStudentRotationToState = (
  state: InternalState,
  rotation: RotationCandidate,
  scoreDelta: number,
  offset: number,
  restrictions: ClassBlockRestrictions
): boolean => {
  const student = state.students.get(rotation.studentId);
  if (!student) {
    return false;
  }

  const uniqueBlocks = new Set(rotation.steps.map((step) => step.fromBlock));
  if (uniqueBlocks.size !== rotation.steps.length) {
    return false;
  }

  const targets = pickRotationTargets(state, student, rotation.steps, offset, restrictions);
  if (!targets) {
    return false;
  }

  const stepIndices = rotation.steps.map((step) => findAssignmentIndex(student, step));
  if (stepIndices.some((index) => index < 0)) {
    return false;
  }

  targets.forEach((target) => {
    const sourceGroup = Array.from(state.groups.values()).find((group) => group.groupId === target.assignment.fromGroupId);
    const destinationGroup = Array.from(state.groups.values()).find((group) => group.groupId === target.targetGroupId);
    if (!sourceGroup || !destinationGroup) {
      return;
    }

    sourceGroup.studentIds.delete(student.id);
    sourceGroup.size = Math.max(0, sourceGroup.size - 1);
    destinationGroup.studentIds.add(student.id);
    destinationGroup.size += 1;
  });

  targets.forEach((target, index) => {
    const assignmentIndex = stepIndices[index];
    const assignment = student.assignments[assignmentIndex];
    assignment.block = target.assignment.toBlock;
    assignment.groupId = target.targetGroupId;
    assignment.groupCode = target.targetGroupCode;

    state.history.push({
      studentId: student.id,
      studentName: student.fullName,
      subjectCode: assignment.subjectCode,
      subjectName: assignment.subjectName,
      fromGroupCode: target.assignment.fromGroupCode,
      fromBlock: target.assignment.fromBlock,
      toGroupCode: target.targetGroupCode,
      toBlock: target.assignment.toBlock,
      reason: 'student-rotation',
      scoreDelta,
      chainStep: index + 1,
    });
  });

  state.studentMoveCounts.set(student.id, (state.studentMoveCounts.get(student.id) || 0) + rotation.steps.length);
  return true;
};

const estimateMoveDelta = (
  state: InternalState,
  move: CandidateMove,
  config: BalancingConfig,
  reason: MoveRecord['reason'],
  offset: number,
  ignoreCapacity: boolean = false
): number => {
  if (!moveIsFeasible(state, move, offset, config.classBlockRestrictions, ignoreCapacity)) {
    return Number.POSITIVE_INFINITY;
  }

  const excludedSubjects = toExcludedSubjectSet(config.excludedSubjects);
  const before = computeScore(state, config.weights, state.history, excludedSubjects).total;
  const cloned = cloneState(state);
  const applied = applyCandidateMoveToState(cloned, move, reason, 0, offset, config.classBlockRestrictions, undefined, ignoreCapacity);
  if (!applied) {
    return Number.POSITIVE_INFINITY;
  }

  const after = computeScore(cloned, config.weights, cloned.history, excludedSubjects).total;
  return after - before;
};

export const buildFlowNetwork = (
  state: InternalState,
  config: BalancingConfig,
  offset: number
): FlowNetwork => {
  const candidates: CandidateMove[] = [];
  const excludedSubjects = toExcludedSubjectSet(config.excludedSubjects);

  state.students.forEach((student) => {
    student.assignments.forEach((assignment) => {
      if (assignment.locked) {
        return;
      }

      if (isExcludedSubjectName(assignment.subjectName, excludedSubjects)) {
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

const buildRotationCandidate = (
  state: InternalState,
  student: StudentNode,
  steps: RotationSpec[],
  config: BalancingConfig,
  offset: number
): RotationCandidate | null => {
  const uniqueBlocks = new Set(steps.map((step) => step.fromBlock));
  if (uniqueBlocks.size !== steps.length) {
    return null;
  }

  const snapshot = cloneState(state);
  const applied = applyStudentRotationToState(snapshot, {
    studentId: student.id,
    steps,
    estimatedScoreDelta: 0,
  }, 0, offset, config.classBlockRestrictions);

  if (!applied) {
    return null;
  }

  const excludedSubjects = toExcludedSubjectSet(config.excludedSubjects);
  const before = computeScore(state, config.weights, state.history, excludedSubjects).total;
  const after = computeScore(snapshot, config.weights, snapshot.history, excludedSubjects).total;
  const delta = after - before;
  if (!Number.isFinite(delta) || delta >= -config.epsilon) {
    return null;
  }

  return {
    studentId: student.id,
    steps,
    estimatedScoreDelta: delta,
  };
};

const generateStudentRotationCandidates = (
  state: InternalState,
  config: BalancingConfig,
  offset: number
): RotationCandidate[] => {
  const rotations: RotationCandidate[] = [];

  const students = Array.from(state.students.values()).sort((left, right) => left.id.localeCompare(right.id));
  students.forEach((student) => {
    const blockCounts = new Map<BlockNumber, number>();
    student.assignments.forEach((assignment) => {
      blockCounts.set(assignment.block, (blockCounts.get(assignment.block) || 0) + 1);
    });

    const eligible = student.assignments
      .filter((assignment) => !assignment.locked)
      .filter((assignment) => (blockCounts.get(assignment.block) || 0) === 1)
      .sort((left, right) => {
        if (left.block !== right.block) {
          return left.block - right.block;
        }
        if (left.subjectCode !== right.subjectCode) {
          return left.subjectCode.localeCompare(right.subjectCode, 'nb', { sensitivity: 'base' });
        }
        return left.groupId.localeCompare(right.groupId, 'nb', { sensitivity: 'base' });
      });

    for (let i = 0; i < eligible.length; i += 1) {
      for (let j = i + 1; j < eligible.length; j += 1) {
        const first = eligible[i];
        const second = eligible[j];
        const pairSteps: RotationSpec[] = [
          {
            subjectCode: first.subjectCode,
            subjectName: first.subjectName,
            fromBlock: first.block,
            fromGroupId: first.groupId,
            fromGroupCode: first.groupCode,
            toBlock: second.block,
          },
          {
            subjectCode: second.subjectCode,
            subjectName: second.subjectName,
            fromBlock: second.block,
            fromGroupId: second.groupId,
            fromGroupCode: second.groupCode,
            toBlock: first.block,
          },
        ];

        const pairCandidate = buildRotationCandidate(state, student, pairSteps, config, offset);
        if (pairCandidate) {
          rotations.push(pairCandidate);
        }

        for (let k = j + 1; k < eligible.length; k += 1) {
          const third = eligible[k];
          const clockwise: RotationSpec[] = [
            {
              subjectCode: first.subjectCode,
              subjectName: first.subjectName,
              fromBlock: first.block,
              fromGroupId: first.groupId,
              fromGroupCode: first.groupCode,
              toBlock: second.block,
            },
            {
              subjectCode: second.subjectCode,
              subjectName: second.subjectName,
              fromBlock: second.block,
              fromGroupId: second.groupId,
              fromGroupCode: second.groupCode,
              toBlock: third.block,
            },
            {
              subjectCode: third.subjectCode,
              subjectName: third.subjectName,
              fromBlock: third.block,
              fromGroupId: third.groupId,
              fromGroupCode: third.groupCode,
              toBlock: first.block,
            },
          ];

          const counterClockwise: RotationSpec[] = [
            {
              subjectCode: first.subjectCode,
              subjectName: first.subjectName,
              fromBlock: first.block,
              fromGroupId: first.groupId,
              fromGroupCode: first.groupCode,
              toBlock: third.block,
            },
            {
              subjectCode: third.subjectCode,
              subjectName: third.subjectName,
              fromBlock: third.block,
              fromGroupId: third.groupId,
              fromGroupCode: third.groupCode,
              toBlock: second.block,
            },
            {
              subjectCode: second.subjectCode,
              subjectName: second.subjectName,
              fromBlock: second.block,
              fromGroupId: second.groupId,
              fromGroupCode: second.groupCode,
              toBlock: first.block,
            },
          ];

          const clockwiseCandidate = buildRotationCandidate(state, student, clockwise, config, offset);
          if (clockwiseCandidate) {
            rotations.push(clockwiseCandidate);
          }

          const counterClockwiseCandidate = buildRotationCandidate(state, student, counterClockwise, config, offset);
          if (counterClockwiseCandidate) {
            rotations.push(counterClockwiseCandidate);
          }
        }
      }
    }
  });

  rotations.sort((left, right) => {
    if (left.estimatedScoreDelta !== right.estimatedScoreDelta) {
      return left.estimatedScoreDelta - right.estimatedScoreDelta;
    }
    if (left.steps.length !== right.steps.length) {
      return left.steps.length - right.steps.length;
    }
    return left.studentId.localeCompare(right.studentId, 'nb', { sensitivity: 'base' });
  });

  return rotations;
};

const tryStudentRotationImprove = (state: InternalState, config: BalancingConfig, offset: number): boolean => {
  const bestRotation = generateStudentRotationCandidates(state, config, offset)[0];
  if (!bestRotation) {
    return false;
  }

  return applyStudentRotationToState(
    state,
    bestRotation,
    bestRotation.estimatedScoreDelta,
    offset,
    config.classBlockRestrictions
  );
};

const repairOvercapacity = (
  state: InternalState,
  config: BalancingConfig,
  offset: number
): number => {
  let appliedCount = 0;
  let improving = true;

  while (improving) {
    improving = false;

    const excludedSubjects = toExcludedSubjectSet(config.excludedSubjects);
    const baselineSeatCount = computeOvercapSeatCountForExcluded(state, excludedSubjects);
    const baselineRaw = computeOvercapRaw(state, excludedSubjects);

    if (baselineSeatCount <= 0) {
      break;
    }

    const candidateEvaluations = generateAllCandidateMoves(state, config, offset)
      .filter((candidate) => candidate.priorityTag === 'overcap')
      .map((candidate) => {
        const snapshot = cloneState(state);
        const applied = applyCandidateMoveToState(snapshot, candidate, 'overcap', 0, offset, config.classBlockRestrictions);
        if (!applied) {
          return null;
        }

        const nextSeatCount = computeOvercapSeatCountForExcluded(snapshot, excludedSubjects);
        const nextRaw = computeOvercapRaw(snapshot, excludedSubjects);
        const strictImprovement = baselineRaw - nextRaw;
        const seatImprovement = baselineSeatCount - nextSeatCount;

        if (strictImprovement <= 0 && seatImprovement <= 0) {
          return null;
        }

        const scoreDelta = estimateMoveDelta(state, candidate, config, 'overcap', offset);
        const repeatMoves = state.studentMoveCounts.get(candidate.studentId) || 0;
        return {
          candidate,
          strictImprovement,
          seatImprovement,
          scoreDelta,
          repeatMoves,
        };
      })
      .filter((entry): entry is NonNullable<typeof entry> => !!entry)
      .sort((left, right) => {
        if (left.strictImprovement !== right.strictImprovement) {
          return right.strictImprovement - left.strictImprovement;
        }
        if (left.seatImprovement !== right.seatImprovement) {
          return right.seatImprovement - left.seatImprovement;
        }
        if (left.repeatMoves !== right.repeatMoves) {
          return left.repeatMoves - right.repeatMoves;
        }
        return compareCandidate(left.candidate, right.candidate);
      });

    const best = candidateEvaluations[0];
    if (!best) {
      break;
    }

    const committed = applyCandidateMoveToState(
      state,
      best.candidate,
      'overcap',
      Number.isFinite(best.scoreDelta) ? best.scoreDelta : 0,
      offset,
      config.classBlockRestrictions
    );

    if (!committed) {
      break;
    }

    appliedCount += 1;
    improving = true;
  }

  return appliedCount;
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
    const excludedSubjects = toExcludedSubjectSet(config.excludedSubjects);
    const baseline = computeScore(state, config.weights, state.history, excludedSubjects).total;
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
      const after = computeScore(snapshot, config.weights, snapshot.history, excludedSubjects).total;
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

        const after = computeScore(snapshot, config.weights, snapshot.history, excludedSubjects).total;
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
    if (!improved && ENABLE_CROSS_BLOCK_ROTATIONS) {
      improved = tryStudentRotationImprove(state, config, offset);
    }
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
  return estimateMoveDelta(state, candidate, config, 'collision-fix', 0, true);
};

// Build all feasible single moves for a student in the current state (ignoring capacity).
// Each entry is tagged `directFix=true` when the move takes a subject out of a colliding
// block and into a block that is currently free for the student — i.e. it immediately
// reduces the collision count by 1.
const buildStudentCollisionMoves = (
  state: InternalState,
  studentId: string,
  config: BalancingConfig
): Array<{ candidate: CandidateMove; score: number; directFix: boolean }> => {
  const student = state.students.get(studentId);
  if (!student) {
    return [];
  }

  const result: Array<{ candidate: CandidateMove; score: number; directFix: boolean }> = [];
  const excludedSubjects = toExcludedSubjectSet(config.excludedSubjects);

  student.assignments.forEach((assignment) => {
    if (assignment.locked) {
      return;
    }

    if (isExcludedSubjectName(assignment.subjectName, excludedSubjects)) {
      return;
    }

    const subjectGroups = state.groupsBySubject.get(assignment.subjectCode) || [];
    const sourceGroup = subjectGroups.find((group) => group.groupId === assignment.groupId);
    if (!sourceGroup) {
      return;
    }

    const sourceColliding = student.assignments.filter((a) => a.block === assignment.block).length > 1;

    subjectGroups
      .filter((group) => group.groupId !== sourceGroup.groupId && group.block !== sourceGroup.block)
      .forEach((target) => {
        const candidate: CandidateMove = {
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
        };

        if (!moveIsFeasible(state, candidate, 0, config.classBlockRestrictions, true)) {
          return;
        }

        const targetFree = !student.assignments.some(
          (a) => a.block === target.block && a.subjectCode !== assignment.subjectCode
        );
        const directFix = sourceColliding && targetFree;
        const score = scoreFallbackCollisionMove(state, config, candidate);

        if (Number.isFinite(score)) {
          result.push({ candidate, score, directFix });
        }
      });
  });

  result.sort((left, right) => left.score - right.score);
  return result;
};

export const repairCollisions = (
  state: InternalState,
  config: BalancingConfig
): { unresolved: string[]; applied: MoveRecord[] } => {
  const unresolved: string[] = [];
  const applied: MoveRecord[] = [];

  const students = Array.from(state.students.values()).sort((left, right) => left.id.localeCompare(right.id));

  students.forEach((student) => {
    let madeProgress = true;

    while (madeProgress) {
      madeProgress = false;

      const current = state.students.get(student.id);
      if (!current || getCollisionCount(current) === 0) {
        break;
      }

      const moves = buildStudentCollisionMoves(state, student.id, config);

      // ── Step 1: direct fix — move a subject out of a colliding block into a free block ──
      const directFix = moves.find((m) => m.directFix);
      if (directFix) {
        const success = applyCandidateMoveToState(
          state,
          directFix.candidate,
          'collision-fix',
          directFix.score,
          0,
          config.classBlockRestrictions,
          undefined,
          true
        );
        if (success) {
          applied.push(state.history[state.history.length - 1]);
          madeProgress = true;
          // restart the while loop to re-evaluate collisions
        }
        continue;
      }

      // ── Step 2: 2-step lookahead ──
      // No single move resolves the collision directly (e.g. every free block is occupied
      // by another subject that could be moved elsewhere first).
      // Try every feasible move for this student; if applying it to a snapshot enables a
      // direct fix, commit both moves to the real state.
      let foundTwoStep = false;
      for (const { candidate: firstCandidate } of moves) {
        const snapshot = cloneState(state);
        const firstApplied = applyCandidateMoveToState(
          snapshot,
          firstCandidate,
          'collision-fix',
          0,
          0,
          config.classBlockRestrictions,
          undefined,
          true
        );
        if (!firstApplied) {
          continue;
        }

        // Check whether a direct fix is now available in the snapshot
        const snapshotMoves = buildStudentCollisionMoves(snapshot, student.id, config);
        const snapshotDirectFix = snapshotMoves.find((m) => m.directFix);

        if (!snapshotDirectFix) {
          continue;
        }

        // Commit first move to the real state
        const firstScore = scoreFallbackCollisionMove(state, config, firstCandidate);
        const committed1 = applyCandidateMoveToState(
          state,
          firstCandidate,
          'collision-fix',
          Number.isFinite(firstScore) ? firstScore : 0,
          0,
          config.classBlockRestrictions,
          undefined,
          true
        );

        if (!committed1) {
          continue;
        }

        applied.push(state.history[state.history.length - 1]);

        // Commit second move to the real state (re-score against updated state)
        const secondScore = scoreFallbackCollisionMove(state, config, snapshotDirectFix.candidate);
        const committed2 = applyCandidateMoveToState(
          state,
          snapshotDirectFix.candidate,
          'collision-fix',
          Number.isFinite(secondScore) ? secondScore : 0,
          0,
          config.classBlockRestrictions,
          undefined,
          true
        );

        if (committed2) {
          applied.push(state.history[state.history.length - 1]);
        }

        madeProgress = true;
        foundTwoStep = true;
        break;
      }

      if (!foundTwoStep) {
        break;
      }
    }

    // Report any collisions that are genuinely unresolvable.
    const finalStudent = state.students.get(student.id);
    if (!finalStudent) {
      return;
    }

    const finalByBlock = new Map<BlockNumber, AssignmentNode[]>();
    finalStudent.assignments.forEach((assignment) => {
      const entries = finalByBlock.get(assignment.block) || [];
      entries.push(assignment);
      finalByBlock.set(assignment.block, entries);
    });

    finalByBlock.forEach((assignmentsInBlock) => {
      if (assignmentsInBlock.length <= 1) {
        return;
      }
      assignmentsInBlock.slice(1).forEach((assignment) => {
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

    const row = nextRows[rowIndex] as unknown as Record<string, string | null>;
    const fromKey = getBlokkKey(move.fromBlock);
    const toKey = getBlokkKey(move.toBlock);

    const fromSubjects = parseSubjects(row[fromKey]);
    const toSubjects = parseSubjects(row[toKey]);

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
  const touchedSubjects = new Set<string>();

  Object.entries(subjectSettingsByName).forEach(([subject, settings]) => {
    nextSettings[subject] = {
      ...settings,
      groups: settings.groups ? [...settings.groups] : settings.groups,
      groupStudentAssignments: { ...(settings.groupStudentAssignments || {}) },
    };
  });

  moveRecords.forEach((move) => {
    const subjectName = move.subjectName;
    touchedSubjects.add(subjectName);
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

  touchedSubjects.forEach((subjectName) => {
    const subjectCode = mapSubjectToCode(subjectName);
    const groupsInState = state.groupsBySubject.get(subjectCode) || [];
    if (groupsInState.length === 0) {
      return;
    }

    const existing = nextSettings[subjectName] || {
      defaultMax: Math.max(...groupsInState.map((group) => group.capacity)),
      groups: [],
      groupStudentAssignments: {},
    };

    const finalizedAssignments: Record<string, string> = {};
    state.students.forEach((student) => {
      student.assignments.forEach((assignment) => {
        if (assignment.subjectCode !== subjectCode) {
          return;
        }

        finalizedAssignments[student.id] = assignment.groupId;
      });
    });

    const hasPersistedGroups = Array.isArray(existing.groups) && existing.groups.length > 0;
    const materializedGroups = hasPersistedGroups
      ? [...(existing.groups || [])]
      : [...groupsInState]
        .sort((left, right) => {
          if (left.block !== right.block) {
            return left.block - right.block;
          }
          if (left.groupCode !== right.groupCode) {
            return left.groupCode.localeCompare(right.groupCode, 'nb', { sensitivity: 'base' });
          }
          return left.groupId.localeCompare(right.groupId, 'nb', { sensitivity: 'base' });
        })
        .map((group) => ({
          id: group.groupId,
          blokk: `Blokk ${group.block}` as `Blokk ${BlockNumber}`,
          sourceBlokk: `Blokk ${group.block}` as `Blokk ${BlockNumber}`,
          enabled: true,
          max: group.capacity,
          createdAt: `balanced-${subjectCode}-${group.groupId}`,
        }));

    nextSettings[subjectName] = {
      ...existing,
      defaultMax: existing.defaultMax || Math.max(...materializedGroups.map((group) => group.max)),
      groups: materializedGroups,
      groupStudentAssignments: finalizedAssignments,
    };
  });

  return nextSettings;
};

const collectSubjectMetrics = (state: InternalState, excludedSubjects: Set<string> = new Set<string>()): SubjectMetrics[] => {
  const metrics: SubjectMetrics[] = [];

  state.groupsBySubject.forEach((groups) => {
    if (groups.length === 0) {
      return;
    }

    if (isExcludedSubjectName(groups[0].subjectName, excludedSubjects)) {
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
    capacityOffsets: Array.isArray(config?.capacityOffsets)
      ? [...config.capacityOffsets]
      : DEFAULT_BALANCING_CONFIG.capacityOffsets,
    excludedSubjects: config?.excludedSubjects || DEFAULT_BALANCING_CONFIG.excludedSubjects,
    lockedAssignmentKeys: config?.lockedAssignmentKeys || DEFAULT_BALANCING_CONFIG.lockedAssignmentKeys,
  };
};

export const progressiveHybridBalance = (
  rows: StandardField[],
  subjectSettingsByName: SubjectSettingsByNameLike,
  config?: Partial<BalancingConfig>,
  onProgress?: (progress: BalancingProgress) => void
): ProgressiveHybridBalanceResult => {
  const mergedConfig = mergeConfig(config);
  const excludedSubjects = toExcludedSubjectSet(mergedConfig.excludedSubjects);
  let state = buildState(rows, subjectSettingsByName, mergedConfig);
  const beforeOvercapSeatCount = computeOvercapSeatCountForExcluded(state, excludedSubjects);

  const beforeScore = computeScore(state, mergedConfig.weights, state.history, excludedSubjects);
  const subjectMetricsBefore = collectSubjectMetrics(state, excludedSubjects);

  // Phase 0: Resolve block-collisions first, ignoring group capacity.
  // Only subjects that truly cannot be placed in separate blocks will remain as warnings.
  repairCollisions(state, mergedConfig);
  let bestStrictState: InternalState | null = null;

  let passesRun = 0;
  let lookaheadAttempts = 0;
  let lookaheadSuccess = 0;
  let lookaheadRollback = 0;

  const capacityOffsets = normalizeCapacityOffsets(mergedConfig.capacityOffsets, mergedConfig.maxRelaxation);
  const totalPasses = capacityOffsets.length;

  onProgress?.({ phase: 'Reparerer kollisjoner...', movesApplied: state.history.length, pass: 0, totalPasses });

  for (const offset of capacityOffsets) {
    passesRun += 1;
    onProgress?.({ phase: `Balanserer runde ${passesRun} / ${totalPasses} (kapasitetsavstand ${offset})...`, movesApplied: state.history.length, pass: passesRun, totalPasses });
    repairOvercapacity(state, mergedConfig, offset);
    let improving = true;
    let flowIterations = 0;

    while (improving && flowIterations < mergedConfig.maxFlowIterationsPerOffset) {
      flowIterations += 1;

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

    onProgress?.({ phase: `Runde ${passesRun} ferdig — ${state.history.length} flytt totalt`, movesApplied: state.history.length, pass: passesRun, totalPasses });

    // Keep only states that satisfy strict capacity constraints as committable output.
    if (offset === 0) {
      bestStrictState = cloneState(state);
    }
  }

  // If a custom capacity schedule somehow omitted offset 0, keep the best computed state.
  state = cloneState(bestStrictState || state);

  const collisionRepair = repairCollisions(state, mergedConfig);
  const afterScore = computeScore(state, mergedConfig.weights, state.history, excludedSubjects);
  const subjectMetricsAfter = collectSubjectMetrics(state, excludedSubjects);

  const updatedData = applyMovesToRows(rows, state.history);
  const updatedSubjectSettingsByName = updateSubjectSettingsAssignments(subjectSettingsByName, state.history, state);

  const uniqueStudentsMoved = new Set(state.history.map((entry) => entry.studentId));
  const repeatedMoveCount = Array.from(state.studentMoveCounts.values()).reduce((sum, value) => sum + Math.max(0, value - 1), 0);

  const diagnostics: BalanceDiagnostics = {
    beforeScore,
    afterScore,
    beforeOvercapSeatCount,
    afterOvercapSeatCount: computeOvercapSeatCountForExcluded(state, excludedSubjects),
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
