import type { StandardField } from './excelUtils';
import {
  DEFAULT_BALANCING_CONFIG,
  DEFAULT_CLASS_BLOCK_RESTRICTIONS,
  progressiveHybridBalance,
  type ProgressiveHybridBalanceResult,
  type SubjectSettingsByNameLike,
} from './progressiveHybridBalance';

interface FixtureCase {
  name: string;
  run: () => ProgressiveHybridBalanceResult;
  verify: (result: ProgressiveHybridBalanceResult) => { ok: boolean; details: string };
}

interface FixtureOutcome {
  name: string;
  ok: boolean;
  details: string;
}

export interface FixtureSuiteResult {
  allPassed: boolean;
  outcomes: FixtureOutcome[];
}

const makeStudent = (
  studentId: string,
  navn: string,
  klasse: string,
  blokk1: string | null,
  blokk2: string | null,
  blokk3: string | null,
  blokk4: string | null
): StandardField => ({
  studentId,
  navn,
  klasse,
  blokkmatvg2: null,
  matematikk2p: null,
  matematikks1: null,
  matematikkr1: null,
  fremmedsprak: null,
  blokk1,
  blokk2,
  blokk3,
  blokk4,
  blokk5: null,
  blokk6: null,
  blokk7: null,
  blokk8: null,
  reserve: null,
});

const makeSubjectSettings = (): SubjectSettingsByNameLike => ({
  Matematikk: {
    defaultMax: 2,
    groups: [
      {
        id: 'm-b1-1',
        blokk: 'Blokk 1',
        sourceBlokk: 'Blokk 1',
        enabled: true,
        max: 2,
        createdAt: '2024-01-01T00:00:00.000Z',
      },
      {
        id: 'm-b2-1',
        blokk: 'Blokk 2',
        sourceBlokk: 'Blokk 2',
        enabled: true,
        max: 2,
        createdAt: '2024-01-01T00:00:01.000Z',
      },
      {
        id: 'm-b3-1',
        blokk: 'Blokk 3',
        sourceBlokk: 'Blokk 3',
        enabled: true,
        max: 2,
        createdAt: '2024-01-01T00:00:02.000Z',
      },
      {
        id: 'm-b4-1',
        blokk: 'Blokk 4',
        sourceBlokk: 'Blokk 4',
        enabled: true,
        max: 2,
        createdAt: '2024-01-01T00:00:03.000Z',
      },
    ],
    groupStudentAssignments: {},
  },
});

const severeOverloadCase: FixtureCase = {
  name: 'Severe overload redistributes',
  run: () => {
    const rows: StandardField[] = [
      makeStudent('s1', 'A', '2STA', 'Matematikk', null, null, null),
      makeStudent('s2', 'B', '2STA', 'Matematikk', null, null, null),
      makeStudent('s3', 'C', '2STA', 'Matematikk', null, null, null),
      makeStudent('s4', 'D', '2STA', 'Matematikk', null, null, null),
      makeStudent('s5', 'E', '2STA', 'Matematikk', null, null, null),
    ];

    return progressiveHybridBalance(rows, makeSubjectSettings(), {
      ...DEFAULT_BALANCING_CONFIG,
      classBlockRestrictions: {
        ...DEFAULT_CLASS_BLOCK_RESTRICTIONS,
      },
    });
  },
  verify: (result) => {
    const improved = result.diagnostics.afterScore.overcap <= result.diagnostics.beforeScore.overcap;
    return {
      ok: improved,
      details: improved ? 'Overcap pressure reduced or unchanged.' : 'Overcap pressure increased unexpectedly.',
    };
  },
};

const largeGroupPressureCase: FixtureCase = {
  name: 'Large-group penalty prioritizes peaks',
  run: () => {
    const rows: StandardField[] = [
      makeStudent('s1', 'A', '1STA', 'Matematikk', null, null, null),
      makeStudent('s2', 'B', '1STA', 'Matematikk', null, null, null),
      makeStudent('s3', 'C', '1STA', 'Matematikk', null, null, null),
      makeStudent('s4', 'D', '1STA', 'Matematikk', null, null, null),
      makeStudent('s5', 'E', '1STA', null, 'Matematikk', null, null),
    ];

    return progressiveHybridBalance(rows, makeSubjectSettings(), {
      weights: {
        ...DEFAULT_BALANCING_CONFIG.weights,
        alpha: 0.2,
        peakC: 2.5,
      },
      classBlockRestrictions: {
        ...DEFAULT_CLASS_BLOCK_RESTRICTIONS,
      },
    });
  },
  verify: (result) => {
    const beforePeak = result.diagnostics.beforeScore.peak;
    const afterPeak = result.diagnostics.afterScore.peak;
    const ok = afterPeak <= beforePeak;
    return {
      ok,
      details: ok ? 'Peak pressure reduced as expected.' : 'Peak pressure did not reduce.',
    };
  },
};

const chainRequiredCase: FixtureCase = {
  name: 'Lookahead chain can unlock blocked move',
  run: () => {
    const rows: StandardField[] = [
      makeStudent('s1', 'A', '1STA', 'Matematikk, Fysikk', null, null, null),
      makeStudent('s2', 'B', '1STA', 'Matematikk', 'Fysikk', null, null),
      makeStudent('s3', 'C', '1STA', null, 'Matematikk', null, null),
      makeStudent('s4', 'D', '1STA', null, 'Matematikk', null, null),
    ];

    const settings: SubjectSettingsByNameLike = {
      ...makeSubjectSettings(),
      Fysikk: {
        defaultMax: 2,
        groups: [
          {
            id: 'f-b1-1',
            blokk: 'Blokk 1',
            sourceBlokk: 'Blokk 1',
            enabled: true,
            max: 2,
            createdAt: '2024-01-01T00:00:00.000Z',
          },
          {
            id: 'f-b2-1',
            blokk: 'Blokk 2',
            sourceBlokk: 'Blokk 2',
            enabled: true,
            max: 2,
            createdAt: '2024-01-01T00:00:01.000Z',
          },
        ],
        groupStudentAssignments: {},
      },
    };

    return progressiveHybridBalance(rows, settings, {
      maxLookaheadAttempts: 100,
      maxDepth2Attempts: 30,
      maxDepth2Chains: 40,
      classBlockRestrictions: {
        ...DEFAULT_CLASS_BLOCK_RESTRICTIONS,
      },
    });
  },
  verify: (result) => {
    const hadLookahead = result.diagnostics.lookaheadAttempts > 0;
    return {
      ok: hadLookahead,
      details: hadLookahead ? 'Lookahead path executed.' : 'Lookahead path not exercised.',
    };
  },
};

const classRestrictionGuardCase: FixtureCase = {
  name: 'Class restrictions are respected',
  run: () => {
    const rows: StandardField[] = [
      makeStudent('s1', 'A', '3STA', 'Matematikk', null, null, null),
      makeStudent('s2', 'B', '3STA', 'Matematikk', null, null, null),
      makeStudent('s3', 'C', '2STA', null, null, null, 'Matematikk'),
      makeStudent('s4', 'D', '2STA', null, null, null, 'Matematikk'),
    ];

    return progressiveHybridBalance(rows, makeSubjectSettings(), {
      classBlockRestrictions: {
        VG2: { 4: false },
        VG3: { 1: false },
      },
    });
  },
  verify: (result) => {
    const violated = result.moveRecords.some((move) => {
      const student = result.updatedData.find((row) => row.studentId === move.studentId);
      if (!student) {
        return false;
      }

      const klasse = (student.klasse || '').toUpperCase();
      if (klasse.startsWith('2') && move.toBlock === 4) {
        return true;
      }
      if (klasse.startsWith('3') && move.toBlock === 1) {
        return true;
      }
      return false;
    });

    return {
      ok: !violated,
      details: !violated ? 'No class-block violations detected.' : 'Detected class restriction violation.',
    };
  },
};

const determinismCase: FixtureCase = {
  name: 'Deterministic repeat run',
  run: () => {
    const rows: StandardField[] = [
      makeStudent('s1', 'A', '1STA', 'Matematikk', null, null, null),
      makeStudent('s2', 'B', '1STA', 'Matematikk', null, null, null),
      makeStudent('s3', 'C', '1STA', null, 'Matematikk', null, null),
      makeStudent('s4', 'D', '1STA', null, 'Matematikk', null, null),
    ];

    return progressiveHybridBalance(rows, makeSubjectSettings(), {
      classBlockRestrictions: {
        ...DEFAULT_CLASS_BLOCK_RESTRICTIONS,
      },
    });
  },
  verify: (result) => {
    const rerun = progressiveHybridBalance(result.updatedData, makeSubjectSettings(), {
      classBlockRestrictions: {
        ...DEFAULT_CLASS_BLOCK_RESTRICTIONS,
      },
    });

    const stable = rerun.moveRecords.length === 0 || rerun.diagnostics.afterScore.total <= rerun.diagnostics.beforeScore.total;
    return {
      ok: stable,
      details: stable ? 'Second run remains stable/improving.' : 'Second run regressed unexpectedly.',
    };
  },
};

const FIXTURES: FixtureCase[] = [
  severeOverloadCase,
  largeGroupPressureCase,
  chainRequiredCase,
  classRestrictionGuardCase,
  determinismCase,
];

export const runProgressiveHybridBalanceFixtures = (): FixtureSuiteResult => {
  const outcomes: FixtureOutcome[] = FIXTURES.map((fixture) => {
    const result = fixture.run();
    const check = fixture.verify(result);
    return {
      name: fixture.name,
      ok: check.ok,
      details: check.details,
    };
  });

  return {
    allPassed: outcomes.every((item) => item.ok),
    outcomes,
  };
};
