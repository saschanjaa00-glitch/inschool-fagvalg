# Progressive Hybrid Balance Architecture

## Pseudocode

```text
progressiveHybridBalance(rows, subjectSettings, config):
  state = buildState(rows, subjectSettings, config)
  before = computeScore(state)

  for offset in [-maxRelaxation .. 0]:
    repeat:
      flowNetwork = buildFlowNetwork(state, offset)
      solved = solveFlow(flowNetwork)
      moves = extractMovesFromFlow(solved)
      applied = applyMoves(state, moves, offset)
    until applied is empty or budget exhausted

    localSearchImprove(state)

  repairCollisions(state)

  after = computeScore(state)
  return {
    updatedData,
    moveRecords,
    diagnostics(before, after)
  }
```

## Module Responsibilities

- buildState:
  - Parse student block assignments.
  - Build subject-group occupancy and capacities.
  - Include empty configured groups.
  - Respect locked assignment keys.
- computeScore:
  - Uses weighted penalties for overcap, imbalance, peak, collision, moves, repeat moves.
- buildFlowNetwork:
  - Build deterministic candidate moves from current assignment state.
  - Estimate delta using snapshot simulation.
- solveFlow:
  - Deterministic greedy min-cost-flow approximation for UI-time constraints.
- extractMovesFromFlow:
  - Stable ordered candidate extraction.
- applyMoves:
  - Constraint-checked move commit with strict/relaxed capacity by offset.
- localSearchImprove:
  - Single-move hill climbing + lookahead chain trigger.
- tryLookaheadChain:
  - Snapshot/rollback, depth-1 and depth-2 transactional chain attempts.
- repairCollisions:
  - Capacity-respecting collision repair first, least-bad fallback second.
- progressiveHybridBalance:
  - Orchestration with progressive capacity schedule and diagnostics.

## Complexity Notes

Let:
- S = number of students
- A = total assignments across all students
- G = total groups
- C = generated candidate moves

Approximate runtime per pass:
- Candidate generation: O(C)
- Delta estimation by simulation: O(C * (A + G)) in current implementation
- Greedy flow selection: O(C log C)
- Local search: bounded by configured attempt caps

Total runtime is bounded in practice by:
- maxPassMillis
- maxLookaheadAttempts
- maxDepth2Chains

## Tuning Guide

- overcapA:
  - Increase when overfilled groups must be solved first.
- imbalanceB:
  - Increase for flatter within-subject distribution.
- peakC:
  - Increase to aggressively reduce largest groups.
- alpha:
  - Increase when large groups should dominate balancing urgency.
- movesE:
  - Increase to reduce total movement.
- repeatF:
  - Increase to avoid moving same student multiple times.
- collisionD:
  - Keep very high to strongly discourage collisions.

Recommended defaults are defined in:
- progressiveHybridBalance.ts
