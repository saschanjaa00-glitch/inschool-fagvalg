// Utility for Forarbeid: Find subject-to-block assignments maximizing unique 3-subject combinations
// Each subject has groupCount, VG-level, and blokk restrictions
// Returns: { assignment: Record<subjectId, blockNumber>, maxCombinations: number, details: ... }

import type { ForarbeidSubject } from '../components/ForarbeidContext';

export interface ForarbeidResult {
  assignment: Record<string, number>;
  maxCombinations: number;
  details: Array<{ blocks: number[]; count: number }>;
}

// Generate all possible assignments of subjects to blocks, respecting restrictions
function* generateAssignments(subjects: ForarbeidSubject[], blokkCount: number): Generator<Record<string, number>> {
  if (subjects.length === 0) {
    yield {};
    return;
  }
  const [first, ...rest] = subjects;
  const allowedBlocks = Array.from({ length: blokkCount }, (_, i) => i + 1).filter(
    b => !first.blokkRestrictions.includes(b) && !(first.vgLevel === 'VG3' && b === 1)
  );
  for (const block of allowedBlocks) {
    for (const subAssignment of generateAssignments(rest, blokkCount)) {
      yield { [first.id]: block, ...subAssignment };
    }
  }
}

// Count unique 3-subject combinations possible for students
function countCombinations(assignment: Record<string, number>): { count: number; details: Array<{ blocks: number[]; count: number }> } {
  // Group subjects by block
  const byBlock: Record<number, string[]> = {};
  for (const [id, block] of Object.entries(assignment)) {
    if (!byBlock[block]) byBlock[block] = [];
    byBlock[block].push(id);
  }
  // For each possible pick of 3 blocks, count combinations
  const blocks = Object.keys(byBlock).map(Number);
  let total = 0;
  const details: Array<{ blocks: number[]; count: number }> = [];
  for (let i = 0; i < blocks.length; ++i) {
    for (let j = i + 1; j < blocks.length; ++j) {
      for (let k = j + 1; k < blocks.length; ++k) {
        const b1 = blocks[i], b2 = blocks[j], b3 = blocks[k];
        const c = (byBlock[b1]?.length || 0) * (byBlock[b2]?.length || 0) * (byBlock[b3]?.length || 0);
        if (c > 0) details.push({ blocks: [b1, b2, b3], count: c });
        total += c;
      }
    }
  }
  return { count: total, details };
}

export function findOptimalBlockAssignment(
  subjects: ForarbeidSubject[],
  blokkCount: number,
  onProgress?: (progress: { current: number; total: number }) => void
): ForarbeidResult {
  let maxCombinations = 0;
  let bestAssignment: Record<string, number> = {};
  let bestDetails: Array<{ blocks: number[]; count: number }> = [];
  // Estimate total assignments
  let total = 1;
  for (const s of subjects) {
    const allowed = Array.from({ length: blokkCount }, (_, i) => i + 1).filter(
      b => !s.blokkRestrictions.includes(b) && !(s.vgLevel === 'VG3' && b === 1)
    );
    total *= allowed.length;
  }
  let current = 0;
  for (const assignment of generateAssignments(subjects, blokkCount)) {
    current++;
    if (onProgress && (current % 1000 === 0 || current === total)) {
      onProgress({ current, total });
    }
    const { count, details } = countCombinations(assignment);
    if (count > maxCombinations) {
      maxCombinations = count;
      bestAssignment = assignment;
      bestDetails = details;
    }
  }
  if (onProgress) onProgress({ current: total, total });
  return { assignment: bestAssignment, maxCombinations, details: bestDetails };
}
