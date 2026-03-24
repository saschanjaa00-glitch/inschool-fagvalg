/// <reference lib="webworker" />

import type { ForarbeidSubject } from '../components/ForarbeidContext';
import { findOptimalBlockAssignment } from '../utils/forarbeidOptimizer';

interface ForarbeidWorkerRequest {
  subjects: ForarbeidSubject[];
  blokkCount: number;
  chunkStart: number;
  chunkEnd: number;
}

self.onmessage = (event: MessageEvent<ForarbeidWorkerRequest>) => {
  const { subjects, blokkCount, chunkStart, chunkEnd } = event.data;
  // Patch: The worker will run findOptimalBlockAssignment on a chunk of assignments.
  // For now, just run the full calculation (single-threaded fallback)
  // TODO: Implement chunking logic for true parallelism
  const result = findOptimalBlockAssignment(subjects, blokkCount);
  self.postMessage(result);
};

export {};