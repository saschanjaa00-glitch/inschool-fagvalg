import type { ForarbeidSubject } from '../components/ForarbeidContext';
import type { ForarbeidResult } from './forarbeidOptimizer';

const WORKER_COUNT = 8;

export function runForarbeidParallel(
  subjects: ForarbeidSubject[],
  blokkCount: number
): Promise<ForarbeidResult> {
  return new Promise((resolve, reject) => {
    const workers: Worker[] = [];
    let finished = 0;
    let bestResult: ForarbeidResult | null = null;
    let errorOccurred = false;

    // Dynamically import the worker script
    const workerUrl = new URL('../workers/forarbeidOptimizer.worker.ts', import.meta.url,).href;

    for (let i = 0; i < WORKER_COUNT; i++) {
      const worker = new Worker(workerUrl, { type: 'module' });
      workers.push(worker);
    }

    // For now, just split by first subject's allowed blocks (for demo, not optimal)
    const allowedBlocks = Array.from({ length: blokkCount }, (_, i) => i + 1);
    const chunkSize = Math.ceil(allowedBlocks.length / WORKER_COUNT);
    const _chunks = Array.from({ length: WORKER_COUNT }, (_, i) =>
      allowedBlocks.slice(i * chunkSize, (i + 1) * chunkSize)
    );

    workers.forEach((worker, _idx) => {
      worker.onmessage = (event: MessageEvent) => {
        if (errorOccurred) return;
        const result: ForarbeidResult = event.data;
        if (!bestResult || result.maxCombinations > bestResult.maxCombinations) {
          bestResult = result;
        }
        finished++;
        if (finished === WORKER_COUNT) {
          workers.forEach(w => w.terminate());
          if (bestResult) resolve(bestResult);
          else reject(new Error('No result'));
        }
      };
      worker.onerror = (err) => {
        errorOccurred = true;
        workers.forEach(w => w.terminate());
        reject(err);
      };
      // For now, all workers get the same data (real chunking would split assignments)
      worker.postMessage({ subjects, blokkCount, chunkStart: 0, chunkEnd: 0 });
    });
  });
}
