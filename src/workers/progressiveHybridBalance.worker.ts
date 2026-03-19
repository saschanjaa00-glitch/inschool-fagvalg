/// <reference lib="webworker" />

import { progressiveHybridBalance, AbortBalancingError } from '../utils/progressiveHybridBalance';
import type { BalancingWorkerInbound, BalancingWorkerOutbound } from './progressiveHybridBalance.worker.types';

let abortRequested = false;

self.onmessage = (event: MessageEvent<BalancingWorkerInbound>) => {
  const message = event.data;

  if (message && message.type === 'stop') {
    abortRequested = true;
    return;
  }

  if (!message || message.type !== 'run') {
    return;
  }

  abortRequested = false;
  const { requestId, payload } = message;

  try {
    const progressCallback = (progress: import('../utils/progressiveHybridBalance').BalancingProgress) => {
      if (abortRequested) {
        throw new AbortBalancingError();
      }
      const progressResponse: BalancingWorkerOutbound = { type: 'progress', requestId, progress };
      self.postMessage(progressResponse);
    };
    const result = progressiveHybridBalance(payload.rows, payload.subjectSettingsByName, payload.config, progressCallback);
    const response: BalancingWorkerOutbound = {
      type: 'success',
      requestId,
      result,
    };

    self.postMessage(response);
  } catch (error) {
    if (error instanceof AbortBalancingError) {
      const response: BalancingWorkerOutbound = {
        type: 'error',
        requestId,
        message: 'Balansering avbrutt av bruker',
      };
      self.postMessage(response);
      return;
    }

    const response: BalancingWorkerOutbound = {
      type: 'error',
      requestId,
      message: error instanceof Error ? error.message : 'Ukjent feil i balanseringsarbeider',
    };

    self.postMessage(response);
  }
};

export {};
