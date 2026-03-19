import type { StandardField } from '../utils/excelUtils';
import type {
  BalancingConfig,
  BalancingProgress,
  ProgressiveHybridBalanceResult,
  SubjectSettingsByNameLike,
} from '../utils/progressiveHybridBalance';

export type { BalancingProgress };

export interface BalancingWorkerRunPayload {
  rows: StandardField[];
  subjectSettingsByName: SubjectSettingsByNameLike;
  config: Partial<BalancingConfig>;
}

export interface BalancingWorkerRunRequest {
  type: 'run';
  requestId: number;
  payload: BalancingWorkerRunPayload;
}

export interface BalancingWorkerSuccessResponse {
  type: 'success';
  requestId: number;
  result: ProgressiveHybridBalanceResult;
}

export interface BalancingWorkerErrorResponse {
  type: 'error';
  requestId: number;
  message: string;
}

export interface BalancingWorkerProgressResponse {
  type: 'progress';
  requestId: number;
  progress: BalancingProgress;
}

export type BalancingWorkerInbound = BalancingWorkerRunRequest;
export type BalancingWorkerOutbound = BalancingWorkerSuccessResponse | BalancingWorkerErrorResponse | BalancingWorkerProgressResponse;
