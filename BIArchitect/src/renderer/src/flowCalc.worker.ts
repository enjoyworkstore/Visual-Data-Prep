import { computeFlowArtifacts, type WorkerFlowPayload, type WorkerFlowResult } from './lib/flowCalcShared';

type WorkerRequest = {
  requestId: number;
  payload: WorkerFlowPayload;
};

type WorkerResponse = {
  requestId: number;
  result: WorkerFlowResult;
};

self.onmessage = (event: MessageEvent<WorkerRequest>) => {
  const { requestId, payload } = event.data;
  const result = computeFlowArtifacts(payload);
  const response: WorkerResponse = { requestId, result };
  self.postMessage(response);
};
