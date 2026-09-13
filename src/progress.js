import { AsyncLocalStorage } from 'node:async_hooks';

const scope = new AsyncLocalStorage();
export const currentProgress = () => scope.getStore();
export const reportProgress = (stage, details = {}) => currentProgress()?.({ stage, ...details });

// Request-local callbacks also let native IPC capture the correct recipient
// without a mutable, process-wide "current tool" or changes to operation receipts.
export function createRequestProgress(request, tool, receivedAt) {
  const startedAt = performance.now(), durations = new Map();
  const token = request._meta?.progressToken;
  const enabled = typeof token === 'string' || Number.isInteger(token);
  let stage, stageStartedAt = startedAt, accepting = true, sequence = 0;
  let lastSentAt = -Infinity, lastKey, sent = Promise.resolve(), notificationErrors = 0;
  const recordDuration = now => {
    if (stage) durations.set(stage, (durations.get(stage) || 0) + now - stageStartedAt);
  };
  const report = event => {
    if (!accepting) return;
    const now = performance.now(), changed = event.stage !== stage;
    if (changed) { recordDuration(now); stage = event.stage; stageStartedAt = now; }
    if (!enabled || request.signal.aborted) return;
    const { completed, total, unit } = event;
    const counted = Number.isInteger(completed) && Number.isInteger(total) && total > 0;
    const key = JSON.stringify([stage, completed, total, unit]);
    if (key === lastKey || (!changed && !(counted && completed === total) && now - lastSentAt < 250)) return;
    lastKey = key; lastSentAt = now;
    const count = counted ? ` (${completed}/${total} ${unit || 'items'})` : '';
    const message = `${tool}: ${stage.replaceAll('-', ' ')}${count}; elapsed ${((now - startedAt) / 1000).toFixed(1)}s; stage ${((now - stageStartedAt) / 1000).toFixed(1)}s`;
    // The whole request has no known step total. Progress is a monotonic event
    // sequence; actual operation/page counts belong to the human-readable message.
    const params = { progressToken: token, progress: ++sequence, message };
    sent = sent.then(() => {
      if (!request.signal.aborted) return request.notify({ method: 'notifications/progress', params });
    }).catch(() => { notificationErrors++; });
  };
  return {
    run: execute => scope.run(report, () => { report({ stage: 'started' }); return execute(); }),
    async finish(outcome) {
      report({ stage: outcome }); accepting = false;
      const endedAt = performance.now(); recordDuration(endedAt);
      await sent;
      return { queueMs: Math.max(0, Math.round(startedAt - receivedAt)), elapsedMs: Math.round(endedAt - startedAt),
        stages: [...durations].map(([name, elapsed]) => ({ stage: name, elapsedMs: Math.round(elapsed) })),
        ...(notificationErrors ? { notificationErrors } : {}) };
    }
  };
}
