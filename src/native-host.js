import { fork } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { randomUUID } from 'node:crypto';
import { once } from 'node:events';
import { PptError, check } from './errors.js';
import { processIdentity, sameProcess, waitForExit, terminateOwnedProcess } from './windows.js';
import { currentProgress, reportProgress } from './progress.js';

export function assertCleanupConfirmed(report) {
  const known = report?.office === 'not_started' || ['completed', 'application_preserved'].includes(report?.cleanup);
  check(known && report.workerExited === true && report.outcome !== 'outcome_unknown' &&
    !report.errors?.length && (!report.lease?.quitRequested || report.officeExited === true),
  'CLEANUP_UNCONFIRMED', 'Native resource cleanup could not be confirmed. Recovery data has been retained.', { report });
}

export class NativeHost {
  constructor(job, onState = async () => {}) { this.job = job; this.onState = onState; this.pending = new Map(); this.broken = false; this.ownerKey = randomUUID(); }

  async start() {
    if (this.worker) { check(!this.broken, 'NATIVE_HOST_UNAVAILABLE', 'Native executor failed; resume from a checkpoint.'); return; }
    reportProgress('starting-native-worker');
    this.worker = fork(fileURLToPath(new URL('./native-worker.js', import.meta.url)), [], { stdio: ['ignore', 'ignore', 'pipe', 'ipc'], windowsHide: true, execArgv: ['--expose-gc'] });
    this.worker.stderr.on('data', chunk => { this.lastError = ((this.lastError || '') + chunk).slice(-4000); });
    this.worker.on('message', message => {
      if (message.stage) {
        this.lastStage = message.stage; if (message.lease) this.lastLease = message.lease;
        const { stage, completed, total, unit } = message;
        this.pending.get(message.id)?.progress?.({ stage, completed, total, unit });
        return;
      }
      const pending = this.pending.get(message.id); if (!pending) return;
      clearTimeout(pending.timer); this.pending.delete(message.id);
      if (message.error) pending.reject(new PptError(message.error.code, message.error.message, message.error.details));
      else pending.resolve(message.result);
    });
    this.worker.on('error', error => this.fail(error));
    this.worker.on('exit', (code, signal) => {
      this.exited = true;
      for (const pending of this.pending.values()) { clearTimeout(pending.timer); pending.reject(new PptError('NATIVE_EXITED', 'Native executor exited before returning a result.', { code, signal, stage: this.lastStage, diagnostics: this.lastError })); }
      this.pending.clear();
    });
    await once(this.worker, 'spawn');
    this.identity = processIdentity(this.worker.pid);
    try { this.job.add(this.identity); } catch (error) { terminateOwnedProcess(this.identity); await waitForExit(this.identity); throw error; }
    await this.onState({ nativeWorker: this.identity });
    await this.request('initialize', {}, 15000);
  }

  fail(error) {
    this.broken = true;
    for (const pending of this.pending.values()) { clearTimeout(pending.timer); pending.reject(error); } this.pending.clear();
  }

  async request(method, args = {}, timeoutMs = 60000) {
    check(this.worker && !this.exited && !this.broken, 'NATIVE_HOST_UNAVAILABLE', 'Native executor is not available.');
    check(sameProcess(this.identity), 'NATIVE_PROCESS_MISMATCH', 'The task worker process identity no longer matches.');
    const id = randomUUID();
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id); this.broken = true;
        try { terminateOwnedProcess(this.identity); } catch {}
        reject(new PptError('NATIVE_TIMEOUT', 'Native request timed out; outcome may be unknown. Resume from the last durable checkpoint.', { method, worker: this.identity, outcome: 'outcome_unknown' }));
      }, timeoutMs);
      this.pending.set(id, { resolve, reject, timer, progress: currentProgress() });
      this.worker.send({ id, method, args, ownerKey: this.ownerKey }, error => { if (error) { clearTimeout(timer); this.pending.delete(id); reject(error); } });
    });
  }

  async shutdown(interrupted = false) {
    if (!this.worker) return { workerExited: true, office: 'not_started' };
    let result, shutdownError;
    if (!this.exited && !this.broken) {
      try { result = await this.request('shutdown', { interrupted }, 30000); }
      catch (error) { shutdownError = { code: error.code || 'NATIVE_SHUTDOWN_FAILED', message: error.message }; }
    }
    reportProgress('confirming-native-exit');
    const workerExited = await waitForExit(this.identity, 10000);
    let officeExited = null;
    const observedOffice = [];
    const lease = result?.lease || this.lastLease;
    const observations = lease?.observedNewProcesses || [];
    if (result?.lease?.quitRequested) officeExited = await waitForExit(result.lease.identity, 10000);
    for (const identity of observations) {
      const isQuitTarget = result?.lease?.quitRequested && identity.pid === result.lease.identity.pid && identity.created === result.lease.identity.created;
      observedOffice.push({ identity, exited: isQuitTarget ? officeExited : await waitForExit(identity, 0) });
    }
    if (officeExited === null && observedOffice.length) officeExited = observedOffice.every(p => p.exited);
    const cleanup = !result ? 'worker_not_responsive' : !workerExited ? 'worker_exit_unconfirmed' : result.errors?.length ? 'document_cleanup_failed' : result.lease?.quitRequested ? officeExited ? 'completed' : 'office_exit_unconfirmed' : result.lease ? 'application_preserved' : 'completed';
    return { ...result, lease, worker: this.identity, workerExited, officeExited, observedOffice, cleanup,
      ...(!result ? { outcome: 'outcome_unknown' } : {}), ...(shutdownError ? { errors: [shutdownError] } : {}) };
  }
}
