import koffi from 'koffi';
import { setTimeout as delay } from 'node:timers/promises';
import { check, PptError } from './errors.js';

check(process.platform === 'win32' && process.arch === 'x64', 'PLATFORM_UNSUPPORTED', 'This build requires Windows x64.');
const kernel = koffi.load('kernel32.dll');
const advapi = koffi.load('advapi32.dll');
const user = koffi.load('user32.dll');
const bind = (library, signature) => library.func(signature);
const OpenProcess = bind(kernel, 'void * __stdcall OpenProcess(uint32_t, int, uint32_t)');
const CloseHandle = bind(kernel, 'int __stdcall CloseHandle(void *)');
const GetLastError = bind(kernel, 'uint32_t __stdcall GetLastError()');
const GetProcessTimes = bind(kernel, 'int __stdcall GetProcessTimes(void *, void *, void *, void *, void *)');
const QueryImage = bind(kernel, 'int __stdcall QueryFullProcessImageNameW(void *, uint32_t, void *, _Inout_ uint32_t *)');
const Wait = bind(kernel, 'uint32_t __stdcall WaitForSingleObject(void *, uint32_t)');
const TerminateProcess = bind(kernel, 'int __stdcall TerminateProcess(void *, uint32_t)');
const CreateJob = bind(kernel, 'void * __stdcall CreateJobObjectW(void *, str16)');
const SetJob = bind(kernel, 'int __stdcall SetInformationJobObject(void *, int, void *, uint32_t)');
const AssignJob = bind(kernel, 'int __stdcall AssignProcessToJobObject(void *, void *)');
const InJob = bind(kernel, 'int __stdcall IsProcessInJob(void *, void *, _Out_ int *)');
const CreateSnapshot = bind(kernel, 'void * __stdcall CreateToolhelp32Snapshot(uint32_t, uint32_t)');
const ProcessFirst = bind(kernel, 'int __stdcall Process32FirstW(void *, void *)');
const ProcessNext = bind(kernel, 'int __stdcall Process32NextW(void *, void *)');
const OpenToken = bind(advapi, 'int __stdcall OpenProcessToken(void *, uint32_t, _Out_ void **)');
const TokenInformation = bind(advapi, 'int __stdcall GetTokenInformation(void *, int, void *, uint32_t, _Out_ uint32_t *)');
const SidToString = bind(advapi, 'int __stdcall ConvertSidToStringSidW(void *, _Out_ void **)');
const LocalFree = bind(kernel, 'void * __stdcall LocalFree(void *)');
const ConvertDescriptor = bind(advapi, 'int __stdcall ConvertStringSecurityDescriptorToSecurityDescriptorW(str16, uint32_t, _Out_ void **, void *)');
const GetDacl = bind(advapi, 'int __stdcall GetSecurityDescriptorDacl(void *, _Out_ int *, _Out_ void **, _Out_ int *)');
const SetNamedSecurity = bind(advapi, 'uint32_t __stdcall SetNamedSecurityInfoW(str16, int, uint32_t, void *, void *, void *, void *)');
const GetForeground = bind(user, 'void * __stdcall GetForegroundWindow()');
const WindowPid = bind(user, 'uint32_t __stdcall GetWindowThreadProcessId(void *, _Out_ uint32_t *)');
const ClipboardSequence = bind(user, 'uint32_t __stdcall GetClipboardSequenceNumber()');
const PeekMessage = bind(user, 'int __stdcall PeekMessageW(void *, void *, uint32_t, uint32_t, uint32_t)');
const TranslateMessage = bind(user, 'int __stdcall TranslateMessage(void *)');
const DispatchMessage = bind(user, 'intptr_t __stdcall DispatchMessageW(void *)');
const EventCallback = koffi.proto('void __stdcall PptForegroundCallback(void *, uint32_t, void *, int32_t, int32_t, uint32_t, uint32_t)');
const SetEventHook = user.func('SetWinEventHook', 'void *', ['uint32_t', 'uint32_t', 'void *', koffi.pointer(EventCallback), 'uint32_t', 'uint32_t', 'uint32_t']);
const Unhook = bind(user, 'int __stdcall UnhookWinEvent(void *)');
const CreateMutex = bind(kernel, 'void * __stdcall CreateMutexW(void *, int, str16)');
const ReleaseMutex = bind(kernel, 'int __stdcall ReleaseMutex(void *)');

function winCheck(result, operation) {
  check(result, 'WINDOWS_API_ERROR', `${operation} failed.`, { win32Error: GetLastError() });
}

function identityFromHandle(handle, pid) {
  const created = Buffer.alloc(8), exited = Buffer.alloc(8), kernelTime = Buffer.alloc(8), userTime = Buffer.alloc(8);
  winCheck(GetProcessTimes(handle, created, exited, kernelTime, userTime), 'GetProcessTimes');
  const image = Buffer.alloc(65536), length = [32768];
  winCheck(QueryImage(handle, 0, image, length), 'QueryFullProcessImageNameW');
  return { pid, created: created.readBigUInt64LE().toString(), image: image.toString('utf16le', 0, length[0] * 2) };
}

export function processIdentity(pid) {
  const handle = OpenProcess(0x101000, 0, pid);
  if (!handle) {
    const code = GetLastError();
    if (code === 87) return null;
    throw new PptError('PROCESS_ACCESS_ERROR', 'Unable to observe process identity.', { pid, win32Error: code });
  }
  try { if (Wait(handle, 0) === 0) return null; return identityFromHandle(handle, pid); }
  finally { CloseHandle(handle); }
}

export function sameProcess(expected) {
  if (!expected) return false;
  const current = processIdentity(expected.pid);
  return Boolean(current && current.created === expected.created && current.image.toLowerCase() === expected.image.toLowerCase());
}

export async function waitForExit(identity, timeoutMs = 10000) {
  const handle = OpenProcess(0x101000, 0, identity.pid);
  if (!handle) { if (GetLastError() === 87) return true; throw new PptError('PROCESS_ACCESS_ERROR', 'Unable to verify process exit.'); }
  try {
    if (Wait(handle, 0) === 0) return true;
    try {
      const created = Buffer.alloc(8), exited = Buffer.alloc(8), kernelTime = Buffer.alloc(8), userTime = Buffer.alloc(8);
      if (GetProcessTimes(handle, created, exited, kernelTime, userTime)) {
        if (created.readBigUInt64LE().toString() !== identity.created) return true;
      }
    } catch {}
    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
      const state = Wait(handle, 0);
      if (state === 0) return true;
      check(state === 258, 'PROCESS_WAIT_ERROR', 'Unable to wait for process exit.');
      await delay(40);
    }
    return Wait(handle, 0) === 0;
  } finally { CloseHandle(handle); }
}

export function terminateOwnedProcess(identity) {
  const handle = OpenProcess(0x101001, 0, identity.pid);
  if (!handle) { check(GetLastError() === 87, 'PROCESS_ACCESS_ERROR', 'Cannot terminate owned worker.'); return; }
  try {
    check(identityFromHandle(handle, identity.pid).created === identity.created, 'PROCESS_IDENTITY_CHANGED', 'Refusing to terminate a reused PID.');
    winCheck(TerminateProcess(handle, 1), 'TerminateProcess');
  } finally { CloseHandle(handle); }
}

export function listProcesses(name) {
  const snapshot = CreateSnapshot(2, 0);
  check(snapshot && BigInt(snapshot) !== 0xffffffffffffffffn, 'PROCESS_ENUMERATION_ERROR', 'Cannot enumerate processes.');
  const entry = Buffer.alloc(568);
  entry.writeUInt32LE(entry.length, 0);
  const results = [];
  try {
    let ok = ProcessFirst(snapshot, entry);
    while (ok) {
      const executable = entry.toString('utf16le', 44, 564).split('\0')[0];
      if (!name || executable.toLowerCase() === name.toLowerCase()) results.push({ pid: entry.readUInt32LE(8), parentPid: entry.readUInt32LE(32), executable });
      ok = ProcessNext(snapshot, entry);
    }
    check(GetLastError() === 18, 'PROCESS_ENUMERATION_ERROR', 'Process enumeration did not complete; an incomplete list cannot establish ownership.');
  } finally { CloseHandle(snapshot); }
  return results;
}

export function createOwnedJob() {
  const handle = CreateJob(null, null);
  winCheck(handle, 'CreateJobObjectW');
  const limits = Buffer.alloc(144);
  // Silent breakaway prevents Office/COM descendants entering the kill scope.
  // Every owned Node child is explicitly assigned before it starts work.
  limits.writeUInt32LE(0x2000 | 0x1000, 16);
  if (!SetJob(handle, 9, limits, limits.length)) { CloseHandle(handle); throw new PptError('JOB_SETUP_FAILED', 'Cannot configure owned-worker cleanup.'); }
  let closed = false;
  return {
    add(identity) {
      check(!closed, 'JOB_CLOSED', 'Worker job is closed.');
      const child = OpenProcess(0x100 | 0x1 | 0x1000, 0, identity.pid);
      winCheck(child, 'OpenProcess for job assignment');
      try {
        check(identityFromHandle(child, identity.pid).created === identity.created, 'PROCESS_IDENTITY_CHANGED', 'Worker PID changed before assignment.');
        winCheck(AssignJob(handle, child), 'AssignProcessToJobObject');
      } finally { CloseHandle(child); }
    },
    contains(pid) {
      const child = OpenProcess(0x1000, 0, pid); winCheck(child, 'OpenProcess for job verification');
      try { const result = [0]; winCheck(InJob(child, handle, result), 'IsProcessInJob'); return Boolean(result[0]); }
      finally { CloseHandle(child); }
    },
    close() { if (!closed) { closed = true; CloseHandle(handle); } }
  };
}

let cachedSid;
export function currentUserSid() {
  if (cachedSid) return cachedSid;
  const current = OpenProcess(0x1000, 0, process.pid), token = [null];
  winCheck(current, 'OpenProcess');
  try { winCheck(OpenToken(current, 8, token), 'OpenProcessToken'); } finally { CloseHandle(current); }
  try {
    const length = [0]; TokenInformation(token[0], 1, null, 0, length);
    const buffer = Buffer.alloc(length[0]);
    winCheck(TokenInformation(token[0], 1, buffer, buffer.length, length), 'GetTokenInformation');
    const sid = koffi.decode(buffer, 'void *'), text = [null];
    winCheck(SidToString(sid, text), 'ConvertSidToStringSidW');
    try { cachedSid = koffi.decode.string16(text[0]); return cachedSid; } finally { LocalFree(text[0]); }
  } finally { CloseHandle(token[0]); }
}

export function restrictAccess(target, { directory = false } = {}) {
  const inheritance = directory ? 'OICI' : '';
  const descriptor = [null], present = [0], dacl = [null], defaulted = [0];
  const sddl = `D:P(A;${inheritance};GA;;;SY)(A;${inheritance};GA;;;${currentUserSid()})`;
  winCheck(ConvertDescriptor(sddl, 1, descriptor, null), 'ConvertStringSecurityDescriptor');
  try {
    winCheck(GetDacl(descriptor[0], present, dacl, defaulted), 'GetSecurityDescriptorDacl');
    const status = SetNamedSecurity(target, 1, 0x80000004, null, null, dacl[0], null);
    check(status === 0, 'ACL_SETUP_FAILED', 'Cannot restrict task resources to the current user.', { win32Error: status });
  } finally { LocalFree(descriptor[0]); }
}

export function windowProcessId(hwnd) {
  const result = [0]; WindowPid(typeof hwnd === 'number' ? BigInt(hwnd >>> 0) : hwnd, result); return result[0];
}

export function pumpMessages() {
  const message = Buffer.alloc(48);
  let count = 0;
  while (count++ < 256 && PeekMessage(message, null, 0, 0, 1)) { TranslateMessage(message); DispatchMessage(message); }
}

export function observeForeground() {
  const events = [], initialClipboard = ClipboardSequence();
  const initial = { hwnd: String(GetForeground()), pid: windowProcessId(GetForeground()), at: Date.now() };
  const callback = koffi.register((_hook, event, hwnd, _object, _child, thread, time) => {
    if (events.length < 10000) events.push({ event, hwnd: String(hwnd), pid: windowProcessId(hwnd), thread, time, at: Date.now() });
  }, koffi.pointer(EventCallback));
  const hook = SetEventHook(3, 3, null, callback, 0, 0, 0);
  if (!hook) { koffi.unregister(callback); throw new PptError('FOREGROUND_OBSERVER_FAILED', 'Cannot observe foreground events.'); }
  const timer = setInterval(pumpMessages, 10);
  return {
    stop() {
      pumpMessages(); clearInterval(timer); Unhook(hook); koffi.unregister(callback);
      return { initial, events, truncated: events.length === 10000, final: { hwnd: String(GetForeground()), pid: windowProcessId(GetForeground()) }, clipboardSequenceUnchanged: initialClipboard === ClipboardSequence() };
    }
  };
}

export function initializeSta() {
  const ole = koffi.load('ole32.dll');
  const initialize = ole.func('int32_t __stdcall CoInitializeEx(void *, uint32_t)');
  const uninitialize = ole.func('void __stdcall CoUninitialize()');
  const result = initialize(null, 2);
  check(result === 0 || result === 1, 'COM_APARTMENT_ERROR', 'Native executor could not enter STA.', { hresult: (result >>> 0).toString(16) });
  return () => uninitialize();
}

export function acquireNativeMutex() {
  const mutex = CreateMutex(null, 0, `Local\\ppt-editor-native-${currentUserSid()}`);
  winCheck(mutex, 'CreateMutexW');
  const result = Wait(mutex, 0);
  if (result !== 0 && result !== 0x80) { CloseHandle(mutex); throw new PptError('NATIVE_BUSY', 'Another ppt-editor task holds the native editing lease.'); }
  let closed = false;
  return () => { if (!closed) { closed = true; ReleaseMutex(mutex); CloseHandle(mutex); } };
}
