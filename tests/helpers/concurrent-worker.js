import { FileLock, TaskHost } from "../../src/task.js";

let heldLock;

process.on("message", async msg => {
  if (msg.action === "init") {
    process.send({ ready: true });
    return;
  }
  if (msg.action === "hold-lock") {
    heldLock = await FileLock.acquire(msg.lockPath);
    process.send({ locked: true, nonce: heldLock.nonce });
    return;
  }
  if (msg.action === "apply") {
    try {
      const host = new TaskHost({ taskId: msg.taskId, baseDir: msg.baseDir });
      const result = await host.apply({
        documentId: msg.documentId,
        expectedRevision: msg.expectedRevision,
        operationId: msg.operationId,
        operations: msg.operations
      });
      process.send({ success: true, result });
    } catch (error) {
      process.send({
        success: false,
        error: { code: error.code || "UNKNOWN", message: error.message, details: error.details }
      });
    }
  }
});
