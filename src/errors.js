export class PptError extends Error {
  constructor(code, message, details = {}) {
    super(message);
    this.name = 'PptError';
    this.code = code;
    this.details = details;
  }
}

export function check(condition, code, message, details) {
  if (!condition) throw new PptError(code, message, details);
}

export function errorResult(error) {
  return { code: error.code || 'INTERNAL_ERROR', message: error.message, details: error.details || {} };
}
