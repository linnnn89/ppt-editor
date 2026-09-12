import { readFile, lstat } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import path from 'node:path';

// Heuristic release check, not anonymisation. Never print matched values.
const rules = [
  ['user-directory', /[A-Za-z]:[\\/](?:Users|Documents and Settings)[\\/][^\s'"`\\/]+|\/(?:home|Users)\/[^\s'"`/]+/i],
  ['github-token', /gh[pousr]_[A-Za-z0-9_]{30,}|github_pat_[A-Za-z0-9_]{30,}/],
  ['api-key', /sk-[A-Za-z0-9_-]{24,}/],
  ['aws-key', /AKIA[0-9A-Z]{16}/],
  ['private-key', /-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----/],
  ['private-sample-path', /['"`][^'"`\r\n]*(?:private[\\/]|local-data[\\/]|\u6d4b\u8bd5\u7528PPT[\\/])[^'"`\r\n]*\.pptx['"`]/i],
  ['identifiable-sample-name', /['"][^'"`\r\n]*[\u3400-\u9fff][^'"`\r\n]*\.pptx['"]/]
];
const forbidden = file => {
  const base = path.posix.basename(file);
  return (base !== '.env.example' && /^\.env(?:\..+)?$/i.test(base)) ||
    /\.(?:pem|key|pfx|p12|secret|pptx?|pptm|docx|pdf|png|jpe?g|dmp)$/i.test(base) ||
    /^(?:id_rsa|credentials\.json)$/i.test(base);
};
const git = args => execFileSync('git', args, { encoding: 'utf8', windowsHide: true, maxBuffer: 16 * 1024 * 1024 });
const issues = [];
let scanned = 0;
function scan(file, text, basis) {
  scanned++;
  if (forbidden(file)) issues.push({ file, basis, rule: 'forbidden-file', line: 1 });
  text.split(/\r?\n/).forEach((line, index) => {
    for (const [rule, regex] of rules) if (regex.test(line)) issues.push({ file, basis, rule, line: index + 1 });
  });
}

try {
  // -z retains spaces, Unicode names and exact Git path boundaries.
  const files = new Set(git(['ls-files', '--cached', '--others', '--exclude-standard', '-z']).split('\0').filter(Boolean));
  for (const file of files) {
    try {
      const stat = await lstat(file);
      if (!stat.isFile() || stat.isSymbolicLink()) {
        issues.push({ file, basis: 'worktree', rule: 'non-regular-file', line: 1 });
        continue;
      }
      if (stat.size > 8 * 1024 * 1024) {
        issues.push({ file, basis: 'worktree', rule: 'oversize-needs-review', line: 1 });
        continue;
      }
      scan(file, await readFile(file, 'utf8'), 'worktree');
    } catch (error) { if (error.code !== 'ENOENT') throw error; }
  }
  // Check index blobs too: a clean working copy can hide a staged secret.
  for (const entry of git(['ls-files', '--stage', '-z']).split('\0').filter(Boolean)) {
    const tab = entry.indexOf('\t');
    const [mode, oid, stage] = entry.slice(0, tab).split(' ');
    const file = entry.slice(tab + 1);
    if (stage !== '0' || !['100644', '100755'].includes(mode)) {
      issues.push({ file, basis: 'index', rule: 'unmerged-or-non-regular-file', line: 1 });
      continue;
    }
    scan(file, git(['cat-file', 'blob', oid]), 'index');
  }
  for (const item of issues) console.error(`[${item.basis}] ${item.file}:${item.line} ${item.rule} (value redacted)`);
  if (issues.length) {
    console.error(`Privacy check failed: ${issues.length} rule match(es).`);
    process.exitCode = 1;
  } else {
    console.log(`Privacy check passed: scanned ${files.size} candidate paths / ${scanned} worktree and index contents; no configured rule matches. Manual review is still required.`);
  }
} catch (error) {
  console.error(`Privacy check unavailable (${error.code || 'SCAN_ERROR'}); release check did not pass.`);
  process.exitCode = 1;
}
