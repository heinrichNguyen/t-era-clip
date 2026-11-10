import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';

// Heuristic scan only. Never print matched credentials or source lines.
const patterns = [
  ['private-key', /-----BEGIN (?:RSA |EC |OPENSSH |DSA )?PRIVATE KEY-----/],
  ['github-token', /\b(?:gh[pousr]_[A-Za-z0-9]{30,}|github_pat_[A-Za-z0-9_]{50,})\b/],
  ['aws-access-id', /\b(?:AKIA|ASIA)[A-Z0-9]{16}\b/],
  ['openai-key', /\bsk-(?:proj-|svcacct-)?[A-Za-z0-9_-]{40,}\b/],
  ['credential-literal', /(?:api[_-]?key|access[_-]?token|client[_-]?secret|password)\s*[=:]\s*["'][A-Za-z0-9_+\/.=-]{20,}["']/i],
];
const git = (...args) => execFileSync('git', args, { maxBuffer: 32 * 1024 * 1024 });
const findings = [];
let scanned = 0;
let skippedBinary = 0;
const inspect = (buffer, source) => {
  if (buffer.includes(0)) { skippedBinary++; return; }
  scanned++;
  const lines = buffer.toString('utf8').split('\n');
  lines.forEach((line, index) => {
    for (const [kind, pattern] of patterns) {
      if (pattern.test(line)) findings.push({ source, line: index + 1, kind });
    }
  });
};

const files = git('ls-files', '-z', '--cached', '--others', '--exclude-standard')
  .toString('utf8').split('\0').filter(Boolean);
for (const file of new Set(files)) {
  if (existsSync(file)) inspect(readFileSync(file), `worktree:${file}`);
}
const walk = (path) => {
  for (const entry of readdirSync(path, { withFileTypes: true })) {
    const file = join(path, entry.name);
    if (entry.isDirectory()) walk(file);
    else if (entry.isFile()) inspect(readFileSync(file), `bundle:${file}`);
  }
};
if (existsSync('dist')) walk('dist');

const objects = git('rev-list', '--objects', '--all').toString('utf8').trim().split('\n');
const checked = new Set();
let historyBlobs = 0;
for (const entry of objects) {
  const [id, ...path] = entry.split(' ');
  if (!id || checked.has(id)) continue;
  checked.add(id);
  if (git('cat-file', '-t', id).toString('utf8').trim() !== 'blob') continue;
  historyBlobs++;
  inspect(git('cat-file', 'blob', id), `git:${id.slice(0, 12)}:${path.join(' ')}`);
}
console.log(JSON.stringify({
  scan: 'heuristic-credential-patterns-not-a-complete-secret-audit',
  commits: Number(git('rev-list', '--count', '--all').toString('utf8').trim()),
  historyBlobs, scannedTextFilesAndBlobs: scanned, skippedBinary,
  findings,
  limits: 'Reachable local Git refs, non-ignored worktree files, and dist only. No reflog, remote-only history, ignored files outside dist, entropy scan, or credential validity checks.',
}, null, 2));
process.exitCode = findings.length ? 1 : 0;
