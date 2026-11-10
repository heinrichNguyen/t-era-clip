import { createHash } from 'node:crypto';
import { cpSync, existsSync, lstatSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const dist = join(root, 'dist-discussions');
const manifest = JSON.parse(readFileSync(join(dist, 'manifest.json'), 'utf8'));
if (manifest.background || manifest.options_page || manifest.host_permissions || manifest.content_scripts ||
  JSON.stringify(manifest.permissions) !== JSON.stringify(['storage', 'activeTab', 'scripting']) || manifest.action.default_popup !== 'discussion-popup.html') throw new Error('Unexpected discussion package permissions or entries.');
const files = [];
function walk(dir) { for (const entry of readdirSync(dir).sort()) { const path = join(dir, entry); const info = lstatSync(path);
  if (info.isSymbolicLink()) throw new Error('Symlink in package');
  if (info.isDirectory()) walk(path); else { const name = relative(dist, path); if (!/\.(?:json|html|js|css|png)$/.test(name) || /demo|import-review|no-review|access-review|options|background/.test(name)) throw new Error(`Unexpected file: ${name}`); files.push(name); }
} }
walk(dist);
if (!['discussion-popup.html', 'review.html', 'project-review.html', 'tadt-logo.png'].every((file) => files.includes(file))) throw new Error('Missing required resource');
const stamp = new Date().toISOString().replace(/[:.]/g, '-');
const out = join(root, 'releases', `discussions-only-${stamp}`);
if (existsSync(out)) throw new Error('Release exists'); mkdirSync(out, { recursive: true });
cpSync(dist, join(out, 'T-eraClip-Thao-luan'), { recursive: true });
const zip = join(out, 't-eraclip-discussions-beta-store.zip');
execFileSync('zip', ['-X', '-q', zip, ...files], { cwd: dist }); execFileSync('unzip', ['-t', zip], { stdio: 'pipe' });
writeFileSync(join(out, 'BUILD-INFO.json'), JSON.stringify({ version: manifest.version, profile: 'discussions-only', packagedAt: new Date().toISOString(),
  revision: execFileSync('git', ['rev-parse', 'HEAD'], { cwd: root, encoding: 'utf8' }).trim(), dirtyWorkingTree: !!execFileSync('git', ['status', '--porcelain'], { cwd: root, encoding: 'utf8' }).trim(),
  permissions: manifest.permissions, files, archive: 't-eraclip-discussions-beta-store.zip', sha256: createHash('sha256').update(readFileSync(zip)).digest('hex'),
  status: 'BETA. Not submitted to or approved by Google. No key-user acceptance claimed.' }, null, 2));
console.log(out);
