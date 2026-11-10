import { createHash } from 'node:crypto';
import { cpSync, existsSync, lstatSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const dist = join(root, 'dist');
const manifest = JSON.parse(readFileSync(join(dist, 'manifest.json'), 'utf8'));
const expectedPermissions = ['storage', 'contextMenus', 'activeTab', 'scripting'];
if (manifest.manifest_version !== 3 || JSON.stringify(manifest.permissions) !== JSON.stringify(expectedPermissions)
  || manifest.host_permissions?.length || manifest.optional_host_permissions?.length || manifest.content_scripts?.length) {
  throw new Error('Manifest changed beyond the reviewed pilot permission scope. Review before packaging.');
}
if (manifest.description.length > 132) throw new Error('Manifest description is too long.');
const files = [];
function walk(directory) {
  for (const name of readdirSync(directory).sort()) {
    const path = join(directory, name);
    const stat = lstatSync(path);
    if (stat.isSymbolicLink()) throw new Error(`Unexpected symlink: ${path}`);
    if (stat.isDirectory()) { walk(path); continue; }
    const entry = relative(dist, path).split('\\').join('/');
    if (!stat.isFile() || (entry !== 'manifest.json' && !/\.(js|css|html|png|woff2?)$/.test(entry))
      || entry.split('/').some((part) => part.startsWith('.'))) {
      throw new Error(`Unexpected build file: ${entry}`);
    }
    files.push(entry);
  }
}
walk(dist);
const resources = [manifest.action.default_popup, manifest.options_page, manifest.background.service_worker,
  ...Object.values(manifest.icons), ...Object.values(manifest.action.default_icon), 'review.html', 'project-review.html'];
for (const resource of resources) {
  if (!files.includes(resource)) throw new Error(`Missing manifest resource: ${resource}`);
}
const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d+Z$/, 'Z');
const releaseName = `t-eraclip-v${manifest.version}-${stamp}`;
const release = join(root, 'releases', releaseName);
if (existsSync(release)) throw new Error('Release directory already exists.');
mkdirSync(release, { recursive: true });
const pilot = join(release, 'pilot');
mkdirSync(pilot);
cpSync(dist, join(pilot, 'T-eraClip'), { recursive: true });
cpSync(join(root, 'docs/INSTALL_INTERNAL_TEST.md'), join(pilot, 'HUONG-DAN-CAI-DAT.md'));
const store = join(release, 'chrome-web-store.zip');
const pilotZip = join(release, 't-eraclip-test-noi-bo.zip');
execFileSync('zip', ['-X', '-q', store, ...files], { cwd: dist });
execFileSync('zip', ['-X', '-q', '-r', pilotZip, 'T-eraClip', 'HUONG-DAN-CAI-DAT.md'], { cwd: pilot });
execFileSync('unzip', ['-t', store], { stdio: 'pipe' });
execFileSync('unzip', ['-t', pilotZip], { stdio: 'pipe' });
const checksums = [pilotZip, store].map((file) => `${createHash('sha256').update(readFileSync(file)).digest('hex')}  ${relative(release, file)}`).join('\n') + '\n';
writeFileSync(join(release, 'SHA256SUMS.txt'), checksums);
const revision = execFileSync('git', ['rev-parse', 'HEAD'], { cwd: root, encoding: 'utf8' }).trim();
const dirty = Boolean(execFileSync('git', ['status', '--porcelain'], { cwd: root, encoding: 'utf8' }).trim());
writeFileSync(join(release, 'BUILD-INFO.json'), JSON.stringify({
  packagedAt: new Date().toISOString(), version: manifest.version, revision, dirtyWorkingTree: dirty,
  permissions: manifest.permissions, files,
  notes: 'Built from the current working tree. Pilot only; not submitted or approved by Google. No browser profile or live 1Office data is packaged.',
}, null, 2) + '\n');
for (const document of ['CHROME_WEB_STORE_RELEASE.md', 'PRIVACY_POLICY_DRAFT.md']) {
  cpSync(join(root, 'docs', document), join(release, document));
}
console.log(release);
console.log(checksums);
