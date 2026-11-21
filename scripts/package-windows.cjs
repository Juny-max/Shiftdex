const { spawn } = require('node:child_process');
const { mkdir, copyFile, cp, readFile } = require('node:fs/promises');
const path = require('node:path');

async function main() {
  const root = path.resolve(__dirname, '..');
  const { version } = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
  if (!/^\d+\.\d+\.\d+$/.test(version)) throw new Error('Expected a numeric release version.');

  const stage = path.join(root, '.packaging', `${version}-${process.pid}`);
  const electronDist = path.join(root, '.packaging', 'electron-dist');
  const release = path.join(root, 'release');
  const builder = path.join(root, 'node_modules', 'electron-builder', 'cli.js');
  const installedElectronDist = path.join(root, 'node_modules', 'electron', 'dist');
  await cp(installedElectronDist, electronDist, {
    recursive: true,
    filter: source => path.basename(source) !== 'default_app.asar',
  });
  const code = await new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [
      builder,
      '--win', 'nsis',
      '--publish', 'never',
      `--config.directories.output=${stage}`,
      `--config.electronDist=${electronDist}`,
    ], { cwd: root, stdio: 'inherit', windowsHide: true });
    child.on('error', reject);
    child.on('exit', resolve);
  });
  if (code !== 0) throw new Error(`Windows packaging failed (${code}); previous installers were preserved.`);

  await mkdir(release, { recursive: true });
  const installer = `ShiftDex-Setup-${version}.exe`;
  await copyFile(path.join(stage, installer), path.join(release, installer));
  console.log(`Installer ready: ${path.join('release', installer)}`);
}

main().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
