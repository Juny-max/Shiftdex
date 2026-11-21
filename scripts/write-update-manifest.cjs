const { createHash } = require('node:crypto');
const { readFile, writeFile } = require('node:fs/promises');
const path = require('node:path');

async function main() {
  const root = path.resolve(__dirname, '..');
  const { version } = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
  const file = `ShiftDex-Setup-${version}.exe`;
  const installerPath = path.join(root, 'release', file);
  const installer = await readFile(installerPath);
  const sha512 = createHash('sha512').update(installer).digest('base64');
  const manifest = [
    `version: ${JSON.stringify(version)}`,
    'files:',
    `  - url: ${JSON.stringify(file)}`,
    `    sha512: ${JSON.stringify(sha512)}`,
    `    size: ${installer.byteLength}`,
    `path: ${JSON.stringify(file)}`,
    `sha512: ${JSON.stringify(sha512)}`,
    `releaseDate: ${JSON.stringify(new Date().toISOString())}`,
    '',
  ].join('\n');
  await writeFile(path.join(root, 'release', 'latest.yml'), manifest, 'utf8');
  console.log('GitHub updater manifest ready: release/latest.yml');
}

main().catch(error => {
  console.error(error.message || error);
  process.exitCode = 1;
});
