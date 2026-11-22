#!/usr/bin/env node
'use strict';

const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const readline = require('node:readline');

const ROOT = path.resolve(__dirname, '..');
const CONFIG = path.join(ROOT, 'worker', 'wrangler.toml');
const DATABASE = 'shift-roster';
// Cloudflare Workers' WebCrypto PBKDF2 limit is 100,000 iterations.
const ITERATIONS = 100000;

function ask(question) {
  return new Promise((resolve, reject) => {
    const terminal = readline.createInterface({ input: process.stdin, output: process.stdout });
    terminal.question(question, answer => { terminal.close(); resolve(answer.trim()); });
    terminal.on('SIGINT', () => { terminal.close(); reject(new Error('Cancelled.')); });
  });
}

function askHidden(question) {
  if (!process.stdin.isTTY || !process.stdin.setRawMode) {
    throw new Error('Use an interactive terminal; passwords must not be echoed or entered as command arguments.');
  }
  return new Promise((resolve, reject) => {
    let value = '';
    process.stdout.write(question);
    const cleanup = () => {
      process.stdin.off('data', onData);
      process.stdin.setRawMode(false);
      process.stdin.pause();
    };
    const onData = buffer => {
      for (const byte of buffer) {
        if (byte === 3) { cleanup(); reject(new Error('Cancelled.')); return; }
        if (byte === 13 || byte === 10) { process.stdout.write('\n'); cleanup(); resolve(value); return; }
        if (byte === 8 || byte === 127) value = value.slice(0, -1);
        else if (byte >= 32 && byte <= 126 && value.length < 256) value += String.fromCharCode(byte);
      }
    };
    process.stdin.setRawMode(true);
    process.stdin.resume();
    process.stdin.on('data', onData);
  });
}

function sqlText(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

function rowsOf(value, predicate) {
  if (Array.isArray(value)) {
    for (const item of value) { const found = rowsOf(item, predicate); if (found) return found; }
  } else if (value && typeof value === 'object') {
    if (Array.isArray(value.results) && value.results.some(predicate)) return value.results;
    for (const child of Object.values(value)) { const found = rowsOf(child, predicate); if (found) return found; }
  }
  return null;
}

function runWrangler(sql) {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'shiftdex-admin-reset-'));
  const sqlFile = path.join(directory, 'reset.sql');
  try {
    fs.writeFileSync(sqlFile, sql, { encoding: 'utf8', flag: 'wx', mode: 0o600 });
    const cli = require.resolve('wrangler');
    const result = spawnSync(process.execPath, [
      cli, 'd1', 'execute', DATABASE, '--remote', '--file', sqlFile,
      '--config', CONFIG, '--json', '--yes',
    ], { cwd: ROOT, encoding: 'utf8', windowsHide: true, maxBuffer: 4 * 1024 * 1024 });
    if (result.error) throw result.error;
    if (result.status !== 0) {
      const message = (result.stderr || result.stdout || 'Wrangler command failed.').trim();
      throw new Error(message.slice(-1200));
    }
    try { return JSON.parse(result.stdout); }
    catch { throw new Error('Could not read Wrangler JSON output. Update Wrangler and run `pnpm exec wrangler whoami` first.'); }
  } finally {
    try { fs.rmSync(directory, { recursive: true, force: true }); } catch {}
  }
}

async function main() {
  if (process.argv.length > 2) throw new Error('This command takes no arguments. Do not pass passwords or emails on the command line.');
  if (!fs.existsSync(CONFIG)) throw new Error('Cloudflare config is missing: worker/wrangler.toml');
  const configText = fs.readFileSync(CONFIG, 'utf8');
  const d1Block = configText.match(/\[\[d1_databases\]\]([\s\S]*?)(?=\n\[|$)/)?.[1] || '';
  const configuredName = d1Block.match(/^\s*database_name\s*=\s*["']([^"']+)["']/m)?.[1];
  const databaseId = d1Block.match(/^\s*database_id\s*=\s*["']([0-9a-f-]{36})["']/im)?.[1];
  if (configuredName !== DATABASE || !databaseId) {
    throw new Error('The local Wrangler config must identify the shift-roster D1 database before this reset can run.');
  }

  console.log('Reset the password for one existing ShiftDex administrator.');
  console.log(`Target: remote D1 “${DATABASE}” (ID ending ${databaseId.slice(-6)}).`);
  console.log('The script cannot create accounts. It requires Wrangler sign-in and D1 write access to this database.');
  console.log('It revokes the selected admin’s existing sessions; it does not change roster or worker data.\n');

  const email = (await ask('Existing admin email: ')).toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 254) throw new Error('Enter a valid administrator email.');
  const lookup = runWrangler(`SELECT id FROM admin_accounts WHERE lower(email)=${sqlText(email)} AND status='active' LIMIT 1;`);
  const accounts = rowsOf(lookup, row => typeof row.id === 'string');
  const account = accounts?.find(row => typeof row.id === 'string');
  if (!account) throw new Error('No active admin account matches that email. No account was changed.');

  const password = await askHidden('New password (minimum 16 characters; input hidden): ');
  if (password.length < 16 || password.length > 256) throw new Error('Choose a password between 16 and 256 characters.');
  const confirm = await askHidden('Confirm new password: ');
  if (password !== confirm) throw new Error('Passwords do not match. No account was changed.');
  const phrase = await ask(`Type RESET ${email} to confirm: `);
  if (phrase !== `RESET ${email}`) throw new Error('Confirmation did not match. No account was changed.');

  const salt = crypto.randomBytes(16).toString('hex');
  const digest = crypto.pbkdf2Sync(password, Buffer.from(salt, 'hex'), ITERATIONS, 32, 'sha256').toString('hex');
  const auditId = crypto.randomUUID();
  const sql = `DELETE FROM admin_sessions WHERE admin_id=${sqlText(account.id)};
UPDATE admin_accounts SET password_salt=${sqlText(salt)},password_hash=${sqlText(digest)},password_iterations=${ITERATIONS},updated_at=CURRENT_TIMESTAMP WHERE id=${sqlText(account.id)} AND status='active';
INSERT INTO audit_logs(id,actor_id,action,entity_type,entity_id,details_json) VALUES(${sqlText(auditId)},NULL,'reset_admin_password','admin_account',${sqlText(account.id)},'{"method":"authorized D1 reset script"}');
SELECT id,password_iterations FROM admin_accounts WHERE id=${sqlText(account.id)} AND status='active' AND password_salt=${sqlText(salt)};`;
  const result = runWrangler(sql);
  const verified = rowsOf(result, row => row.id === account.id && Number(row.password_iterations) === ITERATIONS);
  if (!verified) throw new Error('The reset could not be verified. Check D1 and rerun only after confirming the account state.');
  console.log('\nPassword reset complete. Existing sessions for this admin were revoked. Sign in with the new password.');
}

main().catch(error => {
  console.error(`Admin password reset stopped: ${error.message}`);
  process.exitCode = 1;
});
