#!/usr/bin/env node
'use strict';

const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const readline = require('node:readline');

const ROOT = path.resolve(__dirname, '..');
const CONFIG = path.join(ROOT, 'worker', 'wrangler.toml');
const DATABASE = 'shift-roster';
// Cloudflare Workers rejects PBKDF2 iteration counts above 100,000.
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
    throw new Error('Run this script in an interactive terminal so the password is not echoed or saved in shell history.');
  }
  return new Promise((resolve, reject) => {
    let value = '';
    process.stdout.write(question);
    const onData = buffer => {
      for (const byte of buffer) {
        if (byte === 3) {
          cleanup(); reject(new Error('Cancelled.')); return;
        }
        if (byte === 13 || byte === 10) {
          process.stdout.write('\n'); cleanup(); resolve(value); return;
        }
        if (byte === 8 || byte === 127) value = value.slice(0, -1);
        else if (byte >= 32 && byte <= 126 && value.length < 256) value += String.fromCharCode(byte);
      }
    };
    const cleanup = () => {
      process.stdin.off('data', onData);
      process.stdin.setRawMode(false);
      process.stdin.pause();
    };
    process.stdin.setRawMode(true);
    process.stdin.resume();
    process.stdin.on('data', onData);
  });
}

function sqlText(value) {
  return `'${value.replace(/'/g, "''")}'`;
}

function runWrangler(sql) {
  const cli = require.resolve('wrangler');
  const result = spawnSync(process.execPath, [
    cli, 'd1', 'execute', DATABASE, '--remote', '--command', sql,
    '--config', CONFIG, '--json', '--yes',
  ], { cwd: ROOT, encoding: 'utf8', windowsHide: true, maxBuffer: 4 * 1024 * 1024 });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    const message = (result.stderr || result.stdout || 'Wrangler command failed.').trim();
    throw new Error(message.slice(-1600));
  }
  try { return JSON.parse(result.stdout); }
  catch { throw new Error('Could not read Wrangler’s JSON response. Update Wrangler or run `pnpm exec wrangler whoami` first.'); }
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

async function main() {
  if (!fs.existsSync(CONFIG)) throw new Error('Cloudflare config is missing: worker/wrangler.toml');
  console.log('Create the first ShiftDex administrator account.');
  console.log('This command only adds one admin to the configured remote D1 database. It does not deploy code or change Cloudflare settings.');
  console.log('');
  const email = (await ask('Admin email: ')).toLowerCase();
  const fullName = await ask('Admin name: ');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 254) throw new Error('Enter a valid email address.');
  if (fullName.length < 2 || fullName.length > 120) throw new Error('Enter a name between 2 and 120 characters.');
  const password = await askHidden('Password (minimum 14 characters; input hidden): ');
  if (password.length < 14 || password.length > 256) throw new Error('Choose a password between 14 and 256 characters.');
  const confirm = await askHidden('Confirm password: ');
  if (password !== confirm) throw new Error('Passwords do not match.');
  const phrase = await ask('Type CREATE FIRST ADMIN to confirm: ');
  if (phrase !== 'CREATE FIRST ADMIN') throw new Error('Confirmation did not match. No account was created.');

  console.log('\nChecking Cloudflare authentication and the configured D1 database…');
  const countResult = runWrangler('SELECT COUNT(*) AS count FROM admin_accounts;');
  const countRows = rowsOf(countResult, row => Object.hasOwn(row, 'count'));
  if (!countRows) throw new Error('The admin_accounts table is missing. Apply worker/schema.sql to the configured remote D1 database first.');
  if (Number(countRows[0].count) !== 0) throw new Error('An administrator already exists. Bootstrap is permanently closed; this script will not add or change accounts.');

  const id = crypto.randomUUID();
  const salt = crypto.randomBytes(16).toString('hex');
  const passwordHash = crypto.pbkdf2Sync(password, Buffer.from(salt, 'hex'), ITERATIONS, 32, 'sha256').toString('hex');
  const insert = `INSERT INTO admin_accounts(id,email,full_name,password_salt,password_hash,password_iterations,status)
    SELECT ${sqlText(id)},${sqlText(email)},${sqlText(fullName)},${sqlText(salt)},${sqlText(passwordHash)},${ITERATIONS},'active'
    WHERE NOT EXISTS (SELECT 1 FROM admin_accounts);`;
  runWrangler(insert);
  const verify = runWrangler(`SELECT id,email FROM admin_accounts WHERE id=${sqlText(id)};`);
  const created = rowsOf(verify, row => row.id === id);
  if (!created) throw new Error('The database already has an administrator or the account could not be confirmed. No other account was modified.');
  console.log(`Administrator account created for ${email}. Sign in to ShiftDex with this email and password.`);
  console.log('The password was sent only as a PBKDF2 hash; keep your password manager entry safe.');
}

main().catch(error => {
  console.error(`Admin setup stopped: ${error.message}`);
  process.exitCode = 1;
});
