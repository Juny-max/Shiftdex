const { spawn } = require('node:child_process');

// Some host tools set this flag for their own Node runtime. Electron inherits it
// unless it is removed, which makes Electron run as Node instead of opening a window.
delete process.env.ELECTRON_RUN_AS_NODE;

const executable = require('electron');
const child = spawn(executable, ['.'], {
  stdio: 'inherit',
  windowsHide: false,
  env: process.env,
});

child.on('error', (error) => {
  console.error('Could not start Electron:', error.message);
  process.exitCode = 1;
});
child.on('exit', (code, signal) => {
  if (signal) {
    console.error(`Electron exited after ${signal}.`);
    process.exitCode = 1;
  } else {
    process.exitCode = code ?? 1;
  }
});
