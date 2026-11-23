const { app, BrowserWindow, dialog, Menu, ipcMain, session } = require('electron');
const path = require('node:path');
const fs = require('node:fs');
const { spawn } = require('node:child_process');
const { autoUpdater } = require('electron-updater');
let devServer;
let mainWindow;
let updateState = { status: 'idle', currentVersion: app.getVersion(), version: null, percent: 0, message: '' };

function assertTrustedRenderer(event) {
  const url = event.senderFrame?.url || '';
  const trusted = app.isPackaged
    ? url.startsWith('file://') && url.includes('/dist/admin/index.html')
    : /^http:\/\/127\.0\.0\.1:(517[3-9]|518[0-3])\/admin\/index\.html(?:$|[?#])/.test(url);
  if (!trusted) throw new Error('Untrusted renderer IPC request.');
}

autoUpdater.autoDownload = false;
autoUpdater.autoInstallOnAppQuit = false;

function publishUpdateState(next) {
  updateState = { ...updateState, ...next, currentVersion: app.getVersion() };
  if (mainWindow && !mainWindow.isDestroyed()) mainWindow.webContents.send('updates:status', updateState);
}

autoUpdater.on('checking-for-update', () => publishUpdateState({ status: 'checking', message: 'Checking GitHub for updates…' }));
autoUpdater.on('update-available', info => publishUpdateState({ status: 'available', version: info.version, percent: 0, message: `Version ${info.version} is ready to download.` }));
autoUpdater.on('update-not-available', info => publishUpdateState({ status: 'current', version: info.version, percent: 0, message: 'ShiftDex is up to date.' }));
autoUpdater.on('download-progress', progress => publishUpdateState({ status: 'downloading', version: updateState.version, percent: Math.round(progress.percent || 0), message: `Downloading update… ${Math.round(progress.percent || 0)}%` }));
autoUpdater.on('update-downloaded', info => publishUpdateState({ status: 'downloaded', version: info.version, percent: 100, message: `Version ${info.version} is ready to install.` }));
autoUpdater.on('error', error => publishUpdateState({ status: 'error', message: error?.message || 'Could not check for updates.' }));

async function checkForUpdates() {
  if (!app.isPackaged) {
    publishUpdateState({ status: 'development', message: 'Update checks are available in the installed app.' });
    return updateState;
  }
  if (updateState.status === 'checking' || updateState.status === 'downloading') return updateState;
  try {
    publishUpdateState({ status: 'checking', message: 'Checking GitHub for updates…' });
    await autoUpdater.checkForUpdates();
  } catch (error) {
    publishUpdateState({ status: 'error', message: error?.message || 'Could not check for updates.' });
  }
  return updateState;
}

ipcMain.handle('updates:get-state', event => { assertTrustedRenderer(event); return updateState; });
ipcMain.handle('updates:check', event => { assertTrustedRenderer(event); return checkForUpdates(); });
ipcMain.handle('updates:download', async event => {
  assertTrustedRenderer(event);
  if (updateState.status !== 'available') throw new Error('There is no update ready to download.');
  await autoUpdater.downloadUpdate();
  return updateState;
});
ipcMain.handle('updates:install', event => {
  assertTrustedRenderer(event);
  if (updateState.status !== 'downloaded') throw new Error('Download the update before installing it.');
  autoUpdater.quitAndInstall(false, true);
});
ipcMain.handle('app:uninstall', async event => {
  assertTrustedRenderer(event);
  if (!app.isPackaged || process.platform !== 'win32') throw new Error('Uninstall is available in the installed Windows app.');
  const installDir = path.dirname(app.getPath('exe'));
  const uninstaller = fs.readdirSync(installDir).find(name => /^Uninstall.*\.exe$/i.test(name));
  if (!uninstaller) throw new Error('The ShiftDex uninstaller could not be found.');
  const child = spawn(path.join(installDir, uninstaller), [], { detached: true, stdio: 'ignore', windowsHide: true });
  child.unref();
  setTimeout(() => app.quit(), 500);
  return { started: true };
});

app.setName(app.isPackaged ? 'ShiftDex' : 'ShiftDex Development');
if (process.platform === 'win32') {
  app.setAppUserModelId(app.isPackaged ? 'com.shiftdesk.roster' : 'com.shiftdesk.roster.dev');
}

function createWindow(port) {
  const iconPath = app.isPackaged
    ? path.join(process.resourcesPath, 'shift-desk-mark.ico')
    : path.join(__dirname, '../assets/shift-desk-mark.ico');
  const win = new BrowserWindow({ width: 1500, height: 960, minWidth: 1100, minHeight: 720, backgroundColor: '#f5f6f7', autoHideMenuBar: true, icon: iconPath, webPreferences: { preload: path.join(__dirname, 'preload.cjs'), contextIsolation: true, nodeIntegration: false, sandbox: true, webSecurity: true, allowRunningInsecureContent: false } });
  mainWindow = win;
  win.on('closed', () => { if (mainWindow === win) mainWindow = null; });
  win.setMenu(null);
  win.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  win.webContents.on('will-attach-webview', event => event.preventDefault());
  if (app.isPackaged) {
    win.webContents.on('devtools-opened', () => win.webContents.closeDevTools());
  }
  win.webContents.on('will-navigate', (event, url) => {
    const allowed = app.isPackaged
      ? url.startsWith('file://') && url.includes('/dist/admin/index.html')
      : new URL(url).origin === `http://127.0.0.1:${port}` && new URL(url).pathname === '/admin/index.html';
    if (!allowed) event.preventDefault();
  });
  win.webContents.on('did-fail-load', (_event, code, description, url) => {
    console.error(`Renderer failed to load (${code}): ${description} — ${url}`);
  });
  win.webContents.on('console-message', (_event, details) => {
    if (details.level >= 2) console.error(`Renderer console: ${details.message}`);
  });
  win.webContents.on('render-process-gone', (_event, details) => {
    console.error('Shift Desk renderer stopped:', details.reason, details.exitCode);
  });
  win.webContents.once('did-finish-load', async () => {
    try {
      const snapshot = await win.webContents.executeJavaScript(`(() => { const root = document.querySelector('#app'); return { url: location.href, title: document.title, mounted: !!root && root.children.length > 0, text: root?.innerText?.slice(0, 160) || '' }; })()`);
      console.log('Shift Desk renderer status:', JSON.stringify(snapshot));
      if (!snapshot.mounted) {
        await win.webContents.executeJavaScript(`document.querySelector('#app').innerHTML = '<div style="font:16px Segoe UI,Arial,sans-serif;margin:48px auto;max-width:640px;padding:24px;color:#233248;background:#fff;border:1px solid #dbe2ec;border-radius:14px"><h2>Shift Desk did not finish loading</h2><p>Check the terminal running pnpm dev for the renderer error, then restart the app.</p><button onclick="location.reload()" style="padding:9px 14px;background:#2864e8;color:#fff;border:0;border-radius:7px">Reload</button></div>'`);
      }
    } catch (error) {
      console.error('Renderer status capture failed:', error.message);
    }
  });
  return app.isPackaged
    ? win.loadFile(path.join(__dirname, '../dist/admin/index.html'))
    : win.loadURL(`http://127.0.0.1:${port}/admin/index.html`);
}

app.whenReady().then(async () => {
  Menu.setApplicationMenu(null);
  session.defaultSession.setPermissionRequestHandler((_webContents, _permission, callback) => callback(false));
  session.defaultSession.webRequest.onHeadersReceived((details, callback) => {
    callback({
      responseHeaders: {
        ...details.responseHeaders,
        'Content-Security-Policy': [
          "default-src 'self' file:; script-src 'self' 'unsafe-inline' file:; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com https://fonts.cdnfonts.com; font-src 'self' data: https://fonts.gstatic.com https://fonts.cdnfonts.com; img-src 'self' data: blob: file:; connect-src 'self' https://shift-roster-api.maxgrip500.workers.dev; object-src 'none'; base-uri 'self';"
        ],
        'X-Content-Type-Options': ['nosniff'],
        'X-Frame-Options': ['DENY']
      }
    });
  });
  try {
    if (!app.isPackaged) {
      const { createServer } = await import('vite');
      let port;
      for (let candidatePort = 5173; candidatePort <= 5183; candidatePort++) {
        const candidate = await createServer({
          root: path.resolve(__dirname, '..'),
          configFile: false,
          server: { host: '127.0.0.1', port: candidatePort, strictPort: true },
          appType: 'mpa',
        });
        try {
          await candidate.listen();
          devServer = candidate;
          port = candidatePort;
          break;
        } catch (error) {
          await candidate.close().catch(() => {});
          if (error.code !== 'EADDRINUSE' && !/port \d+ is already in use/i.test(error.message || '')) throw error;
        }
      }
      if (!devServer) throw new Error('No free local port was found between 5173 and 5183.');
      await createWindow(port);
    } else {
      await createWindow();
    }
  } catch (error) {
    console.error('Shift Desk startup failed:', error);
    dialog.showErrorBox('Shift Desk could not start', error.message || String(error));
    app.quit();
  }
  if (app.isPackaged) {
    setTimeout(() => { void checkForUpdates(); }, 15000);
    const timer = setInterval(() => { void checkForUpdates(); }, 6 * 60 * 60 * 1000);
    timer.unref?.();
  }
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length) return;
    if (app.isPackaged) createWindow();
    else if (devServer) createWindow(devServer.httpServer.address().port);
  });
});

app.on('before-quit', () => { devServer?.close(); });
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
