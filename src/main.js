/*  Payroll System — desktop app (Electron main process)
 *  Talks to the shared Google Sheet through the "DesktopApi.gs" web app, so every
 *  staff member on Windows or Mac sees and edits the same live payroll data.        */
const { app, BrowserWindow, ipcMain, shell, dialog, net, clipboard, Menu } = require('electron');
const path = require('path');
const fs = require('fs');

const APP_DIR = __dirname;
const SCRIPT_DIR = path.join(APP_DIR, '..', 'apps-script');
const CHECK_EVERY_MS = 4 * 60 * 60 * 1000;          // look for app updates every 4 hours
let win = null, sheetWin = null;

/* ------------------------------ settings file ------------------------------ */
const cfgPath = () => path.join(app.getPath('userData'), 'config.json');
function loadCfg() {
  try { return JSON.parse(fs.readFileSync(cfgPath(), 'utf8')); } catch (e) { return {}; }
}
function saveCfg(patch) {
  const c = Object.assign(loadCfg(), patch);
  fs.mkdirSync(path.dirname(cfgPath()), { recursive: true });
  fs.writeFileSync(cfgPath(), JSON.stringify(c, null, 2));
  return c;
}

/* ------------------------------- API client -------------------------------- */
const DEMO = process.env.PAYROLL_DEMO === '1';
async function callApi(action, params, override) {
  if (DEMO) { await new Promise(r => setTimeout(r, 150)); return require('./mock')(action, params || {}); }
  const c = Object.assign({}, loadCfg(), override || {});
  if (!c.webAppUrl || !c.key) return { ok: false, error: 'Not connected yet. Open Settings ▸ Connection.', code: 'CONFIG' };
  if (!/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec/.test(c.webAppUrl.trim()))
    return { ok: false, error: 'The Web App URL must look like https://script.google.com/macros/s/…/exec', code: 'CONFIG' };
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), 5.5 * 60 * 1000);   // Apps Script limit is 6 minutes
  try {
    const res = await net.fetch(c.webAppUrl.trim(), {
      method: 'POST', redirect: 'follow', signal: ctrl.signal,
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ key: c.key.trim(), user: c.userName || 'Unknown', action: action, params: params || {} })
    });
    const text = await res.text();
    try { return JSON.parse(text); }
    catch (e) {
      if (/accounts\.google\.com|ServiceLogin|<html/i.test(text))
        return { ok: false, error: 'Google returned a sign-in page. In Apps Script ▸ Deploy ▸ Manage deployments, set "Who has access" to "Anyone" and use the /exec URL.', code: 'DEPLOY' };
      return { ok: false, error: 'Unexpected reply from Google (HTTP ' + res.status + ').' };
    }
  } catch (err) {
    return { ok: false, error: err.name === 'AbortError' ? 'Google took too long to answer. Try again.' : 'No connection to Google: ' + err.message, code: 'NET' };
  } finally { clearTimeout(timer); }
}

/* ------------------------------- windows ----------------------------------- */
function createWindow() {
  const c = loadCfg();
  win = new BrowserWindow({
    width: c.w || 1380, height: c.h || 860, minWidth: 1024, minHeight: 640,
    title: 'Payroll System', backgroundColor: '#F4F6FA', show: false,
    icon: path.join(APP_DIR, '..', 'build', 'icon.png'),
    webPreferences: { preload: path.join(APP_DIR, 'preload.js'), contextIsolation: true, nodeIntegration: false, sandbox: true }
  });
  win.loadFile(path.join(APP_DIR, 'renderer', 'index.html'));
  win.once('ready-to-show', () => win.show());
  win.on('close', () => { const [w, h] = win.getSize(); saveCfg({ w: w, h: h }); });
  win.webContents.setWindowOpenHandler(({ url }) => { if (/^https:\/\//.test(url)) shell.openExternal(url); return { action: 'deny' }; });
  win.webContents.on('will-navigate', e => e.preventDefault());
}

/** The live Google Sheet inside the app (same menus, buttons and scripts as in the browser). */
function openSheetWindow(url) {
  if (sheetWin && !sheetWin.isDestroyed()) { sheetWin.focus(); if (url) sheetWin.loadURL(url); return; }
  sheetWin = new BrowserWindow({
    width: 1400, height: 900, title: 'Payroll System — Google Sheet', parent: undefined,
    icon: path.join(APP_DIR, '..', 'build', 'icon.png'),
    webPreferences: { partition: 'persist:google', contextIsolation: true, nodeIntegration: false, sandbox: true }
  });
  // Google sign-in refuses browsers that announce themselves as "Electron": present a normal Chrome identity.
  const ua = `Mozilla/5.0 (${process.platform === 'darwin' ? 'Macintosh; Intel Mac OS X 10_15_7' : 'Windows NT 10.0; Win64; x64'}) ` +
    `AppleWebKit/537.36 (KHTML, like Gecko) Chrome/${process.versions.chrome} Safari/537.36`;
  sheetWin.webContents.setUserAgent(ua);
  sheetWin.webContents.setWindowOpenHandler(({ url: u }) => {
    if (/^https:\/\/(docs|drive|accounts|script)\.google\.com\//.test(u))
      return { action: 'allow', overrideBrowserWindowOptions: { webPreferences: { partition: 'persist:google' } } };
    shell.openExternal(u); return { action: 'deny' };
  });
  sheetWin.loadURL(url, { userAgent: ua });
  sheetWin.on('closed', () => { sheetWin = null; });
}

/* --------------------------------- IPC ------------------------------------- */
ipcMain.handle('cfg:get', () => { const c = loadCfg(); return Object.assign({}, c, { appVersion: app.getVersion(), platform: process.platform, demo: DEMO }); });
ipcMain.handle('cfg:set', (e, patch) => saveCfg(patch || {}));
ipcMain.handle('api:call', (e, action, params) => callApi(action, params));
ipcMain.handle('api:test', (e, conn) => callApi('ping', {}, conn));

ipcMain.handle('file:savePdf', async (e, base64, fileName) => {
  const r = await dialog.showSaveDialog(win, { defaultPath: path.join(app.getPath('documents'), fileName), filters: [{ name: 'PDF', extensions: ['pdf'] }] });
  if (r.canceled || !r.filePath) return { saved: false };
  fs.writeFileSync(r.filePath, Buffer.from(base64, 'base64'));
  shell.openPath(r.filePath);
  return { saved: true, path: r.filePath };
});
ipcMain.handle('file:saveCsv', async (e, text, fileName) => {
  const r = await dialog.showSaveDialog(win, { defaultPath: path.join(app.getPath('documents'), fileName), filters: [{ name: 'CSV (Excel)', extensions: ['csv'] }] });
  if (r.canceled || !r.filePath) return { saved: false };
  fs.writeFileSync(r.filePath, '﻿' + text, 'utf8');          // BOM so Excel shows ₹ correctly
  return { saved: true, path: r.filePath };
});

ipcMain.handle('sheet:open', (e, inBrowser) => {
  const url = loadCfg().sheetUrl;
  if (!url) return { ok: false, error: 'The Google Sheet link is not set (Settings ▸ Connection — it fills in automatically after "Test & Save").' };
  if (inBrowser) shell.openExternal(url); else openSheetWindow(url);
  return { ok: true };
});
ipcMain.handle('link:open', (e, url) => { if (/^https:\/\//.test(url)) shell.openExternal(url); });

ipcMain.handle('script:files', () => {
  const out = {};
  ['Code.gs', 'DesktopApi.gs'].forEach(f => { try { out[f] = fs.readFileSync(path.join(SCRIPT_DIR, f), 'utf8'); } catch (e) { out[f] = ''; } });
  const m = /SHEET_SCRIPT_VERSION\s*=\s*'([^']+)'/.exec(out['DesktopApi.gs']);
  const a = /DESKTOP_API_VERSION\s*=\s*'([^']+)'/.exec(out['DesktopApi.gs']);
  return { files: Object.keys(out).map(k => ({ name: k, lines: out[k].split('\n').length })), scriptVersion: m ? m[1] : '', apiVersion: a ? a[1] : '' };
});
ipcMain.handle('script:copy', (e, name) => {
  if (['Code.gs', 'DesktopApi.gs'].indexOf(name) < 0) return false;
  clipboard.writeText(fs.readFileSync(path.join(SCRIPT_DIR, name), 'utf8'));
  return true;
});

/* ------------------------------ auto-update -------------------------------- */
let updater = null;
const pkg = require('../package.json');
const pub = (pkg.build && pkg.build.publish && pkg.build.publish[0]) || {};
const releasesUrl = pub.owner ? `https://github.com/${pub.owner}/${pub.repo}/releases/latest` : '';
function sendUpdate(state, info) { if (win && !win.isDestroyed()) win.webContents.send('update:status', Object.assign({ state: state }, info || {})); }

function setupUpdater() {
  if (!app.isPackaged) { sendUpdate('dev'); return; }
  try { updater = require('electron-updater').autoUpdater; } catch (e) { return; }
  // Windows installs updates itself. Mac auto-install needs an Apple-signed build; unsigned Mac builds get a download link.
  const macUnsigned = process.platform === 'darwin' && !pkg.macSigned;
  updater.autoDownload = !macUnsigned;
  updater.autoInstallOnAppQuit = true;
  updater.on('checking-for-update', () => sendUpdate('checking'));
  updater.on('update-not-available', () => sendUpdate('none'));
  updater.on('update-available', i => sendUpdate(macUnsigned ? 'available-manual' : 'downloading', { version: i.version, notes: notesText(i), url: releasesUrl }));
  updater.on('download-progress', p => sendUpdate('downloading', { percent: Math.round(p.percent) }));
  updater.on('update-downloaded', i => sendUpdate('ready', { version: i.version, notes: notesText(i) }));
  updater.on('error', err => sendUpdate('error', { message: String(err && err.message || err).split('\n')[0], url: releasesUrl }));
  const check = () => updater.checkForUpdates().catch(() => { });
  setTimeout(check, 8000);
  setInterval(check, CHECK_EVERY_MS);
}
function notesText(i) {
  const n = i.releaseNotes;
  if (!n) return '';
  const s = Array.isArray(n) ? n.map(x => x.note).join('\n') : String(n);
  return s.replace(/<[^>]+>/g, '').trim().slice(0, 1500);
}
ipcMain.handle('update:check', () => { if (updater) updater.checkForUpdates().catch(e => sendUpdate('error', { message: e.message, url: releasesUrl })); else sendUpdate('dev'); });
ipcMain.handle('update:install', () => { if (updater) updater.quitAndInstall(false, true); });

/* -------------------------------- app life --------------------------------- */
if (!app.requestSingleInstanceLock()) app.quit();
else {
  app.on('second-instance', () => { if (win) { if (win.isMinimized()) win.restore(); win.focus(); } });
  app.whenReady().then(() => {
    const isMac = process.platform === 'darwin';
    Menu.setApplicationMenu(Menu.buildFromTemplate([
      ...(isMac ? [{ role: 'appMenu' }] : []),
      { label: 'File', submenu: [{ label: 'Open Google Sheet', click: () => { const u = loadCfg().sheetUrl; if (u) openSheetWindow(u); } }, { type: 'separator' }, isMac ? { role: 'close' } : { role: 'quit' }] },
      { role: 'editMenu' },
      { label: 'View', submenu: [{ role: 'reload' }, { role: 'resetZoom' }, { role: 'zoomIn' }, { role: 'zoomOut' }, { type: 'separator' }, { role: 'togglefullscreen' }, { role: 'toggleDevTools' }] },
      { label: 'Help', submenu: [{ label: 'Check for updates', click: () => updater ? updater.checkForUpdates().catch(() => { }) : sendUpdate('dev') },
        { label: 'Release notes', click: () => releasesUrl && shell.openExternal(releasesUrl) }] }
    ]));
    createWindow();
    setupUpdater();
    app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
  });
  app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
}
