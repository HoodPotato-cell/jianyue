'use strict';
/* 简阅 Windows · Electron 主进程 */

const { app, BrowserWindow, ipcMain, dialog, shell, Menu } = require('electron');
const path = require('path');
const fs = require('fs');
const { start } = require('./server');

function logError(e) {
  try {
    const dir = process.env.APPDATA ? path.join(process.env.APPDATA, '简阅') : __dirname;
    fs.mkdirSync(dir, { recursive: true });
    fs.appendFileSync(path.join(dir, 'error.log'), `[${new Date().toISOString()}] ${e && e.stack ? e.stack : e}\n`);
  } catch (_) {}
}
process.on('uncaughtException', (e) => { logError(e); dialog.showErrorBox('简阅出错', String(e && e.message ? e.message : e)); });

let srv = null;
let win = null;
const ICON = path.join(__dirname, 'build', 'icon.png');

async function createWindow() {
  srv = await start({
    assetsDir: path.join(__dirname, 'assets', 'web'),
    storeFile: path.join(app.getPath('userData'), 'store.json'),
    roots: [],
    port: 0,
  });

  win = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 960,
    minHeight: 620,
    title: '简阅',
    backgroundColor: '#F5F6F8',
    icon: ICON,
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  win.once('ready-to-show', () => win.show());
  win.loadURL(srv.url);
  win.on('closed', () => { win = null; });
  buildMenu();
}

async function chooseRoot() {
  if (!win) return;
  const r = await dialog.showOpenDialog(win, {
    title: '选择要扫描的文件夹',
    properties: ['openDirectory'],
  });
  if (r.canceled || !r.filePaths.length) return;
  srv.addRoot(r.filePaths[0]);
  win.reload();
}

function buildMenu() {
  const template = [
    {
      label: '文件',
      submenu: [
        { label: '添加扫描文件夹…', accelerator: 'CmdOrCtrl+O', click: chooseRoot },
        { label: '重新扫描', accelerator: 'F5', click: () => { srv.rescan(); if (win) win.reload(); } },
        { type: 'separator' },
        { role: 'quit', label: '退出' },
      ],
    },
    {
      label: '视图',
      submenu: [
        { role: 'reload', label: '重新载入' },
        { role: 'toggleDevTools', label: '开发者工具' },
        { type: 'separator' },
        { role: 'resetZoom', label: '实际大小' },
        { role: 'zoomIn', label: '放大' },
        { role: 'zoomOut', label: '缩小' },
        { type: 'separator' },
        { role: 'togglefullscreen', label: '全屏' },
      ],
    },
    {
      label: '帮助',
      submenu: [
        {
          label: '关于简阅',
          click: () => dialog.showMessageBox(win, {
            type: 'info',
            title: '简阅',
            message: '简阅 1.4 · Windows',
            detail: 'PDF / Word / Excel / PPT 阅读器\n\n· 从列表拖文档到资源管理器、微信、邮件，可直接当文件用\n· 把文件或文件夹拖进窗口，会自动加入扫描目录\n· 书库、删除、缩放、旧版 .doc/.xls/.ppt 与手机版一致',
          }),
        },
      ],
    },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

/* 拖出：用 Electron 原生拖拽，带真实文件路径 */
ipcMain.on('jy:drag', (evt, filePath) => {
  if (!filePath || !fs.existsSync(filePath)) return;
  try {
    evt.sender.startDrag({ file: filePath, icon: ICON });
  } catch (e) {
    console.error('拖拽失败', e);
  }
});

/* 拖入：文件加入其所在目录，文件夹直接加入扫描目录 */
ipcMain.on('jy:drop', (evt, paths) => {
  if (!Array.isArray(paths) || !srv) return;
  for (const p of paths) {
    try {
      const st = fs.statSync(p);
      srv.addRoot(st.isDirectory() ? p : path.dirname(p));
    } catch (e) { /* 跳过不存在的路径 */ }
  }
  setTimeout(() => { if (win) win.reload(); }, 700);
});

ipcMain.handle('jy:open', (evt, p) => shell.openPath(p));
ipcMain.handle('jy:reveal', (evt, p) => { shell.showItemInFolder(p); return true; });
ipcMain.handle('jy:chooseRoot', async () => { await chooseRoot(); return true; });
ipcMain.handle('jy:info', () => ({ platform: process.platform, version: app.getVersion() }));

app.whenReady().then(createWindow).catch((e) => {
  logError(e);
  dialog.showErrorBox('简阅启动失败', String(e && e.stack ? e.stack : e));
  app.quit();
});

app.on('window-all-closed', () => app.quit());
app.on('activate', () => { if (!win) createWindow(); });
