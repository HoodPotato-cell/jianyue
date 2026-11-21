'use strict';
/* 预加载：把桌面能力桥接给页面（保留手机版的 window.Android 接口） */

const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('Android', {
  hasAllFiles: () => true,
  requestAllFiles: () => {},
  openPath: (p) => ipcRenderer.invoke('jy:open', p),
  deviceInfo: () => 'Windows · 简阅 1.4',
});

contextBridge.exposeInMainWorld('desktop', {
  platform: process.platform,
  startDrag: (p) => ipcRenderer.send('jy:drag', p),
  dropPaths: (list) => ipcRenderer.send('jy:drop', list),
  chooseRoot: () => ipcRenderer.invoke('jy:chooseRoot'),
  reveal: (p) => ipcRenderer.invoke('jy:reveal', p),
  info: () => ipcRenderer.invoke('jy:info'),
});

contextBridge.exposeInMainWorld('jyDesktop', {
  pathForFile: (file) => {
    try { return webUtils.getPathForFile(file); } catch (e) { return file && file.path ? file.path : ''; }
  },
});
