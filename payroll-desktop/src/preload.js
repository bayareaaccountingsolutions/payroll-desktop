const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('payroll', {
  getConfig: () => ipcRenderer.invoke('cfg:get'),
  setConfig: patch => ipcRenderer.invoke('cfg:set', patch),
  api: (action, params) => ipcRenderer.invoke('api:call', action, params),
  testConnection: conn => ipcRenderer.invoke('api:test', conn),
  savePdf: (base64, name) => ipcRenderer.invoke('file:savePdf', base64, name),
  saveCsv: (text, name) => ipcRenderer.invoke('file:saveCsv', text, name),
  openSheet: inBrowser => ipcRenderer.invoke('sheet:open', !!inBrowser),
  openLink: url => ipcRenderer.invoke('link:open', url),
  scriptFiles: () => ipcRenderer.invoke('script:files'),
  copyScript: name => ipcRenderer.invoke('script:copy', name),
  checkUpdates: () => ipcRenderer.invoke('update:check'),
  installUpdate: () => ipcRenderer.invoke('update:install'),
  onUpdate: cb => ipcRenderer.on('update:status', (e, s) => cb(s))
});
