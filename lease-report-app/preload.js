const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  isElectron: true,
  loadExcelData: (dirPath, forceRefresh) => ipcRenderer.invoke('load-excel-data', dirPath, forceRefresh),
  checkPath: (dirPath) => ipcRenderer.invoke('check-path', dirPath),
  saveCsv: (csvContent) => ipcRenderer.invoke('save-csv', csvContent),
  savePdf: (fileName, useA3) => ipcRenderer.invoke('save-pdf', fileName, useA3),
  selectFolder: () => ipcRenderer.invoke('select-folder'),
});
