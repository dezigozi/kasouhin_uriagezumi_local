const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  loadExcelData: (dirPath, forceRefresh) => ipcRenderer.invoke('load-excel-data', dirPath, forceRefresh),
  checkPath: (dirPath) => ipcRenderer.invoke('check-path', dirPath),
  saveCsv: (csvContent) => ipcRenderer.invoke('save-csv', csvContent),
  savePdf: (fileName) => ipcRenderer.invoke('save-pdf', fileName),
  selectFolder: () => ipcRenderer.invoke('select-folder'),
});
