const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const { createExcelBackend } = require('./excelBackend');

// EPIPE クラッシュ防止
process.on('uncaughtException', (err) => {
  if (err.code === 'EPIPE' || err.message?.includes('EPIPE')) return;
});

let mainWindow;
let excelBackend;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1600, height: 1000, minWidth: 1200, minHeight: 800,
    icon: path.join(__dirname, 'app-icon.png'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
    title: '特販部リース会社実績レポート',
    show: false,
  });
  const isDev = process.argv.includes('--dev');
  if (isDev) {
    mainWindow.loadURL('http://localhost:9000');
  } else {
    mainWindow.loadFile(path.join(__dirname, 'build', 'index.html'));
  }
  mainWindow.once('ready-to-show', () => mainWindow.show());
}

app.whenReady().then(() => {
  excelBackend = createExcelBackend({
    cacheDir: path.join(app.getPath('userData'), 'excel-cache'),
  });
  createWindow();
});
app.on('window-all-closed', () => app.quit());
app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });

// ===== IPC ハンドラ =====
ipcMain.handle('load-excel-data', async (event, dirPath, forceRefresh) => {
  try {
    if (!excelBackend) return { success: false, error: 'アプリの初期化中です' };
    return await excelBackend.loadExcelData(dirPath, forceRefresh);
  } catch (err) {
    return { success: false, error: `読み込み中にエラー: ${err.message}` };
  }
});

ipcMain.handle('check-path', async (event, dirPath) => {
  if (!excelBackend) return { exists: false, path: dirPath };
  return excelBackend.checkPath(dirPath);
});

ipcMain.handle('save-csv', async (event, csvContent) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: 'CSVファイルを保存',
    defaultPath: `リース実績レポート_${new Date().toISOString().slice(0, 10)}.csv`,
    filters: [{ name: 'CSV', extensions: ['csv'] }],
  });
  if (result.canceled) return { success: false };
  try {
    fs.writeFileSync(result.filePath, '\uFEFF' + csvContent, 'utf8');
    return { success: true, path: result.filePath };
  } catch (err) { return { success: false, error: err.message }; }
});

ipcMain.handle('save-pdf', async (event, fileName, useA3) => {
  const defaultName = fileName || `リース実績レポート_${new Date().toISOString().slice(0, 10)}`;
  const result = await dialog.showSaveDialog(mainWindow, {
    title: 'PDFファイルを保存',
    defaultPath: `${defaultName}.pdf`,
    filters: [{ name: 'PDF', extensions: ['pdf'] }],
  });
  if (result.canceled) return { success: false };
  try {
    const opts = useA3
      ? { landscape: true, pageSize: 'A3', printBackground: true, scale: 0.75, margins: { top: 0.15, bottom: 0.15, left: 0.15, right: 0.15 } }
      : { landscape: false, pageSize: 'A4', printBackground: true, scale: 0.82, margins: { top: 0.2, bottom: 0.2, left: 0.2, right: 0.2 } };
    const pdfData = await mainWindow.webContents.printToPDF(opts);
    fs.writeFileSync(result.filePath, pdfData);
    return { success: true, path: result.filePath };
  } catch (err) { return { success: false, error: err.message }; }
});

ipcMain.handle('select-folder', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory'],
    title: 'Excelデータフォルダを選択',
  });
  if (result.canceled) return { success: false };
  return { success: true, path: result.filePaths[0] };
});
