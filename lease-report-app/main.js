const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const officeCrypto = require('officecrypto-tool');
const crypto = require('crypto');
const zlib = require('zlib');

// EPIPE クラッシュ防止
process.on('uncaughtException', (err) => {
  if (err.code === 'EPIPE' || err.message?.includes('EPIPE')) return;
});

let mainWindow;

// ===== キャッシュ管理 =====
const CACHE_DIR = path.join(app.getPath('userData'), 'excel-cache');

function ensureCacheDir() {
  if (!fs.existsSync(CACHE_DIR)) {
    fs.mkdirSync(CACHE_DIR, { recursive: true });
  }
}

function getCacheKey(dirPath) {
  return crypto.createHash('md5').update(dirPath).digest('hex');
}

function getCachePath(dirPath) {
  return path.join(CACHE_DIR, `${getCacheKey(dirPath)}.gz`);
}

function getFilesFingerprint(dirPath) {
  try {
    const files = fs.readdirSync(dirPath)
      .filter(f => /\.(xlsx|xls)$/i.test(f) && !f.startsWith('~$'))
      .map(f => {
        const stat = fs.statSync(path.join(dirPath, f));
        return `${f}:${stat.size}`;
      })
      .sort();
    return files.join('|');
  } catch {
    return '';
  }
}

function loadCache(dirPath) {
  try {
    ensureCacheDir();
    const cachePath = getCachePath(dirPath);
    console.log(`[Cache] 検索パス: ${cachePath}`);
    console.log(`[Cache] CACHE_DIR: ${CACHE_DIR}`);
    if (!fs.existsSync(cachePath)) {
      console.log('[Cache] キャッシュファイルなし');
      try {
        const files = fs.readdirSync(CACHE_DIR);
        console.log(`[Cache] キャッシュフォルダ内: ${files.join(', ')}`);
      } catch(e) { console.log(`[Cache] フォルダ読み取りエラー: ${e.message}`); }
      return null;
    }

    console.log('[Cache] キャッシュ読み込み中...');
    const compressed = fs.readFileSync(cachePath);
    const json = zlib.gunzipSync(compressed).toString('utf8');
    const cache = JSON.parse(json);

    const currentFingerprint = getFilesFingerprint(dirPath);
    if (cache.fingerprint !== currentFingerprint) {
      console.log('[Cache] フィンガープリント不一致 → キャッシュ無効');
      return null;
    }

    console.log(`[Cache] キャッシュ有効（${cache.data.totalRows}件, ${(compressed.length / 1024 / 1024).toFixed(1)}MB）`);
    return cache.data;
  } catch (err) {
    console.log('[Cache] 読み込み失敗:', err.message);
    return null;
  }
}

function saveCache(dirPath, data) {
  try {
    ensureCacheDir();
    const cachePath = getCachePath(dirPath);
    const fingerprint = getFilesFingerprint(dirPath);

    const slim = {
      ...data,
      rows: data.rows.map(r => ({
        leaseCompany: r.leaseCompany,
        branch: r.branch,
        ordererName: r.ordererName,
        customerName: r.customerName,
        rep: r.rep,
        sales: r.sales,
        profit: r.profit,
        fiscalYear: r.fiscalYear,
        month: r.month,
      })),
      diagnostics: undefined,
    };

    const json = JSON.stringify({ fingerprint, dirPath, data: slim });
    console.log(`[Cache] 圧縮中... (JSON: ${(json.length / 1024 / 1024).toFixed(1)}MB)`);
    const compressed = zlib.gzipSync(json, { level: 6 });
    fs.writeFileSync(cachePath, compressed);
    console.log(`[Cache] 保存完了（${(compressed.length / 1024 / 1024).toFixed(1)}MB）: ${cachePath}`);
  } catch (err) {
    console.error('[Cache] 保存失敗:', err.message);
  }
}

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

app.whenReady().then(createWindow);
app.on('window-all-closed', () => app.quit());
app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });

// ===== Excelファイル一覧取得 =====
function getExcelFiles(dirPath) {
  try {
    if (!fs.existsSync(dirPath)) {
      return { success: false, error: `パスが見つかりません: ${dirPath}` };
    }
    const files = fs.readdirSync(dirPath)
      .filter(f => /\.(xlsx|xls)$/i.test(f) && !f.startsWith('~$'))
      .map(f => ({
        name: f,
        path: path.join(dirPath, f),
        modified: fs.statSync(path.join(dirPath, f)).mtime.toISOString(),
        size: fs.statSync(path.join(dirPath, f)).size,
      }));
    return { success: true, files };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ===== ヘッダー名から列を自動検出 =====
// ヘッダーパターン: 各キーに対して [パターン文字列, ...] で定義
// 'prefer_last' が true のキーは、最後にマッチした列を採用（部店など複数マッチする場合）
const HEADER_PATTERNS = {
  leaseCompany: { patterns: ['リース会社', 'リース', '架装実績', 'カウヒン'], prefer_last: true },
  branch:       { patterns: ['部店', '部店名', '支店', 'プチン'], prefer_last: true },
  ordererName:  { patterns: ['注文社名', '注文者名', '注文者', '注文先', '注文社'], prefer_last: false },
  customerName: { patterns: ['顧客名_漢字', '顧客名', '顧客'], prefer_last: false },
  repLastName:  { patterns: ['担当者名', '担当者', '担当'], prefer_last: false },
  salesAmount:  { patterns: ['金額', '売上金額', '売上'], prefer_last: true },
  grossProfit:  { patterns: ['粗利', '粗利額'], prefer_last: false },
  deliveryDate: { patterns: ['納品日', '納品', '納入日'], prefer_last: false },
  slipDate:     { patterns: ['売上伝票日', '伝票日', '売上日', '計上日'], prefer_last: false },
};

// 強制オーバーライド列番号（診断データから確認済み）
// AH(34)列=リース会社(SMAS等), AI(35)列=部店(MA営業第4部等)
const OVERRIDE_COLUMNS = {
  leaseCompany: 34, // AH列 - リース会社
  branch: 35,       // AI列 - 部店（詳細）
};

function detectColumnsFromRow(row) {
  const mapping = {};
  const totalCols = row.cellCount || 0;

  row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    const val = String(cell.value || '').trim();
    if (!val) return;
    for (const [key, config] of Object.entries(HEADER_PATTERNS)) {
      for (const pat of config.patterns) {
        if (val.includes(pat)) {
          if (config.prefer_last) {
            // 常に上書き（後ろの列を優先）
            mapping[key] = colNumber;
          } else if (!mapping[key]) {
            // 最初にマッチした列を採用
            mapping[key] = colNumber;
          }
          break;
        }
      }
    }
  });

  // オーバーライド列を強制適用（診断データから確定済みの列番号）
  for (const [key, overrideCol] of Object.entries(OVERRIDE_COLUMNS)) {
    if (overrideCol <= totalCols || overrideCol <= 50) {
      mapping[key] = overrideCol;
    }
  }

  return mapping;
}

function colNumToLetter(n) {
  let s = '';
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}

// ===== 日付処理 =====
function parseExcelDate(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  if (typeof val === 'number') {
    const epoch = new Date(1899, 11, 30);
    return new Date(epoch.getTime() + val * 86400000);
  }
  if (typeof val === 'string') {
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

function getFiscalYear(dateVal) {
  const d = parseExcelDate(dateVal);
  if (!d) return null;
  return d.getMonth() + 1 >= 4 ? d.getFullYear() : d.getFullYear() - 1;
}

function getMonth(dateVal) {
  const d = parseExcelDate(dateVal);
  if (!d) return null;
  return d.getMonth() + 1;
}

// ===== セル値変換 =====
function toNumber(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  if (typeof val === 'object' && val.result !== undefined) return toNumber(val.result);
  const n = parseFloat(String(val).replace(/[,¥\\]/g, ''));
  return isNaN(n) ? 0 : n;
}

function toStr(val) {
  if (val === null || val === undefined) return '';
  if (typeof val === 'object' && val.richText) return val.richText.map(r => r.text).join('');
  if (typeof val === 'object' && val.result !== undefined) return String(val.result);
  return String(val).trim();
}

const EXCEL_PASSWORD = '1962';

// ===== パスワード付きExcel読み込み（OLE2暗号化 → 復号 → ストリーミング）=====
async function readExcelWithPassword(filePath, password) {
  const encryptedBuf = fs.readFileSync(filePath);
  console.log(`[Decrypt] 暗号化ファイルサイズ: ${(encryptedBuf.length / 1024 / 1024).toFixed(1)}MB`);
  const decryptedBuf = await officeCrypto.decrypt(encryptedBuf, { password });
  console.log(`[Decrypt] 復号後サイズ: ${(decryptedBuf.length / 1024 / 1024).toFixed(1)}MB`);

  const os = require('os');
  const tmpPath = path.join(os.tmpdir(), `_decrypted_${Date.now()}.xlsx`);
  fs.writeFileSync(tmpPath, decryptedBuf);
  console.log(`[Decrypt] 一時ファイル: ${tmpPath}`);

  try {
    const result = await readExcelStreaming(tmpPath);
    return result;
  } finally {
    try { fs.unlinkSync(tmpPath); } catch {}
  }
}

// ===== ExcelJSストリーミング読み込み（大規模ファイル対応）=====
async function readExcelStreaming(filePath) {
  return new Promise((resolve, reject) => {
    const allRows = [];
    let colMap = null;
    let headerInfo = {};
    let rowCount = 0;
    let sheetName = '';
    let sampleRow = {};

    const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, {
      sharedStrings: 'cache',
      hyperlinks: 'ignore',
      styles: 'ignore',
      worksheets: 'emit',
    });

    workbookReader.on('worksheet', (worksheetReader) => {
      sheetName = worksheetReader.name || 'Sheet';

      worksheetReader.on('row', (row) => {
        rowCount++;

        // 1行目: ヘッダーから列マッピングを検出
        if (rowCount === 1) {
          colMap = detectColumnsFromRow(row);
          // ヘッダー情報を記録
          row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
            const val = toStr(cell.value);
            if (val) {
              headerInfo[colNumToLetter(colNumber) + '(' + colNumber + ')'] = val;
            }
          });
          return;
        }

        // 2行目: サンプルとして記録
        if (rowCount === 2) {
          row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
            const val = toStr(cell.value);
            if (val) {
              sampleRow[colNumToLetter(colNumber) + '(' + colNumber + ')'] = val.substring(0, 50);
            }
          });
        }

        if (!colMap) return;

        // データ行の読み込み
        const leaseCompany = colMap.leaseCompany ? toStr(row.getCell(colMap.leaseCompany).value) : '';
        const branch       = colMap.branch       ? toStr(row.getCell(colMap.branch).value) : '';
        const ordererName  = colMap.ordererName  ? toStr(row.getCell(colMap.ordererName).value) : '';
        const customerName = colMap.customerName ? toStr(row.getCell(colMap.customerName).value) : '';
        const rep          = colMap.repLastName  ? toStr(row.getCell(colMap.repLastName).value) : '';
        const sales        = colMap.salesAmount  ? toNumber(row.getCell(colMap.salesAmount).value) : 0;
        const profit       = colMap.grossProfit  ? toNumber(row.getCell(colMap.grossProfit).value) : 0;

        const dateRaw = colMap.deliveryDate ? row.getCell(colMap.deliveryDate).value
                      : (colMap.slipDate    ? row.getCell(colMap.slipDate).value : null);
        const fiscalYear = getFiscalYear(dateRaw);
        const month = getMonth(dateRaw);

        // 空行スキップ
        if (!leaseCompany && !branch && !ordererName && !customerName && sales === 0) return;

        allRows.push({
          leaseCompany, branch, ordererName, customerName, rep,
          sales, profit, fiscalYear, month,
        });
      });
    });

    workbookReader.on('end', () => {
      resolve({
        rows: allRows,
        rowCount,
        sheetName,
        colMap,
        headerInfo,
        sampleRow,
      });
    });

    workbookReader.on('error', (err) => {
      reject(err);
    });

    workbookReader.read();
  });
}

// ===== メインの集計処理 =====
async function loadAndAggregate(dirPath) {
  const fileResult = getExcelFiles(dirPath);
  if (!fileResult.success) return fileResult;
  if (fileResult.files.length === 0) {
    return { success: false, error: 'Excelファイルが見つかりません' };
  }

  const allRows = [];
  const diagnostics = [];

  for (const fileInfo of fileResult.files) {
    const diag = {
      file: fileInfo.name,
      fileSize: `${(fileInfo.size / 1024 / 1024).toFixed(1)} MB`,
      sheets: [],
      error: null,
      openMethod: 'ExcelJS Streaming',
    };

    try {
      let result;
      try {
        console.log(`[Read] ストリーミング読み込み試行: ${fileInfo.name}`);
        result = await readExcelStreaming(fileInfo.path);
        diag.openMethod = 'ExcelJS Streaming';
        console.log(`[Read] ストリーミング成功: ${result.rows.length}行`);
      } catch (streamErr) {
        console.log(`[Read] ストリーミング失敗: ${streamErr.message}`);
        console.log(`[Read] パスワード付き読み込み試行: ${fileInfo.name}`);
        result = await readExcelWithPassword(fileInfo.path, EXCEL_PASSWORD);
        diag.openMethod = 'ExcelJS (パスワード解除)';
        console.log(`[Read] パスワード読み込み成功: ${result.rows.length}行`);
      }

      const sheetDiag = {
        name: result.sheetName,
        totalRows: result.rowCount,
        dataRowsFound: result.rows.length,
        headers: result.headerInfo,
        sampleRow2: result.sampleRow,
        detectedMapping: {},
        skip: null,
      };

      if (result.colMap) {
        for (const [key, colIdx] of Object.entries(result.colMap)) {
          sheetDiag.detectedMapping[key] = `${colNumToLetter(colIdx)}列(${colIdx})`;
        }
      }
      for (const key of Object.keys(HEADER_PATTERNS)) {
        if (!sheetDiag.detectedMapping[key]) {
          sheetDiag.detectedMapping[key] = '未検出';
        }
      }

      if (result.rows.length === 0 && result.rowCount <= 1) {
        sheetDiag.skip = `データ行なし（総行数: ${result.rowCount}）`;
      }

      diag.sheets.push(sheetDiag);

      for (const row of result.rows) {
        row.sourceFile = fileInfo.name;
        allRows.push(row);
      }

    } catch (err) {
      console.error(`[Read] 完全失敗: ${fileInfo.name} - ${err.message}`);
      diag.error = `読み込みエラー: ${err.message}`;
    }
    diagnostics.push(diag);
  }

  console.log(`[Load] 全ファイル読み込み完了: ${allRows.length}行, ${fileResult.files.length}ファイル`);

  // 年度・リース会社リスト
  const yearsSet = new Set();
  allRows.forEach(r => { if (r.fiscalYear) yearsSet.add(r.fiscalYear); });
  const years = [...yearsSet].sort();

  const leaseSet = new Set();
  allRows.forEach(r => { if (r.leaseCompany) leaseSet.add(r.leaseCompany); });
  const leaseCompanies = [...leaseSet].sort();

  return {
    success: true,
    data: {
      rows: allRows, years, leaseCompanies,
      fileCount: fileResult.files.length,
      totalRows: allRows.length,
      files: fileResult.files,
      diagnostics,
    },
  };
}

// ===== IPC ハンドラ =====
ipcMain.handle('load-excel-data', async (event, dirPath, forceRefresh) => {
  try {
    console.log(`[Load] パス: ${dirPath}, 強制更新: ${forceRefresh}`);

    if (!forceRefresh) {
      const cached = loadCache(dirPath);
      if (cached) {
        console.log(`[Load] キャッシュから返却（${cached.totalRows}件）`);
        return { success: true, data: { ...cached, fromCache: true } };
      }
    }

    console.log('[Load] Excelファイル読み込み開始...');
    const startTime = Date.now();
    const result = await loadAndAggregate(dirPath);
    console.log(`[Load] 読み込み完了（${Date.now() - startTime}ms）`);

    if (result.success && result.data) {
      saveCache(dirPath, result.data);
      result.data.fromCache = false;
    }

    return result;
  } catch (err) {
    return { success: false, error: `読み込み中にエラー: ${err.message}` };
  }
});

ipcMain.handle('check-path', async (event, dirPath) => {
  try { return { exists: fs.existsSync(dirPath), path: dirPath }; }
  catch { return { exists: false, path: dirPath }; }
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
