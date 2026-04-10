const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const officeCrypto = require('officecrypto-tool');
const crypto = require('crypto');
const zlib = require('zlib');
const os = require('os');

const EXCEL_PASSWORD = process.env.EXCEL_PASSWORD || '1962';
const CACHE_VERSION = 2; // 品番・商品名・受注数フィールド追加

function createExcelBackend({ cacheDir }) {
  const CACHE_DIR = cacheDir;

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
      if (!fs.existsSync(cachePath)) return null;

      const compressed = fs.readFileSync(cachePath);
      const json = zlib.gunzipSync(compressed).toString('utf8');
      const cache = JSON.parse(json);

      if (cache.version !== CACHE_VERSION) return null;

      const currentFingerprint = getFilesFingerprint(dirPath);
      if (cache.fingerprint !== currentFingerprint) return null;

      return cache.data;
    } catch {
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
          productCode: r.productCode,
          productName: r.productName,
          quantity: r.quantity,
        })),
        diagnostics: undefined,
      };

      const json = JSON.stringify({ version: CACHE_VERSION, fingerprint, dirPath, data: slim });
      const compressed = zlib.gzipSync(json, { level: 6 });
      fs.writeFileSync(cachePath, compressed);
    } catch (err) {
      console.error('[Cache] 保存失敗:', err.message);
    }
  }

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
    productCode:  { patterns: ['品番'], prefer_last: false },
    productName:  { patterns: ['商品名'], prefer_last: false },
    quantity:     { patterns: ['受注数量', '受注数', '数量'], prefer_last: false },
  };

  const OVERRIDE_COLUMNS = {
    leaseCompany: 34,
    branch:       35,
    productCode:  24, // X列
    productName:  25, // Y列
    quantity:     30, // AD列
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
              mapping[key] = colNumber;
            } else if (!mapping[key]) {
              mapping[key] = colNumber;
            }
            break;
          }
        }
      }
    });

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

  async function readExcelWithPassword(filePath, password) {
    const encryptedBuf = fs.readFileSync(filePath);
    const decryptedBuf = await officeCrypto.decrypt(encryptedBuf, { password });

    const tmpPath = path.join(os.tmpdir(), `_decrypted_${Date.now()}.xlsx`);
    fs.writeFileSync(tmpPath, decryptedBuf);

    try {
      const result = await readExcelStreaming(tmpPath);
      return result;
    } finally {
      try { fs.unlinkSync(tmpPath); } catch {}
    }
  }

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

          if (rowCount === 1) {
            colMap = detectColumnsFromRow(row);
            row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
              const val = toStr(cell.value);
              if (val) {
                headerInfo[colNumToLetter(colNumber) + '(' + colNumber + ')'] = val;
              }
            });
            return;
          }

          if (rowCount === 2) {
            row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
              const val = toStr(cell.value);
              if (val) {
                sampleRow[colNumToLetter(colNumber) + '(' + colNumber + ')'] = val.substring(0, 50);
              }
            });
          }

          if (!colMap) return;

          const leaseCompany = colMap.leaseCompany ? toStr(row.getCell(colMap.leaseCompany).value) : '';
          const branch       = colMap.branch       ? toStr(row.getCell(colMap.branch).value) : '';
          const ordererName  = colMap.ordererName  ? toStr(row.getCell(colMap.ordererName).value) : '';
          const customerName = colMap.customerName ? toStr(row.getCell(colMap.customerName).value) : '';
          const rep          = colMap.repLastName  ? toStr(row.getCell(colMap.repLastName).value) : '';
          const sales        = colMap.salesAmount  ? toNumber(row.getCell(colMap.salesAmount).value) : 0;
          const profit       = colMap.grossProfit  ? toNumber(row.getCell(colMap.grossProfit).value) : 0;
          const productCode  = colMap.productCode  ? toStr(row.getCell(colMap.productCode).value) : '';
          const productName  = colMap.productName  ? toStr(row.getCell(colMap.productName).value) : '';
          const quantity     = colMap.quantity     ? toNumber(row.getCell(colMap.quantity).value) : 0;

          const dateRaw = colMap.deliveryDate ? row.getCell(colMap.deliveryDate).value
                        : (colMap.slipDate    ? row.getCell(colMap.slipDate).value : null);
          const fiscalYear = getFiscalYear(dateRaw);
          const month = getMonth(dateRaw);

          if (!leaseCompany && !branch && !ordererName && !customerName && sales === 0) return;

          allRows.push({
            leaseCompany, branch, ordererName, customerName, rep,
            sales, profit, fiscalYear, month,
            productCode, productName, quantity,
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
          result = await readExcelStreaming(fileInfo.path);
          diag.openMethod = 'ExcelJS Streaming';
        } catch {
          result = await readExcelWithPassword(fileInfo.path, EXCEL_PASSWORD);
          diag.openMethod = 'ExcelJS (パスワード解除)';
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
        diag.error = `読み込みエラー: ${err.message}`;
      }
      diagnostics.push(diag);
    }

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

  async function loadExcelData(dirPath, forceRefresh) {
    if (!forceRefresh) {
      const cached = loadCache(dirPath);
      if (cached) {
        return { success: true, data: { ...cached, fromCache: true } };
      }
    }

    const result = await loadAndAggregate(dirPath);

    if (result.success && result.data) {
      saveCache(dirPath, result.data);
      result.data.fromCache = false;
    }

    return result;
  }

  function checkPath(dirPath) {
    try {
      return { exists: fs.existsSync(dirPath), path: dirPath };
    } catch {
      return { exists: false, path: dirPath };
    }
  }

  return { loadExcelData, checkPath, loadCache, saveCache };
}

module.exports = { createExcelBackend };
