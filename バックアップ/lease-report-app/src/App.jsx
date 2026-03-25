import React, { useState, useMemo, useCallback, useEffect, useRef } from 'react';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
} from 'recharts';
import {
  ChevronRight, Building2, Settings, User, Store,
  ArrowUpRight, ArrowDownRight, LayoutDashboard, Database,
  Calendar, FolderOpen, RefreshCcw, CheckCircle2, FileText, FileSpreadsheet,
  ListFilter, AlertCircle, Loader2, XCircle, ChevronLeft, ArrowUpDown, Eye, EyeOff,
} from 'lucide-react';
import {
  filterRows, aggregateByBranch, aggregateByOrderer, aggregateByCustomer,
  generatePivotData, calcYoY, calcMargin, formatCurrency, formatCurrencyFull,
  generateCsvContent,
} from './utils/aggregator';

// Electron API (preloadで公開)
const api = window.electronAPI || null;

// デフォルトパス
const DEFAULT_PATH = '\\\\192.1.1.103\\share\\特販共有\\見積\\売上データ';

const App = () => {
  // ===== State =====
  const [networkPath, setNetworkPath] = useState(DEFAULT_PATH);
  const [rawData, setRawData] = useState(null); // { rows, years, leaseCompanies, fileCount, totalRows }
  const [isLoading, setIsLoading] = useState(false);
  const [loadError, setLoadError] = useState(null);
  const [connectionStatus, setConnectionStatus] = useState('idle'); // idle | online | offline | loading

  const [selectedLeaseCo, setSelectedLeaseCo] = useState('ALL');
  const [monthRange, setMonthRange] = useState({ start: '4', end: '3' });
  const [viewMode, setViewMode] = useState('dashboard'); // dashboard | pivot_report
  const [activeView, setActiveView] = useState({ level: 'branch', branchName: null, ordererName: null });
  const [isConfigOpen, setIsConfigOpen] = useState(false);
  const [amountUnit, setAmountUnit] = useState('yen'); // 'yen' | 'thousand'
  const [pivotBranch, setPivotBranch] = useState('ALL');
  const [pivotSort, setPivotSort] = useState('sales'); // 'sales' | 'orderer' | 'customer'
  const [showProfit, setShowProfit] = useState(true);

  const printRef = useRef(null);

  // 金額フォーマット（単位切替対応）
  const fmtAmt = useCallback((val) => {
    if (amountUnit === 'thousand') {
      const v = Math.round((val || 0) / 1000);
      return `¥${v.toLocaleString()}`;
    }
    return formatCurrencyFull(val);
  }, [amountUnit]);

  const fmtAmtShort = useCallback((val) => {
    if (amountUnit === 'thousand') {
      const v = (val || 0) / 1000;
      if (Math.abs(v) >= 100000) return `¥${(v / 10000).toFixed(0)}万`;
      return `¥${Math.round(v).toLocaleString()}`;
    }
    return formatCurrency(val);
  }, [amountUnit]);

  // ===== データ読み込み =====
  const loadData = useCallback(async (forceRefresh = false) => {
    if (!api) {
      loadMockData();
      return;
    }
    setIsLoading(true);
    setLoadError(null);
    setConnectionStatus('loading');
    try {
      const result = await api.loadExcelData(networkPath, forceRefresh);
      if (result.success) {
        setRawData(result.data);
        setConnectionStatus('online');
        setLoadError(null);
      } else {
        setLoadError(result.error);
        setConnectionStatus('offline');
      }
    } catch (err) {
      setLoadError(err.message);
      setConnectionStatus('offline');
    } finally {
      setIsLoading(false);
    }
  }, [networkPath]);

  // モックデータ（開発・デモ用）
  const loadMockData = useCallback(() => {
    setIsLoading(true);
    setConnectionStatus('loading');
    setTimeout(() => {
      const years = [2023, 2024, 2025];
      const mockRows = [];
      const branches = ['MA営業第4部', '首都圏営業第2部', '中日本営業部', '西日本営業部'];
      const leases = ['SMAS', 'TCM', 'オリックス', '三菱HC'];
      const orderers = ['相馬 リース課', '福岡 営業所', '東京本部', '大阪支店', '名古屋営業所'];
      const customers = ['(株)西原環境', '山下商事(有)', '日本カーソリューションズ', 'トヨタモビリティパーツ', '(株)丸紅エネルギー', '三菱商事(株)', 'アイシン精機(株)', '(株)豊田自動織機'];
      const reps = ['土岐 暁治', '犬塚 龍', '山本 太郎', '佐藤 花子', '田中 一郎'];

      for (let i = 0; i < 500; i++) {
        const fy = years[Math.floor(Math.random() * years.length)];
        const month = Math.floor(Math.random() * 12) + 1;
        const sales = Math.floor(Math.random() * 500000) + 50000;
        const profitRate = 0.15 + Math.random() * 0.2;
        mockRows.push({
          leaseCompany: leases[Math.floor(Math.random() * leases.length)],
          branch: branches[Math.floor(Math.random() * branches.length)],
          ordererName: orderers[Math.floor(Math.random() * orderers.length)],
          customerName: customers[Math.floor(Math.random() * customers.length)],
          rep: reps[Math.floor(Math.random() * reps.length)],
          sales,
          profit: Math.floor(sales * profitRate),
          fiscalYear: fy,
          month,
          sourceFile: 'mock_data.xlsx',
        });
      }

      setRawData({
        rows: mockRows,
        years,
        leaseCompanies: leases,
        fileCount: 3,
        totalRows: mockRows.length,
      });
      setConnectionStatus('online');
      setIsLoading(false);
    }, 1200);
  }, []);

  // 初回読み込み
  useEffect(() => {
    loadData();
  }, []);

  // データ読み込み後に最新月をデフォルト終了月に設定（会計年度順: 4→...→12→1→2→3）
  useEffect(() => {
    if (!rawData || !rawData.rows.length) return;
    const maxFY = Math.max(...rawData.years);
    const latestMonths = rawData.rows
      .filter(r => r.fiscalYear === maxFY)
      .map(r => r.month);
    if (latestMonths.length > 0) {
      const toFiscalPos = m => (m - 4 + 12) % 12;
      const latest = latestMonths.reduce((best, m) => toFiscalPos(m) > toFiscalPos(best) ? m : best);
      setMonthRange({ start: '4', end: String(latest) });
    }
  }, [rawData]);

  // ===== フィルタ済みデータ =====
  const filteredRows = useMemo(() => {
    if (!rawData) return [];
    return filterRows(rawData.rows, {
      leaseCompany: selectedLeaseCo,
      startMonth: monthRange.start,
      endMonth: monthRange.end,
    });
  }, [rawData, selectedLeaseCo, monthRange]);

  const years = useMemo(() => rawData?.years || [], [rawData]);
  const leaseCompanies = useMemo(() => rawData?.leaseCompanies || [], [rawData]);

  const branches = useMemo(() => {
    if (!filteredRows.length) return [];
    const set = new Set();
    filteredRows.forEach(r => { if (r.branch) set.add(r.branch); });
    return [...set].sort();
  }, [filteredRows]);

  // ===== ドリルダウンデータ =====
  const currentTableData = useMemo(() => {
    if (!filteredRows.length || !years.length) return [];
    switch (activeView.level) {
      case 'branch':
        return aggregateByBranch(filteredRows, years);
      case 'orderer':
        return aggregateByOrderer(filteredRows, years, activeView.branchName);
      case 'customer':
        return aggregateByCustomer(filteredRows, years, activeView.branchName, activeView.ordererName);
      default:
        return [];
    }
  }, [filteredRows, years, activeView]);

  // ===== ピボットデータ =====
  const pivotData = useMemo(() => {
    if (!filteredRows.length || !years.length) return [];
    const rows = pivotBranch === 'ALL'
      ? filteredRows
      : filteredRows.filter(r => r.branch === pivotBranch);
    const data = generatePivotData(rows, years);
    const latestYear = years[years.length - 1];

    if (pivotSort === 'sales') {
      data.sort((a, b) => (b.sales[latestYear] || 0) - (a.sales[latestYear] || 0));
    } else if (pivotSort === 'orderer') {
      const ordererTotals = {};
      data.forEach(r => {
        const key = r.orderer || '';
        ordererTotals[key] = (ordererTotals[key] || 0) + (r.sales[latestYear] || 0);
      });
      data.sort((a, b) => (ordererTotals[b.orderer || ''] || 0) - (ordererTotals[a.orderer || ''] || 0)
        || (b.sales[latestYear] || 0) - (a.sales[latestYear] || 0));
    } else if (pivotSort === 'customer') {
      data.sort((a, b) => (b.sales[latestYear] || 0) - (a.sales[latestYear] || 0));
      const customerTotals = {};
      data.forEach(r => {
        const key = r.customer || '';
        customerTotals[key] = (customerTotals[key] || 0) + (r.sales[latestYear] || 0);
      });
      data.sort((a, b) => (customerTotals[b.customer || ''] || 0) - (customerTotals[a.customer || ''] || 0));
    }
    return data;
  }, [filteredRows, years, pivotBranch, pivotSort]);

  // ===== チャートデータ =====
  const chartData = useMemo(() => {
    if (!currentTableData.length || !years.length) return [];
    const latestYear = years[years.length - 1];
    return currentTableData.slice(0, 8).map(item => ({
      name: item.name.length > 10 ? item.name.slice(0, 10) + '…' : item.name,
      fullName: item.name,
      [`${latestYear}年 売上`]: item.sales[latestYear] || 0,
      [`${latestYear}年 粗利`]: item.profit[latestYear] || 0,
    }));
  }, [currentTableData, years]);

  // ===== ハンドラ =====
  const handleDrillDown = (item) => {
    if (activeView.level === 'branch') {
      setActiveView({ level: 'orderer', branchName: item.name, ordererName: null });
    } else if (activeView.level === 'orderer') {
      setActiveView({ ...activeView, level: 'customer', ordererName: item.name });
    }
  };

  const handleBreadcrumb = (level) => {
    if (level === 'branch') {
      setActiveView({ level: 'branch', branchName: null, ordererName: null });
    } else if (level === 'orderer') {
      setActiveView({ ...activeView, level: 'orderer', ordererName: null });
    }
  };

  const handleRefresh = () => {
    loadData(true); // 強制リフレッシュ（キャッシュ無視）
  };

  const handleSaveCsv = async () => {
    if (!pivotData.length) return;
    const csv = generateCsvContent(pivotData, years);
    if (api) {
      const result = await api.saveCsv(csv);
      if (result.success) {
        alert(`CSVを保存しました: ${result.path}`);
      }
    } else {
      // ブラウザ用フォールバック
      const bom = '\uFEFF';
      const blob = new Blob([bom + csv], { type: 'text/csv;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `リース実績レポート_${new Date().toISOString().slice(0, 10)}.csv`;
      a.click();
      URL.revokeObjectURL(url);
    }
  };

  const handleSavePdf = async () => {
    const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const lease = selectedLeaseCo !== 'ALL' ? selectedLeaseCo : '';
    let label = '';
    if (viewMode === 'pivot_report') {
      const branch = pivotBranch !== 'ALL' ? pivotBranch : '';
      label = [lease, branch].filter(Boolean).join('_');
      label = `架装品実績_一括網羅${label ? '_' + label : ''}_${date}`;
    } else {
      const levelMap = { branch: '担当別', orderer: '担当別', customer: '顧客別' };
      const typeName = levelMap[activeView.level] || '担当別';
      const branch = activeView.branchName || '';
      label = [lease, branch].filter(Boolean).join('_');
      label = `架装品実績_${typeName}${label ? '_' + label : ''}_${date}`;
    }
    if (api) {
      await api.savePdf(label);
    } else {
      window.print();
    }
  };

  const handleConfigSave = () => {
    setIsConfigOpen(false);
    loadData();
  };

  const handleSelectFolder = async () => {
    if (!api) return;
    const result = await api.selectFolder();
    if (result.success) {
      setNetworkPath(result.path);
    }
  };

  // ===== レンダリング =====
  const levelLabels = {
    branch: { icon: Building2, label: '部店名', title: '部店別 年次実績比較' },
    orderer: { icon: Store, label: '注文者名', title: `${activeView.branchName || ''} 内 注文者実績` },
    customer: { icon: User, label: '顧客名', title: `${activeView.ordererName || ''} 内 顧客実績` },
  };

  const currentLevelInfo = levelLabels[activeView.level];
  const LevelIcon = currentLevelInfo.icon;

  return (
    <div className="min-h-screen h-screen bg-slate-50 font-sans text-slate-900 flex select-none">
      {/* ===== Sidebar ===== */}
      <aside className="w-64 bg-slate-900 text-white flex flex-col p-6 sticky top-0 h-screen shadow-2xl z-20 no-print flex-shrink-0">
        <div className="flex items-center gap-3 mb-10 px-2">
          <div className="bg-blue-500 p-2 rounded-xl text-white shadow-lg shadow-blue-500/20">
            <Database size={24} />
          </div>
          <h1 className="text-lg font-black leading-none tracking-tighter uppercase">
            Special<br/>Sales Report
          </h1>
        </div>

        <nav className="flex-1 space-y-2">
          <button
            onClick={() => setViewMode('dashboard')}
            className={`flex items-center gap-3 w-full p-4 rounded-2xl transition-all duration-300 ${
              viewMode === 'dashboard'
                ? 'bg-blue-600 text-white shadow-lg shadow-blue-600/30 scale-105'
                : 'text-slate-400 hover:bg-slate-800'
            }`}
          >
            <LayoutDashboard size={20} />
            <span className="font-black text-sm tracking-tight">分析ダッシュボード</span>
          </button>
          <button
            onClick={() => setViewMode('pivot_report')}
            className={`flex items-center gap-3 w-full p-4 rounded-2xl transition-all duration-300 ${
              viewMode === 'pivot_report'
                ? 'bg-blue-600 text-white shadow-lg shadow-blue-600/30 scale-105'
                : 'text-slate-400 hover:bg-slate-800'
            }`}
          >
            <ListFilter size={20} />
            <span className="font-black text-sm tracking-tight">一括網羅レポート</span>
          </button>

          <div className="pt-8 pb-2 px-4 text-[10px] font-black text-slate-500 uppercase tracking-widest">
            Reports Export
          </div>
          <button
            onClick={handleSavePdf}
            className="flex items-center justify-between w-full p-4 rounded-2xl text-slate-400 hover:bg-slate-800 transition-all border border-transparent hover:border-slate-700"
          >
            <div className="flex items-center gap-3 italic font-bold text-sm text-slate-200">
              <FileText size={18} /> PDF Export
            </div>
          </button>
          <button
            onClick={handleSaveCsv}
            className="flex items-center justify-between w-full p-4 rounded-2xl text-slate-400 hover:bg-slate-800 transition-all border border-transparent hover:border-slate-700"
          >
            <div className="flex items-center gap-3 italic font-bold text-sm text-slate-200">
              <FileSpreadsheet size={18} /> CSV Export
            </div>
          </button>
        </nav>

        {/* Stats */}
        {rawData && (
          <div className="mb-4 p-3 bg-slate-800 rounded-2xl text-xs space-y-1">
            <div className="flex justify-between text-slate-400">
              <span>ファイル数</span>
              <span className="font-mono font-bold text-slate-200">{rawData.fileCount}</span>
            </div>
            <div className="flex justify-between text-slate-400">
              <span>レコード数</span>
              <span className="font-mono font-bold text-slate-200">{rawData.totalRows.toLocaleString()}</span>
            </div>
            <div className="flex justify-between text-slate-400">
              <span>読込</span>
              <span className={`font-mono font-bold ${rawData.fromCache ? 'text-emerald-400' : 'text-blue-400'}`}>
                {rawData.fromCache ? 'Cache' : 'Excel'}
              </span>
            </div>
          </div>
        )}

        <button
          onClick={() => setIsConfigOpen(true)}
          className="flex items-center gap-3 w-full p-4 rounded-2xl text-slate-400 hover:bg-slate-800 transition-all mt-auto border border-slate-800"
        >
          <Settings size={20} />
          <span className="font-bold text-sm">システム設定</span>
        </button>
      </aside>

      {/* ===== Main Content ===== */}
      <main className="flex-1 p-8 overflow-y-auto custom-scrollbar" ref={printRef}>
        {/* Header */}
        <header className="mb-10 space-y-8 no-print">
          <div className="flex justify-between items-end">
            <div>
              <div className="flex items-center gap-4 mb-2">
                <h2 className="text-4xl font-black text-slate-800 tracking-tighter">
                  特販部リース会社実績レポート
                </h2>
                <ConnectionBadge status={connectionStatus} />
              </div>
              <div className="flex items-center gap-2 text-slate-400 font-bold text-sm bg-slate-100 w-fit px-3 py-1 rounded-lg">
                <FolderOpen size={14} /> {networkPath}
              </div>
            </div>

            <button
              onClick={handleRefresh}
              disabled={isLoading}
              className={`group flex items-center gap-2 px-8 py-4 rounded-3xl bg-slate-900 text-white font-black text-sm shadow-2xl hover:bg-blue-600 transition-all duration-300 active:scale-95 disabled:opacity-50 ${
                isLoading ? 'animate-pulse' : ''
              }`}
            >
              <RefreshCcw
                size={18}
                className={isLoading ? 'animate-spin' : 'group-hover:rotate-180 transition-transform duration-500'}
              />
              {isLoading ? 'データを読み込んでいます...' : '最新データに更新'}
            </button>
          </div>

          {/* Filters Bar */}
          <div className="bg-white p-8 rounded-[3rem] shadow-sm border border-slate-100 flex flex-wrap gap-12 items-center">
            <div className="space-y-3">
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-[0.2em] ml-2">
                リース会社選択
              </label>
              <div className="flex gap-2 flex-wrap">
                <button
                  onClick={() => setSelectedLeaseCo('ALL')}
                  className={`px-6 py-2.5 rounded-2xl text-xs font-black transition-all duration-300 ${
                    selectedLeaseCo === 'ALL'
                      ? 'bg-blue-600 text-white shadow-xl shadow-blue-200'
                      : 'bg-slate-100 text-slate-400 hover:bg-slate-200'
                  }`}
                >
                  すべて
                </button>
                {leaseCompanies.map(l => (
                  <button
                    key={l}
                    onClick={() => setSelectedLeaseCo(l)}
                    className={`px-6 py-2.5 rounded-2xl text-xs font-black transition-all duration-300 ${
                      selectedLeaseCo === l
                        ? 'bg-blue-600 text-white shadow-xl shadow-blue-200'
                        : 'bg-slate-100 text-slate-400 hover:bg-slate-200'
                    }`}
                  >
                    {l}
                  </button>
                ))}
              </div>
            </div>

            <div className="space-y-3">
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-[0.2em] ml-2">
                期間指定
              </label>
              <div className="flex items-center gap-4 bg-slate-100 p-2 rounded-2xl">
                <select
                  value={monthRange.start}
                  onChange={e => setMonthRange(prev => ({ ...prev, start: e.target.value }))}
                  className="bg-transparent border-none text-sm font-black px-4 py-1.5 focus:ring-0 text-slate-700 cursor-pointer"
                >
                  {[...Array(12)].map((_, i) => (
                    <option key={i + 1} value={i + 1}>{i + 1}月</option>
                  ))}
                </select>
                <div className="w-4 h-0.5 bg-slate-300 rounded-full" />
                <select
                  value={monthRange.end}
                  onChange={e => setMonthRange(prev => ({ ...prev, end: e.target.value }))}
                  className="bg-transparent border-none text-sm font-black px-4 py-1.5 focus:ring-0 text-slate-700 cursor-pointer"
                >
                  {[...Array(12)].map((_, i) => (
                    <option key={i + 1} value={i + 1}>{i + 1}月</option>
                  ))}
                </select>
              </div>
            </div>

            <div className="space-y-3">
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-[0.2em] ml-2">
                金額単位
              </label>
              <div className="flex bg-slate-100 p-1 rounded-2xl">
                <button
                  onClick={() => setAmountUnit('yen')}
                  className={`px-5 py-2.5 rounded-xl text-xs font-black transition-all duration-300 ${
                    amountUnit === 'yen'
                      ? 'bg-white text-slate-800 shadow-md'
                      : 'text-slate-400 hover:text-slate-600'
                  }`}
                >
                  円
                </button>
                <button
                  onClick={() => setAmountUnit('thousand')}
                  className={`px-5 py-2.5 rounded-xl text-xs font-black transition-all duration-300 ${
                    amountUnit === 'thousand'
                      ? 'bg-white text-slate-800 shadow-md'
                      : 'text-slate-400 hover:text-slate-600'
                  }`}
                >
                  千円
                </button>
              </div>
            </div>

            <div className="space-y-3">
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-[0.2em] ml-2">
                粗利表示
              </label>
              <button
                onClick={() => setShowProfit(p => !p)}
                className={`flex items-center gap-2 px-5 py-2.5 rounded-2xl text-xs font-black transition-all duration-300 ${
                  showProfit
                    ? 'bg-emerald-500 text-white shadow-lg shadow-emerald-200'
                    : 'bg-slate-100 text-slate-400 hover:bg-slate-200'
                }`}
              >
                {showProfit ? <Eye size={14} /> : <EyeOff size={14} />}
                {showProfit ? 'ON' : 'OFF'}
              </button>
            </div>
          </div>
        </header>

        {/* Loading State */}
        {isLoading && !rawData && (
          <LoadingScreen />
        )}

        {/* Error State */}
        {loadError && (
          <div className="bg-rose-50 border border-rose-200 rounded-3xl p-8 mb-8 flex items-center gap-4 animate-fade-in">
            <AlertCircle className="text-rose-500 flex-shrink-0" size={24} />
            <div>
              <h3 className="font-black text-rose-800 mb-1">データ読み込みエラー</h3>
              <p className="text-sm text-rose-600">{loadError}</p>
              <p className="text-xs text-rose-400 mt-1">設定画面からパスを確認するか、ネットワーク接続を確認してください。</p>
            </div>
          </div>
        )}

        {/* Cache / Data Status */}
        {rawData && !isLoading && rawData.fromCache && (
          <div className="bg-emerald-50 border border-emerald-200 rounded-3xl px-8 py-5 mb-8 flex items-center justify-between animate-fade-in no-print">
            <div className="flex items-center gap-3">
              <CheckCircle2 className="text-emerald-500" size={20} />
              <span className="font-bold text-emerald-800 text-sm">
                キャッシュから高速読込（{rawData.totalRows.toLocaleString()}件）
              </span>
              <span className="text-xs text-emerald-500 font-bold">Excelファイルに変更がある場合は自動で再解析されます</span>
            </div>
            <button
              onClick={handleRefresh}
              className="flex items-center gap-2 px-5 py-2 bg-emerald-600 text-white rounded-2xl text-xs font-black hover:bg-emerald-700 transition-all active:scale-95"
            >
              <RefreshCcw size={14} />
              強制更新
            </button>
          </div>
        )}

        {/* Main Content Area */}
        {rawData && !isLoading && (
          <>
            {viewMode === 'dashboard' ? (
              <DashboardView
                data={currentTableData}
                chartData={chartData}
                years={years}
                activeView={activeView}
                levelInfo={currentLevelInfo}
                LevelIcon={LevelIcon}
                onDrillDown={handleDrillDown}
                onBreadcrumb={handleBreadcrumb}
                onSavePdf={handleSavePdf}
                onSaveCsv={handleSaveCsv}
                fmtAmt={fmtAmt}
                fmtAmtShort={fmtAmtShort}
                amountUnit={amountUnit}
                showProfit={showProfit}
                selectedLeaseCo={selectedLeaseCo}
              />
            ) : (
              <PivotView
                data={pivotData}
                years={years}
                branches={branches}
                pivotBranch={pivotBranch}
                onBranchChange={setPivotBranch}
                pivotSort={pivotSort}
                onSortChange={setPivotSort}
                onSavePdf={handleSavePdf}
                onSaveCsv={handleSaveCsv}
                fmtAmt={fmtAmt}
                amountUnit={amountUnit}
                showProfit={showProfit}
                selectedLeaseCo={selectedLeaseCo}
              />
            )}
          </>
        )}
      </main>

      {/* ===== Settings Modal ===== */}
      {isConfigOpen && (
        <SettingsModal
          networkPath={networkPath}
          onPathChange={setNetworkPath}
          onSelectFolder={handleSelectFolder}
          onClose={() => setIsConfigOpen(false)}
          onSave={handleConfigSave}
        />
      )}
    </div>
  );
};

// ===== 接続ステータスバッジ =====
const ConnectionBadge = ({ status }) => {
  const styles = {
    idle: 'bg-slate-100 text-slate-500 border-slate-200',
    loading: 'bg-amber-50 text-amber-600 border-amber-100 animate-pulse',
    online: 'bg-blue-50 text-blue-600 border-blue-100 animate-pulse-glow',
    offline: 'bg-rose-50 text-rose-600 border-rose-100',
  };
  const labels = {
    idle: 'Standby',
    loading: 'Loading...',
    online: 'Live Sync',
    offline: 'Offline',
  };
  const icons = {
    idle: CheckCircle2,
    loading: Loader2,
    online: CheckCircle2,
    offline: AlertCircle,
  };
  const Icon = icons[status];
  return (
    <div className={`flex items-center gap-1.5 px-4 py-1.5 rounded-full text-[10px] font-black uppercase border shadow-sm ${styles[status]}`}>
      <Icon size={12} className={status === 'loading' ? 'animate-spin' : ''} /> {labels[status]}
    </div>
  );
};

// ===== ローディング画面 =====
const LoadingScreen = () => (
  <div className="flex flex-col items-center justify-center h-96 animate-fade-in">
    <div className="relative mb-8">
      <div className="w-20 h-20 border-4 border-slate-200 rounded-full" />
      <div className="absolute top-0 left-0 w-20 h-20 border-4 border-blue-500 border-t-transparent rounded-full animate-spin" />
      <Database className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 text-blue-500" size={28} />
    </div>
    <h3 className="text-xl font-black text-slate-700 mb-2">データを読み込んでいます</h3>
    <p className="text-sm text-slate-400 font-bold">ネットワークフォルダからExcelファイルを解析中...</p>
    <div className="flex gap-1 mt-6">
      {[0, 1, 2].map(i => (
        <div
          key={i}
          className="w-2 h-2 bg-blue-500 rounded-full animate-bounce"
          style={{ animationDelay: `${i * 0.15}s` }}
        />
      ))}
    </div>
  </div>
);

// ===== ダッシュボードビュー =====
const DashboardView = ({ data, chartData, years, activeView, levelInfo, LevelIcon, onDrillDown, onBreadcrumb, onSavePdf, onSaveCsv, fmtAmt, fmtAmtShort, amountUnit, showProfit, selectedLeaseCo }) => {
  const totalRow = useMemo(() => {
    if (activeView.level !== 'branch' || !data.length || !years.length) return null;
    const sales = {};
    const profit = {};
    years.forEach(y => { sales[y] = 0; profit[y] = 0; });
    data.forEach(item => {
      years.forEach(y => {
        sales[y] += item.sales[y] || 0;
        profit[y] += item.profit[y] || 0;
      });
    });
    return { name: '全部店 合計', sales, profit };
  }, [data, years, activeView.level]);

  return (
  <div className="space-y-8 animate-fade-in-up">
    {/* Breadcrumb */}
    <div className="flex items-center gap-2 text-sm font-bold no-print">
      <button
        onClick={() => onBreadcrumb('branch')}
        className={`flex items-center gap-1 transition-colors ${
          activeView.level === 'branch' ? 'text-blue-600' : 'text-slate-400 hover:text-slate-600'
        }`}
      >
        <Building2 size={16} /> 部店一覧
      </button>
      {activeView.level !== 'branch' && (
        <>
          <ChevronRight size={14} className="text-slate-300" />
          <button
            onClick={() => onBreadcrumb('orderer')}
            className={`flex items-center gap-1 transition-colors ${
              activeView.level === 'orderer' ? 'text-blue-600' : 'text-slate-400 hover:text-slate-600'
            }`}
          >
            <Store size={16} /> {activeView.branchName}
          </button>
        </>
      )}
      {activeView.level === 'customer' && (
        <>
          <ChevronRight size={14} className="text-slate-300" />
          <span className="text-blue-600 flex items-center gap-1">
            <User size={16} /> {activeView.ordererName}
          </span>
        </>
      )}
    </div>

    {/* Comparison Table */}
    <div className="bg-white rounded-[3rem] shadow-sm border border-slate-100 overflow-hidden">
      <div className="p-8 border-b border-slate-50 flex justify-between items-center bg-slate-50/30">
        <h3 className="font-black text-slate-800 text-xl flex items-center gap-3">
          <LevelIcon className="text-blue-500" />
          {levelInfo.title}
          {selectedLeaseCo !== 'ALL' && <span className="text-blue-500 text-base">— {selectedLeaseCo}</span>}
        </h3>
        <div className="flex items-center gap-4">
          <div className="text-xs font-bold text-slate-400 flex items-center gap-1.5">
            <Calendar size={14} />
            出力日: {new Date().toLocaleDateString('ja-JP', { year: 'numeric', month: '2-digit', day: '2-digit' })}
          </div>
        <div className="flex gap-2 no-print">
          <button
            onClick={onSavePdf}
            className="flex items-center gap-2 px-4 py-2 bg-white rounded-xl border border-slate-200 text-slate-600 hover:text-blue-500 hover:border-blue-200 transition-all shadow-sm text-sm font-bold"
          >
            <FileText size={16} /> PDF
          </button>
          <button
            onClick={onSaveCsv}
            className="flex items-center gap-2 px-4 py-2 bg-white rounded-xl border border-slate-200 text-slate-600 hover:text-emerald-500 hover:border-emerald-200 transition-all shadow-sm text-sm font-bold"
          >
            <FileSpreadsheet size={16} /> CSV
          </button>
        </div>
        </div>
      </div>

      <div className="overflow-x-auto">
        <table className="w-full text-left border-collapse">
          <thead>
            <tr className="bg-slate-900 text-xl font-black text-white tracking-wide text-center">
              <th className="px-8 py-4 min-w-[200px] text-center">
                {levelInfo.label}
                {amountUnit === 'thousand' && <span className="ml-2 text-amber-400 normal-case tracking-normal text-xs">（単位：千円）</span>}
              </th>
              {years.map(year => (
                <th key={year} className="px-6 py-4 text-center border-l border-slate-800">
                  {year}年度
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {totalRow && (
              <tr className="bg-blue-50/50 border-b-2 border-blue-200">
                <td className="px-8 py-5">
                  <div className="font-black text-blue-700 text-lg">{totalRow.name}</div>
                </td>
                {years.map((year, yIdx) => {
                  const sales = totalRow.sales[year] || 0;
                  const profit = totalRow.profit[year] || 0;
                  const margin = calcMargin(profit, sales);
                  const prevYear = years[yIdx - 1];
                  const yoy = prevYear ? calcYoY(sales, totalRow.sales[prevYear]) : null;
                  return (
                    <td key={year} className="px-6 py-5 border-l border-blue-200">
                      <div className="space-y-3">
                        <div className="flex justify-between items-baseline">
                          <span className="text-xs font-black text-blue-500">売上</span>
                          <div className="text-right">
                            <div className="font-mono font-black text-blue-800">{fmtAmt(sales)}</div>
                            {yoy !== null && (
                              <div className={`text-xs font-black flex items-center justify-end gap-1 ${parseFloat(yoy) >= 0 ? 'text-emerald-500' : 'text-rose-500'}`}>
                                {parseFloat(yoy) >= 0 ? <ArrowUpRight size={12} /> : <ArrowDownRight size={12} />}
                                {yoy}%
                              </div>
                            )}
                          </div>
                        </div>
                        {showProfit && (
                          <div className="flex justify-between items-baseline">
                            <span className="text-xs font-black text-blue-500">粗利(率)</span>
                            <div className="text-right">
                              <div className="font-mono font-black text-emerald-600">{fmtAmt(profit)}</div>
                              <div className="px-2 py-0.5 bg-emerald-50 text-emerald-700 rounded text-[11px] font-black mt-0.5 w-fit ml-auto border border-emerald-100">{margin}%</div>
                            </div>
                          </div>
                        )}
                      </div>
                    </td>
                  );
                })}
              </tr>
            )}
            {data.length === 0 ? (
              <tr>
                <td colSpan={years.length + 1} className="px-8 py-16 text-center text-slate-300 italic">
                  該当するデータがありません
                </td>
              </tr>
            ) : (
              data.map((item, idx) => (
                <tr
                  key={idx}
                  className={`group hover:bg-blue-50/30 transition-all ${
                    activeView.level !== 'customer' ? 'cursor-pointer' : ''
                  }`}
                  onClick={() => activeView.level !== 'customer' && onDrillDown(item)}
                >
                  <td className="px-8 py-6">
                    <div className="font-black text-slate-800 text-lg group-hover:text-blue-600 transition-colors flex items-center gap-2">
                      {item.name}
                      {activeView.level !== 'customer' && (
                        <ChevronRight
                          size={14}
                          className="opacity-0 group-hover:opacity-100 transition-all -translate-x-2 group-hover:translate-x-0"
                        />
                      )}
                    </div>
                    {activeView.level !== 'customer' && (
                      <div className="text-[10px] font-bold text-slate-400 mt-1 uppercase tracking-tighter no-print">
                        クリックで詳細を表示
                      </div>
                    )}
                  </td>
                  {years.map((year, yIdx) => {
                    const sales = item.sales[year] || 0;
                    const profit = item.profit[year] || 0;
                    const margin = calcMargin(profit, sales);
                    const prevYear = years[yIdx - 1];
                    const yoy = prevYear ? calcYoY(sales, item.sales[prevYear]) : null;

                    return (
                      <td key={year} className="px-6 py-6 border-l border-slate-300 group-hover:bg-white/50">
                        <div className="space-y-3">
                          <div className="flex justify-between items-baseline">
                            <span className="text-xs font-black text-slate-600">売上</span>
                            <div className="text-right">
                              <div className="font-mono font-black text-slate-700">
                                {fmtAmt(sales)}
                              </div>
                              {yoy !== null && (
                                <div className={`text-xs font-black flex items-center justify-end gap-1 ${
                                  parseFloat(yoy) >= 0 ? 'text-emerald-500' : 'text-rose-500'
                                }`}>
                                  {parseFloat(yoy) >= 0 ? <ArrowUpRight size={12} /> : <ArrowDownRight size={12} />}
                                  {yoy}%
                                </div>
                              )}
                            </div>
                          </div>
                          {showProfit && (
                            <div className="flex justify-between items-baseline">
                              <span className="text-xs font-black text-slate-600">粗利(率)</span>
                              <div className="text-right">
                                <div className="font-mono font-black text-emerald-600">
                                  {fmtAmt(profit)}
                                </div>
                                <div className="px-2 py-0.5 bg-emerald-50 text-emerald-700 rounded text-[11px] font-black mt-0.5 w-fit ml-auto border border-emerald-100">
                                  {margin}%
                                </div>
                              </div>
                            </div>
                          )}
                        </div>
                      </td>
                    );
                  })}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </div>

    {/* Chart */}
    {chartData.length > 0 && (
      <div className="bg-white p-8 rounded-[3rem] shadow-sm border border-slate-100 no-print">
        <h4 className="text-lg font-black text-slate-800 mb-6 flex items-center gap-2">
          <div className="w-1.5 h-6 bg-blue-500 rounded-full" />
          売上トレンド可視化
        </h4>
        <div className="h-72 w-full">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={chartData}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
              <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{ fill: '#94a3b8', fontWeight: 700, fontSize: 12 }} />
              <YAxis axisLine={false} tickLine={false} tick={{ fill: '#94a3b8', fontWeight: 700, fontSize: 12 }} tickFormatter={v => fmtAmtShort(v)} />
              <Tooltip
                contentStyle={{ borderRadius: '1.5rem', border: 'none', boxShadow: '0 20px 25px -5px rgb(0 0 0 / 0.1)', padding: '1.5rem' }}
                cursor={{ fill: '#f8fafc' }}
                formatter={(value) => fmtAmt(value)}
              />
              <Legend iconType="circle" wrapperStyle={{ paddingTop: '2rem' }} />
              {years.length > 0 && (
                <>
                  <Bar dataKey={`${years[years.length - 1]}年 売上`} fill="#3b82f6" radius={[6, 6, 0, 0]} barSize={24} />
                  <Bar dataKey={`${years[years.length - 1]}年 粗利`} fill="#10b981" radius={[6, 6, 0, 0]} barSize={24} />
                </>
              )}
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    )}
  </div>
  );
};

// ===== ピボットレポートビュー =====
const PivotView = ({ data, years, branches, pivotBranch, onBranchChange, pivotSort, onSortChange, onSavePdf, onSaveCsv, fmtAmt, amountUnit, showProfit, selectedLeaseCo }) => {
  const titleParts = [
    selectedLeaseCo !== 'ALL' ? selectedLeaseCo : null,
    pivotBranch !== 'ALL' ? pivotBranch : null,
  ].filter(Boolean);
  const titleSuffix = titleParts.length > 0 ? titleParts.join(' — ') : null;

  return (
  <div className="space-y-8 animate-fade-in-up">
    {/* 支店フィルタ */}
    <div className="bg-white p-6 rounded-[3rem] shadow-sm border border-slate-100 no-print">
      <div className="space-y-3">
        <label className="text-[10px] font-black text-slate-400 uppercase tracking-[0.2em] ml-2">
          部店で絞り込み
        </label>
        <div className="flex gap-2 flex-wrap">
          <button
            onClick={() => onBranchChange('ALL')}
            className={`px-6 py-2.5 rounded-2xl text-xs font-black transition-all duration-300 ${
              pivotBranch === 'ALL'
                ? 'bg-blue-600 text-white shadow-xl shadow-blue-200'
                : 'bg-slate-100 text-slate-400 hover:bg-slate-200'
            }`}
          >
            すべて
          </button>
          {branches.map(b => (
            <button
              key={b}
              onClick={() => onBranchChange(b)}
              className={`px-6 py-2.5 rounded-2xl text-xs font-black transition-all duration-300 ${
                pivotBranch === b
                  ? 'bg-blue-600 text-white shadow-xl shadow-blue-200'
                  : 'bg-slate-100 text-slate-400 hover:bg-slate-200'
              }`}
            >
              {b}
            </button>
          ))}
        </div>
      </div>
    </div>

    <div className="bg-white rounded-[3rem] shadow-2xl shadow-slate-200/50 border border-slate-100 overflow-hidden">
      <div className="p-8 border-b border-slate-50 bg-slate-50/50 flex justify-between items-center">
        <div className="flex items-center gap-3">
          <div className="p-2 bg-blue-600 text-white rounded-xl shadow-lg shadow-blue-100">
            <ListFilter size={20} />
          </div>
          <div>
            <h3 className="font-black text-slate-800 text-xl tracking-tighter">
              一括網羅ピボットレポート
              {titleSuffix && <span className="text-blue-500 ml-2 text-base">— {titleSuffix}</span>}
            </h3>
            <p className="text-[10px] font-bold text-blue-500 mt-0.5 uppercase tracking-widest italic">
              Unified Comprehensive View for PDF &middot; {data.length} records
              {amountUnit === 'thousand' && <span className="ml-2 text-amber-500 normal-case">（単位：千円）</span>}
            </p>
          </div>
        </div>
        <div className="flex items-center gap-4">
          <div className="text-xs font-bold text-slate-400 flex items-center gap-1.5">
            <Calendar size={14} />
            出力日: {new Date().toLocaleDateString('ja-JP', { year: 'numeric', month: '2-digit', day: '2-digit' })}
          </div>
          <div className="flex gap-2 no-print">
            <button
              onClick={onSavePdf}
              className="flex items-center gap-2 px-4 py-2 bg-white rounded-xl border border-slate-200 text-slate-600 hover:text-blue-500 hover:border-blue-200 transition-all shadow-sm text-sm font-bold"
            >
              <FileText size={16} /> PDF
            </button>
            <button
              onClick={onSaveCsv}
              className="flex items-center gap-2 px-4 py-2 bg-white rounded-xl border border-slate-200 text-slate-600 hover:text-emerald-500 hover:border-emerald-200 transition-all shadow-sm text-sm font-bold"
            >
              <FileSpreadsheet size={16} /> CSV
            </button>
          </div>
        </div>
      </div>

      <div className="overflow-x-auto custom-scrollbar">
        <table className="w-full text-left border-collapse pivot-table">
          <thead>
            <tr className="bg-slate-900 text-xl font-black text-white tracking-wide text-center">
              <th className="px-4 py-4 text-center border-r border-slate-800 min-w-[120px] sticky left-0 bg-slate-900 z-10">
                <button onClick={() => onSortChange('orderer')} className={`w-full flex items-center justify-center gap-1 transition-colors ${pivotSort === 'orderer' ? 'text-blue-300' : 'text-white hover:text-blue-300'}`}>
                  注文者 {pivotSort === 'orderer' && <ArrowUpDown size={12} />}
                </button>
              </th>
              <th className="px-4 py-4 text-center border-r border-slate-800 min-w-[200px]">
                <button onClick={() => onSortChange('customer')} className={`w-full flex items-center justify-center gap-1 transition-colors ${pivotSort === 'customer' ? 'text-blue-300' : 'text-white hover:text-blue-300'}`}>
                  顧客名 {pivotSort === 'customer' && <ArrowUpDown size={12} />}
                </button>
              </th>
              {years.map(year => (
                <th key={year} className="px-3 py-4 border-l border-slate-800 min-w-[160px]">
                  {year}年度 実績
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-300">
            {data.length === 0 ? (
              <tr>
                <td colSpan={2 + years.length} className="px-8 py-16 text-center text-slate-300 italic">
                  該当するデータがありません
                </td>
              </tr>
            ) : (
              data.map((row, idx) => (
                <tr key={idx} className="hover:bg-blue-50/50 transition-colors">
                  <td className="px-4 py-3 border-r border-slate-300 sticky left-0 bg-white z-10 font-black text-slate-800 text-sm">
                    {row.orderer}
                  </td>
                  <td className="px-4 py-3 border-r border-slate-300">
                    <div className="flex items-center gap-1.5">
                      <div className="w-1.5 h-1.5 bg-blue-500 rounded-full flex-shrink-0" />
                      <span className="font-black text-slate-800 text-sm">{row.customer}</span>
                    </div>
                  </td>
                  {years.map((year, yIdx) => {
                    const sales = row.sales[year] || 0;
                    const profit = row.profit[year] || 0;
                    const margin = calcMargin(profit, sales);
                    const prevYear = years[yIdx - 1];
                    const yoy = prevYear ? calcYoY(sales, row.sales[prevYear]) : null;

                    return (
                      <td key={year} className="px-3 py-3 border-l border-slate-300 bg-white/40">
                        <div className="flex flex-col gap-0.5">
                          <div className="flex justify-between items-baseline">
                            <span className="text-[9px] font-black text-slate-400">売上</span>
                            <span className="font-mono font-black text-slate-700 text-[13px]">
                              {fmtAmt(sales)}
                            </span>
                          </div>
                          {showProfit && (
                            <>
                              <div className="flex justify-between items-baseline border-t border-slate-50 pt-0.5">
                                <span className="text-[9px] font-black text-slate-400">粗利</span>
                                <span className="font-mono font-black text-emerald-600 text-[13px]">
                                  {fmtAmt(profit)}
                                </span>
                              </div>
                              <div className="flex justify-between items-center">
                                <div className="px-1 py-0.5 bg-emerald-50 text-emerald-700 rounded text-[9px] font-black">
                                  {margin}%
                                </div>
                                {yoy !== null && (
                                  <div className={`text-[9px] font-black flex items-center gap-0.5 ${
                                    parseFloat(yoy) >= 0 ? 'text-emerald-500' : 'text-rose-500'
                                  }`}>
                                    {parseFloat(yoy) >= 0 ? <ArrowUpRight size={10} /> : <ArrowDownRight size={10} />}
                                    {yoy}%
                                  </div>
                                )}
                              </div>
                            </>
                          )}
                        </div>
                      </td>
                    );
                  })}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </div>
  </div>
  );
};

// ===== 設定モーダル =====
const SettingsModal = ({ networkPath, onPathChange, onSelectFolder, onClose, onSave }) => (
  <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-md z-50 flex items-center justify-center p-4 animate-fade-in">
    <div className="bg-white rounded-[3rem] w-full max-w-xl shadow-2xl animate-scale-in overflow-hidden">
      <div className="p-10 border-b border-slate-50 bg-slate-50/50 flex justify-between items-center">
        <div>
          <h3 className="text-2xl font-black text-slate-800">アプリ設定</h3>
          <p className="text-[10px] text-slate-400 font-black uppercase mt-1 tracking-widest">
            Environment Configuration
          </p>
        </div>
        <button
          onClick={onClose}
          className="w-10 h-10 rounded-full bg-white shadow-sm flex items-center justify-center text-slate-400 hover:text-slate-600 hover:shadow-md transition-all"
        >
          <XCircle size={20} />
        </button>
      </div>

      <div className="p-10 space-y-8">
        <div className="space-y-3">
          <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest flex items-center gap-2 ml-1">
            <FolderOpen size={14} className="text-blue-500" /> Excelデータ保存先パス (UNC)
          </label>
          <div className="flex gap-2">
            <input
              type="text"
              value={networkPath}
              onChange={e => onPathChange(e.target.value)}
              className="flex-1 bg-slate-100 border-none rounded-[1.5rem] p-5 text-sm font-bold text-slate-700 focus:ring-4 focus:ring-blue-500/10 transition-all"
              placeholder="\\192.1.1.103\share\..."
            />
            <button
              onClick={onSelectFolder}
              className="px-5 bg-slate-100 rounded-[1.5rem] text-slate-500 hover:bg-slate-200 hover:text-slate-700 transition-all font-bold text-sm flex items-center gap-2"
            >
              <FolderOpen size={16} /> 参照
            </button>
          </div>
          <p className="text-[10px] text-slate-400 px-4 italic leading-relaxed">
            ※このパスにあるExcelファイル（.xlsx/.xls）を自動的に統合・解析します。
          </p>
        </div>

        <div className="grid grid-cols-2 gap-4 pt-4">
          <div className="p-5 bg-blue-50 rounded-3xl border border-blue-100">
            <div className="w-8 h-8 bg-blue-600 text-white rounded-xl flex items-center justify-center mb-3 shadow-md shadow-blue-200">
              <CheckCircle2 size={16} />
            </div>
            <h4 className="text-xs font-black text-blue-900 mb-1 tracking-tight">自動データ同期</h4>
            <p className="text-[10px] text-blue-700 leading-normal font-bold">
              起動時に最新のExcelを自動的に読み込みます。
            </p>
          </div>
          <div className="p-5 bg-emerald-50 rounded-3xl border border-emerald-100">
            <div className="w-8 h-8 bg-emerald-600 text-white rounded-xl flex items-center justify-center mb-3 shadow-md shadow-emerald-200">
              <FileText size={16} />
            </div>
            <h4 className="text-xs font-black text-emerald-900 mb-1 tracking-tight">PDFレポート</h4>
            <p className="text-[10px] text-emerald-700 leading-normal font-bold">
              ボタン一つで現在の表示内容をPDF保存可能です。
            </p>
          </div>
        </div>
      </div>

      <div className="p-10 pt-0">
        <button
          onClick={onSave}
          className="w-full bg-slate-900 text-white py-5 rounded-[1.5rem] font-black shadow-2xl shadow-slate-900/20 active:scale-[0.98] transition-all flex items-center justify-center gap-2 hover:bg-blue-600"
        >
          <RefreshCcw size={18} />
          設定を保存して再読込
        </button>
      </div>
    </div>
  </div>
);

export default App;
