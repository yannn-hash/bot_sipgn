import { useState, useEffect, useRef } from 'react';
import axios from 'axios';
import { Play, PlayCircle, RefreshCw, Save, LogOut, CheckCircle, FileSpreadsheet, XCircle, ChevronDown, ChevronRight, X, FileText, Camera, Clock, Info, Send, User, Droplet, Sparkles, Trash2 } from 'lucide-react';

const API_BASE = 'http://localhost:8000';

function App() {
  const [data, setData] = useState([]);
  const [headers, setHeaders] = useState([]);
  const [selectedFile, setSelectedFile] = useState('kloter1.xlsx');
  const [logs, setLogs] = useState([]);
  const [isRunning, setIsRunning] = useState(false);
  const [isWaitingInput, setIsWaitingInput] = useState(false);
  const [expandedPorsi, setExpandedPorsi] = useState(null);
  const [expandedDriver, setExpandedDriver] = useState(null);
  const [selectedRowIndex, setSelectedRowIndex] = useState(null);
  const [completedActions, setCompletedActions] = useState({});
  const [showResetPopup, setShowResetPopup] = useState(false);
  const logEndRef = useRef(null);
  
  const REQUIRED_MODES = ['submit', 'mulaicuci', 'berangkat', 'terima', 'ambil', 'selesaicuci', 'kembali'];
  
  const PORSI_COLS = ['porsi_besar', 'porsi_kecil', 'balita', 'bumil', 'busui'];
  const DRIVER_COLS = ['nama_driver', 'plat_kendaraan'];
  
  const getTotalPorsi = (row) => {
    return PORSI_COLS.reduce((sum, col) => sum + (parseInt(row[col]) || 0), 0);
  };

  // Auto-scroll logs
  useEffect(() => {
    logEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [logs]);

  // Poll status saat berjalan
  useEffect(() => {
    let interval;
    if (isRunning) {
      interval = setInterval(async () => {
        try {
          const res = await axios.get(`${API_BASE}/api/status`);
          setIsWaitingInput(res.data.waiting);
        } catch (e) {}
      }, 1000);
    } else {
      setIsWaitingInput(false);
    }
    return () => clearInterval(interval);
  }, [isRunning]);

  const handleResume = async () => {
    try {
      await axios.post(`${API_BASE}/api/resume`);
    } catch (e) {
      console.error(e);
    }
  };

  const handleStop = async () => {
    try {
      await axios.post(`${API_BASE}/api/stop`);
    } catch (e) {
      console.error(e);
    }
  };

  const handleClearCache = async () => {
    try {
      setLogs(prev => [...prev, `[INFO] Sedang membersihkan cache browser...`]);
      const res = await axios.post(`${API_BASE}/api/clear-cache`);
      if (res.data.cleared && res.data.cleared.length > 0) {
        setLogs(prev => [...prev, `✓ Cache berhasil dibersihkan: ${res.data.cleared.join(', ')}`]);
      } else {
        setLogs(prev => [...prev, `✓ Tidak ada file cache yang perlu dibersihkan.`]);
      }
    } catch (e) {
      setLogs(prev => [...prev, `[ERROR] Gagal membersihkan cache.`]);
      console.error(e);
    }
  };

  // Cek apakah semua baris sudah selesai
  useEffect(() => {
    if (data.length > 0) {
      const isAllComplete = data.every(row => {
        const rowId = row._excel_row;
        const acts = completedActions[rowId] || [];
        return REQUIRED_MODES.every(m => acts.includes(m));
      });
      if (isAllComplete) {
        setShowResetPopup(true);
      }
    }
  }, [completedActions, data]);

  // Initial load
  useEffect(() => {
    fetchData();
    // Load state dari localStorage
    const saved = localStorage.getItem(`completedActions_${selectedFile}`);
    if (saved) {
      try {
        setCompletedActions(JSON.parse(saved));
      } catch (e) {
        setCompletedActions({});
      }
    } else {
      setCompletedActions({});
    }
    
    const cleanupWs = setupWebSocket();
    return () => {
      cleanupWs();
    };
  }, [selectedFile]);

  // Simpan progress ke localStorage tiap kali ada perubahan
  useEffect(() => {
    if (Object.keys(completedActions).length > 0 || localStorage.getItem(`completedActions_${selectedFile}`)) {
      localStorage.setItem(`completedActions_${selectedFile}`, JSON.stringify(completedActions));
    }
  }, [completedActions, selectedFile]);

  const setupWebSocket = () => {
    const ws = new WebSocket(`ws://localhost:8000/ws/logs`);
    ws.onmessage = (event) => {
      setLogs((prev) => [...prev, event.data]);
    };
    return () => ws.close();
  };

  const fetchData = async () => {
    try {
      const res = await axios.get(`${API_BASE}/api/data?file=${selectedFile}`);
      setHeaders(res.data.headers);
      setData(res.data.rows);
    } catch (err) {
      console.error(err);
      setLogs(prev => [...prev, `[ERROR] Gagal memuat data: ${err.message}`]);
    }
  };

  const handleCellChange = (rowIndex, column, value) => {
    const newData = [...data];
    newData[rowIndex][column] = value;
    setData(newData);
  };

  const saveData = async () => {
    try {
      await axios.post(`${API_BASE}/api/data?file=${selectedFile}`, { rows: data });
      setLogs(prev => [...prev, `✓ Data berhasil disimpan ke ${selectedFile}`]);
    } catch (err) {
      setLogs(prev => [...prev, `[ERROR] Gagal menyimpan data`]);
    }
  };

  const runBot = async (mode, rowIndex = null) => {
    setIsRunning(true);
    try {
      const payload = { mode, excel_file: selectedFile, row: rowIndex };
      await axios.post(`${API_BASE}/api/run`, payload);
      
      if (mode !== 'login') {
        setCompletedActions(prev => {
          const next = { ...prev };
          if (rowIndex !== null) {
             const acts = next[rowIndex] || [];
             if (!acts.includes(mode)) next[rowIndex] = [...acts, mode];
          } else {
             data.forEach(r => {
               const idx = r._excel_row;
               const acts = next[idx] || [];
               if (!acts.includes(mode)) next[idx] = [...acts, mode];
             });
          }
          return next;
        });
      }
    } catch (err) {
      setLogs(prev => [...prev, `[ERROR] Bot gagal berjalan: ${err.message}`]);
    } finally {
      setIsRunning(false);
    }
  };

  const handleReset = async () => {
    const newData = data.map(r => ({ ...r, uuid: '' }));
    setData(newData);
    setCompletedActions({});
    setShowResetPopup(false);
    
    try {
      await axios.post(`${API_BASE}/api/data?file=${selectedFile}`, { rows: newData });
      setLogs(prev => [...prev, `✓ Kloter selesai! UUID direset dan disimpan ke ${selectedFile}`]);
    } catch (err) {
      setLogs(prev => [...prev, `[ERROR] Gagal mereset UUID ke Excel`]);
    }
  };

  return (
    <div className="h-screen bg-slate-900 text-slate-100 font-sans selection:bg-blue-500 selection:text-white flex flex-col overflow-hidden">
      {/* HEADER */}
      <header className="bg-slate-800 border-b border-slate-700 px-6 py-4 flex items-center justify-between shadow-sm">
        <div className="flex items-center gap-3">
          <div className="bg-blue-600 p-2 rounded-lg shadow-lg shadow-blue-900/50">
            <PlayCircle className="text-white w-6 h-6" />
          </div>
          <div>
            <h1 className="text-xl font-bold bg-gradient-to-r from-blue-400 to-cyan-300 bg-clip-text text-transparent">SIPGN Bot UI</h1>
            <p className="text-xs text-slate-400 font-medium tracking-wide uppercase">Distribution Manager</p>
          </div>
        </div>
        
        <div className="flex gap-3">
          <select 
            className="bg-slate-900 border border-slate-700 rounded-md px-3 py-1.5 text-sm outline-none focus:border-blue-500 transition-colors"
            value={selectedFile}
            onChange={(e) => setSelectedFile(e.target.value)}
          >
            <option value="kloter1.xlsx">kloter1.xlsx</option>
            <option value="kloter2.xlsx">kloter2.xlsx</option>
            <option value="kloter3.xlsx">kloter3.xlsx</option>
            <option value="template_bot_distribusi_sipgn_v2.xlsx">template.xlsx</option>
          </select>
          
          <button 
            onClick={handleClearCache}
            className="flex items-center gap-2 bg-slate-800 hover:bg-slate-700 border border-slate-700 px-3 py-1.5 rounded-md text-sm font-medium transition-colors text-slate-300"
            title="Bersihkan folder cache browser"
          >
            <Trash2 className="w-4 h-4 text-slate-400" /> Bersihkan Cache
          </button>
          
          <button 
            onClick={saveData}
            className="flex items-center gap-2 bg-slate-700 hover:bg-slate-600 px-4 py-1.5 rounded-md text-sm font-medium transition-colors"
          >
            <Save className="w-4 h-4" /> Simpan Data
          </button>
          
          <button 
            onClick={() => runBot('login')}
            disabled={isRunning}
            className="flex items-center gap-2 bg-emerald-600 hover:bg-emerald-500 px-4 py-1.5 rounded-md text-sm font-medium transition-colors disabled:opacity-50"
          >
            <LogOut className="w-4 h-4" /> Setup Sesi Login
          </button>
        </div>
      </header>

      <main className="flex-1 flex overflow-hidden">
        {/* LEFT PANEL: DATA GRID */}
        <section className="flex-1 flex flex-col border-r border-slate-700 bg-slate-900/50">
          <div className="p-4 border-b border-slate-700 flex justify-between items-center bg-slate-800/50 backdrop-blur-sm">
            <h2 className="font-semibold text-slate-200 flex items-center gap-2">
              <FileSpreadsheet className="w-5 h-5 text-blue-400" /> 
              Tabel Data ({selectedFile})
            </h2>
            <div className="flex gap-2">
              <button onClick={() => runBot('submit')} className="text-xs bg-emerald-500/20 text-emerald-300 hover:bg-emerald-500/30 px-3 py-1 rounded border border-emerald-500/30 transition-colors">Submit Semua</button>
              <button onClick={() => runBot('berangkat')} className="text-xs bg-indigo-500/20 text-indigo-300 hover:bg-indigo-500/30 px-3 py-1 rounded border border-indigo-500/30 transition-colors">Berangkat Semua</button>
              <button onClick={() => runBot('terima')} className="text-xs bg-teal-500/20 text-teal-300 hover:bg-teal-500/30 px-3 py-1 rounded border border-teal-500/30 transition-colors">Terima Semua</button>
              <button onClick={() => runBot('ambil')} className="text-xs bg-cyan-500/20 text-cyan-300 hover:bg-cyan-500/30 px-3 py-1 rounded border border-cyan-500/30 transition-colors">Ambil Semua</button>
              <button onClick={() => runBot('kembali')} className="text-xs bg-fuchsia-500/20 text-fuchsia-300 hover:bg-fuchsia-500/30 px-3 py-1 rounded border border-fuchsia-500/30 transition-colors">Kembali Semua</button>
              <button onClick={() => runBot('mulaicuci')} className="text-xs bg-blue-500/20 text-blue-300 hover:bg-blue-500/30 px-3 py-1 rounded border border-blue-500/30 transition-colors">Mulai Cuci Semua</button>
              <button onClick={() => runBot('selesaicuci')} className="text-xs bg-amber-500/20 text-amber-300 hover:bg-amber-500/30 px-3 py-1 rounded border border-amber-500/30 transition-colors">Selesai Cuci Semua</button>
              
              <div className="w-px bg-slate-700/50 mx-1"></div>
              
              <button onClick={() => setShowResetPopup(true)} className="text-xs bg-red-500/20 text-red-400 hover:bg-red-500/30 px-3 py-1 rounded border border-red-500/30 transition-colors font-bold shadow-sm flex items-center gap-1">
                Reset Progress
              </button>
            </div>
          </div>
          
          {/* GLOBAL INPUTS CONTROL PANEL */}
          {data.length > 0 && (
            <div className="px-6 py-5 bg-slate-800/80 border-b border-slate-700/50 relative overflow-hidden">
              <div className="absolute top-0 right-0 w-64 h-64 bg-blue-500/5 rounded-full blur-3xl pointer-events-none transform translate-x-1/3 -translate-y-1/3"></div>
              
              <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-4 flex items-center gap-2">
                <FileText className="w-4 h-4" /> Pengaturan Universal Kloter
              </h3>
              
              <div className="grid grid-cols-2 md:grid-cols-4 xl:grid-cols-7 gap-4">
                <div className="flex flex-col gap-1.5">
                  <label className="text-[10px] font-semibold text-slate-400 uppercase">Tanggal WO</label>
                  <input 
                    type="text" 
                    className="w-full bg-slate-900 border border-slate-700 focus:border-blue-500 text-slate-200 px-3 py-1.5 rounded outline-none transition-colors text-xs shadow-inner"
                    value={data[0]?.tanggal_wo || ''}
                    onChange={(e) => setData(data.map(r => ({ ...r, tanggal_wo: e.target.value })))}
                  />
                </div>
                <div className="flex flex-col gap-1.5">
                  <label className="text-[10px] font-semibold text-slate-400 uppercase">WO ID</label>
                  <input 
                    type="text" 
                    className="w-full bg-slate-900 border border-slate-700 focus:border-blue-500 text-slate-200 px-3 py-1.5 rounded outline-none transition-colors text-xs shadow-inner"
                    value={data[0]?.wo_id || ''}
                    onChange={(e) => setData(data.map(r => ({ ...r, wo_id: e.target.value })))}
                  />
                </div>
                <div className="flex flex-col gap-1.5">
                  <label className="text-[10px] font-semibold text-slate-400 uppercase">Kode WO</label>
                  <input 
                    type="text" 
                    className="w-full bg-slate-900 border border-slate-700 focus:border-blue-500 text-slate-200 px-3 py-1.5 rounded outline-none transition-colors text-xs shadow-inner"
                    value={data[0]?.kode_wo || ''}
                    onChange={(e) => setData(data.map(r => ({ ...r, kode_wo: e.target.value })))}
                  />
                </div>
                <div className="flex flex-col gap-1.5">
                  <label className="text-[10px] font-semibold text-slate-400 uppercase">Urutan Rute</label>
                  <input 
                    type="text" 
                    className="w-full bg-slate-900 border border-slate-700 focus:border-blue-500 text-slate-200 px-3 py-1.5 rounded outline-none transition-colors text-xs shadow-inner"
                    value={data[0]?.urutan_rute || ''}
                    onChange={(e) => setData(data.map(r => ({ ...r, urutan_rute: e.target.value })))}
                  />
                </div>
                <div className="flex flex-col gap-1.5">
                  <label className="text-[10px] font-semibold text-slate-400 uppercase">Batas Waktu</label>
                  <input 
                    type="text" 
                    className="w-full bg-slate-900 border border-slate-700 focus:border-blue-500 text-slate-200 px-3 py-1.5 rounded outline-none transition-colors text-xs shadow-inner"
                    value={data[0]?.batas_waktu || ''}
                    onChange={(e) => setData(data.map(r => ({ ...r, batas_waktu: e.target.value })))}
                  />
                </div>
                <div className="flex flex-col gap-1.5">
                  <label className="text-[10px] font-semibold text-blue-400 uppercase">Mulai Cuci</label>
                  <input 
                    type="text" 
                    className="w-full bg-slate-900 border border-blue-500/30 focus:border-blue-500 text-slate-200 px-3 py-1.5 rounded outline-none transition-colors text-xs shadow-inner"
                    value={data[0]?.mulai_cuci || ''}
                    onChange={(e) => setData(data.map(r => ({ ...r, mulai_cuci: e.target.value })))}
                  />
                </div>
                <div className="flex flex-col gap-1.5">
                  <label className="text-[10px] font-semibold text-amber-400 uppercase">Selesai Cuci</label>
                  <input 
                    type="text" 
                    className="w-full bg-slate-900 border border-amber-500/30 focus:border-amber-500 text-slate-200 px-3 py-1.5 rounded outline-none transition-colors text-xs shadow-inner"
                    value={data[0]?.selesai_cuci || ''}
                    onChange={(e) => setData(data.map(r => ({ ...r, selesai_cuci: e.target.value })))}
                  />
                </div>
              </div>
            </div>
          )}
          
          
          <div className="flex-1 overflow-auto p-4">
            <div className="inline-block min-w-full align-middle border border-slate-700 rounded-lg overflow-hidden shadow-xl bg-slate-800">
              <table className="min-w-full divide-y divide-slate-700 text-sm">
                <thead className="bg-slate-900/80">
                  <tr>
                    <th className="sticky top-0 px-3 py-3 text-left text-xs font-medium text-slate-400 uppercase tracking-wider bg-slate-900 shadow-sm z-10 w-24 border-r border-slate-700/50">Aksi</th>
                    <th className="sticky top-0 px-3 py-3 text-left text-xs font-medium text-slate-400 uppercase tracking-wider bg-slate-900 shadow-sm z-10 whitespace-nowrap align-top border-r border-slate-700/50">
                      <div className="resize-x overflow-hidden min-w-[120px] w-[150px] max-w-[500px] pb-2">
                        Tujuan Nama
                      </div>
                    </th>
                    <th className="sticky top-0 px-3 py-3 text-left text-xs font-medium text-slate-400 uppercase tracking-wider bg-slate-900 shadow-sm z-10 align-top border-r border-slate-700/50">
                      <div className="resize-x overflow-hidden min-w-[150px] w-[200px] max-w-[400px] pb-2">
                        Porsi
                      </div>
                    </th>
                    <th className="sticky top-0 px-3 py-3 text-center text-xs font-medium text-slate-400 uppercase tracking-wider bg-slate-900 shadow-sm z-10 w-24">Detail</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-700/50">
                  {data.map((row, idx) => (
                    <tr key={idx} className="hover:bg-slate-700/30 transition-colors group">
                      <td className="px-2 py-2 whitespace-nowrap border-r border-slate-700/30">
                        <div className="flex gap-1 opacity-50 group-hover:opacity-100 transition-opacity">
                           {[
                             { id: 'submit', icon: Send, title: 'Submit Awal' },
                             { id: 'berangkat', icon: Play, title: 'Berangkat' },
                             { id: 'terima', icon: CheckCircle, title: 'Terima' },
                             { id: 'ambil', icon: RefreshCw, title: 'Ambil' },
                             { id: 'kembali', icon: XCircle, title: 'Kembali' },
                             { id: 'mulaicuci', icon: Droplet, title: 'Mulai Cuci' },
                             { id: 'selesaicuci', icon: Sparkles, title: 'Selesai Cuci' }
                           ].map(action => {
                             const isComplete = completedActions[row._excel_row]?.includes(action.id);
                             const btnClass = isComplete 
                               ? 'p-1 bg-emerald-500/20 text-emerald-400 rounded hover:bg-emerald-500/30 shadow-sm border border-emerald-500/30'
                               : 'p-1 bg-red-500/10 text-red-400 rounded hover:bg-red-500/20 hover:text-red-300 border border-transparent';
                             return (
                               <button 
                                 key={action.id}
                                 onClick={() => runBot(action.id, row._excel_row)} 
                                 title={action.title} 
                                 className={`${btnClass} transition-all`}
                               >
                                 <action.icon className="w-4 h-4" />
                               </button>
                             );
                           })}
                        </div>
                      </td>
                      <td className="px-3 py-2 whitespace-nowrap p-0 border-r border-slate-700/30">
                        <input 
                          type="text" 
                          className="w-full bg-transparent border-transparent focus:border-blue-500 focus:bg-slate-800 text-slate-300 px-2 py-1 rounded outline-none transition-colors"
                          value={row.tujuan_nama || ''} 
                          onChange={(e) => handleCellChange(idx, 'tujuan_nama', e.target.value)}
                        />
                      </td>
                      <td className="px-3 py-2 border-r border-slate-700/30">
                        <div className="flex flex-col gap-1 w-full min-w-full">
                          <div 
                            className="flex items-center justify-between text-xs cursor-pointer text-slate-300 hover:text-blue-400 bg-slate-800 px-2 py-1.5 rounded transition-colors" 
                            onClick={() => setExpandedPorsi(expandedPorsi === idx ? null : idx)}
                          >
                            <span className="font-semibold">{getTotalPorsi(row)} Total Porsi</span>
                            {expandedPorsi === idx ? <ChevronDown className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />}
                          </div>
                          {expandedPorsi === idx && (
                            <div className="mt-1 grid grid-cols-2 gap-2 bg-slate-900/50 p-2 rounded border border-slate-700/50">
                              {PORSI_COLS.map(pc => (
                                <div key={pc} className="flex flex-col">
                                  <label className="text-[9px] font-semibold text-slate-500 uppercase tracking-wider mb-0.5">{pc.replace('porsi_', '')}</label>
                                  <input 
                                    type="number" 
                                    className="w-full bg-slate-800 border border-slate-700 focus:border-blue-500 text-slate-300 px-2 py-1 rounded outline-none transition-colors text-xs"
                                    value={row[pc] || ''} 
                                    onChange={(e) => handleCellChange(idx, pc, e.target.value)}
                                  />
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      </td>
                      <td className="px-3 py-2 text-center whitespace-nowrap">
                        <button 
                          onClick={() => setSelectedRowIndex(idx)}
                          className="text-xs bg-slate-700 hover:bg-blue-600 text-slate-200 px-3 py-1.5 rounded flex items-center justify-center gap-1 transition-colors w-full"
                        >
                          <Info className="w-3.5 h-3.5" /> Detail
                        </button>
                      </td>
                    </tr>
                  ))}
                  {data.length === 0 && (
                    <tr>
                      <td colSpan={headers.length + 2} className="px-6 py-12 text-center text-slate-500">
                        Tidak ada data yang ditemukan.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </section>

        {/* RIGHT PANEL: TERMINAL */}
        <section className="w-96 flex flex-col bg-slate-950 border-l border-slate-800 shadow-2xl">
          <div className="p-4 border-b border-slate-800 bg-slate-900 flex justify-between items-center">
            <h2 className="font-semibold text-slate-300 flex items-center gap-2 text-sm uppercase tracking-wider">
              Terminal Log
              {isRunning && !isWaitingInput && <span className="flex h-2 w-2 relative ml-2"><span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span><span className="relative inline-flex rounded-full h-2 w-2 bg-emerald-500"></span></span>}
              {isWaitingInput && <span className="flex h-2 w-2 relative ml-2"><span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-amber-400 opacity-75"></span><span className="relative inline-flex rounded-full h-2 w-2 bg-amber-500"></span></span>}
            </h2>
            {isRunning && (
              <button 
                onClick={handleStop}
                className="text-xs bg-red-500/20 text-red-300 hover:bg-red-500/30 hover:text-red-200 px-2 py-1 rounded border border-red-500/30 transition-colors flex items-center gap-1 shadow-sm"
              >
                <XCircle className="w-3 h-3" /> Stop
              </button>
            )}
          </div>
          {isWaitingInput && (
            <div className="p-3 bg-amber-900/50 border-b border-amber-500/30">
              <button 
                onClick={handleResume}
                className="w-full bg-amber-500 hover:bg-amber-400 text-amber-950 font-bold py-2 px-4 rounded shadow shadow-amber-500/20 transition-all flex items-center justify-center gap-2 animate-pulse hover:animate-none"
              >
                <CheckCircle className="w-5 h-5" />
                LANJUTKAN BOT
              </button>
            </div>
          )}
          <div className="flex-1 overflow-auto p-4 font-mono text-xs text-slate-400 space-y-1.5 selection:bg-slate-700">
            {logs.length === 0 ? (
              <div className="text-slate-600 italic">Menunggu eksekusi...</div>
            ) : (
              logs.map((log, i) => {
                let color = 'text-slate-400';
                if (log.includes('✓')) color = 'text-emerald-400';
                if (log.includes('⚠') || log.includes('🟡')) color = 'text-amber-400';
                if (log.includes('❌') || log.includes('[ERROR]')) color = 'text-rose-400';
                if (log.includes('===')) color = 'text-blue-400 font-bold';
                return <div key={i} className={color}>{log}</div>;
              })
            )}
            <div ref={logEndRef} />
          </div>
        </section>
      </main>

      {/* SIDE DRAWER */}
      <div 
        className={`fixed inset-y-0 right-0 w-96 bg-slate-900 border-l border-slate-700 shadow-2xl flex flex-col z-50 transform transition-transform duration-300 ease-in-out ${selectedRowIndex !== null ? 'translate-x-0' : 'translate-x-full'}`}
      >
        {selectedRowIndex !== null && data[selectedRowIndex] && (
          <>
            <div className="p-4 border-b border-slate-800 flex justify-between items-center bg-slate-800/80 backdrop-blur">
              <div>
                <h2 className="font-bold text-slate-200">Detail Lengkap</h2>
                <p className="text-xs text-blue-400 truncate w-64">{data[selectedRowIndex].tujuan_nama || `Baris ${selectedRowIndex + 1}`}</p>
              </div>
              <button onClick={() => setSelectedRowIndex(null)} className="p-1.5 hover:bg-slate-700 rounded-md text-slate-400 hover:text-rose-400 transition-colors">
                <X className="w-5 h-5"/>
              </button>
            </div>
            
            <div className="flex-1 overflow-auto p-5 space-y-6 bg-slate-900/50">
              
              {/* DRIVER */}
              <div className="space-y-3">
                <h3 className="text-xs font-bold text-amber-400 uppercase tracking-wider flex items-center gap-2 border-b border-slate-700/50 pb-2"><User className="w-4 h-4" /> Data Driver</h3>
                <div className="grid grid-cols-1 gap-3 bg-slate-800/50 p-3 rounded-lg border border-slate-700/50">
                  {DRIVER_COLS.map(dc => (
                    <div key={dc} className="flex flex-col">
                      <label className="text-[10px] font-semibold text-slate-400 uppercase mb-1">{dc.replace('_', ' ')}</label>
                      <input 
                        type="text" 
                        className="w-full bg-slate-800 border border-slate-700 focus:border-amber-500 text-slate-200 px-3 py-1.5 rounded outline-none transition-colors text-xs shadow-inner" 
                        value={data[selectedRowIndex][dc] || ''} 
                        onChange={(e) => handleCellChange(selectedRowIndex, dc, e.target.value)} 
                      />
                    </div>
                  ))}
                </div>
              </div>

              {/* KEBERANGKATAN */}
              <div className="space-y-3">
                <h3 className="text-xs font-bold text-indigo-400 uppercase tracking-wider flex items-center gap-2 border-b border-slate-700/50 pb-2"><Clock className="w-4 h-4" /> Keberangkatan</h3>
                <div className="grid grid-cols-1 gap-3 bg-slate-800/50 p-3 rounded-lg border border-slate-700/50">
                  <div className="flex flex-col">
                    <label className="text-[10px] font-semibold text-slate-400 uppercase mb-1">Waktu Keberangkatan</label>
                    <input type="text" className="w-full bg-slate-800 border border-slate-700 focus:border-indigo-500 text-slate-200 px-3 py-1.5 rounded outline-none transition-colors text-xs shadow-inner" value={data[selectedRowIndex].waktu_keberangkatan || ''} onChange={(e) => handleCellChange(selectedRowIndex, 'waktu_keberangkatan', e.target.value)} />
                  </div>
                  {['foto_1', 'foto_2', 'foto_3', 'foto_4'].map(f => (
                     <div key={f} className="flex flex-col">
                        <label className="text-[10px] font-semibold text-slate-400 uppercase mb-1">{f.replace('_', ' ')}</label>
                        <input type="text" className="w-full bg-slate-800 border border-slate-700 focus:border-indigo-500 text-slate-200 px-3 py-1.5 rounded outline-none transition-colors text-xs shadow-inner" value={data[selectedRowIndex][f] || ''} onChange={(e) => handleCellChange(selectedRowIndex, f, e.target.value)} />
                     </div>
                  ))}
                </div>
              </div>

              {/* PENGAMBILAN */}
              <div className="space-y-3">
                <h3 className="text-xs font-bold text-cyan-400 uppercase tracking-wider flex items-center gap-2 border-b border-slate-700/50 pb-2"><Camera className="w-4 h-4" /> Pengambilan</h3>
                <div className="grid grid-cols-1 gap-3 bg-slate-800/50 p-3 rounded-lg border border-slate-700/50">
                  <div className="flex flex-col">
                    <label className="text-[10px] font-semibold text-slate-400 uppercase mb-1">Waktu Ambil</label>
                    <input type="text" className="w-full bg-slate-800 border border-slate-700 focus:border-cyan-500 text-slate-200 px-3 py-1.5 rounded outline-none transition-colors text-xs shadow-inner" value={data[selectedRowIndex].waktu_ambil || ''} onChange={(e) => handleCellChange(selectedRowIndex, 'waktu_ambil', e.target.value)} />
                  </div>
                  <div className="flex flex-col">
                    <label className="text-[10px] font-semibold text-slate-400 uppercase mb-1">Bukti Foto Ambil</label>
                    <input type="text" className="w-full bg-slate-800 border border-slate-700 focus:border-cyan-500 text-slate-200 px-3 py-1.5 rounded outline-none transition-colors text-xs shadow-inner" value={data[selectedRowIndex].bukti_foto_ambil || ''} onChange={(e) => handleCellChange(selectedRowIndex, 'bukti_foto_ambil', e.target.value)} />
                  </div>
                </div>
              </div>

              {/* PENERIMAAN */}
              <div className="space-y-3">
                <h3 className="text-xs font-bold text-teal-400 uppercase tracking-wider flex items-center gap-2 border-b border-slate-700/50 pb-2"><Camera className="w-4 h-4" /> Penerimaan</h3>
                <div className="grid grid-cols-1 gap-3 bg-slate-800/50 p-3 rounded-lg border border-slate-700/50">
                  <div className="flex flex-col">
                    <label className="text-[10px] font-semibold text-slate-400 uppercase mb-1">Waktu Diterima</label>
                    <input type="text" className="w-full bg-slate-800 border border-slate-700 focus:border-teal-500 text-slate-200 px-3 py-1.5 rounded outline-none transition-colors text-xs shadow-inner" value={data[selectedRowIndex].waktu_diterima || ''} onChange={(e) => handleCellChange(selectedRowIndex, 'waktu_diterima', e.target.value)} />
                  </div>
                  <div className="flex flex-col">
                    <label className="text-[10px] font-semibold text-slate-400 uppercase mb-1">Foto Bukti Terima</label>
                    <input type="text" className="w-full bg-slate-800 border border-slate-700 focus:border-teal-500 text-slate-200 px-3 py-1.5 rounded outline-none transition-colors text-xs shadow-inner" value={data[selectedRowIndex].foto_bukti_terima || ''} onChange={(e) => handleCellChange(selectedRowIndex, 'foto_bukti_terima', e.target.value)} />
                  </div>
                </div>
              </div>

              {/* KEMBALI */}
              <div className="space-y-3">
                <h3 className="text-xs font-bold text-fuchsia-400 uppercase tracking-wider flex items-center gap-2 border-b border-slate-700/50 pb-2"><Camera className="w-4 h-4" /> Kembali</h3>
                <div className="grid grid-cols-1 gap-3 bg-slate-800/50 p-3 rounded-lg border border-slate-700/50">
                  <div className="flex flex-col">
                    <label className="text-[10px] font-semibold text-slate-400 uppercase mb-1">Waktu Kembali</label>
                    <input type="text" className="w-full bg-slate-800 border border-slate-700 focus:border-fuchsia-500 text-slate-200 px-3 py-1.5 rounded outline-none transition-colors text-xs shadow-inner" value={data[selectedRowIndex].waktu_kembali || ''} onChange={(e) => handleCellChange(selectedRowIndex, 'waktu_kembali', e.target.value)} />
                  </div>
                  <div className="flex flex-col">
                    <label className="text-[10px] font-semibold text-slate-400 uppercase mb-1">Bukti Foto Kembali</label>
                    <input type="text" className="w-full bg-slate-800 border border-slate-700 focus:border-fuchsia-500 text-slate-200 px-3 py-1.5 rounded outline-none transition-colors text-xs shadow-inner" value={data[selectedRowIndex].bukti_foto_kembali || ''} onChange={(e) => handleCellChange(selectedRowIndex, 'bukti_foto_kembali', e.target.value)} />
                  </div>
                </div>
              </div>

              {/* LAINNYA */}
              <div className="space-y-3 pb-8">
                <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider flex items-center gap-2 border-b border-slate-700/50 pb-2"><FileText className="w-4 h-4" /> Lainnya</h3>
                <div className="grid grid-cols-1 gap-3 bg-slate-800/50 p-3 rounded-lg border border-slate-700/50">
                  {['catatan', 'UUID', 'status_bot', 'pesan_bot', 'waktu_submit'].map(f => (
                     <div key={f} className="flex flex-col">
                        <label className="text-[10px] font-semibold text-slate-400 uppercase mb-1">{f.replace('_', ' ')}</label>
                        {f === 'catatan' || f === 'pesan_bot' ? (
                          <textarea className="w-full bg-slate-800 border border-slate-700 focus:border-slate-500 text-slate-200 px-3 py-2 rounded outline-none transition-colors text-xs resize-y min-h-[60px] shadow-inner" value={data[selectedRowIndex]?.[f] || ''} onChange={(e) => handleCellChange(selectedRowIndex, f, e.target.value)} />
                        ) : (
                          <input type="text" className="w-full bg-slate-800 border border-slate-700 focus:border-slate-500 text-slate-200 px-3 py-1.5 rounded outline-none transition-colors text-xs shadow-inner" value={data[selectedRowIndex]?.[f] || ''} onChange={(e) => handleCellChange(selectedRowIndex, f, e.target.value)} disabled={f === 'UUID'} />
                        )}
                     </div>
                  ))}
                </div>
              </div>

            </div>
          </>
        )}
      </div>
      
      {/* RESET POPUP MODAL */}
      {showResetPopup && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-slate-950/80 backdrop-blur-sm">
          <div className="bg-slate-900 border border-slate-700 shadow-2xl rounded-xl w-full max-w-md overflow-hidden transform animate-in fade-in zoom-in-95 duration-200">
            <div className="bg-gradient-to-r from-emerald-600 to-teal-500 p-6 flex flex-col items-center justify-center text-center">
              <div className="bg-white/20 p-3 rounded-full mb-3">
                <Sparkles className="w-8 h-8 text-white" />
              </div>
              <h2 className="text-xl font-bold text-white tracking-wide">Kloter Selesai!</h2>
            </div>
            <div className="p-6 text-center space-y-4">
              <p className="text-slate-300 text-sm">
                Seluruh aksi pada <span className="font-semibold text-white">{selectedFile}</span> telah berhasil dihijaukan. 
              </p>
              <p className="text-slate-400 text-xs bg-slate-800/50 p-3 rounded border border-slate-700/50">
                Klik <strong className="text-emerald-400">Ya, Reset</strong> untuk menghapus semua status hijau dan mengosongkan kolom UUID sebagai persiapan untuk siklus tugas di hari berikutnya.
              </p>
              
              <div className="flex gap-3 pt-4">
                <button 
                  onClick={() => setShowResetPopup(false)}
                  className="flex-1 px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg text-sm font-semibold border border-slate-700 transition-colors"
                >
                  Nanti Saja
                </button>
                <button 
                  onClick={handleReset}
                  className="flex-1 px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-sm font-semibold shadow-lg shadow-emerald-900/50 transition-colors"
                >
                  Ya, Reset
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
