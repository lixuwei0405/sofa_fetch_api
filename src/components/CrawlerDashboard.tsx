import React, { useState, useRef, useEffect, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { Play, RotateCcw, Upload, CheckCircle2, XCircle, Clock, Loader2, Pause, Trash2 } from 'lucide-react';
import { cn } from '../lib/utils';
import { TaskRow } from '../types';

export default function CrawlerDashboard() {
  const [urlTemplate, setUrlTemplate] = useState('http://xxxx:xxxx/api-crawler/sofa/crawl/league/{mpLeagueId}');
  const [httpMethod, setHttpMethod] = useState<'GET' | 'POST'>('GET');
  const [concurrency, setConcurrency] = useState<number>(1);
  const [skipHeader, setSkipHeader] = useState<boolean>(true);
  const [selectedColumnIndex, setSelectedColumnIndex] = useState<number>(0);
  
  const [excelData, setExcelData] = useState<any[][]>([]);
  const [tasks, setTasks] = useState<TaskRow[]>([]);
  const [isRunning, setIsRunning] = useState(false);
  
  const isRunningRef = useRef(isRunning);
  const tasksRef = useRef(tasks);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Sync refs for async access
  useEffect(() => {
    isRunningRef.current = isRunning;
  }, [isRunning]);
  useEffect(() => {
    tasksRef.current = tasks;
  }, [tasks]);

  // Process Excel data into tasks
  useEffect(() => {
    if (excelData.length === 0) {
      setTasks([]);
      return;
    }
    
    const newTasks: TaskRow[] = [];
    const startIndex = skipHeader ? 1 : 0;
    
    for (let i = startIndex; i < excelData.length; i++) {
      const row = excelData[i];
      if (row && row.length > selectedColumnIndex) {
        const idVal = row[selectedColumnIndex];
        if (idVal !== undefined && idVal !== null && idVal !== '') {
          newTasks.push({
            id: String(idVal).trim(),
            originalIndex: i,
            status: 'pending'
          });
        }
      }
    }
    setTasks(newTasks);
  }, [excelData, skipHeader, selectedColumnIndex]);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      if (!bstr) return;
      
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json<any[]>(ws, { header: 1 });
      
      setExcelData(data);
      // Reset input so the same file can be uploaded again if needed
      if (fileInputRef.current) fileInputRef.current.value = '';
    };
    reader.readAsBinaryString(file);
  };

  const clearData = () => {
    setExcelData([]);
    setTasks([]);
  };

  const getConstructedUrl = (id: string) => {
    // Robust, case-insensitive replacement for various placeholder patterns
    return urlTemplate.replace(/\{mpleagueid\}|\{id\}|\{leagueid\}/gi, id);
  };

  const processSingleTask = async (taskIndex: number, taskId: string) => {
    setTasks(prev => {
      const next = [...prev];
      if (next[taskIndex]) {
        next[taskIndex] = { ...next[taskIndex], status: 'running', report: '' };
      }
      return next;
    });

    const targetUrl = getConstructedUrl(taskId);

    try {
      const res = await fetch('/api/proxy', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ targetUrl, method: httpMethod })
      });
      
      const data = await res.json();
      let isSuccess = false;
      let reportText = '';

      if (res.ok && data.message === 'success') {
        const hasErrors = Array.isArray(data.data?.errors) && data.data.errors.length > 0;
        const teamFailed = data.data?.teamStats?.failed || 0;
        const playerFailed = data.data?.playerStats?.failed || 0;

        if (!hasErrors && teamFailed === 0 && playerFailed === 0) {
          isSuccess = true;
        }
        reportText = data.data?.report || JSON.stringify(data.data, null, 2);
      } else {
        reportText = data.message || data.error || JSON.stringify(data, null, 2);
      }

      setTasks(prev => {
        const next = [...prev];
        if (next[taskIndex]) {
          next[taskIndex] = { ...next[taskIndex], status: isSuccess ? 'success' : 'failed', report: reportText };
        }
        return next;
      });
    } catch (err: any) {
      setTasks(prev => {
        const next = [...prev];
        if (next[taskIndex]) {
          next[taskIndex] = { ...next[taskIndex], status: 'failed', report: err.message };
        }
        return next;
      });
    }
  };

  const toggleRunning = async () => {
    if (isRunning) {
      setIsRunning(false);
      isRunningRef.current = false;
      return;
    }

    setIsRunning(true);
    isRunningRef.current = true;
    
    const allPendingOrFailed = tasksRef.current
      .map((t, i) => ({ ...t, idx: i }))
      .filter(t => t.status === 'pending' || t.status === 'failed');

    let currentIndex = 0;

    const workers = Array.from({ length: concurrency }, async () => {
      while (currentIndex < allPendingOrFailed.length) {
        if (!isRunningRef.current) break;
        
        const myIndex = currentIndex++;
        if (myIndex >= allPendingOrFailed.length) break;
        const taskObj = allPendingOrFailed[myIndex];
        
        await processSingleTask(taskObj.idx, taskObj.id);
      }
    });

    await Promise.all(workers);
    setIsRunning(false);
    isRunningRef.current = false;
  };

  const handleManualRun = (idx: number) => {
    if (isRunning) return;
    const task = tasks[idx];
    if (!task || task.status === 'running') return;
    processSingleTask(idx, task.id);
  };

  const stats = {
    total: tasks.length,
    success: tasks.filter(t => t.status === 'success').length,
    failed: tasks.filter(t => t.status === 'failed').length,
    running: tasks.filter(t => t.status === 'running').length,
    pending: tasks.filter(t => t.status === 'pending').length,
  };
  
  const progressPercent = stats.total > 0 ? ((stats.success + stats.failed) / stats.total) * 100 : 0;

  return (
    <div className="flex flex-col h-screen w-screen bg-slate-950 text-slate-200 overflow-hidden p-6 gap-6 font-sans">
      {/* Header Section */}
      <header className="flex items-center justify-between shrink-0">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-indigo-600 rounded-lg flex items-center justify-center shadow-lg shadow-indigo-500/20">
            <svg className="w-6 h-6 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 7v10c0 2.21 3.582 4 8 4s8-1.79 8-4V7M4 7c0 2.21 3.582 4 8 4s8-1.79 8-4M4 7c0-2.21 3.582-4 8-4s8 1.79 8 4m0 5c0 2.21-3.582 4-8 4s-8-1.79-8-4"></path></svg>
          </div>
          <div>
            <h1 className="text-xl font-bold tracking-tight text-white">SOFA Crawler <span className="text-indigo-400">Pro</span></h1>
            <p className="text-xs text-slate-400">v2.4.0 • Enterprise Data Compensation Engine</p>
          </div>
        </div>
        <div className="flex items-center gap-4">
          <div className="flex flex-col items-end">
            <span className="text-[10px] uppercase tracking-widest text-slate-500 font-bold">API Status</span>
            <span className="flex items-center gap-2 text-emerald-400 text-sm font-medium">● Operational</span>
          </div>
        </div>
      </header>

      {/* Main Bento Grid */}
      <div className="grid grid-cols-12 grid-rows-12 flex-1 gap-6 min-h-0">
        
        {/* Configuration Card (Left Column) */}
        <div className="col-span-4 row-span-12 bg-slate-900/50 border border-slate-800 rounded-2xl p-5 flex flex-col gap-6 overflow-y-auto custom-scrollbar">
          <div>
            <h2 className="text-sm font-bold text-slate-300 uppercase tracking-wider mb-4">1. Endpoint Configuration</h2>
            <div className="space-y-4">
              <div>
                <label className="text-[10px] text-slate-500 uppercase font-bold block mb-1.5">API Endpoint Template</label>
                <textarea 
                  className="w-full bg-slate-800 border border-slate-700 rounded px-3 py-2 text-xs font-mono text-indigo-300 focus:ring-1 ring-indigo-500 outline-none"
                  rows={3}
                  value={urlTemplate}
                  onChange={e => setUrlTemplate(e.target.value)}
                  disabled={isRunning}
                  placeholder="http://.../league/{mpLeagueId}"
                />
                <p className="text-[10px] text-slate-500 mt-1">Use <code className="bg-slate-800 px-1 py-0.5 rounded text-indigo-400">{"{mpLeagueId}"}</code> as placeholder.</p>
              </div>
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-[10px] text-slate-500 uppercase font-bold block mb-1.5">Method</label>
                  <select 
                    className="w-full bg-slate-800 border border-slate-700 text-xs px-2 py-2 rounded focus:ring-1 ring-indigo-500 outline-none text-slate-200"
                    value={httpMethod}
                    onChange={e => setHttpMethod(e.target.value as any)}
                    disabled={isRunning}
                  >
                    <option value="GET">GET</option>
                    <option value="POST">POST</option>
                  </select>
                </div>
                <div>
                  <label className="text-[10px] text-slate-500 uppercase font-bold block mb-1.5">Threads</label>
                  <input 
                    type="number"
                    min={1}
                    max={20}
                    className="w-full bg-slate-800 border border-slate-700 rounded px-3 py-2 text-xs text-slate-200 focus:ring-1 ring-indigo-500 outline-none"
                    value={concurrency}
                    onChange={e => setConcurrency(parseInt(e.target.value) || 1)}
                    disabled={isRunning}
                  />
                </div>
              </div>
            </div>
          </div>

          <div>
            <h2 className="text-sm font-bold text-slate-300 uppercase tracking-wider mb-4">2. Data Source</h2>
            
            <input 
              type="file" 
              accept=".xlsx, .xls, .csv" 
              className="hidden" 
              ref={fileInputRef}
              onChange={handleFileUpload}
              disabled={isRunning}
            />
            
            <div 
              onClick={() => !isRunning && fileInputRef.current?.click()}
              className={cn(
                "p-6 border-2 border-dashed border-slate-700 rounded-xl flex flex-col items-center justify-center text-center cursor-pointer transition-colors",
                isRunning ? "bg-slate-900/30 opacity-50 cursor-not-allowed" : "bg-slate-800/30 hover:bg-slate-800/50 hover:border-slate-600"
              )}
            >
              <Upload className="w-8 h-8 text-slate-500 mb-2" />
              <p className="text-xs font-medium text-slate-400">Select Excel File</p>
              {excelData.length > 0 && <p className="text-[10px] text-emerald-500 mt-1">{excelData.length} Rows Loaded</p>}
            </div>

            {excelData.length > 0 && (
              <div className="mt-4 space-y-3 p-4 border border-slate-800 rounded-xl bg-slate-900/30">
                <label className="flex items-center cursor-pointer">
                  <input 
                    type="checkbox" 
                    checked={skipHeader} 
                    onChange={e => setSkipHeader(e.target.checked)} 
                    disabled={isRunning}
                    className="accent-indigo-500 mr-2 w-4 h-4 rounded border-slate-700 bg-slate-800"
                  />
                  <span className="text-xs text-slate-400 font-medium">Skip First Row (Header)</span>
                </label>
                
                <div>
                  <label className="text-[10px] text-slate-500 uppercase font-bold block mb-1.5">ID Column</label>
                  <select 
                    className="w-full bg-slate-800 border border-slate-700 text-xs px-2 py-2 rounded focus:ring-1 ring-indigo-500 outline-none text-slate-200"
                    value={selectedColumnIndex}
                    onChange={e => setSelectedColumnIndex(parseInt(e.target.value))}
                    disabled={isRunning}
                  >
                    {excelData[0]?.map((col: any, idx: number) => (
                      <option key={idx} value={idx}>
                        Col {String.fromCharCode(65 + idx)} {skipHeader && col ? `(${String(col).substring(0,20)})` : ''}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
            )}
          </div>

          <div className="mt-auto flex flex-col gap-3 pt-6">
            <button
              onClick={toggleRunning}
              disabled={tasks.length === 0 || (!isRunning && stats.running > 0)}
              className={cn(
                "w-full font-bold py-3 rounded-xl transition-all shadow-lg text-sm flex items-center justify-center gap-2",
                isRunning ? "bg-amber-600 hover:bg-amber-500 text-white shadow-amber-600/20" : "bg-indigo-600 hover:bg-indigo-500 text-white shadow-indigo-600/20",
                (tasks.length === 0 || (!isRunning && stats.running > 0)) && "opacity-50 cursor-not-allowed bg-slate-800 hover:bg-slate-800 shadow-none text-slate-500"
              )}
            >
              {isRunning ? (
                <><Pause className="w-4 h-4" /> Pause Execution</>
              ) : stats.running > 0 ? (
                <><Loader2 className="w-4 h-4 animate-spin" /> Single Run Active</>
              ) : (
                <><Play className="w-4 h-4" /> {stats.pending === tasks.length ? 'Execute Batch Process' : 'Resume Execution'}</>
              )}
            </button>
            
            {tasks.length > 0 && !isRunning && (
              <button 
                onClick={clearData}
                className="w-full bg-slate-800 hover:bg-slate-700 text-rose-400 hover:text-rose-300 font-bold py-3 rounded-xl transition-all text-sm flex items-center justify-center gap-2"
              >
                <Trash2 className="w-4 h-4" />
                Clear Data
              </button>
            )}
          </div>
        </div>

        {/* Statistics Card (Top Right) */}
        <div className="col-span-8 row-span-3 bg-slate-900 border border-slate-800 rounded-2xl p-5 flex flex-col justify-between shadow-xl">
          <div className="flex justify-between items-start">
            <div>
              <p className="text-[10px] text-slate-500 uppercase font-bold tracking-widest">Task Progress</p>
              <h3 className="text-2xl font-bold text-white">{(stats.success + stats.failed)} <span className="text-slate-500 text-sm font-normal">/ {stats.total}</span></h3>
            </div>
            <div className="flex gap-4">
              <div className="text-right text-emerald-500">
                <p className="text-[10px] uppercase font-bold opacity-70">Success</p>
                <p className="text-lg font-bold">{stats.success}</p>
              </div>
              <div className="text-right text-rose-500">
                <p className="text-[10px] uppercase font-bold opacity-70">Failed</p>
                <p className="text-lg font-bold">{stats.failed}</p>
              </div>
              <div className="text-right text-indigo-400">
                <p className="text-[10px] uppercase font-bold opacity-70">Running</p>
                <p className="text-lg font-bold">{stats.running}</p>
              </div>
              <div className="text-right text-slate-400">
                <p className="text-[10px] uppercase font-bold opacity-70">Pending</p>
                <p className="text-lg font-bold">{stats.pending}</p>
              </div>
            </div>
          </div>
          <div className="w-full bg-slate-800 h-3 rounded-full overflow-hidden mt-4">
            <div 
              className="bg-indigo-500 h-full shadow-[0_0_12px_rgba(99,102,241,0.5)] transition-all duration-300 ease-in-out" 
              style={{ width: `${progressPercent}%` }}
            />
          </div>
        </div>

        {/* Main Table View (Bottom Right) */}
        <div className="col-span-8 row-span-9 bg-slate-900/40 border border-slate-800 rounded-2xl overflow-hidden flex flex-col">
          <div className="bg-slate-800/50 px-4 py-2 border-b border-slate-700 flex justify-between items-center">
            <span className="text-[10px] font-bold uppercase text-slate-400 tracking-wider">Active Execution Queue</span>
          </div>
          
          <div className="flex-1 overflow-y-auto font-mono text-[11px] custom-scrollbar">
            {tasks.length === 0 ? (
              <div className="h-full flex flex-col items-center justify-center text-slate-500">
                <p>No Data Loaded</p>
              </div>
            ) : (
              <table className="w-full text-left">
                <thead className="bg-slate-900 text-slate-500 sticky top-0 z-10 shadow-sm">
                  <tr>
                    <th className="p-3 font-normal border-b border-slate-800 w-12">#</th>
                    <th className="p-3 font-normal border-b border-slate-800 w-48">League ID</th>
                    <th className="p-3 font-normal border-b border-slate-800 w-32">Status</th>
                    <th className="p-3 font-normal border-b border-slate-800 w-24">Action</th>
                    <th className="p-3 font-normal border-b border-slate-800">Report</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/50">
                  {tasks.map((task, idx) => (
                    <tr key={idx} className={cn(
                      "transition-colors",
                      task.status === 'success' ? "bg-slate-800/20" : 
                      task.status === 'failed' ? "bg-rose-500/5" : 
                      task.status === 'running' ? "bg-slate-800/40" : "opacity-70 hover:bg-slate-800/20"
                    )}>
                      <td className="p-3 text-slate-500">{task.originalIndex + 1}</td>
                      <td className="p-3">
                        <div className="text-indigo-400 font-bold">{task.id}</div>
                        <div className="text-[9px] text-slate-500 truncate max-w-[220px]" title={getConstructedUrl(task.id)}>
                          {getConstructedUrl(task.id)}
                        </div>
                      </td>
                      <td className="p-3">
                        {task.status === 'success' && <span className="bg-emerald-500/10 text-emerald-500 px-2 py-0.5 rounded flex items-center w-fit gap-1"><CheckCircle2 className="w-3 h-3"/> SUCCESS</span>}
                        {task.status === 'failed' && <span className="bg-rose-500/10 text-rose-500 px-2 py-0.5 rounded flex items-center w-fit gap-1"><XCircle className="w-3 h-3"/> FAILED</span>}
                        {task.status === 'running' && <span className="flex items-center gap-2 text-indigo-400"><Loader2 className="w-3 h-3 animate-spin"/> Processing</span>}
                        {task.status === 'pending' && <span className="text-slate-500 flex items-center gap-1"><Clock className="w-3 h-3"/> PENDING</span>}
                      </td>
                      <td className="p-3">
                        <button
                          onClick={() => handleManualRun(idx)}
                          disabled={isRunning || task.status === 'running'}
                          className="text-slate-500 hover:text-white disabled:opacity-30 flex items-center gap-1"
                        >
                          {task.status === 'failed' ? <><RotateCcw className="w-3 h-3"/> Retry</> : <><Play className="w-3 h-3"/> Run</>}
                        </button>
                      </td>
                      <td className="p-3 text-[10px] text-slate-400">
                        <div className="max-h-24 overflow-y-auto custom-scrollbar whitespace-pre-wrap font-mono">
                          {task.report || "-"}
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </div>

      </div>

      {/* Footer Bar */}
      <footer className="shrink-0 flex items-center justify-between text-[11px] text-slate-500 border-t border-slate-800/50 pt-4">
        <div className="flex items-center gap-6">
          <p>Active Thread Pool: <span className="text-indigo-400">0x0{concurrency}</span></p>
          <p>Success Rate: <span className="text-emerald-500">{stats.total > 0 ? ((stats.success / stats.total) * 100).toFixed(2) : '0.00'}%</span></p>
        </div>
        <div className="flex gap-4">
          <span className="px-2 py-0.5 bg-slate-800 rounded">System OK</span>
          <span className="px-2 py-0.5 bg-indigo-500/10 text-indigo-400 rounded">Auto-Analyze Enabled</span>
        </div>
      </footer>
    </div>
  );
}
