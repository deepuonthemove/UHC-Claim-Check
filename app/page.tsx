'use client';

import { FormEvent, useMemo, useRef, useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { fetchEventSource } from '@microsoft/fetch-event-source';

// ── Bot columns ───────────────────────────────────────────────────────────────
const BOT_HEADERS = [
  'BotClaimNumber', 'BotClaimStatus', 'BotPaidAmount', 'BotBilledAmount',
  'BotCheckEFTNumber', 'BotDenialReasonCode', 'BotDenialDescription',
  'BotRemarkCodes', 'BotProcessedDate', 'BotClaimDetails',
  'BotUpdateTime', 'BotStatus', 'BotStatusError',
] as const;

// ── Apply a row_update to the in-memory ExcelJS worksheet ─────────────────────
function applyRowUpdate(
  worksheet: ExcelJS.Worksheet,
  headerMap: Map<string, number>,
  excelRowIndex: number,
  update: Record<string, string>
) {
  const row = worksheet.getRow(excelRowIndex);
  for (const [key, value] of Object.entries(update)) {
    const colNum = headerMap.get(key);
    if (colNum !== undefined) {
      row.getCell(colNum).value = value ?? null;
    }
  }
  row.commit();
}

// ── Post-process: copy Bot columns to duplicate (Subscriber No + Service Date) rows ──
function postProcessWorksheet(
  worksheet: ExcelJS.Worksheet,
  headerMap: Map<string, number>
) {
  const subNoCol   = headerMap.get('Subscriber No')   ?? headerMap.get('Member ID') ?? -1;
  const svcDateCol = headerMap.get('Service Date')    ?? headerMap.get('DOS') ?? -1;
  const botColNums = [...headerMap.entries()]
    .filter(([k]) => k.startsWith('Bot'))
    .map(([, v]) => v);

  if (subNoCol < 0 || svcDateCol < 0 || botColNums.length === 0) return;

  const groups = new Map<string, number[]>();
  worksheet.eachRow((row, rowNum) => {
    if (rowNum === 1) return;
    const sub  = String(row.getCell(subNoCol).value ?? '').trim();
    const date = String(row.getCell(svcDateCol).value ?? '').trim();
    if (!sub) return;
    const key = `${sub}|${date}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(rowNum);
  });

  let dupeCount = 0;
  for (const [, rows] of groups) {
    if (rows.length < 2) continue;
    const srcRow = worksheet.getRow(rows[0]);
    for (let i = 1; i < rows.length; i++) {
      const tgt = worksheet.getRow(rows[i]);
      for (const col of botColNums) {
        tgt.getCell(col).value = srcRow.getCell(col).value;
      }
      tgt.commit();
      dupeCount++;
    }
  }
  if (dupeCount > 0) console.log(`Post-processing: copied Bot columns to ${dupeCount} duplicate rows.`);
}

// ── Types ─────────────────────────────────────────────────────────────────────
interface ClaimRow {
  rowIndex: number;
  subscriberNo: string;
  patientDOB: string;
  serviceDate: string;
  [key: string]: unknown;
}

// ── Main Page Component ───────────────────────────────────────────────────────
export default function HomePage() {
  const [loginFile,    setLoginFile]    = useState<File | null>(null);
  const [claimFileHandle, setClaimFileHandle] = useState<FileSystemFileHandle | null>(null);
  const [fallbackClaimFile, setFallbackClaimFile] = useState<File | null>(null);
  const [claimFileName, setClaimFileName] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [status,       setStatus]       = useState('');
  const [logs,         setLogs]         = useState<string[]>([]);
  const [errorScreenshots, setErrorScreenshots] = useState<{ index: number; rowIndex?: number; attempt?: number; image: string }[]>([]);
  const [progress,     setProgress]     = useState<{ completed: number; total: number } | null>(null);
  const [browserType,  setBrowserType]  = useState<'chrome' | 'firefox'>('chrome');
  const [hasFilePickerAPI, setHasFilePickerAPI] = useState(false);

  const logsEndRef = useRef<HTMLDivElement>(null);

  // In-memory ExcelJS workbook (lives in client, mutated by row_update events)
  const excelWb    = useRef<ExcelJS.Workbook | null>(null);
  const worksheet  = useRef<ExcelJS.Worksheet | null>(null);
  const headerMap  = useRef<Map<string, number>>(new Map());
  const claimRows  = useRef<ClaimRow[]>([]);
  const consecutiveRetries = useRef(0);

  // Auto-scroll logs
  useEffect(() => {
    logsEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [logs]);

  useEffect(() => {
    setHasFilePickerAPI(typeof window !== 'undefined' && 'showOpenFilePicker' in window);
  }, []);

  const canSubmit = useMemo(
    () => Boolean(loginFile && (claimFileHandle || fallbackClaimFile) && !isProcessing),
    [loginFile, claimFileHandle, fallbackClaimFile, isProcessing]
  );

  // ── Pick claims file with File System Access API ───────────────────────────
  async function selectClaimFile() {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const [handle] = await (window as any).showOpenFilePicker({
        types: [{
          description: 'Excel Files',
          accept: {
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
            'application/vnd.ms-excel': ['.xls'],
          },
        }],
        excludeAcceptAllOption: true,
        multiple: false,
      });
      setClaimFileHandle(handle);
      const file = await handle.getFile();
      setClaimFileName(file.name);
    } catch (err) {
      if ((err as Error).name !== 'AbortError') {
        console.error('File picker error:', err);
      }
    }
  }

  // ── Queue-based in-place Excel write ──────────────────────────────────────
  const writeQueue = useRef(Promise.resolve());

  function queueWrite() {
    writeQueue.current = writeQueue.current.then(async () => {
      if (!excelWb.current || !claimFileHandle) return;
      try {
        const buf      = await excelWb.current.xlsx.writeBuffer();
        const writable = await claimFileHandle.createWritable();
        await writable.write(buf);
        await writable.close();
      } catch (err) {
        console.error('Write error:', err);
      }
    });
  }

  // ── Build header map from worksheet ────────────────────────────────────────
  function buildHeaderMap(ws: ExcelJS.Worksheet): Map<string, number> {
    const map = new Map<string, number>();
    ws.getRow(1).eachCell((cell, colNum) => {
      map.set(String(cell.value ?? '').trim(), colNum);
    });

    // Ensure Bot columns exist
    let nextCol = (ws.getRow(1).cellCount ?? 0) + 1;
    for (const col of BOT_HEADERS) {
      if (!map.has(col)) {
        ws.getRow(1).getCell(nextCol).value = col;
        map.set(col, nextCol);
        nextCol++;
      }
    }
    ws.getRow(1).commit();
    return map;
  }

  // ── Parse claim rows from SheetJS (fast, no style preservation needed) ─────
  function parseClaimRows(arrayBuffer: ArrayBuffer): ClaimRow[] {
    const xlsxWb   = XLSX.read(arrayBuffer, { type: 'array', cellDates: false });
    const sheet    = xlsxWb.Sheets[xlsxWb.SheetNames[0]];
    const rawRows  = XLSX.utils.sheet_to_json(sheet) as Record<string, unknown>[];

    return rawRows.map((row, idx) => {
      // Flexible column name matching
      const subNo = String(
        row['Subscriber No'] ?? row['Subscriber Number'] ?? row['Member ID'] ?? ''
      ).trim();
      const dobRaw = row['Patient DOB'] ?? row['DOB'] ?? row['Date Of Birth'] ?? '';
      const svcRaw = row['Service Date'] ?? row['DOS'] ?? row['Date Of Service'] ?? '';

      // Normalize dates: XLSX returns numbers (serial) or strings
      const normDate = (v: unknown): string => {
        if (!v) return '';
        if (typeof v === 'number') {
          // Excel serial date
          const d = XLSX.SSF.parse_date_code(v);
          if (d) {
            const m = String(d.m).padStart(2, '0');
            const day = String(d.d).padStart(2, '0');
            return `${m}/${day}/${d.y}`;
          }
        }
        return String(v).trim();
      };

      return {
        rowIndex:    idx + 2,  // 1-based, row 1 = header
        subscriberNo: subNo,
        patientDOB:  normDate(dobRaw),
        serviceDate: normDate(svcRaw),
        ...row,
      } as ClaimRow;
    }).filter(r => r.subscriberNo);
  }

  const processChunk = async (startIndex: number, totalRows: number) => {
    const formData = new FormData();
    formData.append('loginExcel',  loginFile!);
    formData.append('claimRows',   JSON.stringify(claimRows.current));
    formData.append('startIndex',  String(startIndex));
    formData.append('browserType', browserType);
    formData.append('attempt',     String(consecutiveRetries.current + 1));

    const ctrl = new AbortController();
    let currentCompleted = startIndex;
    let chunkHasError    = false;
    let completedSuccessfully = false;

    try {
      await fetchEventSource('/api/process', {
        method: 'POST',
        body:   formData,
        signal: ctrl.signal,

        async onopen(response) {
          if (response.ok && response.headers.get('content-type')?.includes('text/event-stream')) {
            return; // everything's good
          } else {
            const text = await response.text().catch(() => '');
            throw new Error(`Failed to establish connection: ${response.status} ${response.statusText}. ${text}`);
          }
        },

        async onmessage(ev) {
          try {
            if (!ev.data || ev.data.startsWith(':')) return;
            const event = JSON.parse(ev.data);

            if (event.type === 'log') {
              setLogs(prev => [...prev, event.message]);

            } else if (event.type === 'progress') {
              currentCompleted = event.completed ?? currentCompleted;
              setProgress({ completed: event.completed, total: event.total });

            } else if (event.type === 'row_update') {
              // Apply update to the in-memory workbook
              if (worksheet.current && event.rowIndex) {
                applyRowUpdate(worksheet.current, headerMap.current, event.rowIndex, event.update ?? {});
                queueWrite();
              }

            } else if (event.type === 'error_screenshot') {
              setErrorScreenshots(prev => [...prev, { index: event.index, rowIndex: event.rowIndex, attempt: event.attempt, image: event.image }]);
              
              // Auto-download error screenshot as well
              try {
                const byteCharacters = atob(event.image);
                const byteNumbers = new Array(byteCharacters.length);
                for (let i = 0; i < byteCharacters.length; i++) {
                  byteNumbers[i] = byteCharacters.charCodeAt(i);
                }
                const byteArray = new Uint8Array(byteNumbers);
                const blob = new Blob([byteArray], { type: 'image/jpeg' });
                const url  = URL.createObjectURL(blob);
                const a    = document.createElement('a');
                a.href     = url;
                const rowLabel = event.rowIndex ? `row_${event.rowIndex}` : (event.index === -1 ? 'login' : `row_${event.index + 2}`);
                const attemptLabel = event.attempt ? `_attempt_${event.attempt}` : '';
                a.download = `error_screenshot_${rowLabel}${attemptLabel}.jpg`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
              } catch (err) {
                console.error('Failed to auto-download screenshot', err);
              }

            } else if (event.type === 'debug_html') {
              // Auto-download debug HTML so user can inspect the browser state
              // Stagger by 1000ms to avoid browser blocking multiple concurrent downloads
              setTimeout(() => {
                try {
                  const blob = new Blob([event.html], { type: 'text/html' });
                  const url  = URL.createObjectURL(blob);
                  const a    = document.createElement('a');
                  a.href     = url;
                  const rowLabel = event.rowIndex ? `row_${event.rowIndex}` : (event.index === -1 ? 'login' : `row_${event.index + 2}`);
                  const attemptLabel = event.attempt ? `_attempt_${event.attempt}` : '';
                  a.download = `debug_dom_${rowLabel}${attemptLabel}.html`;
                  document.body.appendChild(a);
                  a.click();
                  document.body.removeChild(a);
                  URL.revokeObjectURL(url);
                } catch (err) {
                  console.error('Failed to auto-download HTML', err);
                }
              }, 1000);

            } else if (event.type === 'error') {
              setStatus(`❌ Error: ${event.message}`);
              setLogs(prev => [...prev, `❌ Error: ${event.message}`]);
              chunkHasError = true;
              ctrl.abort(); // Cancel the request and prevent auto-retry

            } else if (event.type === 'done') {
              completedSuccessfully = true;
              consecutiveRetries.current = 0; // Reset retries on successful batch completion
              ctrl.abort(); // Done processing, close SSE stream cleanly
            }
          } catch (parseErr) {
            console.error('Failed to parse SSE event', parseErr);
          }
        },

        onclose() {
          if (!completedSuccessfully && !chunkHasError) {
            chunkHasError = true;
            const msg = `❌ Connection closed unexpectedly by the server. This may indicate a Vercel timeout or server crash.`;
            setLogs(prev => [...prev, msg]);
            setStatus(msg);
          }
          // If the connection was closed by the server, abort to prevent reconnect
          ctrl.abort();
        },

        onerror(err) {
          console.error('SSE stream error:', err);
          const errorMsg = err instanceof Error ? err.message : String(err);
          setLogs(prev => [...prev, `❌ Stream error: ${errorMsg}`]);
          chunkHasError = true;
          ctrl.abort(); // Stop fetchEventSource retry
          throw err;
        },
      });

      // Wait for all queued writes to finish
      await writeQueue.current;

      // Double check: if the promise resolved but we didn't receive 'done' or 'error' event,
      // it means the stream ended abruptly without sending either.
      if (!completedSuccessfully && !chunkHasError) {
        chunkHasError = true;
        const msg = `❌ Stream ended abruptly without a terminal 'done' or 'error' event from the server.`;
        setLogs(prev => [...prev, msg]);
        setStatus(msg);
      }

    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : String(err);
      const isAbortError = 
        errorMsg.toLowerCase().includes('abort') || 
        errorMsg.includes('DOMException');

      if (ctrl.signal.aborted && !chunkHasError) {
        // Clean close
      } else if (isAbortError) {
        // Aborted due to tracked error, already logged, do nothing
      } else {
        console.error('fetchEventSource failed:', err);
        setLogs(prev => [...prev, `❌ Connection error: ${errorMsg}`]);
        chunkHasError = true;
      }
    }

    // Auto-resume, retry, or post-process
    if (chunkHasError) {
      if (consecutiveRetries.current < 3) {
        consecutiveRetries.current += 1;
        const retryDelay = 3000;
        const targetExcelRow = claimRows.current[currentCompleted]?.rowIndex ?? (currentCompleted + 2);
        setStatus(`⚠️ Failure occurred. Retrying automatically in 3s (attempt ${consecutiveRetries.current}/3) from Excel Row ${targetExcelRow}...`);
        setLogs(prev => [...prev, `⚠️ Failure occurred. Retrying automatically in 3s (attempt ${consecutiveRetries.current}/3) from Excel Row ${targetExcelRow}...`]);
        await new Promise(resolve => setTimeout(resolve, retryDelay));
        await processChunk(currentCompleted, totalRows);
      } else {
        const targetExcelRow = claimRows.current[currentCompleted]?.rowIndex ?? (currentCompleted + 2);
        setStatus(`❌ Process failed after 3 automatic retries at Excel Row ${targetExcelRow}.`);
        setLogs(prev => [...prev, `❌ Auto-retry limit (3 attempts) reached. Stopping automation.`]);
        setIsProcessing(false);
      }
      return;
    }

    if (currentCompleted < totalRows) {
      const targetExcelRow = claimRows.current[currentCompleted]?.rowIndex ?? (currentCompleted + 2);
      setStatus(`⏩ Auto-resuming from Excel Row ${targetExcelRow}...`);
      await processChunk(currentCompleted, totalRows);
    } else {
      // All batches done — run post-processing, then write once
      try {
        setStatus('🔄 Running post-processing (deduplication)...');
        postProcessWorksheet(worksheet.current!, headerMap.current);
        await writeQueue.current; // flush any pending
        queueWrite();             // write post-processed version
        await writeQueue.current;
        
        if (!claimFileHandle && excelWb.current) {
          setStatus('💾 Downloading completed Excel file...');
          const buf = await excelWb.current.xlsx.writeBuffer();
          const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = claimFileName ? `updated_${claimFileName}` : 'updated_claims.xlsx';
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          URL.revokeObjectURL(url);
          setStatus('✅ Processing complete! Excel file downloaded successfully.');
        } else {
          setStatus('✅ Processing complete! Excel file updated on disk.');
        }
      } catch (err) {
        setStatus(`⚠️ Post-processing failed: ${err instanceof Error ? err.message : String(err)}`);
      } finally {
        setIsProcessing(false);
      }
    }
  };

  async function onSubmit(e: FormEvent<HTMLFormElement>) {
    e.preventDefault();
    if (!loginFile || (!claimFileHandle && !fallbackClaimFile)) return;

    setIsProcessing(true);
    setStatus('📂 Reading claim file...');
    setLogs([]);
    setErrorScreenshots([]);
    setProgress(null);
    consecutiveRetries.current = 0;

    try {
      let file: File;
      if (claimFileHandle) {
        // Ensure readwrite permission
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const handle = claimFileHandle as any;
        if ((await handle.queryPermission({ mode: 'readwrite' })) !== 'granted') {
          if ((await handle.requestPermission({ mode: 'readwrite' })) !== 'granted') {
            throw new Error('Write permission denied. Cannot update Excel file.');
          }
        }
        file = await claimFileHandle.getFile();
      } else {
        file = fallbackClaimFile!;
      }

      // Load file into memory (both XLSX for quick parsing and ExcelJS for style-preserving writes)
      const arrayBuffer = await file.arrayBuffer();

      // SheetJS parse for claim rows (fast)
      const rows = parseClaimRows(arrayBuffer);
      if (rows.length === 0) {
        throw new Error('No valid claim rows found. Check that "Subscriber No" and "Service Date" columns exist.');
      }
      claimRows.current = rows;

      // ExcelJS load for in-memory style-preserving worksheet
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(arrayBuffer);
      excelWb.current   = wb;
      worksheet.current = wb.getWorksheet(1)!;
      headerMap.current = buildHeaderMap(worksheet.current);

      setStatus(`🚀 Starting: ${rows.length} claim row(s) to process...`);

      await processChunk(0, rows.length);
    } catch (err) {
      setStatus(`❌ Failed: ${err instanceof Error ? err.message : String(err)}`);
      setIsProcessing(false);
    }
  }

  // ── Derived state ──────────────────────────────────────────────────────────
  const pct = progress
    ? Math.round((progress.completed / progress.total) * 100)
    : 0;

  // ── Render ─────────────────────────────────────────────────────────────────
  return (
    <main className="main">
      <header className="header">
        <div className="header-inner">
          <div className="logo">
            <div className="logo-icon">UHC</div>
            <div>
              <div className="logo-title">Claim Status Automation</div>
              <div className="logo-sub">UHC Provider Portal · Playwright + Next.js</div>
            </div>
          </div>
          <div className={`status-badge ${isProcessing ? 'status-badge--running' : status.startsWith('✅') ? 'status-badge--done' : status.startsWith('❌') ? 'status-badge--error' : ''}`}>
            {isProcessing ? '◎ Running' : status.startsWith('✅') ? '✓ Complete' : status.startsWith('❌') ? '✕ Error' : '● Ready'}
          </div>
        </div>
      </header>

      <div className="page-body">
        {/* ── Left sidebar ── */}
        <aside className="sidebar">
          <form onSubmit={onSubmit} style={{ display: 'contents' }}>

            {/* Step 1: Login file */}
            <div className="card">
              <div className="card-label">Step 1 — Login Credentials</div>
              <div className="card-desc">Upload <code>UHC-Website Details.xlsx</code></div>
              <label className={`file-drop ${loginFile ? 'file-drop--loaded' : ''}`} htmlFor="login-file-input">
                <input
                  id="login-file-input"
                  type="file"
                  accept=".xlsx,.xls"
                  className="sr-only"
                  onChange={e => setLoginFile(e.target.files?.[0] ?? null)}
                />
                <span className="file-icon">{loginFile ? '✓' : '📂'}</span>
                <span>{loginFile ? loginFile.name : 'Click to select login file'}</span>
              </label>
            </div>

            {/* Step 2: Claims file */}
            <div className="card">
              <div className="card-label">Step 2 — Claims File</div>
              {hasFilePickerAPI ? (
                <>
                  <div className="card-desc">
                    Pick <code>TPm UHC claims details.xlsx</code> — updates this file in-place
                  </div>
                  <button
                    type="button"
                    className={`btn btn--secondary btn--full ${claimFileName ? 'file-drop--loaded' : ''}`}
                    onClick={selectClaimFile}
                    id="pick-claims-btn"
                  >
                    {claimFileName ? `✓  ${claimFileName}` : '📂  Pick Claims File…'}
                  </button>
                  <div className="hint">
                    Uses the File System Access API — grants direct write access so data is saved after every row.
                  </div>
                </>
              ) : (
                <>
                  <div className="card-desc">
                    Upload <code>TPm UHC claims details.xlsx</code> (Direct writing is not supported in this browser; file will download at the end)
                  </div>
                  <label className={`file-drop ${fallbackClaimFile ? 'file-drop--loaded' : ''}`} htmlFor="claim-file-input">
                    <input
                      id="claim-file-input"
                      type="file"
                      accept=".xlsx,.xls"
                      className="sr-only"
                      onChange={e => {
                        const file = e.target.files?.[0] ?? null;
                        setFallbackClaimFile(file);
                        setClaimFileName(file ? file.name : '');
                      }}
                    />
                    <span className="file-icon">{fallbackClaimFile ? '✓' : '📂'}</span>
                    <span>{fallbackClaimFile ? fallbackClaimFile.name : 'Click to select claims file'}</span>
                  </label>
                </>
              )}
            </div>

            {/* Browser choice */}
            <div className="card">
              <div className="card-label">Step 2b — Browser Type</div>
              <div className="card-desc">Choose local automation browser</div>
              <div style={{ display: 'flex', gap: '16px', marginTop: '8px' }}>
                <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', fontSize: '0.85rem' }}>
                  <input
                    type="radio"
                    name="browserType"
                    value="chrome"
                    checked={browserType === 'chrome'}
                    onChange={() => setBrowserType('chrome')}
                    style={{ cursor: 'pointer' }}
                  />
                  <span>Chrome</span>
                </label>
                <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', fontSize: '0.85rem' }}>
                  <input
                    type="radio"
                    name="browserType"
                    value="firefox"
                    checked={browserType === 'firefox'}
                    onChange={() => setBrowserType('firefox')}
                    style={{ cursor: 'pointer' }}
                  />
                  <span>Firefox</span>
                </label>
              </div>
            </div>

            {/* Step 3: Start */}
            <div className="card" style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
              <div className="card-label">Step 3 — Process</div>
              <button
                type="submit"
                id="start-btn"
                disabled={!canSubmit}
                className={`btn btn--primary btn--full ${!canSubmit ? 'btn--disabled' : ''}`}
              >
                {isProcessing ? '⏳  Processing…' : progress && progress.completed > 0 ? '🚀  Start from Beginning' : '🚀  Start Processing'}
              </button>
              
              {!isProcessing && progress && progress.completed > 0 && progress.completed < progress.total && (
                <button
                  type="button"
                  id="resume-btn"
                  className="btn btn--secondary btn--full"
                  onClick={async () => {
                    setIsProcessing(true);
                    consecutiveRetries.current = 0; // Reset retries on manual resume
                    const nextExcelRow = claimRows.current[progress.completed]?.rowIndex ?? (progress.completed + 2);
                    setStatus(`⏩ Resuming manually from Excel Row ${nextExcelRow}...`);
                    await processChunk(progress.completed, progress.total);
                  }}
                >
                  ⏩ Resume from Excel Row {claimRows.current[progress.completed]?.rowIndex ?? (progress.completed + 2)}
                </button>
              )}
            </div>

            {/* Progress bar */}
            {progress && (
              <div className="card">
                <div className="card-label">Progress</div>
                <div className="progress-bar-outer">
                  <div className="progress-bar-inner" style={{ width: `${pct}%` }} />
                </div>
                <div className="progress-text">{progress.completed} / {progress.total} rows ({pct}%)</div>
              </div>
            )}

            {/* Status */}
            {status && (
              <div className={`card ${status.startsWith('❌') ? 'card--error' : ''}`}>
                <div className="card-label">Status</div>
                <div style={{ fontSize: '0.85rem' }}>{status}</div>
              </div>
            )}

            {/* Download button for fallback mode */}
            {!claimFileHandle && progress && (
              <div className="card">
                <div className="card-label">Download Output</div>
                <button
                  type="button"
                  className="btn btn--secondary btn--full"
                  disabled={isProcessing}
                  onClick={async () => {
                    if (!excelWb.current) return;
                    const buf = await excelWb.current.xlsx.writeBuffer();
                    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = claimFileName ? `updated_${claimFileName}` : 'updated_claims.xlsx';
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    URL.revokeObjectURL(url);
                  }}
                >
                  💾 Download Updated Excel
                </button>
              </div>
            )}

            {/* Bot columns reference */}
            <div className="card card--muted">
              <div className="card-label">Output Bot Columns</div>
              <div className="bot-cols">
                {BOT_HEADERS.map(col => (
                  <span key={col} className="bot-col-badge">{col}</span>
                ))}
              </div>
            </div>
          </form>
        </aside>

        {/* ── Right: logs + screenshots ── */}
        <section className="log-section">
          {/* Live log */}
          <div className="log-container">
            <div className="log-header">
              <span className="log-title">Live Log</span>
              <span className="log-count">{logs.length} events</span>
            </div>
            <div className="log-body">
              {logs.length === 0 && (
                <div className="log-empty">Waiting for automation to start…</div>
              )}
              {logs.map((line, idx) => {
                const type = line.includes('❌') || line.includes('Failed') || line.includes('Error')
                  ? 'error'
                  : line.includes('✅') || line.includes('Success')
                  ? 'success'
                  : line.includes('ℹ️') || line.includes('Resuming') || line.includes('📊')
                  ? 'info'
                  : 'log';
                return (
                  <div key={idx} className={`log-entry log-entry--${type}`}>
                    <span className="log-time">
                      {new Date().toLocaleTimeString('en-US', { hour12: false })}
                    </span>
                    <span className="log-msg">{line}</span>
                  </div>
                );
              })}
              <div ref={logsEndRef} />
            </div>
          </div>

          {/* Error screenshots */}
          {errorScreenshots.length > 0 && (
            <div style={{ marginTop: '16px', display: 'flex', flexDirection: 'column', gap: '12px' }}>
              {errorScreenshots.map((err, i) => {
                const rowLabel = err.rowIndex ? `Excel Row ${err.rowIndex}` : err.index === -1 ? 'Login' : `Row ${err.index + 2}`;
                const attemptLabel = err.attempt ? ` — Attempt ${err.attempt}` : '';
                const downloadFilename = `error_screenshot_${err.rowIndex ? `row_${err.rowIndex}` : err.index === -1 ? 'login' : `row_${err.index + 2}`}${err.attempt ? `_attempt_${err.attempt}` : ''}.jpg`;
                
                return (
                  <div key={i} className="card card--error">
                    <div className="card-label" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <span>📸 {rowLabel}{attemptLabel} Error Screenshot</span>
                      <a
                        href={`data:image/jpeg;base64,${err.image}`}
                        download={downloadFilename}
                        className="btn btn--secondary"
                        style={{ padding: '2px 8px', fontSize: '0.75rem', width: 'auto', display: 'inline-block', textDecoration: 'none' }}
                      >
                        Download JPG
                      </a>
                    </div>
                    {/* eslint-disable-next-line @next/next/no-img-element */}
                    <img
                      src={`data:image/jpeg;base64,${err.image}`}
                      alt="Browser state at error"
                      style={{ maxWidth: '100%', borderRadius: '6px', marginTop: '8px' }}
                    />
                  </div>
                );
              })}
            </div>
          )}
        </section>
      </div>
    </main>
  );
}
