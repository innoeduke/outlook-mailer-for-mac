import React, { useState, useEffect, useRef } from 'react';
import Papa from 'papaparse';
import { marked } from 'marked';
import { 
  Mail, 
  FileSpreadsheet, 
  Play, 
  Pause, 
  Square, 
  RotateCcw, 
  RefreshCw, 
  AlertCircle, 
  CheckCircle2, 
  Settings, 
  Eye, 
  HelpCircle,
  Clock,
  Send
} from 'lucide-react';

interface LogEntry {
  timestamp: string;
  type: 'info' | 'success' | 'error' | 'warning';
  message: string;
}

export default function App() {
  // Outlook Running Status
  const [isOutlookRunning, setIsOutlookRunning] = useState<boolean | null>(null);
  const [isCheckingStatus, setIsCheckingStatus] = useState<boolean>(false);

  // CSV Data
  const [csvFile, setCsvFile] = useState<File | null>(null);
  const [csvHeaders, setCsvHeaders] = useState<string[]>([]);
  const [csvRows, setCsvRows] = useState<Record<string, string>[]>([]);
  const [isDragging, setIsDragging] = useState(false);

  // Mappings
  const [emailColumn, setEmailColumn] = useState<string>('');
  const [nameColumn, setNameColumn] = useState<string>('');

  // Email Template
  const [emailSubject, setEmailSubject] = useState<string>('Greetings {{First Name}}!');
  const [emailBody, setEmailBody] = useState<string>(
    `<p>Dear {{First Name}},</p>\n<p>I hope this email finds you well.</p>\n<p>This is a batch test email sent via Outlook for Mac.</p>\n<p>Best regards,<br>Kyle</p>`
  );
  const [editorMode, setEditorMode] = useState<'html' | 'markdown'>('html');

  // Preview Index
  const [previewIndex, setPreviewIndex] = useState<number>(0);

  // Sending State
  const [sendingStatus, setSendingStatus] = useState<'idle' | 'sending' | 'paused' | 'done'>('idle');
  const [currentIndex, setCurrentIndex] = useState<number>(0);
  const [successCount, setSuccessCount] = useState<number>(0);
  const [errorCount, setErrorCount] = useState<number>(0);
  const [sendingDelay, setSendingDelay] = useState<number>(3); // seconds
  const [logs, setLogs] = useState<LogEntry[]>([]);

  // Refs for tracking mutable states in the sending loop
  const isSendingRef = useRef<boolean>(false);
  const currentIndexRef = useRef<number>(0);
  const csvRowsRef = useRef<Record<string, string>[]>([]);
  const emailColumnRef = useRef<string>('');
  const sendingDelayRef = useRef<number>(3);

  // Sync refs with state
  useEffect(() => {
    csvRowsRef.current = csvRows;
  }, [csvRows]);

  useEffect(() => {
    emailColumnRef.current = emailColumn;
  }, [emailColumn]);

  useEffect(() => {
    sendingDelayRef.current = sendingDelay;
  }, [sendingDelay]);

  // Check Outlook Status on Mount and Poll
  const checkOutlookStatus = async () => {
    setIsCheckingStatus(true);
    try {
      const response = await fetch('http://localhost:3001/api/status');
      const data = await response.json();
      setIsOutlookRunning(data.running);
    } catch (error) {
      console.error('Failed to contact backend status endpoint:', error);
      setIsOutlookRunning(false);
    } finally {
      setIsCheckingStatus(false);
    }
  };

  useEffect(() => {
    checkOutlookStatus();
    const interval = setInterval(checkOutlookStatus, 5000);
    return () => clearInterval(interval);
  }, []);

  // Add Log Entry
  const addLog = (type: LogEntry['type'], message: string) => {
    const time = new Date().toLocaleTimeString();
    setLogs((prev) => [{ timestamp: time, type, message }, ...prev]);
  };

  // CSV Drag and Drop Handlers
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      handleCsvFile(files[0]);
    }
  };

  const insertTextAtCursor = (textToInsert: string) => {
    const textarea = document.getElementById('body-input') as HTMLTextAreaElement;
    if (!textarea) return;

    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const text = textarea.value;
    const before = text.substring(0, start);
    const after  = text.substring(end, text.length);

    setEmailBody(before + textToInsert + after);
    
    setTimeout(() => {
      textarea.focus();
      textarea.selectionStart = textarea.selectionEnd = start + textToInsert.length;
    }, 10);
  };

  const handleImageUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const currentMode = editorMode;

    if (file.size > 2 * 1024 * 1024) {
      addLog('warning', 'Image file is larger than 2MB. Large embedded images might be blocked by some email clients.');
    }

    const reader = new FileReader();
    reader.onload = (event) => {
      const dataUrl = event.target?.result as string;
      const imgTag = currentMode === 'markdown'
        ? `![${file.name}](${dataUrl})`
        : `<img src="${dataUrl}" alt="${file.name}" style="max-width: 100%; height: auto; display: block; margin: 1rem 0;" />`;
      insertTextAtCursor(imgTag);
      addLog('success', `Inserted image: ${file.name}`);
    };
    reader.onerror = () => {
      addLog('error', 'Failed to read image file.');
    };
    reader.readAsDataURL(file);
    e.target.value = '';
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      handleCsvFile(files[0]);
    }
  };

  const handleCsvFile = (file: File) => {
    if (!file.name.endsWith('.csv')) {
      addLog('error', 'Invalid file. Please upload a .csv file.');
      return;
    }

    setCsvFile(file);
    addLog('info', `Loading file: ${file.name}`);

    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (results) => {
        if (results.data && results.data.length > 0) {
          const headers = Object.keys(results.data[0]);
          setCsvHeaders(headers);
          
          // Cast data as array of records
          const rows = results.data as Record<string, string>[];
          setCsvRows(rows);
          setPreviewIndex(0);
          setCurrentIndex(0);
          currentIndexRef.current = 0;
          setSuccessCount(0);
          setErrorCount(0);
          setSendingStatus('idle');

          // Attempt smart matching for email and name columns
          const emailMatch = headers.find(h => /email/i.test(h));
          const nameMatch = headers.find(h => /name/i.test(h) || /first/i.test(h));
          
          if (emailMatch) setEmailColumn(emailMatch);
          else setEmailColumn(headers[0] || '');

          if (nameMatch) setNameColumn(nameMatch);
          else setNameColumn(headers[1] || headers[0] || '');

          addLog('success', `Parsed ${rows.length} contacts from CSV.`);
        } else {
          addLog('error', 'CSV file is empty or has no headers.');
        }
      },
      error: (error) => {
        addLog('error', `CSV Parse Error: ${error.message}`);
      }
    });
  };

  // Replace placeholders in email template
  const interpolateTemplate = (template: string, row: Record<string, string>): string => {
    if (!row) return template;
    let result = template;
    Object.entries(row).forEach(([key, value]) => {
      const escapedKey = key.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
      const regex = new RegExp(`{{\\s*${escapedKey}\\s*}}`, 'gi');
      result = result.replace(regex, value || '');
    });
    return result;
  };

  // Preview fields
  const currentPreviewRow = csvRows[previewIndex];
  const renderedSubject = currentPreviewRow ? interpolateTemplate(emailSubject, currentPreviewRow) : '';
  const renderedBody = currentPreviewRow 
    ? (editorMode === 'markdown' 
        ? (marked.parse(interpolateTemplate(emailBody, currentPreviewRow)) as string)
        : interpolateTemplate(emailBody, currentPreviewRow))
    : '';

  // Sending Loop
  const sendBatch = async () => {
    isSendingRef.current = true;
    setSendingStatus('sending');
    addLog('info', 'Starting batch send operation...');

    while (isSendingRef.current && currentIndexRef.current < csvRowsRef.current.length) {
      const idx = currentIndexRef.current;
      const row = csvRowsRef.current[idx];
      const recipient = row[emailColumnRef.current];

      if (!recipient) {
        addLog('warning', `[Row ${idx + 1}] Skipped: No email address found in mapped column.`);
        currentIndexRef.current = idx + 1;
        setCurrentIndex(idx + 1);
        continue;
      }

      // Check if Outlook is still running before sending this email
      let statusOk = false;
      try {
        const statusRes = await fetch('http://localhost:3001/api/status');
        const statusData = await statusRes.json();
        statusOk = statusData.running;
        setIsOutlookRunning(statusOk);
      } catch (err) {
        statusOk = false;
      }

      if (!statusOk) {
        addLog('error', 'Send paused: Microsoft Outlook is not open/running on your Mac.');
        pauseSending();
        break;
      }

      // Prepare template values
      const subject = interpolateTemplate(emailSubject, row);
      const rawBody = interpolateTemplate(emailBody, row);
      const body = editorMode === 'markdown' ? (marked.parse(rawBody) as string) : rawBody;

      addLog('info', `Sending email to ${recipient}...`);

      try {
        const response = await fetch('http://localhost:3001/api/send-email', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({ recipient, subject, body })
        });

        const data = await response.json();

        if (response.ok && data.success) {
          setSuccessCount(prev => prev + 1);
          addLog('success', `[${idx + 1}/${csvRowsRef.current.length}] Successfully sent email to ${recipient}`);
        } else {
          setErrorCount(prev => prev + 1);
          addLog('error', `[${idx + 1}/${csvRowsRef.current.length}] Failed sending to ${recipient}: ${data.error || 'Unknown error'}`);
        }
      } catch (error: any) {
        setErrorCount(prev => prev + 1);
        addLog('error', `[${idx + 1}/${csvRowsRef.current.length}] Connection error sending to ${recipient}: ${error.message}`);
      }

      // Advance index
      const nextIdx = idx + 1;
      currentIndexRef.current = nextIdx;
      setCurrentIndex(nextIdx);

      // Throttling Delay
      if (nextIdx < csvRowsRef.current.length && isSendingRef.current) {
        const delayMs = sendingDelayRef.current * 1000;
        addLog('info', `Waiting ${sendingDelayRef.current}s before next send...`);
        await new Promise(resolve => setTimeout(resolve, delayMs));
      }
    }

    if (currentIndexRef.current >= csvRowsRef.current.length) {
      setSendingStatus('done');
      isSendingRef.current = false;
      addLog('success', 'Batch send operation completed.');
    }
  };

  const pauseSending = () => {
    isSendingRef.current = false;
    setSendingStatus('paused');
    addLog('warning', 'Sending execution paused.');
  };

  const stopSending = () => {
    isSendingRef.current = false;
    setSendingStatus('idle');
    currentIndexRef.current = 0;
    setCurrentIndex(0);
    setSuccessCount(0);
    setErrorCount(0);
    addLog('info', 'Sending execution stopped and counters reset.');
  };

  const resumeSending = () => {
    if (sendingStatus === 'paused') {
      sendBatch();
    }
  };

  const resetAll = () => {
    stopSending();
    setCsvFile(null);
    setCsvHeaders([]);
    setCsvRows([]);
    setLogs([]);
    addLog('info', 'Workspace cleared.');
  };

  const downloadDemoCsv = () => {
    const headers = ['First Name', 'Last Name', 'Email', 'Company'];
    const rows = [
      ['Kyle', 'Wu', 'kyle.wu@example.com', 'Leadership Toastmasters'],
      ['John', 'Doe', 'john.doe@example.com', 'Acme Corp'],
      ['Jane', 'Smith', 'jane.smith@example.com', 'Tech Innovators']
    ];
    
    const csvContent = [
      headers.join(','),
      ...rows.map(row => row.map(val => `"${val.replace(/"/g, '""')}"`).join(','))
    ].join('\n');
    
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', 'demo_contacts.csv');
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    addLog('info', 'Downloaded demo CSV template (demo_contacts.csv).');
  };

  const progressPercent = csvRows.length > 0 ? (currentIndex / csvRows.length) * 100 : 0;

  return (
    <div className="app-container">
      <header>
        <div className="logo-container">
          <span className="logo-icon">✉️</span>
          <div>
            <h1>Outlook Batch Mailer</h1>
            <p style={{ fontSize: '0.85rem', color: 'var(--text-secondary)' }}>Automated AppleScript-based email merge for macOS</p>
          </div>
        </div>
        <div className="status-badge">
          <div className={`status-dot ${isOutlookRunning ? 'active' : 'inactive'}`}></div>
          <span>Outlook Status: <strong>{isOutlookRunning === null ? 'Checking...' : isOutlookRunning ? 'Connected' : 'Not Running'}</strong></span>
          <button 
            onClick={checkOutlookStatus} 
            disabled={isCheckingStatus} 
            className="btn btn-secondary" 
            style={{ padding: '0.3rem 0.6rem', fontSize: '0.75rem', borderRadius: '0.5rem' }}
          >
            <RefreshCw size={12} className={isCheckingStatus ? 'animate-spin' : ''} />
          </button>
        </div>
      </header>

      {!isOutlookRunning && isOutlookRunning !== null && (
        <div style={{
          background: 'rgba(244, 63, 94, 0.1)',
          border: '1px solid var(--error)',
          borderRadius: '1rem',
          padding: '1.25rem',
          marginBottom: '2rem',
          display: 'flex',
          gap: '1rem',
          alignItems: 'center'
        }}>
          <AlertCircle color="var(--error)" size={24} style={{ flexShrink: 0 }} />
          <div>
            <h4 style={{ color: 'var(--error)', fontWeight: '600', marginBottom: '0.25rem' }}>Microsoft Outlook App Not Detected</h4>
            <p style={{ fontSize: '0.9rem', color: 'var(--text-secondary)', lineHeight: '1.4' }}>
              Please open the Microsoft Outlook desktop app on your Mac and ensure it is logged in to your account.
              The application will automatically connect and turn green when Outlook is active.
            </p>
          </div>
        </div>
      )}

      <div className="dashboard-grid">
        {/* Left Side: Upload & Configuration */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
          
          {/* Card 1: CSV Upload */}
          <div className="card">
            <div className="card-title">
              <FileSpreadsheet size={20} />
              <span>1. Import Contacts CSV</span>
            </div>

            {!csvFile ? (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem' }}>
                <div 
                  className={`upload-zone ${isDragging ? 'dragging' : ''}`}
                  onDragOver={handleDragOver}
                  onDragLeave={handleDragLeave}
                  onDrop={handleDrop}
                  onClick={() => document.getElementById('csv-input')?.click()}
                >
                  <FileSpreadsheet className="upload-icon" />
                  <p style={{ fontWeight: '500' }}>Drag & drop your contacts CSV file here</p>
                  <p style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>or click to browse local files</p>
                  <input 
                    type="file" 
                    id="csv-input" 
                    accept=".csv" 
                    onChange={handleFileChange} 
                  />
                </div>
                <div style={{ textAlign: 'center' }}>
                  <button 
                    onClick={downloadDemoCsv}
                    className="btn btn-secondary" 
                    style={{ fontSize: '0.85rem', width: '100%', padding: '0.6rem' }}
                  >
                    📥 Download Demo CSV Template
                  </button>
                </div>
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem' }}>
                <div style={{
                  display: 'flex',
                  justifyContent: 'space-between',
                  alignItems: 'center',
                  background: 'rgba(255, 255, 255, 0.03)',
                  padding: '1rem',
                  borderRadius: '0.75rem',
                  border: '1px solid var(--border-light)'
                }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
                    <CheckCircle2 color="var(--success)" size={20} />
                    <div>
                      <p style={{ fontSize: '0.95rem', fontWeight: '500' }}>{csvFile.name}</p>
                      <p style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>{csvRows.length} contacts found</p>
                    </div>
                  </div>
                  <button onClick={resetAll} className="btn btn-secondary" style={{ padding: '0.5rem 1rem', fontSize: '0.85rem' }}>
                    Clear
                  </button>
                </div>

                {/* Column Mappings */}
                <div className="form-group">
                  <label>Field Mapping Config</label>
                  <div className="mapping-grid">
                    <div className="form-group">
                      <label style={{ fontSize: '0.75rem' }}>Recipient Email Column</label>
                      <select 
                        value={emailColumn} 
                        onChange={(e) => setEmailColumn(e.target.value)}
                        disabled={sendingStatus === 'sending'}
                      >
                        {csvHeaders.map(header => (
                          <option key={header} value={header}>{header}</option>
                        ))}
                      </select>
                    </div>
                    <div className="form-group">
                      <label style={{ fontSize: '0.75rem' }}>First Name / Reference Column</label>
                      <select 
                        value={nameColumn} 
                        onChange={(e) => setNameColumn(e.target.value)}
                        disabled={sendingStatus === 'sending'}
                      >
                        {csvHeaders.map(header => (
                          <option key={header} value={header}>{header}</option>
                        ))}
                      </select>
                    </div>
                  </div>
                </div>

                {/* Available Variables Info */}
                <div style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>
                  <strong>Available dynamic variables:</strong>
                  <div style={{ display: 'flex', gap: '0.5rem', flexWrap: 'wrap', marginTop: '0.4rem' }}>
                    {csvHeaders.map(h => (
                      <span key={h} className="tag tag-indigo">{`{{${h}}}`}</span>
                    ))}
                  </div>
                </div>

                {/* Micro contacts list */}
                <div className="contacts-table-wrapper">
                  <table>
                    <thead>
                      <tr>
                        <th>#</th>
                        <th>Email ({emailColumn})</th>
                        <th>Name ({nameColumn})</th>
                      </tr>
                    </thead>
                    <tbody>
                      {csvRows.slice(0, 50).map((row, idx) => (
                        <tr key={idx} style={idx === currentIndex ? { background: 'rgba(99, 102, 241, 0.1)' } : {}}>
                          <td>{idx + 1}</td>
                          <td>{row[emailColumn] || <span style={{ color: 'var(--error)' }}>Missing</span>}</td>
                          <td>{row[nameColumn] || '-'}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>

          {/* Card 2: Email Template Composer */}
          <div className="card">
            <div className="card-title">
              <Mail size={20} />
              <span>2. Compose Email Template</span>
            </div>

            <div className="form-group">
              <label htmlFor="subject-input">Subject Line</label>
              <input 
                type="text" 
                id="subject-input" 
                value={emailSubject} 
                onChange={(e) => setEmailSubject(e.target.value)}
                placeholder="Greetings {{First Name}}!"
                disabled={sendingStatus === 'sending'}
              />
            </div>

            <div className="form-group">
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '0.5rem' }}>
                <label htmlFor="body-input">Email Body ({editorMode === 'markdown' ? 'Markdown' : 'HTML'})</label>
                <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
                  {/* HTML / Markdown Toggler */}
                  <div style={{ display: 'flex', gap: '0.2rem', background: 'rgba(255,255,255,0.04)', padding: '0.2rem', borderRadius: '0.5rem', border: '1px solid var(--border-light)' }}>
                    <button
                      type="button"
                      onClick={() => setEditorMode('html')}
                      className={`btn ${editorMode === 'html' ? 'btn-primary' : 'btn-secondary'}`}
                      style={{ padding: '0.25rem 0.6rem', fontSize: '0.75rem', borderRadius: '0.35rem', boxShadow: 'none' }}
                      disabled={sendingStatus === 'sending'}
                    >
                      HTML
                    </button>
                    <button
                      type="button"
                      onClick={() => setEditorMode('markdown')}
                      className={`btn ${editorMode === 'markdown' ? 'btn-primary' : 'btn-secondary'}`}
                      style={{ padding: '0.25rem 0.6rem', fontSize: '0.75rem', borderRadius: '0.35rem', boxShadow: 'none' }}
                      disabled={sendingStatus === 'sending'}
                    >
                      Markdown
                    </button>
                  </div>

                  <button
                    type="button"
                    onClick={() => document.getElementById('image-uploader')?.click()}
                    className="btn btn-secondary"
                    style={{ padding: '0.35rem 0.75rem', fontSize: '0.75rem', borderRadius: '0.4rem' }}
                    disabled={sendingStatus === 'sending'}
                  >
                    🖼️ Insert Image
                  </button>
                  <input
                    type="file"
                    id="image-uploader"
                    accept="image/*"
                    onChange={handleImageUpload}
                    style={{ display: 'none' }}
                  />
                </div>
              </div>
              <textarea 
                id="body-input" 
                value={emailBody} 
                onChange={(e) => setEmailBody(e.target.value)}
                placeholder="<p>Dear {{First Name}},</p>\n<p>Body here...</p>"
                disabled={sendingStatus === 'sending'}
              />
            </div>
          </div>
        </div>

        {/* Right Side: Preview & Progress Log */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
          
          {/* Card 3: Preview Panel */}
          <div className="card">
            <div className="card-title">
              <Eye size={20} />
              <span>3. Live Email Preview</span>
            </div>

            {csvRows.length === 0 ? (
              <div style={{ 
                height: '250px', 
                border: '1px dashed var(--border-light)',
                borderRadius: '0.75rem',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                color: 'var(--text-muted)',
                fontSize: '0.9rem'
              }}>
                Upload a CSV to view customized previews
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem' }}>
                <div className="preview-container">
                  <div className="preview-header">
                    <div className="preview-field">
                      <span className="preview-field-label">To:</span>
                      <span>{currentPreviewRow?.[emailColumn] || '(No Email Found)'}</span>
                    </div>
                    <div className="preview-field" style={{ fontWeight: '600' }}>
                      <span className="preview-field-label">Subject:</span>
                      <span>{renderedSubject}</span>
                    </div>
                  </div>
                  <div 
                    className="preview-body-html"
                    dangerouslySetInnerHTML={{ __html: renderedBody }}
                  />
                </div>

                <div className="preview-nav">
                  <span style={{ fontSize: '0.85rem', color: 'var(--text-secondary)' }}>
                    Previewing contact <strong>{previewIndex + 1}</strong> of {csvRows.length}
                  </span>
                  <div style={{ display: 'flex', gap: '0.5rem' }}>
                    <button 
                      className="btn btn-secondary" 
                      style={{ padding: '0.4rem 0.8rem', fontSize: '0.85rem' }}
                      disabled={previewIndex === 0}
                      onClick={() => setPreviewIndex(prev => prev - 1)}
                    >
                      Previous
                    </button>
                    <button 
                      className="btn btn-secondary" 
                      style={{ padding: '0.4rem 0.8rem', fontSize: '0.85rem' }}
                      disabled={previewIndex === csvRows.length - 1}
                      onClick={() => setPreviewIndex(prev => prev + 1)}
                    >
                      Next
                    </button>
                  </div>
                </div>
              </div>
            )}
          </div>

          {/* Card 4: Send Panel & Logs */}
          <div className="card">
            <div className="card-title">
              <Send size={20} />
              <span>4. Execution Controls & Logs</span>
            </div>

            {/* Delay Speed setting */}
            <div className="form-group" style={{ background: 'rgba(255,255,255,0.02)', padding: '1rem', borderRadius: '0.75rem', border: '1px solid var(--border-light)' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <label style={{ display: 'flex', alignItems: 'center', gap: '0.4rem' }}>
                  <Clock size={14} />
                  <span>Delay between sends:</span>
                </label>
                <span style={{ fontSize: '0.9rem', color: 'var(--primary)', fontWeight: '600' }}>{sendingDelay} seconds</span>
              </div>
              <input 
                type="range" 
                min="1" 
                max="10" 
                value={sendingDelay} 
                onChange={(e) => setSendingDelay(parseInt(e.target.value))}
                disabled={sendingStatus === 'sending'}
                style={{ width: '100%', marginTop: '0.5rem', cursor: 'pointer' }}
              />
              <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)', marginTop: '0.2rem', display: 'block' }}>
                A delay of 3+ seconds is recommended to allow Outlook for Mac time to queue message transfers cleanly.
              </span>
            </div>

            {/* Progress Bar */}
            <div className="progress-container">
              <div className="progress-stats">
                <span>Progress: {currentIndex} / {csvRows.length}</span>
                <span>{progressPercent.toFixed(0)}%</span>
              </div>
              <div className="progress-bar-wrapper">
                <div className="progress-bar-fill" style={{ width: `${progressPercent}%` }}></div>
              </div>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem', marginTop: '0.5rem', textAlign: 'center' }}>
                <div style={{ background: 'rgba(16, 185, 129, 0.05)', padding: '0.5rem', borderRadius: '0.5rem', border: '1px solid rgba(16, 185, 129, 0.1)' }}>
                  <div style={{ fontSize: '0.75rem', color: 'var(--text-secondary)' }}>Successfully Sent</div>
                  <div style={{ fontSize: '1.25rem', fontWeight: '700', color: 'var(--success)' }}>{successCount}</div>
                </div>
                <div style={{ background: 'rgba(244, 63, 94, 0.05)', padding: '0.5rem', borderRadius: '0.5rem', border: '1px solid rgba(244, 63, 94, 0.1)' }}>
                  <div style={{ fontSize: '0.75rem', color: 'var(--text-secondary)' }}>Failed Sends</div>
                  <div style={{ fontSize: '1.25rem', fontWeight: '700', color: 'var(--error)' }}>{errorCount}</div>
                </div>
              </div>
            </div>

            {/* Control Buttons */}
            <div className="control-buttons">
              {sendingStatus === 'idle' || sendingStatus === 'done' ? (
                <button 
                  className="btn btn-primary" 
                  style={{ flex: 1 }}
                  onClick={sendBatch}
                  disabled={csvRows.length === 0 || !isOutlookRunning}
                >
                  <Play size={16} />
                  <span>Start Batch Send</span>
                </button>
              ) : null}

              {sendingStatus === 'sending' ? (
                <button 
                  className="btn btn-secondary" 
                  style={{ flex: 1 }}
                  onClick={pauseSending}
                >
                  <Pause size={16} />
                  <span>Pause Send</span>
                </button>
              ) : null}

              {sendingStatus === 'paused' ? (
                <button 
                  className="btn btn-primary" 
                  style={{ flex: 1 }}
                  onClick={resumeSending}
                  disabled={!isOutlookRunning}
                >
                  <Play size={16} />
                  <span>Resume Send</span>
                </button>
              ) : null}

              {sendingStatus !== 'idle' ? (
                <button 
                  className="btn btn-danger" 
                  style={{ flexShrink: 0 }}
                  onClick={stopSending}
                >
                  <Square size={16} />
                  <span>Stop & Reset</span>
                </button>
              ) : null}
            </div>

            {/* Live Logs */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
              <label>Console Activity Log</label>
              <div className="log-console">
                {logs.length === 0 ? (
                  <div style={{ color: 'var(--text-muted)', textAlign: 'center', marginTop: '2rem' }}>
                    Activity events will be printed here in real-time
                  </div>
                ) : (
                  logs.map((log, index) => (
                    <div key={index} className={`log-entry ${log.type}`}>
                      <span className="log-timestamp">[{log.timestamp}]</span>
                      <span className="log-msg">{log.message}</span>
                    </div>
                  ))
                )}
              </div>
            </div>

          </div>
        </div>
      </div>
    </div>
  );
}
