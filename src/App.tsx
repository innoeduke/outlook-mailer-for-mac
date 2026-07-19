import React, { useState, useEffect, useRef } from 'react';
import Papa from 'papaparse';
import { marked } from 'marked';
import { 
  Mail, 
  FileSpreadsheet, 
  Play, 
  Pause, 
  Square, 
  RefreshCw, 
  AlertCircle, 
  CheckCircle2, 
  Eye, 
  Clock,
  Send,
  Save,
  Plus,
  Trash2,
  History,
  Download,
  Search,
  ExternalLink,
  RotateCcw
} from 'lucide-react';

interface LogEntry {
  timestamp: string;
  type: 'info' | 'success' | 'error' | 'warning';
  message: string;
}

interface TemplateSummary {
  id: string;
  name: string;
  subject: string;
  editorMode: 'html' | 'markdown';
}

const DEFAULT_HTML_BODY = `<p>Dear {{First Name}},</p>\n<p>I hope this email finds you well.</p>\n<p>This is a batch test email sent via Outlook for Mac.</p>\n<p>Best regards,<br>Kyle</p>`;

const DEFAULT_MARKDOWN_BODY = `Dear {{First Name}},\n\nI hope this email finds you well.\n\nThis is a batch test email sent via Outlook for Mac.\n\nBest regards,\nKyle`;

const htmlToMarkdown = (html: string): string => {
  if (!html) return '';
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');

  const walk = (node: Node): string => {
    if (node.nodeType === Node.TEXT_NODE) {
      return node.textContent || '';
    }
    if (node.nodeType !== Node.ELEMENT_NODE) {
      return '';
    }

    const element = node as HTMLElement;
    const tagName = element.tagName.toUpperCase();

    let childrenContent = '';
    for (let i = 0; i < element.childNodes.length; i++) {
      childrenContent += walk(element.childNodes[i]);
    }

    switch (tagName) {
      case 'P':
        return `\n\n${childrenContent.trim()}\n\n`;
      case 'BR':
        return '\n';
      case 'STRONG':
      case 'B':
        return `**${childrenContent}**`;
      case 'EM':
      case 'I':
        return `*${childrenContent}*`;
      case 'H1':
        return `\n\n# ${childrenContent.trim()}\n\n`;
      case 'H2':
        return `\n\n## ${childrenContent.trim()}\n\n`;
      case 'H3':
        return `\n\n### ${childrenContent.trim()}\n\n`;
      case 'H4':
        return `\n\n#### ${childrenContent.trim()}\n\n`;
      case 'H5':
        return `\n\n##### ${childrenContent.trim()}\n\n`;
      case 'H6':
        return `\n\n###### ${childrenContent.trim()}\n\n`;
      case 'A': {
        const href = element.getAttribute('href') || '';
        return `[${childrenContent}](${href})`;
      }
      case 'UL':
      case 'OL':
        return `\n\n${childrenContent}\n\n`;
      case 'LI': {
        const parent = element.parentElement;
        const isOrdered = parent?.tagName.toUpperCase() === 'OL';
        if (isOrdered) {
          const index = Array.from(parent?.children || []).indexOf(element) + 1;
          return `${index}. ${childrenContent.trim()}\n`;
        }
        return `* ${childrenContent.trim()}\n`;
      }
      case 'DIV':
        return `\n${childrenContent}\n`;
      default:
        return childrenContent;
    }
  };

  let markdown = walk(doc.body);

  markdown = markdown
    .replace(/\r\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

  return markdown;
};

export default function App() {
  // Outlook Running Status
  const [isOutlookRunning, setIsOutlookRunning] = useState<boolean | null>(null);
  const [isCheckingStatus, setIsCheckingStatus] = useState<boolean>(false);
  const [outlookEmail, setOutlookEmail] = useState<string>('');

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
  const [emailBody, setEmailBody] = useState<string>(DEFAULT_HTML_BODY);
  const [editorMode, setEditorMode] = useState<'html' | 'markdown'>('html');
  const [imagesMap, setImagesMap] = useState<Record<string, string>>({});

  // Saved Templates List
  const [templatesList, setTemplatesList] = useState<TemplateSummary[]>([]);
  const [selectedTemplateId, setSelectedTemplateId] = useState<string>('');

  // Preview Index
  const [previewIndex, setPreviewIndex] = useState<number>(0);

  // Sending State
  const [sendingStatus, setSendingStatus] = useState<'idle' | 'sending' | 'paused' | 'done'>('idle');
  const [currentIndex, setCurrentIndex] = useState<number>(0);
  const [successCount, setSuccessCount] = useState<number>(0);
  const [errorCount, setErrorCount] = useState<number>(0);
  const [sendingDelay, setSendingDelay] = useState<number>(3); // seconds
  const [logs, setLogs] = useState<LogEntry[]>([]);

  // Navigation & Delivery History
  const [activeTab, setActiveTab] = useState<'dashboard' | 'history'>('dashboard');
  const [historyLogs, setHistoryLogs] = useState<any[]>([]);
  const [historyFilter, setHistoryFilter] = useState<'all' | 'success' | 'failed'>('all');
  const [selectedLogIds, setSelectedLogIds] = useState<string[]>([]);
  const [isSyncingOutlook, setIsSyncingOutlook] = useState<boolean>(false);
  const [selectedLogForDetail, setSelectedLogForDetail] = useState<any | null>(null);
  const [searchQuery, setSearchQuery] = useState<string>('');

  // Workspaces session management
  interface WorkspaceSummary {
    id: string;
    createdAt: string;
    total: number;
    successCount: number;
    errorCount: number;
  }
  const [currentWorkspaceId, setCurrentWorkspaceId] = useState<string>('');
  const [workspacesList, setWorkspacesList] = useState<WorkspaceSummary[]>([]);
  const [selectedHistoryWorkspaceId, setSelectedHistoryWorkspaceId] = useState<string>('');

  const formatWorkspaceName = (id: string) => {
    if (!id) return '';
    const isCli = id.endsWith('_CLI');
    const cleanId = id.replace('_CLI', '');
    const match = cleanId.match(/^(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})$/);
    if (match) {
      const [_, yyyy, mm, dd, hh, min, ss] = match;
      return `${yyyy}-${mm}-${dd} ${hh}:${min}:${ss}${isCli ? ' (CLI Run)' : ' (Web Run)'}`;
    }
    return id;
  };

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
      if (data.email) {
        setOutlookEmail(data.email);
      }
    } catch (error) {
      console.error('Failed to contact backend status endpoint:', error);
      setIsOutlookRunning(false);
    } finally {
      setIsCheckingStatus(false);
    }
  };

  const fetchWorkspaces = async (selectLatest = false, forceCreate = false) => {
    try {
      const response = await fetch('http://localhost:3001/api/workspaces');
      if (response.ok) {
        const data: WorkspaceSummary[] = await response.json();
        setWorkspacesList(data);
        
        if (data.length > 0) {
          if (selectLatest || !currentWorkspaceId) {
            const latest = data[0].id;
            setCurrentWorkspaceId(latest);
            setSelectedHistoryWorkspaceId(latest);
          }
        } else if (forceCreate || data.length === 0) {
          await createNewWorkspace();
        }
      }
    } catch (error) {
      console.error('Failed to retrieve workspaces:', error);
    }
  };

  const createNewWorkspace = async () => {
    try {
      const response = await fetch('http://localhost:3001/api/workspaces/new', {
        method: 'POST'
      });
      if (response.ok) {
        const data = await response.json();
        const newId = data.workspaceId;
        setCurrentWorkspaceId(newId);
        setSelectedHistoryWorkspaceId(newId);
        await fetchWorkspaces(false, false);
        return newId;
      }
    } catch (error) {
      console.error('Failed to create new workspace:', error);
    }
    return null;
  };

  const fetchHistoryLogsForWorkspace = async (workspaceId: string) => {
    if (!workspaceId) return;
    try {
      const response = await fetch(`http://localhost:3001/api/workspaces/${workspaceId}/delivery-log`);
      if (response.ok) {
        const data = await response.json();
        setHistoryLogs(data);
      }
    } catch (error) {
      console.error(`Failed to retrieve logs for workspace ${workspaceId}:`, error);
    }
  };

  useEffect(() => {
    if (selectedHistoryWorkspaceId) {
      fetchHistoryLogsForWorkspace(selectedHistoryWorkspaceId);
    } else {
      setHistoryLogs([]);
    }
  }, [selectedHistoryWorkspaceId]);

  useEffect(() => {
    checkOutlookStatus();
    fetchWorkspaces(true, true);
    const interval = setInterval(checkOutlookStatus, 5000);
    return () => clearInterval(interval);
  }, []);

  // Fetch templates list from backend
  const fetchTemplates = async () => {
    try {
      const response = await fetch('http://localhost:3001/api/templates');
      if (response.ok) {
        const data = await response.json();
        setTemplatesList(data);
      }
    } catch (error) {
      console.error('Failed to contact backend templates endpoint:', error);
      addLog('error', 'Failed to retrieve saved templates list.');
    }
  };

  useEffect(() => {
    fetchTemplates();
  }, []);

  const handleLoadTemplate = async (id: string) => {
    setSelectedTemplateId(id);
    if (!id) return;

    try {
      const response = await fetch(`http://localhost:3001/api/templates/${id}`);
      if (response.ok) {
        const template = await response.json();
        setEmailSubject(template.subject);
        setEmailBody(template.body);
        setEditorMode(template.editorMode);
        setImagesMap(template.imagesMap || {});
        addLog('success', `Loaded template: ${template.name}`);
      } else {
        addLog('error', 'Failed to load selected template details.');
      }
    } catch (error) {
      console.error('Error loading template:', error);
      addLog('error', 'Error loading template details.');
    }
  };

  const handleSaveAsNewTemplate = async () => {
    const name = prompt('Enter a name for the new template:');
    if (!name) return;

    const newTemplate = {
      name,
      subject: emailSubject,
      body: emailBody,
      editorMode,
      imagesMap
    };

    try {
      const response = await fetch('http://localhost:3001/api/templates', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(newTemplate)
      });

      if (response.ok) {
        const data = await response.json();
        addLog('success', `Template "${name}" saved successfully.`);
        await fetchTemplates();
        setSelectedTemplateId(data.template.id);
      } else {
        const errorData = await response.json();
        addLog('error', `Failed to save template: ${errorData.error}`);
      }
    } catch (error) {
      addLog('error', 'Failed to connect to backend to save template.');
    }
  };

  const handleOverwriteTemplate = async () => {
    if (!selectedTemplateId) return;
    const currentTemplate = templatesList.find(t => t.id === selectedTemplateId);
    if (!currentTemplate) return;

    if (!confirm(`Are you sure you want to overwrite template "${currentTemplate.name}" with your current edits?`)) {
      return;
    }

    const updatedTemplate = {
      id: selectedTemplateId,
      name: currentTemplate.name,
      subject: emailSubject,
      body: emailBody,
      editorMode,
      imagesMap
    };

    try {
      const response = await fetch('http://localhost:3001/api/templates', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(updatedTemplate)
      });

      if (response.ok) {
        addLog('success', `Template "${currentTemplate.name}" updated successfully.`);
        await fetchTemplates();
      } else {
        const errorData = await response.json();
        addLog('error', `Failed to update template: ${errorData.error}`);
      }
    } catch (error) {
      addLog('error', 'Failed to connect to backend to update template.');
    }
  };

  const handleDeleteTemplate = async () => {
    if (!selectedTemplateId) return;
    const currentTemplate = templatesList.find(t => t.id === selectedTemplateId);
    if (!currentTemplate) return;

    if (!confirm(`Are you sure you want to delete template "${currentTemplate.name}"?`)) {
      return;
    }

    try {
      const response = await fetch(`http://localhost:3001/api/templates/${selectedTemplateId}`, {
        method: 'DELETE'
      });

      if (response.ok) {
        addLog('success', `Template "${currentTemplate.name}" deleted.`);
        setSelectedTemplateId('');
        await fetchTemplates();
      } else {
        const errorData = await response.json();
        addLog('error', `Failed to delete template: ${errorData.error}`);
      }
    } catch (error) {
      addLog('error', 'Failed to connect to backend to delete template.');
    }
  };

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
      setImagesMap(prev => ({ ...prev, [file.name]: dataUrl }));
      const imgTag = currentMode === 'markdown'
        ? `![${file.name}]({{image:${file.name}}})`
        : `<img src="{{image:${file.name}}}" alt="${file.name}" style="max-width: 100%; height: auto; display: block; margin: 1rem 0;" />`;
      insertTextAtCursor(imgTag);
      addLog('success', `Uploaded and inserted placeholder for image: ${file.name}`);
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
          const headers = Object.keys(results.data[0] as object);
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

  const resolveImagePlaceholders = (text: string, map: Record<string, string>): string => {
    let result = text;
    Object.entries(map).forEach(([filename, base64Data]) => {
      const escapedName = filename.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
      const regex = new RegExp(`{{\\s*image:${escapedName}\\s*}}`, 'gi');
      result = result.replace(regex, base64Data);
    });
    return result;
  };

  // Preview fields
  const currentPreviewRow = csvRows[previewIndex];
  const renderedSubject = currentPreviewRow ? interpolateTemplate(emailSubject, currentPreviewRow) : '';
  const renderedBody = currentPreviewRow 
    ? (editorMode === 'markdown' 
        ? (marked.parse(resolveImagePlaceholders(interpolateTemplate(emailBody, currentPreviewRow), imagesMap)) as string)
        : resolveImagePlaceholders(interpolateTemplate(emailBody, currentPreviewRow), imagesMap))
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
      const resolvedBody = resolveImagePlaceholders(rawBody, imagesMap);
      const body = editorMode === 'markdown' ? (marked.parse(resolvedBody) as string) : resolvedBody;

      addLog('info', `Sending email to ${recipient}...`);

      try {
        const response = await fetch('http://localhost:3001/api/send-email', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({ recipient, subject, body, workspaceId: currentWorkspaceId })
        });

        const data = await response.json();

        if (response.ok && data.success) {
          setSuccessCount(prev => prev + 1);
          addLog('success', `[${idx + 1}/${csvRowsRef.current.length}] Successfully sent email to ${recipient}`);
        } else {
          setErrorCount(prev => prev + 1);
          addLog('error', `[${idx + 1}/${csvRowsRef.current.length}] Failed sending to ${recipient}: ${data.error || 'Unknown error'}`);
        }
        fetchHistoryLogsForWorkspace(currentWorkspaceId);
        fetchWorkspaces(false, false);
      } catch (error: any) {
        setErrorCount(prev => prev + 1);
        addLog('error', `[${idx + 1}/${csvRowsRef.current.length}] Connection error sending to ${recipient}: ${error.message}`);
        fetchHistoryLogsForWorkspace(currentWorkspaceId);
        fetchWorkspaces(false, false);
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
    setImagesMap({});
    setLogs([]);
    addLog('info', 'Workspace cleared.');
  };

  const handleFullReset = async () => {
    if (!confirm('Are you sure you want to reset the mailer dashboard? This will clear your CSV contacts, reset the email template to default, clear logs, and reset all sending configurations.')) {
      return;
    }
    
    // Stop any active sending
    stopSending();
    
    // Reset CSV data
    setCsvFile(null);
    setCsvHeaders([]);
    setCsvRows([]);
    
    // Reset Mapping Config
    setEmailColumn('');
    setNameColumn('');
    
    // Reset Template to Initial Default
    setEmailSubject('Greetings {{First Name}}!');
    setEmailBody(DEFAULT_HTML_BODY);
    setEditorMode('html');
    setImagesMap({});
    setSelectedTemplateId('');
    
    // Reset Delay
    setSendingDelay(3);
    
    // Reset Log console
    setLogs([]);
    
    // Create new workspace session (keeps old files intact)
    await createNewWorkspace();
    
    addLog('info', 'Mailer dashboard has been reset to initial default settings and a new workspace session created.');
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

  const exportRecipients = (type: 'valid' | 'invalid') => {
    const targetStatus = type === 'valid' ? 'success' : 'failed';
    const filtered = historyLogs.filter(log => log.status === targetStatus);
    
    if (filtered.length === 0) {
      alert(`No ${type} recipients found in history to export.`);
      return;
    }

    const headers = ['Timestamp', 'Recipient Email', 'Subject', 'Source', 'Status', 'Error/Reason'];
    const rows = filtered.map(log => [
      log.timestamp,
      log.recipient,
      log.subject,
      log.source,
      log.status,
      log.error || ''
    ]);

    const csvContent = [
      headers.join(','),
      ...rows.map(row => row.map(val => `"${val.replace(/"/g, '""')}"`).join(','))
    ].join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `${type}_recipients_${Date.now()}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    addLog('info', `Exported ${filtered.length} ${type} recipients.`);
  };

  const handleRetrySelected = async () => {
    if (selectedLogIds.length === 0) return;
    addLog('info', `Retrying ${selectedLogIds.length} failed delivery attempts...`);
    try {
      const response = await fetch(`http://localhost:3001/api/workspaces/${selectedHistoryWorkspaceId}/retry`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ids: selectedLogIds })
      });
      if (response.ok) {
        const data = await response.json();
        setHistoryLogs(data.logs);
        setSelectedLogIds([]);
        const successes = data.results.filter((r: any) => r.success).length;
        const failures = data.results.length - successes;
        addLog('success', `Retry completed: ${successes} succeeded, ${failures} failed.`);
        fetchWorkspaces(false, false);
      } else {
        addLog('error', 'Failed to retry selected deliveries.');
      }
    } catch (error: any) {
      addLog('error', `Error retrying deliveries: ${error.message}`);
    }
  };

  const handleSyncOutlook = async () => {
    setIsSyncingOutlook(true);
    addLog('info', 'Syncing deliveries with Microsoft Outlook...');
    try {
      const response = await fetch(`http://localhost:3001/api/workspaces/${selectedHistoryWorkspaceId}/sync-outlook`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ limitDays: 30 })
      });
      if (response.ok) {
        const data = await response.json();
        setHistoryLogs(data.logs);
        addLog('success', `Outlook sync completed. Synced ${data.count} new failure events.`);
        fetchWorkspaces(false, false);
      } else {
        addLog('error', 'Failed to sync with Outlook.');
      }
    } catch (error: any) {
      addLog('error', `Error syncing with Outlook: ${error.message}`);
    } finally {
      setIsSyncingOutlook(false);
    }
  };

  const handleClearHistory = async (ids?: string[]) => {
    const confirmMsg = ids 
      ? `Are you sure you want to delete the selected log entry?` 
      : 'Are you sure you want to delete this entire workspace log session? This will remove the file from history.';
    if (!confirm(confirmMsg)) return;

    try {
      const response = await fetch(`http://localhost:3001/api/workspaces/${selectedHistoryWorkspaceId}`, {
        method: 'DELETE',
        headers: { 'Content-Type': 'application/json' },
        body: ids ? JSON.stringify({ ids }) : undefined
      });
      if (response.ok) {
        const data = await response.json();
        if (ids) {
          setHistoryLogs(data.logs);
          setSelectedLogIds(prev => prev.filter(id => !ids.includes(id)));
          if (selectedLogForDetail && ids.includes(selectedLogForDetail.id)) {
            setSelectedLogForDetail(null);
          }
          addLog('success', 'Selected log entries deleted.');
          fetchWorkspaces(false, false);
        } else {
          setSelectedLogIds([]);
          setSelectedLogForDetail(null);
          addLog('success', 'Workspace log session deleted.');
          await fetchWorkspaces(true, true);
        }
      } else {
        addLog('error', 'Failed to clear history log.');
      }
    } catch (error: any) {
      addLog('error', `Error clearing history: ${error.message}`);
    }
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

      {/* Tab Navigation */}
      <div className="tab-navigation" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div style={{ display: 'flex', gap: '0.75rem' }}>
          <button 
            className={`tab-btn ${activeTab === 'dashboard' ? 'active' : ''}`}
            onClick={() => setActiveTab('dashboard')}
          >
            <Mail size={16} />
            <span>Mailer Dashboard</span>
          </button>
          <button 
            className={`tab-btn ${activeTab === 'history' ? 'active' : ''}`}
            onClick={() => setActiveTab('history')}
          >
            <History size={16} />
            <span>Delivery History</span>
            {historyLogs.length > 0 && (
              <span className="tab-badge">{historyLogs.length}</span>
            )}
          </button>
        </div>
        {activeTab === 'dashboard' && (
          <button
            onClick={handleFullReset}
            className="btn btn-danger btn-outline-danger"
            style={{ padding: '0.5rem 1rem', fontSize: '0.85rem', display: 'flex', alignItems: 'center', gap: '0.35rem', borderRadius: '0.6rem' }}
            title="Reset all settings and workspace to default"
          >
            <RotateCcw size={14} />
            <span>Reset Workspace</span>
          </button>
        )}
      </div>

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

      {activeTab === 'dashboard' ? (
        <div className="dashboard-grid animate-fadeIn">
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

              {/* Template Selector & Manager */}
              <div className="form-group" style={{ 
                background: 'rgba(255, 255, 255, 0.02)', 
                padding: '1.25rem', 
                borderRadius: '0.85rem', 
                border: '1px solid var(--border-light)',
                display: 'flex',
                flexDirection: 'column',
                gap: '0.75rem',
                boxShadow: 'inset 0 1px 3px rgba(0, 0, 0, 0.2)'
              }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <label style={{ fontWeight: '600', color: 'var(--text-primary)' }}>Template Presets</label>
                  {selectedTemplateId && (
                    <span style={{ fontSize: '0.75rem', color: 'var(--secondary)', display: 'inline-flex', alignItems: 'center', gap: '0.25rem' }}>
                      <span style={{ display: 'inline-block', width: '6px', height: '6px', borderRadius: '50%', backgroundColor: 'var(--secondary)' }}></span>
                      Active template loaded
                    </span>
                  )}
                </div>
                
                <div style={{ display: 'flex', gap: '0.5rem', alignItems: 'center', flexWrap: 'wrap' }}>
                  <select
                    value={selectedTemplateId}
                    onChange={(e) => handleLoadTemplate(e.target.value)}
                    style={{ flex: 1, minWidth: '180px' }}
                    disabled={sendingStatus === 'sending'}
                  >
                    <option value="">-- Load/Select Template --</option>
                    {templatesList.map(t => (
                      <option key={t.id} value={t.id}>{t.name} ({t.editorMode.toUpperCase()})</option>
                    ))}
                  </select>

                  <div style={{ display: 'flex', gap: '0.4rem', flexShrink: 0 }}>
                    {selectedTemplateId && (
                      <button 
                        type="button"
                        onClick={handleOverwriteTemplate} 
                        className="btn btn-secondary" 
                        title="Save current subject & body to this template"
                        disabled={sendingStatus === 'sending'}
                        style={{ padding: '0.5rem 0.85rem', fontSize: '0.8rem', display: 'flex', alignItems: 'center', gap: '0.35rem', borderRadius: '0.6rem' }}
                      >
                        <Save size={14} />
                        <span>Overwrite</span>
                      </button>
                    )}
                    
                    <button 
                      type="button"
                      onClick={handleSaveAsNewTemplate} 
                      className="btn btn-primary"
                      title="Save current layout as a new template"
                      disabled={sendingStatus === 'sending'}
                      style={{ padding: '0.5rem 0.85rem', fontSize: '0.8rem', display: 'flex', alignItems: 'center', gap: '0.35rem', borderRadius: '0.6rem' }}
                    >
                      <Plus size={14} />
                      <span>Save As New</span>
                    </button>

                    {selectedTemplateId && (
                      <button 
                        type="button"
                        onClick={handleDeleteTemplate} 
                        className="btn btn-danger"
                        title="Delete this template"
                        disabled={sendingStatus === 'sending'}
                        style={{ padding: '0.5rem', display: 'flex', alignItems: 'center', justifyContent: 'center', borderRadius: '0.6rem', boxShadow: 'none' }}
                      >
                        <Trash2 size={14} />
                      </button>
                    )}
                  </div>
                </div>
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
                        onClick={() => {
                          setEditorMode('html');
                          if (emailBody.trim() === DEFAULT_MARKDOWN_BODY.trim()) {
                            setEmailBody(DEFAULT_HTML_BODY);
                          } else {
                            try {
                              const converted = (marked.parse(emailBody) as string).trim();
                              setEmailBody(converted);
                            } catch (err) {
                              console.error('Failed to convert Markdown to HTML:', err);
                            }
                          }
                        }}
                        className={`btn ${editorMode === 'html' ? 'btn-primary' : 'btn-secondary'}`}
                        style={{ padding: '0.25rem 0.6rem', fontSize: '0.75rem', borderRadius: '0.35rem', boxShadow: 'none' }}
                        disabled={sendingStatus === 'sending'}
                      >
                        HTML
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          setEditorMode('markdown');
                          if (emailBody.trim() === DEFAULT_HTML_BODY.trim()) {
                            setEmailBody(DEFAULT_MARKDOWN_BODY);
                          } else {
                            try {
                              const converted = htmlToMarkdown(emailBody);
                              setEmailBody(converted);
                            } catch (err) {
                              console.error('Failed to convert HTML to Markdown:', err);
                            }
                          }
                        }}
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
                        <span className="preview-field-label">From:</span>
                        <span>{outlookEmail || 'outlook-user@example.com'}</span>
                      </div>
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
      ) : (
        /* Delivery History Tab View */
        <div className="history-container card animate-fadeIn">
          {/* Workspace Run Selector */}
          <div style={{
            display: 'flex',
            alignItems: 'center',
            gap: '0.75rem',
            background: 'rgba(255, 255, 255, 0.02)',
            border: '1px solid var(--border-light)',
            padding: '0.75rem 1.25rem',
            borderRadius: '0.75rem',
            flexWrap: 'wrap',
            marginBottom: '1.25rem'
          }}>
            <span style={{ fontSize: '0.9rem', fontWeight: '600', color: 'var(--text-secondary)' }}>Select Workspace Run:</span>
            <select
              value={selectedHistoryWorkspaceId}
              onChange={(e) => setSelectedHistoryWorkspaceId(e.target.value)}
              style={{
                flex: 1,
                minWidth: '250px',
                padding: '0.4rem 0.75rem',
                borderRadius: '0.5rem',
                background: 'var(--card-bg)',
                color: 'var(--text-primary)',
                border: '1px solid var(--border-light)',
                cursor: 'pointer'
              }}
            >
              {workspacesList.length === 0 ? (
                <option value="">-- No Workspaces Found --</option>
              ) : (
                workspacesList.map(ws => (
                  <option key={ws.id} value={ws.id}>
                    {formatWorkspaceName(ws.id)} (Total: {ws.total} | Success: {ws.successCount} | Failed: {ws.errorCount})
                  </option>
                ))
              )}
            </select>
          </div>

          <div className="card-title" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '1rem' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
              <History size={20} />
              <span>Email Delivery Log History</span>
            </div>
            
            <div className="history-global-actions" style={{ display: 'flex', gap: '0.5rem', flexWrap: 'wrap' }}>
              <button 
                onClick={handleSyncOutlook}
                disabled={isSyncingOutlook || !isOutlookRunning}
                className="btn btn-primary"
                style={{ padding: '0.5rem 0.9rem', fontSize: '0.85rem' }}
              >
                <RefreshCw size={14} className={isSyncingOutlook ? 'animate-spin' : ''} />
                <span>{isSyncingOutlook ? 'Syncing...' : 'Sync Outlook'}</span>
              </button>
              
              <button 
                onClick={() => exportRecipients('valid')}
                className="btn btn-secondary btn-outline-success"
                style={{ padding: '0.5rem 0.9rem', fontSize: '0.85rem' }}
                title="Export Succeeded Recipients"
              >
                <Download size={14} />
                <span>Export Valid</span>
              </button>

              <button 
                onClick={() => exportRecipients('invalid')}
                className="btn btn-secondary btn-outline-danger"
                style={{ padding: '0.5rem 0.9rem', fontSize: '0.85rem' }}
                title="Export Failed Recipients"
              >
                <Download size={14} />
                <span>Export Invalid</span>
              </button>

              <button 
                onClick={() => handleClearHistory()}
                disabled={historyLogs.length === 0}
                className="btn btn-danger"
                style={{ padding: '0.5rem 0.9rem', fontSize: '0.85rem' }}
              >
                <Trash2 size={14} />
                <span>Clear All Log</span>
              </button>
            </div>
          </div>

          {/* Delivery Stats Bar */}
          <div className="history-stats-bar" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '1rem', margin: '0.5rem 0 1.5rem 0' }}>
            <div className="stat-box all" style={{ background: 'rgba(255,255,255,0.02)', border: '1px solid var(--border-light)', padding: '1rem', borderRadius: '0.75rem', textAlign: 'center' }}>
              <div className="stat-title" style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>Total Attempts</div>
              <div className="stat-value" style={{ fontSize: '1.75rem', fontWeight: '700', color: 'var(--text-primary)', marginTop: '0.25rem' }}>{historyLogs.length}</div>
            </div>
            <div className="stat-box success" style={{ background: 'rgba(16, 185, 129, 0.03)', border: '1px solid rgba(16, 185, 129, 0.15)', padding: '1rem', borderRadius: '0.75rem', textAlign: 'center' }}>
              <div className="stat-title" style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>Valid (Succeeded)</div>
              <div className="stat-value" style={{ fontSize: '1.75rem', fontWeight: '700', color: 'var(--success)', marginTop: '0.25rem' }}>{historyLogs.filter(l => l.status === 'success').length}</div>
            </div>
            <div className="stat-box failure" style={{ background: 'rgba(244, 63, 94, 0.03)', border: '1px solid rgba(244, 63, 94, 0.15)', padding: '1rem', borderRadius: '0.75rem', textAlign: 'center' }}>
              <div className="stat-title" style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>Invalid (Failed)</div>
              <div className="stat-value" style={{ fontSize: '1.75rem', fontWeight: '700', color: 'var(--error)', marginTop: '0.25rem' }}>{historyLogs.filter(l => l.status === 'failed').length}</div>
            </div>
          </div>

          {/* Filters and Search */}
          <div className="history-filter-bar" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '1rem', flexWrap: 'wrap', marginBottom: '1rem' }}>
            <div style={{ display: 'flex', gap: '0.25rem', background: 'rgba(255,255,255,0.04)', padding: '0.2rem', borderRadius: '0.5rem', border: '1px solid var(--border-light)' }}>
              <button 
                onClick={() => setHistoryFilter('all')} 
                className={`btn ${historyFilter === 'all' ? 'btn-primary' : 'btn-secondary'}`}
                style={{ padding: '0.35rem 0.8rem', fontSize: '0.8rem', borderRadius: '0.35rem', boxShadow: 'none' }}
              >
                All
              </button>
              <button 
                onClick={() => setHistoryFilter('success')} 
                className={`btn ${historyFilter === 'success' ? 'btn-primary' : 'btn-secondary'}`}
                style={{ padding: '0.35rem 0.8rem', fontSize: '0.8rem', borderRadius: '0.35rem', boxShadow: 'none' }}
              >
                Succeeded
              </button>
              <button 
                onClick={() => setHistoryFilter('failed')} 
                className={`btn ${historyFilter === 'failed' ? 'btn-primary' : 'btn-secondary'}`}
                style={{ padding: '0.35rem 0.8rem', fontSize: '0.8rem', borderRadius: '0.35rem', boxShadow: 'none' }}
              >
                Failed
              </button>
            </div>

            <div className="search-box-wrapper" style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', flex: 1, maxWidth: '350px', position: 'relative' }}>
              <Search size={16} className="search-icon" style={{ position: 'absolute', left: '0.8rem', color: 'var(--text-muted)' }} />
              <input 
                type="text" 
                placeholder="Search recipient or subject..." 
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                style={{ width: '100%', paddingLeft: '2.2rem', fontSize: '0.85rem' }}
              />
            </div>
          </div>

          {/* Selected Action Banner */}
          {selectedLogIds.length > 0 && (
            <div className="selected-action-banner" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: 'rgba(99, 102, 241, 0.1)', border: '1px solid var(--primary)', borderRadius: '0.75rem', padding: '0.75rem 1.25rem', marginBottom: '1rem', animation: 'fadeIn 0.2s ease' }}>
              <span style={{ fontSize: '0.9rem' }}><strong>{selectedLogIds.length}</strong> items selected</span>
              <div style={{ display: 'flex', gap: '0.5rem' }}>
                <button 
                  onClick={handleRetrySelected}
                  className="btn btn-primary"
                  style={{ padding: '0.4rem 0.85rem', fontSize: '0.8rem' }}
                >
                  <RefreshCw size={12} />
                  <span>Retry Selected Failures</span>
                </button>
                <button 
                  onClick={() => handleClearHistory(selectedLogIds)}
                  className="btn btn-danger"
                  style={{ padding: '0.4rem 0.85rem', fontSize: '0.8rem' }}
                >
                  <Trash2 size={12} />
                  <span>Delete Selected</span>
                </button>
              </div>
            </div>
          )}

          {/* History List Table */}
          <div className="history-table-wrapper" style={{ overflowX: 'auto', border: '1px solid var(--border-light)', borderRadius: '0.75rem', background: 'rgba(0, 0, 0, 0.1)' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '0.9rem' }}>
              <thead>
                <tr style={{ borderBottom: '1px solid var(--border-light)', background: 'rgba(255,255,255,0.02)' }}>
                  <th style={{ width: '40px', padding: '1rem' }}>
                    <input 
                      type="checkbox"
                      checked={
                        (() => {
                          const filtered = historyLogs.filter(log => {
                            if (historyFilter === 'success') return log.status === 'success';
                            if (historyFilter === 'failed') return log.status === 'failed';
                            return true;
                          }).filter(log => 
                            log.recipient.toLowerCase().includes(searchQuery.toLowerCase()) || 
                            log.subject.toLowerCase().includes(searchQuery.toLowerCase())
                          );
                          return filtered.length > 0 && filtered.every(log => selectedLogIds.includes(log.id));
                        })()
                      }
                      onChange={(e) => {
                        const filtered = historyLogs.filter(log => {
                          if (historyFilter === 'success') return log.status === 'success';
                          if (historyFilter === 'failed') return log.status === 'failed';
                          return true;
                        }).filter(log => 
                          log.recipient.toLowerCase().includes(searchQuery.toLowerCase()) || 
                          log.subject.toLowerCase().includes(searchQuery.toLowerCase())
                        );

                        if (e.target.checked) {
                          setSelectedLogIds(prev => {
                            const newIds = [...prev];
                            filtered.forEach(log => {
                              if (!newIds.includes(log.id)) newIds.push(log.id);
                            });
                            return newIds;
                          });
                        } else {
                          setSelectedLogIds(prev => prev.filter(id => !filtered.some(log => log.id === id)));
                        }
                      }}
                    />
                  </th>
                  <th style={{ padding: '1rem' }}>Timestamp</th>
                  <th style={{ padding: '1rem' }}>Recipient Email</th>
                  <th style={{ padding: '1rem' }}>Subject</th>
                  <th style={{ padding: '1rem' }}>Source</th>
                  <th style={{ padding: '1rem' }}>Status</th>
                  <th style={{ padding: '1rem' }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {(() => {
                  const filtered = historyLogs.filter(log => {
                    if (historyFilter === 'success') return log.status === 'success';
                    if (historyFilter === 'failed') return log.status === 'failed';
                    return true;
                  }).filter(log => 
                    log.recipient.toLowerCase().includes(searchQuery.toLowerCase()) || 
                    log.subject.toLowerCase().includes(searchQuery.toLowerCase())
                  );

                  if (filtered.length === 0) {
                    return (
                      <tr>
                        <td colSpan={7} style={{ textAlign: 'center', padding: '3rem 0', color: 'var(--text-muted)', fontSize: '0.9rem' }}>
                          No delivery attempts recorded matching criteria.
                        </td>
                      </tr>
                    );
                  }

                  return filtered.map((log) => (
                    <tr key={log.id} style={{ borderBottom: '1px solid var(--border-light)', transition: 'background-color 0.2s' }} className={`history-row ${log.status === 'failed' ? 'row-failed' : ''}`}>
                      <td style={{ padding: '1rem' }}>
                        <input 
                          type="checkbox"
                          checked={selectedLogIds.includes(log.id)}
                          onChange={(e) => {
                            if (e.target.checked) {
                              setSelectedLogIds(prev => [...prev, log.id]);
                            } else {
                              setSelectedLogIds(prev => prev.filter(id => id !== log.id));
                            }
                          }}
                        />
                      </td>
                      <td style={{ padding: '1rem', fontSize: '0.8rem', color: 'var(--text-secondary)' }}>
                        {new Date(log.timestamp).toLocaleString()}
                      </td>
                      <td style={{ padding: '1rem', fontWeight: '500' }}>
                        {log.recipient}
                      </td>
                      <td style={{ padding: '1rem', maxWidth: '300px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={log.subject}>
                        {log.subject}
                      </td>
                      <td style={{ padding: '1rem' }}>
                        <span className={`tag ${log.source === 'CLI' ? 'tag-purple' : 'tag-indigo'}`} style={{ padding: '0.2rem 0.5rem', borderRadius: '0.25rem', fontSize: '0.75rem', fontWeight: '600' }}>
                          {log.source}
                        </span>
                      </td>
                      <td style={{ padding: '1rem' }}>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '0.2rem' }}>
                          <span className={`badge ${log.status === 'success' ? 'badge-success' : 'badge-danger'}`} style={{ padding: '0.25rem 0.6rem', borderRadius: '9999px', fontSize: '0.75rem', fontWeight: '600', width: 'fit-content' }}>
                            {log.status === 'success' ? 'Succeeded' : 'Failed'}
                          </span>
                          {log.status === 'failed' && log.error && (
                            <span style={{ fontSize: '0.75rem', color: 'var(--error)', maxWidth: '200px', display: 'block', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={log.error}>
                              {log.error}
                            </span>
                          )}
                        </div>
                      </td>
                      <td style={{ padding: '0.75rem 1rem' }}>
                        <div style={{ display: 'flex', gap: '0.4rem' }}>
                          <button 
                            className="btn btn-secondary" 
                            style={{ padding: '0.35rem 0.6rem', fontSize: '0.75rem', display: 'inline-flex', alignItems: 'center', gap: '0.25rem', borderRadius: '0.4rem' }}
                            onClick={() => setSelectedLogForDetail(log)}
                          >
                            <ExternalLink size={12} />
                            <span>View</span>
                          </button>
                          {log.status === 'failed' && (
                            <button 
                              className="btn btn-primary" 
                              style={{ padding: '0.35rem 0.6rem', fontSize: '0.75rem', display: 'inline-flex', alignItems: 'center', gap: '0.25rem', borderRadius: '0.4rem' }}
                              onClick={async () => {
                                const response = await fetch(`http://localhost:3001/api/workspaces/${selectedHistoryWorkspaceId}/retry`, {
                                  method: 'POST',
                                  headers: { 'Content-Type': 'application/json' },
                                  body: JSON.stringify({ ids: [log.id] })
                                });
                                if (response.ok) {
                                  const data = await response.json();
                                  setHistoryLogs(data.logs);
                                  addLog('success', `Retried sending to ${log.recipient}.`);
                                  fetchWorkspaces(false, false);
                                }
                              }}
                            >
                              <RefreshCw size={12} />
                              <span>Retry</span>
                            </button>
                          )}
                          <button 
                            className="btn btn-danger" 
                            style={{ padding: '0.35rem', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', borderRadius: '0.4rem' }}
                            onClick={() => handleClearHistory([log.id])}
                          >
                            <Trash2 size={12} />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ));
                })()}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Slide-out Detail Drawer */}
      {selectedLogForDetail && (
        <div className="drawer-overlay" onClick={() => setSelectedLogForDetail(null)}>
          <div className="drawer-content animate-slideLeft" onClick={(e) => e.stopPropagation()}>
            <div className="drawer-header">
              <h2>Delivery Details</h2>
              <button className="drawer-close-btn" onClick={() => setSelectedLogForDetail(null)}>×</button>
            </div>
            
            <div className="drawer-body">
              <div className="detail-row">
                <span className="detail-label">Status</span>
                <span className={`badge ${selectedLogForDetail.status === 'success' ? 'badge-success' : 'badge-danger'}`}>
                  {selectedLogForDetail.status === 'success' ? 'Succeeded (Valid)' : 'Failed (Invalid)'}
                </span>
              </div>

              {selectedLogForDetail.error && (
                <div style={{
                  background: 'rgba(244, 63, 94, 0.08)',
                  border: '1px solid var(--error)',
                  borderRadius: '0.5rem',
                  padding: '0.75rem',
                  marginBottom: '1.25rem',
                  fontSize: '0.85rem',
                  color: 'var(--text-primary)',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '0.25rem'
                }}>
                  <strong style={{ color: 'var(--error)' }}>Error Message:</strong>
                  <span>{selectedLogForDetail.error}</span>
                </div>
              )}

              <div className="detail-row">
                <span className="detail-label">Recipient</span>
                <span className="detail-value font-mono">{selectedLogForDetail.recipient}</span>
              </div>

              <div className="detail-row">
                <span className="detail-label">Subject</span>
                <span className="detail-value" style={{ fontWeight: '600' }}>{selectedLogForDetail.subject}</span>
              </div>

              <div className="detail-row">
                <span className="detail-label">Timestamp</span>
                <span className="detail-value">{new Date(selectedLogForDetail.timestamp).toLocaleString()}</span>
              </div>

              <div className="detail-row">
                <span className="detail-label">Source</span>
                <span className={`tag ${selectedLogForDetail.source === 'CLI' ? 'tag-purple' : 'tag-indigo'}`}>
                  {selectedLogForDetail.source}
                </span>
              </div>

              <div className="detail-body-preview">
                <span className="detail-label" style={{ marginBottom: '0.5rem', display: 'block' }}>Email Content Body</span>
                <div 
                  className="preview-body-content"
                  style={{ background: 'rgba(0,0,0,0.3)', border: '1px solid var(--border-light)', borderRadius: '0.5rem', padding: '1rem', maxHeight: '350px', overflowY: 'auto', fontSize: '0.875rem', lineHeight: '1.5' }}
                  dangerouslySetInnerHTML={{ __html: selectedLogForDetail.body }}
                />
              </div>
            </div>

            <div className="drawer-footer">
              {selectedLogForDetail.status === 'failed' && (
                <button 
                  className="btn btn-primary"
                  style={{ width: '100%', marginBottom: '0.5rem' }}
                  onClick={async () => {
                    const response = await fetch(`http://localhost:3001/api/workspaces/${selectedHistoryWorkspaceId}/retry`, {
                      method: 'POST',
                      headers: { 'Content-Type': 'application/json' },
                      body: JSON.stringify({ ids: [selectedLogForDetail.id] })
                    });
                    if (response.ok) {
                      const data = await response.json();
                      setHistoryLogs(data.logs);
                      const updatedLog = data.logs.find((l: any) => l.id === selectedLogForDetail.id);
                      setSelectedLogForDetail(updatedLog || null);
                      addLog('success', `Retried sending to ${selectedLogForDetail.recipient}.`);
                      fetchWorkspaces(false, false);
                    }
                  }}
                >
                  <RefreshCw size={14} />
                  <span>Retry Sending Now</span>
                </button>
              )}
              <button 
                className="btn btn-secondary"
                style={{ width: '100%' }}
                onClick={() => setSelectedLogForDetail(null)}
              >
                Close Drawer
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
