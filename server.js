import express from 'express';
import cors from 'cors';
import { spawn } from 'child_process';
import fs from 'fs';
import path from 'path';

const app = express();
const port = 3001;

app.use(cors());
app.use(express.json({ limit: '10mb' })); // Support larger payloads for batch mail contents

// Helper to execute AppleScript via stdin (prevents shell escaping issues)
function runAppleScript(scriptText) {
  return new Promise((resolve, reject) => {
    const child = spawn('osascript', []);
    let stdout = '';
    let stderr = '';

    child.stdout.on('data', (data) => {
      stdout += data.toString();
    });

    child.stderr.on('data', (data) => {
      stderr += data.toString();
    });

    child.on('close', (code) => {
      if (code === 0) {
        resolve(stdout.trim());
      } else {
        reject(new Error(stderr.trim() || `Exit code ${code}`));
      }
    });

    child.stdin.write(scriptText);
    child.stdin.end();
  });
}

// Escape strings specifically for AppleScript string literals
function escapeForAppleScript(str) {
  if (!str) return '';
  return str
    .replace(/\\/g, '\\\\')  // Escape backslashes
    .replace(/"/g, '\\"');   // Escape double quotes
}

const logsDir = path.join(process.cwd(), 'logs');

// Helper to get log file path for a workspace ID
function getLogFilePath(workspaceId) {
  const safeId = String(workspaceId).replace(/[^a-zA-Z0-9_-]/g, '');
  return path.join(logsDir, `workspace_${safeId}.json`);
}

// Ensure logs directory exists
function ensureLogsDir() {
  if (!fs.existsSync(logsDir)) {
    fs.mkdirSync(logsDir, { recursive: true });
  }
}

// Read log file for a specific workspace
function getDeliveryLogs(workspaceId) {
  ensureLogsDir();
  const filePath = getLogFilePath(workspaceId);
  if (!fs.existsSync(filePath)) {
    return [];
  }
  try {
    const data = fs.readFileSync(filePath, 'utf8');
    return JSON.parse(data);
  } catch (error) {
    console.error(`Error reading delivery log for workspace ${workspaceId}:`, error);
    return [];
  }
}

// Save log file for a specific workspace
function saveDeliveryLogs(workspaceId, logs) {
  ensureLogsDir();
  const filePath = getLogFilePath(workspaceId);
  try {
    fs.writeFileSync(filePath, JSON.stringify(logs, null, 2));
  } catch (error) {
    console.error(`Error saving delivery log for workspace ${workspaceId}:`, error);
  }
}

// Log a single delivery attempt for a workspace
function logDeliveryAttempt(workspaceId, recipient, subject, body, status, errorMsg = null, source = 'Web App') {
  const logs = getDeliveryLogs(workspaceId);
  const newEntry = {
    id: `del_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
    timestamp: new Date().toISOString(),
    recipient,
    subject,
    body,
    status, // 'success' or 'failed'
    error: errorMsg,
    source
  };
  logs.push(newEntry);
  saveDeliveryLogs(workspaceId, logs);
  return newEntry;
}

// Helper to clean subjects for matching bounce messages
function cleanSubject(subject) {
  if (!subject) return '';
  return subject
    .replace(/^(Undeliverable|Undelivered|Delivery Failure|Failed|Failure Notice|Returned mail|Diagnostic|FW|RE):\s*/i, '')
    .trim();
}

// Helper to parse AppleScript dates robustly
function parseAppleScriptDate(dateStr) {
  if (!dateStr) return new Date();
  const cleaned = dateStr.replace(/\s+at\s+/i, ' ').trim();
  const d = new Date(cleaned);
  return isNaN(d.getTime()) ? new Date() : d;
}

// Generate timestamp-based workspace ID: YYYYMMDD_HHmmss
function generateTimestampId() {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  const yyyy = now.getFullYear();
  const mm = pad(now.getMonth() + 1);
  const dd = pad(now.getDate());
  const hh = pad(now.getHours());
  const min = pad(now.getMinutes());
  const ss = pad(now.getSeconds());
  return `${yyyy}${mm}${dd}_${hh}${min}${ss}`;
}

// Endpoint to check Outlook's status
app.get('/api/status', async (req, res) => {
  try {
    // Check if the Microsoft Outlook process is running
    const script = `tell application "System Events" to set isRunning to (name of processes) contains "Microsoft Outlook"
return isRunning`;
    
    const result = await runAppleScript(script);
    
    if (result === 'true') {
      const emailScript = `tell application "Microsoft Outlook"
    try
        set inboxFolders to every folder whose name is "Inbox"
        repeat with f in inboxFolders
            try
                set acc to account of f
                if acc is not missing value then
                    return email address of acc
                end if
            end try
        end repeat
        return ""
    on error
        return ""
    end try
end tell`;
      const email = await runAppleScript(emailScript);
      res.json({ running: true, email: email.trim() });
    } else {
      res.json({ running: false, email: '' });
    }
  } catch (error) {
    console.error('Error checking Outlook status:', error);
    res.json({ running: false, email: '', error: error.message });
  }
});

// Endpoint to send a single email via AppleScript + Outlook
app.post('/api/send-email', async (req, res) => {
  const { recipient, subject, body, workspaceId } = req.body;

  if (!recipient || !subject || !body) {
    return res.status(400).json({ error: 'Missing recipient, subject, or body' });
  }

  const activeWorkspaceId = workspaceId || generateTimestampId();

  const escapedRecipient = escapeForAppleScript(recipient);
  const escapedSubject = escapeForAppleScript(subject);
  const escapedBody = escapeForAppleScript(body);

  // Construct the AppleScript to send email via Outlook for Mac
  const appleScript = `
tell application "Microsoft Outlook"
    try
        -- Create a new outgoing message
        set newMessage to make new outgoing message with properties {subject:"${escapedSubject}", content:"${escapedBody}"}
        
        -- Add recipient
        make new recipient at newMessage with properties {email address:{address:"${escapedRecipient}"}}
        
        -- Send the message
        send newMessage
        
        return "SUCCESS"
    on error errMsg number errNum
        return "ERROR: " & errMsg & " (Code " & errNum & ")"
    end try
end tell
  `;

  try {
    const result = await runAppleScript(appleScript);
    if (result === 'SUCCESS') {
      logDeliveryAttempt(activeWorkspaceId, recipient, subject, body, 'success');
      res.json({ success: true, workspaceId: activeWorkspaceId });
    } else {
      logDeliveryAttempt(activeWorkspaceId, recipient, subject, body, 'failed', result);
      res.status(500).json({ success: false, error: result, workspaceId: activeWorkspaceId });
    }
  } catch (error) {
    console.error(`Failed to send email to ${recipient}:`, error);
    logDeliveryAttempt(activeWorkspaceId, recipient, subject, body, 'failed', error.message);
    res.status(500).json({ success: false, error: error.message, workspaceId: activeWorkspaceId });
  }
});

// Endpoint to get list of all workspaces and stats
app.get('/api/workspaces', (req, res) => {
  ensureLogsDir();
  try {
    const files = fs.readdirSync(logsDir);
    const workspaces = [];

    for (const file of files) {
      if (file.startsWith('workspace_') && file.endsWith('.json')) {
        const filePath = path.join(logsDir, file);
        const workspaceId = file.substring('workspace_'.length, file.length - '.json'.length);
        
        try {
          const stats = fs.statSync(filePath);
          const data = JSON.parse(fs.readFileSync(filePath, 'utf8'));
          
          const total = data.length;
          const successCount = data.filter(item => item.status === 'success').length;
          const errorCount = data.filter(item => item.status === 'failed').length;
          
          workspaces.push({
            id: workspaceId,
            createdAt: stats.birthtime || stats.mtime,
            total,
            successCount,
            errorCount
          });
        } catch (err) {
          console.error(`Error parsing workspace log ${file}:`, err);
        }
      }
    }

    // Sort by createdAt descending
    workspaces.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
    res.json(workspaces);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Endpoint to create a new workspace ID
app.post('/api/workspaces/new', (req, res) => {
  try {
    const workspaceId = generateTimestampId();
    saveDeliveryLogs(workspaceId, []);
    res.json({ success: true, workspaceId });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Endpoint to get delivery log for a specific workspace
app.get('/api/workspaces/:workspaceId/delivery-log', (req, res) => {
  try {
    res.json(getDeliveryLogs(req.params.workspaceId));
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Endpoint to retry failed deliveries in a specific workspace
app.post('/api/workspaces/:workspaceId/retry', async (req, res) => {
  const { workspaceId } = req.params;
  const { ids } = req.body;
  if (!ids || !Array.isArray(ids)) {
    return res.status(400).json({ error: 'Missing log entry ids to retry' });
  }

  const logs = getDeliveryLogs(workspaceId);
  const results = [];

  for (const id of ids) {
    const logEntry = logs.find(l => l.id === id);
    if (!logEntry) {
      results.push({ id, success: false, error: 'Log entry not found' });
      continue;
    }

    const { recipient, subject, body } = logEntry;
    const escapedRecipient = escapeForAppleScript(recipient);
    const escapedSubject = escapeForAppleScript(subject);
    const escapedBody = escapeForAppleScript(body);

    const appleScript = `
tell application "Microsoft Outlook"
    try
        set newMessage to make new outgoing message with properties {subject:"${escapedSubject}", content:"${escapedBody}"}
        make new recipient at newMessage with properties {email address:{address:"${escapedRecipient}"}}
        send newMessage
        return "SUCCESS"
    on error errMsg number errNum
        return "ERROR: " & errMsg & " (Code " & errNum & ")"
    end try
end tell
    `;

    try {
      const result = await runAppleScript(appleScript);
      if (result === 'SUCCESS') {
        logEntry.status = 'success';
        logEntry.error = null;
        logEntry.timestamp = new Date().toISOString();
        results.push({ id, success: true });
      } else {
        logEntry.status = 'failed';
        logEntry.error = result;
        logEntry.timestamp = new Date().toISOString();
        results.push({ id, success: false, error: result });
      }
    } catch (error) {
      logEntry.status = 'failed';
      logEntry.error = error.message;
      logEntry.timestamp = new Date().toISOString();
      results.push({ id, success: false, error: error.message });
    }
  }

  saveDeliveryLogs(workspaceId, logs);
  res.json({ success: true, results, logs: getDeliveryLogs(workspaceId) });
});

// Endpoint to delete items or delete the workspace file entirely
app.delete('/api/workspaces/:workspaceId', (req, res) => {
  const { workspaceId } = req.params;
  const { ids } = req.body || {};

  try {
    const filePath = getLogFilePath(workspaceId);
    
    if (ids && Array.isArray(ids)) {
      const logs = getDeliveryLogs(workspaceId);
      const filteredLogs = logs.filter(l => !ids.includes(l.id));
      saveDeliveryLogs(workspaceId, filteredLogs);
      res.json({ success: true, logs: filteredLogs });
    } else {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
      res.json({ success: true, logs: [] });
    }
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Endpoint to sync with Outlook for a specific workspace
app.post('/api/workspaces/:workspaceId/sync-outlook', async (req, res) => {
  const { workspaceId } = req.params;
  const { limitDays } = req.body || {};
  const daysLimit = parseInt(limitDays) || 30;

  // AppleScript to scan Inbox for bounces and Outbox for stuck mails
  const appleScript = `
tell application "Microsoft Outlook"
    set limitDate to (current date) - (${daysLimit} * days)
    set outputLines to {}
    
    -- Scan Inboxes for bounces
    set inboxFolders to every folder whose name is "Inbox"
    repeat with i from 1 to count of inboxFolders
        set f to item i of inboxFolders
        try
            set msgs to (every message of f whose time received > limitDate and (subject contains "Undeliverable" or subject contains "Failure" or subject contains "Delivery Status" or subject contains "Returned mail"))
            repeat with j from 1 to count of msgs
                set msg to item j of msgs
                set msgId to id of msg
                set msgSubject to subject of msg
                set msgSender to sender of msg
                set senderName to name of msgSender
                set senderAddr to address of msgSender
                set msgDate to (time received of msg) as string
                set msgContent to ""
                try
                    set msgContent to plain text content of msg
                end try
                if msgContent is "" then
                    try
                        set msgContent to content of msg
                    end try
                end if
                
                -- Keep first 500 chars and strip double quotes to avoid script breakages
                if length of msgContent > 500 then
                    set msgContent to text 1 thru 500 of msgContent
                end if
                
                -- Simple escape for line breaks
                set msgContentEsc to ""
                repeat with paragraphItem in paragraphs of msgContent
                    set msgContentEsc to msgContentEsc & paragraphItem & " "
                end repeat
                
                copy ("BOUNCE||" & msgId & "||" & msgSubject & "||" & senderName & "||" & senderAddr & "||" & msgDate & "||" & msgContentEsc) to end of outputLines
            end repeat
        end try
    end repeat
    
    -- Scan Outbox for stuck messages
    set outboxFolders to every folder whose name is "Outbox"
    repeat with i from 1 to count of outboxFolders
        set f to item i of outboxFolders
        try
            set msgs to every message of f
            repeat with j from 1 to count of msgs
                set msg to item j of msgs
                set msgId to id of msg
                set msgSubject to subject of msg
                set msgDate to ""
                try
                    set msgDate to (modification date of msg) as string
                end try
                set recipList to {}
                try
                    set recips to to recipients of msg
                    repeat with k from 1 to count of recips
                        set r to item k of recips
                        copy (address of email address of r) to end of recipList
                    end repeat
                end try
                copy ("OUTBOX||" & msgId & "||" & msgSubject & "||" & (recipList as string) & "||" & msgDate) to end of outputLines
            end repeat
        end try
    end repeat
    
    set AppleScript's text item delimiters to "\\n"
    return outputLines as string
end tell
  `;

  try {
    const rawResult = await runAppleScript(appleScript);
    const lines = rawResult.split('\\n').map(l => l.trim()).filter(Boolean);
    const logs = getDeliveryLogs(workspaceId);
    let newEntriesCount = 0;

    for (const line of lines) {
      const parts = line.split('||');
      if (parts[0] === 'BOUNCE') {
        const [_, msgId, subject, senderName, senderAddr, dateStr, content] = parts;
        const origSubject = cleanSubject(subject);
        
        let recipient = '';
        if (content) {
          const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}/g;
          const foundEmails = content.match(emailRegex) || [];
          const filteredEmails = foundEmails.filter(email => 
            !/daemon|postmaster|system|admin/i.test(email) && 
            email.toLowerCase() !== senderAddr.toLowerCase()
          );
          if (filteredEmails.length > 0) {
            recipient = filteredEmails[0];
          }
        }
        
        const matchingLog = logs.find(l => cleanSubject(l.subject) === origSubject);
        if (matchingLog && !recipient) {
          recipient = matchingLog.recipient;
        }
        
        if (!recipient) {
          recipient = 'Unknown Recipient';
        }

        const duplicate = logs.find(l => 
          l.recipient === recipient && 
          cleanSubject(l.subject) === origSubject && 
          l.status === 'failed' &&
          (l.error && l.error.includes('Outlook Bounce'))
        );

        if (!duplicate) {
          const successLog = logs.find(l => 
            l.recipient === recipient && 
            cleanSubject(l.subject) === origSubject && 
            l.status === 'success'
          );

          if (successLog) {
            successLog.status = 'failed';
            successLog.error = `Outlook Bounce: ${subject} (${senderName})`;
            successLog.timestamp = parseAppleScriptDate(dateStr).toISOString();
          } else {
            logs.push({
              id: `del_sync_${msgId}`,
              timestamp: parseAppleScriptDate(dateStr).toISOString(),
              recipient,
              subject: origSubject || subject,
              body: `Bounce notification body: ${content || 'No text content available'}`,
              status: 'failed',
              error: `Outlook Bounce: ${subject} (${senderName})`,
              source: 'Outlook Sync'
            });
          }
          newEntriesCount++;
        }
      } else if (parts[0] === 'OUTBOX') {
        const [_, msgId, subject, recipientsStr, dateStr] = parts;
        const recipients = recipientsStr ? recipientsStr.split(',').map(r => r.trim()) : ['Unknown Recipient'];

        for (const recipient of recipients) {
          const duplicate = logs.find(l => 
            l.recipient === recipient && 
            cleanSubject(l.subject) === cleanSubject(subject) && 
            l.status === 'failed' && 
            l.error === 'Stuck in Outlook Outbox'
          );

          if (!duplicate) {
            const successLog = logs.find(l => 
              l.recipient === recipient && 
              cleanSubject(l.subject) === cleanSubject(subject) && 
              l.status === 'success'
            );

            if (successLog) {
              successLog.status = 'failed';
              successLog.error = 'Stuck in Outlook Outbox';
            } else {
              logs.push({
                id: `del_sync_${msgId}_${recipient}`,
                timestamp: parseAppleScriptDate(dateStr).toISOString(),
                recipient,
                subject,
                body: 'Email stuck in Outlook Outbox.',
                status: 'failed',
                error: 'Stuck in Outlook Outbox',
                source: 'Outlook Sync'
              });
            }
            newEntriesCount++;
          }
        }
      }
    }

    if (newEntriesCount > 0) {
      saveDeliveryLogs(workspaceId, logs);
    }
    
    res.json({ success: true, count: newEntriesCount, logs: getDeliveryLogs(workspaceId) });
  } catch (error) {
    console.error('Error syncing Outlook:', error);
    res.status(500).json({ error: error.message });
  }
});

// Templates directory path
const templatesDir = path.join(process.cwd(), 'templates');

// Ensure templates folder and default templates exist
if (!fs.existsSync(templatesDir)) {
  fs.mkdirSync(templatesDir, { recursive: true });
}

const defaultHtmlPath = path.join(templatesDir, 'default-html.json');
if (!fs.existsSync(defaultHtmlPath)) {
  const defaultTemplate = {
    id: 'default-html',
    name: 'Default HTML Template',
    subject: 'Greetings {{First Name}}!',
    body: `<p>Dear {{First Name}},</p>\n<p>I hope this email finds you well.</p>\n<p>This is a batch test email sent via Outlook for Mac.</p>\n<p>Best regards,<br>Kyle</p>`,
    editorMode: 'html',
    imagesMap: {}
  };
  fs.writeFileSync(defaultHtmlPath, JSON.stringify(defaultTemplate, null, 2));
}

const defaultMarkdownPath = path.join(templatesDir, 'default-markdown.json');
if (!fs.existsSync(defaultMarkdownPath)) {
  const defaultMarkdownTemplate = {
    id: 'default-markdown',
    name: 'Default Markdown Template',
    subject: 'Greetings {{First Name}}!',
    body: `Dear {{First Name}},\n\nI hope this email finds you well.\n\nThis is a batch test email sent via Outlook for Mac.\n\nBest regards,\nKyle`,
    editorMode: 'markdown',
    imagesMap: {}
  };
  fs.writeFileSync(defaultMarkdownPath, JSON.stringify(defaultMarkdownTemplate, null, 2));
}

// 1. Get templates list summaries (lightweight metadata)
app.get('/api/templates', (req, res) => {
  try {
    if (!fs.existsSync(templatesDir)) {
      return res.json([]);
    }
    const files = fs.readdirSync(templatesDir);
    const templates = [];
    for (const file of files) {
      if (file.endsWith('.json')) {
        const filePath = path.join(templatesDir, file);
        try {
          const data = JSON.parse(fs.readFileSync(filePath, 'utf8'));
          templates.push({
            id: data.id,
            name: data.name,
            subject: data.subject,
            editorMode: data.editorMode
          });
        } catch (parseError) {
          console.error(`Error parsing template file ${file}:`, parseError);
        }
      }
    }
    res.json(templates);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 2. Get specific template details
app.get('/api/templates/:id', (req, res) => {
  try {
    const templatePath = path.join(templatesDir, `${req.params.id}.json`);
    if (!fs.existsSync(templatePath)) {
      return res.status(404).json({ error: 'Template not found' });
    }
    const data = JSON.parse(fs.readFileSync(templatePath, 'utf8'));
    res.json(data);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 3. Save or update template
app.post('/api/templates', (req, res) => {
  try {
    const { id, name, subject, body, editorMode, imagesMap } = req.body;
    if (!name || !subject || !body || !editorMode) {
      return res.status(400).json({ error: 'Missing required template fields (name, subject, body, editorMode)' });
    }
    
    const templateId = id || `template_${Date.now()}`;
    const template = {
      id: templateId,
      name,
      subject,
      body,
      editorMode,
      imagesMap: imagesMap || {}
    };
    
    const filePath = path.join(templatesDir, `${templateId}.json`);
    fs.writeFileSync(filePath, JSON.stringify(template, null, 2));
    res.json({ success: true, template });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 4. Delete template
app.delete('/api/templates/:id', (req, res) => {
  try {
    const templatePath = path.join(templatesDir, `${req.params.id}.json`);
    if (fs.existsSync(templatePath)) {
      fs.unlinkSync(templatePath);
      res.json({ success: true });
    } else {
      res.status(404).json({ error: 'Template not found' });
    }
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.listen(port, () => {
  console.log(`Backend server running at http://localhost:${port}`);
});
