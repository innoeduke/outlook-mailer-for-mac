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

// Endpoint to check Outlook's status
app.get('/api/status', async (req, res) => {
  try {
    // Check if the Microsoft Outlook process is running
    const script = `tell application "System Events" to set isRunning to (name of processes) contains "Microsoft Outlook"
return isRunning`;
    
    const result = await runAppleScript(script);
    
    // Result will be "true" or "false"
    res.json({ running: result === 'true' });
  } catch (error) {
    console.error('Error checking Outlook status:', error);
    res.json({ running: false, error: error.message });
  }
});

// Endpoint to send a single email via AppleScript + Outlook
app.post('/api/send-email', async (req, res) => {
  const { recipient, subject, body } = req.body;

  if (!recipient || !subject || !body) {
    return res.status(400).json({ error: 'Missing recipient, subject, or body' });
  }

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
      res.json({ success: true });
    } else {
      res.status(500).json({ success: false, error: result });
    }
  } catch (error) {
    console.error(`Failed to send email to ${recipient}:`, error);
    res.status(500).json({ success: false, error: error.message });
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
