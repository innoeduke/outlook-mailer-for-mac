import express from 'express';
import cors from 'cors';
import { spawn } from 'child_process';

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

app.listen(port, () => {
  console.log(`Backend server running at http://localhost:${port}`);
});
