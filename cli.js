#!/usr/bin/env node

import fs from 'fs';
import path from 'path';
import { spawn } from 'child_process';
import Papa from 'papaparse';
import { marked } from 'marked';

// Color codes for CLI output
const RESET = '\x1b[0m';
const GREEN = '\x1b[32m';
const RED = '\x1b[31m';
const YELLOW = '\x1b[33m';
const BLUE = '\x1b[34m';
const BOLD = '\x1b[1m';

function printUsage() {
  console.log(`
${BOLD}Outlook Batch Mailer - Command Line Interface${RESET}

${BOLD}Usage:${RESET}
  node cli.js --csv <csv-file> --subject "<subject>" --body <body-html-file-or-string> [--delay <seconds>] [--email-col <column-name>]

${BOLD}Options:${RESET}
  --csv        Path to the contacts CSV file (Required)
  --subject    Subject line. Can contain {{Placeholders}} matching CSV headers (Required)
  --body       Path to an HTML file OR a raw HTML string. Can contain {{Placeholders}} (Required)
  --delay      Time in seconds to wait between emails (Optional, Default: 3)
  --email-col  Exact header name of the recipient email column (Optional, auto-detects if omitted)

${BOLD}Examples:${RESET}
  node cli.js --csv contacts.csv --subject "Hello {{First Name}}!" --body template.html --delay 2
  node cli.js --csv contacts.csv --subject "Quick update" --body "<p>Hi {{First Name}}, this is a test.</p>" --email-col "Primary Email"
`);
}

// Simple argv parser
const args = {};
for (let i = 2; i < process.argv.length; i++) {
  if (process.argv[i].startsWith('--')) {
    const key = process.argv[i].slice(2);
    const val = process.argv[i + 1];
    
    // Check if value exists and doesn't start with --
    if (val && !val.startsWith('--')) {
      args[key] = val;
      i++;
    } else {
      args[key] = true;
    }
  }
}

async function run() {
  const csvPath = args.csv;
  const subjectTemplate = args.subject;
  const bodyInput = args.body;
  const delaySec = parseFloat(args.delay) || 3;
  let emailCol = args['email-col'];

  if (!csvPath || !subjectTemplate || !bodyInput) {
    printUsage();
    process.exit(1);
  }

  // 1. Resolve Body Template (File or raw String)
  let bodyTemplate = '';
  if (fs.existsSync(bodyInput)) {
    try {
      bodyTemplate = fs.readFileSync(bodyInput, 'utf8');
      console.log(`${BLUE}ℹ Load body template from file: ${bodyInput}${RESET}`);
    } catch (err) {
      console.error(`${RED}✗ Error reading body file: ${err.message}${RESET}`);
      process.exit(1);
    }
  } else {
    bodyTemplate = bodyInput;
  }

  // 2. Read and Parse CSV
  if (!fs.existsSync(csvPath)) {
    console.error(`${RED}✗ CSV file not found at path: ${csvPath}${RESET}`);
    process.exit(1);
  }

  let csvContent = '';
  try {
    csvContent = fs.readFileSync(csvPath, 'utf8');
  } catch (err) {
    console.error(`${RED}✗ Error reading CSV file: ${err.message}${RESET}`);
    process.exit(1);
  }

  const parsed = Papa.parse(csvContent, {
    header: true,
    skipEmptyLines: true
  });

  if (parsed.errors && parsed.errors.length > 0) {
    console.warn(`${YELLOW}⚠ CSV Parsing encountered errors: ${parsed.errors[0].message}${RESET}`);
  }

  const rows = parsed.data;
  if (!rows || rows.length === 0) {
    console.error(`${RED}✗ No contacts parsed from CSV file.${RESET}`);
    process.exit(1);
  }

  const headers = Object.keys(rows[0]);
  console.log(`${GREEN}✔ Parsed ${rows.length} contacts from ${csvPath}${RESET}`);

  // 3. Determine Email Column
  if (!emailCol) {
    const detected = headers.find(h => /email/i.test(h));
    if (detected) {
      emailCol = detected;
      console.log(`${BLUE}ℹ Auto-detected email column: "${emailCol}"${RESET}`);
    } else {
      emailCol = headers[0];
      console.log(`${YELLOW}⚠ Could not auto-detect email column. Using first column: "${emailCol}"${RESET}`);
    }
  } else if (!headers.includes(emailCol)) {
    console.error(`${RED}✗ Specified email column "${emailCol}" not found in CSV. Headers: [${headers.join(', ')}]${RESET}`);
    process.exit(1);
  }

  // 4. Verify Microsoft Outlook is running
  console.log(`${BLUE}ℹ Verifying Outlook status...${RESET}`);
  const isOutlookRunning = await checkOutlookRunning();
  if (!isOutlookRunning) {
    console.error(`${RED}✗ Error: Microsoft Outlook is not open or running on your Mac.${RESET}`);
    console.error(`${RED}  Please launch Outlook and sign in before executing.${RESET}`);
    process.exit(1);
  }
  console.log(`${GREEN}✔ Outlook is active and connected.${RESET}`);

  // 5. Placeholder Substitution Function
  const interpolate = (template, row) => {
    let result = template;
    Object.entries(row).forEach(([key, value]) => {
      const escapedKey = key.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
      const regex = new RegExp(`{{\\s*${escapedKey}\\s*}}`, 'gi');
      result = result.replace(regex, value || '');
    });
    return result;
  };

  // 6. Sending Loop
  console.log(`\n${BOLD}Starting batch send...${RESET}`);
  console.log(`Delay set to ${delaySec} seconds between emails.\n`);

  let successCount = 0;
  let errorCount = 0;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const recipient = row[emailCol];

    if (!recipient) {
      console.log(`${YELLOW}⚠ [${i + 1}/${rows.length}] Skipped: No email address in column "${emailCol}"${RESET}`);
      continue;
    }

    const subject = interpolate(subjectTemplate, row);
    const rawBody = interpolate(bodyTemplate, row);
    const body = marked.parse(rawBody);

    // Double check Outlook is still open before sending
    const active = await checkOutlookRunning();
    if (!active) {
      console.error(`\n${RED}✗ Lost connection to Outlook! Process aborted. Progress paused at item ${i + 1}/${rows.length}.${RESET}`);
      break;
    }

    process.stdout.write(`Sending to ${recipient}... `);

    try {
      const sendResult = await sendEmailViaAppleScript(recipient, subject, body);
      if (sendResult === 'SUCCESS') {
        successCount++;
        console.log(`${GREEN}Success${RESET}`);
      } else {
        errorCount++;
        console.log(`${RED}Failed (${sendResult})${RESET}`);
      }
    } catch (err) {
      errorCount++;
      console.log(`${RED}Failed Error: ${err.message}${RESET}`);
    }

    // Delay before next send
    if (i < rows.length - 1) {
      await new Promise(resolve => setTimeout(resolve, delaySec * 1000));
    }
  }

  console.log(`\n${BOLD}Summary:${RESET}`);
  console.log(`${GREEN}✔ Sent Successfully: ${successCount}${RESET}`);
  console.log(`${RED}✗ Failed: ${errorCount}${RESET}`);
  console.log(`Progress: ${successCount + errorCount}/${rows.length} completed.`);
}

// Check Outlook status via AppleScript
function checkOutlookRunning() {
  return new Promise((resolve) => {
    const script = `tell application "System Events" to set isRunning to (name of processes) contains "Microsoft Outlook"
return isRunning`;
    const child = spawn('osascript', []);
    let stdout = '';
    child.stdout.on('data', (d) => stdout += d.toString());
    child.on('close', () => resolve(stdout.trim() === 'true'));
    child.stdin.write(script);
    child.stdin.end();
  });
}

// Escape strings for AppleScript
function escapeForAppleScript(str) {
  if (!str) return '';
  return str.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}

// Send single email via AppleScript
function sendEmailViaAppleScript(recipient, subject, body) {
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

  return new Promise((resolve) => {
    const child = spawn('osascript', []);
    let stdout = '';
    let stderr = '';
    
    child.stdout.on('data', (d) => stdout += d.toString());
    child.stderr.on('data', (d) => stderr += d.toString());
    
    child.on('close', (code) => {
      if (code === 0) {
        resolve(stdout.trim());
      } else {
        resolve(`SHELL_ERROR: ${stderr.trim()}`);
      }
    });
    
    child.stdin.write(appleScript);
    child.stdin.end();
  });
}

run();
