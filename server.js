const express = require('express');
const multer = require('multer');
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');
const cors = require('cors');

const app = express();
const PORT = 3000;

// הגדרות
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// הגדרת תיקיות
const uploadsDir = path.join(__dirname, 'uploads');
const draftsDir = path.join(__dirname, 'drafts');

if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir);
}

if (!fs.existsSync(draftsDir)) {
  fs.mkdirSync(draftsDir);
}

// הגדרת multer
const storage = multer.diskStorage({
  destination: uploadsDir,
  filename: (req, file, cb) => {
    cb(null, Date.now() + '-' + file.originalname);
  }
});
const upload = multer({ storage });

// API endpoint
app.post('/create-draft', upload.single('attachment'), (req, res) => {
  try {
    const { subject, recipients, body } = req.body;
    const attachmentPath = req.file ? req.file.path : null;
    const recipientList = recipients.split(',').map(r => r.trim());

    // יצירת PowerShell script עבור כל recipient
    recipientList.forEach((recipient, index) => {
      // שמור בתיקיית TEMP של Windows
      const tempDir = process.env.TEMP || 'C:\\Windows\\Temp';
      const psScriptPath = path.join(tempDir, `outlook_draft_${Date.now()}_${index}.ps1`);
      const jsonPath = path.join(tempDir, `outlook_data_${Date.now()}_${index}.json`);
      
      // שמור את הנתונים בJSON
      const data = {
        subject: subject,
        body: body,
        recipient: recipient,
        attachmentPath: attachmentPath
      };
      
      fs.writeFileSync(jsonPath, JSON.stringify(data, null, 2), 'utf8');
      
      // PowerShell script שקורא מJSON
      const psScript = `
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

try {
    $jsonPath = "${jsonPath.replace(/\\/g, '\\\\')}"
    $jsonContent = Get-Content -Path $jsonPath -Encoding UTF8 | ConvertFrom-Json
    
    $subject = $jsonContent.subject
    $body = $jsonContent.body
    $recipient = $jsonContent.recipient
    $attachment = $jsonContent.attachmentPath
    
    Write-Host "Creating Outlook..."
    $outlook = New-Object -ComObject Outlook.Application
    
    $namespace = $outlook.GetNamespace("MAPI")
    $mailItem = $outlook.CreateItem(0)
    
    $mailItem.Subject = $subject
    $mailItem.Body = $body
    $mailItem.To = $recipient
    
    if ($attachment -and $attachment -ne "") {
        $mailItem.Attachments.Add($attachment) | Out-Null
    }
    
    $mailItem.Save()
    $mailItem.Display()
    Write-Host "Done!"
    
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mailItem) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
}
catch {
    Write-Host "Error: $_"
    exit 1
}
`;

      fs.writeFileSync(psScriptPath, psScript, 'utf8');

      try {
        const proc = spawn('powershell.exe', [
          '-NoProfile',
          '-ExecutionPolicy', 'Bypass',
          '-File', psScriptPath
        ], {
          detached: false,
          stdio: ['ignore', 'pipe', 'pipe'],
          shell: true
        });
        
        let stdout = '';
        let stderr = '';
        
        proc.stdout.on('data', (data) => {
          stdout += data.toString('utf8');
          console.log(`📤 PowerShell output: ${data}`);
        });
        
        proc.stderr.on('data', (data) => {
          stderr += data.toString('utf8');
          console.log(`❌ PowerShell error: ${data}`);
        });
        
        proc.on('close', (code) => {
          console.log(`PowerShell סיים עם קוד: ${code}`);
          if (code !== 0) {
            console.log(`Full error output: ${stderr}`);
          }
          
          // ניקוי הקבצים
          setTimeout(() => {
            [psScriptPath, jsonPath].forEach(file => {
              try {
                if (fs.existsSync(file)) {
                  fs.unlinkSync(file);
                  console.log(`Cleaned: ${file}`);
                }
              } catch (e) {}
            });
          }, 2000);
        });
        
        
      } catch (error) {
        console.error(`❌ שגיאה:`, error.message);
      } finally {
        setTimeout(() => {
          if (fs.existsSync(psScriptPath)) {
            fs.unlinkSync(psScriptPath);
          }
        }, 3000);
      }
    });

    res.json({ success: true, message: `${recipientList.length} טיוטות נוצרו!` });

  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// טעינת ה-Web App
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// הפעלת השרת
app.listen(PORT, () => {
  console.log(`🚀 Server is running on http://localhost:${PORT}`);
  console.log('💡 פתח את http://localhost:3000 בדפדפן שלך');
});