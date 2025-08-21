// Code.gs

/**
 * Adds “Email → Upload & Send Emails” menu (optional if using in‑sheet button).
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Email')
    .addItem('Upload & Send Emails', 'showFilePicker')
    .addToUi();
}

/**
 * Pops up the inline picker with a fixed 600×800 preview at 30% scale,
 * plus spinner/status while the send script is running.
 */
function showFilePicker() {
  var scale = 0.3;
  var previewW = 300, previewH = 400;
  var iframeW = previewW / scale, iframeH = previewH / scale;

  var html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: sans-serif; text-align:center; padding:16px; }
    #fileInput { display:none; }
    .btn {
      background:#1a73e8; color:#fff; border:2px solid #1a73e8;
      border-radius:4px; padding:8px 16px; margin:8px;
      font-size:14px; cursor:pointer;
      transition:background .2s,border-width .2s;
      display:inline-block;
    }
    .btn:hover { background:#1558b0; border-width:3px; }
    #fileNameDisplay { font-weight:bold; margin-left:8px; }

    /* preview pane */
    #previewContainer {
      width:${previewW}px;
      height:${previewH}px;
      overflow:hidden;
      border:1px solid #ccc;
      margin:12px auto;
    }
    #previewFrame {
      width:${iframeW}px;
      height:${iframeH}px;
      transform:scale(${scale});
      transform-origin:top left;
      border:none;
      display:block;
    }

    /* spinner */
    .spinner {
      display: none;
      width:24px; height:24px;
      border:3px solid #f3f3f3;
      border-top:3px solid #1a73e8;
      border-radius:50%;
      animation:spin 1s linear infinite;
      margin-right:8px;
      vertical-align:middle;
    }
    @keyframes spin { to { transform:rotate(360deg); } }

    /* status text */
    #status {
      display: none;
      font-size:14px;
      margin-top:8px;
      color:#333;
    }
  </style>
</head>
<body>
  <div style="margin-bottom:12px;">
    <label class="btn" for="fileInput">Choose HTML File</label>
    <input type="file" id="fileInput" accept=".html" onchange="handleFile()" />
    <span id="fileNameDisplay"></span>
  </div>

  <div id="previewContainer">
    <iframe id="previewFrame"></iframe>
  </div>

  <div style="margin-top:12px;">
    <span id="spinner" class="spinner"></span>
    <button class="btn" id="sendBtn" disabled onclick="sendEmails()">Upload & Send</button>
  </div>
  <div id="status">Sending… please wait</div>

  <script>
    var rawHtml = '';
    function handleFile() {
      var input = document.getElementById('fileInput');
      if (!input.files.length) return;
      var file = input.files[0];
      document.getElementById('fileNameDisplay').textContent = file.name;
      var reader = new FileReader();
      reader.onload = function(e) {
        rawHtml = e.target.result;
        document.getElementById('previewFrame').srcdoc = rawHtml;
        document.getElementById('sendBtn').disabled = false;
      };
      reader.readAsText(file);
    }
    function sendEmails() {
      // Show spinner & status
      document.getElementById('spinner').style.display = 'inline-block';
      document.getElementById('status').style.display  = 'block';
      document.getElementById('sendBtn').disabled      = true;

      // Call server function
      google.script.run
        .withSuccessHandler(function(msg) {
          alert(msg);
          google.script.host.close();
        })
        .withFailureHandler(function(err) {
          alert('Error: ' + err.message);
          // Hide spinner & status, re-enable button
          document.getElementById('spinner').style.display = 'none';
          document.getElementById('status').style.display  = 'none';
          document.getElementById('sendBtn').disabled      = false;
        })
        .processTemplateAndSend(rawHtml);
    }
  </script>
</body>
</html>`;

  // Size the dialog to show everything at once
  var dialog = HtmlService.createHtmlOutput(html)
    .setWidth(previewW + 80)   // ~680px
    .setHeight(previewH + 200); // ~1000px
  SpreadsheetApp.getUi().showModalDialog(dialog, 'Upload & Preview HTML');
}

/**
 * Server‐side: Reads the HTML, replaces [First Name]/[Last Name], 
 * sends personalized emails, logs quota/status, and returns a summary.
 */
function processTemplateAndSend(rawHtml) {
  Logger.log('▶️ Script start');
  var startQuota = MailApp.getRemainingDailyQuota();
  Logger.log('Starting quota: ' + startQuota);

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  if (!sheet) throw new Error('Sheet1 not found');

  var HEADER=1, FN=1, LN=2, EM=3, SU=4;
  var last = sheet.getLastRow();
  Logger.log('lastRow=' + last);
  if (last <= HEADER) return 'No data to send';

  var data = sheet.getRange(HEADER+1,1,last-HEADER,sheet.getLastColumn()).getValues();
  Logger.log('Loaded rows: ' + data.length);

  var count = 0;
  data.forEach(function(r,i){
    var email  = r[EM-1].toString().trim();
    if (!email) return;
    var first  = r[FN-1].toString().trim();
    var last   = r[LN-1].toString().trim();
    var subj   = r[SU-1].toString().trim();

    var body = rawHtml
      .replace(/\[first name\]/gi, first)
      .replace(/\[last name\]/gi,  last);

    MailApp.sendEmail({ to: email, subject: subj, htmlBody: body });
    count++;
    Logger.log('Sent to ' + email + '; remaining quota: ' + MailApp.getRemainingDailyQuota());
  });

  var endQuota = MailApp.getRemainingDailyQuota();
  Logger.log('✅ Done: sent ' + count + ', final quota ' + endQuota);

  return [
    '✅ Sent ' + count + ' emails.',
    'Starting quota: ' + startQuota,
    'Remaining quota: ' + endQuota
  ].join('\n');
}
