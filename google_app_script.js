// ─── Landmark CRM — Google Apps Script ────────────────────────────────────
// Handles both:
//   GET  (no params)      → return all leads as JSON
//   GET  (?action=updateReply&phone=...&lastReply=...&...) → write back reply

function doGet(e) {
  try {
    var params = e && e.parameter ? e.parameter : {};
    // ── Add new lead mode ─────────────────────────────────────────────
if (params.action === 'addLead') {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.appendRow([
    params.Name     || '',
    params.Phone    || '',
    params.LastMsg  || '',
    params.Interest || '',
    params.Status   || 'New',
    params.Time     || new Date().toLocaleTimeString('en-IN',{hour:'2-digit',minute:'2-digit'}),
    '', ''
  ]);
  return buildResponse({ success: true, added: params.Name });
}

    // ── Write-back mode ──────────────────────────────────────────────────
    if (params.action === 'updateReply') {
      return handleUpdateReply(params);
    }

    // ── Read all leads mode ──────────────────────────────────────────────
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data  = sheet.getDataRange().getValues();
    var keys  = data[0];

    var leads = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      // Skip completely empty rows
      if (!row[0] && !row[1]) continue;
      var obj = {};
      for (var j = 0; j < keys.length; j++) {
        obj[String(keys[j])] = row[j] !== undefined && row[j] !== null ? String(row[j]) : '';
      }
      leads.push(obj);
    }

    return buildResponse({ leads: leads, total: leads.length });

  } catch(err) {
    return buildResponse({ error: err.message });
  }
}

// ─── Write reply back to the matching row ─────────────────────────────────
function handleUpdateReply(params) {
  try {
    var sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data    = sheet.getDataRange().getValues();
    var headers = data[0];

    // Find column indexes (case-insensitive)
    var phoneCol   = findCol(headers, 'Phone');
    var replyCol   = findCol(headers, 'LastReply');
    var repliedCol = findCol(headers, 'RepliedAt');
    var statusCol  = findCol(headers, 'Status');

    var bodyPhone  = String(params.phone  || '').replace(/\D/g, '');
    var lastReply  = String(params.lastReply  || '');
    var repliedAt  = String(params.repliedAt  || '');
    var newStatus  = String(params.status     || 'Contacted');

    if (!bodyPhone) {
      return buildResponse({ success: false, error: 'No phone number provided' });
    }

    // Find the row with this phone number
    for (var i = 1; i < data.length; i++) {
      var rowPhone = String(data[i][phoneCol] || '').replace(/\D/g, '');
      // Match last 10 digits to handle country code differences
      if (rowPhone.slice(-10) === bodyPhone.slice(-10) && bodyPhone.length >= 10) {
        if (replyCol   >= 0) sheet.getRange(i + 1, replyCol   + 1).setValue(lastReply);
        if (repliedCol >= 0) sheet.getRange(i + 1, repliedCol + 1).setValue(repliedAt);
        if (statusCol  >= 0) sheet.getRange(i + 1, statusCol  + 1).setValue(newStatus);
        return buildResponse({ success: true, row: i + 1, phone: rowPhone });
      }
    }

    // Phone not found — append as a note in a new row instead of failing silently
    return buildResponse({ success: false, error: 'Phone not found: ' + bodyPhone });

  } catch(err) {
    return buildResponse({ success: false, error: err.message });
  }
}

// ─── Helper: find column index by name (case-insensitive) ─────────────────
function findCol(headers, name) {
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).toLowerCase() === name.toLowerCase()) return i;
  }
  return -1;
}

// ─── Helper: build JSON response with CORS headers ────────────────────────
function buildResponse(obj) {
  var output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ─── doPost: kept for future webhook integrations ─────────────────────────
function doPost(e) {
  try {
    var body  = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.appendRow([
      body.Name     || '',
      body.Phone    || '',
      body.LastMsg  || '',
      body.Interest || '',
      body.Status   || 'New',
      new Date().toLocaleTimeString('en-IN', {hour:'2-digit', minute:'2-digit'}),
      '', ''
    ]);
    return buildResponse({ success: true });
  } catch(err) {
    return buildResponse({ error: err.message });
  }
}
