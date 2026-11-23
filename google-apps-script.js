/**
 * BTS Treas Tools — Google Drive, Gmail & Sheets Automation Script (v3)
 *
 * SETUP INSTRUCTIONS:
 * 1. Go to https://script.google.com and open your project.
 * 2. Replace all code in Code.gs with this entire file and press Ctrl+S (Save).
 * 3. (Recommended — for Automatic Gmail Signature, Blank "To" field & Fast Draft Search):
 *    In the left sidebar, click "+" next to "Services", select "Gmail API", and click "Add".
 * 4. At the top toolbar, select "doGet" in the function dropdown and click "Run" once
 *    to approve the Google permissions prompt ("Review permissions" -> "Advanced" -> "Allow").
 *    (This approves Drive, Gmail, Google Sheets, and External Token Verification scopes.)
 * 5. Click "Deploy" (top right) -> "Manage deployments" -> Edit -> Version: "New version" -> "Deploy".
 *    - Type: "Web app"
 *    - Execute as: "Me"
 *    - Who has access: "Anyone" (Requests still require a verified treas@bts-crew.com Google OAuth token)
 */

var APP_SCRIPT_VERSION = 3;
var AUTHORIZED_EMAIL = 'treas@bts-crew.com';
var DEFAULT_ALLOWED_CLIENT_IDS = [
  '523917169511-qp7hvnb54uhho8huevg0r6e76bil27ub.apps.googleusercontent.com'
];

function doGet(e) {
  try {
    // Touch Drive, Gmail, Sheets, and UrlFetch scopes so running doGet once approves all permissions
    DriveApp.getRootFolder().getName();
    GmailApp.getInboxThreads(0, 1);
    try {
      SpreadsheetApp.getActiveSpreadsheet();
    } catch (ssErr) {}
    UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo', { muteHttpExceptions: true });

    var hasAdvancedGmail = (typeof Gmail !== 'undefined' && Boolean(Gmail.Users));
    var statusText = [
      '✓ BTS Treas Tools Google Apps Script (v' + APP_SCRIPT_VERSION + ') is active and authorized!',
      '',
      '• Google Drive scope: Connected',
      '• Gmail scope: Connected',
      '• Google Sheets scope: Connected',
      '• Token verification scope (UrlFetchApp): Connected',
      '• Advanced Gmail API service enabled: ' + (hasAdvancedGmail ? 'Yes' : 'No (optional — enable "Gmail API" under Services +)')
    ].join('\n');

    return ContentService
      .createTextOutput(statusText)
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService
      .createTextOutput('Error checking permissions: ' + (err && err.message ? err.message : String(err)))
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function doPost(e) {
  try {
    var rawContents = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    // Reject oversized payloads (>25MB) to prevent quota exhaustion or abuse
    if (rawContents.length > 25 * 1024 * 1024) {
      return jsonResponse({ success: false, scriptVersion: APP_SCRIPT_VERSION, error: 'Payload exceeds maximum limit (25MB).' });
    }

    var payload = JSON.parse(rawContents);
    var auth = verifyGoogleAuth(payload);
    if (!auth.authorized) {
      return jsonResponse({
        success: false,
        code: 'UNAUTHORIZED',
        scriptVersion: APP_SCRIPT_VERSION,
        error: auth.error
      });
    }

    var action = String(payload.action || 'saveDocument');
    if (action === 'status') {
      return jsonResponse({
        success: true,
        scriptVersion: APP_SCRIPT_VERSION,
        email: auth.email,
        hasAdvancedGmail: (typeof Gmail !== 'undefined' && Boolean(Gmail.Users))
      });
    }
    if (action === 'syncOverviewSheet') {
      return jsonResponse(handleSyncOverviewSheet(payload));
    }

    return jsonResponse(handleSaveDocument(payload));
  } catch (err) {
    return jsonResponse({
      success: false,
      scriptVersion: APP_SCRIPT_VERSION,
      error: err && err.message ? err.message : String(err)
    });
  }
}

function getAllowedClientIds() {
  var ids = DEFAULT_ALLOWED_CLIENT_IDS.slice();
  try {
    var extra = PropertiesService.getScriptProperties().getProperty('ALLOWED_CLIENT_IDS');
    if (extra) {
      var parts = extra.split(',');
      for (var i = 0; i < parts.length; i++) {
        var trimmed = parts[i].trim();
        if (trimmed && ids.indexOf(trimmed) === -1) {
          ids.push(trimmed);
        }
      }
    }
  } catch (err) {
    Logger.log('Could not read ALLOWED_CLIENT_IDS property: ' + err);
  }
  return ids;
}

function verifyGoogleAuth(payload) {
  var token = String(payload && payload.authToken ? payload.authToken : '').trim();
  var tokenType = String(payload && payload.authTokenType ? payload.authTokenType : 'access_token').toLowerCase();

  if (!token) {
    return {
      authorized: false,
      error: 'Authentication required. Please sign in with Google (' + AUTHORIZED_EMAIL + ').'
    };
  }

  var paramName = (tokenType === 'id_token') ? 'id_token' : 'access_token';
  var url = 'https://oauth2.googleapis.com/tokeninfo?' + paramName + '=' + encodeURIComponent(token);

  try {
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) {
      return {
        authorized: false,
        error: 'Google session expired or invalid. Please sign in again with ' + AUTHORIZED_EMAIL + '.'
      };
    }

    var info = JSON.parse(res.getContentText() || '{}');
    var email = String(info.email || '').trim().toLowerCase();
    var emailVerified = (info.email_verified === true || String(info.email_verified).toLowerCase() === 'true');
    if (!emailVerified || email !== AUTHORIZED_EMAIL.toLowerCase()) {
      return {
        authorized: false,
        error: 'Access denied for account "' + (info.email || 'unknown') + '". Only ' + AUTHORIZED_EMAIL + ' is authorized.'
      };
    }

    var allowedIds = getAllowedClientIds();
    var tokenAud = String(info.aud || '').trim();
    var tokenAzp = String(info.azp || '').trim();
    if (allowedIds.indexOf(tokenAud) === -1 && allowedIds.indexOf(tokenAzp) === -1) {
      return {
        authorized: false,
        error: 'Token client ID is not authorized for BTS Treas Tools.'
      };
    }

    if (info.expires_in !== undefined) {
      var expiresIn = parseInt(info.expires_in, 10);
      if (isNaN(expiresIn) || expiresIn <= 0) {
        return { authorized: false, error: 'Google token has expired. Please try again.' };
      }
    } else if (info.exp !== undefined) {
      var expMs = Number(info.exp) * 1000;
      if (isNaN(expMs) || expMs <= Date.now()) {
        return { authorized: false, error: 'Google ID token has expired. Please try again.' };
      }
    }

    return { authorized: true, email: info.email };
  } catch (err) {
    return {
      authorized: false,
      error: 'Could not verify Google authentication: ' + (err && err.message ? err.message : String(err))
    };
  }
}

function resolveEventFolder(parentFolder, eventName, requestedFolderId) {
  if (requestedFolderId) {
    try {
      var byId = DriveApp.getFolderById(requestedFolderId);
      if (byId && !byId.isTrashed()) {
        return { folder: byId, existed: true };
      }
    } catch (folderLookupErr) {
      Logger.log('Stored folder ID lookup failed, falling back to name lookup: ' + folderLookupErr);
    }
  }

  var existingFolders = parentFolder.getFoldersByName(eventName);
  if (existingFolders.hasNext()) {
    return { folder: existingFolders.next(), existed: true };
  }
  return { folder: parentFolder.createFolder(eventName), existed: false };
}

function deleteMatchingDrafts(subject) {
  var replaced = false;
  try {
    if (typeof Gmail !== 'undefined' && Gmail.Users && Gmail.Users.Drafts) {
      var cleanQuery = subject.replace(/["\\]/g, ' ').trim();
      var listRes = Gmail.Users.Drafts.list('me', { q: 'subject:"' + cleanQuery + '"' });
      var drafts = (listRes && listRes.drafts) ? listRes.drafts : [];
      for (var i = 0; i < drafts.length; i++) {
        var detail = Gmail.Users.Drafts.get('me', drafts[i].id, {
          format: 'metadata',
          metadataHeaders: ['Subject']
        });
        var headers = (detail && detail.message && detail.message.payload && detail.message.payload.headers)
          ? detail.message.payload.headers
          : [];
        var subjHeader = '';
        for (var h = 0; h < headers.length; h++) {
          if (String(headers[h].name || '').toLowerCase() === 'subject') {
            subjHeader = headers[h].value || '';
            break;
          }
        }
        if (subjHeader === subject) {
          Gmail.Users.Drafts.remove('me', drafts[i].id);
          replaced = true;
        }
      }
      return replaced;
    }

    var existingDrafts = GmailApp.getDrafts();
    for (var d = 0; d < existingDrafts.length; d++) {
      if (existingDrafts[d].getMessage().getSubject() === subject) {
        existingDrafts[d].deleteDraft();
        replaced = true;
      }
    }
  } catch (draftCheckErr) {
    Logger.log('Could not check existing drafts: ' + draftCheckErr);
  }
  return replaced;
}

function handleSaveDocument(payload) {
  var eventName = (payload.eventName || 'Untitled Event').trim();
  var fileName = (payload.fileName || '').trim();
  var pdfBase64 = payload.pdfBase64 || '';
  var parentFolderId = extractFolderId(payload.parentFolderId || '');
  var requestedFolderId = extractFolderId(payload.folderId || '');
  var documentType = String(payload.documentType || (fileName.indexOf('INV') === 0 ? 'invoice' : 'blue_form')).toLowerCase();
  var isInvoice = (documentType === 'invoice');

  // Strip control characters only so Drive folders and email drafts keep punctuation such as colons
  eventName = eventName.replace(/[\x00-\x1f\x7f]/g, '').substring(0, 100).trim() || 'Untitled Event';

  if (!fileName) {
    fileName = isInvoice ? ('INV_' + eventName + '.pdf') : ('BF_' + eventName + '.pdf');
  }

  if (!pdfBase64) {
    return { success: false, scriptVersion: APP_SCRIPT_VERSION, error: 'Missing PDF data.' };
  }
  if (pdfBase64.length > 20 * 1024 * 1024) {
    return { success: false, scriptVersion: APP_SCRIPT_VERSION, error: 'PDF data exceeds maximum size limit (20MB).' };
  }

  // Strip control characters and replace path-unsafe characters (preserving colons) in Drive/Gmail filenames
  fileName = fileName.replace(/[\x00-\x1f\x7f]/g, '').replace(/[\\/*?"<>|]/g, '_').substring(0, 100).trim() || (isInvoice ? 'INV_Event.pdf' : 'BF_Event.pdf');

  // 1. Decode the PDF bytes
  var pdfBytes = Utilities.base64Decode(pdfBase64);
  var pdfBlob = Utilities.newBlob(pdfBytes, 'application/pdf', fileName);

  // 2. Find or create the "{event name}" folder inside the Parent Drive Folder (or root Drive)
  var parentFolder = parentFolderId
    ? DriveApp.getFolderById(parentFolderId)
    : DriveApp.getRootFolder();

  var resolvedFolder = resolveEventFolder(parentFolder, eventName, requestedFolderId);
  var eventFolder = resolvedFolder.folder;
  var folderExisted = resolvedFolder.existed;

  // Trash any existing file with the exact same name in this folder
  var existingFiles = eventFolder.getFilesByName(fileName);
  var fileExisted = existingFiles.hasNext();
  while (existingFiles.hasNext()) {
    var oldFile = existingFiles.next();
    try {
      oldFile.setTrashed(true);
    } catch (trashErr) {
      Logger.log('Could not trash old file: ' + trashErr);
    }
  }

  // Save the PDF file into the event folder
  var savedFile = eventFolder.createFile(pdfBlob);

  // 3. Build the Gmail Draft only for Blue Forms (skipped for Invoices)
  var skipDraft = (payload.skipDraft === true || isInvoice);
  var draftCreated = false;
  var draftExisted = false;
  if (!skipDraft) {
    var subject = 'Blue Form - ' + eventName;
    var plainBody = 'Hi,\n\nPlease find attached the Blue Form for ' + eventName + '.';
    var htmlBody =
      '<div style="font-family: Arial, sans-serif; font-size: 14px; color: #222222; line-height: 1.5;">' +
      '<p style="margin: 0 0 12px 0;">Hi,</p>' +
      '<p style="margin: 0 0 12px 0;">Please find attached the Blue Form for ' + escapeHtml(eventName) + '.</p>' +
      '</div>';

    var signatureHtml = getGmailSignatureHtml();
    if (signatureHtml) {
      htmlBody += signatureHtml;
    }

    draftExisted = deleteMatchingDrafts(subject);
    createGmailDraft(subject, plainBody, htmlBody, pdfBlob);
    draftCreated = true;
  }

  return {
    success: true,
    scriptVersion: APP_SCRIPT_VERSION,
    folderId: eventFolder.getId(),
    folderName: eventFolder.getName(),
    folderUrl: eventFolder.getUrl(),
    folderExisted: folderExisted,
    fileUrl: savedFile.getUrl(),
    fileReplaced: fileExisted,
    draftCreated: draftCreated,
    draftReplaced: draftExisted
  };
}

function extractSpreadsheetId(input) {
  var str = String(input || '').trim();
  if (!str) return '';
  var sheetMatch = str.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (sheetMatch) return sheetMatch[1];
  var idMatch = str.match(/^[a-zA-Z0-9_-]{20,}$/);
  return idMatch ? idMatch[0] : '';
}

function normalizeHeaderText(val) {
  if (val === null || val === undefined) return '';
  return String(val)
    .replace(/[\r\n\t\u00a0]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function stripHeaderDecorations(norm) {
  return String(norm || '')
    .replace(/\s*\([\s\S]*?\)\s*/g, ' ')
    .replace(/[£$€:*#\-–—]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function classifyHeaderCell(val) {
  var raw = normalizeHeaderText(val);
  if (!raw) return { event: 0, client: 0, total: 0 };
  var clean = stripHeaderDecorations(raw);

  var eventScore = 0;
  if (raw === 'event name') eventScore = 4;
  else if (clean === 'event name') eventScore = 3;
  else if (
    clean === 'event' ||
    clean === 'event title' ||
    clean === 'name of event' ||
    clean === 'event / show' ||
    clean === 'show name' ||
    clean === 'production name' ||
    clean === 'booking name' ||
    clean === 'event description'
  ) {
    eventScore = 2;
  }

  var clientScore = 0;
  if (raw === 'client') clientScore = 4;
  else if (clean === 'client' || clean === 'client name') clientScore = 3;
  else if (
    clean === 'client / society' ||
    clean === 'society / client' ||
    clean === 'client / soc' ||
    clean === 'client / dept' ||
    clean === 'society' ||
    clean === 'club / society' ||
    clean === 'customer' ||
    clean === 'organisation' ||
    clean === 'organization'
  ) {
    clientScore = 2;
  }

  var totalScore = 0;
  if (raw === 'total') totalScore = 4;
  else if (clean === 'total') totalScore = 3;
  else if (!/sub[\s\-]?total/.test(raw)) {
    if (
      clean === 'total cost' ||
      clean === 'total charge' ||
      clean === 'total price' ||
      clean === 'total amount' ||
      clean === 'event total' ||
      clean === 'invoice total' ||
      clean === 'grand total' ||
      clean === 'total due' ||
      clean === 'net total' ||
      clean === 'gross total' ||
      clean === 'total ex vat' ||
      clean === 'total exc vat' ||
      clean === 'total inc vat' ||
      clean === 'total cost to client'
    ) {
      totalScore = 2;
    }
  }

  return { event: eventScore, client: clientScore, total: totalScore };
}

function pickBestColumnsFromScores(colScores) {
  // Choose best distinct columns for event, client, total from per-column score objects
  var best = null;
  for (var e = 0; e < colScores.length; e++) {
    if (colScores[e].event <= 0) continue;
    for (var cl = 0; cl < colScores.length; cl++) {
      if (cl === e || colScores[cl].client <= 0) continue;
      for (var t = 0; t < colScores.length; t++) {
        if (t === e || t === cl || colScores[t].total <= 0) continue;
        var sum = colScores[e].event + colScores[cl].client + colScores[t].total;
        if (!best || sum > best.scoreSum) {
          best = {
            eventCol: e,
            clientCol: cl,
            totalCol: t,
            scoreSum: sum
          };
        }
      }
    }
  }
  return best;
}

function buildOverviewHeaderDiagnostics(values) {
  var scanLimit = Math.min(values.length, 25);
  var bestPartial = null;
  var bestPartialCount = 0;
  var previewRows = [];

  for (var r = 0; r < scanLimit; r++) {
    var row = values[r] || [];
    var nonEmpty = [];
    var eCol = -1;
    var cCol = -1;
    var tCol = -1;

    for (var c = 0; c < row.length; c++) {
      var text = String(row[c] === null || row[c] === undefined ? '' : row[c])
        .replace(/[\r\n\t\u00a0]+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
      if (text && nonEmpty.length < 15) nonEmpty.push(text);
      var scores = classifyHeaderCell(row[c]);
      if (scores.event > 0 && eCol === -1) eCol = c;
      if (scores.client > 0 && cCol === -1) cCol = c;
      if (scores.total > 0 && tCol === -1) tCol = c;
    }

    if (nonEmpty.length > 0 && previewRows.length < 6) {
      previewRows.push({
        rowNumber: r + 1,
        cells: nonEmpty
      });
    }

    var count = (eCol !== -1 ? 1 : 0) + (cCol !== -1 ? 1 : 0) + (tCol !== -1 ? 1 : 0);
    if (count > bestPartialCount) {
      bestPartialCount = count;
      bestPartial = {
        rowNumber: r + 1,
        found: {
          client: cCol !== -1,
          eventName: eCol !== -1,
          total: tCol !== -1
        },
        cells: nonEmpty
      };
    }
  }

  return {
    bestPartial: bestPartial,
    previewRows: previewRows
  };
}

function findOverviewColumnIndices(values) {
  var scanLimit = Math.min(values.length, 60);

  // Pass 1: single-row header detection (highest score sum wins; earliest row breaks ties)
  var bestSingle = null;
  for (var r = 0; r < scanLimit; r++) {
    var row = values[r] || [];
    var scores = [];
    for (var c = 0; c < row.length; c++) {
      scores.push(classifyHeaderCell(row[c]));
    }
    var match = pickBestColumnsFromScores(scores);
    if (match && (!bestSingle || match.scoreSum > bestSingle.scoreSum)) {
      bestSingle = {
        headerRowIndex: r,
        eventCol: match.eventCol,
        clientCol: match.clientCol,
        totalCol: match.totalCol,
        scoreSum: match.scoreSum
      };
    }
  }
  if (bestSingle) return bestSingle;

  // Pass 2: vertically merged / 2-row stacked header block (top row + immediate sub-header row)
  var bestMerged = null;
  for (var r2 = 0; r2 < scanLimit - 1; r2++) {
    var topRow = values[r2] || [];
    var subRow = values[r2 + 1] || [];
    var maxCols = Math.max(topRow.length, subRow.length);
    var combined = [];
    var fromTopCols = {};
    var fromSubCols = {};

    for (var c2 = 0; c2 < maxCols; c2++) {
      var topScore = classifyHeaderCell(topRow[c2]);
      var subScore = classifyHeaderCell(subRow[c2]);
      var subBlank = !normalizeHeaderText(subRow[c2]);

      // Allow top-row header when the sub-header cell below it is blank (vertically merged header)
      // or when sub-row itself provides the header label
      var eScore = Math.max(subBlank ? topScore.event : 0, subScore.event);
      var clScore = Math.max(subBlank ? topScore.client : 0, subScore.client);
      var tScore = Math.max(subBlank ? topScore.total : 0, subScore.total);

      if (subBlank && (topScore.event > 0 || topScore.client > 0 || topScore.total > 0)) {
        fromTopCols[c2] = true;
      }
      if (subScore.event > 0 || subScore.client > 0 || subScore.total > 0) {
        fromSubCols[c2] = true;
      }

      combined.push({ event: eScore, client: clScore, total: tScore });
    }

    var mergedMatch = pickBestColumnsFromScores(combined);
    if (mergedMatch) {
      var colsUsed = [mergedMatch.eventCol, mergedMatch.clientCol, mergedMatch.totalCol];
      var usedTop = false;
      var usedSub = false;
      for (var cu = 0; cu < colsUsed.length; cu++) {
        if (fromTopCols[colsUsed[cu]]) usedTop = true;
        if (fromSubCols[colsUsed[cu]]) usedSub = true;
      }
      if (usedTop && usedSub && (!bestMerged || mergedMatch.scoreSum > bestMerged.scoreSum)) {
        bestMerged = {
          headerRowIndex: r2 + 1,
          eventCol: mergedMatch.eventCol,
          clientCol: mergedMatch.clientCol,
          totalCol: mergedMatch.totalCol,
          scoreSum: mergedMatch.scoreSum
        };
      }
    }
  }

  return bestMerged;
}

function handleSyncOverviewSheet(payload) {
  var sheetId = extractSpreadsheetId(payload && payload.spreadsheetId ? payload.spreadsheetId : '');
  if (!sheetId) {
    return {
      success: false,
      scriptVersion: APP_SCRIPT_VERSION,
      error: 'Invalid or missing Google Spreadsheet link/ID. Add your spreadsheet URL in Settings.'
    };
  }

  var ss;
  try {
    ss = SpreadsheetApp.openById(sheetId);
  } catch (openErr) {
    return {
      success: false,
      scriptVersion: APP_SCRIPT_VERSION,
      error: 'Could not open spreadsheet. Check the URL/ID and ensure treas@bts-crew.com has view access.'
    };
  }

  var overviewSheet = ss.getSheetByName('Overview');
  if (!overviewSheet) {
    var allSheets = ss.getSheets();
    for (var s = 0; s < allSheets.length; s++) {
      if (String(allSheets[s].getName() || '').trim().toLowerCase() === 'overview') {
        overviewSheet = allSheets[s];
        break;
      }
    }
  }

  if (!overviewSheet) {
    return {
      success: false,
      scriptVersion: APP_SCRIPT_VERSION,
      error: 'Could not find a sheet tab named "Overview" in this spreadsheet.'
    };
  }

  var values = overviewSheet.getDataRange().getValues();
  if (!values || values.length === 0) {
    return {
      success: true,
      scriptVersion: APP_SCRIPT_VERSION,
      sheetTitle: ss.getName(),
      tabName: overviewSheet.getName(),
      rows: []
    };
  }

  var cols = findOverviewColumnIndices(values);
  if (!cols) {
    var diag = buildOverviewHeaderDiagnostics(values);
    var detailParts = [];
    if (diag.bestPartial) {
      var missing = [];
      if (!diag.bestPartial.found.client) missing.push('"Client"');
      if (!diag.bestPartial.found.eventName) missing.push('"Event Name"');
      if (!diag.bestPartial.found.total) missing.push('"Total"');
      if (missing.length > 0 && missing.length < 3) {
        detailParts.push('Missing ' + missing.join(' & ') + ' near row ' + diag.bestPartial.rowNumber);
      }
      if (diag.bestPartial.cells.length > 0) {
        detailParts.push('Row ' + diag.bestPartial.rowNumber + ' headers seen: [' + diag.bestPartial.cells.slice(0, 8).join(' | ') + ']');
      }
    } else if (diag.previewRows.length > 0) {
      var firstRow = diag.previewRows[0];
      detailParts.push('Row ' + firstRow.rowNumber + ' cells seen: [' + firstRow.cells.slice(0, 8).join(' | ') + ']');
    }
    return {
      success: false,
      scriptVersion: APP_SCRIPT_VERSION,
      debugHeaders: diag,
      error: 'Could not locate "Client", "Event Name", and "Total" columns in "' + overviewSheet.getName() + '".' + (detailParts.length > 0 ? ' ' + detailParts.join(' — ') : '')
    };
  }

  var rows = [];
  for (var r = cols.headerRowIndex + 1; r < values.length; r++) {
    var row = values[r];
    var eventName = String(row[cols.eventCol] === null || row[cols.eventCol] === undefined ? '' : row[cols.eventCol]).trim();
    if (!eventName) continue;

    var client = String(row[cols.clientCol] === null || row[cols.clientCol] === undefined ? '' : row[cols.clientCol]).trim();
    var rawTotal = row[cols.totalCol];
    rows.push({
      rowNumber: r + 1,
      eventName: eventName,
      client: client,
      total: (rawTotal === null || rawTotal === undefined) ? '' : rawTotal
    });
  }

  return {
    success: true,
    scriptVersion: APP_SCRIPT_VERSION,
    sheetTitle: ss.getName(),
    tabName: overviewSheet.getName(),
    rows: rows
  };
}

function extractFolderId(input) {
  var str = String(input || '').trim();
  if (!str) return '';
  var folderMatch = str.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (folderMatch) return folderMatch[1];
  var idMatch = str.match(/[-\w]{25,}/);
  return idMatch ? idMatch[0] : str;
}

function getGmailSignatureHtml() {
  try {
    if (typeof Gmail !== 'undefined' && Gmail.Users && Gmail.Users.Settings) {
      var res = Gmail.Users.Settings.SendAs.list('me');
      if (res && res.sendAs) {
        for (var i = 0; i < res.sendAs.length; i++) {
          if (res.sendAs[i].isDefault && res.sendAs[i].signature) {
            return res.sendAs[i].signature;
          }
        }
      }
    }
  } catch (e) {
    Logger.log('Gmail service signature lookup skipped: ' + e);
  }
  return '';
}

function createGmailDraft(subject, plainBody, htmlBody, pdfBlob) {
  if (typeof Gmail !== 'undefined' && Gmail.Users && Gmail.Users.Drafts) {
    try {
      var boundary = 'treastrack_boundary_' + Utilities.getUuid();
      var base64Pdf = Utilities.base64Encode(pdfBlob.getBytes());
      var safeFileName = pdfBlob.getName().replace(/"/g, '');

      var mimeLines = [
        'MIME-Version: 1.0',
        'Subject: =?UTF-8?B?' + Utilities.base64Encode(subject, Utilities.Charset.UTF_8) + '?=',
        'Content-Type: multipart/mixed; boundary="' + boundary + '"',
        '',
        '--' + boundary,
        'Content-Type: text/html; charset="UTF-8"',
        'Content-Transfer-Encoding: base64',
        '',
        Utilities.base64Encode(htmlBody, Utilities.Charset.UTF_8),
        '',
        '--' + boundary,
        'Content-Type: application/pdf; name="' + safeFileName + '"',
        'Content-Disposition: attachment; filename="' + safeFileName + '"',
        'Content-Transfer-Encoding: base64',
        '',
        base64Pdf,
        '--' + boundary + '--'
      ];

      var rawMime = mimeLines.join('\r\n');
      var encodedRaw = Utilities.base64EncodeWebSafe(rawMime, Utilities.Charset.UTF_8);
      Gmail.Users.Drafts.create({ message: { raw: encodedRaw } }, 'me');
      return;
    } catch (mimeErr) {
      Logger.log('Advanced Gmail draft failed, falling back to GmailApp: ' + mimeErr);
    }
  }

  try {
    GmailApp.createDraft('', subject, plainBody, {
      htmlBody: htmlBody,
      attachments: [pdfBlob]
    });
  } catch (emptyRecipErr) {
    var fallbackRecipient = Session.getActiveUser().getEmail() || 'recipient@example.com';
    GmailApp.createDraft(fallbackRecipient, subject, plainBody, {
      htmlBody: htmlBody,
      attachments: [pdfBlob]
    });
  }
}

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
