/**
 * BTS Treas Tools — Google Drive & Gmail Automation Script (v2)
 *
 * SETUP INSTRUCTIONS:
 * 1. Go to https://script.google.com and open your project.
 * 2. Replace all code in Code.gs with this entire file and press Ctrl+S (Save).
 * 3. (Recommended — for Automatic Gmail Signature, Blank "To" field & Fast Draft Search):
 *    In the left sidebar, click "+" next to "Services", select "Gmail API", and click "Add".
 * 4. At the top toolbar, select "doGet" in the function dropdown and click "Run" once
 *    to approve the Google permissions prompt ("Review permissions" -> "Advanced" -> "Allow").
 *    (This approves Drive, Gmail, and External Token Verification scopes.)
 * 5. Click "Deploy" (top right) -> "Manage deployments" -> Edit -> Version: "New version" -> "Deploy".
 *    - Type: "Web app"
 *    - Execute as: "Me"
 *    - Who has access: "Anyone" (Requests still require a verified treas@bts-crew.com Google OAuth token)
 */

var APP_SCRIPT_VERSION = 2;
var AUTHORIZED_EMAIL = 'treas@bts-crew.com';
var DEFAULT_ALLOWED_CLIENT_IDS = [
  '523917169511-qp7hvnb54uhho8huevg0r6e76bil27ub.apps.googleusercontent.com'
];

function doGet(e) {
  try {
    // Touch Drive, Gmail, and UrlFetch scopes so running doGet once approves all permissions
    DriveApp.getRootFolder().getName();
    GmailApp.getInboxThreads(0, 1);
    UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo', { muteHttpExceptions: true });

    var hasAdvancedGmail = (typeof Gmail !== 'undefined' && Boolean(Gmail.Users));
    var statusText = [
      '✓ BTS Treas Tools Google Apps Script (v' + APP_SCRIPT_VERSION + ') is active and authorized!',
      '',
      '• Google Drive scope: Connected',
      '• Gmail scope: Connected',
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
