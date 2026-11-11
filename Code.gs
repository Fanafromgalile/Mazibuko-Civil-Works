/**
 * Mazibuko Civil Works (MCW) - Comprehensive Backend Script
 * Target Spreadsheet ID: 1b2ldt35HijbAdsroXUCcZoGLWDuctfwOugmCwMd2QFo
 */

const SPREADSHEET_ID = "1b2ldt35HijbAdsroXUCcZoGLWDuctfwOugmCwMd2QFo";
const ADMIN_EMAIL = "mcw.services2024@outlook.com";
const DRIVE_FOLDER_ID = ""; // Optional: Paste your Google Drive Folder ID here for file uploads
const WHATSAPP_API_URL = ""; // Optional: CallMeBot or Twilio webhook URL for instant WhatsApp alerts
const LOCK_WAIT_TIMEOUT = 10000; // 10 seconds

/**
 * Serves your HTML frontend file directly when opening the Web App URL
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Mazibuko Civil Works (MCW) - Making things Possible')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Handles form submissions securely from the frontend via google.script.run or doPost
 */
function processFormSubmission(data) {
  const lock = LockService.getScriptLock();
  
  try {
    if (!lock.tryLock(LOCK_WAIT_TIMEOUT)) {
      return { status: 'error', message: 'Server busy. Please try again shortly.' };
    }

    data = data || {};

    // 1. Anti-Spam: Check Honeypot field (bots fill hidden fields)
    if (data.website_hp) {
      return { status: 'error', message: 'Submission flagged as spam.' };
    }

    // 2. Anti-Spam: Rate Limiting (Max 3 submissions per IP/Email every 5 minutes)
    const clientEmail = getClientEmail(data);
    if (!clientEmail) {
      return { status: 'error', message: 'Please provide a valid email address.' };
    }

    const cache = CacheService.getScriptCache();
    const rateKey = 'rate_' + clientEmail.replace(/[^a-z0-9]/g, '_');
    const recentSubmissions = parseInt(cache.get(rateKey) || '0', 10);
    if (recentSubmissions >= 3) {
      return { status: 'error', message: 'Too many submissions. Please try again later.' };
    }
    cache.put(rateKey, (recentSubmissions + 1).toString(), 300); // 5 minute window

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const timestamp = new Date();
    const submissionType = data.type ? data.type : 'quote';

    // 3. Handle File Uploads to Google Drive if provided
    let fileUrl = 'None';
    if (data.file && data.file.data) {
      fileUrl = saveFileToDrive(data.file, data.fullName || clientEmail);
    }

    if (submissionType === 'quote') {
      let sheet = ss.getSheetByName('Quotes') || ss.insertSheet('Quotes');

      if (sheet.getLastRow() === 0 || sheet.getRange(1, 2).getValue() === 'Untitled Question') {
        sheet.clear();
        sheet.appendRow(['Timestamp', 'Full Name', 'Contact Info', 'Service Requested', 'Project Location', 'Project Details', 'Email', 'File Attachment', 'Status']);
        sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
      }
      ensureHeader(sheet, 7, 'Email');
      ensureHeader(sheet, 8, 'File Attachment');
      ensureHeader(sheet, 9, 'Status');

      sheet.appendRow([
        timestamp,
        data.fullName || '',
        data.contactInfo || '',
        data.service || '',
        data.location || '',
        data.details || '',
        clientEmail,
        fileUrl,
        'New' // Lead Status Pipeline
      ]);

      // Generate PDF Quote Attachment
      const pdfBlob = generateQuotePdf(data);

      // Send Confirmation Email with PDF attachment to Client
      const clientName = escapeHtml(data.fullName || "Valued Client");
      const clientEmailBody = getEmailTemplate(
        "Quote Request Received & Official Estimate",
        `Dear <strong>${clientName}</strong>,`,
        `Thank you for contacting <strong>Mazibuko Civil Works (MCW)</strong>.<br><br>We have received your request and attached your preliminary quotation details for your review. We will be in touch shortly.`,
        `Need urgent assistance?`,
        `Connect directly with us on WhatsApp at <strong>+27 69 877 7757</strong>.`,
        `#487138`
      );
      sendMcwEmail(clientEmail, "MCW Quote Request Confirmation & Estimate", clientEmailBody, [pdfBlob]);

      // Notify Admin (Email + WhatsApp)
      sendAdminNotification('quote', data, clientEmail, fileUrl);
      sendWhatsAppAlert(`New Quote from ${data.fullName || 'Client'}: ${data.service || 'General'}`);

      return { status: 'success', message: 'Quote recorded, PDF estimate emailed, and team notified.' };

    } else if (submissionType === 'faq') {
      let sheet = ss.getSheetByName('FAQ') || ss.insertSheet('FAQ');

      if (sheet.getLastRow() === 0) {
        sheet.appendRow(['Timestamp', 'Contact Info', 'Question Topic', 'Question Text', 'Email', 'Full Name', 'Status']);
        sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
      }
      ensureHeader(sheet, 5, 'Email');
      ensureHeader(sheet, 6, 'Full Name');
      ensureHeader(sheet, 7, 'Status');

      sheet.appendRow([
        timestamp,
        data.contactInfo || '',
        data.topic || '',
        data.question || '',
        clientEmail,
        data.fullName || '',
        'New' // Lead Status Pipeline
      ]);

      // Send Confirmation to Client
      const topicSubject = data.topic ? ` regarding <strong>${escapeHtml(data.topic)}</strong>` : '';
      const clientEmailBody = getEmailTemplate(
        "Inquiry Received",
        `Dear <strong>${escapeHtml(data.fullName || "Valued Client")}</strong>,`,
        `Thank you for reaching out to <strong>Mazibuko Civil Works (MCW)</strong> with your inquiry${topicSubject}.<br><br>Our team is reviewing your question and will reply shortly.`,
        `Quick Assistance`,
        `Feel free to reach us on WhatsApp at <strong>+27 69 877 7757</strong>.`,
        `#B5882D`
      );
      sendMcwEmail(clientEmail, "MCW Inquiry Confirmation", clientEmailBody);

      // Notify Admin (Email + WhatsApp)
      sendAdminNotification('faq', data, clientEmail, fileUrl);
      sendWhatsAppAlert(`New Inquiry from ${data.fullName || 'Client'}: ${data.topic || 'General'}`);

      return { status: 'success', message: 'Inquiry recorded, confirmation sent, and team notified.' };

    } else {
      return { status: 'error', message: 'Invalid payload type.' };
    }

  } catch (error) {
    return { status: 'error', message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Feature 2: Saves base64 encoded client files to Google Drive
 */
function saveFileToDrive(fileObj, clientName) {
  try {
    const folder = DRIVE_FOLDER_ID ? DriveApp.getFolderById(DRIVE_FOLDER_ID) : DriveApp.getRootFolder();
    const splitData = fileObj.data.split(',');
    const contentType = splitData[0].match(/:(.*?);/)[1];
    const decodedBytes = Utilities.base64Decode(splitData[1]);
    const blob = Utilities.newBlob(decodedBytes, contentType, `${clientName}_${fileObj.name || 'attachment'}`);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    Logger.log("File upload failed: " + e.toString());
    return 'Upload Failed';
  }
}

/**
 * Feature 3: Auto-generates a PDF Quotation Blob
 */
function generateQuotePdf(data) {
  const htmlContent = `
    <html>
      <body style="font-family: Arial, sans-serif; padding: 30px; color: #333;">
        <h1 style="color: #487138; border-bottom: 2px solid #B5882D; padding-bottom: 10px;">MAZIBUKO CIVIL WORKS (MCW)</h1>
        <h3>Official Project Estimate / Quote Summary</h3>
        <p><strong>Date:</strong> ${new Date().toLocaleDateString()}</p>
        <hr>
        <p><strong>Client Name:</strong> ${escapeHtml(data.fullName || 'N/A')}</p>
        <p><strong>Contact / Phone:</strong> ${escapeHtml(data.contactInfo || 'N/A')}</p>
        <p><strong>Service Requested:</strong> ${escapeHtml(data.service || 'N/A')}</p>
        <p><strong>Project Location:</strong> ${escapeHtml(data.location || 'N/A')}</p>
        <br>
        <h4>Project Details:</h4>
        <p style="background: #f4f6f3; padding: 15px; border-left: 4px solid #487138;">${escapeHtml(data.details || 'N/A').replace(/\n/g, '<br>')}</p>
        <br><br>
        <p style="font-size: 12px; color: #777; text-align: center;">Mazibuko Civil Works (MCW) | Phone / WhatsApp: +27 69 877 7757 | Email: mcw.services2024@outlook.com</p>
      </body>
    </html>
  `;
  const blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf');
  blob.setName(`MCW_Quote_${(data.fullName || 'Client').replace(/\s+/g, '_')}.pdf`);
  return blob;
}

/**
 * Resilient Helper: Sends branded MCW emails via GmailApp with fallback and attachments
 */
function sendMcwEmail(recipient, subject, htmlBody, attachments = []) {
  if (MailApp.getRemainingDailyQuota() <= 0) return false;

  const options = {
    name: "Mazibuko Civil Works",
    htmlBody: htmlBody,
    attachments: attachments
  };

  try {
    GmailApp.sendEmail(recipient, subject, "", options);
    return true;
  } catch (err1) {
    Logger.log("GmailApp failed: " + err1.toString());
    try {
      MailApp.sendEmail(recipient, subject, "", options);
      return true;
    } catch (err2) {
      return false;
    }
  }
}

/**
 * Feature 5: Sends instant WhatsApp alerts via webhook (e.g. CallMeBot / Twilio)
 */
function sendWhatsAppAlert(message) {
  if (!WHATSAPP_API_URL) return;
  try {
    UrlFetchApp.fetch(`${WHATSAPP_API_URL}&text=${encodeURIComponent(message)}`, {
      method: 'get',
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log("WhatsApp alert failed: " + e.toString());
  }
}

/**
 * Sends Admin Notification Email
 */
function sendAdminNotification(type, data, clientEmail, fileUrl) {
  if (MailApp.getRemainingDailyQuota() <= 0) return;

  const subject = `🔔 New MCW ${type.toUpperCase()} Submission: ${data.fullName || 'Client'}`;
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; padding: 20px; color: #333;">
      <h2 style="color: #487138;">New Website ${type.toUpperCase()} Submission</h2>
      <hr>
      <p><strong>Full Name:</strong> ${escapeHtml(data.fullName || 'N/A')}</p>
      <p><strong>Email:</strong> ${escapeHtml(clientEmail)}</p>
      <p><strong>Phone:</strong> ${escapeHtml(data.contactInfo || 'N/A')}</p>
      <p><strong>Service / Topic:</strong> ${escapeHtml(data.service || data.topic || 'N/A')}</p>
      <p><strong>Location:</strong> ${escapeHtml(data.location || 'N/A')}</p>
      <p><strong>Details:</strong><br>${escapeHtml(data.details || data.question || 'N/A').replace(/\n/g, '<br>')}</p>
      <p><strong>File Attachment:</strong> <a href="${fileUrl}" target="_blank">View in Google Drive</a></p>
    </div>
  `;

  try {
    GmailApp.sendEmail(ADMIN_EMAIL, subject, "", {
      name: "MCW Notification System",
      htmlBody: htmlBody
    });
  } catch (e) {}
}

/**
 * Helper: Reusable HTML Email Template Builder
 */
function getEmailTemplate(title, salutation, bodyText, alertTitle, alertText, alertBorderColor) {
  return `
    <div style="background-color: #f4f6f3; padding: 20px; font-family: Arial, sans-serif;">
      <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden; border: 1px solid #e0e0e0;">
        <div style="background-color: #487138; padding: 25px 20px; text-align: center; border-bottom: 4px solid #B5882D;">
          <h2 style="color: #ffffff; margin: 0;">MAZIBUKO CIVIL WORKS</h2>
          <p style="color: #e0e8dc; margin: 5px 0 0 0; font-size: 12px; font-style: italic;">Making things Possible</p>
        </div>
        <div style="padding: 30px; color: #333333; line-height: 1.6;">
          <h3 style="color: #487138; margin-top: 0;">${title}</h3>
          <p>${salutation}</p>
          <p>${bodyText}</p>
          <div style="background-color: #f4f7f2; border-left: 4px solid ${alertBorderColor}; padding: 15px; margin: 20px 0;">
            <p style="margin: 0; font-size: 13px; color: ${alertBorderColor}; font-weight: bold;">${alertTitle}</p>
            <p style="margin: 5px 0 0 0; font-size: 13px;">${alertText}</p>
          </div>
          <p>Kind regards,</p>
          <p style="margin: 0; font-weight: bold; color: #487138;">Team MCW</p>
        </div>
      </div>
    </div>
  `;
}

function doPost(e) {
  let data = {};
  try {
    data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return jsonResponse({ status: 'error', message: 'Invalid JSON payload.' });
  }
  return jsonResponse(processFormSubmission(data));
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getClientEmail(data) {
  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const candidates = [data.email, data.contactInfo];
  for (let i = 0; i < candidates.length; i++) {
    const value = String(candidates[i] || '').trim();
    if (emailPattern.test(value)) return value;
  }
  return '';
}

function ensureHeader(sheet, column, title) {
  const cell = sheet.getRange(1, column);
  if (!cell.getValue()) {
    cell.setValue(title).setFontWeight('bold');
  }
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
