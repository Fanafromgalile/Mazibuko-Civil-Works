function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Mazibuko Civil Works (MCW) - Making things Possible')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    let data = {};

    if (e && e.postData && e.postData.contents) {
      const raw = e.postData.contents;
      if (typeof raw === 'string' && raw.trim()) {
        data = JSON.parse(raw);
      }
    }

    if (e && e.parameter && Object.keys(e.parameter).length) {
      data = { ...data, ...e.parameter };
    }

    const result = processFormSubmission(data || {});
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    const result = {
      status: 'error',
      message: error && error.toString ? error.toString() : 'Unexpected server error.'
    };

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handles form submissions securely from the frontend via google.script.run
 */
function processFormSubmission(data) {
  const lock = LockService.getScriptLock();

  try {
    if (!lock.tryLock(LOCK_WAIT_TIMEOUT)) {
      return { status: 'error', message: 'Server busy. Please try again shortly.' };
    }

    data = normalizeSubmissionData(data || {});

    if (data.website_hp) {
      return { status: 'error', message: 'Submission flagged as spam.' };
    }

    const clientEmail = getClientEmail(data);
    if (!clientEmail) {
      return { status: 'error', message: 'Please provide a valid email address.' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const timestamp = new Date();
    const submissionType = String(data.type || 'quote').toLowerCase();

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
        'New'
      ]);

      const pdfBlob = generateQuotePdf(data);
      const clientName = escapeHtml(data.fullName || 'Valued Client');

      const clientEmailBody = getProfessionalEmailTemplate(
        'Official Quotation Request Confirmation',
        `Dear <strong>${clientName}</strong>,`,
        `Thank you for contacting <strong>Mazibuko Civil Works (MCW)</strong>.<br><br>
        We acknowledge receipt of your project inquiry. Our technical team is currently reviewing your specifications to ensure we provide you with an accurate and comprehensive proposal.<br><br>
        For your convenience, we have attached an itemized summary PDF of your quote request to this email for your reference. We will follow up with you shortly to discuss next steps.`,
        'Require Immediate Assistance?',
        'You may connect directly with our operations team via WhatsApp at <strong>+27 69 877 7757</strong> or reply directly to this communication.',
        '#487138'
      );

      sendMcwEmail(clientEmail, 'MCW Quote Request Confirmation & Reference Summary', clientEmailBody, [pdfBlob]);
      sendAdminNotifications('quote', data, clientEmail, fileUrl);

      return { status: 'success', message: 'Quote recorded and emails dispatched successfully.' };

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
        'New'
      ]);

      const clientName = escapeHtml(data.fullName || 'Valued Client');
      const topicSubject = data.topic ? ` regarding <strong>${escapeHtml(data.topic)}</strong>` : '';

      const clientEmailBody = getProfessionalEmailTemplate(
        'Inquiry Received - MCW Support',
        `Dear <strong>${clientName}</strong>,`,
        `Thank you for reaching out to <strong>Mazibuko Civil Works (MCW)</strong>${topicSubject}.<br><br>
        We have successfully received your inquiry and routed it to the appropriate department. A specialist will review your question and respond with detailed information shortly.`,
        'Direct Support Channel',
        'For urgent inquiries, please feel free to reach our team on WhatsApp at <strong>+27 69 877 7757</strong>.',
        '#B5882D'
      );

      sendMcwEmail(clientEmail, 'MCW Inquiry Acknowledgment', clientEmailBody);
      sendAdminNotifications('faq', data, clientEmail, fileUrl);

      return { status: 'success', message: 'Inquiry recorded and emails dispatched successfully.' };

    } else {
      return { status: 'error', message: 'Invalid payload type.' };
    }

  } catch (error) {
    Logger.log('FATAL ERROR in processFormSubmission: ' + error.toString());
    return { status: 'error', message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function normalizeSubmissionData(data) {
  const raw = data || {};
  const contactInfo = String(raw.contactInfo || '').trim();
  const email = String(raw.email || '').trim();
  const fullName = String(raw.fullName || '').trim();

  return {
    ...raw,
    type: String(raw.type || 'quote').toLowerCase(),
    email: email || (isValidEmail(contactInfo) ? contactInfo : ''),
    contactInfo,
    fullName,
    topic: String(raw.topic || '').trim(),
    question: String(raw.question || raw.details || '').trim(),
    details: String(raw.details || '').trim(),
    service: String(raw.service || '').trim(),
    location: String(raw.location || '').trim()
  };
}

function isValidEmail(value) {
  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailPattern.test(String(value || '').trim());
}

/**
 * Robust Email Dispatcher
 */
function sendMcwEmail(recipient, subject, htmlBody, attachments = []) {
  const options = {
    name: 'Mazibuko Civil Works',
    htmlBody: htmlBody,
    attachments: attachments
  };

  try {
    GmailApp.sendEmail(recipient, subject, '', options);
    Logger.log('Successfully sent email to: ' + recipient);
    return true;
  } catch (err1) {
    Logger.log('GmailApp failed for ' + recipient + ': ' + err1.toString() + '. Attempting MailApp fallback...');
    try {
      MailApp.sendEmail(recipient, subject, '', options);
      return true;
    } catch (err2) {
      Logger.log('CRITICAL: Email dispatch failed for ' + recipient + ': ' + err2.toString());
      return false;
    }
  }
}

/**
 * Sends notifications to all Admin Emails (Outlook & info@mazibukocivilworks.co.za)
 */
function sendAdminNotifications(type, data, clientEmail, fileUrl) {
  const subject = `🔔 New MCW ${type.toUpperCase()} Submission: ${data.fullName || 'Client'}`;
  let detailsHtml = '';

  if (type === 'quote') {
    detailsHtml = `
      <p><strong>Service Requested:</strong> ${escapeHtml(data.service || 'N/A')}</p>
      <p><strong>Project Location:</strong> ${escapeHtml(data.location || 'N/A')}</p>
      <p><strong>Project Details:</strong><br>${escapeHtml(data.details || 'N/A').replace(/\n/g, '<br>')}</p>
      <p><strong>File Attachment:</strong> <a href="${fileUrl}" target="_blank">View in Google Drive</a></p>
    `;
  } else {
    detailsHtml = `
      <p><strong>Question Topic:</strong> ${escapeHtml(data.topic || 'N/A')}</p>
      <p><strong>Question Text:</strong><br>${escapeHtml(data.question || 'N/A').replace(/\n/g, '<br>')}</p>
    `;
  }

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; padding: 20px; color: #333;">
      <h2 style="color: #487138;">New Website ${type.toUpperCase()} Submission</h2>
      <hr style="border: 1px solid #ddd;">
      <p><strong>Full Name:</strong> ${escapeHtml(data.fullName || 'N/A')}</p>
      <p><strong>Email Address:</strong> ${escapeHtml(clientEmail)}</p>
      <p><strong>Contact Info / Phone:</strong> ${escapeHtml(data.contactInfo || 'N/A')}</p>
      ${detailsHtml}
      <hr style="border: 1px solid #ddd;">
      <p style="font-size: 12px; color: #777;">Automated notification from Mazibuko Civil Works website backend.</p>
    </div>
  `;

  ADMIN_EMAILS.forEach(adminAddr => {
    if (adminAddr) {
      sendMcwEmail(adminAddr, subject, htmlBody);
    }
  });
}

/**
 * Saves base64 encoded files to Google Drive
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
    Logger.log('File upload failed: ' + e.toString());
    return 'Upload Failed';
  }
}

/**
 * Auto-generates a PDF Quotation Blob
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
        <p style="font-size: 12px; color: #777; text-align: center;">Mazibuko Civil Works (Pty) Ltd | Contact: +27 69 877 7757 | Email: info@mazibukocivilworks.co.za</p>
      </body>
    </html>
  `;
  const blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf');
  blob.setName(`MCW_Quote_${(data.fullName || 'Client').replace(/\s+/g, '_')}.pdf`);
  return blob;
}

/**
 * Professional Corporate Client Email Template Builder
 */
function getProfessionalEmailTemplate(title, salutation, bodyText, alertTitle, alertText, alertBorderColor) {
  return `
    <div style="background-color: #f4f6f3; padding: 30px; font-family: 'Inter', Arial, sans-serif;">
      <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden; border: 1px solid #dcdcdc; box-shadow: 0 2px 8px rgba(0,0,0,0.04);">
        <div style="background-color: #487138; padding: 25px 20px; text-align: center; border-bottom: 4px solid #B5882D;">
          <h1 style="color: #ffffff; margin: 0; font-size: 22px; font-weight: 700; letter-spacing: 0.5px;">MAZIBUKO CIVIL WORKS</h1>
          <p style="color: #e0e8dc; margin: 6px 0 0 0; font-size: 12px; font-style: italic; letter-spacing: 0.3px;">Making things Possible</p>
        </div>

        <div style="padding: 35px 30px; color: #333333; line-height: 1.6; font-size: 14px;">
          <h2 style="color: #487138; margin-top: 0; font-size: 18px; border-bottom: 1px solid #eee; padding-bottom: 10px;">${title}</h2>
          <p style="font-size: 15px; font-weight: 600; color: #1c2b16; margin-top: 20px;">${salutation}</p>
          <p>${bodyText}</p>

          <div style="background-color: #f8faf7; border-left: 4px solid ${alertBorderColor}; padding: 15px 18px; margin: 25px 0; border-radius: 0 6px 6px 0;">
            <p style="margin: 0; font-size: 13px; color: ${alertBorderColor}; font-weight: bold; text-transform: uppercase; letter-spacing: 0.3px;">${alertTitle}</p>
            <p style="margin: 5px 0 0 0; font-size: 13px; color: #444;">${alertText}</p>
          </div>

          <p style="margin-top: 30px;">Kind regards,</p>
          <p style="margin: 0; font-weight: bold; color: #487138; font-size: 15px;">Management &amp; Operations Team</p>
          <p style="margin: 2px 0 0 0; font-size: 13px; color: #555;">Mazibuko Civil Works (Pty) Ltd</p>
        </div>

        <div style="background-color: #f4f6f3; padding: 20px; text-align: center; border-top: 1px solid #e2e8df; font-size: 11px; color: #555;">
          <p style="margin: 0 0 4px 0; font-weight: bold; color: #487138;">Mazibuko Civil Works (Pty) Ltd</p>
          <p style="margin: 0 0 4px 0;">Telephone / WhatsApp: +27 69 877 7757 | Email: info@mazibukocivilworks.co.za</p>
          <p style="margin: 0; color: #777; font-style: italic;">Professional Property Maintenance &amp; Civil Site Support</p>
        </div>
      </div>
    </div>
  `;
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
