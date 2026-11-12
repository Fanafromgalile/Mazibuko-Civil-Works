// ==========================================
// 1. CONFIGURATION VARIABLES
// ==========================================
// Tip: You can paste your IDs directly below OR set them in Project Settings > Script Properties.
const DEFAULT_SPREADSHEET_ID = '1b2ldt35HijbAdsroXUCcZoGLWDuctfwOugmCwMd2QFo';
const DEFAULT_DRIVE_FOLDER_ID = ''; // Optional: Paste a Google Drive Folder ID for site attachments
const ADMIN_EMAILS = ['mcw.services2024@outlook.com', 'info@mazibukocivilworks.co.za'];
const LOCK_WAIT_TIMEOUT = 10000; // 10 seconds
const BASE_LOCATION = 'Soweto, Gauteng, South Africa';

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    spreadsheetId: props.getProperty('SPREADSHEET_ID') || DEFAULT_SPREADSHEET_ID,
    driveFolderId: props.getProperty('DRIVE_FOLDER_ID') || DEFAULT_DRIVE_FOLDER_ID
  };
}

// ==========================================
// 2. WEB APP ENTRY POINTS (doGet / doPost)
// ==========================================

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
    Logger.log('doPost Error: ' + error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: 'Unexpected server error while processing your request.'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 3. CORE FORM SUBMISSION HANDLER
// ==========================================

function processFormSubmission(data) {
  const lock = LockService.getScriptLock();

  try {
    if (!lock.tryLock(LOCK_WAIT_TIMEOUT)) {
      return { status: 'error', message: 'Server is currently busy. Please try again in a few seconds.' };
    }

    const config = getConfig();
    if (!config.spreadsheetId || config.spreadsheetId === 'YOUR_SPREADSHEET_ID_HERE') {
      return { status: 'error', message: 'Server configuration incomplete: Spreadsheet ID is missing.' };
    }

    data = normalizeSubmissionData(data || {});

    // 1. Server-side Honeypot Spam Check
    if (data.website_hp) {
      Logger.log('Blocked spam submission via honeypot.');
      return { status: 'error', message: 'Submission flagged as spam.' };
    }

    // 2. Validate Required Client Contact Details
    const clientEmail = getClientEmail(data);
    if (!clientEmail) {
      return { status: 'error', message: 'Please provide a valid email address.' };
    }

    if (!data.fullName) {
      return { status: 'error', message: 'Please provide your full name.' };
    }

    const ss = SpreadsheetApp.openById(config.spreadsheetId);
    const timestamp = new Date();
    const refId = generateReferenceId();
    const submissionType = data.type;

    // 3. Handle Optional Site Photo / PDF Upload
    let fileUrl = 'None';
    if (data.file && data.file.data) {
      fileUrl = saveFileToDrive(data.file, `${refId}_${data.fullName}`, config.driveFolderId);
    }

    // 4. Route by Submission Type ('quote' vs 'faq')
    if (submissionType === 'quote') {
      if (!data.location) {
        return { status: 'error', message: 'Please enter your suburb or project location.' };
      }

      let sheet = ss.getSheetByName('Quotes') || ss.insertSheet('Quotes');

      // Run Secured Suburb Distance & Callout Fee Engine
      const routing = calculateDistanceAndFee(data.location);
      data.routing = routing;
      data.refId = refId;

      if (sheet.getLastRow() === 0) {
        sheet.appendRow([
          'Timestamp',
          'Reference ID',
          'Full Name',
          'Email',
          'Phone / WhatsApp',
          'Service Requested',
          'Project Location',
          'Distance to Site',
          'Callout Fee',
          'Project Details',
          'File Attachment',
          'Status'
        ]);
        sheet.getRange(1, 1, 1, 12).setFontWeight('bold');
        sheet.setFrozenRows(1);
      }

      const feeCellText = routing.status === 'OK'
        ? (routing.fee > 0 ? `R ${routing.fee}` : 'Free (<= 25 km)')
        : 'Pending Review';

      sheet.appendRow([
        timestamp,
        refId,
        data.fullName,
        clientEmail,
        data.contactInfo,
        data.service,
        data.location,
        routing.distanceText,
        feeCellText,
        data.details,
        fileUrl,
        'New'
      ]);

      const pdfBlob = generateQuotePdf(data);
      const clientName = escapeHtml(data.fullName);

      const clientEmailBody = getProfessionalEmailTemplate(
        `Quotation Request Confirmation (${refId})`,
        `Dear <strong>${clientName}</strong>,`,
        `Thank you for contacting <strong>Mazibuko Civil Works (MCW)</strong>.<br><br>
        We acknowledge receipt of your project inquiry (Reference: <strong>${refId}</strong>) for <strong>${escapeHtml(data.service)}</strong> in <strong>${escapeHtml(data.location)}</strong>. Our operations team is reviewing your specifications to prepare a formal quotation.<br><br>
        For your records, an itemized PDF summary—including your estimated site distance and callout details—is attached to this email.`,
        'Require Immediate Assistance?',
        `Quote Reference: <strong>${refId}</strong><br>Connect directly with our dispatch team on WhatsApp at <strong>+27 69 877 7757</strong> or reply to this email.`,
        '#487138'
      );

      sendMcwEmail(
        clientEmail,
        `MCW Quote Request Confirmation [${refId}]`,
        clientEmailBody,
        [pdfBlob]
      );
      sendAdminNotifications('quote', data, clientEmail, fileUrl, refId);

      return {
        status: 'success',
        refId: refId,
        message: `Thank you, ${data.fullName}! Your quote request (${refId}) has been submitted and a confirmation PDF has been sent to ${clientEmail}.`
      };

    } else if (submissionType === 'faq') {
      let sheet = ss.getSheetByName('FAQ') || ss.insertSheet('FAQ');
      data.refId = refId;

      if (sheet.getLastRow() === 0) {
        sheet.appendRow([
          'Timestamp',
          'Reference ID',
          'Full Name',
          'Email',
          'Phone / WhatsApp',
          'Inquiry Topic',
          'Question Text',
          'Status'
        ]);
        sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
        sheet.setFrozenRows(1);
      }

      sheet.appendRow([
        timestamp,
        refId,
        data.fullName,
        clientEmail,
        data.contactInfo,
        data.topic,
        data.question,
        'New'
      ]);

      const clientName = escapeHtml(data.fullName);
      const topicSubject = data.topic ? ` regarding <strong>${escapeHtml(data.topic)}</strong>` : '';

      const clientEmailBody = getProfessionalEmailTemplate(
        `Inquiry Received (${refId}) - MCW Support`,
        `Dear <strong>${clientName}</strong>,`,
        `Thank you for reaching out to <strong>Mazibuko Civil Works (MCW)</strong>${topicSubject}.<br><br>
        We have logged your inquiry under reference <strong>${refId}</strong> and routed it to our team. A specialist will review your message and respond shortly.`,
        'Direct Support Channel',
        'For urgent site assistance, contact our team directly on WhatsApp at <strong>+27 69 877 7757</strong>.',
        '#B5882D'
      );

      sendMcwEmail(
        clientEmail,
        `MCW Inquiry Acknowledgment [${refId}]`,
        clientEmailBody
      );
      sendAdminNotifications('faq', data, clientEmail, fileUrl, refId);

      return {
        status: 'success',
        refId: refId,
        message: `Thank you, ${data.fullName}! Your inquiry (${refId}) has been received and an acknowledgment has been sent to ${clientEmail}.`
      };

    } else {
      return { status: 'error', message: 'Invalid submission type.' };
    }

  } catch (error) {
    Logger.log('FATAL ERROR in processFormSubmission: ' + error.toString());
    return { status: 'error', message: 'Server error: ' + error.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 4. SECURED SUBURB DISTANCE & CALLOUT ENGINE
// ==========================================

function calculateDistanceAndFee(clientLocation) {
  let dest = String(clientLocation || '').trim();
  if (!dest) {
    return {
      distanceText: 'No Location Provided',
      distanceKm: 0,
      fee: 0,
      status: 'NO_LOCATION',
      feeMsg: 'Manual location check required'
    };
  }

  // Restrict geocoding within South Africa
  if (!dest.toLowerCase().includes('south africa') && !dest.toLowerCase().includes('sa')) {
    dest += ', South Africa';
  }

  try {
    const directions = Maps.newDirectionFinder()
      .setOrigin(BASE_LOCATION)
      .setDestination(dest)
      .setMode(Maps.DirectionFinder.Mode.DRIVING)
      .getDirections();

    if (directions.status !== 'OK' || !directions.routes || !directions.routes[0]) {
      return {
        distanceText: 'Unverified Suburb',
        distanceKm: 0,
        fee: 0,
        status: 'NOT_FOUND',
        feeMsg: 'Address could not be auto-verified; callout confirmed on review'
      };
    }

    const leg = directions.routes[0].legs[0];
    const distanceMeters = leg.distance.value;
    const distanceKm = parseFloat((distanceMeters / 1000).toFixed(1));
    const distanceText = leg.distance.text;

    let fee = 0;
    let feeMsg = 'Free Callout (Within 25 km of Soweto base)';

    if (distanceKm <= 25) {
      fee = 0;
    } else if (distanceKm <= 50) {
      fee = 150;
      feeMsg = 'Standard Gauteng Outlying Zone (26–50 km)';
    } else if (distanceKm <= 100) {
      fee = 350;
      feeMsg = 'Extended Gauteng Zone (51–100 km)';
    } else {
      fee = 850;
      feeMsg = 'Long Distance / Out-of-Province (>100 km — Subject to Operations Review)';
    }

    return {
      distanceText: distanceText,
      distanceKm: distanceKm,
      fee: fee,
      status: 'OK',
      feeMsg: feeMsg
    };
  } catch (e) {
    Logger.log('Maps API Error: ' + e.message);
    return {
      distanceText: 'Routing Unavailable',
      distanceKm: 0,
      fee: 0,
      status: 'ERROR',
      feeMsg: 'Callout fee will be confirmed on formal quotation'
    };
  }
}

// ==========================================
// 5. DATA NORMALIZATION & HELPERS
// ==========================================

function normalizeSubmissionData(data) {
  const raw = data || {};
  const contactInfo = String(raw.contactInfo || '').trim();
  const email = String(raw.email || '').trim();
  const fullName = String(raw.fullName || '').trim();
  const details = String(raw.details || '').trim();

  return {
    ...raw,
    type: String(raw.type || 'quote').toLowerCase(),
    website_hp: String(raw.website_hp || '').trim(),
    email: email || (isValidEmail(contactInfo) ? contactInfo : ''),
    contactInfo: contactInfo,
    fullName: fullName,
    topic: String(raw.topic || 'General Inquiry').trim(),
    question: String(raw.question || details).trim(),
    details: details,
    service: String(raw.service || 'General Service').trim(),
    location: String(raw.location || '').trim()
  };
}

function generateReferenceId() {
  const year = new Date().getFullYear();
  const randomCode = Math.floor(1000 + Math.random() * 9000);
  return `MCW-${year}-${randomCode}`;
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
}

function getClientEmail(data) {
  const candidates = [data.email, data.contactInfo];
  for (let i = 0; i < candidates.length; i++) {
    const value = String(candidates[i] || '').trim();
    if (isValidEmail(value)) return value;
  }
  return '';
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// ==========================================
// 6. GOOGLE DRIVE ATTACHMENT STORAGE
// ==========================================

function saveFileToDrive(fileObj, prefixName, driveFolderId) {
  try {
    if (!fileObj || !fileObj.data || !fileObj.data.includes(',')) {
      return 'None';
    }

    const folder = driveFolderId
      ? DriveApp.getFolderById(driveFolderId)
      : DriveApp.getRootFolder();

    const splitData = fileObj.data.split(',');
    const mimeMatch = splitData[0].match(/:(.*?);/);
    const contentType = mimeMatch ? mimeMatch[1] : 'application/octet-stream';
    const decodedBytes = Utilities.base64Decode(splitData[1]);

    const safePrefix = String(prefixName || 'Client').replace(/[^a-zA-Z0-9_-]/g, '_');
    const safeFileName = String(fileObj.name || 'attachment').replace(/[^a-zA-Z0-9._-]/g, '_');

    const blob = Utilities.newBlob(decodedBytes, contentType, `${safePrefix}_${safeFileName}`);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    Logger.log('File upload failed: ' + e.toString());
    return 'Upload Failed';
  }
}

// ==========================================
// 7. PDF & EMAIL DISPATCHERS
// ==========================================

function sendMcwEmail(recipient, subject, htmlBody, attachments = []) {
  const options = {
    name: 'Mazibuko Civil Works',
    htmlBody: htmlBody,
    attachments: attachments
  };

  try {
    GmailApp.sendEmail(recipient, subject, '', options);
    Logger.log('Sent email via GmailApp to: ' + recipient);
    return true;
  } catch (err1) {
    Logger.log('GmailApp failed for ' + recipient + ': ' + err1.toString());
    try {
      MailApp.sendEmail(recipient, subject, '', options);
      Logger.log('Sent email via MailApp fallback to: ' + recipient);
      return true;
    } catch (err2) {
      Logger.log('CRITICAL: Email failed for ' + recipient + ': ' + err2.toString());
      return false;
    }
  }
}

function sendAdminNotifications(type, data, clientEmail, fileUrl, refId) {
  const subject = `🔔 [${refId}] New MCW ${type.toUpperCase()}: ${data.fullName} (${data.location || 'Inquiry'})`;
  let detailsHtml = '';

  if (type === 'quote') {
    const isOutOfZone = data.routing && data.routing.distanceKm > 100;
    const zoneBannerColor = isOutOfZone ? '#dc2626' : '#B5882D';
    const zoneBgColor = isOutOfZone ? '#fef2f2' : '#fdf5e6';

    let calloutInfo = `<strong>Callout Status:</strong> ${escapeHtml(data.routing ? data.routing.feeMsg : 'Pending manual review')}`;
    if (data.routing && data.routing.status === 'OK') {
      const feeText = data.routing.fee === 0 ? 'Free (Within 25 km)' : `R ${data.routing.fee}`;
      calloutInfo = `
        <strong>Driving Distance from Base:</strong> ${escapeHtml(data.routing.distanceText)}<br>
        <strong>Calculated Callout Fee:</strong> ${feeText}<br>
        <strong>Zone Classification:</strong> ${escapeHtml(data.routing.feeMsg)}
      `;
    }

    const attachmentHtml = (fileUrl && fileUrl !== 'None' && fileUrl !== 'Upload Failed')
      ? `<a href="${fileUrl}" target="_blank" style="color: #487138; font-weight: bold;">Open Site Attachment in Google Drive</a>`
      : escapeHtml(fileUrl || 'None');

    detailsHtml = `
      <p><strong>Service Requested:</strong> ${escapeHtml(data.service)}</p>
      <p><strong>Project Location:</strong> ${escapeHtml(data.location)}</p>
      <div style="background: ${zoneBgColor}; padding: 12px 15px; border-left: 4px solid ${zoneBannerColor}; margin: 12px 0;">
        ${calloutInfo}
      </div>
      <p><strong>Project Details:</strong><br>${escapeHtml(data.details).replace(/\n/g, '<br>')}</p>
      <p><strong>Site Attachment:</strong> ${attachmentHtml}</p>
    `;
  } else {
    detailsHtml = `
      <p><strong>Inquiry Topic:</strong> ${escapeHtml(data.topic)}</p>
      <p><strong>Question Text:</strong><br>${escapeHtml(data.question).replace(/\n/g, '<br>')}</p>
    `;
  }

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; padding: 20px; color: #333;">
      <h2 style="color: #487138; margin-bottom: 4px;">New Website ${type.toUpperCase()} Submission</h2>
      <p style="margin-top: 0; font-size: 13px; color: #666;">Reference ID: <strong>${escapeHtml(refId)}</strong></p>
      <hr style="border: 1px solid #ddd;">
      <p><strong>Full Name:</strong> ${escapeHtml(data.fullName)}</p>
      <p><strong>Email Address:</strong> <a href="mailto:${escapeHtml(clientEmail)}">${escapeHtml(clientEmail)}</a></p>
      <p><strong>Phone / WhatsApp:</strong> ${escapeHtml(data.contactInfo)}</p>
      ${detailsHtml}
      <hr style="border: 1px solid #ddd;">
      <p style="font-size: 12px; color: #777;">Automated dispatch notification from Mazibuko Civil Works website backend.</p>
    </div>
  `;

  ADMIN_EMAILS.forEach(adminAddr => {
    if (adminAddr) {
      sendMcwEmail(adminAddr, subject, htmlBody);
    }
  });
}

function generateQuotePdf(data) {
  const refId = escapeHtml(data.refId || 'MCW-REF');
  let feeDisplayHtml = '';

  if (data.routing && data.routing.status === 'OK') {
    const feeText = data.routing.fee === 0 ? 'Free (Within 25 km standard radius)' : `R ${data.routing.fee}`;
    feeDisplayHtml = `
      <div style="background: #fdf5e6; padding: 12px 15px; margin-top: 15px; border-left: 4px solid #B5882D; font-size: 13px;">
        <strong>Est. Driving Distance from Soweto Base:</strong> ${escapeHtml(data.routing.distanceText)}<br>
        <strong>Transport / Callout Fee:</strong> ${feeText}<br>
        <span style="font-size: 11px; color: #666;">*${escapeHtml(data.routing.feeMsg)}</span>
      </div>
    `;
  } else if (data.routing) {
    feeDisplayHtml = `
      <div style="background: #f8faf7; padding: 12px 15px; margin-top: 15px; border-left: 4px solid #487138; font-size: 13px;">
        <strong>Transport / Callout Fee:</strong> To be confirmed upon address verification<br>
        <span style="font-size: 11px; color: #666;">*${escapeHtml(data.routing.feeMsg)}</span>
      </div>
    `;
  }

  const htmlContent = `
    <html>
      <body style="font-family: Arial, sans-serif; padding: 30px; color: #333;">
        <div style="border-bottom: 3px solid #B5882D; padding-bottom: 12px; margin-bottom: 20px;">
          <h1 style="color: #487138; margin: 0;">MAZIBUKO CIVIL WORKS (MCW)</h1>
          <p style="margin: 4px 0 0 0; font-size: 12px; color: #666; font-style: italic;">Making things Possible | Property Maintenance &amp; Civil Site Support</p>
        </div>

        <table style="width: 100%; border-collapse: collapse; margin-bottom: 15px; font-size: 13px;">
          <tr>
            <td><h3 style="margin: 0; color: #141e11;">Quotation Request Summary</h3></td>
            <td style="text-align: right;"><strong>Ref:</strong> ${refId}<br><strong>Date:</strong> ${new Date().toLocaleDateString('en-ZA')}</td>
          </tr>
        </table>

        <hr style="border: 0; border-top: 1px solid #ddd;">

        <p style="font-size: 13px;"><strong>Client Name:</strong> ${escapeHtml(data.fullName)}</p>
        <p style="font-size: 13px;"><strong>Email Address:</strong> ${escapeHtml(data.email)}</p>
        <p style="font-size: 13px;"><strong>Phone / WhatsApp:</strong> ${escapeHtml(data.contactInfo)}</p>
        <p style="font-size: 13px;"><strong>Service Requested:</strong> ${escapeHtml(data.service)}</p>
        <p style="font-size: 13px;"><strong>Project Location:</strong> ${escapeHtml(data.location)}</p>

        ${feeDisplayHtml}

        <h4 style="margin-top: 22px; margin-bottom: 8px; color: #487138;">Project Details &amp; Scope:</h4>
        <p style="background: #f4f6f3; padding: 15px; border-left: 4px solid #487138; font-size: 13px; line-height: 1.5;">
          ${escapeHtml(data.details).replace(/\n/g, '<br>')}
        </p>

        <div style="margin-top: 40px; padding-top: 15px; border-top: 1px solid #ddd; font-size: 11px; color: #777; text-align: center;">
          Mazibuko Civil Works (Pty) Ltd | Soweto, Gauteng<br>
          Telephone / WhatsApp: +27 69 877 7757 | Email: info@mazibukocivilworks.co.za
        </div>
      </body>
    </html>
  `;

  const safeClient = String(data.fullName || 'Client').replace(/[^a-zA-Z0-9_-]/g, '_');
  return HtmlService.createHtmlOutput(htmlContent)
    .getAs('application/pdf')
    .setName(`${refId}_Quote_${safeClient}.pdf`);
}

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
