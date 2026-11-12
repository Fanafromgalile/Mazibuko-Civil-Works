// ====================// ==========================================
// 1. CONFIGURATION VARIABLES
// ==========================================
// Tip: You can paste your IDs directly below OR set them in Project Settings > Script Properties.
const DEFAULT_SPREADSHEET_ID = '1b2ldt35HijbAdsroXUCcZoGLWDuctfwOugmCwMd2QFo';
const DEFAULT_DRIVE_FOLDER_ID = ''; // Optional: Paste a Google Drive Folder ID for site attachments
const DEFAULT_CALENDAR_ID = '';     // Optional: Paste a specific Google Calendar ID, or leave blank for primary calendar
const ADMIN_EMAILS = ['mcw.services2024@outlook.com', 'info@mazibukocivilworks.co.za'];
const LOCK_WAIT_TIMEOUT = 10000; // 10 seconds
const BASE_LOCATION = 'Soweto, Gauteng, South Africa';
const MAX_UPLOAD_BYTES = 2.5 * 1024 * 1024; // 2.5 MB hard server limit

const ALLOWED_URGENCIES = ['Standard / Flexible', 'This Week', 'Urgent / Same-Day'];
const ALLOWED_MIME_TYPES = ['image/jpeg', 'image/jpg', 'image/png', 'image/webp', 'application/pdf'];

const QUOTE_HEADERS = [
  'Timestamp',
  'Reference ID',
  'Full Name',
  'Email',
  'Phone / WhatsApp',
  'Service Requested',
  'Project Location',
  'Distance to Site',
  'Callout Fee',
  'Preferred Date',
  'Urgency',
  'Project Details',
  'File Attachment',
  'Status',
  'WhatsApp Quick Reply',
  'Calendar Event ID'
];

const FAQ_HEADERS = [
  'Timestamp',
  'Reference ID',
  'Full Name',
  'Email',
  'Phone / WhatsApp',
  'Inquiry Topic',
  'Question Text',
  'Status',
  'WhatsApp Quick Reply'
];

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    spreadsheetId: props.getProperty('SPREADSHEET_ID') || DEFAULT_SPREADSHEET_ID,
    driveFolderId: props.getProperty('DRIVE_FOLDER_ID') || DEFAULT_DRIVE_FOLDER_ID,
    calendarId: props.getProperty('DISPATCH_CALENDAR_ID') || DEFAULT_CALENDAR_ID
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

    // Route lightweight estimator distance checks via POST if hosted externally
    if (data && data.action === 'calculateDistance') {
      const routing = checkSuburbDistance(data.location || '');
      return ContentService
        .createTextOutput(JSON.stringify(routing))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Route self-service quote status lookups via POST if hosted externally
    if (data && data.action === 'checkQuoteStatus') {
      const statusResult = checkQuoteStatus(data.refId || '', data.verification || '');
      return ContentService
        .createTextOutput(JSON.stringify(statusResult))
        .setMimeType(ContentService.MimeType.JSON);
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

// Callable directly from index.html via google.script.run.checkSuburbDistance(location)
function checkSuburbDistance(location) {
  const safeLocation = clampString(location, 150);
  return calculateDistanceAndFee(safeLocation);
}

// ==========================================
// 3. SECURED SELF-SERVICE QUOTE STATUS LOOKUP
// ==========================================

function checkQuoteStatus(refIdInput, verificationInput) {
  try {
    const refId = String(refIdInput || '').trim().toUpperCase();
    const verification = String(verificationInput || '').trim();

    if (!/^MCW-\d{4}-\d{3,6}$/.test(refId)) {
      return {
        status: 'error',
        message: 'Please enter a valid Reference ID (e.g., MCW-2026-0101).'
      };
    }

    if (!verification || verification.length > 120) {
      return {
        status: 'error',
        message: 'Please enter the email address or phone number used when requesting your quote.'
      };
    }

    // Rate-limit lookup attempts per Reference ID (max 8 checks per 10 minutes)
    const cache = CacheService.getScriptCache();
    const rateKey = 'MCW_TRACK_' + refId;
    const attempts = parseInt(cache.get(rateKey) || '0', 10);
    if (attempts >= 8) {
      return {
        status: 'error',
        message: 'Too many status lookup attempts for this reference. Please wait 10 minutes or message us on WhatsApp.'
      };
    }
    cache.put(rateKey, String(attempts + 1), 600);

    const config = getConfig();
    if (!config.spreadsheetId || config.spreadsheetId === 'YOUR_SPREADSHEET_ID_HERE') {
      return { status: 'error', message: 'Tracking service is temporarily unavailable.' };
    }

    const ss = SpreadsheetApp.openById(config.spreadsheetId);
    const sheet = ss.getSheetByName('Quotes');
    if (!sheet || sheet.getLastRow() <= 1) {
      return {
        status: 'error',
        message: 'No matching quote found. Please verify your Reference ID and contact details.'
      };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());

    const idxRef = headers.indexOf('reference id');
    const idxName = headers.indexOf('full name');
    const idxEmail = headers.indexOf('email');
    const idxPhone = headers.indexOf('phone / whatsapp');
    const idxService = headers.indexOf('service requested');
    const idxLocation = headers.indexOf('project location');
    const idxFee = headers.indexOf('callout fee');
    const idxPrefDate = headers.indexOf('preferred date');
    const idxUrgency = headers.indexOf('urgency');
    const idxStatus = headers.indexOf('status');
    const idxTimestamp = headers.indexOf('timestamp');

    if (idxRef === -1) {
      return { status: 'error', message: 'Quote records could not be read.' };
    }

    const inputEmail = isValidEmail(verification) ? verification.toLowerCase() : '';
    const inputDigits = verification.replace(/\D/g, '');
    const inputLast9 = inputDigits.length >= 9 ? inputDigits.slice(-9) : '';

    if (!inputEmail && !inputLast9) {
      return {
        status: 'error',
        message: 'Please enter a valid email address or 10-digit South African phone number.'
      };
    }

    // Search newest rows first
    for (let r = data.length - 1; r >= 1; r--) {
      const row = data[r];
      const rowRef = String(row[idxRef] || '').trim().toUpperCase();
      if (rowRef !== refId) continue;

      const rowEmail = idxEmail >= 0 ? String(row[idxEmail] || '').trim().toLowerCase() : '';
      const rowPhoneDigits = idxPhone >= 0 ? String(row[idxPhone] || '').replace(/\D/g, '') : '';
      const rowLast9 = rowPhoneDigits.length >= 9 ? rowPhoneDigits.slice(-9) : '';

      const emailMatches = Boolean(inputEmail && rowEmail && inputEmail === rowEmail);
      const phoneMatches = Boolean(inputLast9 && rowLast9 && inputLast9 === rowLast9);

      if (emailMatches || phoneMatches) {
        const rawFullName = idxName >= 0 ? String(row[idxName] || '').trim() : '';
        const firstName = rawFullName ? rawFullName.split(/\s+/)[0] : 'Client';

        let submittedStr = '';
        if (idxTimestamp >= 0 && row[idxTimestamp]) {
          const ts = row[idxTimestamp] instanceof Date ? row[idxTimestamp] : new Date(row[idxTimestamp]);
          if (!isNaN(ts.getTime())) {
            submittedStr = Utilities.formatDate(ts, 'Africa/Johannesburg', 'dd MMM yyyy');
          }
        }

        let prefDateStr = idxPrefDate >= 0 ? String(row[idxPrefDate] || '').trim() : '';
        if (idxPrefDate >= 0 && row[idxPrefDate] instanceof Date) {
          prefDateStr = Utilities.formatDate(row[idxPrefDate], 'Africa/Johannesburg', 'yyyy-MM-dd');
        }

        // Return only non-sensitive summary fields
        return {
          status: 'success',
          quote: {
            refId: rowRef,
            firstName: firstName,
            service: idxService >= 0 ? String(row[idxService] || 'Property & Civil Service') : 'Property & Civil Service',
            location: idxLocation >= 0 ? String(row[idxLocation] || '') : '',
            calloutFee: idxFee >= 0 ? String(row[idxFee] || 'Pending Review') : 'Pending Review',
            preferredDate: prefDateStr || 'Flexible',
            urgency: idxUrgency >= 0 ? String(row[idxUrgency] || 'Standard / Flexible') : 'Standard / Flexible',
            quoteStatus: idxStatus >= 0 ? String(row[idxStatus] || 'New') : 'New',
            submittedDate: submittedStr
          }
        };
      }
    }

    // Generic response prevents Reference ID enumeration
    return {
      status: 'error',
      message: 'No matching quote found. Please verify your Reference ID and email/phone number.'
    };
  } catch (err) {
    Logger.log('checkQuoteStatus Error: ' + err.toString());
    return {
      status: 'error',
      message: 'Unable to check quote status right now. Please try again shortly.'
    };
  }
}

// ==========================================
// 4. CORE FORM SUBMISSION HANDLER
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

    // 3. Handle Optional Site Photo / PDF Upload (with server-side MIME & size validation)
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

      const feeCellText = routing.status === 'OK'
        ? (routing.fee > 0 ? `R ${routing.fee}` : 'Free (<= 25 km)')
        : 'Pending Review';

      const whatsappFormula = buildWhatsAppSheetFormula(
        data.contactInfo,
        data.fullName,
        refId,
        `your ${data.service} quote request`
      );

      appendMappedRow(sheet, QUOTE_HEADERS, {
        'Timestamp': timestamp,
        'Reference ID': refId,
        'Full Name': sanitizeSheetCell(data.fullName),
        'Email': sanitizeSheetCell(clientEmail),
        'Phone / WhatsApp': sanitizeSheetCell(data.contactInfo),
        'Service Requested': sanitizeSheetCell(data.service),
        'Project Location': sanitizeSheetCell(data.location),
        'Distance to Site': sanitizeSheetCell(routing.distanceText),
        'Callout Fee': sanitizeSheetCell(feeCellText),
        'Preferred Date': sanitizeSheetCell(data.preferredDate || 'Flexible'),
        'Urgency': sanitizeSheetCell(data.urgency),
        'Project Details': sanitizeSheetCell(data.details),
        'File Attachment': fileUrl,
        'Status': 'New',
        'WhatsApp Quick Reply': whatsappFormula,
        'Calendar Event ID': ''
      });

      const pdfBlob = generateQuotePdf(data);
      const clientName = escapeHtml(data.fullName);
      const scheduleLine = data.preferredDate
        ? ` Preferred date: <strong>${escapeHtml(data.preferredDate)}</strong> (${escapeHtml(data.urgency)}).`
        : ` Priority: <strong>${escapeHtml(data.urgency)}</strong>.`;

      const clientEmailBody = getProfessionalEmailTemplate(
        `Quotation Request Confirmation (${refId})`,
        `Dear <strong>${clientName}</strong>,`,
        `Thank you for contacting <strong>Mazibuko Civil Works (MCW)</strong>.<br><br>
        We acknowledge receipt of your project inquiry (Reference: <strong>${refId}</strong>) for <strong>${escapeHtml(data.service)}</strong> in <strong>${escapeHtml(data.location)}</strong>.${scheduleLine} Our operations team is reviewing your specifications to prepare a formal quotation.<br><br>
        For your records, an itemized PDF summary—including your estimated site distance and callout details—is attached to this email.`,
        'Track Your Quote Online Anytime',
        `Quote Reference: <strong>${refId}</strong><br>Use the <strong>Track My Quote</strong> tab on our website with your Reference ID and email/phone number, or message our dispatch team on WhatsApp at <strong>+27 69 877 7757</strong>.`,
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

      const whatsappFormula = buildWhatsAppSheetFormula(
        data.contactInfo,
        data.fullName,
        refId,
        `your inquiry (${data.topic})`
      );

      appendMappedRow(sheet, FAQ_HEADERS, {
        'Timestamp': timestamp,
        'Reference ID': refId,
        'Full Name': sanitizeSheetCell(data.fullName),
        'Email': sanitizeSheetCell(clientEmail),
        'Phone / WhatsApp': sanitizeSheetCell(data.contactInfo),
        'Inquiry Topic': sanitizeSheetCell(data.topic),
        'Question Text': sanitizeSheetCell(data.question),
        'Status': 'New',
        'WhatsApp Quick Reply': whatsappFormula
      });

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
// 5. SHEETS STATUS & GOOGLE CALENDAR DISPATCH SYNC
// ==========================================

/**
 * Run this function ONCE from the Apps Script editor (Select 'setupEditTrigger' > Run)
 * to authorize Spreadsheet, Email, and Google Calendar permissions and install the trigger.
 */
function setupEditTrigger() {
  const config = getConfig();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);

  // Touch CalendarApp during setup so Google prompts for Calendar scope upfront
  CalendarApp.getDefaultCalendar().getName();

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onSheetEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('onSheetEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  Logger.log('Installed onSheetEdit trigger with Calendar sync successfully.');
}

function onSheetEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== 'Quotes') return;

    const row = e.range.getRow();
    if (row <= 1) return;

    const newStatus = String(e.value || '').trim();
    const oldStatus = String(e.oldValue || '').trim();
    if (!newStatus || newStatus.toLowerCase() === oldStatus.toLowerCase()) return;

    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim().toLowerCase());
    const statusCol = headers.indexOf('status') + 1;

    if (statusCol === 0 || e.range.getColumn() !== statusCol) return;

    const rowValues = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
    const getVal = (headerName) => {
      const idx = headers.indexOf(headerName.toLowerCase());
      if (idx < 0) return '';
      const rawVal = rowValues[idx];
      if (rawVal instanceof Date) {
        return Utilities.formatDate(rawVal, 'Africa/Johannesburg', 'yyyy-MM-dd');
      }
      return String(rawVal || '').replace(/^'/, '').trim();
    };

    const refId = getVal('reference id') || 'MCW-REF';
    const fullName = getVal('full name') || 'Valued Client';
    const clientEmail = getVal('email');
    const phone = getVal('phone / whatsapp');
    const service = getVal('service requested') || 'Property & Civil Services';
    const location = getVal('project location') || 'Site Location TBC';
    const distance = getVal('distance to site');
    const calloutFee = getVal('callout fee');
    const preferredDate = getVal('preferred date');
    const urgency = getVal('urgency');
    const details = getVal('project details');
    const fileAttachment = getVal('file attachment');
    const existingEventId = getVal('calendar event id');

    const normalized = newStatus.toLowerCase();
    const safeName = escapeHtml(fullName);
    const safeService = escapeHtml(service);
    const safeLocation = escapeHtml(location);
    const safeRef = escapeHtml(refId);

    if (normalized === 'site visit scheduled' || normalized === 'confirmed') {
      const eventLabel = normalized === 'site visit scheduled' ? 'Site Visit' : 'Confirmed Job';
      const eventId = syncDispatchCalendarEvent({
        existingEventId: existingEventId,
        eventLabel: eventLabel,
        refId: refId,
        fullName: fullName,
        phone: phone,
        clientEmail: clientEmail,
        service: service,
        location: location,
        distance: distance,
        calloutFee: calloutFee,
        preferredDate: preferredDate,
        urgency: urgency,
        details: details,
        fileAttachment: fileAttachment
      });

      if (eventId) {
        let calColIdx = headers.indexOf('calendar event id') + 1;
        if (calColIdx === 0) {
          calColIdx = sheet.getLastColumn() + 1;
          sheet.getRange(1, calColIdx).setValue('Calendar Event ID').setFontWeight('bold');
        }
        sheet.getRange(row, calColIdx).setValue(eventId);
      }
    }

    if (!isValidEmail(clientEmail)) return;

    if (normalized === 'site visit scheduled') {
      const body = getProfessionalEmailTemplate(
        `Site Inspection Scheduled (${safeRef})`,
        `Dear <strong>${safeName}</strong>,`,
        `Our operations team has scheduled a site visit for your <strong>${safeService}</strong> request in <strong>${safeLocation}</strong> (Reference: <strong>${safeRef}</strong>).<br><br>
        A site supervisor will reach out via phone or WhatsApp shortly to confirm your preferred arrival window.`,
        'Need to Reschedule or Share Pin Location?',
        `Reply to this email or send your live WhatsApp pin to <strong>+27 69 877 7757</strong> quoting <strong>${safeRef}</strong>.`,
        '#B5882D'
      );
      sendMcwEmail(clientEmail, `MCW Site Visit Scheduled [${refId}]`, body);
    } else if (normalized === 'confirmed') {
      const body = getProfessionalEmailTemplate(
        `Booking Confirmed (${safeRef})`,
        `Dear <strong>${safeName}</strong>,`,
        `Your booking with <strong>Mazibuko Civil Works (MCW)</strong> for <strong>${safeService}</strong> in <strong>${safeLocation}</strong> (Reference: <strong>${safeRef}</strong>) is officially <strong>confirmed</strong>.<br><br>
        Our crew and equipment are being allocated for your site. We will send a dispatch notice prior to arrival.`,
        'Operations & Dispatch Desk',
        `Quote Reference: <strong>${safeRef}</strong><br>Direct WhatsApp Line: <strong>+27 69 877 7757</strong>`,
        '#487138'
      );
      sendMcwEmail(clientEmail, `MCW Booking Confirmed [${refId}]`, body);
    } else if (normalized === 'completed') {
      const body = getProfessionalEmailTemplate(
        `Project Completed (${safeRef})`,
        `Dear <strong>${safeName}</strong>,`,
        `Thank you for choosing <strong>Mazibuko Civil Works (MCW)</strong> for your <strong>${safeService}</strong> project in <strong>${safeLocation}</strong> (Reference: <strong>${safeRef}</strong>).<br><br>
        Your job has been marked as completed. If you require ongoing scheduled maintenance or additional site support, our team is ready to assist.`,
        'We Value Your Feedback',
        `Have feedback or want to schedule recurring service? Message us anytime on WhatsApp at <strong>+27 69 877 7757</strong>.`,
        '#487138'
      );
      sendMcwEmail(clientEmail, `Thank You from Mazibuko Civil Works [${refId}]`, body);
    }
  } catch (err) {
    Logger.log('onSheetEdit Error: ' + err.toString());
  }
}

function syncDispatchCalendarEvent(info) {
  try {
    const config = getConfig();
    const calendar = config.calendarId
      ? CalendarApp.getCalendarById(config.calendarId)
      : CalendarApp.getDefaultCalendar();

    if (!calendar) {
      Logger.log('Dispatch calendar not found.');
      return '';
    }

    // Parse Preferred Date (YYYY-MM-DD) in SAST (+02:00), or default to tomorrow at 09:00 SAST
    let startTime;
    if (/^\d{4}-\d{2}-\d{2}$/.test(info.preferredDate)) {
      startTime = new Date(`${info.preferredDate}T09:00:00+02:00`);
    }
    if (!startTime || isNaN(startTime.getTime())) {
      const tomorrow = new Date();
      tomorrow.setDate(tomorrow.getDate() + 1);
      const ymd = Utilities.formatDate(tomorrow, 'Africa/Johannesburg', 'yyyy-MM-dd');
      startTime = new Date(`${ymd}T09:00:00+02:00`);
    }
    const endTime = new Date(startTime.getTime() + 90 * 60 * 1000); // 1.5-hour dispatch window

    const title = `[MCW ${info.eventLabel}] ${info.service} - ${info.fullName} (${info.location})`;
    const waPhone = formatSaPhoneForWhatsApp(info.phone);
    const waLink = waPhone
      ? `https://wa.me/${waPhone}?text=${encodeURIComponent(`Hi ${info.fullName}, MCW Dispatch confirming our visit for ${info.refId}.`)}`
      : 'Not provided';

    const description = [
      `Reference ID: ${info.refId}`,
      `Dispatch Stage: ${info.eventLabel}`,
      `Client Name: ${info.fullName}`,
      `Phone / WhatsApp: ${info.phone}`,
      `WhatsApp Direct Chat: ${waLink}`,
      `Client Email: ${info.clientEmail}`,
      `Service: ${info.service}`,
      `Location: ${info.location}`,
      `Distance / Callout: ${info.distance || 'N/A'} (${info.calloutFee || 'Pending'})`,
      `Urgency: ${info.urgency || 'Standard / Flexible'}`,
      `Site Attachment: ${info.fileAttachment || 'None'}`,
      ``,
      `Project Scope / Notes:`,
      info.details || 'No additional notes.'
    ].join('\n');

    // Update existing event if present so status changes stay idempotent
    if (info.existingEventId) {
      try {
        const existingEvent = calendar.getEventById(info.existingEventId);
        if (existingEvent) {
          existingEvent.setTitle(title);
          existingEvent.setTime(startTime, endTime);
          existingEvent.setLocation(info.location);
          existingEvent.setDescription(description);
          return existingEvent.getId();
        }
      } catch (lookupErr) {
        Logger.log('Existing calendar event not found; creating new event.');
      }
    }

    const createdEvent = calendar.createEvent(title, startTime, endTime, {
      location: info.location,
      description: description
    });

    return createdEvent.getId();
  } catch (calErr) {
    Logger.log('Calendar Sync Error: ' + calErr.toString());
    return '';
  }
}

// ==========================================
// 6. SECURED SUBURB DISTANCE & CALLOUT ENGINE
// ==========================================

function calculateDistanceAndFee(clientLocation) {
  let dest = String(clientLocation || '').trim();
  if (!dest) {
    return {
      distanceText: 'No Location Provided',
      distanceKm: 0,
      fee: 0,
      status: 'NO_LOCATION',
      feeMsg: 'Enter a suburb to check callout fee'
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
        feeMsg: 'Suburb not auto-verified; callout confirmed on quote review'
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
// 7. DATA NORMALIZATION & SECURITY HELPERS
// ==========================================

function clampString(value, maxLen) {
  const str = String(value || '').trim();
  return str.length > maxLen ? str.substring(0, maxLen) : str;
}

// Neutralizes Spreadsheet Formula Injection (=, +, -, @, tab, CR)
function sanitizeSheetCell(value) {
  const str = String(value == null ? '' : value);
  if (/^[=+\-@\t\r]/.test(str)) {
    return "'" + str;
  }
  return str;
}

// Appends a row aligned to existing headers (or initializes headers on a new sheet)
function appendMappedRow(sheet, defaultHeaders, recordMap) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(defaultHeaders);
    sheet.getRange(1, 1, 1, defaultHeaders.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  let currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());

  // Automatically add any new columns (e.g., Preferred Date, Urgency, Calendar Event ID) to existing sheets
  Object.keys(recordMap).forEach(key => {
    const exists = currentHeaders.some(h => h.toLowerCase() === key.toLowerCase());
    if (!exists) {
      const newColIdx = currentHeaders.length + 1;
      sheet.getRange(1, newColIdx).setValue(key).setFontWeight('bold');
      currentHeaders.push(key);
    }
  });

  const rowValues = currentHeaders.map(header => {
    const matchKey = Object.keys(recordMap).find(k => k.toLowerCase() === header.toLowerCase());
    return matchKey !== undefined ? recordMap[matchKey] : '';
  });

  sheet.appendRow(rowValues);
}

function normalizeSubmissionData(data) {
  const raw = data || {};
  const rawType = String(raw.type || 'quote').trim().toLowerCase();
  const type = (rawType === 'quote' || rawType === 'faq') ? rawType : 'invalid';

  const contactInfo = clampString(raw.contactInfo, 40);
  const email = clampString(raw.email, 120);
  const fullName = clampString(raw.fullName, 100);
  const details = clampString(raw.details, 3000);
  const question = clampString(raw.question || details, 3000);

  const rawDate = clampString(raw.preferredDate, 20);
  const preferredDate = /^\d{4}-\d{2}-\d{2}$/.test(rawDate) ? rawDate : '';

  const rawUrgency = clampString(raw.urgency, 40);
  const urgency = ALLOWED_URGENCIES.includes(rawUrgency) ? rawUrgency : 'Standard / Flexible';

  return {
    type: type,
    website_hp: clampString(raw.website_hp, 100),
    email: email || (isValidEmail(contactInfo) ? contactInfo : ''),
    contactInfo: contactInfo,
    fullName: fullName,
    topic: clampString(raw.topic || 'General Inquiry', 100),
    question: question,
    details: details,
    service: clampString(raw.service || 'General Service', 80),
    location: clampString(raw.location, 150),
    preferredDate: preferredDate,
    urgency: urgency,
    file: (raw.file && typeof raw.file === 'object') ? raw.file : null
  };
}

// Generates sequential IDs (e.g., MCW-2026-0101) safely inside LockService
function generateReferenceId() {
  const year = new Date().getFullYear();
  const props = PropertiesService.getScriptProperties();
  const key = `REF_SEQ_${year}`;
  const current = parseInt(props.getProperty(key) || '100', 10);
  const next = current + 1;
  props.setProperty(key, String(next));
  return `MCW-${year}-${String(next).padStart(4, '0')}`;
}

function formatSaPhoneForWhatsApp(phoneInput) {
  let digits = String(phoneInput || '').replace(/\D/g, '');
  if (!digits) return '';
  if (digits.startsWith('0') && digits.length === 10) {
    digits = '27' + digits.substring(1);
  }
  return digits.length >= 11 ? digits : '';
}

function buildWhatsAppSheetFormula(phoneInput, clientName, refId, contextLabel) {
  const waPhone = formatSaPhoneForWhatsApp(phoneInput);
  if (!waPhone) return 'No Valid Mobile';
  const safeClient = String(clientName || 'Client').replace(/[\r\n"]/g, ' ').trim();
  const safeContext = String(contextLabel || 'your request').replace(/[\r\n"]/g, ' ').trim();
  const greeting = `Hi ${safeClient}, this is Mazibuko Civil Works (MCW) regarding ${safeContext} [${refId}].`;
  const url = `https://wa.me/${waPhone}?text=${encodeURIComponent(greeting)}`;
  const safeUrl = url.replace(/"/g, '""');
  return `=HYPERLINK("${safeUrl}", "Chat on WhatsApp")`;
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
// 8. SECURED GOOGLE DRIVE ATTACHMENT STORAGE
// ==========================================

function saveFileToDrive(fileObj, prefixName, driveFolderId) {
  try {
    if (!fileObj || typeof fileObj.data !== 'string' || !fileObj.data.includes(',')) {
      return 'None';
    }

    const splitData = fileObj.data.split(',');
    const mimeMatch = splitData[0].match(/^data:([a-zA-Z0-9/+.-]+);base64$/);
    if (!mimeMatch) {
      Logger.log('Blocked attachment with malformed data URI header.');
      return 'Invalid File Format';
    }

    const contentType = mimeMatch[1].toLowerCase();
    if (!ALLOWED_MIME_TYPES.includes(contentType)) {
      Logger.log('Blocked unsupported attachment MIME type: ' + contentType);
      return 'Unsupported File Type';
    }

    const decodedBytes = Utilities.base64Decode(splitData[1]);
    if (decodedBytes.length > MAX_UPLOAD_BYTES) {
      Logger.log('Blocked oversized attachment: ' + decodedBytes.length + ' bytes.');
      return 'File Exceeded Size Limit';
    }

    let folder;
    try {
      folder = driveFolderId
        ? DriveApp.getFolderById(driveFolderId)
        : DriveApp.getRootFolder();
    } catch (folderErr) {
      Logger.log('Invalid Drive Folder ID; falling back to root folder: ' + folderErr.toString());
      folder = DriveApp.getRootFolder();
    }

    const safePrefix = String(prefixName || 'Client').replace(/[^a-zA-Z0-9_-]/g, '_').substring(0, 50);
    const safeFileName = String(fileObj.name || 'attachment').replace(/[^a-zA-Z0-9._-]/g, '_').substring(0, 60);

    const blob = Utilities.newBlob(decodedBytes, contentType, `${safePrefix}_${safeFileName}`);
    const file = folder.createFile(blob);

    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      Logger.log('Could not set public link sharing (restricted by Workspace policy): ' + shareErr.toString());
    }

    return file.getUrl();
  } catch (e) {
    Logger.log('File upload failed: ' + e.toString());
    return 'Upload Failed';
  }
}

// ==========================================
// 9. PDF & EMAIL DISPATCHERS
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
  const isUrgent = type === 'quote' && data.urgency === 'Urgent / Same-Day';
  const urgentPrefix = isUrgent ? '🚨 URGENT ' : '🔔 ';
  const subject = `${urgentPrefix}[${refId}] New MCW ${type.toUpperCase()}: ${data.fullName} (${data.location || 'Inquiry'})`;
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

    const isSafeUrl = typeof fileUrl === 'string' && fileUrl.startsWith('https://');
    const attachmentHtml = isSafeUrl
      ? `<a href="${escapeHtml(fileUrl)}" target="_blank" rel="noopener" style="color: #487138; font-weight: bold;">Open Site Attachment in Google Drive</a>`
      : escapeHtml(fileUrl || 'None');

    const urgentBanner = isUrgent
      ? `<div style="background: #fef2f2; color: #b91c1c; padding: 10px 14px; border: 1px solid #fecaca; border-radius: 6px; font-weight: bold; margin-bottom: 12px;">🚨 URGENT / SAME-DAY DISPATCH REQUESTED</div>`
      : '';

    detailsHtml = `
      ${urgentBanner}
      <p><strong>Service Requested:</strong> ${escapeHtml(data.service)}</p>
      <p><strong>Project Location:</strong> ${escapeHtml(data.location)}</p>
      <p><strong>Preferred Date:</strong> ${escapeHtml(data.preferredDate || 'Flexible')}</p>
      <p><strong>Urgency Level:</strong> ${escapeHtml(data.urgency)}</p>
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
        <p style="font-size: 13px;"><strong>Preferred Date:</strong> ${escapeHtml(data.preferredDate || 'Flexible')} &nbsp;|&nbsp; <strong>Priority:</strong> ${escapeHtml(data.urgency)}</p>

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
