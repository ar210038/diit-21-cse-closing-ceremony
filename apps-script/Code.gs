// Adds a private setup command to the Google Sheet menu.
// Open or reload the Sheet, then use DIIT Ceremony > Initialize registration sheet.
function onOpen() {
  SpreadsheetApp.getUi().createMenu('DIIT Ceremony')
    .addItem('Initialize registration sheet', 'setup_')
    .addSeparator()
    .addItem('Resend confirmation for selected student', 'resendConfirmationForSelectedRow_')
    .addToUi();
}

// Google Sheets → Extensions → Apps Script. Paste as Code.gs.
// Bind this script to the Google Sheet that will be your private admin dashboard.
const SHEET_NAME_ = 'Registrations';
const DEADLINE_ = '2026-10-27';
const ZONE_ = 'Asia/Dhaka';
// Exact public, working DIIT Workspace web-app deployment. Keep this synced with docs/config.js.
const TICKET_VERIFICATION_BASE_URL_ = 'https://script.google.com/a/macros/diit.edu.bd/s/AKfycbz8G1aheVeGIxKzsxIMXl783fUMAG-7TZ_XW2pFvoprh6BYJYj_xIl59_B5vPDbFqWB9A/exec';
 // Approved green-and-gold ticket image, hosted publicly with the GitHub Pages site.
 // The filename really has two .png extensions in the repository.
 const TICKET_TEMPLATE_URL_ = 'https://ar210038.github.io/diit-21-cse-closing-ceremony/assets/ticket-template.png.png';

const HEADERS_ = ['Submitted at', 'Reference', 'Student name', 'Student ID', 'Email', 'Phone', 'Guests', 'Total (BDT)', 'Payment method', 'Transaction ID', 'Payment received', 'Status', 'Ticket token', 'Confirmed at', 'Email sent at', 'Email error'];

// Reload the Sheet and choose DIIT Ceremony > Initialize registration sheet; authorize when prompted.
// The setup and payment-edit methods have trailing underscores: web visitors cannot invoke them.
function setup_() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  if (!book) throw Error('Create this script using Extensions > Apps Script from a Google Sheet.');
  let sheet = book.getSheetByName(SHEET_NAME_);
  if (!sheet) sheet = book.insertSheet(SHEET_NAME_);
  sheet.getRange(1, 1, 1, HEADERS_.length).setValues([HEADERS_]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, HEADERS_.length).setBackground('#123d34').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setColumnWidth(3, 170);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(12, 145);
  const existing = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'onPaymentEdited_');
  if (!existing.length) ScriptApp.newTrigger('onPaymentEdited_').forSpreadsheet(book).onEdit().create();
  return 'Ready: only trusted editors should have access to this Google Sheet.';
}

function registrations_() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = book && book.getSheetByName(SHEET_NAME_);
  if (!sheet) throw Error('Run setup_() in Apps Script before opening registration.');
  return sheet;
}
function nowDate_() {
  return Utilities.formatDate(new Date(), ZONE_, 'yyyy-MM-dd');
}
function clean_(value) { return String(value == null ? '' : value).trim(); }
function escape_(value) {
  return clean_(value).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}
function doGet(e) {
  const ticket = e && e.parameter ? clean_(e.parameter.ticket) : '';
  if (ticket) return verifyTicket_(ticket);
  return HtmlService.createHtmlOutputFromFile('Registration')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('DIIT 21st CSE Registration')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// This is the only mutation that public registration visitors may call.
function submitRegistration(form) {
  if (nowDate_() > DEADLINE_) throw Error('Registration closed on 27 October 2026.');
  form = form || {};
  const name = clean_(form.name), studentId = clean_(form.studentId),
    email = clean_(form.email).toLowerCase(),
    phone = clean_(form.phone).replace(/[\s-]/g, ''),
    guests = Number(form.guests), method = clean_(form.paymentMethod).toLowerCase(),
    txn = clean_(form.transactionId);
  if (!/^[\p{L}][\p{L} .'-]{1,99}$/u.test(name)) throw Error('Enter your full name.');
  if (!/^[A-Za-z0-9][A-Za-z0-9/_-]{2,44}$/.test(studentId)) throw Error('Enter a valid student ID.');
  if (!/^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/.test(email) || email.length > 180) throw Error('Enter a valid email address.');
  if (!/^(?:\+?88)?01[3-9]\d{8}$/.test(phone)) throw Error('Enter a valid Bangladeshi phone number.');
  if (!Number.isInteger(guests) || guests < 0 || guests > 10) throw Error('Choose a valid guest count (0–10).');
  if (!['cash','bkash'].includes(method)) throw Error('Choose cash or bKash.');
  if (txn && !/^[A-Za-z0-9_-]{3,40}$/.test(txn)) throw Error('Enter a valid transaction ID or leave it blank.');
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sheet = registrations_();
    const last = sheet.getLastRow();
    if (last > 1 && sheet.getRange(2, 4, last - 1, 1).getDisplayValues()
        .some(row => row[0].toLowerCase() === studentId.toLowerCase())) {
      throw Error('This student ID is already registered. Contact the organizer if you need changes.');
    }
    const reference = 'CSE21-' + Utilities.getUuid().replace(/-/g, '').substring(0, 10).toUpperCase();
    const total = 1000 + guests * 500;
    sheet.appendRow([new Date(), reference, name, studentId, email, phone, guests, total,
      method, txn, false, 'PENDING', '', '', '', '']);
    sheet.getRange(sheet.getLastRow(), 11).insertCheckboxes();
    SpreadsheetApp.flush();
    return {reference:reference, amount:total, status:'PENDING'};
  } finally { lock.releaseLock(); }
}

// Installable on-edit trigger, created by setup_. Only spreadsheet editors can tick Payment received.
function onPaymentEdited_(e) {
  if (!e || !e.range || !e.source) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME_ || e.range.getColumn() !== 11 ||
      e.range.getRow() < 2 || e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1 ||
      e.range.getValue() !== true) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const row = e.range.getRow(), data = sheet.getRange(row, 1, 1, 16).getValues()[0];
    if (String(data[14]).trim()) return; // Already emailed; don't resend accidentally.
    let token = clean_(data[12]);
    if (!token) {
      token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
      sheet.getRange(row, 13).setValue(token);
    }
    sheet.getRange(row, 12).setValue('CONFIRMED');
    if (!data[13]) sheet.getRange(row, 14).setValue(new Date());
    SpreadsheetApp.flush();
    try {
      sendTicketEmail_(data, token);
      sheet.getRange(row, 15).setValue(new Date());
      sheet.getRange(row, 16).clearContent();
    } catch (err) {
      sheet.getRange(row, 16).setValue(String(err).slice(0, 500));
      // Payment remains confirmed; uncheck and recheck the payment box to retry failed email.
    }
  } finally { lock.releaseLock(); }
}

// Private Google Sheets menu action. Use only for a confirmed student whose email needs a correction.
// Existing QR token stays unchanged so the student retains the same ticket identity.
function resendConfirmationForSelectedRow_() {
  const ui = SpreadsheetApp.getUi();
  const book = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = book && book.getActiveSheet();
  const selected = sheet && sheet.getActiveRange();
  if (!sheet || sheet.getName() !== SHEET_NAME_ || !selected ||
      selected.getNumRows() !== 1 || selected.getRow() < 2) {
    ui.alert('Select one student row in the Registrations sheet first.');
    return;
  }
  const row = selected.getRow();
  const record = sheet.getRange(row, 1, 1, HEADERS_.length).getValues()[0];
  if (String(record[11]) !== 'CONFIRMED' || !/^[a-f0-9]{64}$/i.test(String(record[12]))) {
    ui.alert('This student must be CONFIRMED and have a valid ticket token before resending.');
    return;
  }
  const answer = ui.alert('Resend confirmation email?',
      'Student: ' + String(record[2]) + '\\nEmail: ' + String(record[4]) +
      '\\nThis sends a new email with the corrected ticket link and QR image.',
      ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    sendTicketEmail_(record, String(record[12]));
    sheet.getRange(row, 15).setValue(new Date());
    sheet.getRange(row, 16).clearContent();
    ui.alert('Confirmation email sent. Test both the link and QR code in a different browser.');
  } catch (err) {
    sheet.getRange(row, 16).setValue(String(err).slice(0, 500));
    ui.alert('Email failed: ' + String(err));
  } finally { lock.releaseLock(); }
}


/**
 * Build one personalized PDF ticket from the approved landscape PNG.
 * Google Slides renders the original design, the student's real name, and
 * the actual QR code together into a single downloadable one-page PDF.
 * Temporary Slides files are put in trash after PDF conversion.
 * New Slide/Drive permissions must be granted by the sheet owner once.
 */
function createPersonalizedTicketPdf_(studentName, reference, ticketUrl) {
  const templateResponse = UrlFetchApp.fetch(TICKET_TEMPLATE_URL_, {muteHttpExceptions:true});
  if (templateResponse.getResponseCode() !== 200 ||
      !/image\/png/i.test(String(templateResponse.getHeaders()['Content-Type'] ||
        templateResponse.getHeaders()['content-type'] || ''))) {
    throw Error('Ticket artwork unavailable. Verify TICKET_TEMPLATE_URL_ is public (HTTP ' +
      templateResponse.getResponseCode() + ').');
  }
  const artwork = templateResponse.getBlob().setName('diit-ticket-artwork.png');

  const qrResponse = UrlFetchApp.fetch(
    'https://quickchart.io/qr?format=png&size=800&margin=2&text=' + encodeURIComponent(ticketUrl),
    {muteHttpExceptions:true});
  if (qrResponse.getResponseCode() !== 200 ||
      !/image\/png/i.test(String(qrResponse.getHeaders()['Content-Type'] ||
        qrResponse.getHeaders()['content-type'] || ''))) {
    throw Error('Ticket QR generation failed (HTTP ' + qrResponse.getResponseCode() + '). Retry later.');
  }
  const qrImage = qrResponse.getBlob().setName('ticket-qr.png');

  let presentationId = null;
  try {
    const deck = SlidesApp.create('Temporary DIIT CSE21 ticket - ' + reference);
    presentationId = deck.getId();
    // Default Slides canvas is 16:9. Center the complete 3:1 ticket image
    // without stretching/cropping its approved artwork or QR placement.
    const initialSlides = deck.getSlides();
    const slide = deck.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    initialSlides.forEach(oldSlide => oldSlide.remove());
    slide.getBackground().setSolidFill('#073B30');

    const pageW = deck.getPageWidth(), pageH = deck.getPageHeight();
    const ticketW = Math.min(pageW * 0.98, pageH * 2.96);
    const scale = ticketW / 2172; // Approved artwork: 2172 × 724 pixels.
    const x = (pageW - ticketW) / 2;
    const y = (pageH - 724 * scale) / 2;
    slide.insertImage(artwork, x, y, 2172 * scale, 724 * scale);

    // Cover ONLY the printed sample name inside its ornamental frame.
    // Preserve the decorative gold outline, DIIT logo and ceremony title.
    const namePlate = slide.insertShape(SlidesApp.ShapeType.RECTANGLE,
      x + 674 * scale, y + 491 * scale, 816 * scale, 88 * scale);
    namePlate.getFill().setSolidFill('#FBF5E9');
    namePlate.getBorder().setTransparent();
    namePlate.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    const nameText = namePlate.getText();
    nameText.setText(String(studentName));
    const count = [...String(studentName)].length;
    const fontSize = count > 46 ? 10 : count > 34 ? 12 : count > 25 ? 15 : 18;
    nameText.getTextStyle().setFontFamily('Georgia').setFontSize(fontSize)
      .setBold(true).setForegroundColor('#0B2F26');
    nameText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // Place the REAL QR over the decorative placeholder, inside the gold frame.
    // This image is embedded into the PDF, not linked to a third-party website.
    // Mask the sample QR completely before putting the real QR within the frame.
    const qrMask = slide.insertShape(SlidesApp.ShapeType.RECTANGLE,
      x + 1730 * scale, y + 234 * scale, 283 * scale, 283 * scale);
    qrMask.getFill().setSolidFill('#FFFDFC');
    qrMask.getBorder().setTransparent();
    slide.insertImage(qrImage,
      x + 1750 * scale, y + 254 * scale, 243 * scale, 243 * scale);
    deck.saveAndClose();

    return DriveApp.getFileById(presentationId).getAs(MimeType.PDF)
      .setName('DIIT-CSE21-' + reference + '-Ticket.pdf');
  } finally {
    if (presentationId) {
      try { DriveApp.getFileById(presentationId).setTrashed(true); }
      catch (cleanupError) { console.log('Temporary ticket slide cleanup: ' + cleanupError); }
    }
  }
}

function sendTicketEmail_(data, token) {
  // ScriptApp.getService().getUrl() may return a different, inaccessible Workspace deployment.
  const ticketUrl = TICKET_VERIFICATION_BASE_URL_ + '?ticket=' + encodeURIComponent(token);
  const name = String(data[2]), reference = String(data[1]),
    email = String(data[4]), guests = Number(data[6]), total = Number(data[7]);
  const subject = 'DIIT 21st CSE Closing Ceremony — Registration Confirmed';
  const plain = 'Dear ' + name + ',\n\nYour payment has been confirmed!\nReference: ' +
    reference + '\nStudent ID: ' + data[3] + '\nGuests: ' + guests + '\nAmount: BDT ' +
    total + '\nVenue: Green Garden\nCeremony date: to be announced after final defence.' +
    '\nYour entry ticket: ' + ticketUrl + '\n\nAl-Rafi Chowdhury · 01940372926';
  const html = '<div style="font:16px Arial,sans-serif;max-width:580px;color:#123d34">' +
    '<h2>🎓 DIIT 21st CSE Closing Ceremony</h2><h3>Payment confirmed!</h3>' +
    '<p>Dear ' + escape_(name) + ', your registration is confirmed.</p>' +
    '<p><b>Reference:</b> ' + escape_(reference) + '<br><b>Student ID:</b> ' + escape_(data[3]) +
    '<br><b>Guests:</b> ' + guests + '<br><b>Total paid:</b> BDT ' + total +
    '</p><p><b>Venue:</b> <a href="https://greengarden.com.bd/">Green Garden</a><br>' +
    '<b>Date:</b> To be announced after the final defence.</p>' +
    '<p>Your personalized ceremony entry ticket is attached as a PDF. Show its QR code at entry. <a href="' + escape_(ticketUrl) +
    '">Open your digital ticket</a>.</p>' +
    '<p>Contact: Al-Rafi Chowdhury · 01940372926</p></div>';
  // A successful email now carries the personalized, scannable PDF instead of a bare QR PNG.
  // A generation failure propagates to column P, so the organizer can retry safely.
  const ticketPdf = createPersonalizedTicketPdf_(name, reference, ticketUrl);
  const options = {to:email, subject:subject, body:plain, htmlBody:html,
    attachments:[ticketPdf]};
  MailApp.sendEmail(options);
}

function verifyTicket_(token) {
  const sheet = registrations_();
  let found = null, rows = sheet.getLastRow();
  if (/^[a-fA-F0-9]{64}$/.test(token) && rows > 1) {
    const values = sheet.getRange(2, 1, rows - 1, 16).getValues();
    found = values.find(r => String(r[12]) === token && String(r[11]) === 'CONFIRMED') || null;
  }
  const body = found ?
    '<h1>✓ Verified ticket</h1><p><b>Student:</b> ' + escape_(found[2]) +
    '</p><p><b>Student ID:</b> ' + escape_(found[3]) +
    '</p><p><b>Guests:</b> ' + Number(found[6]) +
    '</p><p><b>Reference:</b> ' + escape_(found[1]) + '</p>' :
    '<h1>Ticket not verified</h1><p>Invalid or unconfirmed ticket.</p>';
  return HtmlService.createHtmlOutput('<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><title>DIIT Ticket</title></head>' +
    '<body style="background:#f4f6ef;color:#123d34;padding:30px;font:17px Arial,sans-serif">' +
    '<main style="max-width:500px;margin:auto;background:white;padding:32px;border-radius:12px">' +
    '<p>🎓 DIIT · 21st CSE Closing Ceremony</p>' + body +
    '<p>Venue: <a href="https://greengarden.com.bd/">Green Garden</a></p>' +
    '</main></body></html>');
}
