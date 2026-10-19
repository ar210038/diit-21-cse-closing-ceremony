// Adds a private setup command to the Google Sheet menu.
// Open or reload the Sheet, then use DIIT Ceremony > Initialize registration sheet.
function onOpen() {
  SpreadsheetApp.getUi().createMenu('DIIT Ceremony')
    .addItem('Initialize registration sheet', 'setup_')
    .addItem('Create / refresh registration totals', 'refreshRegistrationTotals_')
    .addSeparator()
    .addItem('Confirm payment for selected student', 'confirmPaymentForSelectedRow_')
    .addItem('Preview sample ticket PDF', 'testTicketPdf_')
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
// Google Slides ticket template: a single 3:1 slide with {{NAME}} and QR_PLACEHOLDER.
// The script copies this presentation per ticket; it never edits the original.
const TICKET_SLIDES_TEMPLATE_ID_ = '1sO-h3YSwKfhnlxO9s_4GUzxvAzovGWyixFmRjfI-kEk';

const HEADERS_ = ['Submitted at', 'Reference', 'Student name', 'Student ID', 'Email', 'Phone', 'Guests', 'Total (BDT)', 'Payment method', 'Transaction ID', 'Payment received', 'Status', 'Ticket token', 'Confirmed at', 'Email sent at', 'Email error'];

// Reload the Sheet and choose DIIT Ceremony > Initialize registration sheet; authorize when prompted.
// Sheet-only payment methods end with underscores: public registration visitors cannot invoke them.
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
  ensureRegistrationTotals_(sheet);
  return 'Ready: registration summary is visible below the students. Keep the sheet restricted to trusted editors.';
}

// A four-row, formula-driven summary sits immediately AFTER the last registration.
// Row 1 of the summary contains guests in G and total expected fees in H.
// Later rows contain received fees, outstanding fees, and student count.
// A marker in column A lets submissions insert ABOVE it instead of overwriting it.
const REGISTRATION_TOTALS_MARKER_ = 'REGISTRATION SUMMARY';
const REGISTRATION_TOTALS_ROWS_ = 4;

function registrationTotalsRow_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const columnA = sheet.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (let i = columnA.length - 1; i >= 0; i--) {
    if (String(columnA[i][0]).trim() === REGISTRATION_TOTALS_MARKER_) {
      return i + 2;
    }
  }
  return null;
}

function ensureRegistrationTotals_(sheet) {
  let summaryRow = registrationTotalsRow_(sheet);
  if (summaryRow === null) {
    // Existing student rows stay where they are; create the totals after them.
    summaryRow = sheet.getLastRow() + 1;
    const requiredRows = summaryRow + REGISTRATION_TOTALS_ROWS_ - 1;
    if (requiredRows > sheet.getMaxRows()) {
      sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
    }
    sheet.getRange(summaryRow, 1).setValue(REGISTRATION_TOTALS_MARKER_);
  }
  updateRegistrationTotals_(sheet, summaryRow);
  return summaryRow;
}

function updateRegistrationTotals_(sheet, summaryRow) {
  const dataEnd = summaryRow - 1;
  // With no registrations yet, avoid self-referencing or reversed ranges.
  const guestTotal = dataEnd < 2 ? '=0' :
    '=SUM(G2:G' + dataEnd + ')';
  const feeTotal = dataEnd < 2 ? '=0' :
    '=SUM(H2:H' + dataEnd + ')';
  const received = dataEnd < 2 ? '=0' :
    '=SUMIFS(H2:H' + dataEnd + ',K2:K' + dataEnd +
      ',TRUE,L2:L' + dataEnd + ',"CONFIRMED")';
  const studentCount = dataEnd < 2 ? '=0' :
    '=COUNTIF(D2:D' + dataEnd + ',"<>")';

  sheet.getRange(summaryRow, 1).setValue(REGISTRATION_TOTALS_MARKER_);
  sheet.getRange(summaryRow, 6).setValue('TOTAL');
  sheet.getRange(summaryRow, 7).setFormula(guestTotal);
  sheet.getRange(summaryRow, 8).setFormula(feeTotal);
  sheet.getRange(summaryRow + 1, 6).setValue('Amount received');
  sheet.getRange(summaryRow + 1, 8).setFormula(received);
  sheet.getRange(summaryRow + 2, 6).setValue('Amount pending');
  sheet.getRange(summaryRow + 2, 8)
    .setFormula('=H' + summaryRow + '-H' + (summaryRow + 1));
  sheet.getRange(summaryRow + 3, 6).setValue('Students registered');
  sheet.getRange(summaryRow + 3, 8).setFormula(studentCount);

  // The totals are displayed in the existing Guests (G) / Amount (H)
  // columns, while labels sit in F (Phone) and A (summary title).
  sheet.getRange(summaryRow, 1, REGISTRATION_TOTALS_ROWS_, HEADERS_.length)
    .setBackground('#eaf2ed').setFontColor('#123d34');
  sheet.getRange(summaryRow, 1, 1, HEADERS_.length)
    .setBackground('#123d34').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange(summaryRow + 1, 6, 3, 3).setFontWeight('bold');
  sheet.getRange(summaryRow, 7).setNumberFormat('#,##0');
  sheet.getRange(summaryRow, 8, 3, 1).setNumberFormat('"৳"#,##0');
  sheet.getRange(summaryRow + 3, 8).setNumberFormat('#,##0');
  sheet.setColumnWidth(6, Math.max(sheet.getColumnWidth(6), 175));
  sheet.getRange(summaryRow, 7).setNote('Total guests across all registrations.');
  sheet.getRange(summaryRow, 8).setNote('Expected amount: student fees plus guest fees.');
  sheet.getRange(summaryRow + 1, 8).setNote('Received amount: only confirmed payments with Payment received checked.');
  sheet.getRange(summaryRow + 2, 8).setNote('Outstanding amount = expected amount minus received amount.');
  sheet.getRange(summaryRow + 3, 8).setNote('Number of registrations with a student ID.');
}

// Run once after copying new Code.gs; safe to run again without duplicating totals.
// Does not alter the payment verification or email process.
function refreshRegistrationTotals_() {
  const sheet = registrations_();
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const row = ensureRegistrationTotals_(sheet);
    SpreadsheetApp.flush();
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Guests, expected fees, received fees, pending fees and student count are shown at row ' + row + '.',
      'Registration totals', 7);
  } finally {
    lock.releaseLock();
  }
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
    const totalsRow = ensureRegistrationTotals_(sheet);
    const lastStudentRow = totalsRow - 1;
    if (lastStudentRow > 1 &&
        sheet.getRange(2, 4, lastStudentRow - 1, 1).getDisplayValues()
          .some(row => row[0].toLowerCase() === studentId.toLowerCase())) {
      throw Error('This student ID is already registered. Contact the organizer if you need changes.');
    }
    const reference = 'CSE21-' + Utilities.getUuid().replace(/-/g, '').substring(0, 10).toUpperCase();
    const total = 1000 + guests * 500;
    // Always insert BEFORE the footer. appendRow() would place a student after it.
    sheet.insertRowsBefore(totalsRow, 1);
    sheet.getRange(totalsRow, 1, 1, HEADERS_.length).setValues([[
      new Date(), reference, name, studentId, email, phone, guests, total,
      method, txn, false, 'PENDING', '', '', '', ''
    ]]);
    sheet.getRange(totalsRow, 11).insertCheckboxes();
    // Rebuild formula ranges so the newly inserted student is included.
    updateRegistrationTotals_(sheet, totalsRow + 1);
    SpreadsheetApp.flush();
    return {reference:reference, amount:total, status:'PENDING'};
  } finally { lock.releaseLock(); }
}

// The installable on-edit trigger is deliberately NOT a payment approval action.
// Google Sheets does not support a reliable blocking Yes/No dialog BEFORE a checkbox
// edit. Revert direct checkbox edits, including pasted multi-cell ranges.
// Use DIIT Ceremony > Confirm payment for selected student instead.
function onPaymentEdited_(e) {
  if (!e || !e.range || !e.source) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME_) return;
  const firstCol = e.range.getColumn();
  const lastCol = firstCol + e.range.getNumColumns() - 1;
  if (firstCol > 11 || lastCol < 11) return;
  const firstRow = Math.max(2, e.range.getRow());
  const lastRow = e.range.getRow() + e.range.getNumRows() - 1;
  if (firstRow > lastRow) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const records = sheet.getRange(firstRow, 11, lastRow - firstRow + 1, 4)
      .getValues(); // K: Payment received, L: status, M: token, N: confirmed at
    records.forEach((values, offset) => {
      const confirmed = String(values[1]) === 'CONFIRMED' &&
        /^[a-f0-9]{64}$/i.test(String(values[2])) && Boolean(values[3]);
      if (values[0] !== confirmed) {
        sheet.getRange(firstRow + offset, 11).setValue(confirmed);
      }
    });
    SpreadsheetApp.flush();
    e.source.toast(
      'Checkbox edits do not approve payments. Select a student row, then use DIIT Ceremony > Confirm payment for selected student.',
      'Payment safety', 8);
  } finally {
    lock.releaseLock();
  }
}

// Private Sheet menu command, never callable by a public registration visitor.
// The payment checkbox, status and email change ONLY after the organizer clicks Yes.
function confirmPaymentForSelectedRow_() {
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
  const data = sheet.getRange(row, 1, 1, HEADERS_.length).getValues()[0];
  if (String(data[11]) === 'CONFIRMED' || Boolean(data[14])) {
    ui.alert('This registration is already confirmed or an email was already sent. Use the separate resend option if needed.');
    return;
  }
  if (String(data[11]) !== 'PENDING' || data[10] === true) {
    ui.alert('This registration is not in a normal pending state. Check columns K–P before proceeding.');
    return;
  }

  const details = [
    'Student: ' + String(data[2]),
    'Student ID: ' + String(data[3]),
    'Reference: ' + String(data[1]),
    'Email: ' + String(data[4]),
    'Amount: BDT ' + String(data[7]),
    'Payment method: ' + String(data[8]),
    'Transaction ID: ' + (String(data[9] || '') || '(none)'),
    '',
    'Have you VERIFIED receiving the correct payment?',
    'YES will tick Payment received, confirm the student and email the PDF ticket.',
    'NO will leave everything unchanged.'
  ].join('\n');
  const answer = ui.alert('Confirm payment and send ticket?', details, ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  let emailError = '';
  try {
    // Re-check after the dialog: the sheet might have changed while it was open.
    const current = sheet.getRange(row, 1, 1, HEADERS_.length).getValues()[0];
    if (String(current[1]) !== String(data[1]) ||
        String(current[11]) !== 'PENDING' ||
        current[10] === true || Boolean(current[14])) {
      ui.alert('The selected registration changed or has already been processed. No email was sent. Refresh and check this row.');
      return;
    }
    const token = clean_(current[12]) ||
      Utilities.getUuid().replace(/-/g, '') +
      Utilities.getUuid().replace(/-/g, '');
    sheet.getRange(row, 13).setValue(token);        // M: unique ticket token
    sheet.getRange(row, 14).setValue(new Date());    // N: confirmed at
    sheet.getRange(row, 12).setValue('CONFIRMED');   // L: status
    sheet.getRange(row, 11).setValue(true);          // K: payment received
    sheet.getRange(row, 16).clearContent();         // P: clear previous email error
    SpreadsheetApp.flush();

    // Script edits do not fire onEdit triggers. Email is sent exactly once here.
    try {
      sendTicketEmail_(current, token);
      sheet.getRange(row, 15).setValue(new Date()); // O: email sent at
    } catch (err) {
      emailError = String(err).slice(0, 500);
      sheet.getRange(row, 16).setValue(emailError); // P: email error
    }
  } finally {
    lock.releaseLock();
  }
  if (emailError) {
    ui.alert('Payment CONFIRMED, but the ticket email failed.\n\n' + emailError +
      '\n\nCheck column P, then use DIIT Ceremony > Resend confirmation for selected student.');
  } else {
    ui.alert('Payment confirmed. The checkbox is checked, status is CONFIRMED, and the ticket email was sent. Check column O.');
  }
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
      '\\nThis sends a new email with the personalized ticket PDF and verification link.',
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
 * Build a one-page ticket PDF from the owner's editable Google Slides template.
 * The slide must have a {{NAME}} text placeholder and a page element whose
 * Alt text TITLE is QR_PLACEHOLDER. QR positioning comes from that element,
 * not hard-coded pixel coordinates. Never edit the original template.
 * The temporary copy is deleted after exporting the PDF.
 */
// Run once from the Apps Script editor to approve Slides/Drive permissions
// and inspect the ticket layout. Creates only one preview PDF in your Drive,
// with a dummy (invalid) QR token; it does not email or alter registrations.
function testTicketPdf_() {
  const sample = createPersonalizedTicketPdf_(
    'Sample Participant', 'PREVIEW',
    TICKET_VERIFICATION_BASE_URL_ + '?ticket=' + '0'.repeat(64));
  const file = DriveApp.createFile(sample);
  const url = file.getUrl();
  console.log('Ticket preview saved in Drive: ' + url);
  // Safe to use from the private Google Sheets menu; keep this function private.
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(
      '<div style="font:15px Arial,sans-serif;padding:16px;line-height:1.6">' +
      '<p>Your sample ceremony ticket PDF was saved to Google Drive.</p>' +
      '<p><a href="' + escape_(url) + '" target="_blank" rel="noopener noreferrer">' +
      'Open ticket preview PDF ↗</a></p></div>'
    ).setWidth(430).setHeight(175),
    'Ticket preview ready'
  );
  return url;
}

function createPersonalizedTicketPdf_(studentName, reference, ticketUrl) {
  // Only the name and QR are personalized. All decorative layout is maintained
  // in the original editable Slides template.
  const qrResponse = UrlFetchApp.fetch(
    'https://quickchart.io/qr?format=png&size=800&margin=2&text=' +
      encodeURIComponent(ticketUrl),
    {muteHttpExceptions:true});
  if (qrResponse.getResponseCode() !== 200 ||
      !/image\/png/i.test(String(qrResponse.getHeaders()['Content-Type'] ||
        qrResponse.getHeaders()['content-type'] || ''))) {
    throw Error('Ticket QR generation failed (HTTP ' +
      qrResponse.getResponseCode() + ').');
  }
  const qrBlob = qrResponse.getBlob().setName('ticket-qr.png');

  let copyId = null;
  try {
    // Diagnose access to the ORIGINAL template separately from opening its copy.
    // A Sheet owner and a Slides owner can be different Google accounts.
    let template;
    try {
      template = DriveApp.getFileById(TICKET_SLIDES_TEMPLATE_ID_);
      if (template.getMimeType() !== MimeType.GOOGLE_SLIDES) {
        throw Error('The template ID does not refer to a Google Slides presentation.');
      }
    } catch (err) {
      throw Error('Cannot access the Google Slides ticket template. Share the template ' +
        'with the Google account running this registration Sheet (Editor access), ' +
        'or make a copy in that account and update TICKET_SLIDES_TEMPLATE_ID_. ' +
        'Original error: ' + err);
    }

    let copy;
    try {
      copy = template.makeCopy('Temporary DIIT CSE21 ticket - ' + reference);
    } catch (err) {
      throw Error('The ticket template was found, but Apps Script could not copy it. ' +
        'Check that copying is allowed and your account has sufficient access. ' +
        'Original error: ' + err);
    }
    copyId = copy.getId();

    // Google may take a few seconds to index a newly copied presentation.
    let deck = null, openError = null;
    for (let attempt = 0; attempt < 4; attempt++) {
      try {
        deck = SlidesApp.openById(copyId);
        break;
      } catch (err) {
        openError = err;
        if (attempt < 3) Utilities.sleep(1500);
      }
    }
    if (!deck) {
      throw Error('A temporary copy was created, but Google Slides could not open it. ' +
        'Check that Google Slides is enabled for this account. ' +
        'Original error: ' + openError);
    }
    const slides = deck.getSlides();
    if (slides.length !== 1) {
      throw Error('Ticket template must contain exactly one slide.');
    }
    const slide = slides[0];
    const replacedNames = slide.replaceAllText('{{NAME}}', String(studentName));
    if (replacedNames !== 1) {
      throw Error('Ticket template must contain exactly one editable {{NAME}} placeholder. Found: ' +
        replacedNames);
    }

    // Alt text TITLE = QR_PLACEHOLDER on a shape/image. Use its exact geometry.
    const qrPlaceholders = slide.getPageElements()
      .filter(el => el.getTitle() === 'QR_PLACEHOLDER');
    if (qrPlaceholders.length !== 1) {
      throw Error('Ticket template must have exactly one element with Alt text Title ' +
        'QR_PLACEHOLDER. Found: ' + qrPlaceholders.length);
    }
    const place = qrPlaceholders[0];
    const side = Math.min(place.getWidth(), place.getHeight());
    if (side <= 0) throw Error('QR placeholder must have a visible square area.');
    const left = place.getLeft() + (place.getWidth() - side) / 2;
    const top = place.getTop() + (place.getHeight() - side) / 2;

    place.remove();
    slide.insertImage(qrBlob, left, top, side, side);
    deck.saveAndClose();

    return DriveApp.getFileById(copyId).getAs(MimeType.PDF)
      .setName('DIIT-CSE21-' + reference + '-Ticket.pdf');
  } finally {
    if (copyId) {
      try { DriveApp.getFileById(copyId).setTrashed(true); }
      catch (cleanupError) {
        console.log('Temporary ticket Slides cleanup failed: ' + cleanupError);
      }
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
