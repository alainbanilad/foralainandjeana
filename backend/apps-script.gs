// ============================================================
// Alain + Jeana -- Combined Apps Script
// ============================================================
// PART 1: RSVP Email to Google Sheets (your existing script,
//         unchanged) -- runs hourly via trigger
// PART 2: Menu Selection endpoint (new) -- receives POSTs from
//         menu.html and writes to the "Menu Selections" tab
// ============================================================
// SETUP FOR PART 2:
// 1. Replace your project's code with this entire file
// 2. Deploy -> New deployment -> Web app
//    Execute as: Me | Who has access: Anyone
// 3. Copy the Web app URL into menu.html
// Your existing hourly trigger keeps working untouched --
// do NOT re-run createTrigger().
// ============================================================


// ============================================================
// PART 1 -- RSVP EMAIL PARSER (existing, unchanged)
// ============================================================

function parseRSVPs() {

  const SEARCH_QUERY    = 'subject:"New RSVP from"';
  const SHEET_NAME      = 'RSVPs';
  const PROCESSED_LABEL = 'rsvp-processed';

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'Timestamp', 'Name', 'Email', 'Attendance',
      'Event(s)', 'Dietary Requirements', 'Song Request', 'Message', 'Flag'
    ]);
    const header = sheet.getRange(1, 1, 1, 9);
    header.setFontWeight('bold');
    header.setBackground('#7A1C2E');
    header.setFontColor('#FAF0F2');
    sheet.setFrozenRows(1);
  }

  let label = GmailApp.getUserLabelByName(PROCESSED_LABEL);
  if (!label) label = GmailApp.createLabel(PROCESSED_LABEL);

  const threads = GmailApp.search(SEARCH_QUERY + ' -label:' + PROCESSED_LABEL);

  threads.forEach(thread => {
    const messages = thread.getMessages();

    messages.forEach(msg => {
      const body      = msg.getPlainBody();
      const timestamp = msg.getDate();

      const name       = extract(body, 'Name');
      const email      = extract(body, 'Email');
      const attendance = extract(body, 'Attendance');
      const events     = extract(body, 'Event(s) selected');
      const dietary    = extract(body, 'Dietary requirements');
      const song       = extract(body, 'Song requests');
      const message    = extract(body, 'Message');

      let flag = '';
      if (attendance === 'Not attending' && events !== '--' && events !== '') {
        flag = 'Not attending but ticked: ' + events;
      }
      if (attendance === 'Attending' && (events === '--' || events === '')) {
        flag = 'Attending but no event selected';
      }

      sheet.appendRow([timestamp, name, email, attendance, events, dietary, song, message, flag]);

      if (email && email !== '--') {
        sendReply(name, email, attendance, events, dietary, song);
      }
    });

    thread.addLabel(label);
  });

  sheet.autoResizeColumns(1, 9);
  Logger.log('Done! Processed ' + threads.length + ' new RSVP(s).');
}


function sendReply(name, guestEmail, attendance, events, dietary, song) {

  const firstName = name.split(' ')[0];
  const isAttending = attendance.toLowerCase().includes('attending') &&
                      !attendance.toLowerCase().includes('not');

  const hasROM    = events.indexOf('ROM') !== -1;
  const hasDinner = events.indexOf('Dinner') !== -1;

  let eventNames = '';
  if (hasROM && hasDinner) {
    eventNames = 'the ROM Solemnization and the Dinner Celebration';
  } else if (hasROM) {
    eventNames = 'the ROM Solemnization';
  } else if (hasDinner) {
    eventNames = 'the Dinner Celebration';
  } else {
    eventNames = 'us';
  }

  if (!isAttending) {
    const missedLine = (hasROM || hasDinner)
      ? "You'll be missed at " + eventNames + ".\n\n"
      : '';

    const body =
      'Hi ' + firstName + ',\n\n' +
      "Thank you so much for letting us know - we really appreciate it, and " +
      "we totally understand that schedules and distances don't always line " +
      'up.\n\n' +
      missedLine +
      "We're sad you can't make it, but please know you're in our hearts " +
      "either way. We'll be sure to share photos and stories from the " +
      "day(s) once it's all over (and maybe even some of the chaos, " +
      'knowing us).\n\n' +
      "If anything changes and you're able to join after all, just reply " +
      "to this email - there's always a seat and a glass of wine waiting " +
      'for you.\n\n' +
      'Thank you again for being part of our story, even from afar.\n\n' +
      'With love,\n' +
      'Alain + Jeana';

    GmailApp.sendEmail(
      guestEmail,
      "We'll miss you, " + firstName,
      body,
      { cc: 'alain.banilad@gmail.com,jeanathegreat@gmail.com' }
    );
    Logger.log('Not-attending reply sent to: ' + guestEmail);
    return;
  }

  let eventDetails = '';

  if (hasROM) {
    eventDetails +=
      '* ROM Solemnization\n' +
      'Monday, 11 August 2026 | 3:30 PM - 3:45 PM\n' +
      'Registry of Marriages | Commitment Room, The Esplanade Mall\n' +
      'Dress code: Formal - no colour restrictions, just come as your best dressed self\n';
  }

  if (hasROM && hasDinner) {
    eventDetails += '\n';
  }

  if (hasDinner) {
    eventDetails +=
      '* Dinner Celebration\n' +
      'Saturday, 15 August 2026 | 6:00 PM til late\n' +
      'Botanico at The Summerhouse, 3 Park Lane\n' +
      'Dress code: Semi-Formal / Cocktail - wine red, silvery grays & elegance welcome\n';
  }

  if (!hasROM && !hasDinner) {
    eventDetails = 'Event(s): ' + events + '\n';
  }

  const dietaryLine = (dietary && dietary !== '--') ? dietary : 'None';
  const songLine    = (song && song !== '--') ? song : '--';

  const bottleRequest = hasDinner
    ? "\nOne cheeky reminder: if you said you'd bring a bottle for the " +
      "dinner, don't forget it. We're counting on you to keep the night going!\n"
    : '';

  const divider = '----------------------------';

  const body =
    'Hi ' + firstName + ',\n\n' +
    'Wonderful news - we got your RSVP and we are SO ready!\n\n' +
    "You're joining us for " + eventNames + ". Here's a recap:\n\n" +
    divider + '\n' +
    'EVENT DETAILS\n' +
    divider + '\n\n' +
    eventDetails + '\n' +
    divider + '\n' +
    'YOUR DETAILS ON FILE\n' +
    divider + '\n' +
    'Dietary requirements: ' + dietaryLine + '\n' +
    'Song request: ' + songLine + '\n\n' +
    'If anything above looks off - wrong dietary info, name typo, change ' +
    "of plans - just reply to this email and we'll sort it out before " +
    'the big day(s).\n' +
    bottleRequest + '\n' +
    'We genuinely cannot wait to celebrate with you.\n\n' +
    'With love (and mild chaos),\n' +
    'Alain + Jeana';

  GmailApp.sendEmail(
    guestEmail,
    'See you there, ' + firstName + '!',
    body,
    { cc: 'alain.banilad@gmail.com,jeanathegreat@gmail.com' }
  );
  Logger.log('Attending reply sent to: ' + guestEmail + ' | Events: ' + events);
}


function extract(body, fieldName) {
  const escaped = fieldName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

  if (fieldName === 'Message') {
    const regex = new RegExp(escaped + ':\\s*([\\s\\S]*)', 'i');
    const match = body.match(regex);
    if (!match) return '--';
    const cleaned = match[1]
      .replace(/\s*Email sent via EmailJS\.com\s*\[?https?:\/\/[^\]]*\]?/gi, '')
      .trim();
    return cleaned || '--';
  }

  const regex = new RegExp(escaped + ':\\s*(.+)', 'i');
  const match = body.match(regex);
  return match ? match[1].trim() : '--';
}


function createTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('parseRSVPs').timeBased().everyHours(1).create();
  Logger.log('Trigger created! Script will run every hour.');
}


// ============================================================
// PART 2 -- MENU SELECTION ENDPOINT (new)
// Receives POSTs from menu.html, writes to "Menu Selections"
// ============================================================

var MENU_SHEET_NAME = 'Menu Selections';

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); // avoid two guests writing at the exact same moment

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(MENU_SHEET_NAME);

    // Create the tab with headers on first run (styled like your RSVPs tab)
    if (!sheet) {
      sheet = ss.insertSheet(MENU_SHEET_NAME);
      sheet.appendRow(['Timestamp', 'Name', 'Email', 'Main Course', 'Dietary Notes']);
      var header = sheet.getRange(1, 1, 1, 5);
      header.setFontWeight('bold');
      header.setBackground('#7A1C2E');
      header.setFontColor('#FAF0F2');
      sheet.setFrozenRows(1);
    }

    var p = e.parameter;

    // Basic sanity check so random hits don't create junk rows
    if (!p.name || !p.email || !p.main_course) {
      return jsonReply({ status: 'error', message: 'Missing required fields' });
    }

    sheet.appendRow([
      new Date(),
      p.name,
      p.email,
      p.main_course,
      p.dietary || ''
    ]);

    return jsonReply({ status: 'ok' });

  } catch (err) {
    return jsonReply({ status: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Visiting the web app URL in a browser gives a friendly response
// (also handy as a quick "is it deployed?" check)
function doGet() {
  return jsonReply({ status: 'ok', message: 'Menu selection endpoint is live' });
}

function jsonReply(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
