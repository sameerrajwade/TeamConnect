/**
 * @OnlyCurrentDoc
 */

// Paste the ID from your "Three60" Google Sheet URL: docs.google.com/spreadsheets/d/{THIS_PART}/edit
const SHEET_ID = 'YOUR_THREE60_SHEET_ID_HERE';
const NOTIFY_EMAIL = 'sameer.rajwade@capitalone.com';

function doGet(e) {
  const page = e && e.parameter && e.parameter.page;
  if (page === 'directory') {
    return HtmlService.createHtmlOutputFromFile('directory')
      .setTitle('Three60 — Full Directory')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'allwins') {
    return HtmlService.createHtmlOutputFromFile('allwins')
      .setTitle('Three60 — All Wins')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'pastevents') {
    return HtmlService.createHtmlOutputFromFile('pastevents')
      .setTitle('Three60 — Past Events')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Three60 — Three Teams · 360° Collaboration')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return {
    events:      getSheetData(ss, 'Events'),
    showAndTell: getSheetData(ss, 'ShowAndTell'),
    wins:        getSheetData(ss, 'Wins'),
    directory:   getSheetData(ss, 'Directory'),
    config:      getConfigData(ss),
    ratings:     getSheetData(ss, 'Ratings')
  };
}

function getSheetData(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1)
    .filter(r => r.some(c => c !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i] instanceof Date ? row[i].toISOString() : row[i];
      });
      return obj;
    });
}

function getConfigData(ss) {
  const sheet = ss.getSheetByName('Config');
  if (!sheet) return {};
  const rows = sheet.getDataRange().getValues();
  const config = {};
  rows.forEach(r => {
    if (r[0]) config[r[0]] = r[1];
  });
  return config;
}

function submitWin(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Wins');
  if (!sheet) return { success: false, error: 'Wins sheet not found' };
  sheet.appendRow([
    new Date(),
    data.name,
    data.team,
    data.description
  ]);
  return { success: true };
}

function submitShowAndTell(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('ShowAndTell');
  if (!sheet) return { success: false, error: 'ShowAndTell sheet not found' };
  sheet.appendRow([
    new Date(),
    data.name,
    data.team,
    data.topic,
    data.description,
    'Pending'
  ]);
  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: 'Three60 — New Show & Tell Submission',
    body: 'New S&T topic submitted:\n\n' +
          'Name: ' + data.name + '\n' +
          'Team: ' + data.team + '\n' +
          'Topic: ' + data.topic + '\n' +
          'Description: ' + data.description
  });
  return { success: true };
}

function submitRating(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Ratings');
  if (!sheet) return { success: false, error: 'Ratings sheet not found' };
  sheet.appendRow([
    new Date(),
    data.eventTitle,
    data.rating
  ]);
  return { success: true };
}

function submitBuzzPoll(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('BuzzResponses');
  if (!sheet) return { success: false, error: 'BuzzResponses sheet not found' };
  sheet.appendRow([
    new Date(),
    'Poll',
    data.question,
    data.answer
  ]);
  return { success: true };
}

function submitBuzzAnswer(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('BuzzResponses');
  if (!sheet) return { success: false, error: 'BuzzResponses sheet not found' };
  sheet.appendRow([
    new Date(),
    'Question',
    data.question,
    data.answer
  ]);
  return { success: true };
}

function setupEmailReminder() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'checkAndSendReminders') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('checkAndSendReminders')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
}

function checkAndSendReminders() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const events = getSheetData(ss, 'Events');
  const now = new Date();
  const sevenDays = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const targetDay = new Date(sevenDays.getFullYear(), sevenDays.getMonth(), sevenDays.getDate());

  events.forEach(evt => {
    if (!evt.PlannedDate) return;
    const evtDate = new Date(evt.PlannedDate);
    const evtDay = new Date(evtDate.getFullYear(), evtDate.getMonth(), evtDate.getDate());
    if (evtDay.getTime() === targetDay.getTime()) {
      const st = events.length > 0 ? getSheetData(ss, 'ShowAndTell').find(
        s => s.EventTitle === evt.Title && s.Status === 'Approved'
      ) : null;

      let body = 'Reminder: Three60 Connect is in 7 days!\n\n' +
                 'Event: ' + evt.Title + '\n' +
                 'Date: ' + evtDate.toDateString() + '\n' +
                 'Time: ' + (evt.Time || '2:00 PM ET') + '\n' +
                 'Host Team: ' + (evt.HostTeam || '') + '\n\n' +
                 'Agenda:\n' +
                 '  30 min — ' + (evt.Activity || 'Team Activity') + '\n' +
                 '  30 min — Show & Tell' + (st ? ' by ' + st.PresenterName + ': "' + st.Topic + '"' : '') + '\n' +
                 '  30 min — Open Discussion\n\n' +
                 (evt.MeetLink ? 'Join: ' + evt.MeetLink + '\n\n' : '') +
                 'See you there!';

      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: 'Three60 Reminder — ' + evt.Title + ' in 7 days',
        body: body
      });
    }
  });
}
