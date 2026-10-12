/**
 * Life app – ticks sync (Google Apps Script web app).
 * Paste into a Google Sheet: Extensions → Apps Script → replace Code.gs → Save.
 * Run setup once, Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 * Set Script property EXEC_URL to the /exec URL, SITE_URL to the GitHub Pages URL.
 * Menu Life app → Phone link → open on phone once (pairs url|key into localStorage).
 *
 * GET  ?action=load&key=K              → current ticks JSON (+ history if history=1)
 * POST text/plain {key, daily, weekly, points} → save current week
 */
var TZ = 'Europe/London';
var TAB = '_ticks';
var HIST = '_ticks_history';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Life app')
    .addItem('Phone link', 'showLink')
    .addItem('New secret key', 'rotateKey')
    .addToUi();
}

function setup() {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('APP_KEY')) props.setProperty('APP_KEY', Utilities.getUuid().replace(/-/g, ''));
  ticksTab_(); histTab_();
  Logger.log('Key ready. Deploy as web app, set EXEC_URL, then use Life app → Phone link.');
}
function rotateKey() {
  PropertiesService.getScriptProperties().setProperty('APP_KEY', Utilities.getUuid().replace(/-/g, ''));
  showLink();
}
function showLink() {
  var key = PropertiesService.getScriptProperties().getProperty('APP_KEY');
  var execUrl = PropertiesService.getScriptProperties().getProperty('EXEC_URL');
  if (!execUrl) {
    SpreadsheetApp.getUi().alert('Add Script property EXEC_URL (web app URL ending in /exec) in Project Settings.');
    return;
  }
  var site = PropertiesService.getScriptProperties().getProperty('SITE_URL') || 'https://oliverrising96.github.io/life-app/';
  var link = site.replace(/\/$/, '') + '/#pair=' + encodeURIComponent(execUrl) + '|' + key;
  var html = HtmlService.createHtmlOutput(
    '<p style="font:14px sans-serif">Open once on your phone (it remembers):</p>' +
    '<textarea style="width:100%;height:120px">' + link + '</textarea>'
  ).setWidth(520).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, 'Life app phone link');
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function authed_(k) {
  var key = PropertiesService.getScriptProperties().getProperty('APP_KEY');
  return key && k === key;
}
function ticksTab_() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(TAB);
  if (!sh) {
    sh = ss.insertSheet(TAB);
    sh.getRange(1, 1, 1, 2).setValues([['key', 'json']]);
    sh.hideSheet();
  }
  return sh;
}
function histTab_() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(HIST);
  if (!sh) {
    sh = ss.insertSheet(HIST);
    sh.getRange(1, 1, 1, 3).setValues([['week', 'json', 'archivedAt']]);
    sh.hideSheet();
  }
  return sh;
}
function isoWeek_() {
  var d = new Date();
  // Europe/London date parts
  var s = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  var parts = s.split('-').map(Number);
  var dt = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  var day = dt.getUTCDay() || 7;
  dt.setUTCDate(dt.getUTCDate() + 4 - day);
  var yearStart = new Date(Date.UTC(dt.getUTCFullYear(), 0, 1));
  var week = Math.ceil((((dt - yearStart) / 86400000) + 1) / 7);
  return dt.getUTCFullYear() + '-W' + ('0' + week).slice(-2);
}
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function empty_(week, date) {
  return {
    date: date, week: week, daily: {}, weekly: {},
    points: { Career: 0, Fitness: 0, Learning: 0, Relationships: 0 },
    updatedAt: new Date().toISOString()
  };
}
function loadCurrent_() {
  var sh = ticksTab_(), last = sh.getLastRow();
  if (last < 2) return null;
  var vals = sh.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][0] === 'current' && vals[i][1]) {
      try { return JSON.parse(vals[i][1]); } catch (e) { return null; }
    }
  }
  return null;
}
function saveCurrent_(obj) {
  var sh = ticksTab_(), last = sh.getLastRow();
  var row = 0;
  if (last >= 2) {
    var keys = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) if (keys[i][0] === 'current') { row = i + 2; break; }
  }
  var json = JSON.stringify(obj);
  if (row) sh.getRange(row, 1, 1, 2).setValues([['current', json]]);
  else sh.appendRow(['current', json]);
}
function archive_(state) {
  if (!state || !state.week) return;
  var sh = histTab_();
  var hist = {
    week: state.week,
    pillars: {},
    archivedAt: new Date().toISOString()
  };
  ['Career', 'Fitness', 'Learning', 'Relationships'].forEach(function (p) {
    hist.pillars[p] = {
      weekly: (state.weekly && state.weekly[p]) || {},
      daily: (state.daily && state.daily[p]) || {},
      points: (state.points && state.points[p]) || 0
    };
  });
  // upsert by week
  var last = sh.getLastRow(), row = 0;
  if (last >= 2) {
    var weeks = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < weeks.length; i++) if (weeks[i][0] === state.week) { row = i + 2; break; }
  }
  if (row) sh.getRange(row, 1, 1, 3).setValues([[state.week, JSON.stringify(hist), hist.archivedAt]]);
  else sh.appendRow([state.week, JSON.stringify(hist), hist.archivedAt]);
}
function ensureCurrent_() {
  var week = isoWeek_(), date = today_();
  var cur = loadCurrent_() || empty_(week, date);
  if (cur.week && cur.week !== week) {
    archive_(cur);
    cur = empty_(week, date);
  }
  if (cur.date !== date) {
    cur.date = date;
    cur.daily = {};
    cur.updatedAt = new Date().toISOString();
  }
  if (!cur.week) cur.week = week;
  saveCurrent_(cur);
  return cur;
}
function loadHistory_(weeks) {
  var sh = histTab_(), last = sh.getLastRow(), rows = [];
  if (last < 2) return rows;
  var vals = sh.getRange(2, 1, last - 1, 2).getValues();
  vals.forEach(function (r) {
    if (!r[1]) return;
    try { rows.push(JSON.parse(r[1])); } catch (e) {}
  });
  rows.sort(function (a, b) { return (a.week || '').localeCompare(b.week || ''); });
  return rows.slice(-weeks);
}

function doGet(e) {
  e = e || { parameter: {} };
  if (!authed_(e.parameter.key)) return json_({ ok: false, error: 'bad key' });
  var cur = ensureCurrent_();
  if (e.parameter.history === '1' || e.parameter.action === 'history') {
    var n = Math.min(26, parseInt(e.parameter.weeks || '8', 10) || 8);
    return json_({ ok: true, history: loadHistory_(n), current: cur });
  }
  return json_(Object.assign({ ok: true }, cur));
}

function doPost(e) {
  var body;
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'bad json' }); }
  if (!authed_(body.key)) return json_({ ok: false, error: 'bad key' });
  var week = isoWeek_(), date = today_();
  var cur = ensureCurrent_();
  if (body.daily) cur.daily = body.daily;
  if (body.weekly) cur.weekly = body.weekly;
  if (body.points) cur.points = body.points;
  cur.week = week;
  cur.date = date;
  cur.updatedAt = new Date().toISOString();
  saveCurrent_(cur);
  return json_(Object.assign({ ok: true }, cur));
}
