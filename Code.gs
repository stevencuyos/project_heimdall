function doGet(e) {
  const t = HtmlService.createTemplateFromFile('index');
  const preview = String((e && e.parameter && e.parameter.as) || '').toLowerCase().replace(/[^a-z0-9._-]/g, '');
  t.preview = preview;

  // Server-side fast bootstrap: pre-fetch agent data so client renders immediately on first frame
  let bootstrap = null;
  try {
    bootstrap = getAgentData(preview);
  } catch (err) {
    bootstrap = null;
  }
  t.bootstrap = bootstrap;

  return t.evaluate()
    .setTitle('go/myplayschedule')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getMatrix() {
  const ss = SpreadsheetApp.openById("1YnQun0_ANNmteT2anXtOD_VIYk3X_gihHMGPPdh3OSc");
  const sheet = ss.getSheetByName("Matrix Play-April 2026");

  const data = sheet.getRange("E17:AM36").getValues();
  return data;
}

/**
 * ------------------------------------------------------------------
 * GO/MYPLAYSCHEDULE - CONFIG
 * ------------------------------------------------------------------
 */
const HOME_ID = '1YnQun0_ANNmteT2anXtOD_VIYk3X_gihHMGPPdh3OSc';
const MPS_TZ = 'Asia/Manila';
const SHIFT_HOURS = 9;                 // 8h paid + 1h unpaid lunch
const SRC_TAB = 'Schedules';           // tab name inside each Overall Schedule file
const ADMINS = ['stevenjosephc'];      // can preview any agent via ?as=ldap
const SH_SOURCES = 'Schedule Sources'; // links in B3 down, status written to C
const SH_AGENTS = 'Agents List';       // LDAPs in B3 down
const SH_SHIFTS = 'Shifts';
const SH_TAGS = 'Tags';
const SH_ROLES = 'Roles';
const SH_TAGLOG = 'TagLog';
const SH_CORR = 'Corrections';
const SH_CLOG = 'CorrLog';
const SH_ARCH = 'ShiftsArchive';
const SH_ERR = 'ErrorLog';
const SH_COBRA = 'Cobra Attendance';
const SH_PROD = 'Productivity';
const SH_YTD = 'YTD-ATT';
const OT_FORM_URL = 'https://script.google.com/a/macros/google.com/s/AKfycbyvgg3UiYFPYsQS1rtYYLk5ooCByzOzj-nP7KjJpiOKt4YXOSujZCCIhWBRU9p6te_vyw/exec';
const SHIFT_HEAD = ['Date', 'LDAP', 'EmpID', 'Name', 'Team', 'Supervisor', 'Channel', 'Code', 'Prev', 'ChangedAt', 'Synced'];
const TAG_HEAD = ['Date', 'LDAP', 'Tag', 'By', 'At', 'Note'];

let SHIFTS_CHANGED_ = false;
let HOME_MEMO_ = null, ROS_MEMO_ = null, ROLES_MEMO_ = null; // reset on every execution
function home_() { return HOME_MEMO_ || (HOME_MEMO_ = SpreadsheetApp.openById(HOME_ID)); }

function roles_() {
  if (ROLES_MEMO_) return ROLES_MEMO_;
  const out = {};
  const sh = home_().getSheetByName(SH_ROLES);
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(row => {
      const l = String(row[0]).trim().toLowerCase();
      if (!l || out[l]) return;
      out[l] = { role: String(row[1]).trim().toUpperCase(), teams: String(row[2]).split(',').map(s => s.trim()).filter(Boolean) };
    });
  }
  return (ROLES_MEMO_ = out);
}

// ADMINS in code is now only a lock-out safety net. Add admins in the Roles tab: LDAP | ADMIN | *
function isAdmin_(ldap) {
  if (!ldap) return false;
  if (ADMINS.indexOf(ldap) > -1) return true;
  const r = roles_()[ldap];
  return !!(r && r.role === 'ADMIN');
}

function ds_(v) {
  return v instanceof Date ? Utilities.formatDate(v, MPS_TZ, 'yyyy-MM-dd') : String(v);
}

/**
 * ------------------------------------------------------------------
 * SETUP (run once) - creates Shifts + Tags tabs, 15-min sync trigger
 * ------------------------------------------------------------------
 */
function setupScheduleApp() {
  const ss = home_();
  [[SH_SHIFTS, SHIFT_HEAD], [SH_TAGS, TAG_HEAD]].forEach(([name, head]) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), head.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncSchedules')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncSchedules').timeBased().everyMinutes(15).create();
  syncSchedules();
}

/**
 * ------------------------------------------------------------------
 * INGEST - Schedule Sources -> Shifts
 * ------------------------------------------------------------------
 */
function syncSchedules() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = home_();
    const allow = readAgentList_(ss);
    if (!Object.keys(allow).length) throw new Error('Agents List is empty');

    const srcSheet = ss.getSheetByName(SH_SOURCES);
    const stamp = Utilities.formatDate(new Date(), MPS_TZ, 'yyyy-MM-dd HH:mm');
    const incoming = [];
    let okCount = 0, errCount = 0;

    const last = srcSheet.getLastRow();
    if (last >= 3) {
      const numRows = last - 2;
      const srcVals = srcSheet.getRange(3, 2, numRows, 2).getValues();
      const rtvs = srcSheet.getRange(3, 2, numRows, 1).getRichTextValues();
      const statusCol = [];

      for (let i = 0; i < numRows; i++) {
        const r = i + 3;
        const valB = String(srcVals[i][0]).trim();
        if (!valB) {
          statusCol.push([srcVals[i][1]]);
          continue;
        }
        let msg;
        try {
          const cell = srcSheet.getRange(r, 2);
          const fileId = sourceId_(cell, srcVals[i][0], rtvs[i][0]);
          if (!fileId) throw new Error('No link found. Paste the plain link in this cell.');
          const res = readSource_(fileId, allow);
          res.rows.forEach(x => incoming.push(x));
          okCount++;
          msg = 'OK · ' + res.agents + ' agents · ' + res.days + ' days · ' + stamp;
        } catch (err) {
          msg = 'ERROR · ' + err.message; errCount++;
        }
        statusCol.push([msg]);
      }

      srcSheet.getRange(3, 3, numRows, 1).setValues(statusCol);
    }

    PropertiesService.getScriptProperties().setProperty('syncErrs', String(errCount));
    if (okCount) {
      upsertShifts_(ss, incoming, stamp);
      const props = PropertiesService.getScriptProperties();
      props.setProperty('syncedAt', Utilities.formatDate(new Date(), MPS_TZ, "yyyy-MM-dd'T'HH:mm:ssXXX"));
      if (SHIFTS_CHANGED_) props.setProperty('sver', String(Date.now()));
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * ------------------------------------------------------------------
 * TEST SUITE - Manual integration test for syncSchedules()
 * ------------------------------------------------------------------
 */
function testSyncSchedules() {
  console.log('Starting testSyncSchedules()...');
  const ss = home_();
  const srcSheet = ss.getSheetByName(SH_SOURCES);
  if (!srcSheet) {
    throw new Error(`Sheet "${SH_SOURCES}" not found.`);
  }

  const lastRow = srcSheet.getLastRow();
  const configuredSources = lastRow >= 3 ? lastRow - 2 : 0;
  console.log(`Configured sources count: ${configuredSources}`);

  // Execute synchronization
  syncSchedules();

  const props = PropertiesService.getScriptProperties();
  const syncedAt = props.getProperty('syncedAt');
  const syncErrs = props.getProperty('syncErrs');

  console.log(`Sync test completed. Timestamp: ${syncedAt}, Error count: ${syncErrs}`);

  if (Number(syncErrs) > 0) {
    console.warn('Sync finished with errors on one or more sources. Check the Status column in Schedule Sources.');
  } else {
    console.log('TEST SUCCESS: syncSchedules() executed with zero errors.');
  }
}

function chipUrls_(sheet, r, c) {
  try {
    const a1 = "'" + sheet.getName().replace(/'/g, "''") + "'!" + sheet.getRange(r, c).getA1Notation();
    const res = Sheets.Spreadsheets.get(sheet.getParent().getId(), {
      ranges: [a1],
      fields: 'sheets(data(rowData(values(chipRuns))))'
    });
    const out = [];
    res.sheets.forEach(s => (s.data || []).forEach(d => (d.rowData || []).forEach(rd => (rd.values || []).forEach(v => (v.chipRuns || []).forEach(cr => {
      if (cr.chip && cr.chip.richLinkProperties) out.push(cr.chip.richLinkProperties.uri);
    })))));
    return out;
  } catch (e) {
    return [];
  }
}

function readAgentList_(ss) {
  const sh = ss.getSheetByName(SH_AGENTS);
  const set = {};
  if (!sh || sh.getLastRow() < 3) return set;
  sh.getRange(3, 2, sh.getLastRow() - 2, 1).getValues().forEach(r => {
    const v = String(r[0]).trim().toLowerCase();
    if (v) set[v] = true;
  });
  return set;
}

function sourceId_(cell, valOverride, rtvOverride) {
  const urls = [];
  const rtv = rtvOverride || cell.getRichTextValue();
  if (rtv) {
    urls.push(rtv.getLinkUrl());
    rtv.getRuns().forEach(r => urls.push(r.getLinkUrl()));
  }
  chipUrls_(cell.getSheet(), cell.getRow(), cell.getColumn()).forEach(u => urls.push(u));
  urls.push(String(valOverride !== undefined ? valOverride : cell.getValue()));
  for (let i = 0; i < urls.length; i++) {
    const u = urls[i];
    if (!u) continue;
    const m = String(u).match(/\/d\/([a-zA-Z0-9_-]{20,})/) || String(u).match(/^([a-zA-Z0-9_-]{25,})$/);
    if (m) return m[1];
  }
  return '';
}

function readSource_(fileId, allow) {
  const src = SpreadsheetApp.openById(fileId);
  const sh = src.getSheetByName(SRC_TAB);
  if (!sh) throw new Error('Tab "' + SRC_TAB + '" not found');
  const vals = sh.getDataRange().getDisplayValues();

  let hr = -1;
  for (let i = 0; i < Math.min(vals.length, 10); i++) {
    if (vals[i].some(c => String(c).trim().toUpperCase() === 'LDAP')) { hr = i; break; }
  }
  if (hr < 0) throw new Error('Header row with LDAP not found');

  const head = vals[hr].map(c => String(c).trim().toLowerCase());
  const cLdap = head.indexOf('ldap');
  const cEmp = head.indexOf('emp id');
  const cName = head.indexOf('complete name');
  const cSup = head.indexOf('immediate supervisor');
  const cTeam = head.indexOf('team name');
  const cChan = head.indexOf('channel');

  const now = new Date();
  const dateCols = [];
  vals[hr].forEach((h, j) => {
    const d = parseHeaderDate_(h, now);
    if (d) dateCols.push({ j: j, d: d });
  });
  if (!dateCols.length) throw new Error('No date columns found in header row');

  const rows = [];
  let agents = 0;
  for (let i = hr + 1; i < vals.length; i++) {
    const ldap = String(vals[i][cLdap]).trim().toLowerCase();
    if (!ldap || !allow[ldap]) continue;
    agents++;
    const meta = [
      String(vals[i][cEmp] || '').trim(),
      String(vals[i][cName] || '').trim(),
      String(vals[i][cTeam] || '').trim(),
      String(vals[i][cSup] || '').trim(),
      String(vals[i][cChan] || '').trim()
    ];
    dateCols.forEach(dc => {
      const code = normCode_(vals[i][dc.j]);
      if (code) rows.push([dc.d, ldap, meta[0], meta[1], meta[2], meta[3], meta[4], code]);
    });
  }
  return { rows: rows, agents: agents, days: dateCols.length };
}

function parseHeaderDate_(s, now) {
  s = String(s).trim();
  let mo, d, y = null, m = s.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
  if (m) {
    mo = +m[1]; d = +m[2];
    if (m[3]) y = +m[3] < 100 ? 2000 + +m[3] : +m[3];
  } else {
    m = s.match(/^([A-Za-z]{3})[a-z]*[-\s](\d{1,2})$/);
    if (!m) return null;
    mo = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'].indexOf(m[1].toLowerCase()) + 1;
    d = +m[2];
  }
  if (!mo || mo > 12 || d < 1 || d > 31) return null;
  if (!y) y = nearestYear_(mo, d, now);
  return y + '-' + ('0' + mo).slice(-2) + '-' + ('0' + d).slice(-2);
}

function nearestYear_(mo, d, now) {
  const cy = +Utilities.formatDate(now, MPS_TZ, 'yyyy');
  let best = cy, bd = Infinity;
  [cy - 1, cy, cy + 1].forEach(y => {
    const diff = Math.abs(new Date(y, mo - 1, d).getTime() - now.getTime());
    if (diff < bd) { bd = diff; best = y; }
  });
  return best;
}

function normCode_(v) {
  v = String(v).trim().toUpperCase();
  if (!v || v === '-') return '';
  const m = v.match(/^(\d{1,2}):(\d{2})/);
  if (m) return ('0' + m[1]).slice(-2) + ':' + m[2];
  return v;
}

function upsertShifts_(ss, incoming, stamp) {
  const sh = ss.getSheetByName(SH_SHIFTS) || ss.insertSheet(SH_SHIFTS);
  const n = SHIFT_HEAD.length;
  const last = sh.getLastRow();
  const old = last > 1 ? sh.getRange(2, 1, last - 1, n).getValues() : [];
  const idx = {};
  old.forEach((r, i) => {
    r[0] = ds_(r[0]);
    idx[r[0] + '|' + String(r[1]).toLowerCase()] = i;
  });

  const keepFrom = keepFrom_();
  incoming.forEach(x => {
    if (x[0] < keepFrom) return; // don't re-add archived days
    const k = x[0] + '|' + x[1];
    if (k in idx) {
      const r = old[idx[k]];
      if (String(r[7]) !== x[7]) { r[8] = r[7]; r[9] = stamp; r[7] = x[7]; SHIFTS_CHANGED_ = true; }
      if (String(r[3]) !== x[3] || String(r[4]) !== x[4] || String(r[5]) !== x[5] || String(r[6]) !== x[6]) SHIFTS_CHANGED_ = true; r[2] = x[2]; r[3] = x[3]; r[4] = x[4]; r[5] = x[5]; r[6] = x[6]; r[10] = stamp;
    } else {
      idx[k] = old.length; SHIFTS_CHANGED_ = true;
      old.push([x[0], x[1], x[2], x[3], x[4], x[5], x[6], x[7], '', '', stamp]);
    }
  });

  if (sh.getMaxRows() < old.length + 1) {
    sh.insertRowsAfter(sh.getMaxRows(), old.length + 1 - sh.getMaxRows());
  }
  const rng = sh.getRange(2, 1, old.length, n);
  rng.setNumberFormat('@');
  rng.setValues(old);
}

/**
 * ------------------------------------------------------------------
 * AGENT VIEW - data endpoint
 * ------------------------------------------------------------------
 */
function getAgentData(asLdap) {
  const email = Session.getActiveUser().getEmail();
  const self = email ? email.split('@')[0].toLowerCase() : '';
  const as = String(asLdap || '').toLowerCase().replace(/[^a-z0-9._-]/g, '');
  const isAdmin = isAdmin_(self);
  const ldap = (as && isAdmin) ? as : self;
  if (!ldap) return { ok: false, reason: 'no_identity' };

  const props = PropertiesService.getScriptProperties();
  const key = 'agent5:' + ldap + ':' + (props.getProperty('sver') || '0') + ':' + rosterSig_();
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  let out;
  if (hit) {
    out = JSON.parse(hit);
  } else {
    out = buildAgent_(home_(), ldap, props.getProperty('syncedAt') || '');
    try { cache.put(key, JSON.stringify(out), 21600); } catch (e) {}
  }
    out.preview = !!(as && isAdmin && as !== self);
  if (out.ok) { out.tags = tagsOf_(ldap); out.syncedAt = props.getProperty('syncedAt') || out.syncedAt || ''; }
  const rl = roleFor_(self);
  if (!out.ok && rl && ldap === self) out = { ok: true, ldap: ldap, me: { ldap: ldap, name: ldap, emp: (roster_()[ldap] || {}).emp || '' }, cfg: { shiftHours: SHIFT_HOURS }, shifts: {}, tags: {}, people: [], team: {}, syncedAt: out.syncedAt || '', preview: false };
  // Hydrate with cached weather if available so client has weather without waiting
  try {
    const wxHit = cache.get('wx2');
    out.weather = wxHit ? JSON.parse(wxHit) : {};
  } catch (e) {
    out.weather = {};
  }
  // Hydrate with cached OT and Perf if available in cache
  try {
    const otHit = cache.get('ot1:' + ldap);
    if (otHit) {
      const otParsed = JSON.parse(otHit);
      out.cachedOt = otParsed.items || [];
    }
  } catch (e) {}
  try {
    const pfHit = cache.get('pf1:' + ldap);
    if (pfHit) {
      out.cachedPerf = JSON.parse(pfHit);
    }
  } catch (e) {}
  out.syncErrs = isAdmin_(self) ? +(props.getProperty('syncErrs') || 0) : 0;
  out.corr = corrFor_(ldap);
  out.crN = rl ? corrOpenCount_(rl) : 0;
  out.otUrl = OT_FORM_URL.indexOf('http') === 0 ? OT_FORM_URL : '';
  out.role = rl ? { all: rl.all, teams: rl.teams } : null;
  return out;
}

function weatherMap_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('wx2');
  if (hit) return JSON.parse(hit);
  const map = {};
  try {
    const url = 'https://api.open-meteo.com/v1/forecast?latitude=10.3157&longitude=123.8854&daily=weathercode,temperature_2m_max,temperature_2m_min,precipitation_probability_max&timezone=Asia%2FManila&past_days=45&forecast_days=16';
    const data = JSON.parse(UrlFetchApp.fetch(url).getContentText());
    data.daily.time.forEach((d, i) => {
      map[d] = [
        Math.round(data.daily.temperature_2m_max[i]),
        data.daily.weathercode[i],
        Math.round(data.daily.temperature_2m_min[i]),
        data.daily.precipitation_probability_max[i] || 0
      ];
    });
    cache.put('wx2', JSON.stringify(map), 10800);
  } catch (e) {
    cache.put('wx2', '{}', 300);
  }
  return map;
}

function buildAgent_(ss, ldap, syncedAt) {
  const out = { ok: true, ldap: ldap, me: { ldap: ldap }, cfg: { shiftHours: SHIFT_HOURS }, shifts: {}, tags: {}, syncedAt: syncedAt };
  const sh = ss.getSheetByName(SH_SHIFTS);
  if (!sh || sh.getLastRow() < 2) { out.ok = false; out.reason = 'no_data'; return out; }

  const all = sh.getRange(2, 1, sh.getLastRow() - 1, SHIFT_HEAD.length).getValues();
  const ros = roster_();
  const ft = {};
  all.forEach(r => { ft[String(r[1]).toLowerCase()] = String(r[4]); });
  const teamOf = l => (ros[l] && ros[l].team) || ft[l] || '';
  let team = '';
  all.forEach(r => {
    if (String(r[1]).toLowerCase() !== ldap) return;
    out.shifts[ds_(r[0])] = { c: String(r[7]), p: String(r[8]), u: String(r[9]) };
    out.me.name = String(r[3]); out.me.team = teamOf(ldap); out.me.sup = (ros[ldap] && ros[ldap].sup) || String(r[5]); out.me.ch = String(r[6]); out.me.emp = (ros[ldap] && ros[ldap].emp) || String(r[2]);
    team = teamOf(ldap);
  });
  out.people = [];
  out.team = {};
  if (team) { // teammates = same Team Name; compare r[5] (Supervisor) instead to group by supervisor
    const from = Utilities.formatDate(new Date(Date.now() - 92 * 864e5), MPS_TZ, 'yyyy-MM-dd');
    const pidx = {};
    all.forEach(r => {
      const l = String(r[1]).toLowerCase();
      if (l === ldap || tkey_(teamOf(l)) !== tkey_(team)) return;
      const d = ds_(r[0]);
      if (d < from) return;
      if (!(l in pidx)) { pidx[l] = out.people.length; out.people.push({ l: l, n: String(r[3]), ch: String(r[6]) }); }
      (out.team[d] = out.team[d] || []).push([pidx[l], String(r[7])]);
    });
  }
  if (!out.me.name) { out.ok = false; out.reason = 'not_found'; return out; }

  const tg = ss.getSheetByName(SH_TAGS);
  if (tg && tg.getLastRow() > 1) {
    tg.getRange(2, 1, tg.getLastRow() - 1, TAG_HEAD.length).getValues().forEach(r => {
      if (String(r[1]).toLowerCase() !== ldap) return;
      out.tags[ds_(r[0])] = { t: String(r[2]).toUpperCase(), by: String(r[3]), at: String(r[4]), n: String(r[5]) };
    });
  }
  return out;
}

/**
 * ------------------------------------------------------------------
 * ESTART CORRECTIONS - AUTO SIGNATURE
 * ------------------------------------------------------------------
 */
function onStatusChange(e) {
  if (!e) return;

  const range = e.range;
  const sheet = range.getSheet();
  const name = sheet.getName();

  if (name !== "Estart Corrections" && name !== SH_CORR) return;

  const col = range.getColumn();
  const row = range.getRow();

  if (col === 12 && row >= 3) {
    const timestamp = Utilities.formatDate(new Date(), "Asia/Manila", "MMM d, HH:mm");
    const email = Session.getActiveUser().getEmail();
    
    let signature = timestamp;

    if (email) {
      const username = email.split("@")[0];
      const name = username.charAt(0).toUpperCase() + username.slice(1).toLowerCase();
      signature = `${name} - ${timestamp}`;
    }

    range.offset(0, 1).setValue(signature);
  }
}

/**
 * ------------------------------------------------------------------
 * WEATHER FORECAST AUTOMATION (ROW 12 - CEBU CITY, AUTO-MONTH)
 * ------------------------------------------------------------------
 */

// 1. Adds a custom menu at the top of your Google Sheet
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⛅ Weather Tool')
    .addItem('Update Row 12 Weather (Cebu City)', 'updateWeatherForecast')
    .addItem('Sync Schedules Now', 'syncSchedules')
    .addToUi();
}

// 2. Automatically detects the current month's sheet, reads Row 16 dates, and writes to Row 12
function updateWeatherForecast() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Dynamically generate current month's tab name (e.g., "Matrix Play-November 2026")
  const now = new Date();
  const currentMonthYear = Utilities.formatDate(now, "Asia/Manila", "MMMM yyyy");
  const targetSheetName = `Matrix Play-${currentMonthYear}`;
  const currentYear = parseInt(Utilities.formatDate(now, "Asia/Manila", "yyyy"), 10);
  
  const sheet = ss.getSheetByName(targetSheetName);
  
  if (!sheet) {
    Logger.log(`Sheet '${targetSheetName}' not found!`);
    return;
  }

  // Coordinates for Cebu City
  const lat = 10.3157;
  const lon = 123.8854;
  
  // Fetch up to 16 days of forecast + past 7 days
  const url = `https://api.open-meteo.com/v1/forecast?latitude=${lat}&longitude=${lon}&daily=weathercode,temperature_2m_max,temperature_2m_min&timezone=Asia%2FManila&past_days=7&forecast_days=16`;
  
  const response = UrlFetchApp.fetch(url);
  const data = JSON.parse(response.getContentText());
  
  // Build lookup dictionary: { "YYYY-MM-DD": "🌧️ 31°C" }
  const weatherMap = {};
  for (let i = 0; i < data.daily.time.length; i++) {
    const dateStr = data.daily.time[i];
    const maxTemp = Math.round(data.daily.temperature_2m_max[i]);
    const code = data.daily.weathercode[i];
    const icon = getWeatherIcon(code);
    
    weatherMap[dateStr] = `${icon} ${maxTemp}°C`;
  }

  // Configuration for Row 16 (Dates) and Row 12 (Weather Forecast)
  const dateRow = 16;    // Row 16 has Nov-1, Nov-2, etc.
  const weatherRow = 12; // Row 12 is the top Weather Forecast bar
  const startCol = 10;   // Column J
  const numCols = 31;    // Columns J to AN (31 days)
  
  const dateRowValues = sheet.getRange(dateRow, startCol, 1, numCols).getValues()[0];
  const weatherRowOutput = [];

  for (let i = 0; i < dateRowValues.length; i++) {
    const cellVal = dateRowValues[i];
    let formattedDate = null;

    if (cellVal instanceof Date) {
      const d = new Date(cellVal);
      d.setFullYear(currentYear);
      formattedDate = Utilities.formatDate(d, "Asia/Manila", "yyyy-MM-dd");
    } else if (typeof cellVal === "string" && cellVal.trim() !== "") {
      const parsed = new Date(`${cellVal} ${currentYear}`);
      if (!isNaN(parsed.getTime())) {
        formattedDate = Utilities.formatDate(parsed, "Asia/Manila", "yyyy-MM-dd");
      }
    }

    if (formattedDate && weatherMap[formattedDate]) {
      weatherRowOutput.push(weatherMap[formattedDate]);
    } else {
      weatherRowOutput.push("-");
    }
  }

  // Write results into Row 12 (J12:AN12)
  sheet.getRange(weatherRow, startCol, 1, numCols).setValues([weatherRowOutput]);
}

// Helper function to turn weather codes into compact emojis
function getWeatherIcon(code) {
  if (code === 0) return "☀️"; // Clear sky
  if (code <= 3) return "⛅";  // Partly cloudy
  if (code <= 48) return "🌫️"; // Foggy
  if (code <= 67) return "🌧️"; // Rain
  if (code <= 82) return "🌦️"; // Rain showers
  if (code >= 95) return "⛈️"; // Thunderstorm
  return "☁️";
}

/**
 * ------------------------------------------------------------------
 * SUPERVISOR VIEW
 * ------------------------------------------------------------------
 */
const TAG_CODES = ['DONE', 'LATE', 'OVL', 'DNC', 'ABS', 'NCNS', 'INACTIVE', 'SL', 'VL', 'MVL', 'MWL', 'LOA', 'LWOP', 'OFF', 'RDOT'];

function tagOk_(tag, code) {
  if (!tag) return true;
  if (tag === 'RDOT') return code === 'OFF';
  return /^\d{2}:\d{2}$/.test(code);
}
const ROLE_HEAD = ['LDAP', 'Role', 'Teams'];
const LOG_HEAD = ['Timestamp', 'By', 'Date', 'LDAP', 'Old', 'New', 'Note'];

function setupSupervisorView() {
  const ss = home_();
  [[SH_ROLES, ROLE_HEAD], [SH_TAGLOG, LOG_HEAD], [SH_TAGS, TAG_HEAD]].forEach(([name, head]) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), head.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  const rs = ss.getSheetByName(SH_ROLES);
  rs.getRange('B1').setNote('SUPERVISOR = can tag only the Team Names listed in column C (comma-separated). MANAGER or ADMIN = all teams.');
  if (rs.getLastRow() < 2) {
    rs.getRange(2, 1, ADMINS.length, 3).setValues(ADMINS.map(a => [a, 'ADMIN', '*']));
  }
}

function who_() {
  const email = Session.getActiveUser().getEmail();
  return email ? email.split('@')[0].toLowerCase() : '';
}

const AUTO_SUPERVISORS = true; // anyone listed as Immediate Supervisor in the Agents List can tag their own teams

function roleFor_(ldap) {
  if (!ldap) return null;
  let r = null;
  const row = roles_()[ldap];
  if (row) r = { all: row.role === 'ADMIN' || row.role === 'MANAGER' || row.teams.indexOf('*') > -1, teams: row.teams };
  if (isAdmin_(ldap)) r = { all: true, teams: [] };
  if (AUTO_SUPERVISORS && !(r && r.all)) {
    const ros = roster_(), mine = {};
    Object.keys(ros).forEach(l => { if (ros[l].sup === ldap && ros[l].team) mine[ros[l].team] = true; });
    const t = Object.keys(mine);
    if (t.length) r = { all: false, teams: ((r && r.teams) || []).concat(t) };
  }
  return r;
}

function addDaysStr_(s, n) {
  const p = s.split('-');
  const t = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n));
  return t.getUTCFullYear() + '-' + ('0' + (t.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + t.getUTCDate()).slice(-2);
}

function getTeamWeek(team, ws) {
  const who = who_();
  const role = roleFor_(who);
  if (!role) throw new Error('Not authorized');
  ws = String(ws || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ws)) throw new Error('Bad week');

  const ss = home_();
  const sh = ss.getSheetByName(SH_SHIFTS);
  if (!sh || sh.getLastRow() < 2) return { ok: false, reason: 'no_data' };

  const ros = roster_();
  const known = {};
  Object.keys(ros).forEach(l => { const t = ros[l].team; if (t && !known[tkey_(t)]) known[tkey_(t)] = t; });
  const roleKeys = role.teams.map(tkey_);
  const teams = (role.all
    ? Object.keys(known).map(k => known[k])
    : Object.keys(known).filter(k => roleKeys.indexOf(k) > -1).map(k => known[k])).sort();
  if (!teams.length) return { ok: false, reason: 'no_team' };
  const tk = teams.map(tkey_).indexOf(tkey_(team)) > -1 ? tkey_(team) : tkey_(teams[0]);
  team = known[tk];

  const days = [];
  for (let i = 0; i < 7; i++) days.push(addDaysStr_(ws, i));

  const props = PropertiesService.getScriptProperties();
  const cache = CacheService.getScriptCache();
  const ck = 'tw2:' + (props.getProperty('sver') || '0') + ':' + rosterSig_() + ':' + tk.replace(/[^a-z0-9]/g, '') + ':' + ws;
  let people = null;
  try { const hit = cache.get(ck); if (hit) people = JSON.parse(hit); } catch (e) {}

  if (!people) {
    const skip = nonAgents_(ros);
    const inWk = {};
    days.forEach(d => { inWk[d] = true; });
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, SHIFT_HEAD.length).getValues();
    const map = {};
    rows.forEach(r => {
      const l = String(r[1]).toLowerCase();
      if (skip[l]) return;
      if (tkey_(ros[l] ? ros[l].team : String(r[4])) !== tk) return;
      const d = ds_(r[0]);
      if (!inWk[d]) return;
      const p = map[l] || (map[l] = { l: l, n: String(r[3]), ch: String(r[6]), emp: (ros[l] && ros[l].emp) || String(r[2]), s: {}, t: {} });
      p.s[d] = String(r[7]);
    });
    const tgs = ss.getSheetByName(SH_TAGS);
    if (tgs && tgs.getLastRow() > 1) {
      tgs.getRange(2, 1, tgs.getLastRow() - 1, TAG_HEAD.length).getValues().forEach(r => {
        const d = ds_(r[0]), l = String(r[1]).toLowerCase();
        if (inWk[d] && map[l]) map[l].t[d] = { t: String(r[2]).toUpperCase(), by: String(r[3]), at: String(r[4]) };
      });
    }
    people = Object.keys(map).map(k => map[k]).sort((a, b) => a.n.localeCompare(b.n));
    try { cache.put(ck, JSON.stringify(people), 21600); } catch (e) {}
  }

  // OT approvals: cached for 5 minutes per team-week, fails soft if the tracker can't be read
  let otm = null;
  const otk = 'otw1:' + tk.replace(/[^a-z0-9]/g, '') + ':' + ws;
  try { const h = cache.get(otk); if (h) otm = JSON.parse(h); } catch (e) {}
  if (!otm) {
    try { otm = otTeam_(people.map(p => p.l), days[0], days[6]); cache.put(otk, JSON.stringify(otm), 300); } catch (e) { otm = {}; }
  }
  people.forEach(p => { p.o = {}; days.forEach(d => { const k = p.l + '|' + d; if (otm[k]) p.o[d] = otm[k]; }); });

  const tmap = tagsWeek_(days);
  people.forEach(p => { p.t = {}; days.forEach(d => { const g = tmap[p.l + '|' + d]; if (g) p.t[d] = g; }); });

  // Cobra attendance per team-week (5 min cache, only cached when both tabs read fine)
  let pfm = null;
  const pfk = 'pfw1:' + tk.replace(/[^a-z0-9]/g, '') + ':' + ws;
  try { const h = cache.get(pfk); if (h) pfm = JSON.parse(h); } catch (e) {}
  if (!pfm) {
    pfm = {};
    try {
      const pf = perf_();
      people.forEach(p => days.forEach(d => {
        const a = pf.att[p.l] && pf.att[p.l][d];
        if (a != null) pfm[p.l + '|' + d] = a;
      }));
      if (!pf.err) cache.put(pfk, JSON.stringify(pfm), 300);
    } catch (e) { pfm = {}; }
  }
  people.forEach(p => { p.a = {}; days.forEach(d => { const v = pfm[p.l + '|' + d]; if (v != null) p.a[d] = v; }); });

  // open eStart requests are never cached, so badges are always current
  const open = {};
  readCorr_().forEach(o => { if (o.st === 'PENDING' || o.st === 'INFO') open[o.l + '|' + o.d] = { id: o.id, st: o.st, t: o.t }; });
  people.forEach(p => { p.c = {}; days.forEach(d => { const k = p.l + '|' + d; if (open[k]) p.c[d] = open[k]; }); });

  return { ok: true, team: team, teams: teams, ws: ws, days: days, people: people, me: who, syncedAt: props.getProperty('syncedAt') || '' };
}

function applyTags(items) {
  const who = who_();
  const role = roleFor_(who);
  if (!role) throw new Error('Not authorized');
  if (!items || !items.length) return { ok: true, applied: 0, skippedKeys: [] };
  if (items.length > 700) throw new Error('Too many cells selected');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = home_();
    const want = {};
    items.forEach(it => {
      const d = String((it && it.d) || '');
      const l = String((it && it.l) || '').toLowerCase();
      const t = String((it && it.t) || '').toUpperCase();
      if (/^\d{4}-\d{2}-\d{2}$/.test(d) && l && (t === '' || TAG_CODES.indexOf(t) > -1)) want[d + '|' + l] = t;
    });

    const sh = ss.getSheetByName(SH_SHIFTS);
    const info = {};
    const ros = roster_();
    let nm = '';
    const skipSet = nonAgents_(ros);
    sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), SHIFT_HEAD.length).getValues().forEach(r => {
      const l = String(r[1]).toLowerCase();
      if (l === who) nm = String(r[3]);
      const k = ds_(r[0]) + '|' + l;
      if (k in want) info[k] = { code: String(r[7]), team: (ros[l] && ros[l].team) || String(r[4]) };
    });

    const ok = {}, skipped = [];
    Object.keys(want).forEach(k => {
      const i = info[k];
      const l = k.split('|')[1];
      if (i && !skipSet[l] && tagOk_(want[k], i.code) && (role.all || role.teams.map(tkey_).indexOf(tkey_(i.team)) > -1)) ok[k] = want[k];
      else skipped.push(k);
    });

    const tg = ss.getSheetByName(SH_TAGS) || ss.insertSheet(SH_TAGS);
    const n = TAG_HEAD.length;
    const last = tg.getLastRow();
    const rows = last > 1 ? tg.getRange(2, 1, last - 1, n).getValues() : [];
    const idx = {};
    rows.forEach((r, i) => { r[0] = ds_(r[0]); idx[r[0] + '|' + String(r[1]).toLowerCase()] = i; });

    const by = nm ? (nm.indexOf(',') > -1 ? nm.split(',')[1].trim() : nm) : (who.charAt(0).toUpperCase() + who.slice(1));
    const stamp = Utilities.formatDate(new Date(), MPS_TZ, 'MMM d, HH:mm');
    const logStamp = Utilities.formatDate(new Date(), MPS_TZ, 'yyyy-MM-dd HH:mm:ss');
    const log = [];
    let applied = 0;

    Object.keys(ok).forEach(k => {
      const t = ok[k], p = k.split('|');
      const i = (k in idx) ? idx[k] : -1;
      const old = (i > -1 && rows[i]) ? String(rows[i][2]).toUpperCase() : '';
      if (t === old) { applied++; return; }
      if (t === '') {
        rows[i] = null;
        delete idx[k];
      } else if (i > -1) {
        rows[i] = [p[0], p[1], t, by, stamp, String(rows[i][5] || '')];
      } else {
        idx[k] = rows.length;
        rows.push([p[0], p[1], t, by, stamp, '']);
      }
      log.push([logStamp, who, p[0], p[1], old, t, '']);
      applied++;
    });

    const out = rows.filter(Boolean);
    if (out.length) {
      if (tg.getMaxRows() < out.length + 1) tg.insertRowsAfter(tg.getMaxRows(), out.length + 1 - tg.getMaxRows());
      const rng = tg.getRange(2, 1, out.length, n);
      rng.setNumberFormat('@');
      rng.setValues(out);
    }
    if (last - 1 > out.length) tg.getRange(out.length + 2, 1, last - 1 - out.length, n).clearContent();

    if (log.length) {
      const lg = ss.getSheetByName(SH_TAGLOG) || ss.insertSheet(SH_TAGLOG);
      const ll = Math.max(lg.getLastRow(), 1);
      if (lg.getMaxRows() < ll + log.length) lg.insertRowsAfter(lg.getMaxRows(), ll + log.length - lg.getMaxRows());
      const r2 = lg.getRange(ll + 1, 1, log.length, LOG_HEAD.length);
      r2.setNumberFormat('@');
      r2.setValues(log);
      PropertiesService.getScriptProperties().setProperty('ver', String(Date.now()));
    }
    return { ok: true, applied: applied, skippedKeys: skipped, by: by, at: stamp };
  } finally {
    lock.releaseLock();
  }
}
/**
 * ------------------------------------------------------------------
 * TEAM ROSTER - source of truth is the Agents List tab
 * B = Google LDAP, C = Immediate Supervisor LDAP, D = Team / Supporting Team
 * ------------------------------------------------------------------
 */
function roster_() {
  if (ROS_MEMO_) return ROS_MEMO_;
  const sh = home_().getSheetByName(SH_AGENTS);
  const map = {};
  if (sh && sh.getLastRow() >= 3) {
    sh.getRange(3, 2, sh.getLastRow() - 2, 4).getValues().forEach(r => {
      const l = String(r[0]).trim().toLowerCase();
      if (!l || map[l]) return;
      map[l] = { emp: String(r[1]).trim(), sup: String(r[2]).trim().toLowerCase(), team: String(r[3]).replace(/\s+/g, ' ').trim() };
    });
  }
  return (ROS_MEMO_ = map);
}

function tkey_(s) {
  return String(s || '').replace(/\s+/g, ' ').trim().toLowerCase();
}

function rosterSig_() {
  const ros = roster_();
  const s = Object.keys(ros).sort().map(k => k + '=' + ros[k].team + '/' + ros[k].sup + '/' + ros[k].emp).join(';');
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s)).slice(0, 8);
}

function nonAgents_(ros) {
  const set = {};
  Object.keys(ros).forEach(l => { if (ros[l].sup) set[ros[l].sup] = true; });
  ADMINS.forEach(a => { set[a] = true; });
  Object.keys(roles_()).forEach(l => { set[l] = true; });
  return set;
}

/**
 * ------------------------------------------------------------------
 * ESTART CORRECTIONS
 * ------------------------------------------------------------------
 */
const CORR_HEAD = ['ID', 'Submitted', 'LDAP', 'Name', 'Team', 'Supervisor', 'Date', 'Scheduled', 'Type', 'Login', 'ShiftType', 'Comment', 'Status', 'ReviewedBy', 'ReviewedAt', 'Note', 'Reply'];
const CORR_TYPES = ['Leave', 'Rest day', 'LOA (Workday)', 'No/Incorrect schedule', 'Tools/Login/eStart issue', 'Overtime', 'Shift cross-over (Absence)', 'Shift cross-over (Leave)'];
const CORR_SHIFT_TYPES = ['Regular shift - no OT', 'RD OT', 'Pre-shift OT', 'Post-shift OT'];
const CORR_EMAILS = false; // true = email the supervisor on submit and the agent on review

function corrObj_(r) {
  return {
    id: String(r[0]), at: String(r[1]), l: String(r[2]).toLowerCase(), n: String(r[3]), team: String(r[4]),
    sup: String(r[5]), d: ds_(r[6]), sc: String(r[7]), t: String(r[8]), lt: String(r[9]), sty: String(r[10]),
    c: String(r[11]), st: String(r[12]) || 'PENDING', by: String(r[13]), rat: String(r[14]), rn: String(r[15]), rp: String(r[16])
  };
}

function decor_(o, ros) {
  const m = ros[o.l];
  if (m) { o.team = m.team; o.sup = m.sup; o.emp = m.emp; } else { o.emp = ''; }
  return o;
}

let CORR_MEMO_ = null;
function readCorr_() {
  if (CORR_MEMO_) return CORR_MEMO_;
  const sh = home_().getSheetByName(SH_CORR);
  if (!sh || sh.getLastRow() < 2) return [];
  const ros = roster_();
  return (CORR_MEMO_ = sh.getRange(2, 1, sh.getLastRow() - 1, CORR_HEAD.length).getValues()
    .map(corrObj_).filter(o => o.id).map(o => decor_(o, ros)));
}

const CLOG_HEAD = ['Timestamp', 'By', 'ID', 'LDAP', 'Date', 'Action', 'From', 'To', 'Note'];

function corrLog_(by, o, action, from, to, note) {
  try {
    const ss = home_();
    let sh = ss.getSheetByName(SH_CLOG);
    if (!sh) {
      sh = ss.insertSheet(SH_CLOG);
      sh.getRange(1, 1, sh.getMaxRows(), CLOG_HEAD.length).setNumberFormat('@');
      sh.getRange(1, 1, 1, CLOG_HEAD.length).setValues([CLOG_HEAD]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    const r = Math.max(sh.getLastRow(), 1) + 1;
    if (sh.getMaxRows() < r) sh.insertRowsAfter(sh.getMaxRows(), r - sh.getMaxRows());
    sh.getRange(r, 1, 1, CLOG_HEAD.length).setValues([[
      Utilities.formatDate(new Date(), MPS_TZ, 'yyyy-MM-dd HH:mm:ss'), by, o.id, o.l, o.d, action, from || '', to || '', note || ''
    ]]);
  } catch (e) {}
}

function corrFor_(ldap) {
  return readCorr_().filter(o => o.l === ldap).sort((a, b) => a.at < b.at ? 1 : a.at > b.at ? -1 : 0);
}

function corrOpenCount_(role) {
  const who = who_();
  const keys = role.teams.map(tkey_);
  return readCorr_().filter(o => o.st === 'PENDING' && o.l !== who && (role.all || keys.indexOf(tkey_(o.team)) > -1)).length;
}

function corrSheet_() {
  const ss = home_();
  let sh = ss.getSheetByName(SH_CORR);
  if (!sh) {
    sh = ss.insertSheet(SH_CORR);
    sh.getRange(1, 1, sh.getMaxRows(), CORR_HEAD.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, CORR_HEAD.length).setValues([CORR_HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function findCorrRow_(sh, id) {
  if (!sh || sh.getLastRow() < 2) return 0;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function shiftInfo_(ldap, d) {
  const sh = home_().getSheetByName(SH_SHIFTS);
  const out = { name: '', code: '' };
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues().forEach(r => {
    if (String(r[1]).toLowerCase() !== ldap) return;
    out.name = String(r[3]);
    if (d && ds_(r[0]) === d) out.code = String(r[7]);
  });
  return out;
}

function corrMail_(kind, o) {
  if (!CORR_EMAILS) return;
  try {
    const url = ScriptApp.getService().getUrl();
    if (kind === 'new') {
      if (!o.sup) return;
      MailApp.sendEmail({
        to: o.sup + '@google.com',
        subject: 'eStart correction · ' + o.n + ' · ' + o.d,
        body: o.n + ' (' + o.team + ') submitted an eStart correction.\n\nType: ' + o.t + '\nDate: ' + o.d + '\nScheduled: ' + o.sc + '\nComment: ' + o.c + '\n\nReview it here: ' + url
      });
    } else {
      MailApp.sendEmail({
        to: o.l + '@google.com',
        subject: 'Your eStart correction for ' + o.d + ': ' + o.st,
        body: 'Status: ' + o.st + '\nReviewed by: ' + o.by + '\nNote: ' + (o.rn || '-') + '\n\nOpen your schedule: ' + url
      });
    }
  } catch (e) {}
}

function submitCorrection(p) {
  const who = who_();
  if (!who) throw new Error('Can\u2019t verify your account');
  p = p || {};
  const me = roster_()[who];
  if (!me) throw new Error('You are not on the Agents List');
  const d = String(p.d || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) throw new Error('Pick a valid shift date');
  const today = Utilities.formatDate(new Date(), MPS_TZ, 'yyyy-MM-dd');
  if (d < addDaysStr_(today, -30) || d > addDaysStr_(today, 1)) throw new Error('Date must be within the last 30 days');
  const type = String(p.t || '');
  if (CORR_TYPES.indexOf(type) < 0) throw new Error('Choose what went wrong');
  const sty = CORR_SHIFT_TYPES.indexOf(String(p.sty || '')) > -1 ? String(p.sty) : CORR_SHIFT_TYPES[0];
  const lt = /^\d{2}:\d{2}$/.test(String(p.lt || '')) ? String(p.lt) : '';
  const c = String(p.c || '').trim();
  if (c.length < 3) throw new Error('Add a short comment or a Snip-It link');
  if (c.length > 800) throw new Error('Comment is too long (800 characters max)');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const dup = readCorr_().some(o => o.l === who && o.d === d && o.t === type && (o.st === 'PENDING' || o.st === 'INFO'));
    if (dup) throw new Error('You already have an open "' + type + '" request for ' + d);
    const info = shiftInfo_(who, d);
    const id = 'C' + Date.now().toString(36).toUpperCase();
    const stamp = Utilities.formatDate(new Date(), MPS_TZ, 'yyyy-MM-dd HH:mm');
    const row = [id, stamp, who, info.name, me.team, me.sup, d, info.code, type, lt, sty, c, 'PENDING', '', '', '', ''];
    const sh = corrSheet_();
    const r = Math.max(sh.getLastRow(), 1) + 1;
    if (sh.getMaxRows() < r) sh.insertRowsAfter(sh.getMaxRows(), r - sh.getMaxRows());
    const rng = sh.getRange(r, 1, 1, CORR_HEAD.length);
    rng.setNumberFormat('@');
    rng.setValues([row]);
    const item = corrObj_(row);
    corrLog_(who, item, 'SUBMITTED', '', 'PENDING', c);
    corrMail_('new', item);
    return { ok: true, item: item };
  } finally {
    lock.releaseLock();
  }
}

function getCorrQueue() {
  const who = who_();
  const role = roleFor_(who);
  if (!role) throw new Error('Not authorized');
  const ros = roster_();
  const known = {};
  Object.keys(ros).forEach(l => { const t = ros[l].team; if (t && !known[tkey_(t)]) known[tkey_(t)] = t; });
  const keys = role.teams.map(tkey_);
  const teams = Object.keys(known).filter(k => role.all || keys.indexOf(k) > -1).map(k => known[k]).sort();
  const allow = {};
  teams.forEach(t => { allow[tkey_(t)] = true; });
  const from = addDaysStr_(Utilities.formatDate(new Date(), MPS_TZ, 'yyyy-MM-dd'), -90);

  const tmap = {};
  const tg = home_().getSheetByName(SH_TAGS);
  if (tg && tg.getLastRow() > 1) {
    tg.getRange(2, 1, tg.getLastRow() - 1, 3).getValues().forEach(r => { tmap[ds_(r[0]) + '|' + String(r[1]).toLowerCase()] = String(r[2]).toUpperCase(); });
  }
  const items = readCorr_()
    .filter(o => allow[tkey_(o.team)] && o.st !== 'WITHDRAWN' && (o.st === 'PENDING' || o.st === 'INFO' || o.d >= from))
    .map(o => Object.assign({}, o, { mine: o.sup === who, tag: tmap[o.d + '|' + o.l] || '' }));
  return { ok: true, me: who, teams: teams, items: items };
}

function reviewCorrection(id, status, note, tag) {
  const who = who_();
  const role = roleFor_(who);
  if (!role) throw new Error('Not authorized');
  status = String(status || '').toUpperCase();
  if (['PENDING', 'INFO', 'DONE', 'REJECTED'].indexOf(status) < 0) throw new Error('Bad status');
  note = String(note || '').trim().slice(0, 500);
  if ((status === 'REJECTED' || status === 'INFO') && !note) throw new Error('Add a note for the agent');
  tag = String(tag || '').toUpperCase();
  if (status !== 'DONE' || (tag !== 'CLEAR' && TAG_CODES.indexOf(tag) < 0)) tag = '';

  let res;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = home_().getSheetByName(SH_CORR);
    const r = findCorrRow_(sh, id);
    if (!r) throw new Error('Request not found');
    const rng = sh.getRange(r, 1, 1, CORR_HEAD.length);
    const row = rng.getValues()[0];
    const o = decor_(corrObj_(row), roster_());
    if (o.l === who) throw new Error('You can\u2019t review your own request');
    if (!role.all && role.teams.map(tkey_).indexOf(tkey_(o.team)) < 0) throw new Error('This request belongs to another team');
    if (o.st === 'WITHDRAWN') throw new Error('The agent withdrew this request');

    const nm = shiftInfo_(who, '').name;
    const by = nm ? (nm.indexOf(',') > -1 ? nm.split(',')[1].trim() : nm) : (who.charAt(0).toUpperCase() + who.slice(1));
    row[12] = status;
    row[13] = by;
    row[14] = Utilities.formatDate(new Date(), MPS_TZ, 'yyyy-MM-dd HH:mm');
    row[15] = note || (status === 'PENDING' ? String(row[15]) : '');
    rng.setNumberFormat('@');
    rng.setValues([row]);
    const out = decor_(corrObj_(row), roster_());
    corrLog_(who, out, status === 'PENDING' ? 'REOPENED' : status, o.st, status, note);
    if (status !== 'PENDING') corrMail_('review', out);
    res = { ok: true, item: Object.assign({}, out, { mine: out.sup === who }) };
  } finally {
    lock.releaseLock();
  }

  // applyTags takes its own lock, so it runs after ours is released
  if (tag) {
    try {
      const t = applyTags([{ d: res.item.d, l: res.item.l, t: tag === 'CLEAR' ? '' : tag }]);
      res.tagged = t.skippedKeys.length ? 'skipped' : (tag === 'CLEAR' ? 'cleared' : tag);
    } catch (e) {
      res.tagged = 'skipped';
    }
  }
  return res;
}

function agentCorrAction(id, action, text) {
  const who = who_();
  if (!who) throw new Error('Can\u2019t verify your account');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = home_().getSheetByName(SH_CORR);
    const r = findCorrRow_(sh, id);
    if (!r) throw new Error('Request not found');
    const rng = sh.getRange(r, 1, 1, CORR_HEAD.length);
    const row = rng.getValues()[0];
    const o = corrObj_(row);
    if (o.l !== who) throw new Error('This isn\u2019t your request');
    if (action === 'wd') {
      if (o.st !== 'PENDING' && o.st !== 'INFO') throw new Error('This request was already reviewed');
      row[12] = 'WITHDRAWN';
    } else if (action === 'reply') {
      if (o.st !== 'INFO') throw new Error('Nothing to reply to');
      text = String(text || '').trim().slice(0, 500);
      if (!text) throw new Error('Write a reply first');
      row[16] = text;
      row[12] = 'PENDING';
    } else {
      throw new Error('Bad action');
    }
    rng.setNumberFormat('@');
    rng.setValues([row]);
    corrLog_(who, o, action === 'wd' ? 'WITHDRAWN' : 'AGENT REPLY', o.st, String(row[12]), action === 'reply' ? text : '');
    return { ok: true, item: corrObj_(row) };
  } finally {
    lock.releaseLock();
  }
}
/**
 * ------------------------------------------------------------------
 * ARCHIVE - keeps the Shifts tab small and fast
 * ------------------------------------------------------------------
 */
const KEEP_DAYS = 120;

function keepFrom_() {
  return addDaysStr_(Utilities.formatDate(new Date(), MPS_TZ, 'yyyy-MM-dd'), -KEEP_DAYS);
}

function archiveOldShifts() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = home_();
    const sh = ss.getSheetByName(SH_SHIFTS);
    if (!sh || sh.getLastRow() < 2) return;
    const n = SHIFT_HEAD.length;
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, n).getValues();
    const cut = keepFrom_(), keep = [], old = [];
    rows.forEach(r => { r[0] = ds_(r[0]); (r[0] < cut ? old : keep).push(r); });
    if (!old.length) return;

    let ar = ss.getSheetByName(SH_ARCH);
    if (!ar) {
      ar = ss.insertSheet(SH_ARCH);
      ar.getRange(1, 1, 1, n).setValues([SHIFT_HEAD]).setFontWeight('bold');
      ar.setFrozenRows(1);
    }
    const start = Math.max(ar.getLastRow(), 1) + 1;
    if (ar.getMaxRows() < start + old.length - 1) ar.insertRowsAfter(ar.getMaxRows(), start + old.length - 1 - ar.getMaxRows());
    const a = ar.getRange(start, 1, old.length, n);
    a.setNumberFormat('@');
    a.setValues(old);

    if (keep.length) {
      const k = sh.getRange(2, 1, keep.length, n);
      k.setNumberFormat('@');
      k.setValues(keep);
    }
    sh.getRange(keep.length + 2, 1, rows.length - keep.length, n).clearContent();
    PropertiesService.getScriptProperties().setProperty('ver', String(Date.now()));
  } finally {
    lock.releaseLock();
  }
}

// Run once from the editor
function setupArchive() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'archiveOldShifts').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('archiveOldShifts').timeBased().onMonthDay(1).atHour(3).create();
  archiveOldShifts();
}

/**
 * ------------------------------------------------------------------
 * ONE-TIME IMPORT of the old form responses
 * ------------------------------------------------------------------
 */
const OLD_CORR_FILE_ID = '';                  // '' = the same spreadsheet as this app
const OLD_CORR_TAB = 'Estart Corrections';    // set to the real tab name (the tab strip shows "EST COR")
const IMPORT_SINCE = '';                      // e.g. '2026-01-01' to skip older rows
const IMPORT_BLANK_AS = 'PENDING';            // status for old rows with an empty Status cell

function importOldCorrectionsPreview() { importOldCorrections_(true); }
function importOldCorrectionsRun() { importOldCorrections_(false); }

function hid_(s) {
  return 'H' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('').slice(0, 10).toUpperCase();
}

function importOldCorrections_(dry) {
  const src = OLD_CORR_FILE_ID ? SpreadsheetApp.openById(OLD_CORR_FILE_ID) : home_();
  const sh = src.getSheetByName(OLD_CORR_TAB);
  if (!sh) throw new Error('Tab "' + OLD_CORR_TAB + '" not found. Tabs here: ' + src.getSheets().map(s => s.getName()).join(', '));
  const vals = sh.getDataRange().getDisplayValues();

  let hr = -1;
  for (let i = 0; i < Math.min(vals.length, 6); i++) if (vals[i].some(c => /ldap/i.test(c))) { hr = i; break; }
  if (hr < 0) throw new Error('No header row containing "LDAP" in the first 6 rows');
  const head = vals[hr].map(c => String(c).trim().toLowerCase());
  const col = re => head.findIndex(h => re.test(h));
  const c = { ts: col(/timestamp/), ldap: col(/ldap/), date: col(/shift date/), start: col(/shift start/), type: col(/correction type/), comment: col(/comment/), sty: col(/^shift type/), status: col(/status/) };
  if (c.status < 0) c.status = 11; // column L, as used by the old signature script
  if (c.ldap < 0 || c.date < 0 || c.type < 0) throw new Error('Could not find LDAP / Shift Date / Correction Type columns. Headers: ' + head.join(' | '));
  const g = (r, i) => (i > -1 && r[i] != null) ? String(r[i]).trim() : '';

  const ymd = s => {
    s = String(s).trim();
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return m[1] + '-' + m[2] + '-' + m[3];
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); // assumes month/day/year
    if (m) return m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
    const d = new Date(s);
    return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, MPS_TZ, 'yyyy-MM-dd');
  };
  const hhmm = s => {
    const m = String(s).trim().match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*([AaPp][Mm])?$/);
    if (!m) return '';
    let h = +m[1]; const ap = (m[3] || '').toUpperCase();
    if (ap === 'PM' && h < 12) h += 12;
    if (ap === 'AM' && h === 12) h = 0;
    return ('0' + h).slice(-2) + ':' + m[2];
  };
  const mapStatus = s => {
    s = String(s).trim().toLowerCase();
    if (!s) return IMPORT_BLANK_AS;
    if (/(reject|declin|deny|denied|invalid|disapprov)/.test(s)) return 'REJECTED';
    if (/(done|complete|correct|approv|fix|resolv|updated)/.test(s)) return 'DONE';
    return 'PENDING';
  };
  const tmap = { rd: 'Rest day' };
  CORR_TYPES.forEach(t => { tmap[t.toLowerCase()] = t; });
  const mapSty = s => {
    s = String(s).toLowerCase();
    if (!s || /regular/.test(s)) return CORR_SHIFT_TYPES[0];
    if (/rd/.test(s)) return CORR_SHIFT_TYPES[1];
    if (/pre/.test(s)) return CORR_SHIFT_TYPES[2];
    if (/post/.test(s)) return CORR_SHIFT_TYPES[3];
    return CORR_SHIFT_TYPES[0];
  };

  const ros = roster_(), names = {}, codes = {}, have = {};
  const shs = home_().getSheetByName(SH_SHIFTS);
  if (shs && shs.getLastRow() > 1) shs.getRange(2, 1, shs.getLastRow() - 1, 8).getValues().forEach(r => {
    const l = String(r[1]).toLowerCase();
    names[l] = String(r[3]);
    codes[ds_(r[0]) + '|' + l] = String(r[7]);
  });
  readCorr_().forEach(o => { have[o.id] = true; });

  const out = [], stat = {}, skipped = { noLdap: 0, badDate: 0, old: 0, dup: 0 };
  for (let i = hr + 1; i < vals.length; i++) {
    const r = vals[i];
    const ldap = g(r, c.ldap).toLowerCase();
    if (!ldap) { skipped.noLdap++; continue; }
    const d = ymd(g(r, c.date));
    if (!d) { skipped.badDate++; continue; }
    if (IMPORT_SINCE && d < IMPORT_SINCE) { skipped.old++; continue; }
    const type = tmap[g(r, c.type).toLowerCase()] || g(r, c.type) || 'Other';
    const ts = g(r, c.ts);
    const at = ts ? (ymd(ts) + ' ' + (hhmm(ts.split(/\s+/).slice(1).join(' ')) || '00:00')) : d + ' 00:00';
    const id = hid_([ts, ldap, d, type].join('|'));
    if (have[id]) { skipped.dup++; continue; }
    have[id] = true;
    const raw = g(r, c.status), st = mapStatus(raw), sig = g(r, c.status + 1), dash = sig.indexOf(' - ');
    const m = ros[ldap] || {};
    out.push([
      id, at, ldap, names[ldap] || '', m.team || '', m.sup || '', d, codes[d + '|' + ldap] || '', type,
      hhmm(g(r, c.start)), mapSty(g(r, c.sty)), g(r, c.comment).slice(0, 800) || '(imported)', st,
      st === 'PENDING' ? '' : (dash > -1 ? sig.slice(0, dash) : sig),
      st === 'PENDING' ? '' : (dash > -1 ? sig.slice(dash + 3) : ''),
      (st !== 'PENDING' && raw && !/^(done|rejected)$/i.test(raw)) ? 'Imported status: ' + raw : '', ''
    ]);
    stat[st] = (stat[st] || 0) + 1;
  }
  out.sort((a, b) => a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : 0);

  Logger.log('Header row ' + (hr + 1) + ' · columns (0-based): ' + JSON.stringify(c));
  Logger.log('Would import ' + out.length + ' · by status ' + JSON.stringify(stat) + ' · skipped ' + JSON.stringify(skipped));
  out.slice(0, 3).forEach(x => Logger.log(JSON.stringify(x)));
  if (dry || !out.length) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const dst = corrSheet_();
    const start = Math.max(dst.getLastRow(), 1) + 1;
    const need = start + out.length - 1;
    if (dst.getMaxRows() < need) dst.insertRowsAfter(dst.getMaxRows(), need - dst.getMaxRows());
    const rng = dst.getRange(start, 1, out.length, CORR_HEAD.length);
    rng.setNumberFormat('@');
    rng.setValues(out);
    Logger.log('Imported ' + out.length + ' rows');
  } finally {
    lock.releaseLock();
  }
}

/**
 * ------------------------------------------------------------------
 * OVERTIME - read from the Play OT Tracker team tabs
 * ------------------------------------------------------------------
 */
const OT_SS_ID = '1RZPGO2tntjaGGO0LI2rGMmJPpNcPtt5kvF5tzbcHpHM';
const OT_DAYS = 120;

function otStr_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm');
  const m = String(v).trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})[ T](\d{1,2}):(\d{2})/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) + ' ' + ('0' + m[4]).slice(-2) + ':' + m[5] : '';
}

function otMs_(s) {
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})$/);
  return m ? Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]) : NaN;
}

function readOtLive_(ldap) {
  const ss = SpreadsheetApp.openById(OT_SS_ID);
  const tz = ss.getSpreadsheetTimeZone();
  const from = addDaysStr_(Utilities.formatDate(new Date(), MPS_TZ, 'yyyy-MM-dd'), -OT_DAYS);
  const seen = {}, out = [];
  ss.getSheets().forEach(sh => {
    if (/^overtime data$/i.test(sh.getName()) || sh.getLastRow() < 2) return;
    const vals = sh.getDataRange().getValues();
    const head = vals[0].map(h => String(h).trim().toLowerCase());
    const cL = head.indexOf('ldap'), cS = head.indexOf('ot start'), cE = head.indexOf('ot end');
    if (cL < 0 || cS < 0 || cE < 0) return;
    const cT = head.indexOf('ot type');
    const cR = head.findIndex(h => /reason|case/.test(h));
    let cK = head.findIndex(h => /scrub/.test(h));   // "Scrubbed in eStart?"
    if (cK < 0 && head.length >= 19) cK = 18;        // column S
    for (let i = 1; i < vals.length; i++) {
      const r = vals[i];
      if (String(r[cL]).trim().toLowerCase() !== ldap) continue;
      const a = otStr_(r[cS], tz), b = otStr_(r[cE], tz);
      if (!a || !b || a.slice(0, 10) < from || seen[a + '|' + b]) continue;
      seen[a + '|' + b] = 1;
      let h = (otMs_(b) - otMs_(a)) / 36e5;
      if (h >= 9) h -= 1;                            // lunch auto-deduction, same as the OT form
      if (!(h > 0)) h = 0;
      const ty = String(cT > -1 ? r[cT] : '').toLowerCase();
      const raw = cK > -1 ? String(r[cK]).trim() : '';
      const lo = raw.toLowerCase();
      out.push({
        id: sh.getName() + '!' + (i + 1), a: a, b: b, d: a.slice(0, 10), h: Math.round(h * 100) / 100,
        t: ty.indexOf('rd') === 0 ? 'Rest day' : ty.indexOf('pre') === 0 ? 'Pre-shift' : ty.indexOf('post') === 0 ? 'Post-shift' : 'Overtime',
        r: cR > -1 ? String(r[cR]).slice(0, 200) : '',
        st: /(reject|invalid|denied|declin)/.test(lo) ? 'REJECTED' : /(done|approv|valid)/.test(lo) ? 'DONE' : 'PENDING'
      });
    }
  });
  return out.sort((x, y) => x.a < y.a ? 1 : -1);
}

function getMyOt(asLdap) {
  const self = who_();
  const as = String(asLdap || '').toLowerCase().replace(/[^a-z0-9._-]/g, '');
  const ldap = (as && isAdmin_(self)) ? as : self;
  if (!ldap) return { ok: false, message: 'no identity', items: [] };
  const cache = CacheService.getScriptCache(), key = 'ot1:' + ldap;
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  let res;
  try { res = { ok: true, items: readOt_(ldap) }; }
  catch (e) { return { ok: false, message: e.message, items: [] }; }
  try { cache.put(key, JSON.stringify(res), 300); } catch (e) {}
  return res;
}

function otStatus_(raw) {
  const lo = String(raw).trim().toLowerCase();
  return /(reject|invalid|denied|declin)/.test(lo) ? 'REJECTED' : /(done|approv|valid)/.test(lo) ? 'DONE' : 'PENDING';
}

// One pass over the OT tracker for a whole team-week: returns { 'ldap|yyyy-mm-dd': {h, st, n} }
function otTeamLive_(ldaps, from, to) {
  const want = {};
  ldaps.forEach(l => { want[l] = 1; });
  const ss = SpreadsheetApp.openById(OT_SS_ID), tz = ss.getSpreadsheetTimeZone();
  const rank = { DONE: 0, PENDING: 1, REJECTED: 2 };
  const seen = {}, out = {};
  ss.getSheets().forEach(sh => {
    if (/^overtime data$/i.test(sh.getName()) || sh.getLastRow() < 2) return;
    const vals = sh.getDataRange().getValues();
    const head = vals[0].map(h => String(h).trim().toLowerCase());
    const cL = head.indexOf('ldap'), cS = head.indexOf('ot start'), cE = head.indexOf('ot end');
    if (cL < 0 || cS < 0 || cE < 0) return;
    let cK = head.findIndex(h => /scrub/.test(h));
    if (cK < 0 && head.length >= 19) cK = 18;
    for (let i = 1; i < vals.length; i++) {
      const r = vals[i], l = String(r[cL]).trim().toLowerCase();
      if (!want[l]) continue;
      const a = otStr_(r[cS], tz), b = otStr_(r[cE], tz);
      if (!a || !b) continue;
      const d = a.slice(0, 10);
      if (d < from || d > to || seen[l + a + b]) continue;
      seen[l + a + b] = 1;
      let h = (otMs_(b) - otMs_(a)) / 36e5;
      if (h >= 9) h -= 1;
      if (!(h > 0)) h = 0;
      const st = otStatus_(cK > -1 ? r[cK] : '');
      const o = out[l + '|' + d] || (out[l + '|' + d] = { h: 0, st: 'DONE', n: 0 });
      o.h += h; o.n++;
      if (rank[st] > rank[o.st]) o.st = st;
    }
  });
  Object.keys(out).forEach(k => { out[k].h = Math.round(out[k].h * 100) / 100; });
  return out;
}

/**
 * ------------------------------------------------------------------
 * CLIENT ERROR LOG + AGENT SUMMARY PANEL
 * ------------------------------------------------------------------
 */
function logClientError(where, msg) {
  const who = who_();
  const cache = CacheService.getScriptCache(), k = 'err:' + who;
  const n = +(cache.get(k) || 0);
  if (n >= 5) return;                       // max 5 per user per 10 min
  cache.put(k, String(n + 1), 600);
  const ss = home_();
  let sh = ss.getSheetByName(SH_ERR);
  if (!sh) {
    sh = ss.insertSheet(SH_ERR);
    sh.getRange(1, 1, 1, 4).setValues([['Timestamp', 'LDAP', 'Where', 'Message']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  sh.appendRow([Utilities.formatDate(new Date(), MPS_TZ, 'yyyy-MM-dd HH:mm:ss'), who, String(where).slice(0, 40), String(msg).slice(0, 300)]);
}

function otCached_(ldap) {
  const cache = CacheService.getScriptCache(), key = 'ot1:' + ldap;
  try { const h = cache.get(key); if (h) return JSON.parse(h).items || []; } catch (e) {}
  try {
    const items = readOt_(ldap);
    try { cache.put(key, JSON.stringify({ ok: true, items: items }), 300); } catch (e) {}
    return items;
  } catch (e) { return null; }
}

function getAgentSummary(ldap, ym) {
  const who = who_(), role = roleFor_(who);
  if (!role) throw new Error('Not authorized');
  ldap = String(ldap || '').toLowerCase();
  const today = Utilities.formatDate(new Date(), MPS_TZ, 'yyyy-MM-dd');
  ym = /^\d{4}-\d{2}$/.test(String(ym)) ? String(ym) : today.slice(0, 7);
  const m = roster_()[ldap];
  if (!m) throw new Error('This agent is not on the Agents List');
  if (!role.all && role.teams.map(tkey_).indexOf(tkey_(m.team)) < 0) throw new Error('This agent belongs to another team');

  const ss = home_(), days = {};

  const sh = ss.getSheetByName(SH_SHIFTS);
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues().forEach(r => {
      if (String(r[1]).toLowerCase() !== ldap) return;
      const d = ds_(r[0]);
      if (d.slice(0, 7) === ym) days[d] = { c: String(r[7]), t: '' };
    });
  }
  const tg = ss.getSheetByName(SH_TAGS);
  if (tg && tg.getLastRow() > 1) {
    tg.getRange(2, 1, tg.getLastRow() - 1, 3).getValues().forEach(r => {
      if (String(r[1]).toLowerCase() !== ldap) return;
      const d = ds_(r[0]);
      if (d.slice(0, 7) === ym) { days[d] = days[d] || { c: '', t: '' }; days[d].t = String(r[2]).toUpperCase(); }
    });
  }
  const c = { att: 0, abs: 0, late: 0, ovl: 0, dnc: 0, leave: 0, off: 0, pend: 0 };
  Object.keys(days).forEach(d => {
    const x = days[d], time = /^\d{2}:\d{2}$/.test(x.c);
    if (time && d <= today && !x.t) c.pend++;
    if (x.t === 'DONE') c.att++;
    else if (x.t === 'LATE') { c.att++; c.late++; }
    else if (x.t === 'OVL') { c.att++; c.ovl++; }
    else if (x.t === 'DNC') { c.att++; c.dnc++; }
    else if (x.t === 'ABS' || x.t === 'NCNS') c.abs++;
    else if (['SL', 'VL', 'MVL', 'MWL', 'LOA', 'LWOP'].indexOf(x.t) > -1) c.leave++;
    else if (!x.t && x.c && !time && x.c !== 'OFF') c.leave++;
    else if (x.c === 'OFF' && x.t !== 'RDOT') c.off++;
  });
  const ot = otCached_(ldap);
  return {
    ok: true, ldap: ldap, ym: ym, team: m.team, sup: m.sup, emp: m.emp, c: c, days: days,
    ot: (ot || []).filter(o => o.d.slice(0, 7) === ym), otErr: ot === null,
    corr: readCorr_().filter(o => o.l === ldap && o.d.slice(0, 7) === ym && o.st !== 'WITHDRAWN').map(o => ({ id: o.id, d: o.d, t: o.t, st: o.st }))
  };
}

/**
 * ------------------------------------------------------------------
 * COBRA ATTENDANCE / PRODUCTIVITY / YTD
 * ------------------------------------------------------------------
 */
const PERF_DAYS = 120;
const COL_COBRA = { ldap: 1, date: 0, val: 4 };   // B = ldap, A = date, E = attendance
const COL_PROD  = { ldap: 4, date: 3, val: 6 };   // E = agent, D = date, G = productivity
const COL_YTD   = { ldap: 0, month: 1, val: 2, ord: 3 }; // A, B, C, D
let PERF_MEMO_ = null;

function dKey_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const m = String(v).trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : '';
}

// 0.94 -> 94 ; "94%" -> 94 ; blank -> null
function pctOf_(v) {
  if (v === '' || v == null) return null;
  const s = String(v).trim();
  const n = typeof v === 'number' ? v : parseFloat(s.replace('%', ''));
  if (isNaN(n)) return null;
  return Math.round((s.indexOf('%') > -1 ? n : n * 100) * 10) / 10;
}

function perfRead_(name, c, from) {
  const ss = home_(), sh = ss.getSheetByName(name), out = {};
  if (!sh) throw new Error('Tab "' + name + '" not found');
  if (sh.getLastRow() < 2) return out;
  const tz = ss.getSpreadsheetTimeZone();
  const n = Math.max(c.ldap, c.date, c.val) + 1;
  sh.getRange(2, 1, sh.getLastRow() - 1, n).getValues().forEach(r => {
    const l = String(r[c.ldap]).trim().toLowerCase();
    if (!l) return;
    const d = dKey_(r[c.date], tz);
    if (!d || d < from) return;
    const p = pctOf_(r[c.val]);
    if (p === null) return;
    const o = out[l] || (out[l] = {});
    const e = o[d] || (o[d] = { s: 0, n: 0 });
    e.s += p; e.n++;
  });
  Object.keys(out).forEach(l => Object.keys(out[l]).forEach(d => {
    const e = out[l][d];
    out[l][d] = Math.round(e.s / e.n * 10) / 10;   // duplicate rows for one day are averaged
  }));
  return out;
}

function perfLive_() {
  if (PERF_MEMO_) return PERF_MEMO_;
  const from = addDaysStr_(Utilities.formatDate(new Date(), MPS_TZ, 'yyyy-MM-dd'), -PERF_DAYS);
  const res = { att: {}, prd: {}, err: '' };
  try { res.att = perfRead_(SH_COBRA, COL_COBRA, from); } catch (e) { res.err += 'Attendance: ' + e.message + ' '; }
  try { res.prd = perfRead_(SH_PROD, COL_PROD, from); } catch (e) { res.err += 'Productivity: ' + e.message; }
  return (PERF_MEMO_ = res);
}

function getMyPerf(asLdap) {
  const self = who_();
  const as = String(asLdap || '').toLowerCase().replace(/[^a-z0-9._-]/g, '');
  const ldap = (as && isAdmin_(self)) ? as : self;
  if (!ldap) return { ok: false, message: 'no identity', a: {}, q: {}, err: '' };
  const cache = CacheService.getScriptCache(), key = 'pf1:' + ldap;
  try { const h = cache.get(key); if (h) return JSON.parse(h); } catch (e) {}
  const p = perf_();
  const res = { ok: true, a: p.att[ldap] || {}, q: p.prd[ldap] || {}, err: p.err.trim() };
  if (!res.err) { try { cache.put(key, JSON.stringify(res), 600); } catch (e) {} }
  return res;
}

function getMyYtd(asLdap) {
  const self = who_();
  const as = String(asLdap || '').toLowerCase().replace(/[^a-z0-9._-]/g, '');
  const ldap = (as && isAdmin_(self)) ? as : self;
  if (!ldap) return { ok: false, message: 'no identity' };
  const ss = home_(), sh = ss.getSheetByName(SH_YTD);
  if (!sh) return { ok: false, message: 'Tab "' + SH_YTD + '" not found' };
  const rows = [];
  if (sh.getLastRow() > 1) {
    const tz = ss.getSpreadsheetTimeZone();
    sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(r => {
      if (String(r[COL_YTD.ldap]).trim().toLowerCase() !== ldap) return;
      rows.push({ m: dKey_(r[COL_YTD.month], tz), v: pctOf_(r[COL_YTD.val]), o: dKey_(r[COL_YTD.ord], tz) });
    });
  }
  rows.sort((a, b) => (a.o + a.m) < (b.o + b.m) ? 1 : -1);   // latest first
  const cur = rows.filter(r => r.v !== null)[0] || null;
  const seen = {}, months = [];
  rows.forEach(r => {
    if (r.v === null || !cur || r.m === cur.m || seen[r.m]) return;
    seen[r.m] = 1; months.push({ m: r.m, v: r.v });
  });
  return { ok: true, v: cur ? cur.v : null, asOf: cur ? cur.o : '', months: months };
}

/**
 * ------------------------------------------------------------------
 * SPEED: versioned caches, fresh tags, background snapshots
 * ------------------------------------------------------------------
 */
function getWeather() { return weatherMap_(); }

// Run this after manually editing the Shifts tab, to drop the cached schedules
function clearCaches() { PropertiesService.getScriptProperties().setProperty('sver', String(Date.now())); }

function tagsOf_(ldap) {
  const sh = home_().getSheetByName(SH_TAGS), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, TAG_HEAD.length).getValues().forEach(r => {
    if (String(r[1]).toLowerCase() !== ldap) return;
    out[ds_(r[0])] = { t: String(r[2]).toUpperCase(), by: String(r[3]), at: String(r[4]), n: String(r[5]) };
  });
  return out;
}

function tagsWeek_(days) {
  const sh = home_().getSheetByName(SH_TAGS), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  const inWk = {}; days.forEach(d => { inWk[d] = 1; });
  sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(r => {
    const d = ds_(r[0]);
    if (inWk[d]) out[String(r[1]).toLowerCase() + '|' + d] = { t: String(r[2]).toUpperCase(), by: String(r[3]), at: String(r[4]) };
  });
  return out;
}

// Large objects split across several cache entries (100 KB limit each)
function bigPut_(key, obj, ttl) {
  const s = JSON.stringify(obj), n = Math.ceil(s.length / 90000) || 1, m = {};
  for (let i = 0; i < n; i++) m[key + ':' + i] = s.slice(i * 90000, (i + 1) * 90000);
  m[key + ':n'] = String(n);
  CacheService.getScriptCache().putAll(m, ttl);
}
function bigGet_(key) {
  const cache = CacheService.getScriptCache(), n = +(cache.get(key + ':n') || 0);
  if (!n) return null;
  const ks = []; for (let i = 0; i < n; i++) ks.push(key + ':' + i);
  const m = cache.getAll(ks); let s = '';
  for (let i = 0; i < n; i++) { const p = m[key + ':' + i]; if (p == null) return null; s += p; }
  try { return JSON.parse(s); } catch (e) { return null; }
}

// ---- attendance / productivity snapshot ----
function perf_() {
  if (PERF_MEMO_) return PERF_MEMO_;
  const snap = bigGet_('pfall1');
  if (snap) return (PERF_MEMO_ = snap);
  const res = perfLive_();
  if (!res.err) { try { bigPut_('pfall1', res, 1800); } catch (e) {} }
  return res;
}

// ---- overtime snapshot (one pass over the tracker, grouped by agent) ----
const OT_TAIL = 1500;   // newest rows read per tab
let OTSNAP_MEMO_ = null;

function otAll_() {
  const ss = SpreadsheetApp.openById(OT_SS_ID), tz = ss.getSpreadsheetTimeZone();
  const from = addDaysStr_(Utilities.formatDate(new Date(), MPS_TZ, 'yyyy-MM-dd'), -OT_DAYS);
  const seen = {}, out = {};
  ss.getSheets().forEach(sh => {
    const last = sh.getLastRow(), cols = sh.getLastColumn();
    if (/^overtime data$/i.test(sh.getName()) || last < 2 || cols < 1) return;
    const head = sh.getRange(1, 1, 1, cols).getValues()[0].map(h => String(h).trim().toLowerCase());
    const cL = head.indexOf('ldap'), cS = head.indexOf('ot start'), cE = head.indexOf('ot end');
    if (cL < 0 || cS < 0 || cE < 0) return;
    const cT = head.indexOf('ot type'), cR = head.findIndex(h => /reason|case/.test(h));
    let cK = head.findIndex(h => /scrub/.test(h));
    if (cK < 0 && cols >= 19) cK = 18;
    const start = Math.max(2, last - OT_TAIL + 1);
    sh.getRange(start, 1, last - start + 1, cols).getValues().forEach((r, i) => {
      const l = String(r[cL]).trim().toLowerCase();
      if (!l) return;
      const a = otStr_(r[cS], tz), b = otStr_(r[cE], tz);
      if (!a || !b || a.slice(0, 10) < from || seen[l + a + b]) return;
      seen[l + a + b] = 1;
      let h = (otMs_(b) - otMs_(a)) / 36e5;
      if (h >= 9) h -= 1;
      if (!(h > 0)) h = 0;
      const ty = String(cT > -1 ? r[cT] : '').toLowerCase();
      (out[l] = out[l] || []).push({
        id: sh.getName() + '!' + (start + i), a: a, b: b, d: a.slice(0, 10), h: Math.round(h * 100) / 100,
        t: ty.indexOf('rd') === 0 ? 'Rest day' : ty.indexOf('pre') === 0 ? 'Pre-shift' : ty.indexOf('post') === 0 ? 'Post-shift' : 'Overtime',
        r: cR > -1 ? String(r[cR]).slice(0, 200) : '',
        st: otStatus_(cK > -1 ? r[cK] : '')
      });
    });
  });
  Object.keys(out).forEach(l => out[l].sort((x, y) => x.a < y.a ? 1 : -1));
  return out;
}

function otSnap_() {
  if (OTSNAP_MEMO_) return OTSNAP_MEMO_;
  let s = bigGet_('otall1');
  if (!s) { s = otAll_(); try { bigPut_('otall1', s, 1800); } catch (e) {} }
  return (OTSNAP_MEMO_ = s);
}

function readOt_(ldap) { return otSnap_()[ldap] || []; }
function otTeam_(ldaps, from, to) { return otTeamLive_(ldaps, from, to); }

// ---- background refresh (every 5 minutes) ----
function warmCaches() {
  const p = perfLive_();
  if (!p.err) bigPut_('pfall1', p, 1800);
  bigPut_('otall1', otAll_(), 1800);
}

// Run once from the editor, then approve the permissions
function setupWarm() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'warmCaches').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('warmCaches').timeBased().everyMinutes(5).create();
  warmCaches();
}
