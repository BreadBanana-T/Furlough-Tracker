/**
 * WFM COMMAND CENTER - BACKEND v3.2
 * Fix: Uses IDP "Open" column as source of truth for Supply.
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('WFM Command Center')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- API 1: GET HEATMAP DATA ---
function getHeatmapData(selectedDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSched = ss.getSheetByName('DB_SCHEDULE');
  const dbIDP = ss.getSheetByName('DB_IDP');
  const dbFurlough = ss.getSheetByName('DB_FURLOUGH');
  
  if (!dbIDP) throw new Error("Please Import Data first.");

  const rotation = PropertiesService.getScriptProperties().getProperty('CURRENT_ROTATION') || 'Week A';

  // Fetch Data
  const idpData = dbIDP.getDataRange().getValues(); // [Date, Time, Req, Open]
  const furloughs = dbFurlough ? dbFurlough.getDataRange().getValues() : [];
  
  idpData.shift(); // remove header
  furloughs.shift();

  // Init Buckets
  let buckets = Array.from({length: 96}, (_, i) => ({
    index: i,
    label: indexToTime(i),
    supply: 0,
    demand: 0,
    net: 0
  }));

  // 1. Map IDP Data (Source of Truth)
  // We use the IDP file for BOTH Supply (Open) and Demand (Required)
  idpData.forEach(row => {
    // row[0]=Date, row[1]=Time, row[2]=Req, row[3]=Open
    if (formatDate(row[0]) === selectedDate) {
      let idx = timeToBucketIndex(row[1]);
      if (idx > -1) {
        buckets[idx].demand += Number(row[2] || 0);
        buckets[idx].supply += Number(row[3] || 0);
      }
    }
  });

  // 2. Adjust Supply based on Furloughs (Early Release)
  // Since "Open" includes everyone scheduled, we must subtract anyone we furlough using this tool.
  const activeFurloughs = furloughs.filter(f => formatDate(f[2]) === selectedDate);
  
  activeFurloughs.forEach(f => {
    // f[1]=Agent, f[3]=StartTime
    let fStartIdx = timeToBucketIndex(f[3]);
    
    // If agent is furloughed from X time, they are removed from supply from X until End of Day (96)
    // (Assuming their shift ends by midnight for simplicity, or we let them fall off naturally)
    if (fStartIdx > -1) {
      for (let i = fStartIdx; i < 96; i++) {
        // Only subtract if we have supply to subtract from
        if (buckets[i].supply > 0) {
           buckets[i].supply -= 1;
        }
      }
    }
  });

  // 3. Calculate Net
  buckets.forEach(b => {
    b.net = parseFloat((b.supply - b.demand).toFixed(2));
    b.supply = parseFloat(b.supply.toFixed(2));
    b.demand = parseFloat(b.demand.toFixed(2));
  });

  // 4. Furlough Log
  let log = activeFurloughs.map(f => ({ 
    id: f[0], agent: f[1], time: formatTime(f[3]), type: f[4], rotation: f[5] 
  }));

  return { grid: buckets, furloughs: log, rotation: rotation };
}

// --- API 2: GET BUCKET DETAILS (Drill Down) ---
// This uses the Schedule file ONLY to show WHO is working, not for the math.
function getBucketDetails(date, bucketIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSched = ss.getSheetByName('DB_SCHEDULE');
  if (!dbSched) return { time: indexToTime(bucketIndex), agents: [] };

  const sched = dbSched.getDataRange().getValues();
  sched.shift();

  let agents = [];
  
  // Date Logic for Overnight Support
  const selDateObj = new Date(date.replace(/-/g, '/'));
  const prevDateObj = new Date(selDateObj);
  prevDateObj.setDate(selDateObj.getDate() - 1);
  const prevDateStr = Utilities.formatDate(prevDateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");

  const EXCLUSIONS = ['Off', 'Solicited Time Off', 'Maladie', 'Sick'];

  sched.forEach(row => {
    let name = row[0];
    let sDate = formatDate(row[1]);
    let act = row[2];
    let start = timeToBucketIndex(row[3]);
    let end = timeToBucketIndex(row[4]);
    
    // Skip if exclusion
    if (EXCLUSIONS.some(ex => act.includes(ex))) return;

    let isOvernight = end < start;
    let isActive = false;

    // Check Today's Shifts
    if (sDate === date) {
      let actualEnd = isOvernight ? 96 : end;
      if (bucketIndex >= start && bucketIndex < actualEnd) isActive = true;
    }
    // Check Yesterday's Overnight Shifts (spilling into today)
    else if (sDate === prevDateStr && isOvernight) {
      if (bucketIndex >= 0 && bucketIndex < end) isActive = true;
    }

    if (isActive) {
      agents.push({ 
        name: name, 
        activity: act, 
        shiftStart: formatTime(row[3]), 
        shiftEnd: formatTime(row[4]) 
      });
    }
  });

  return { time: indexToTime(bucketIndex), agents: agents };
}

// --- STANDARD UTILS ---
function timeToBucketIndex(val) {
  if (!val) return -1;
  if (val instanceof Date) return (val.getHours()*4) + Math.floor(val.getMinutes()/15);
  let parts = String(val).match(/(\d+):(\d+)\s?([AP]M)/i);
  if (parts) {
    let h = parseInt(parts[1]);
    let m = parseInt(parts[2]);
    let amp = parts[3].toUpperCase();
    if (amp === 'PM' && h < 12) h += 12;
    if (amp === 'AM' && h === 12) h = 0;
    return (h * 4) + Math.floor(m / 15);
  }
  return -1;
}

function indexToTime(i) {
  let h = Math.floor(i/4);
  let m = (i%4)*15;
  return `${h<10?'0'+h:h}:${m===0?'00':m}`;
}

function formatDate(d) {
  if (!d) return "";
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function formatTime(d) {
  if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), "HH:mm");
  return String(d);
}

// Bridges
function submitFurlough(a,d,t) { 
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('DB_FURLOUGH');
    if (!sheet) {
      sheet = ss.insertSheet('DB_FURLOUGH');
      sheet.appendRow(['ID', 'Agent Name', 'Date', 'Start Time', 'Type', 'WeekRotation']);
    }
    const id = new Date().getTime();
    const rot = PropertiesService.getScriptProperties().getProperty('CURRENT_ROTATION') || 'Week A';
    sheet.appendRow([id, a, d, t, 'Early Release', rot]);
    return "Success";
  } catch(e) { return "Error: System busy."; } finally { lock.releaseLock(); }
}

function setRotation(r) { PropertiesService.getScriptProperties().setProperty('CURRENT_ROTATION', r); }
function importRawData(s, i) { return ImportHandler.processPaste(s, i); }
function getStatsReport() { /* Keep existing stats logic */ return { total: 0, byRot: {} }; }
function archiveOldData() { /* Keep existing archive logic */ return "Archived"; }
