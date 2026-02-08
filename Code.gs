/**
 * WFM COMMAND CENTER - BACKEND v3.1
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
  
  if (!dbSched || !dbIDP) throw new Error("System not initialized. Please Import Data.");

  const rotation = PropertiesService.getScriptProperties().getProperty('CURRENT_ROTATION') || 'Week A';

  // Fetch Data
  const schedData = dbSched.getDataRange().getValues();
  const idpData = dbIDP.getDataRange().getValues();
  const furloughs = dbFurlough ? dbFurlough.getDataRange().getValues() : [];

  schedData.shift(); idpData.shift(); furloughs.shift();

  // Init Buckets
  let buckets = Array.from({length: 96}, (_, i) => ({
    index: i,
    label: indexToTime(i),
    supply: 0,
    demand: 0,
    net: 0
  }));

  // 1. Demand (IDP)
  idpData.forEach(row => {
    if (formatDate(row[0]) === selectedDate) {
      let idx = timeToBucketIndex(row[1]);
      if (idx > -1) buckets[idx].demand += Number(row[2]);
    }
  });

  // 2. Supply (Schedule)
  // UPDATED EXCLUSION LIST based on your raw data
  const NON_SUPPLY_ACTIVITIES = [
    'Break/Pause', 
    'Lunch/Repas', 
    '(ADT) Repas payé / Paid Lunch',
    'ACSU Libération volontaire / Solicited Time Off', // Furlough code
    'STDP (RFT/RPT) Maladie longue durée / Long Term Sick',
    'Off'
  ];
  
  // Date Logic for Overnight
  const selDateObj = new Date(selectedDate.replace(/-/g, '/'));
  const prevDateObj = new Date(selDateObj);
  prevDateObj.setDate(selDateObj.getDate() - 1);
  const prevDateStr = Utilities.formatDate(prevDateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");

  schedData.forEach(row => {
    let agent = row[0];
    let sDate = formatDate(row[1]);
    let act = row[2];
    let startStr = row[3];
    let endStr = row[4];

    // SKIP if activity is a break, lunch, or time off
    if (NON_SUPPLY_ACTIVITIES.some(x => act.includes(x))) return;

    let startIdx = timeToBucketIndex(startStr);
    let endIdx = timeToBucketIndex(endStr);
    let isOvernight = endIdx < startIdx;

    if (sDate === selectedDate) {
      let actualEnd = isOvernight ? 96 : endIdx;
      fillBuckets(buckets, startIdx, actualEnd, agent, selectedDate, furloughs);
    }

    if (sDate === prevDateStr && isOvernight) {
      fillBuckets(buckets, 0, endIdx, agent, selectedDate, furloughs);
    }
  });

  // Calculate Net
  buckets.forEach(b => {
    b.net = parseFloat((b.supply - b.demand).toFixed(2));
    b.supply = parseFloat(b.supply.toFixed(2));
  });

  // Get Furlough Log
  let todayFurloughs = furloughs
    .filter(f => formatDate(f[2]) === selectedDate)
    .map(f => ({ 
      id: f[0], agent: f[1], time: formatTime(f[3]), type: f[4], rotation: f[5] 
    }));

  return { grid: buckets, furloughs: todayFurloughs, rotation: rotation };
}

function fillBuckets(buckets, start, end, agent, dateStr, furloughs) {
  let cutOffIndex = 96;

  // Check manual furloughs added via Tool
  furloughs.forEach(f => {
    if (f[1] === agent && formatDate(f[2]) === dateStr) {
      let fStart = timeToBucketIndex(f[3]);
      if (fStart > -1 && fStart < cutOffIndex) cutOffIndex = fStart;
    }
  });

  let actualEnd = (end > cutOffIndex) ? cutOffIndex : end;

  for (let i = start; i < actualEnd; i++) {
    if (buckets[i]) buckets[i].supply += 1;
  }
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

// Bridge for UI
function submitFurlough(a,d,t) { /* Use previous code */ }
function getStatsReport() { /* Use previous code */ }
function setRotation(r) { PropertiesService.getScriptProperties().setProperty('CURRENT_ROTATION', r); }
function importRawData(s, i) { return ImportHandler.processPaste(s, i); }
function getBucketDetails(d, i) { /* Use previous code */ }
function archiveOldData() { /* Use previous code */ }
