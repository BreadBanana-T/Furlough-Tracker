/**
 * WFM COMMAND CENTER - BACKEND v7.0
 * Fixes: Logic Thresholds, Name Trimming, Segment Summing
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

// --- API 1: GET HEATMAP & STATS ---
function getHeatmapData(selectedDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbIDP = ss.getSheetByName('DB_IDP');
  const dbSched = ss.getSheetByName('DB_SCHEDULE');
  const dbFurlough = ss.getSheetByName('DB_FURLOUGH');
  
  if (!dbIDP || !dbSched) throw new Error("Data missing. Please import IDP & Schedule.");

  const rotation = PropertiesService.getScriptProperties().getProperty('CURRENT_ROTATION') || 'Week A';

  // 1. Get Day Name (e.g., "Friday")
  const selDateObj = new Date(selectedDate.replace(/-/g, '/'));
  const days = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  const dayName = days[selDateObj.getDay()];

  // 2. Fetch Data
  const idpData = dbIDP.getDataRange().getValues();
  const schedData = dbSched.getDataRange().getValues();
  const furloughs = dbFurlough ? dbFurlough.getDataRange().getValues() : [];

  idpData.shift(); schedData.shift(); furloughs.shift();

  // 3. Init Buckets
  let buckets = Array.from({length: 96}, (_, i) => ({
    index: i,
    label: indexToTime(i),
    supply: 0,
    demand: 0,
    net: 0
  }));

  // 4. Map IDP (Demand & Supply Base)
  idpData.forEach(row => {
    if (row[0] === dayName) { 
      let idx = timeToBucketIndex(row[1]);
      if (idx > -1) {
        buckets[idx].demand += Number(row[2] || 0);
        buckets[idx].supply += Number(row[3] || 0);
      }
    }
  });

  // 5. Map Furloughs
  let furloughMap = {}; 
  furloughs.forEach(row => {
    let fDate = formatDate(row[2]);
    if (fDate === selectedDate) {
      let agent = String(row[1]).trim(); // TRIM FIX
      let fStart = timeToBucketIndex(row[3]);
      // If EndTime is missing or invalid, assume end of day (96)
      let fEnd = (row.length > 6 && row[6]) ? timeToBucketIndex(row[6]) : 96;
      if (fEnd === -1) fEnd = 96;
      
      // Handle Overnight Furlough Entry (e.g. 23:00 to 01:00) on current day view
      if (fEnd < fStart) fEnd = 96; 

      if (fStart > -1) {
        if (!furloughMap[agent]) furloughMap[agent] = [];
        furloughMap[agent].push({
          start: fStart,
          end: fEnd,
          startTimeStr: formatTime(row[3]),
          endTimeStr: (row.length > 6 && row[6]) ? formatTime(row[6]) : "End",
          rotation: row[5],
          id: row[0],
          intervalsSaved: 0
        });
      }
    }
  });

  // 6. Iterate Schedule (Supply Adjustment)
  const EXCLUSIONS = ['Break', 'Lunch', 'Off', 'Solicited', 'Sick', 'Maladie'];
  
  const prevDateObj = new Date(selDateObj);
  prevDateObj.setDate(selDateObj.getDate() - 1);
  const prevDateStr = Utilities.formatDate(prevDateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");

  schedData.forEach(row => {
    let agent = String(row[0]).trim(); // TRIM FIX
    let sDate = formatDate(row[1]);
    let act = row[2];
    
    if (EXCLUSIONS.some(ex => act.includes(ex))) return;

    let sStart = timeToBucketIndex(row[3]);
    let sEnd = timeToBucketIndex(row[4]);
    let isOvernight = sEnd < sStart;
    
    // Determine working segment for SELECTED DATE
    let segStart = -1, segEnd = -1;

    if (sDate === selectedDate) {
      segStart = sStart;
      segEnd = isOvernight ? 96 : sEnd; 
    } else if (sDate === prevDateStr && isOvernight) {
      segStart = 0;
      segEnd = sEnd;
    }

    // Intersection Calculation
    if (segStart > -1 && segEnd > -1 && furloughMap[agent]) {
      furloughMap[agent].forEach(f => {
        let overlapStart = Math.max(segStart, f.start);
        let overlapEnd = Math.min(segEnd, f.end);
        let overlap = overlapEnd - overlapStart;
        
        if (overlap > 0) {
          f.intervalsSaved += overlap;
          // Reduce Supply Grid
          for (let i = overlapStart; i < overlapEnd; i++) {
            if (buckets[i].supply > 0) buckets[i].supply -= 1;
          }
        }
      });
    }
  });

  // 7. Net Calculation
  buckets.forEach(b => {
    b.net = parseFloat((b.supply - b.demand).toFixed(2));
  });

  // 8. Prepare Log
  let logData = [];
  let totalHoursDay = 0;
  
  Object.keys(furloughMap).forEach(agent => {
    furloughMap[agent].forEach(f => {
      let hours = (f.intervalsSaved * 15) / 60;
      totalHoursDay += hours;
      logData.push({
        id: f.id,
        agent: agent,
        time: `${f.startTimeStr} - ${f.endTimeStr}`,
        rotation: f.rotation,
        hours: parseFloat(hours.toFixed(2))
      });
    });
  });

  return { 
    grid: buckets, 
    furloughs: logData, 
    rotation: rotation,
    totals: { day: parseFloat(totalHoursDay.toFixed(2)) }
  };
}

// --- API 2: SUBMIT FURLOUGH ---
function submitFurlough(agentName, dateStr, startTimeStr, endTimeStr) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('DB_FURLOUGH');
    if (!sheet) {
      sheet = ss.insertSheet('DB_FURLOUGH');
      sheet.appendRow(['ID', 'Agent Name', 'Date', 'Start Time', 'Type', 'WeekRotation', 'End Time']);
    }
    const rot = PropertiesService.getScriptProperties().getProperty('CURRENT_ROTATION') || 'Week A';
    
    // Trim agent name to ensure matches
    sheet.appendRow([
      new Date().getTime(), 
      agentName.trim(), 
      "'" + dateStr, 
      startTimeStr, 
      'Early Release', 
      rot,
      endTimeStr
    ]);
    return "Success";
  } catch (e) {
    return "Error: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

// --- UTILS ---
function formatDate(d) {
  if (!d) return "";
  if (typeof d === 'string' && d.match(/^\d{4}-\d{2}-\d{2}$/)) return d;
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), "yyyy-MM-dd");
}
function formatTime(d) {
  if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), "HH:mm");
  return String(d);
}
function timeToBucketIndex(val) {
  if (!val) return -1;
  if (val instanceof Date) return (val.getHours()*4) + Math.floor(val.getMinutes()/15);
  let parts = String(val).match(/(\d+):(\d+)\s?([AP]M)?/i);
  if (parts) {
    let h = parseInt(parts[1]);
    let m = parseInt(parts[2]);
    let amp = parts[3] ? parts[3].toUpperCase() : null;
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
function setRotation(r) { PropertiesService.getScriptProperties().setProperty('CURRENT_ROTATION', r); }
function importRawData(s, i) { return ImportHandler.processPaste(s, i); }
function getBucketDetails(d, i) { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSched = ss.getSheetByName('DB_SCHEDULE');
  if (!dbSched) return { time: indexToTime(i), agents: [] };
  const sched = dbSched.getDataRange().getValues(); sched.shift();
  let agents = [];
  const selDateObj = new Date(d.replace(/-/g, '/'));
  const prevDateObj = new Date(selDateObj); prevDateObj.setDate(selDateObj.getDate()-1);
  const prevDateStr = Utilities.formatDate(prevDateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const EXCLUSIONS = ['Break', 'Lunch', 'Off', 'Solicited', 'Sick', 'Maladie'];

  sched.forEach(row => {
    let act = row[2];
    if (EXCLUSIONS.some(ex => act.includes(ex))) return;
    let sDate = formatDate(row[1]);
    let start = timeToBucketIndex(row[3]);
    let end = timeToBucketIndex(row[4]);
    let overnight = end < start;
    let isActive = false;
    if (sDate === d) {
      let actualEnd = overnight ? 96 : end;
      if (i >= start && i < actualEnd) isActive = true;
    } else if (sDate === prevDateStr && overnight) {
      if (i >= 0 && i < end) isActive = true;
    }
    if (isActive) agents.push({ name: row[0], activity: act, shiftStart: formatTime(row[3]), shiftEnd: formatTime(row[4]) });
  });
  return { time: indexToTime(i), agents: agents };
}
