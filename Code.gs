/**
 * WFM COMMAND CENTER - BACKEND v9.0
 * Verified for: Stacked IDP Headers, Strict Date Matching, Supply/Demand Logic
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Furlough Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- API: GET HEATMAP ---
function getHeatmapData(selectedDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbIDP = ss.getSheetByName('DB_IDP');
  const dbSched = ss.getSheetByName('DB_SCHEDULE');
  const dbFurlough = ss.getSheetByName('DB_FURLOUGH');

  if (!dbIDP || !dbSched) throw new Error("Database missing. Please click 'Feed Data' to import.");

  // Get Rotation (Week A/B)
  const rotation = PropertiesService.getScriptProperties().getProperty('CURRENT_ROTATION') || 'Week A';
  
  // Normalize User Date (YYYY-MM-DD)
  const selDateStr = selectedDate; 

  // Fetch Data
  const idpData = dbIDP.getDataRange().getValues();
  const schedData = dbSched.getDataRange().getValues();
  const furloughs = dbFurlough ? dbFurlough.getDataRange().getValues() : [];

  // Remove Headers
  idpData.shift(); schedData.shift(); furloughs.shift();

  // Init 96 Buckets (15 min intervals)
  let buckets = Array.from({length: 96}, (_, i) => ({
    index: i,
    label: indexToTime(i),
    supply: 0,
    demand: 0,
    net: 0
  }));

  // --- 1. MAP IDP (Supply & Demand) ---
  idpData.forEach(row => {
    // Row: [Date, Interval, Required, Open/Supply]
    let rowDate = formatDate(row[0]); 
    
    // Strict Match: Database Date vs Selected Date
    if (rowDate === selDateStr) { 
      let idx = timeToBucketIndex(row[1]);
      if (idx > -1) {
        buckets[idx].demand += Number(row[2] || 0); // "Requirements"
        buckets[idx].supply += Number(row[3] || 0); // "Open" (treated as Staffed/Supply)
      }
    }
  });

  // --- 2. MAP FURLOUGHS (Subtract from Supply) ---
  let furloughMap = {}; 
  furloughs.forEach(row => {
    let fDate = formatDate(row[2]);
    if (fDate === selDateStr) {
      let agent = String(row[1]).trim();
      let fStart = timeToBucketIndex(row[3]);
      // Handle missing End Time or Wrap
      let fEnd = (row.length > 6 && row[6]) ? timeToBucketIndex(row[6]) : 96;
      if (fEnd === -1 || fEnd < fStart) fEnd = 96;

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

  // --- 3. CALCULATE SAVINGS (Schedule Overlap) ---
  const EXCLUSIONS = ['Break', 'Lunch', 'Off', 'Solicited', 'Sick', 'Maladie'];
  
  schedData.forEach(row => {
    let agent = String(row[0]).trim();
    let sDate = formatDate(row[1]);
    let act = row[2];
    
    // Skip non-productive codes
    if (EXCLUSIONS.some(ex => act.includes(ex))) return;

    if (sDate === selDateStr && furloughMap[agent]) {
      let sStart = timeToBucketIndex(row[3]);
      let sEnd = timeToBucketIndex(row[4]);
      if (sEnd < sStart) sEnd = 96; // Cap overnight for single-day view

      furloughMap[agent].forEach(f => {
        let overlapStart = Math.max(sStart, f.start);
        let overlapEnd = Math.min(sEnd, f.end);
        
        if (overlapEnd > overlapStart) {
          f.intervalsSaved += (overlapEnd - overlapStart);
          // Reduce Supply in Grid
          for (let i = overlapStart; i < overlapEnd; i++) {
            buckets[i].supply = Math.max(0, buckets[i].supply - 1);
          }
        }
      });
    }
  });

  // --- 4. CALCULATE NET ---
  buckets.forEach(b => {
    b.net = parseFloat((b.supply - b.demand).toFixed(2));
  });

  // --- 5. LOG PREP ---
  let logData = [];
  let totalHoursDay = 0;
  Object.keys(furloughMap).forEach(agent => {
    furloughMap[agent].forEach(f => {
      let hours = (f.intervalsSaved * 15) / 60;
      if (hours > 0) {
        totalHoursDay += hours;
        logData.push({
          id: f.id,
          agent: agent,
          time: `${f.startTimeStr} - ${f.endTimeStr}`,
          rotation: f.rotation,
          hours: parseFloat(hours.toFixed(2))
        });
      }
    });
  });

  return { 
    grid: buckets, 
    furloughs: logData, 
    rotation: rotation,
    totals: { day: parseFloat(totalHoursDay.toFixed(2)) }
  };
}

// --- UTILS ---
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

// Fetch agents for the Modal
function getBucketDetails(d, i) { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSched = ss.getSheetByName('DB_SCHEDULE');
  if (!dbSched) return { time: indexToTime(i), agents: [] };
  
  const sched = dbSched.getDataRange().getValues(); sched.shift();
  let agents = [];
  const EXCLUSIONS = ['Break', 'Lunch', 'Off', 'Solicited', 'Sick', 'Maladie'];
  
  sched.forEach(row => {
    let act = row[2];
    if (EXCLUSIONS.some(ex => act.includes(ex))) return;
    
    let sDate = formatDate(row[1]);
    let start = timeToBucketIndex(row[3]);
    let end = timeToBucketIndex(row[4]);
    let isActive = false;
    
    if (sDate === d) {
       let actualEnd = (end < start) ? 96 : end;
       if (i >= start && i < actualEnd) isActive = true;
    }
    
    if (isActive) {
      agents.push({ 
        name: row[0], 
        activity: act, 
        shiftStart: formatTime(row[3]), 
        shiftEnd: formatTime(row[4]) 
      });
    }
  });
  return { time: indexToTime(i), agents: agents };
}
