/**
 * Furlough Tracker - BACKEND
 * Features: Range Analytics (Day/Week/Month/Qtr), Precision Math, Auto-Detection
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

// --- API: GET ANALYTICS ---
function getAnalyticsData(mode, refDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbIDP = ss.getSheetByName('DB_IDP');
  const dbSched = ss.getSheetByName('DB_SCHEDULE');
  const dbFurlough = ss.getSheetByName('DB_FURLOUGH');

  if (!dbIDP || !dbSched) throw new Error("Database missing. Feed Data first.");

  // 1. CALCULATE DATE RANGE
  const dateObj = new Date(refDate + 'T00:00:00');
  let startDate = new Date(dateObj);
  let endDate = new Date(dateObj);
  let label = "";

  if (mode === 'day') {
    // Single Day
    label = formatDate(startDate);
  } else if (mode === 'week') {
    // Start of week (Monday)
    let day = startDate.getDay();
    let diff = startDate.getDate() - day + (day == 0 ? -6 : 1); 
    startDate.setDate(diff);
    endDate = new Date(startDate);
    endDate.setDate(startDate.getDate() + 6);
    label = `Week of ${formatDate(startDate)}`;
  } else if (mode === 'month') {
    startDate.setDate(1);
    endDate = new Date(startDate.getFullYear(), startDate.getMonth() + 1, 0);
    label = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MMMM yyyy");
  } else if (mode === 'quarter') {
    const q = Math.floor((startDate.getMonth() + 3) / 3);
    startDate = new Date(startDate.getFullYear(), (q - 1) * 3, 1);
    endDate = new Date(startDate.getFullYear(), startDate.getMonth() + 3, 0);
    label = `Q${q} ${startDate.getFullYear()}`;
  }

  const startStr = formatDate(startDate);
  const endStr = formatDate(endDate);

  // 2. FETCH RAW DATA
  const idpData = dbIDP.getDataRange().getValues();
  const schedData = dbSched.getDataRange().getValues();
  const furloughsManual = dbFurlough ? dbFurlough.getDataRange().getValues() : [];

  idpData.shift(); schedData.shift(); furloughsManual.shift();

  // 3. INIT BUCKETS (Only relevant for Day view)
  let buckets = [];
  if (mode === 'day') {
    buckets = Array.from({length: 96}, (_, i) => ({
      index: i,
      label: indexToTime(i),
      supply: 0,
      demand: 0,
      net: 0
    }));
    
    // Fill Supply/Demand for single day
    idpData.forEach(row => {
      if (formatDate(row[0]) === startStr) { 
        let idx = timeToBucketIndex(row[1]);
        if (idx > -1) {
          buckets[idx].demand += Number(row[2] || 0);
          buckets[idx].supply += Number(row[3] || 0);
        }
      }
    });
  }

  // 4. PROCESS FURLOUGHS (Over the Range)
  let combinedFurloughs = [];
  const ACSU_CODES = ['acsu', 'solicited', 'libération', 'voluntary'];
  let productiveMap = {};
  const EXCLUSIONS = ['break', 'lunch', 'off', 'sick', 'maladie', ...ACSU_CODES];

  // A. Build Productive Map & Auto-Detect ACSU
  schedData.forEach(row => {
    let sDateStr = formatDate(row[1]);
    
    // Check if row date is within range
    if (sDateStr >= startStr && sDateStr <= endStr) {
        let agent = String(row[0]).trim();
        let act = String(row[2]).toLowerCase();
        let sStart = timeToBucketIndex(row[3]);
        let sEnd = timeToBucketIndex(row[4]);
        if (sEnd < sStart) sEnd = 96;

        // Auto-Detect ACSU
        if (ACSU_CODES.some(c => act.includes(c))) {
             let exactHours = getDurationInHours(row[3], row[4]);
             let shiftType = getShiftType(sStart);
             
             combinedFurloughs.push({
               type: 'auto',
               date: sDateStr,
               agent: row[0],
               time: `${formatTime(row[3])} - ${formatTime(row[4])}`,
               hours: exactHours,
               shift: shiftType,
               startIdx: sStart,
               endIdx: sEnd
             });
        } 
        // Build Productive Map for Manual Overlap Check
        else if (!EXCLUSIONS.some(ex => act.includes(ex))) {
             if (!productiveMap[sDateStr]) productiveMap[sDateStr] = {};
             if (!productiveMap[sDateStr][agent]) productiveMap[sDateStr][agent] = [];
             productiveMap[sDateStr][agent].push({ start: sStart, end: sEnd });
        }
    }
  });

  // B. Process Manual Furloughs
  furloughsManual.forEach(row => {
    let fDateStr = formatDate(row[2]);
    if (fDateStr >= startStr && fDateStr <= endStr) {
        let agent = String(row[1]).trim();
        let fStart = timeToBucketIndex(row[3]);
        let fEnd = (row.length > 6 && row[6]) ? timeToBucketIndex(row[6]) : 96;
        if (fEnd < fStart) fEnd = 96;

        let intervalsSaved = 0;
        if (productiveMap[fDateStr] && productiveMap[fDateStr][agent]) {
            productiveMap[fDateStr][agent].forEach(shift => {
                let overlapStart = Math.max(fStart, shift.start);
                let overlapEnd = Math.min(fEnd, shift.end);
                if (overlapEnd > overlapStart) {
                   intervalsSaved += (overlapEnd - overlapStart);
                   // Update Grid Supply ONLY if in Day mode
                   if (mode === 'day') {
                      for (let i = overlapStart; i < overlapEnd; i++) {
                         buckets[i].supply = Math.max(0, buckets[i].supply - 1);
                      }
                   }
                }
            });
        }

        if (intervalsSaved > 0) {
            let shiftType = getShiftType(fStart);
            combinedFurloughs.push({
                type: 'manual',
                date: fDateStr,
                agent: agent,
                time: `${formatTime(row[3])} - ${(row.length > 6 && row[6]) ? formatTime(row[6]) : "End"}`,
                hours: (intervalsSaved * 15) / 60,
                shift: shiftType,
                startIdx: fStart,
                endIdx: fEnd
            });
        }
    }
  });

  // C. Update Grid Supply for Auto Furloughs (Day Mode Only)
  if (mode === 'day') {
     combinedFurloughs.filter(f => f.type === 'auto').forEach(f => {
        for (let i = f.startIdx; i < f.endIdx; i++) {
           if (i >= 0 && i < 96) buckets[i].supply = Math.max(0, buckets[i].supply - 1);
        }
     });
     // Final Net Calc
     buckets.forEach(b => { b.net = parseFloat((b.supply - b.demand).toFixed(2)); });
  }

  // 5. AGGREGATE TOTALS
  let totals = { all: 0, morning: 0, evening: 0, night: 0, count: combinedFurloughs.length };
  combinedFurloughs.forEach(f => {
      totals.all += f.hours;
      if (f.shift === 'Morning') totals.morning += f.hours;
      else if (f.shift === 'Evening') totals.evening += f.hours;
      else totals.night += f.hours;
  });

  // Rounding
  totals.all = parseFloat(totals.all.toFixed(2));
  totals.morning = parseFloat(totals.morning.toFixed(2));
  totals.evening = parseFloat(totals.evening.toFixed(2));
  totals.night = parseFloat(totals.night.toFixed(2));

  return {
    mode: mode,
    label: label,
    grid: buckets, // Empty if not 'day'
    furloughs: combinedFurloughs, // List for the whole period
    totals: totals,
    rotation: PropertiesService.getScriptProperties().getProperty('CURRENT_ROTATION') || 'Week A'
  };
}

// --- UTILS ---
function getShiftType(idx) {
  if (idx >= 28 && idx < 60) return 'Morning'; // 07:00 - 15:00
  if (idx >= 60 && idx < 92) return 'Evening'; // 15:00 - 23:00
  return 'Night';
}

function getDurationInHours(t1, t2) {
  let d1 = (t1 instanceof Date) ? t1 : parseTimeObj(t1);
  let d2 = (t2 instanceof Date) ? t2 : parseTimeObj(t2);
  if (!d1 || !d2) return 0;
  let diffMs = d2 - d1;
  if (diffMs < 0) diffMs += (24 * 60 * 60 * 1000);
  return parseFloat((diffMs / (1000 * 60 * 60)).toFixed(2));
}

function parseTimeObj(str) {
  if (!str) return null;
  let d = new Date();
  let parts = String(str).match(/(\d+):(\d+)\s?([AP]M)?/i);
  if (parts) {
    let h = parseInt(parts[1]);
    let m = parseInt(parts[2]);
    let amp = parts[3] ? parts[3].toUpperCase() : null;
    if (amp === 'PM' && h < 12) h += 12;
    if (amp === 'AM' && h === 12) h = 0;
    d.setHours(h, m, 0, 0);
    return d;
  }
  return null;
}

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
    sheet.appendRow([new Date().getTime(), agentName.trim(), "'" + dateStr, startTimeStr, 'Early Release', rot, endTimeStr]);
    return "Success";
  } catch (e) { return "Error: " + e.message; } 
  finally { lock.releaseLock(); }
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
function getBucketDetails(d, i) { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSched = ss.getSheetByName('DB_SCHEDULE');
  if (!dbSched) return { time: indexToTime(i), agents: [] };
  const sched = dbSched.getDataRange().getValues(); sched.shift();
  let agents = [];
  const EXCLUSIONS = ['Break', 'Lunch', 'Off', 'Solicited', 'Sick', 'Maladie', 'ACSU', 'Libération'];
  sched.forEach(row => {
    let act = row[2];
    if (EXCLUSIONS.some(ex => String(act).includes(ex))) return;
    let sDate = formatDate(row[1]);
    let start = timeToBucketIndex(row[3]);
    let end = timeToBucketIndex(row[4]);
    let isActive = false;
    if (sDate === d) {
       let actualEnd = (end < start) ? 96 : end;
       if (i >= start && i < actualEnd) isActive = true;
    }
    if (isActive) agents.push({ name: row[0], activity: act, shiftStart: formatTime(row[3]), shiftEnd: formatTime(row[4]) });
  });
  return { time: indexToTime(i), agents: agents };
}
