/**
 * WFM COMMAND CENTER - BACKEND v2.0
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

  if (!dbSched || !dbIDP) throw new Error("System not initialized.");

  // 1. Get Raw Data
  const schedData = dbSched.getDataRange().getValues();
  const idpData = dbIDP.getDataRange().getValues();
  // Furlough Data: [ID, Agent, Date, StartTime, Type]
  const furloughs = dbFurlough ? dbFurlough.getDataRange().getValues() : [];

  schedData.shift(); idpData.shift(); furloughs.shift();

  // 2. Init Buckets
  let buckets = Array.from({length: 96}, (_, i) => ({
    index: i,
    label: indexToTime(i),
    supply: 0,
    demand: 0,
    net: 0,
    furloughedCount: 0
  }));

  // 3. Map IDP (Demand)
  idpData.forEach(row => {
    if (formatDate(row[0]) === selectedDate) {
      let idx = timeToBucketIndex(row[1]);
      if (idx > -1) buckets[idx].demand += Number(row[2]);
    }
  });

  // 4. Map Schedule (Supply) - MINUS Furloughs
  const EXCLUSIONS = ['Break/Pause', 'Lunch/Repas', '(ADT) Repas payé / Paid Lunch'];
  
  schedData.forEach(row => {
    let sDate = formatDate(row[1]);
    let activity = row[2];
    let agentName = row[0];

    if (sDate === selectedDate && !EXCLUSIONS.includes(activity)) {
      let startIdx = timeToBucketIndex(row[3]);
      let endIdx = timeToBucketIndex(row[4]);

      // Check for Furlough (Early Release)
      // If agent is furloughed on this date from 14:00, then endIdx becomes 14:00
      let furloughCutoff = 96;
      let isFurloughed = false;

      furloughs.forEach(f => {
        if (f[1] === agentName && formatDate(f[2]) === selectedDate) {
          let fStart = timeToBucketIndex(f[3]);
          if (fStart > -1 && fStart < furloughCutoff) {
            furloughCutoff = fStart;
            isFurloughed = true;
          }
        }
      });

      // Clip End Time by Furlough
      if (endIdx > furloughCutoff) endIdx = furloughCutoff;

      if (startIdx > -1 && endIdx > -1) {
        if (endIdx < startIdx) endIdx = 96; // Midnight wrap
        for (let i = startIdx; i < endIdx; i++) {
          buckets[i].supply += 1;
        }
        // Track stats
        if(isFurloughed) {
           // We don't add to supply for the cut part, but we can track 'hours saved' elsewhere
        }
      }
    }
  });

  // 5. Calculate Net
  buckets.forEach(b => {
    b.net = parseFloat((b.supply - b.demand).toFixed(2));
    b.supply = parseFloat(b.supply.toFixed(2));
  });

  // 6. Get Furlough Log for UI
  let todayFurloughs = furloughs
    .filter(f => formatDate(f[2]) === selectedDate)
    .map(f => ({ agent: f[1], time: formatTime(f[3]), type: f[4] }));

  return { grid: buckets, furloughs: todayFurloughs };
}

// --- API 2: GET AGENT DETAILS FOR BUCKET ---
function getBucketDetails(date, bucketIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sched = ss.getSheetByName('DB_SCHEDULE').getDataRange().getValues();
  const timeLabel = indexToTime(bucketIndex);
  
  let agents = [];
  sched.shift();
  
  sched.forEach(row => {
    if (formatDate(row[1]) === date) {
      let start = timeToBucketIndex(row[3]);
      let end = timeToBucketIndex(row[4]);
      if (end < start) end = 96;

      // If this agent works during this bucket
      if (bucketIndex >= start && bucketIndex < end) {
        agents.push({
          name: row[0],
          activity: row[2],
          shiftStart: formatTime(row[3]),
          shiftEnd: formatTime(row[4])
        });
      }
    }
  });
  return { time: timeLabel, agents: agents };
}

// --- API 3: SUBMIT FURLOUGH ---
function submitFurlough(agentName, dateStr, timeStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('DB_FURLOUGH');
  if (!sheet) {
    sheet = ss.insertSheet('DB_FURLOUGH');
    sheet.appendRow(['ID', 'Agent Name', 'Date', 'Start Time', 'Type']);
  }
  
  const id = new Date().getTime();
  sheet.appendRow([id, agentName, dateStr, timeStr, 'Early Release']);
  return "Success";
}

// --- API 4: IMPORT DATA (Text Paste) ---
function importRawData(schedText, idpText) {
  return ImportHandler.processPaste(schedText, idpText);
}

// --- UTILS ---
function timeToBucketIndex(val) {
  if (val instanceof Date) return (val.getHours()*4) + Math.floor(val.getMinutes()/15);
  let d = new Date(`2000/01/01 ${val}`);
  return isNaN(d) ? -1 : (d.getHours()*4) + Math.floor(d.getMinutes()/15);
}
function indexToTime(i) {
  let h = Math.floor(i/4);
  let m = (i%4)*15;
  return `${h<10?'0'+h:h}:${m===0?'00':m}`;
}
function formatDate(d) {
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), "yyyy-MM-dd");
}
function formatTime(d) {
  if(d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), "HH:mm");
  return String(d);
}
