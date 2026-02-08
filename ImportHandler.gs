var ImportHandler = {
  
  processPaste: function(schedRaw, idpRaw) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Ensure DBs exist
    ['DB_SCHEDULE', 'DB_IDP', 'DB_FURLOUGH'].forEach(n => {
      if(!ss.getSheetByName(n)) ss.insertSheet(n);
    });
    
    // --- 1. PROCESS SCHEDULE ---
    if (schedRaw && schedRaw.trim().length > 0) {
      const rows = parseCSV(schedRaw);
      let cleanSched = [];
      let curAgent = "", curDate = "";
      
      rows.forEach(row => {
        let col0 = row[0];
        // Agent Header Detection
        if (col0 && col0.includes('Agent:')) {
           let parts = col0.split(':');
           if(parts.length > 1) curAgent = parts[1].replace(/^\d+/, '').trim(); // Remove ID
        }
        
        // Data Row Detection (Col 2=Date, 6=Act, 7=Start, 10=End)
        // Adjust indices for 0-based array from CSV parser
        if (row.length > 10) {
          let dateVal = row[2];
          let act = row[6];
          let start = row[7];
          let end = row[10];

          if (act && start && end) {
             if (dateVal && dateVal.includes('/')) curDate = parseDate(dateVal);
             
             if (curAgent && curDate) {
               cleanSched.push([curAgent, curDate, act, start, end]);
             }
          }
        }
      });
      
      // Save to DB
      if (cleanSched.length > 0) {
        const db = ss.getSheetByName('DB_SCHEDULE');
        db.clear();
        db.appendRow(['Agent Name', 'Date', 'Activity', 'Start Time', 'End Time']);
        db.getRange(2,1,cleanSched.length,5).setValues(cleanSched);
      }
    }

    // --- 2. PROCESS IDP ---
    if (idpRaw && idpRaw.trim().length > 0) {
      const rows = parseCSV(idpRaw);
      let cleanIDP = [];
      
      // Find Header
      let hIdx = rows.findIndex(r => r[0] && r[0].toLowerCase().includes('time'));
      if (hIdx > -1) {
        let headers = rows[hIdx];
        let dateMap = {};
        
        headers.forEach((h, i) => {
          if (h && h.includes('Requirements')) {
             let dStr = h.replace('Requirements', '').trim();
             dateMap[i] = parseDateLong(dStr); // Need custom parser for "Friday, Feb..."
          }
        });
        
        for (let i = hIdx+1; i < rows.length; i++) {
          let time = rows[i][0];
          Object.keys(dateMap).forEach(k => {
            let val = rows[i][k];
            if (val) cleanIDP.push([dateMap[k], time, val]);
          });
        }
      }
      
      // Save to DB
      if (cleanIDP.length > 0) {
        const db = ss.getSheetByName('DB_IDP');
        db.clear();
        db.appendRow(['Date', 'Interval', 'Required']);
        db.getRange(2,1,cleanIDP.length,3).setValues(cleanIDP);
      }
    }

    return "Import Successful";
  }
};

// Simple CSV Parser handling quotes
function parseCSV(str) {
  const arr = [];
  let quote = false;  // 'true' means we're inside a quoted field
  let row = 0, col = 0, c = 0;
  for (; c < str.length; c++) {
    let cc = str[c], nc = str[c+1];
    arr[row] = arr[row] || []; arr[row][col] = arr[row][col] || '';
    if (cc == '"' && quote && nc == '"') { arr[row][col] += cc; ++c; continue; }
    if (cc == '"') { quote = !quote; continue; }
    if (cc == ',' && !quote) { ++col; continue; }
    if (cc == '\r' && nc == '\n' && !quote) { ++row; col = 0; ++c; continue; }
    if (cc == '\n' && !quote) { ++row; col = 0; continue; }
    if (cc == '\r' && !quote) { ++row; col = 0; continue; }
    arr[row][col] += cc;
  }
  return arr;
}

function parseDate(s) {
  // 2/8/26 -> 2026-02-08
  if(!s) return "";
  let p = s.split('/');
  return `20${p[2]}-${p[0].padStart(2,'0')}-${p[1].padStart(2,'0')}`;
}

function parseDateLong(s) {
  // "Friday, February 6, 2026" -> YYYY-MM-DD
  let d = new Date(s);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
