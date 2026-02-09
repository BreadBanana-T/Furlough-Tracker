var ImportHandler = {
  processPaste: function(schedRaw, idpRaw) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ['DB_SCHEDULE', 'DB_IDP', 'DB_FURLOUGH'].forEach(n => {
       if(!ss.getSheetByName(n)) ss.insertSheet(n);
    });

    let msg = [];

    // --- 1. SCHEDULE PARSER ---
    if (schedRaw && schedRaw.trim().length > 0) {
      let cleanSched = [];
      const lines = schedRaw.split(/\r?\n/).filter(l => l.trim().length > 0);
      let currentAgent = "", currentDate = "";
      
      // Matches: Activity Name (with noise) + Time + Time
      const segmentRegex = /([a-zA-ZÀ-ÿ0-9\/\(\)\s\-\.&]+?)\s+(\d{1,2}:\d{2}(?:\s?[AP]M)?)\s+(\d{1,2}:\d{2}(?:\s?[AP]M)?)\s*$/i;
      const dateRegex = /(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/;

      lines.forEach(line => {
        let text = line.trim();
        if(text.startsWith('Agent Name') || text.startsWith('"Agent Name"')) return;

        // A. Agent
        if (text.includes('Agent:')) {
          let parts = text.split(':');
          if (parts.length > 1) currentAgent = parts[1].replace(/^\s*\d+\s+/, '').trim();
          return;
        } 
        // B. CSV-style Agent (If pasting rows like "Name",Date...)
        else if (text.includes('"') && text.includes(',')) {
          let csvParts = parseCSVLine(text);
          if (csvParts.length >= 5) {
            let act = cleanActivity(csvParts[2]);
            cleanSched.push([csvParts[0], parseDate(csvParts[1]), act, csvParts[3], csvParts[4]]);
            return; 
          }
        }

        // C. Date
        let dMatch = text.match(dateRegex);
        if (dMatch) currentDate = parseDate(dMatch[1]);

        // D. Activity Segment
        if (currentAgent && currentDate) {
          let segMatch = text.match(segmentRegex);
          if (segMatch) {
            let rawAct = segMatch[1].trim();
            // Clean "00 PM" noise specifically seen in your file
            let act = cleanActivity(rawAct);
            
            if (!act.toLowerCase().includes('activity') && !act.toLowerCase().includes('scheduled')) {
               cleanSched.push([currentAgent, currentDate, act, segMatch[2].trim(), segMatch[3].trim()]);
            }
          }
        }
      });

      if (cleanSched.length > 0) {
        const db = ss.getSheetByName('DB_SCHEDULE');
        db.clear();
        db.appendRow(['Agent Name', 'Date', 'Activity', 'Start Time', 'End Time']);
        db.getRange(2,1,cleanSched.length,5).setValues(cleanSched);
        msg.push(`✔ Schedule: Imported ${cleanSched.length} rows.`);
      }
    }

    // --- 2. IDP PARSER (Combined Headers) ---
    if (idpRaw && idpRaw.trim().length > 0) {
      let cleanIDP = [];
      const lines = idpRaw.split(/\r?\n/).filter(l => l.trim().length > 0);
      
      // Look for the header row with "Requirements" and "Open"
      let headerIdx = -1;
      for (let i = 0; i < lines.length; i++) {
        if (lines[i].includes('Requirements') && lines[i].includes('Open')) {
          headerIdx = i;
          break;
        }
      }

      if (headerIdx > -1) {
        let headers = parseCSVLine(lines[headerIdx]);
        let colMap = {}; 

        headers.forEach((h, i) => {
          let lower = h.toLowerCase();
          // Extract Date from string like "Requirements Friday, February 6, 2026"
          let dateMatch = h.match(/(\w+\s\d{1,2},\s\d{4})/);
          if (dateMatch) {
            let dateStr = parseDate(dateMatch[1]);
            // Map columns
            if (lower.includes('req')) colMap[i] = { date: dateStr, type: 'req' };
            else if (lower.includes('open') && !lower.includes('+/-')) colMap[i] = { date: dateStr, type: 'open' };
          }
        });

        // Parse Data Rows
        let dataByDay = {};
        for (let i = headerIdx + 1; i < lines.length; i++) {
          let cols = parseCSVLine(lines[i]);
          let timeStr = cols[0]; // Interval (e.g. 00:00:00) is Col 0
          
          if (timeStr && timeStr.includes(':')) {
            let tNorm = formatTimeStr(timeStr);
            
            Object.keys(colMap).forEach(idx => {
               if (cols[idx]) {
                 let info = colMap[idx];
                 if (!dataByDay[info.date]) dataByDay[info.date] = {};
                 if (!dataByDay[info.date][tNorm]) dataByDay[info.date][tNorm] = { req:0, open:0 };
                 
                 let val = parseFloat(cols[idx]);
                 if (isNaN(val)) val = 0;
                 if (info.type === 'req') dataByDay[info.date][tNorm].req = val;
                 if (info.type === 'open') dataByDay[info.date][tNorm].open = val;
               }
            });
          }
        }
        
        // Flatten
        Object.keys(dataByDay).forEach(day => {
          Object.keys(dataByDay[day]).forEach(time => {
            cleanIDP.push([day, time, dataByDay[day][time].req, dataByDay[day][time].open]);
          });
        });

        if (cleanIDP.length > 0) {
          const db = ss.getSheetByName('DB_IDP');
          db.clear();
          db.appendRow(['Day', 'Interval', 'Required', 'Open']);
          db.getRange(2,1,cleanIDP.length,4).setValues(cleanIDP);
          msg.push(`✔ IDP: Imported ${cleanIDP.length} intervals.`);
        }
      } else {
        msg.push("❌ IDP: Could not find header row with 'Requirements' and 'Open'. Check selection.");
      }
    }
    
    return msg.length ? msg.join('\n') : "No valid data found to import.";
  }
};

// Utils
function parseDate(s) { 
  if(!s) return "";
  let d = new Date(s);
  return isNaN(d.getTime()) ? "" : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function formatTimeStr(t) {
  let d = new Date(`2000/01/01 ${t}`);
  return isNaN(d.getTime()) ? t : Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm');
}
function cleanActivity(s) {
  // Removes "00 PM" artifacts from strings
  return s.replace(/\d{2}\s?[AP]M/gi, '').trim();
}
// Robust CSV line parser handles quotes
function parseCSVLine(text) {
  let ret = [];
  let inQuote = false;
  let token = "";
  for(let i=0; i<text.length; i++) {
    let char = text[i];
    if(char === '"') { inQuote = !inQuote; continue; }
    if(char === ',' && !inQuote) { ret.push(token); token = ""; }
    else { token += char; }
  }
  ret.push(token);
  return ret;
}
