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
      const segmentRegex = /([a-zA-ZÀ-ÿ0-9\/\(\)\s\-\.&]+?)\s+(\d{1,2}:\d{2}(?:\s?[AP]M)?)\s+(\d{1,2}:\d{2}(?:\s?[AP]M)?)\s*$/i;
      const dateRegex = /(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/;

      lines.forEach(line => {
        let text = line.trim();
        if(text.startsWith('Agent Name') || text.startsWith('"Agent Name"')) return;

        if (text.includes('Agent:')) {
          let parts = text.split(':');
          if (parts.length > 1) currentAgent = parts[1].replace(/^\s*\d+\s+/, '').trim();
          return;
        } 
        else if (text.includes('"') && text.includes(',')) {
          let csvParts = parseCSVLine(text);
          if (csvParts.length >= 5) {
            let act = cleanActivity(csvParts[2]);
            cleanSched.push([csvParts[0], parseDate(csvParts[1]), act, csvParts[3], csvParts[4]]);
            return; 
          }
        }

        let dMatch = text.match(dateRegex);
        if (dMatch) currentDate = parseDate(dMatch[1]);

        if (currentAgent && currentDate) {
          let segMatch = text.match(segmentRegex);
          if (segMatch) {
            let rawAct = segMatch[1].trim();
            let act = cleanActivity(rawAct);
            if (!act.toLowerCase().match(/^activity|^scheduled/)) {
               cleanSched.push([currentAgent, currentDate, act, segMatch[2].trim(), segMatch[3].trim()]);
            }
          }
        }
      });

      if (cleanSched.length > 0) {
        upsertHistoricalData('DB_SCHEDULE', cleanSched, 1, ['Agent Name', 'Date', 'Activity', 'Start Time', 'End Time']);
        msg.push(`✔ Schedule: Imported/Updated ${cleanSched.length} rows.`);
      }
    }

    // --- 2. IDP PARSER (Excel TSV Support) ---
    if (idpRaw && idpRaw.trim().length > 0) {
      let cleanIDP = [];
      const lines = idpRaw.split(/\r?\n/).filter(l => l.trim().length > 0);
      let headerIdx = -1;
      
      for (let i = 0; i < lines.length; i++) {
        let lowerLine = lines[i].toLowerCase();
        if (lowerLine.includes('req') && lowerLine.includes('open')) {
          headerIdx = i;
          break;
        }
      }

      if (headerIdx > -1) {
        let headers = parseCSVLine(lines[headerIdx]);
        let colMap = {}; 

        headers.forEach((h, i) => {
          let lower = h.toLowerCase();
          let dateMatch = h.match(/(\w+\s\d{1,2},?\s\d{4})/);
          if (dateMatch) {
            let dateStr = parseDate(dateMatch[1]);
            if (lower.includes('req')) colMap[i] = { date: dateStr, type: 'req' };
            else if (lower.includes('open') && !lower.includes('+/-')) colMap[i] = { date: dateStr, type: 'open' };
          }
        });

        let dataByDay = {};
        for (let i = headerIdx + 1; i < lines.length; i++) {
          let cols = parseCSVLine(lines[i]);
          let timeStr = cols[0]; 
          
          if (timeStr && timeStr.includes(':')) {
            let tNorm = formatTimeStr(timeStr);
            Object.keys(colMap).forEach(idx => {
               if (cols[idx] !== undefined) {
                 let info = colMap[idx];
                 if (!dataByDay[info.date]) dataByDay[info.date] = {};
                 if (!dataByDay[info.date][tNorm]) dataByDay[info.date][tNorm] = { req:0, open:0 };
                
                 // Clean numbers from Excel formatting
                 let rawVal = String(cols[idx]).replace(/,/g, '');
                 let val = parseFloat(rawVal);
                 if (isNaN(val)) val = 0;
                 if (info.type === 'req') dataByDay[info.date][tNorm].req = val;
                 if (info.type === 'open') dataByDay[info.date][tNorm].open = val;
               }
            });
          }
        }
        
        Object.keys(dataByDay).forEach(day => {
          Object.keys(dataByDay[day]).forEach(time => {
            cleanIDP.push([day, time, dataByDay[day][time].req, dataByDay[day][time].open]);
          });
        });

        if (cleanIDP.length > 0) {
          upsertHistoricalData('DB_IDP', cleanIDP, 0, ['Day', 'Interval', 'Required', 'Open']);
          msg.push(`✔ IDP: Imported/Updated ${cleanIDP.length} intervals.`);
        }
      } else {
        msg.push("❌ IDP: Could not find header row with 'Requirements' and 'Open'. Check selection.");
      }
    }
    
    return msg.length ? msg.join('\n') : "No valid data found to import.";
  }
};

// --- CORE UTILS ---

// Upsert Logic: Removes old data ONLY for the dates being imported, preserving history.
function upsertHistoricalData(sheetName, newRows, dateColIdx, headersArray) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  // Extract unique dates from incoming data
  const incomingDates = new Set(newRows.map(row => String(row[dateColIdx]).trim()));
  
  const existingData = sheet.getDataRange().getValues();
  const headers = existingData.length > 0 ? existingData.shift() : headersArray;
  
  // Keep rows that do NOT belong to the incoming dates
  const retainedRows = existingData.filter(row => {
     let rowDate = parseDate(row[dateColIdx]);
     return !incomingDates.has(String(rowDate).trim());
  });
  
  const combined = retainedRows.concat(newRows);
  
  sheet.clearContents();
  sheet.appendRow(headers);
  if (combined.length > 0) {
      sheet.getRange(2, 1, combined.length, combined[0].length).setValues(combined);
  }
}

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
  return s.replace(/\d{2}\s?[AP]M/gi, '').trim();
}

// Upgraded to handle Excel Tabs (TSV) natively
function parseCSVLine(text) {
  if (text.includes('\t')) {
     return text.split('\t').map(s => s.trim());
  }
  let ret = [];
  let inQuote = false;
  let token = "";
  for(let i=0; i<text.length; i++) {
    let char = text[i];
    if(char === '"') { inQuote = !inQuote; continue; }
    if(char === ',' && !inQuote) { ret.push(token.trim()); token = ""; }
    else { token += char; }
  }
  ret.push(token.trim());
  return ret;
}
