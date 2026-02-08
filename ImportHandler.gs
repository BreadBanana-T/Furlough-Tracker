var ImportHandler = {
  
  processPaste: function(schedRaw, idpRaw) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Ensure DBs exist
    if(!ss.getSheetByName('DB_SCHEDULE')) ss.insertSheet('DB_SCHEDULE');
    if(!ss.getSheetByName('DB_IDP')) ss.insertSheet('DB_IDP');
    if(!ss.getSheetByName('DB_FURLOUGH')) {
      let s = ss.insertSheet('DB_FURLOUGH');
      s.appendRow(['ID', 'Agent Name', 'Date', 'Start Time', 'Type', 'WeekRotation']);
    }

    // --- 1. PROCESS SCHEDULE (HIERARCHICAL TEXT PARSING) ---
    if (schedRaw && schedRaw.trim().length > 0) {
      let cleanSched = [];
      
      // Split by newline and remove empty lines
      const lines = schedRaw.split(/\r?\n/).filter(l => l.trim().length > 0);
      
      let currentAgent = "";
      let currentDate = "";
      
      // Regex to find "Activity StartTime EndTime" at the end of a line
      // Matches: "Open/Ouvert 6:30 AM 8:20 AM"
      // Captures group 1: Activity, group 2: Start, group 3: End
      const segmentRegex = /([a-zA-ZÀ-ÿ0-9\/\(\)\s\-\.]+)[\t\s]+(\d{1,2}:\d{2}\s?[AP]M)[\t\s]+(\d{1,2}:\d{2}\s?[AP]M)\s*$/i;
      
      // Regex to find Date at start of line (e.g. 2/8/26)
      const dateRegex = /^[\t\s]*(\d{1,2}\/\d{1,2}\/\d{2,4})/;

      lines.forEach(line => {
        const text = line.trim();
        
        // 1. Detect Agent
        if (text.includes('Agent:')) {
          const parts = text.split(':');
          if (parts.length > 1) {
            // Remove ID numbers if present (e.g. "773560 Abouda, Mohammed" -> "Abouda, Mohammed")
            currentAgent = parts[1].replace(/^\s*\d+\s+/, '').trim();
          }
          return; // Skip to next line
        }
        
        // 2. Detect Date (Update context)
        const dateMatch = line.match(dateRegex);
        if (dateMatch) {
          currentDate = parseDate(dateMatch[1]);
        }
        
        // 3. Detect Activity Segments
        // We only care if we have an agent and a date context
        if (currentAgent && currentDate) {
          const segMatch = line.match(segmentRegex);
          if (segMatch) {
            let activity = segMatch[1].trim();
            let start = segMatch[2].trim();
            let end = segMatch[3].trim();
            
            // Cleanup activity name (sometimes trailing tabs get caught)
            // Also, your raw data sometimes has "Date Start End" before the activity.
            // The regex matches the *last* 3 parts which is usually the segment breakdown.
            
            // Filter out lines that are just headers like "Date Start End"
            if (!activity.toLowerCase().includes('scheduled activity') && 
                !start.toLowerCase().includes('start') &&
                !end.toLowerCase().includes('end')) {
                  
               cleanSched.push([currentAgent, currentDate, activity, start, end]);
            }
          }
        }
      });
      
      // Save to DB_SCHEDULE
      if (cleanSched.length > 0) {
        const db = ss.getSheetByName('DB_SCHEDULE');
        db.clear();
        db.appendRow(['Agent Name', 'Date', 'Activity', 'Start Time', 'End Time']);
        db.getRange(2,1,cleanSched.length,5).setValues(cleanSched);
      }
    }

    // --- 2. PROCESS IDP (COLUMN MAPPING) ---
    if (idpRaw && idpRaw.trim().length > 0) {
      let cleanIDP = [];
      const lines = idpRaw.split(/\r?\n/).filter(l => l.trim().length > 0);
      
      // Find header row index
      let hIdx = lines.findIndex(l => l.toLowerCase().includes('time') && l.toLowerCase().includes('requirements'));
      
      if (hIdx > -1) {
        // Split header by tab (preferred) or multiple spaces
        // The raw data looks tab-separated
        let headerLine = lines[hIdx];
        let headers = headerLine.split(/\t+/);
        if(headers.length < 2) headers = headerLine.split(/\s{2,}/); // Fallback to spaces
        
        let dateMap = {}; // Col Index -> Date String
        
        // Map columns to dates
        headers.forEach((h, i) => {
          if (h.toLowerCase().includes('requirements')) {
             // Extract "Friday, February 6, 2026"
             let dStr = h.replace(/requirements/i, '').trim();
             let pDate = parseDateLong(dStr);
             if(pDate) dateMap[i] = pDate;
          }
        });

        // Process Data Rows
        for (let i = hIdx + 1; i < lines.length; i++) {
          let line = lines[i];
          let cols = line.split(/\t+/);
          if(cols.length < 2) cols = line.split(/\s{2,}/); // Fallback
          
          let time = cols[0];
          // Basic validation that first col is a time
          if (!time || (!time.includes('AM') && !time.includes('PM'))) continue;

          // For each mapped column, get the value
          Object.keys(dateMap).forEach(k => {
            let colIdx = parseInt(k);
            if (cols.length > colIdx) {
              let val = cols[colIdx];
              // Remove empty strings or non-numbers
              if (val && !isNaN(parseFloat(val))) {
                 cleanIDP.push([dateMap[k], time, val]);
              }
            }
          });
        }
      }
      
      // Save to DB_IDP
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

// --- UTILITIES ---

function parseDate(s) {
  // Handle "2/8/26"
  if(!s) return "";
  let d = new Date(s);
  if(isNaN(d.getTime())) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function parseDateLong(s) {
  // Handle "Friday, February 6, 2026"
  // Remove "Open" or "Occupied" if regex grabbed too much
  s = s.replace(/Open|Occupied|Seats|\+\/\-/g, '').trim();
  let d = new Date(s);
  if (isNaN(d.getTime())) return null;
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
