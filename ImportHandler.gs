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

    // --- 1. PROCESS SCHEDULE (Context & Details) ---
    if (schedRaw && schedRaw.trim().length > 0) {
      let cleanSched = [];
      const lines = schedRaw.split(/\r?\n/).filter(l => l.trim().length > 0);
      
      let currentAgent = "";
      let currentDate = "";
      
      // Regex: "Activity 6:30 AM 8:00 AM"
      const segmentRegex = /([a-zA-ZÀ-ÿ0-9\/\(\)\s\-\.]+)[\t\s]+(\d{1,2}:\d{2}\s?[AP]M)[\t\s]+(\d{1,2}:\d{2}\s?[AP]M)\s*$/i;
      const dateRegex = /^[\t\s]*(\d{1,2}\/\d{1,2}\/\d{2,4})/;

      lines.forEach(line => {
        const text = line.trim();
        
        // Agent Header
        if (text.includes('Agent:')) {
          const parts = text.split(':');
          if (parts.length > 1) {
            currentAgent = parts[1].replace(/^\s*\d+\s+/, '').trim();
          }
          return;
        }
        
        // Date Header
        const dateMatch = line.match(dateRegex);
        if (dateMatch) {
          currentDate = parseDate(dateMatch[1]);
        }
        
        // Activity Rows
        if (currentAgent && currentDate) {
          const segMatch = line.match(segmentRegex);
          if (segMatch) {
            let activity = segMatch[1].trim();
            let start = segMatch[2].trim();
            let end = segMatch[3].trim();
            
            // Filter garbage headers
            if (!activity.toLowerCase().includes('scheduled activity') && 
                !start.toLowerCase().includes('start')) {
               cleanSched.push([currentAgent, currentDate, activity, start, end]);
            }
          }
        }
      });
      
      // Save Schedule
      if (cleanSched.length > 0) {
        const db = ss.getSheetByName('DB_SCHEDULE');
        db.clear();
        db.appendRow(['Agent Name', 'Date', 'Activity', 'Start Time', 'End Time']);
        db.getRange(2,1,cleanSched.length,5).setValues(cleanSched);
      }
    }

    // --- 2. PROCESS IDP (Demand AND Supply) ---
    if (idpRaw && idpRaw.trim().length > 0) {
      let cleanIDP = []; // [Date, Time, Req, Open]
      
      const lines = idpRaw.split(/\r?\n/).filter(l => l.trim().length > 0);
      
      // Find header row (Requirements... Open...)
      let hIdx = lines.findIndex(l => l.toLowerCase().includes('time') && l.toLowerCase().includes('requirements'));
      
      if (hIdx > -1) {
        let headerLine = lines[hIdx];
        // Split by tabs or multiple spaces
        let headers = headerLine.split(/\t+/);
        if(headers.length < 2) headers = headerLine.split(/\s{2,}/);
        
        // Map columns: index -> {date: "YYYY-MM-DD", type: "req"|"open"}
        let colMap = {}; 
        
        headers.forEach((h, i) => {
          let lower = h.toLowerCase();
          
          // Parse "Requirements Friday, Feb..."
          if (lower.includes('requirements')) {
             let dStr = h.replace(/requirements/i, '').trim();
             let pDate = parseDateLong(dStr);
             if(pDate) colMap[i] = { date: pDate, type: 'req' };
          }
          // Parse "Open Friday, Feb..." (Exclude "Open +/-")
          else if (lower.includes('open') && !lower.includes('+') && !lower.includes('-') && !lower.includes('difference')) {
             let dStr = h.replace(/open/i, '').trim();
             let pDate = parseDateLong(dStr);
             if(pDate) colMap[i] = { date: pDate, type: 'open' };
          }
        });

        // Parse Data Rows
        // We need to aggregate by Date+Time because Req and Open might be in different columns
        let tempMap = {}; // "Date|Time" -> {req:0, open:0}

        for (let i = hIdx + 1; i < lines.length; i++) {
          let line = lines[i];
          let cols = line.split(/\t+/);
          if(cols.length < 2) cols = line.split(/\s{2,}/); 
          
          let time = cols[0];
          if (!time || (!time.includes('AM') && !time.includes('PM'))) continue;

          Object.keys(colMap).forEach(k => {
            let colIdx = parseInt(k);
            if (cols.length > colIdx) {
              let val = parseFloat(cols[colIdx]);
              if (isNaN(val)) val = 0;
              
              let info = colMap[colIdx];
              let key = info.date + "|" + time;
              
              if (!tempMap[key]) tempMap[key] = { date: info.date, time: time, req: 0, open: 0 };
              
              if (info.type === 'req') tempMap[key].req = val;
              if (info.type === 'open') tempMap[key].open = val;
            }
          });
        }
        
        // Convert map to array
        Object.values(tempMap).forEach(o => {
          cleanIDP.push([o.date, o.time, o.req, o.open]);
        });
      }
      
      // Save IDP
      if (cleanIDP.length > 0) {
        const db = ss.getSheetByName('DB_IDP');
        db.clear();
        db.appendRow(['Date', 'Interval', 'Required', 'Open']); // Added 'Open' column
        db.getRange(2,1,cleanIDP.length,4).setValues(cleanIDP);
      }
    }

    return "Import Successful";
  }
};

// Utils
function parseDate(s) {
  if(!s) return "";
  let d = new Date(s);
  return isNaN(d.getTime()) ? "" : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function parseDateLong(s) {
  // Clean string from extra words if regex matched loosely
  s = s.replace(/Requirements|Open|Occupied|Seats|\+\/\-/gi, '').trim();
  let d = new Date(s);
  return isNaN(d.getTime()) ? null : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
