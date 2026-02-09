var ImportHandler = {
  processPaste: function(schedRaw, idpRaw) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ['DB_SCHEDULE', 'DB_IDP', 'DB_FURLOUGH'].forEach(n => {
       if(!ss.getSheetByName(n)) {
         let s = ss.insertSheet(n);
         if(n==='DB_FURLOUGH') s.appendRow(['ID', 'Agent Name', 'Date', 'Start Time', 'Type', 'WeekRotation', 'End Time']);
       }
    });

    // 1. SCHEDULE
    if (schedRaw && schedRaw.trim().length > 0) {
      let cleanSched = [];
      const lines = schedRaw.split(/\r?\n/).filter(l => l.trim().length > 0);
      let currentAgent = "", currentDate = "";
      // Regex handling AM/PM and no-AM/PM cases if needed
      const segmentRegex = /([a-zA-ZÀ-ÿ0-9\/\(\)\s\-\.]+)[\t\s]+(\d{1,2}:\d{2}\s?[AP]M?)[\t\s]+(\d{1,2}:\d{2}\s?[AP]M?)\s*$/i;
      const dateRegex = /^[\t\s]*(\d{1,2}\/\d{1,2}\/\d{2,4})/;

      lines.forEach(line => {
        const text = line.trim();
        if (text.includes('Agent:')) {
          const parts = text.split(':');
          if (parts.length > 1) currentAgent = parts[1].replace(/^\s*\d+\s+/, '').trim();
          return;
        }
        const dateMatch = line.match(dateRegex);
        if (dateMatch) currentDate = parseDate(dateMatch[1]);
        
        if (currentAgent && currentDate) {
          const segMatch = line.match(segmentRegex);
          if (segMatch) {
            let activity = segMatch[1].trim();
            if (!activity.toLowerCase().includes('scheduled activity')) {
               cleanSched.push([currentAgent, currentDate, activity, segMatch[2].trim(), segMatch[3].trim()]);
            }
          }
        }
      });
      
      if (cleanSched.length > 0) {
        const db = ss.getSheetByName('DB_SCHEDULE');
        db.clear();
        db.appendRow(['Agent Name', 'Date', 'Activity', 'Start Time', 'End Time']);
        db.getRange(2,1,cleanSched.length,5).setValues(cleanSched);
      }
    }

    // 2. IDP (Original Format)
    if (idpRaw && idpRaw.trim().length > 0) {
      let cleanIDP = []; 
      const lines = idpRaw.split(/\r?\n/).filter(l => l.trim().length > 0);
      let hIdx = lines.findIndex(l => l.toLowerCase().includes('time') && l.toLowerCase().includes('requirements'));
      
      if (hIdx > -1) {
        let headerLine = lines[hIdx];
        let headers = headerLine.split(/[\t,]| {2,}/);
        let colMap = {}; 
        let currentDay = "";
        
        headers.forEach((h, i) => {
          let lower = h.trim().toLowerCase();
          // Detect Day names in header or previous header row logic if applicable
          // Assuming structure is Time, Sunday, Req, Open... or similar.
          // Based on your file: Row 1 has days, Row 2 has Req/Open. 
          // The previous parser logic was better suited for that specific CSV structure.
          // Let's use column index mapping assuming standard 3-col blocks per day.
          
          if(lower.includes('req')) colMap[i] = 'req';
          else if(lower.includes('open') && !lower.includes('dif')) colMap[i] = 'open';
        });
        
        // We need to map which set of Req/Open belongs to which day.
        // Simplified approach: Order of days is Sun, Mon, Tue...
        // We will assume the file is structured chronologically left to right.
      }
      
      // Re-using the ROBUST parser from V6 which worked for your file
      // (Included in previous turn, ensuring it is used here)
      // ... [Insert V6 Import Logic Here] ... 
      // For brevity, I will stick to the V6 logic which you confirmed worked for import.
      return ImportHandlerV6(schedRaw, idpRaw); 
    }
    return "Import Successful";
  }
};

// ... V6 Import Logic Wrapper ...
function ImportHandlerV6(schedRaw, idpRaw) {
    // (Paste the logic from V6 ImportHandler here - it was correct for your files)
    // I will include the full V6 logic in the final block below to ensure copy-paste works.
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ['DB_SCHEDULE', 'DB_IDP', 'DB_FURLOUGH'].forEach(n => {
       if(!ss.getSheetByName(n)) {
         let s = ss.insertSheet(n);
         if(n==='DB_FURLOUGH') s.appendRow(['ID', 'Agent Name', 'Date', 'Start Time', 'Type', 'WeekRotation', 'End Time']);
       }
    });
    // Schedule
    if (schedRaw && schedRaw.trim().length > 0) {
      let cleanSched = [];
      const lines = schedRaw.split(/\r?\n/).filter(l => l.trim().length > 0);
      let currentAgent = "", currentDate = "";
      const segmentRegex = /([a-zA-ZÀ-ÿ0-9\/\(\)\s\-\.]+)[\t\s]+(\d{1,2}:\d{2}\s?[AP]M?)[\t\s]+(\d{1,2}:\d{2}\s?[AP]M?)\s*$/i;
      const dateRegex = /^[\t\s]*(\d{1,2}\/\d{1,2}\/\d{2,4})/;
      lines.forEach(line => {
        const text = line.trim();
        if (text.includes('Agent:')) {
          let parts = text.split(':'); if(parts.length>1) currentAgent = parts[1].replace(/^\s*\d+\s+/, '').trim(); return;
        }
        let dMatch = line.match(dateRegex); if(dMatch) currentDate = parseDate(dMatch[1]);
        if(currentAgent && currentDate) {
          let sMatch = line.match(segmentRegex);
          if(sMatch) {
             let act = sMatch[1].trim();
             if(!act.toLowerCase().includes('scheduled')) cleanSched.push([currentAgent, currentDate, act, sMatch[2].trim(), sMatch[3].trim()]);
          }
        }
      });
      if(cleanSched.length>0){const db=ss.getSheetByName('DB_SCHEDULE');db.clear();db.appendRow(['Agent Name','Date','Activity','Start Time','End Time']);db.getRange(2,1,cleanSched.length,5).setValues(cleanSched);}
    }
    // IDP
    if (idpRaw && idpRaw.trim().length > 0) {
      let cleanIDP=[]; const lines=idpRaw.split(/\r?\n/).filter(l=>l.trim().length>0);
      let dayRowIdx=-1;
      for(let i=0;i<lines.length;i++){if(lines[i].toLowerCase().includes('sunday')||lines[i].toLowerCase().includes('monday')){dayRowIdx=i;break;}}
      if(dayRowIdx>-1 && dayRowIdx+1<lines.length){
        const dayRow=lines[dayRowIdx].split(/[\t,]/); const metricRow=lines[dayRowIdx+1].split(/[\t,]/);
        let colMap={}; let currentDayStr="";
        for(let c=0;c<dayRow.length;c++){
          let dVal=dayRow[c].trim(); if(dVal) currentDayStr=dVal;
          let m=metricRow[c]?metricRow[c].trim().toLowerCase():"";
          if(m.includes('req')) colMap[c]={day:currentDayStr,type:'req'};
          else if(m.includes('open')) colMap[c]={day:currentDayStr,type:'open'};
        }
        for(let i=dayRowIdx+2;i<lines.length;i++){
          let cols=lines[i].split(/[\t,]/);
          let t=cols.find(c=>c.includes(':')&&(c.includes('00')||c.includes('15')||c.includes('30')||c.includes('45')));
          if(!t)continue;
          // Normalize time
          let d=new Date(`2000/01/01 ${t}`); let tNorm=isNaN(d.getTime())?t:Utilities.formatDate(d,Session.getScriptTimeZone(),'HH:mm');
          let rowData={};
          Object.keys(colMap).forEach(idx=>{
            if(cols[idx]){
              let info=colMap[idx]; if(!rowData[info.day]) rowData[info.day]={req:0,open:0};
              let val=parseFloat(cols[idx]); if(isNaN(val)) val=0;
              if(info.type==='req') rowData[info.day].req=val; if(info.type==='open') rowData[info.day].open=val;
            }
          });
          Object.keys(rowData).forEach(day=>cleanIDP.push([day,tNorm,rowData[day].req,rowData[day].open]));
        }
      }
      if(cleanIDP.length>0){const db=ss.getSheetByName('DB_IDP');db.clear();db.appendRow(['Day','Interval','Required','Open']);db.getRange(2,1,cleanIDP.length,4).setValues(cleanIDP);}
    }
    return "Import Successful";
}

function parseDate(s) { if(!s)return ""; let d=new Date(s); return isNaN(d.getTime())?"":Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd'); }
function parseDateLong(s) { let d=new Date(s.replace(/Requirements|Open|Occupied|Seats|\+\/\-/gi,'').trim()); return isNaN(d.getTime())?null:Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd'); }
