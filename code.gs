const SHEET_ID = '12OWxEDQot_E0YQsIfhIJmEw26a5ujTrV2etoXYntt_0';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Timetable Sync')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSavedSettings() {
  let props = PropertiesService.getUserProperties();
  return {
    sem: props.getProperty('SYNC_SEM') || '',
    calendarId: props.getProperty('SYNC_CALENDAR_ID') || '',
    section: props.getProperty('SYNC_SECTION') || 'A',
    batch: props.getProperty('SYNC_BATCH') || 'Both',
    dates: JSON.parse(props.getProperty('SYNC_DATES') || '{}'),
    courseConfig: JSON.parse(props.getProperty('SYNC_COURSE_CONFIG') || '{}'),
    selectedCourses: JSON.parse(props.getProperty('SYNC_SELECTED_COURSES') || '[]')
  };
}

function getSheetNames() {
  return SpreadsheetApp.openById(SHEET_ID).getSheets().map(s => s.getName());
}

function getCalendarList() {
  let list = [];
  try {
    let response = Calendar.CalendarList.list({ minAccessRole: 'writer' });
    let items = response.items || [];
    for (let i = 0; i < items.length; i++) {
      let c = items[i];
      list.push({ id: c.id, name: c.summary + (c.primary ? ' (Primary)' : '') });
    }
  } catch (err) {
    let primary = CalendarApp.getDefaultCalendar();
    if (primary) list.push({ id: primary.getId(), name: primary.getName() + ' (Primary)' });
  }
  return list;
}

function getCourses(semName, section, batch) {
  let ss = SpreadsheetApp.openById(SHEET_ID);
  let slotSheet = ss.getSheetByName("Slots");
  let semSheet = ss.getSheetByName(semName);

  if (!slotSheet || !semSheet) return {};

  let slotMapping = readSlots(slotSheet);
  let parsed = parseTimetable(semSheet, slotMapping, section, batch);
  return parsed.courses;
}

function saveSettingsAndSync(payload) {
  let props = PropertiesService.getUserProperties();
  props.setProperty('SYNC_SEM', payload.semName);
  props.setProperty('SYNC_CALENDAR_ID', payload.calendarId);
  props.setProperty('SYNC_SECTION', payload.section);
  props.setProperty('SYNC_BATCH', payload.batch);
  props.setProperty('SYNC_DATES', JSON.stringify(payload.dates));
  props.setProperty('SYNC_COURSE_CONFIG', JSON.stringify(payload.courseConfig));
  props.setProperty('SYNC_SELECTED_COURSES', JSON.stringify(payload.selectedCourses));

  let triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) ScriptApp.deleteTrigger(triggers[i]);

  ScriptApp.newTrigger('syncCalendarEvent').timeBased().everyDays(1).atHour(7).create();
  syncCalendarEvent();
}

function syncCalendarEvent() {
  let props = PropertiesService.getUserProperties();
  let semName = props.getProperty('SYNC_SEM');
  let calendarId = props.getProperty('SYNC_CALENDAR_ID');
  if (!semName || !calendarId) return;
  executeSync(semName, calendarId);
}

function executeSync(semName, calendarId) {
  let ss = SpreadsheetApp.openById(SHEET_ID);
  let slotSheet = ss.getSheetByName("Slots");
  let semSheet = ss.getSheetByName(semName);
  if (!slotSheet || !semSheet) return;
  
  let props = PropertiesService.getUserProperties();
  let section = props.getProperty('SYNC_SECTION') || 'A';
  let batch = props.getProperty('SYNC_BATCH') || 'Both';
  let dates = JSON.parse(props.getProperty('SYNC_DATES') || '{}');
  let courseConfig = JSON.parse(props.getProperty('SYNC_COURSE_CONFIG') || '{}');
  let selectedCourses = JSON.parse(props.getProperty('SYNC_SELECTED_COURSES') || '[]');
  
  let slotMapping = readSlots(slotSheet);
  let parsed = parseTimetable(semSheet, slotMapping, section, batch);
  
  let desiredEvents = parsed.events.filter(ev => selectedCourses.includes(ev.rawTitle));
  
  let existingEventsMap = {};
  let pageToken = null;
  do {
    let response = Calendar.Events.list(calendarId, { privateExtendedProperty: 'app=iiitdwd_sync', maxResults: 2500, pageToken: pageToken });
    let items = response.items || [];
    for (let i = 0; i < items.length; i++) {
      if (items[i].status !== "cancelled" && items[i].extendedProperties?.private?.syncId) {
        existingEventsMap[items[i].extendedProperties.private.syncId] = items[i];
      }
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  
  let processedSyncIds = {};
  
  // Calculate precise boundaries (x-1, y+1, end-1)
  let preMidEnd = dates.midStart ? new Date(dates.midStart) : new Date(dates.endStart);
  if (dates.midStart) preMidEnd.setDate(preMidEnd.getDate() - 1);
  
  let postMidStart = dates.midEnd ? new Date(dates.midEnd) : new Date(dates.semStart);
  if (dates.midEnd) postMidStart.setDate(postMidStart.getDate() + 1);
  
  let finalEnd = dates.endStart ? new Date(dates.endStart) : new Date();
  if (dates.endStart) finalEnd.setDate(finalEnd.getDate() - 1);
  
  for (let i = 0; i < desiredEvents.length; i++) {
    let evData = desiredEvents[i];
    let durationType = courseConfig[evData.rawTitle] || 'full';
    
    // Split scheduling into separate periods to avoid exam overlap
    let periods = [];
    if (durationType === 'pre') {
      periods.push({ start: dates.semStart, end: preMidEnd, suffix: '_pre' });
    } else if (durationType === 'post') {
      periods.push({ start: postMidStart, end: finalEnd, suffix: '_post' });
    } else {
      // Full sem: Split into two completely separate recurring events
      periods.push({ start: dates.semStart, end: preMidEnd, suffix: '_pre' });
      periods.push({ start: postMidStart, end: finalEnd, suffix: '_post' });
    }
    
    for (let p of periods) {
      if (!p.start || !p.end) continue;
      
      let specificSyncId = evData.syncId + p.suffix;
      processedSyncIds[specificSyncId] = true;
      
      let startEnd = getFirstOccurrence(p.start, evData.day, evData.timeStr, evData.slot);
      if (!startEnd) continue;
      
      let untilDate = new Date(p.end); 
      untilDate.setHours(23, 59, 59);
      
      // Safety: Skip event generation if the first class starts after the period ends
      if (startEnd.start > untilDate) continue;
      
      let untilStr = untilDate.toISOString().replace(/[-:]/g, '').split('.')[0] + 'Z';
      
      let eventResource = {
        summary: evData.title,
        location: evData.location,
        start: { dateTime: startEnd.start.toISOString(), timeZone: Session.getScriptTimeZone() },
        end: { dateTime: startEnd.end.toISOString(), timeZone: Session.getScriptTimeZone() },
        recurrence: ["RRULE:FREQ=WEEKLY;UNTIL=" + untilStr],
        extendedProperties: { private: { app: 'iiitdwd_sync', syncId: specificSyncId } }
      };
      
      let existing = existingEventsMap[specificSyncId];
      if (existing) {
        let exStart = existing.start.dateTime || existing.start.date;
        let d1 = new Date(exStart); let d2 = new Date(eventResource.start.dateTime);
        let exEnd = existing.end.dateTime || existing.end.date;
        let e1 = new Date(exEnd); let e2 = new Date(eventResource.end.dateTime);
        
        if (existing.summary !== eventResource.summary || existing.location !== eventResource.location ||
            d1.getHours() !== d2.getHours() || d1.getMinutes() !== d2.getMinutes() || 
            e1.getHours() !== e2.getHours() || e1.getMinutes() !== e2.getMinutes() || 
            !existing.recurrence || !existing.recurrence[0].includes(untilStr)) {
          Calendar.Events.update(eventResource, calendarId, existing.id);
        }
      } else {
        Calendar.Events.insert(eventResource, calendarId);
      }
    }
  }
  
  // This loop automatically deletes the old, flawed full-sem events
  let existingKeys = Object.keys(existingEventsMap);
  for (let i = 0; i < existingKeys.length; i++) {
    if (!processedSyncIds[existingKeys[i]]) {
      try { Calendar.Events.remove(calendarId, existingEventsMap[existingKeys[i]].id); } catch(e) {}
    }
  }
}

function readSlots(sheet) {
  let data = sheet.getDataRange().getValues();
  let slotMapping = {};
  let timeRowIdx = -1;
  for (let i = 0; i < Math.min(20, data.length); i++) {
    if (data[i].join("").toLowerCase().includes("time")) { timeRowIdx = i; break; }
  }
  let timeRow = data[timeRowIdx];
  let currentDay = "";
  let validDays = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"];

  for (let r = timeRowIdx + 1; r < data.length; r++) {
    let col0 = String(data[r][0]).trim().toUpperCase();
    let col1 = String(data[r][1]).trim().toUpperCase();
    if (validDays.includes(col0)) currentDay = col0;
    else if (validDays.includes(col1)) currentDay = col1;
    if (!currentDay) continue;

    for (let c = 2; c < data[r].length; c++) {
      let cellVal = String(data[r][c]).trim();
      if (cellVal && cellVal !== "NaN") {
        let timeStr = String(timeRow[c]).trim();
        if (timeStr && timeStr.includes("-")) {
          if (!slotMapping[cellVal]) slotMapping[cellVal] = [];
          slotMapping[cellVal].push({ day: currentDay, time: timeStr });
        }
      }
    }
  }
  return slotMapping;
}

function parseTimetable(sheet, slotMapping, section, batch) {
  let data = sheet.getDataRange().getValues();
  let headerRowIdx = -1, titleColIdx = -1, sectionColIdx = -1;

  for (let i = 0; i < Math.min(30, data.length); i++) {
    for (let j = 0; j < data[i].length; j++) {
      let val = String(data[i][j]).toLowerCase();
      if (val.includes("course title") || val.includes("course name")) { headerRowIdx = i; titleColIdx = j; }
      if (val.includes("section:") || val.includes("slot, classroom") || val === "section") { sectionColIdx = j; }
    }
    if (headerRowIdx !== -1 && sectionColIdx !== -1) break;
  }

  let desiredEvents = [];
  let coursesInfo = {};

  if (headerRowIdx === -1 || sectionColIdx === -1) return { events: desiredEvents, courses: coursesInfo };

  for (let r = headerRowIdx + 1; r < data.length; r++) {
    let courseTitle = String(data[r][titleColIdx]).trim();
    let sectionData = String(data[r][sectionColIdx]).trim();
    if (!courseTitle || courseTitle === "NaN" || !sectionData || sectionData === "NaN") continue;

    if (sectionData.includes('{')) {
      let regex = /(?:([^:{]+):)?\s*\{([^}]+)\}/g;
      let match;
      let matchedSection = false;

      while ((match = regex.exec(sectionData)) !== null) {
        let prefix = match[1] ? match[1].trim().toUpperCase() : "";
        if (prefix && section && !prefix.includes(section.toUpperCase())) continue;

        matchedSection = true;
        let pairs = match[2].trim().split(';');

        for (let pair of pairs) {
          let parts = pair.replace(/,/g, ' ').trim().replace(/\s+/g, ' ').split(' ');
          if (parts.length > 0) {
            let slot = parts[0];
            let room = "TBA";

            if (parts.length >= 3) {
              if (batch === '1') room = parts[1];
              else if (batch === '2') room = parts[2];
              else room = parts[1] + " / " + parts[2];
            } else if (parts.length === 2) {
              room = parts[1];
            }

            if (slotMapping[slot]) {
              for (let t of slotMapping[slot]) {
                desiredEvents.push({
                  syncId: (courseTitle + "_" + slot + "_" + t.day + "_" + t.time).replace(/[^a-zA-Z0-9]/g, ""),
                  title: courseTitle,
                  rawTitle: courseTitle,
                  location: room,
                  day: t.day,
                  timeStr: t.time,
                  slot: slot
                });
              }
            }
          }
        }
      }
      if (matchedSection && !coursesInfo[courseTitle]) coursesInfo[courseTitle] = { isCore: true };
    }
    else {
      let tokens = sectionData.replace(/,/g, ' ').trim().replace(/\s+/g, ' ').split(' ');
      let foundSlot = false;
      let room = "TBA";
      let slot = null;

      for (let i = 0; i < tokens.length; i++) {
        let cleanToken = tokens[i].trim().toUpperCase();
        if (slotMapping[cleanToken]) {
          slot = cleanToken;
          foundSlot = true;
          if (i + 1 < tokens.length && !slotMapping[tokens[i + 1].toUpperCase()]) room = tokens[i + 1];
          break;
        }
      }

      if (foundSlot) {
        if (!coursesInfo[courseTitle]) coursesInfo[courseTitle] = { isCore: false };
        for (let t of slotMapping[slot]) {
          desiredEvents.push({
            syncId: (courseTitle + "_" + slot + "_" + t.day + "_" + t.time).replace(/[^a-zA-Z0-9]/g, ""),
            title: courseTitle,
            rawTitle: courseTitle,
            location: room,
            day: t.day,
            timeStr: t.time,
            slot: slot
          });
        }
      }
    }
  }
  return { events: desiredEvents, courses: coursesInfo };
}

function getFirstOccurrence(boundaryDateStr, dayStr, timeStr, slotName) {
  const days = { "SUN": 0, "MON": 1, "TUE": 2, "WED": 3, "THU": 4, "FRI": 5, "SAT": 6 };
  let targetDay = days[dayStr.toUpperCase()];
  if (targetDay === undefined) return null;

  let startD = new Date(boundaryDateStr); startD.setHours(0, 0, 0, 0);
  let daysToAdd = (targetDay - startD.getDay() + 7) % 7;
  startD.setDate(startD.getDate() + daysToAdd);

  // Extract only the START time from the sheet's column header
  let timeParts = timeStr.split('-');
  let st = { h: parseInt(timeParts[0].split(':')[0], 10), m: parseInt(timeParts[0].split(':')[1], 10) || 0 };

  let startObj = new Date(startD);
  startObj.setHours(st.h, st.m, 0, 0);

  // Hardcoded durations based on slot naming convention
  let durationMins = 90; // Default: Classes = 1.5 hrs
  if (slotName) {
    let s = slotName.toUpperCase();
    if (s.endsWith('-T')) {
      durationMins = 60; // Tutorial = 1 hr
    } else if (s.startsWith('L') && s.length > 1 && !isNaN(parseInt(s.charAt(1)))) {
      durationMins = 120; // Lab = 2 hrs
    }
  }

  // Dynamically calculate the precise end time
  let endObj = new Date(startObj.getTime() + durationMins * 60000);

  return { start: startObj, end: endObj };
}
