/************** CONFIGURATION **************/

// Sheet names
const MAIN_SHEET_NAME = "Main Sheet";
const SUBJECTS_SHEET_NAME = "Subjects";
const FILTERED_SHEET_NAME = "My Schedule";   // new sheet showing only your subjects

// Calendar configuration
// Leave as null to use your primary Google Calendar.
const CALENDAR_ID = null;

// Column names (must match header text exactly)
const COL_TRACK = "Track";
const COL_COURSE_CODE = "Course Code";
const COL_COURSE_NAME = "Course Name";
const COL_DATE = "Date";
const COL_TIME_RANGE = "Time";       // e.g. "07.00 PM - 09.45 PM"
const COL_WEEK_DAY = "Week Day";     // optional
const COL_REMARKS = "Remarks";       // optional

// New column that will be created on Main Sheet to store event IDs
const COL_EVENT_ID = "Calendar Event ID";

/************** CUSTOM MENU **************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Schedule Automation")
    .addItem("1) Refresh My Schedule", "refreshFilteredList_")
    .addItem("2) Sync Calendar Events", "syncCalendar_")
    .addItem("Run All (1 → 2)", "runAll_")
    .addToUi();
}

function runAll_() {
  refreshFilteredList_();
  syncCalendar_();
}

/************** STEP 1 – BUILD 'MY SCHEDULE' SHEET **************/

function refreshFilteredList_() {
  const ss = SpreadsheetApp.getActive();
  const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  const subjectsSheet = ss.getSheetByName(SUBJECTS_SHEET_NAME);

  if (!mainSheet || !subjectsSheet) {
    throw new Error("Check that sheets named 'Main Sheet' and 'Subjects' exist (or update names in the script).");
  }

  // Make sure Main Sheet has the Calendar Event ID column
  ensureEventIdColumn_(mainSheet);

  // Build the set of (Track|CourseCode) pairs from Subjects
  const subjectKeys = getSubjectKeys_(subjectsSheet);

  // Read Main Sheet headers and data
  const lastRow = mainSheet.getLastRow();
  const lastCol = mainSheet.getLastColumn();
  if (lastRow < 2) {
    throw new Error("Main Sheet has no data rows.");
  }
  const headerRow = mainSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const mainValues = mainSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const headerIndex = indexHeaders_(headerRow);

  const trackCol = headerIndex[COL_TRACK];
  const courseCodeCol = headerIndex[COL_COURSE_CODE];

  if (trackCol == null || courseCodeCol == null) {
    throw new Error("Could not find 'Track' or 'Course Code' columns on Main Sheet.");
  }

  // Build data for My Schedule
  const filteredData = [];
  filteredData.push(headerRow); // headers first

  mainValues.forEach(row => {
    const track = (row[trackCol] || "").toString().trim();
    const courseCode = (row[courseCodeCol] || "").toString().trim();
    if (!track && !courseCode) return;

    const key = buildKey_(track, courseCode);
    if (subjectKeys.has(key)) {
      filteredData.push(row);
    }
  });

  // Create / clear My Schedule sheet
  let filteredSheet = ss.getSheetByName(FILTERED_SHEET_NAME);
  if (!filteredSheet) {
    filteredSheet = ss.insertSheet(FILTERED_SHEET_NAME);
  }
  filteredSheet.clearContents();

  if (filteredData.length > 0) {
    filteredSheet.getRange(1, 1, filteredData.length, headerRow.length).setValues(filteredData);
    filteredSheet.autoResizeColumns(1, headerRow.length);
  }
}

/************** STEP 2 & 3 – CREATE CALENDAR EVENTS (NO DUPLICATES) **************/

function syncCalendar_() {
  const ss = SpreadsheetApp.getActive();
  const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  const subjectsSheet = ss.getSheetByName(SUBJECTS_SHEET_NAME);

  if (!mainSheet || !subjectsSheet) {
    throw new Error("Main Sheet or Subjects sheet not found.");
  }

  // Ensure event ID column
  ensureEventIdColumn_(mainSheet);

  const subjectKeys = getSubjectKeys_(subjectsSheet);

  const lastRow = mainSheet.getLastRow();
  const lastCol = mainSheet.getLastColumn();
  if (lastRow < 2) return; // nothing to do

  const headerRow = mainSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const headerIndex = indexHeaders_(headerRow);

  const trackCol = headerIndex[COL_TRACK];
  const courseCodeCol = headerIndex[COL_COURSE_CODE];
  const courseNameCol = headerIndex[COL_COURSE_NAME];
  const dateCol = headerIndex[COL_DATE];
  const timeCol = headerIndex[COL_TIME_RANGE];
  const weekDayCol = headerIndex[COL_WEEK_DAY];
  const remarksCol = headerIndex[COL_REMARKS];
  const eventIdCol = headerIndex[COL_EVENT_ID];

  if ([trackCol, courseCodeCol, courseNameCol, dateCol, timeCol, eventIdCol].some(v => v == null)) {
    throw new Error("Check that Main Sheet has columns: Track, Course Code, Course Name, Date, Time, and Calendar Event ID.");
  }

  const dataRange = mainSheet.getRange(2, 1, lastRow - 1, lastCol);
  const values = dataRange.getValues();

  const calendar = CALENDAR_ID
    ? CalendarApp.getCalendarById(CALENDAR_ID)
    : CalendarApp.getDefaultCalendar();

  if (!calendar) {
    throw new Error("Could not access calendar. Check CALENDAR_ID or permissions.");
  }

  values.forEach((row, idx) => {
    // Skip completely empty rows
    if (row.every(cell => cell === "")) return;

    const track = (row[trackCol] || "").toString().trim();
    const courseCode = (row[courseCodeCol] || "").toString().trim();
    if (!track && !courseCode) return;

    const key = buildKey_(track, courseCode);
    if (!subjectKeys.has(key)) return; // not one of your chosen subjects

    // Skip if event already created
    const existingEventId = (row[eventIdCol] || "").toString().trim();
    if (existingEventId) return;

    const title = (row[courseNameCol] || "").toString().trim();
    const dateValue = row[dateCol];
    const timeRange = (row[timeCol] || "").toString().trim();
    if (!title || !dateValue || !timeRange) return;

    let startEnd;
    try {
      startEnd = combineDateAndTimeFromRange_(dateValue, timeRange);
    } catch (e) {
      // If time parsing fails, just skip this row instead of crashing everything
      Logger.log("Skipping row " + (idx + 2) + " due to time parse error: " + e);
      return;
    }

    const [startDateTime, endDateTime] = startEnd;

    const weekDay = weekDayCol != null ? row[weekDayCol] : "";
    const remarks = remarksCol != null ? row[remarksCol] : "";

    const descriptionLines = [];
    descriptionLines.push("Course Code: " + courseCode);
    descriptionLines.push("Track: " + track);
    if (weekDay) descriptionLines.push("Day: " + weekDay);
    if (remarks) descriptionLines.push("Remarks: " + remarks);
    const description = descriptionLines.join("\n");

    const event = calendar.createEvent(title, startDateTime, endDateTime, {
      description: description || undefined
    });

    // Save event ID back to Main Sheet row
    row[eventIdCol] = event.getId();
  });

  // Write updated values (with event IDs) back to the sheet
  dataRange.setValues(values);
}

/************** HELPERS **************/

function ensureEventIdColumn_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    // Empty sheet (shouldn't happen for your case)
    sheet.getRange(1, 1).setValue(COL_EVENT_ID);
    return;
  }
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  if (!headers.includes(COL_EVENT_ID)) {
    sheet.getRange(1, lastCol + 1).setValue(COL_EVENT_ID);
  }
}

function indexHeaders_(headerRow) {
  const map = {};
  headerRow.forEach((name, idx) => {
    const key = name.toString().trim();
    if (key) map[key] = idx;
  });
  return map;
}

function getSubjectKeys_(subjectsSheet) {
  const lastRow = subjectsSheet.getLastRow();
  const lastCol = subjectsSheet.getLastColumn();
  const keys = new Set();
  if (lastRow < 2) return keys;

  const headerRow = subjectsSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const headerIndex = indexHeaders_(headerRow);

  const trackCol = headerIndex[COL_TRACK];
  const courseCodeCol = headerIndex[COL_COURSE_CODE];

  if (trackCol == null || courseCodeCol == null) {
    throw new Error("Subjects sheet must have 'Course Code' and 'Track' columns.");
  }

  const values = subjectsSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  values.forEach(row => {
    const track = (row[trackCol] || "").toString().trim();
    const courseCode = (row[courseCodeCol] || "").toString().trim();
    if (track && courseCode) {
      keys.add(buildKey_(track, courseCode));
    }
  });

  return keys;
}

function buildKey_(track, courseCode) {
  return `${track}||${courseCode}`.toUpperCase();
}

/**
 * Given a 'Date' value and a 'Time' value like "07.00 PM - 09.45 PM",
 * returns [startDateTime, endDateTime] as JavaScript Date objects.
 */
function combineDateAndTimeFromRange_(dateValue, timeRangeStr) {
  const date = parseDateValue_(dateValue);
  const parts = timeRangeStr.split("-");
  if (parts.length !== 2) {
    throw new Error("Time range should look like '07.00 PM - 09.45 PM'. Got: " + timeRangeStr);
  }

  const startTimeInfo = parseTimeString_(parts[0]);
  const endTimeInfo = parseTimeString_(parts[1]);

  const start = new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
    startTimeInfo.h,
    startTimeInfo.m,
    0
  );
  const end = new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
    endTimeInfo.h,
    endTimeInfo.m,
    0
  );

  return [start, end];
}

function parseDateValue_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) {
    return value;
  }
  const str = value.toString().trim();
  // Expecting dd-mm-yyyy (e.g. 12-12-2025)
  const m = str.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (!m) {
    throw new Error("Cannot parse date: " + str);
  }
  const day = Number(m[1]);
  const month = Number(m[2]) - 1; // 0-based
  const year = Number(m[3]);
  return new Date(year, month, day);
}

function parseTimeString_(timeStr) {
  // Convert "07.00 PM" -> "07:00 PM"
  const clean = timeStr.replace(/\./g, ":").trim().toUpperCase();
  const m = clean.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/);
  if (!m) {
    throw new Error("Cannot parse time: " + timeStr);
  }
  let hour = Number(m[1]);
  const minute = Number(m[2]);
  const ampm = m[3];

  if (ampm === "AM") {
    if (hour === 12) hour = 0;
  } else if (ampm === "PM") {
    if (hour !== 12) hour += 12;
  }

  return { h: hour, m: minute };
}
