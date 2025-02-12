/**
 * Aggregates calendar events by prefix.
 */
function aggregateCalendarEvents() {
  const config = getConfig();
  const prefixes = getPrefixConfig();

  // Validation
  if (!config || config.startDate > config.endDate) {
    showErrorDialog('Check the start and end dates.');
    return;
  }
  if (!prefixes || prefixes.length === 0) {
    showErrorDialog('Set prefixes.');
    return;
  }

  const calendar = CalendarApp.getCalendarById(config.calendarId);
  if (!calendar) {
    showErrorDialog('Could not retrieve calendar.');
    return;
  }

  const events = calendar.getEvents(config.startDate, config.endDate);
  const wholeDate = getDatesBetween(config.startDate, config.endDate);
  const aggregation = aggregate(prefixes, events, wholeDate);

  const formattedDateRange = formatDateRange(config.startDate, config.endDate);
  outputResult(aggregation, wholeDate, formattedDateRange);
}

/**
 * Aggregates events.
 * @param {string[]} prefixes Prefix list
 * @param {GoogleAppsScript.Calendar.CalendarEvent[]} events Event list
 * @param {string[]} wholeDate Date list
 * @return {object} Aggregation result
 */
function aggregate(prefixes, events, wholeDate) {
  const prefixDurations = {};
  const prefixDailyDurations = {};

  wholeDate.forEach(date => {
    prefixDailyDurations[date] = {};
    prefixes.forEach(prefix => {
      prefixDailyDurations[date][prefix] = 0;
    });
  });

  events.forEach(event => {
    const eventDate = formatDate(event.getStartTime(), 'yyyy-MM-dd');
    if (wholeDate.includes(eventDate)) {
      prefixes.some(prefix => {
        if (event.getTitle().startsWith(prefix)) {
          const duration = event.getEndTime().getTime() - event.getStartTime().getTime();
          prefixDurations[prefix] = (prefixDurations[prefix] || 0) + duration;
          prefixDailyDurations[eventDate][prefix] += duration;
          return true; // break
        }
      });
    }
  });

  return { prefixDurations, prefixDailyDurations };
}

/**
 * Outputs the aggregation result.
 * @param {object} aggregation Aggregation result
 * @param {string[]} wholeDate Date list
 * @param {string} outputSheetname Output sheet name
 */
function outputResult(aggregation, wholeDate, outputSheetname) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(outputSheetname);
  if (!sheet) {
    sheet = ss.insertSheet(outputSheetname);
  }
  sheet.clearContents();

  const sortedPrefixes = Object.keys(aggregation.prefixDurations).sort();

  const header = ['Date', ...sortedPrefixes, 'Total'];
  sheet.appendRow(header);

  const totalRow = ['Total by prefix'];
  let finalTotal = 0;
  sortedPrefixes.forEach(prefix => {
    const duration = aggregation.prefixDurations[prefix] || 0;
    const hours = formatDurationToHours(duration);
    finalTotal += hours;
    totalRow.push(hours);
  });
  totalRow.push(finalTotal);
  sheet.appendRow(totalRow);

  const detailHeader = ['Breakdown', ...sortedPrefixes.map(() => '---'), 'Total by day'];
  sheet.appendRow(detailHeader);

  wholeDate.forEach(date => {
    const row = [formatDate(date)];
    let total = 0;
    sortedPrefixes.forEach(prefix => {
      const duration = aggregation.prefixDailyDurations[date][prefix] || 0;
      const hours = formatDurationToHours(duration);
      row.push(hours);
      total += hours;
    });
    row.push(total);
    sheet.appendRow(row);
  });
}

/**
 * Formats date.
 * @param {Date|string} date Date
 * @param {string} format Format string (yyyy-MM-dd etc.)
 * @return {string} Formatted date string
 */
function formatDate(date, format = 'yyyy年M月d日') {
  return Utilities.formatDate(new Date(date), 'JST', format);
}

/**
 * Formats date range.
 * @param {Date} startDate Start date
 * @param {Date} endDate End date
 * @return {string} Formatted date range string
 */
function formatDateRange(startDate, endDate) {
  const formattedStartDate = formatDate(startDate);
  const formattedEndDate = formatDate(endDate);
  return `${formattedStartDate}~${formattedEndDate}`;
}

/**
 * Formats duration to hours (milliseconds -> hours).
 * @param {number} duration Milliseconds
 * @return {number} Hours
 */
function formatDurationToHours(duration) {
  return duration ? duration / (1000 * 60 * 60) : 0;
}

/**
 * Gets date list.
 * @param {Date} startDate Start date
 * @param {Date} endDate End date
 * @return {string[]} Date list (yyyy-MM-dd format)
 */
function getDatesBetween(startDate, endDate) {
  const dates = [];
  let currentDate = new Date(startDate);
  while (currentDate <= endDate) {
    dates.push(formatDate(currentDate, 'yyyy-MM-dd'));
    currentDate.setDate(currentDate.getDate() + 1);
  }
  return dates;
}

/**
 * Gets sheet.
 * @param {string} sheetName Sheet name
 * @return {GoogleAppsScript.Spreadsheet.Sheet} Sheet
 */
function getSheetFromActiveSpreadSheet(sheetName) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}

/**
 * Gets configuration.
 * @return {object} Configuration
 */
function getConfig() {
  const sheet = getSheetFromActiveSpreadSheet('config_main');
  if (!sheet) {
    showErrorDialog('config_main sheet not found.');
    return null;
  }
  const calendarId = sheet.getRange('B2').getValue();
  const startDateStr = sheet.getRange('B3').getValue();
  const endDateStr = sheet.getRange('B4').getValue();

  // Validate before converting strings to Date objects
  if (!startDateStr || !endDateStr) {
    return null;
  }

  const startDate = new Date(startDateStr);
  const endDate = new Date(endDateStr);

  return { calendarId, startDate, endDate };
}

/**
 * Gets prefixes.
 * @return {string[]} Prefix list
 */
function getPrefixConfig() {
  const sheet = getSheetFromActiveSpreadSheet('config_prefixes');
  if (!sheet) {
    showErrorDialog('config_prefixes sheet not found.');
    return null;
  }
  const prefixRange = sheet.getRange('A2:A' + sheet.getLastRow());
  return prefixRange.getValues().flat().filter(String);
}

/**
 * Shows error dialog.
 * @param {string} message Message
 */
function showErrorDialog(message) {
  Browser.msgBox('Error', message, Browser.Buttons.OK);
}