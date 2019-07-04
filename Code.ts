const defaultTime = {start: '18:30', end:'20:30'}

function onOpen(): void {
  SpreadsheetApp.getUi()
      .createMenu('Calendar')
      .addItem('Sync now', 'main')
      .addToUi();
}

function main(): void {
  // SpreadsheetApp.getUi()
  //     .alert('Not yet implemented');
  const sheet = SpreadsheetApp.getActive().getActiveSheet();
  const calName: string = SpreadsheetApp.getActiveSheet().getName();
  
  if (CalendarApp.getCalendarsByName(calName).length == 0) {
    // Calendar named same as sheet page doesn't exist. Make it
    CalendarApp.createCalendar(calName);
  }
  const cal = CalendarApp.getCalendarsByName(calName)[0];
  
  let rowN: number = findDataHeight(sheet)
  for (let rowI = 1; rowI <= rowN; rowI++) {
    // For each row. rowI goes from 1 .. last row (inclusive inclusive)
    createEvent(sheet.getRange(rowI, 1, rowI, 8), cal);
  }
}

function createEvent(row: GoogleAppsScript.Spreadsheet.Range, cal: GoogleAppsScript.Calendar.Calendar): GoogleAppsScript.Calendar.CalendarEvent {
  // Get Title
  let title: string = row.getCell(1,6).getValue();
  // Get location and description
  let optionals = {location: row.getCell(1,7).getValue(), description: row.getCell(1,8).getValue()};
  // Get days. startDate uses same value as endDate if no startDate specified
  let startDate: Date = row.getCell(1,1).getValue() === '' ? row.getCell(1,3).getValue() : row.getCell(1,1).getValue();
  let endDate: Date = row.getCell(1,3).getValue();
  // Get times or all day
  if (row.getCell(1,4).getValue() !== '') {
    // Time specified. Use it
    startDate.setHours(row.getCell(1,4).getValue().getHours());
    startDate.setMinutes(row.getCell(1,4).getValue().getMinutes());
    endDate.setHours(row.getCell(1,5).getValue().getHours());
    endDate.setMinutes(row.getCell(1,5).getValue().getMinutes());
  } else if (row.getBackground() === '#ffffff' && defaultTime !== null) {
    // Backround white, use default time
    startDate.setHours(Number(defaultTime.start.substring(0, 2)))
    startDate.setMinutes(Number(defaultTime.start.substring(3, 5)))
    endDate.setHours(Number(defaultTime.end.substring(0, 2)))
    endDate.setMinutes(Number(defaultTime.end.substring(3, 5)))
  } else {
    // Create as all day event
    return cal.createAllDayEvent(title, startDate, endDate, optionals)
  }
  // Create as normal event
  return cal.createEvent(title, startDate, endDate, optionals)
}

function findDataHeight(sheet: GoogleAppsScript.Spreadsheet.Sheet): number {
  let i: number = 1
  let cellText: string
  do {
    cellText = sheet.getRange(i, 3).getValue();

    i++;
  } while (cellText !== '')
  return i - 1;
}
