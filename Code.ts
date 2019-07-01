function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Calendar')
      .addItem('Sync now', 'syncToCalendar')
      .addToUi();
}

function syncToCalendar() {
  SpreadsheetApp.getUi()
      .alert('Not yet implemented');
}
