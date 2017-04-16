// INSTALLATION
// Open: https://docs.google.com/spreadsheets/d/1mTrftieWpNqLJi1o9ZUFYeWtSDDB_x5MN4G8340uK7w/copy
// Go to: Tools > Script Editor...
// Copy and paste this file
// Then in the spreadsheet: Eventi Calendario > Crea
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Crea', functionName: 'createEvent_'}
  ];
  spreadsheet.addMenu('Eventi Calendario', menuItems);
}

function createEvent_() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheets()[0];
  sheet.activate();
  
  var cell = sheet.getRange("E1");
  var client = cell.getValue();
  cell = sheet.getRange("E2");
  var daysBuffer = cell.getValue();
  cell = sheet.getRange("E3");
  var callReminderDate = cell.getValue();
  cell = sheet.getRange("E4");
  var averageCapsulesPerDay = cell.getValue();
  
  var cal = CalendarApp.getDefaultCalendar();
  var description = 
      'Cliente: ' + client + 
      '\nGiorni di buffer: ' + daysBuffer + 
      '\nMedia capsule al giorno: ' + averageCapsulesPerDay;
  
  var event = cal.createAllDayEvent(
    "Chiamare " + client, 
    new Date(callReminderDate),
    { description: description });
  Logger.log('Event ID: ' + event.getId());
}