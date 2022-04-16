function onOpen() {

  "use strict";

  var spreadsheet;
  var menuItems = [{
    name: "Importar",
    functionName: "importFromCalendar"
  }];
  
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.addMenu('Importar agenda', menuItems);

}

function adjustDate(date) {

  if (date == '')
    throw ("Selecione a célula com a data que deseja importar!");
  
  return date.substring(3, 5) + "/" + date.substring(0, 2) + "/" + date.substring(6);

}

function importFromCalendar() {

  var sheet = SpreadsheetApp.getActiveSheet(); // The Sheet object
  
  // Parameters to access Google Calendar and import events
  var calendarId = sheet.getRange("C3").getValue().toString();
  
  if (calendarId == '')
    throw ("A célula C3 precisa conter o ID da agenda.");
  
  var calendar = CalendarApp.getCalendarById(calendarId); // The Calendar object

  // Filters to get events from calendar
  var startDate = new Date(adjustDate(sheet.getActiveCell().getValue().toString()) + " 00:00:00");
  var endDate = new Date(adjustDate(sheet.getActiveCell().getValue().toString()) + " 23:59:59");

  var column = sheet.getActiveCell().getColumn();

  // Retrieve events from calendar
  var events = calendar.getEvents(startDate, endDate);

  // Clear data before importing new ones
  var range = sheet.getRange(5, column, 95, 4);
  range.clearContent();

  if (events.length > 0) {

    for (i = 1; i < events.length - 1; i++) {

      var row = i + 4;

      // Formula to calculate duration
      var startDateCell = sheet.getRange(row, column + 1).getA1Notation();
      var endDateCell = sheet.getRange(row, column + 2).getA1Notation();
      var formula = "=(HOUR(" + endDateCell +")+MINUTE(" + endDateCell + ")/60)-(HOUR(" + startDateCell + ")+MINUTE(" + startDateCell + ")/60)";

      var event = [[events[i].getTitle(), events[i].getStartTime(), events[i].getEndTime(), formula]];
      
      var range = sheet.getRange(row, column, 1, 4);
      range.setValues(event);

      // Some formatting
      var cell = sheet.getRange(row, column);
      cell.setHorizontalAlignment('left');
      cell = sheet.getRange(row, column + 1);
      cell.setHorizontalAlignment('center');
      cell.setNumberFormat("dd/mm/yyyy hh:mm");
      cell = sheet.getRange(row, column + 2);
      cell.setHorizontalAlignment('center');
      cell.setNumberFormat("dd/mm/yyyy hh:mm");
      cell = sheet.getRange(row, column + 3);
      cell.setHorizontalAlignment('center');
      cell.setNumberFormat("0.00");

    }

  }

}
