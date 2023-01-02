function exportToCalendar() {
    var calendarId = "google calendar ID"; 
    var calendar = CalendarApp.getCalendarById(calendarId);
   
    var sheet = SpreadsheetApp.getActiveSheet();
   
    var events = sheet.getRange("first cell :last cell").getValues();
   
    for (x=0; x<events.length; x++) {
      
      var evt = events[x];
      // var startTime = evt[4];
      var caseId = evt[1];
      var title = evt[2];
      var action = evt[3];
      var endTime = evt[4];
   
      calendar.createAllDayEvent(title+" "+caseId+" "+action,endTime);
    }
    
  }