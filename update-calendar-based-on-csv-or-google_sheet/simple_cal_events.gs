//push new events to calendar
function pushToCalendar() {
  
  //spreadsheet variables
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow(); 
  var range = sheet.getRange(2,1,lastRow,5);
  var values = range.getValues();   
  var updateRange = sheet.getRange('I1'); 
  
  //calendar variables
  var calendar = CalendarApp.getCalendarById('cal_id_or_email_here')
  
  //show updating message
  updateRange.setFontColor('red');
  
  var numValues = 0;
  for (var i = 0; i < values.length; i++) {     
    //check to see if name and type are filled out - date is left off because length is "undefined"
    if (values[i][0].length > 0) {
      
      //check if it's been entered before          
      if (values[i][2] != 'y') {                       
        
        //create event https://developers.google.com/apps-script/class_calendarapp#createEvent
        var newEventTitle = values[i][0];
        var setDescription = values[i][3];
        var setLocation = values[i][4];
        var setStartTime = values[i][1];
        var newEvent = calendar.createAllDayEvent(newEventTitle, setStartTime);
        
        //get ID
        var newEventId = newEvent.getId();
        
        //mark as entered, enter ID
        sheet.getRange(i+2,3).setValue('y');
        
      } //could edit here with an else statement
    }
    numValues++;
  }
  
  //hide updating message
  updateRange.setFontColor('white');

}

//add a menu when the spreadsheet is opened
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];  
  menuEntries.push({name: "CONFIRM Update Calendar", functionName: "pushToCalendar"}); 
  sheet.addMenu("CLICK HERE TO UPDATE CALENDAR", menuEntries);  
  
}
