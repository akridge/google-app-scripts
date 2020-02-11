/*
Project: Master Calendar Script
File: MASTER_CALENDAR_SCRIPT.gs
Michael Akridge
------------------------------------------------------------------------------------------
Description:
------------------------------------------------------------------------------------------
This script is designed to update a master calendar with individual calendar events from multi-user calendars. 

Events are only copied when they contain one of the following tags:
Leave (i.e. Annaul Leave, Sick Leave, Family Leave)
Sick
Vacation
AWS
Telework
Offsite
Off-site

The above tags are grouped as follows:
Leave: Leave, Sick, Vacation
AWS: AWS
Telework: Telework
Off-site; Off-site, Offsite
------------------------------------------------------------------------------------------
Additional Setup:
------------------------------------------------------------------------------------------
1. Define the master calendar Id below to populate the master calendar
2. Define the email_list google folder id below 
3. Populate export.csv google sheet with emails
4. Setup time-based triggers:
----a. Calendars must be shared with the user who is running the script in order to capture the events
----b. setup two triggers. 1am and 10am
------------------------------------------------------------------------------------------
Change Log:
------------------------------------------------------------------------------------------
Date: 2019
Note: Combined update with master. Fine tuned. google sheet/CSV email list code
Author: M. Akridge

Date: 2018
Note:  fork & bug fixes
Author: B. Gough

Date: 2018
Note: Google Sheet for staffing list Update & bug fixes
Author: N. Weeks

Date: 2016
Note: Source Code
Author: S. Mehta
------------------------------------------------------------------------------------------
*/

/*
------------------------------------------------------------------------------------------
Define Variable Start
------------------------------------------------------------------------------------------
*/
// Define number of days into the future to update
var updateDays = 14;

// Define the tags used to flag events that should be copied from an employee's calendar into the master calendar
var eventTags = [
  {value: 'Leave', group: 'Leave'}
  ,{value: 'Sick', group: 'Leave'}
  ,{value: 'Vacation', group: 'Leave'}
  ,{value: 'AWS', group: 'AWS'}
  ,{value: 'Flex', group: 'AWS'}
  ,{value: 'Telework', group: 'Telework'}
  ,{value: 'Off-site', group: 'Off-site'}
  ,{value: 'Offsite', group: 'Off-site'}
];
for (var i = 0, l = eventTags.length; i < l; i++) {
  eventTags[i].re = new RegExp('\\b' + eventTags[i].value + '\\b', 'i') // Build regular expression for each tag
}

// Define the master calendar Id
//  TEST CAL: 'INSERT_ID_HERE'
//  master calendar id: 'INSERT_ID_HERE'

var masterCalendarId = 'INSERT_ID_HERE';

// Define employee calendar Ids
var employeeCalendarIds = [];
employeeCalendarIds = importData("exports.csv");

/*
------------------------------------------------------------------------------------------
function importData start

Details:
This function will import a CSV email list thats is defined in a google folder. 
importdata2 below will work with google sheets.

Setup: 
// Your google drive should contain a folder called  email_list2. Define the folder id below.
// in this folder place your exports.csv file which contains the emails of staff
// the emails in the file are basically the unique calendar IDs of each staff's calendar, however, some self created calendars might have
// a calendar ID that is  composed of random numeric and alphabetical characters
------------------------------------------------------------------------------------------
*/ 
function importData(employeeFile) {

  //email_list folder id
  var id = 'INSERT_folder_id_number_here';
  var employee_ID_array = [];
  
  var fSource = DriveApp.getFolderById(id);
  var fi = fSource.getFilesByName(employeeFile);

  if ( fi.hasNext() ) { // continue if "exports.csv" file exists in the reports folder
    
    var file = fi.next();
    
    var csv = file.getBlob().getDataAsString();
    var csvData = CSVToArray(csv); 
     
    for ( var i=0, lenCsv=csvData.length; i<lenCsv; i++ ) {
      
      
      if(csvData[i][0] != '' | null ){
        
       employee_ID_array[i] = csvData[i][0];
       //Logger.log(employee_ID_array[i]);
        
      }
      
      
    }
    
    if (employee_ID_array[0].indexOf("@") == -1){
      employee_ID_array = importData2(employeeFile);
    }
   
  } else {
    throw new Error("File: " + employeeFile + " not found. Has the file name changed?");
  }
  
  //Logs the size of the employee id array and shows all the emails in there
  Logger.log("Array size: " + employee_ID_array.length);
  Logger.log("Array contents: " + employee_ID_array);
  return(employee_ID_array);
}
/*
------------------------------------------------------------------------------------------
function CSVToArray start
------------------------------------------------------------------------------------------
*/ 
  
  function CSVToArray( strData, strDelimiter ) {
  // Check to see if the delimiter is defined. If not,
  // then default to COMMA.
  strDelimiter = (strDelimiter || ",");

  // Create a regular expression to parse the CSV values.
  var objPattern = new RegExp(
    (
      // Delimiters.
      "(\\" + strDelimiter + "|\\r?\\n|\\r|^)" +

      // Quoted fields.
      "(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|" +

      // Standard fields.
      "([^\"\\" + strDelimiter + "\\r\\n]*))"
    ),
    "gi"
  );

  // Create an array to hold our data. Give the array
  // a default empty first row.
  var arrData = [[]];

  // Create an array to hold our individual pattern
  // matching groups.
  var arrMatches = null;

  // Keep looping over the regular expression matches
  // until we can no longer find a match.
  while (arrMatches = objPattern.exec( strData )){

    // Get the delimiter that was found.
    var strMatchedDelimiter = arrMatches[ 1 ];

    // Check to see if the given delimiter has a length
    // (is not the start of string) and if it matches
    // field delimiter. If id does not, then we know
    // that this delimiter is a row delimiter.
    if (
      strMatchedDelimiter.length &&
      (strMatchedDelimiter != strDelimiter)
    ){

      // Since we have reached a new row of data,
      // add an empty row to our data array.
      arrData.push( [] );

    }

    // Now that we have our delimiter out of the way,
    // let's check to see which kind of value we
    // captured (quoted or unquoted).
    if (arrMatches[ 2 ]){

      // We found a quoted value. When we capture
      // this value, unescape any double quotes.
      var strMatchedValue = arrMatches[ 2 ].replace(
        new RegExp( "\"\"", "g" ),
        "\""
      );

    } else {

      // We found a non-quoted value.
      var strMatchedValue = arrMatches[ 3 ];

    }

    // Now that we have our value string, let's add
    // it to the data array.
    arrData[ arrData.length - 1 ].push( strMatchedValue );
  }

  // Return the parsed data.
  return( arrData );
}

/*
------------------------------------------------------------------------------------------
function importData2 start

Details:
This function is called in the case that importData fails because it is calling a google sheet instead of a normal csv. 
This will then work with the googlesheet of the same name
------------------------------------------------------------------------------------------
*/ 
function importData2(file){ 
  var employee_ID_array = [];
  try{
    var sheet = SpreadsheetApp.openById(DriveApp.getFilesByName(file).next().getId());
    var values = sheet.getSheetValues(1, 1, sheet.getLastRow(), 1);
    for (var i = 0; i < sheet.getLastRow(); i++){
      employee_ID_array.push(values[i][0]);
    }
    return employee_ID_array;
  }catch(e){
    throw new Error("File: " + file + " not found");
  }
}
  
/*
------------------------------------------------------------------------------------------
function listCalendars start - uncomment to allow


function listCalendars() {
  // List Name and Id for all subscribed calendars
  var calendars = CalendarApp.getAllCalendars();
  for (var i = 0, l = calendars.length; i < l; i++) {
    var calendar = calendars[i];
    Logger.log("Calendar " + i) 
    Logger.log("Name = " + calendar.getName());
    Logger.log("ID = " + calendar.getId()); 
  }
}
------------------------------------------------------------------------------------------
*/ 

/*
------------------------------------------------------------------------------------------
function updateMasterCalendar start
------------------------------------------------------------------------------------------
*/ 
function updateMasterCalendar() {

  // Get date range to update
  var startDate = new Date();
  startDate.setHours(0, 0, 0);
  var endDate = new Date(startDate.getTime());
  endDate.setDate(endDate.getDate() + updateDays);
  
  // Initialize object to hold events that will be added to masterCalendar
  var newMasterCalendarEvents = {};
  
  // Cycle through all employee calendars and get events that have been tagged to the master caledar
  for (var i = 0, l = employeeCalendarIds.length; i < l; i++) {
      
    var employeeCaledarId = employeeCalendarIds[i];
 
    // Get employee's name from ID where ID is firstname.middle.lastname@noaa.gov
    var employeeNameParts = employeeCaledarId.split('@')[0].split('.');
    var employeeFirstName = employeeNameParts[0];
    employeeFirstName = employeeFirstName.charAt(0).toUpperCase() + employeeFirstName.substr(1).toLowerCase();
    var employeeLastName = employeeNameParts[employeeNameParts.length - 1];
    employeeLastName = employeeLastName.charAt(0).toUpperCase() + employeeLastName.substr(1).toLowerCase();
    var employeeDisplayName = employeeFirstName.substr(0,1) + '.' + employeeLastName;
    
    // Get employee's calendar
    var employeeCalendar = CalendarApp.getCalendarById(employeeCaledarId);

    if (!employeeCalendar) {
      try {
        employeeCalendar = CalendarApp.subscribeToCalendar(employeeCaledarId, { hidden: false });
        Logger.log('Subscribed to calendar ' + employeeCalendar.getName());
      } catch (e) {
        Logger.log('Unable to subscribe to calendar ' + employeeCalendar + ", error: " + e.message);
        continue;
      }
    }

    // Cycle through all events in employee's calendar
    var employeeCalendarEvents = employeeCalendar.getEvents(startDate, endDate);
    
    for (var j = 0, m = employeeCalendarEvents.length; j < m; j++) {
    
      var employeeCalendarEvent = employeeCalendarEvents[j];         
      
      // Cycle through all event tags
      for (var k = 0, n = eventTags.length; k < n; k++) {
        
        // If event title contains tag, save event information so that it can be added to the master event
        // Note: We don't add the event now because we want to group similar events for multiple employees
        // into a single event on the master calendar to save space.
        if (eventTags[k].re.test(employeeCalendarEvent.getTitle()) === true) {
          
          // Get event information and use it to build a unique event key
          var eventGroup = eventTags[k].group;
          var eventIsAllDayEvent = employeeCalendarEvent.isAllDayEvent();
          if (eventIsAllDayEvent) {
            var eventStartDateTime = employeeCalendarEvent.getAllDayStartDate();
            var eventEndDateTime = employeeCalendarEvent.getAllDayEndDate();
          } else {
            var eventStartDateTime = employeeCalendarEvent.getStartTime();
            var eventEndDateTime = employeeCalendarEvent.getEndTime();
          }
          var eventKey = eventGroup + eventStartDateTime.toString() + eventEndDateTime.toString();
          
          if (!newMasterCalendarEvents.hasOwnProperty(eventKey)) {
            // Create event
      
            newMasterCalendarEvents[eventKey] = {
              title: eventGroup + ': ' + employeeDisplayName, 
              startDateTime: eventStartDateTime,
              endDateTime: eventEndDateTime,
              isAllDayEvent: eventIsAllDayEvent,
            };
          } else {
            // Add employee to event
            newMasterCalendarEvents[eventKey].title = newMasterCalendarEvents[eventKey].title + ', ' + employeeDisplayName;
          }
          
          break; // Only add event of first tag found
          
        }
        
      } // for eventTags  
    
    } // for employeeCalendarEvents

  } // for employeeCalendarIds
  
  // Get reference to master calendar
  var masterCalendar = CalendarApp.getCalendarById(masterCalendarId);
  
  // Delete exists events from master calendar
  var masterCalendarEvents = masterCalendar.getEvents(startDate, endDate);
  for (var i = 0, l = masterCalendarEvents.length; i < l; i++) {
    Utilities.sleep(1000);
    masterCalendarEvents[i].deleteEvent();
  }
   
  // Add new events to master calendar
  for (var eventKey in newMasterCalendarEvents) {
    if (newMasterCalendarEvents.hasOwnProperty(eventKey)){
      var event = newMasterCalendarEvents[eventKey];
      if (event.isAllDayEvent) {
        Utilities.sleep(1000);
        masterCalendar.createAllDayEvent(event.title, event.startDateTime, event.endDateTime);
      } else {
        Utilities.sleep(1000);
        masterCalendar.createEvent(event.title, event.startDateTime, event.endDateTime);
      }
    }
  }
  
}
