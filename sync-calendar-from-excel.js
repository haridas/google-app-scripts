/*
 *
 *
 */

// Some global variables.

// The Common calendar name.
var calendarName = "haridas.nss@gmail.com";

// Keep the list list of object with date, and description.
// We can use it to unbounded events from the calender.
var newEvents = new Object();

/*
  Workflow

1. User will add or remove rows from the sheet

2. Finaly User will click on SynctoCalendar menu.

3. Loop through all the rows and insert every events to the calender.

4. Remove duplicate events by checking event time, and title. Because of the AllDay events, Calendar will save the time in UTC format dispite of the all 
   Timezone settings.
   a. Get date from the sheet row, convert it to UTC time, that date is the startDate for our event.
   b. Find the End Date by adding 24 hours to startDate UTC.
   c. Get the events between this from calendar, since the calender will return all the events that are starting or ending on this intervel.
   d. We are only interested in ending events. To identify it check whether the event endDate comes inbetween our input UTC date Range.
   c. Push first event to hash, and remove all others which satisfy this condition.

5. Remove those events that are not been linked to the sheet,
   a. Get all events between predefined range of time. One year from now.
   b. Get each event and then loop through all sheet dates, to get a valid match for that event.
   c. Here also, we have to conver the sheet date to UTC equivalent. Then get the event start and end date. This will be in UTC.
   d. Check the row UTC date comes inbetween event start and end date, also make sure that the event name is equal.

6. Remove old Events.
   a. Pick a predefined range of time. One year back from now.
   b. Remove them. No need of UTC mess.
  

*/

function UpdateCalendar() {

  
  // The Range of date, we use it for
  // Deduplication. Only the dates comes under this range
  // will gets cleaned up.
  var dateRange = new Date();
  dateRange.setDate(dateRange.getDate() + 366);
  
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process
  var numRows = 5;   // Number of rows to process
  
  var columnCount = sheet.getLastColumn();
  var rowCount = sheet.getLastRow();
  
  var dataRange = sheet.getRange(startRow, 1, rowCount, columnCount);
  
  //Browser.msgBox(dataRange);
  
  var data = dataRange.getValues();
  
  var cal = CalendarApp.getCalendarById(calendarName);
  
  /*
  // DEBUG
  var eField = "";
  for(j in data){
    var row = data[j];
    eField += row[0] + " " + row[1] + "\n";
  }
  Browser.msgBox(eField);
  */
  
  //Browser.msgBox(data.length);
  
  // Add event for all the rows in the spread sheet.
  for (i in data) {
    
    var row = data[i];
    var title1 = row[0];  // First column
    var eventDate = row[1];       // Second column
    var title2 = row[3];
    var titlespace = ' ';
    var title = title1 + titlespace + title2; 
        
    // Remove partial dates
    if (eventDate == "" || eventDate == undefined)
    {
      continue;
    }
    
    var desc = row[4];
    var calendarDate = new Date(eventDate);
    calendarDate.setHours(0,0,0,0);
    //var allDayEventDate = new Date(calendarDate.getTime() + 86400000); 
    //var duplicate = eventDeduplicate(calendarDate,desc);    
    
    // Adding event to the calendar.
    var test_c = cal.createAllDayEvent(title, new Date(calendarDate), {description:desc});  
    
    var ts = test_c.getStartTime()
    var te = test_c.getEndTime()
  }


// Now appying the deduplication operation.


  var eventDate = "";
  var calendarDate = "";  
  
  for( i in data){
   
   var events = new Array();
   
   var row = data[i];
   var title1 = row[0];
   eventDate = row[1];
   var desc = row[3];
   var title2 = row[3];
   var titlespace = ' ';
   var title = title1 + titlespace + title2; 
   
    // Remove partial dates
    if (eventDate == "" || eventDate == undefined)
    {
      continue;
    }    
    
    calendarDate = new Date(eventDate); 
    
    //var newstarTime = Utilities.formatDate(calendarDate, "GMT-0530", "yyyy-MM-dd HH:mm:ss");
    
    calendarDate.setHours(0,0,0,0);    
    var startOfDay = new Date(calendarDate);
    startOfDay.setHours(0);
    startOfDay.setMinutes(0);
    startOfDay.setMilliseconds(0);
    
    startOfDay.setUTCHours(0);
    startOfDay.setMinutes(0);
    startOfDay.setSeconds(0);
    startOfDay.setMilliseconds(0);  
    var endOfDay = new Date(startOfDay.getTime() + 24 * 60 * 60 * 1000);

    
    //var sDate = calendarDate;
    //var eDate = new Date(calendarDate);
    //eDate.setHours(23,0,0,0);
 
     
    // We got this much duplicate dates.
    events = cal.getEvents(startOfDay, endOfDay);
    
    //var newEvents = cal.getEvents(sDate, eDate);
    
    //var scriptTz = Session.getTimeZone();
    //var calTz = cal.getTimeZone();
    
    var dflag = false;
    var eventTitles = [];
    
    // New implementation
    var dupEvents = new Object();
    for(index in events)
    {
      
      var event = events[index];
      
      var std = event.getStartTime();
      var end = event.getEndTime();
      
      
      if(end <= endOfDay && end >= startOfDay ){
      // Only events that finish before the end of the day would 
      // consider for the deduplication. Others are part of the next 
      // days event.
        if(dupEvents[event.getTitle()] == undefined){
          dupEvents[event.getTitle()] = event;
        }else{
          // There is already a same title available under the same date.
          // So delete this event
          event.deleteEvent();
        }
      }  
      
    }
    
    /*
    for(index in events){
      // keep one date for the reference.
      var event = events[index];
      
      //keep track of other titles also.
      eventTitles.push(event.getTitle());
      
      if(event.getTitle() == title && dflag == false){
        dflag = true;
        continue;
      }else{
        
        // There is duplicate with other title also.
        var count = 0;
        for(t in eventTitles){
          var dupt = eventTitles[t];
          if( event.getTitle() == dupt ){
            count ++;
          }
        }        
        
        // This is duplicate one.
        if(count > 1)
        {
         event.deleteEvent();
        }
        
      }    
    }  */
    
   }
   
   
    //Browser.msgBox(events.length);    
    
    // Above step removed all the duplicate dates that are there in the calender.
    // Now we need to remove the the events that are not part of the excel sheet.activate()
    // Because of this problem is completely adjoint one, we have to loop through the all the
    // One year events to wipe out unlinked events
    
    var now = new Date();
    var startDate = new Date((now.getMonth() + 1) + "/" + now.getDate() + "/" + now.getFullYear());
    var endDate = new Date((dateRange.getMonth() + 1) + "/" + dateRange.getDate() + "/" + dateRange.getFullYear());  
    var allEvents = cal.getEvents(startDate, endDate);
    
    //Browser.msgBox(allEvents.length);
    //Browser.msgBox(allEvents[0].getTitle() + allEvents[0].getEndTime() + allEvents[0]);    
    
      // DEBUG
      //var eField = "";
      //for(j in data){
      //  var row = data[j];
      //  eField += row[0] + " " + row[1] + "\n";
      //}
      //Browser.msgBox(eField);

    
    
    for(index in allEvents){
    
      var calEvent = allEvents[index];
      // Add event for all the rows in the spread sheet.

      
      var invalid = true;
      for (i in data) {
        
        var row = data[i];
        var title1 = row[0];  // First column
        var eventDate = row[1];       // Second column
        var title2 = row[3];
        var titlespace = ' ';
        var title = title1 + titlespace + title2; 
        
        // Remove partial dates
        if (eventDate == "" || eventDate == undefined)
        {
          continue;
        }
        
        var desc = row[3];
        var calendarDate = new Date(eventDate);    
        var allDayEventDate = new Date(calendarDate.getTime() + 86400000);    
       
         
        // Check for non exisiting events UTC compatible testing.
        var evStartDate = calEvent.getStartTime();
        var evEndDate = calEvent.getEndTime();        
        var evName = calEvent.getTitle();
                    
        var rowDate = new Date(calendarDate);
        rowDate.setUTCHours(0);
        rowDate.setMinutes(0);
        rowDate.setSeconds(0);
        rowDate.setMilliseconds(0);       
                
       
        //Check the event date and event name to remove unlinked
        //Calendar entries.      
        // Match using the UTC time, because we are using Allday events
        if( rowDate >= evStartDate && rowDate <= evEndDate && title == evName){
          
            // Found maching sheet element.
            invalid = false;
            break;
        }
     
      }
      
      if(invalid == true){
        // Delete Invalid event
        calEvent.deleteEvent();
      }    
      
    }
  
  }


/*


function deduplicaeEvents(){
  // Full cleaning of unlinked events from the calendar. 
  // and make sure that calendar has only those dates from
  // the excel sheet.
  
}


 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  
  var entries = [
    {
      name : "Update Calendar",
      functionName : "UpdateCalendar"
    },
    
  ];
    
  sheet.addMenu("Calendar", entries);
    
};


/*
Onedit event. This function will get invoked after finishing edit on a 
cell. So we can check for when it is 
*/

/*
function onEdit(e) {
  
  //Browser.msgBox("Edited at: " + new Date().toTimeString() );
  //Browser.msgBox(Object.keys(e) + e.source);  
  var sheet = e.source;
  var cellValue = sheet.getActiveRange().getValue();
  
  //Browser.msgBox(sheet.getActiveRange().getValue());
  //Browser.msgBox(sheet.range.columnStart + " " + sheet.range.rowStart);
  
  var dt = new Date(cellValue);
  
  var range = sheet.getDataRange();
  Browser.msgBox(range.getRow() + " " + range.getColumn() + " " + range.getA1Notation());
  
  // Validate for date.
  if(isNaN(dt.getDate())){
    return;
  }
  
  Browser.msgBox("New Date: " + dt);
  
   
}*/


//Remove already expired events.
function RemoveExpiredEvents(){

    var cal = CalendarApp.getCalendarById(calendarName);
    
    var oldDate = new Date();
    oldDate.setDate(oldDate.getDate() - 355);

    var now = new Date();
    now.setDate(now.getDate() - 2);
    
    var startDate = new Date((now.getMonth() + 1) + "/" + now.getDate() + "/" + now.getFullYear());
    var endDate = new Date((oldDate.getMonth() + 1) + "/" + oldDate.getDate() + "/" + oldDate.getFullYear());  
    var allEvents = cal.getEvents(endDate,startDate);
    
    for(i in allEvents){
      var event = allEvents[i];
      event.deleteEvent();  
    }    
