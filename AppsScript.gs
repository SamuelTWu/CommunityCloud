function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Process Events')
      .addItem('Add Approved Events to Calendar', 'fillCalendar')
      .addToUi();
}
 
 
function addDataValidation(string_list){
  //define parameters
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Calendar Event Creator");
  var Bvals = sheet.getRange("B9:B").getValues();
  var lastRow = Bvals.filter(String).length;
  
  //place all the dropdowns
  var rule = SpreadsheetApp.newDataValidation()
     .requireValueInList(["Pending","Approved"], true)
    .setAllowInvalid(false)
    .build();

  //pull groups from library
  const dict = string_list.split(/\r?\n/);

  //SpreadsheetApp.getUi().alert(dict[0]);
  
  var setApprovalStatus =  sheet.getRange(9,1,lastRow).setDataValidation(rule).setValue("Pending");
  for(var i = 9; i < 9+lastRow; i++)
  {
    for(var j = 0; j < lastRow; j++)
    {
      //SpreadsheetApp.getUi().alert(dict[j]);
      var getID = sheet.getRange(i,2).getValue();
      if(getID == dict[j])
      {
        var setApprovalStatus =  sheet.getRange(i,1).setDataValidation(rule).setValue("Approved");
      }
    }
  }


  
}

function onEdit() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var studentDB = ss.getSheetByName("Student DB");
  var studentOrg = ss.getRange("B2").getValues();
  
  var GROUPIND = 11;
  var NAMEIND = 3;

  var name =  SpreadsheetApp.getActiveSheet().getName();
  //SpreadsheetApp.getUi().alert(name);

  var row = SpreadsheetApp.getActiveRange().getRow();
  //SpreadsheetApp.getUi().alert(row);

  var col = SpreadsheetApp.getActiveRange().getColumn();
  //SpreadsheetApp.getUi().alert(col);

  
  if(name == "Calendar Event Creator" && row == 2 && col == 2)
  {
    var Bvals = studentDB.getRange("B1:B").getValues();
    var lastRow = Bvals.filter(String).length;
    
    //find studentRow
    var studentRow = 0;
    for(var i = 2; i <= lastRow; i++)
    {
      var studentName = studentDB.getRange(i,NAMEIND).getValue();
      if(studentName == studentOrg)
      {
        studentRow = i;
      }
    }

    var groups = "No Data"
    if(studentRow != 0)
    {
      groups = studentDB.getRange(studentRow,GROUPIND).getValue();
    }
    addDataValidation(groups);
  }
}
   

function fillCalendar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calendarSheet = ss.getSheetByName("Calendar Event Creator");

  var serviceSheet = ss.getSheetByName("Service DB");
  var studentSheet = ss.getSheetByName("Student DB");

  var Avals_service = serviceSheet.getRange("A2:A").getValues();
  var lastRow_service = Avals_service.filter(String).length;

  var Avals_student = studentSheet.getRange("A2:A").getValues();
  var lastRow_student = Avals_student.filter(String).length;
  

  startRow = 9;
  indApprov = 1;
  indTitle = 4;
  indDescription = 17;
  indVolunteers = 16;

  indRecurring = 5;

  indStartDate_recurring = 6;
  indEndDate_recurring = 7;
  indStartTime_recurring = 8;
  indEndTime_recurring = 9;
  indWeekDays_recurring = 10;

  indStartDate_single = 11;
  indStartTime_single = 12;
  indEndTime_single = 13;
  
  indLocation = 15;
  indServiceID = 3;
  indCalID = 5;
  indsignUpLink = 24;
  
  indEventConfirm = 7;

  var servicePartnerLibrary = "";
  
  var calID = calendarSheet.getRange(2,indCalID).getValue();
  
 
  for (var i = 2; i <= lastRow_service ; i++) 
  {
      //data called from Service DB
      var title = serviceSheet.getRange(i,indTitle).getValue();
      var description = serviceSheet.getRange(i,indDescription).getValue();
      var numVolunteers = serviceSheet.getRange(i,indVolunteers).getValue();
      var location = serviceSheet.getRange(i,indLocation).getValue();
      var signUpLink = serviceSheet.getRange(i,indsignUpLink).getValue();
      var isRecurring = serviceSheet.getRange(i,indRecurring).getValue();
      var serviceID = serviceSheet.getRange(i,indServiceID).getValue();
      var finalDescription = description + "\nPreferred number of volunteers: "+numVolunteers + "\nSign up Link: " + signUpLink;

      //data called from Calendar Event Creator
      var approvalStatus = calendarSheet.getRange(i+7,indApprov).getValue();
      var eventConfirm = calendarSheet.getRange(i+7,indEventConfirm).getValue();//TODO: gotta link this to a dict with the service ID in the student DB
      
      if (approvalStatus == "Approved" && eventConfirm == "")
      {
          Logger.log(title+" is Approved"); 
          if (isRecurring== "Yes, this is a weekly event")
          {
              //Logger.log(title+" is a recurring event"); 
              var startDate = serviceSheet.getRange(i,indStartDate_recurring).getValue();
              var endDate = serviceSheet.getRange(i,indEndDate_recurring).getValue();
              var startTime = serviceSheet.getRange(i,indStartTime_recurring).getValue();
              var endTime = serviceSheet.getRange(i,indEndTime_recurring).getValue();
              var weekDays = serviceSheet.getRange(i,indWeekDays_recurring).getValue();
            
              var formattedStart = Utilities.formatDate(new Date(startDate), 'America/New_York', 'MMMM dd, yyyy');
              var formattedEnd = Utilities.formatDate(new Date(endDate), 'America/New_York', 'MMMM dd, yyyy');
              var formattedSTime = Utilities.formatDate(new Date(startTime), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "HH:mm:ss")
              var formattedETime = Utilities.formatDate(new Date(endTime), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "HH:mm:ss")

              var startDateandTime = (formattedStart+" "+formattedSTime);
              var endDateandTime = (formattedStart+" "+formattedETime);

              
              
              var days = weekDays.split(', ').map(function(i) { return CalendarApp.Weekday[i]; });

              var eventSeries = CalendarApp.getCalendarById(calID).createEventSeries(
                title,
                new Date(startDateandTime),
                new Date(endDateandTime),
                CalendarApp.newRecurrence().addWeeklyRule()
                    .onlyOnWeekdays(days)
                    .until(new Date(formattedEnd)),
                    {location: location, description: finalDescription});
            
            //var setEventConfirm = sheet.getRange(i,indEventConfirm).setValue('Event Series ID: ' + eventSeries.getId());
          }
          else if(isRecurring == "No, this is a single event")
          {
              Logger.log(title+" is a single event"); 
              var startDate = serviceSheet.getRange(i,indStartDate_single).getValue();
              var startTime = serviceSheet.getRange(i,indStartTime_single).getValue();
              var endTime = serviceSheet.getRange(i,indEndTime_single).getValue();
            
              var formattedStart = Utilities.formatDate(new Date(startDate), 'America/New_York', 'MMMM dd, yyyy');
              var formattedSTime = Utilities.formatDate(new Date(startTime), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "HH:mm:ss")
              var formattedETime = Utilities.formatDate(new Date(endTime), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "HH:mm:ss")

              var startDateandTime = (formattedStart+" "+formattedSTime);
              var endDateandTime = (formattedStart+" "+formattedETime);
              
              Logger.log(startDateandTime);
              Logger.log(endDateandTime);

              var eventSeries = CalendarApp.getCalendarById(calID).createEvent(
                title,
                new Date(startDateandTime),
                new Date(endDateandTime),
                {location: location, description: finalDescription});
            //Logger.log('Event Series ID: ' + eventSeries.getId());
            //var setEventStatus = sheet.getRange(i,indEventStatus).setValue('Event Series ID: ' + eventSeries.getId());
          }

          //update "Service Partners" library
          servicePartnerLibrary+=serviceID+"\n";
      }
      else
      {
         Logger.log("No Events Found"); 
      }
  }
  //input Service Partners Library in Student DB
  updateServicePartnersLibrary(lastRow_student, studentSheet, calID, servicePartnerLibrary);

}

function updateServicePartnersLibrary(lastRow, studentSheet, calID, servicePartnerLibrary)
{
  indCalID = 10;
  indServicePartner = 11;
  for(var i = 2; i < lastRow; i++)
  {
    var tempCalID = studentSheet.getRange(i,indCalID).getValue();
    if(tempCalID == calID)
    {
      var cell = studentSheet.getRange(i,indServicePartner).setValue(servicePartnerLibrary);
    }
  }
}

