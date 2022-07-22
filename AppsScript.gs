calendarSheet_FIRSTROW = 9;
calendarSheet_STATUS = 1;
calendarSheet_SERVICEID = 2;
calendarSheet_EVENTCONFIRM = 6;
calendarSheet_CALID = 5;
calendarSheet_EVENTCONFIRMATION = 6;

studentSheet_EVENTCONFIRM = 12;
studentSheet_FIRSTROW = 2;
studentSheet_GROUP = 11;
studentSheet_NAME = 3; 
studentSheet_CALID  = 10;
studentSheet_EVENTCONFIRMATION = 12;
studentSheet_SERVICEPARTNER = 11;

serviceSheet_FIRSTROW = 2;
serviceSheet_APPROVE = 2;
serviceSheet_TIMESTAMP = 1;
serviceSheet_TITLE = 5;
serviceSheet_DATE = 7;
serviceSheet_DESCRIPTION = 11;
serviceSheet_VOLUNTEERS = 10;
serviceSheet_LOCATION = 9;
serviceSheet_SERVICEID = 4;
serviceSheet_SIGNUPLINK = 17;
serviceSheet_CONFIRM_SIGNUP = 16; 




function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Process Events')
      .addItem('Add Approved Events to Calendar', 'fillCalendar')
      .addToUi();
}
 
 
function addDataValidation(string_list){
  //define parameters
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calendarSheet = ss.getSheetByName("Calendar Event Creator");
  var serviceSheet = ss.getSheetByName("Service DB");

  var Bvals = calendarSheet.getRange("B9:B").getValues();
  var calendar_lastRow = Bvals.filter(String).length;

  var Avals = serviceSheet.getRange("A2:A").getValues();
  var service_lastRow = Avals.filter(String).length;
  
  //place all the dropdowns
  var rule = SpreadsheetApp.newDataValidation()
     .requireValueInList(["Pending","Approved", "Unapproved"], true)
    .setAllowInvalid(false)
    .build();

  //pull groups from library
  const dict = string_list.split(/\r?\n/);
  
  var setApprovalStatus =  calendarSheet.getRange(calendarSheet_FIRSTROW,calendarSheet_STATUS,calendar_lastRow).setDataValidation(rule).setValue("Pending");

  for(var i = calendarSheet_FIRSTROW; i < calendarSheet_FIRSTROW+calendar_lastRow; i++)
  {
    var getID = calendarSheet.getRange(i,calendarSheet_SERVICEID).getValue();

    //check if service is approved
    for(var j = serviceSheet_FIRSTROW; j < serviceSheet_FIRSTROW+service_lastRow; j++)
    {
      
      var serviceID = serviceSheet.getRange(j, serviceSheet_SERVICEID).getValue();

      if(serviceID == getID)
      {
        var approvalStatus = serviceSheet.getRange(j, serviceSheet_APPROVE).getValue();
        if(approvalStatus == "Unapproved")
        {
          calendarSheet.getRange(i,calendarSheet_STATUS).setDataValidation(rule).setValue("Unapproved");
        }
      }
    }

    //check if event is in student service partners
    for(var j = 0; j < dict.length; j++)
    {
      if(getID == dict[j])
      {
        var setApprovalStatus =  calendarSheet.getRange(i,calendarSheet_STATUS).setDataValidation(rule).setValue("Approved");
      }
    }
  }
}

function addApprovalServiceDB()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var serviceSheet = ss.getSheetByName("Service DB");
  var Avals = serviceSheet.getRange("A1:A").getValues();
  var lastRow = Avals.filter(String).length; 

  //dropdown rules
  var rule = SpreadsheetApp.newDataValidation()
     .requireValueInList(["Unapproved","Approved"], true)
    .setAllowInvalid(false)
    .build();
  
  serviceSheet.getRange(lastRow,serviceSheet_APPROVE).setDataValidation(rule).setValue("Unapproved");
  
}

function addEventConfirm(string_list){
  
  //define parameters
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Calendar Event Creator");
  var Bvals = sheet.getRange("B9:B").getValues();
  var lastRow = Bvals.filter(String).length;

  //clear all Event Confirm
  sheet.getRange(calendarSheet_FIRSTROW,calendarSheet_EVENTCONFIRM,calendarSheet_FIRSTROW+lastRow).setValue("");

  //create dict from string_list
  const dict = string_list.split(/\r?\n/);

  //fill Event Confirm from dict
  for(var i = calendarSheet_FIRSTROW; i < calendarSheet_FIRSTROW+lastRow; i++)
  {
    for(var j = 0; j < dict.length; j++)
    {
      var getID = sheet.getRange(i,calendarSheet_SERVICEID).getValue();
      var confirmIndex = dict[j].split(":")[0];
      var confirmID = dict[j].split(":")[1];

      if(getID == confirmIndex)
      {
        //fill cell with event ID
        sheet.getRange(i,calendarSheet_EVENTCONFIRM).setValue(confirmID);
      }
    }
  }

}

function onEdit() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var studentDB = ss.getSheetByName("Student DB");
  var studentOrg = ss.getRange("B2").getValues();
  
  var name =  SpreadsheetApp.getActiveSheet().getName();
  //SpreadsheetApp.getUi().alert(name);

  var row = SpreadsheetApp.getActiveRange().getRow();
  //SpreadsheetApp.getUi().alert(row);

  var col = SpreadsheetApp.getActiveRange().getColumn();
  //SpreadsheetApp.getUi().alert(col);

  
  if(name == "Calendar Event Creator" && row == 2 && col == 2)
  {
    SpreadsheetApp.getUi().alert("Loading "+studentOrg+" Information");
    var Bvals = studentDB.getRange("B1:B").getValues();
    var lastRow = Bvals.filter(String).length;
    
    //find studentRow
    var studentRow = -1;
    for(var i = calendarSheet_SERVICEID; i <= calendarSheet_SERVICEID+lastRow; i++)
    {
      var studentName = studentDB.getRange(i,studentSheet_NAME).getValue();
      if(studentName == studentOrg)
      {
        studentRow = i;
      }
    }

    var groups
    var eventDict;

    if(studentRow != -1)
    {
      groups = studentDB.getRange(studentRow,studentSheet_GROUP).getValue();
      eventDict = studentDB.getRange(studentRow,studentSheet_EVENTCONFIRM).getValue();
    }
    addDataValidation(groups);
    addEventConfirm(eventDict);
    SpreadsheetApp.getUi().alert("Successfully Retrieved "+studentOrg+" Information");
  }
}
  
function addEvents(i, calID){
  //data called from Service DB
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calendarSheet = ss.getSheetByName("Calendar Event Creator");

  var serviceSheet = ss.getSheetByName("Service DB");
  var studentSheet = ss.getSheetByName("Student DB");

  var title = serviceSheet.getRange(i,serviceSheet_TITLE).getValue();
  var description = serviceSheet.getRange(i,serviceSheet_DESCRIPTION).getValue();
  var numVolunteers = serviceSheet.getRange(i,serviceSheet_VOLUNTEERS).getValue();
  var location = serviceSheet.getRange(i,serviceSheet_LOCATION).getValue();
  var signUpLink = serviceSheet.getRange(i,serviceSheet_SIGNUPLINK).getValue();

  var finalDescription = description + "\nPreferred number of volunteers: "+numVolunteers + "\nSign up Link: " + signUpLink;

  var date_time = serviceSheet.getRange(i,serviceSheet_DATE).getValue();
  const date_time_list = date_time.split(/\r?\n/);

  
  
  
  for(var j = 0; j < date_time_list.length; j++)
  {
    var date_time = date_time_list[j].split(",");
    var type = date_time[0];
    
    var startDate = date_time[1].split("/");
    startDate = startDate[1]+"/"+startDate[0]+"/"+startDate[2];

    var endDate = date_time[3].split("/");
    endDate = endDate[1]+"/"+endDate[0]+"/"+endDate[2];
    
    var startTime = date_time[2];
    var endTime = date_time[4];

    //Logger.log(endDate);
    if(type=="SINGLE")
    {
      var formattedStart = Utilities.formatDate(new Date(startDate), 'America/New_York', 'MMMM dd, yyyy');
      var formattedEnd = Utilities.formatDate(new Date(endDate), 'America/New_York', 'MMMM dd, yyyy');
      

      var startDateandTime = (formattedStart+" "+startTime);
      var endDateandTime = (formattedEnd+" "+endTime);
      Logger.log(startDateandTime +", "+endDateandTime);
      

      var eventSeries = CalendarApp.getCalendarById(calID).createEvent(
        title,
        new Date(startDateandTime),
        new Date(endDateandTime),
        {location: location, description: finalDescription});
    }

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
  
  //These are static variables for the google sheets index

  var servicePartnerLibrary = "";
  var eventConfirmLibrary = "";
  
  var calID = calendarSheet.getRange(serviceSheet_FIRSTROW,calendarSheet_CALID).getValue();
  
  Logger.log(lastRow_service)
  for (var i = serviceSheet_FIRSTROW; i < serviceSheet_FIRSTROW+lastRow_service; i++) 
  {
      //data called from Calendar Event Creator
      var approvalStatus = calendarSheet.getRange(i+7,calendarSheet_STATUS).getValue();
      var eventConfirm = calendarSheet.getRange(i+7,calendarSheet_EVENTCONFIRMATION).getValue();
      
      if (approvalStatus == "Approved")
      {
          if(eventConfirm == "")
          {
            //create calendar events from list
            addEvents(i, calID);
            calendarSheet.getRange(i+7,calendarSheet_EVENTCONFIRMATION).setValue("Confirmed");
          }
          //update servicePartnerLibrary
          var serviceID = serviceSheet.getRange(i,serviceSheet_SERVICEID).getValue();
          servicePartnerLibrary+=serviceID+"\n";

          //update eventConfirmationLibrary (This has depricated due to mutliple event dates being present)
          var confirmID =  calendarSheet.getRange(i+7,calendarSheet_EVENTCONFIRMATION).getValue();
          eventConfirmLibrary += (serviceID+":"+confirmID+"\n");
      }
      else
      {
         Logger.log("No Events Found For Row: "+(i+7)); 
      }
  }
  //update Service Partners Library in Student DB
  updateServicePartnersLibrary(lastRow_student, studentSheet, calID, servicePartnerLibrary);

  //update Service Confirmation (This has depricated due to mutliple event dates being present)
  updateServiceConfirmation(lastRow_student, studentSheet, calID, eventConfirmLibrary);

  //re-fill calendar event sheet
  onEdit();

}

function updateServiceConfirmation(lastRow, studentSheet, calID, eventConfirmLibrary){
  for(var i = studentSheet_FIRSTROW; i < studentSheet_FIRSTROW+lastRow; i++)
  {
    var tempCalID = studentSheet.getRange(i,studentSheet_CALID).getValue();
    if(tempCalID == calID)
    {
      var cell = studentSheet.getRange(i,studentSheet_EVENTCONFIRMATION).setValue(eventConfirmLibrary);
    }
  }
}

function updateServicePartnersLibrary(lastRow, studentSheet, calID, servicePartnerLibrary){
  for(var i = studentSheet_FIRSTROW; i < studentSheet_FIRSTROW+lastRow; i++)
  {
    var tempCalID = studentSheet.getRange(i,studentSheet_CALID).getValue();
    if(tempCalID == calID)
    {
      var cell = studentSheet.getRange(i,studentSheet_SERVICEPARTNER).setValue(servicePartnerLibrary);
    }
  }
}

function fillSignupLink(link){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var serviceSheet = ss.getSheetByName("Service DB");
  var Avals_service = serviceSheet.getRange("A2:A").getValues();
  var lastRow = Avals_service.filter(String).length;

  for(i = serviceSheet_FIRSTROW; i < serviceSheet_FIRSTROW+lastRow; i++)
  {
    var tempCell = serviceSheet.getRange(i, serviceSheet_CONFIRM_SIGNUP).getValues();
    if(tempCell=="No, I want one to be provided")
    {
      serviceSheet.getRange(i,serviceSheet_SIGNUPLINK).setValue(link);
    }
  }
}

function createSignupForm(eventName, eventDate)
{
  var nameForm = 'Signup Form: ' + eventName +":"+eventDate
  var nameSheet = 'Signup Sheet: ' + eventName +":"+eventDate
  var folderId = '1HPQ0P2d81UrNwj2-IRmr7cN6ovmgVGLl'

  var form = FormApp.create(nameForm);
  form.setTitle(eventName + " "+ eventDate+" Sign Up Form")
    .setDescription('Please Respond to the Attendance Questions Below')
    .setConfirmationMessage('Thanks for responding!')
    
  var question1 = form.addTextItem();
  question1.setTitle('What is your name?');
  question1.setRequired(true);
  

  var question2 = form.addTextItem();
  question2.setTitle('What is your BU ID?');
  question2.setRequired(true);

  var question3 = form.addMultipleChoiceItem();
  question3.setTitle('Do you have an approved CORI Form through the CSC?');
  question3.setChoiceValues(['Yes','No']);
  question3.setRequired(true);


  var resource = {
  title: nameSheet,
  mimeType: MimeType.GOOGLE_SHEETS,
  parents: [{ id: folderId }]
  }
  var fileJson = Drive.Files.insert(resource)
  var fileId = fileJson.id

  form.setDestination(FormApp.DestinationType.SPREADSHEET, fileId);

  //move form to signup sheet folder
  file = DriveApp.getFileById(form.getId());
  folder = DriveApp.getFolderById(folderId);
  file.moveTo(folder);

  return form.getPublishedUrl();
}

function hash(string) {
  var hash = 0;
  if (string.length == 0) return hash;
  for (x = 0; x <string.length; x++) {
  ch = string.charCodeAt(x);
          hash = ((hash <<5) - hash) + ch;
          hash = hash & hash;
      }
  hash = Math.abs(hash);
  return hash.toString().substring(0,5);
}

function generateID()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var serviceSheet = ss.getSheetByName("Service DB");
  var Avals = serviceSheet.getRange("A1:A").getValues();
  var lastRow = Avals.filter(String).length; 

  Logger.log(lastRow);
  
  var name = serviceSheet.getRange(lastRow,serviceSheet_TITLE).getValue().replace(/\s/g, '').substring(0,3);
  var ind1 = serviceSheet.getRange(lastRow,serviceSheet_TIMESTAMP).getValue();

  
  var id = "ID-"+name+"-"+hash(ind1.toString())+"-"+lastRow;

  serviceSheet.getRange(lastRow,serviceSheet_SERVICEID).setValue(id);
}



