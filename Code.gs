//Copyright 2013 Google Inc. All Rights Reserved.

/*
   DISCLAIMER:

   (i) GOOGLE INC. ("GOOGLE") PROVIDES YOU ALL CODE HEREIN "AS IS" WITHOUT ANY
   WARRANTIES OF ANY KIND, EXPRESS, IMPLIED, STATUTORY OR OTHERWISE, INCLUDING,
   WITHOUT LIMITATION, ANY IMPLIED WARRANTY OF MERCHANTABILITY, FITNESS FOR A
   PARTICULAR PURPOSE AND NON-INFRINGEMENT; AND

   (ii) IN NO EVENT WILL GOOGLE BE LIABLE FOR ANY LOST REVENUES, PROFIT OR DATA,
   OR ANY DIRECT, INDIRECT, SPECIAL, CONSEQUENTIAL, INCIDENTAL OR PUNITIVE
   DAMAGES, HOWEVER CAUSED AND REGARDLESS OF THE THEORY OF LIABILITY, EVEN IF
   GOOGLE HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES, ARISING OUT OF
   THE USE OR INABILITY TO USE, MODIFICATION OR DISTRIBUTION OF THIS CODE OR
   ITS DERIVATIVES.
*/

// setRowsData fills in one row of data per object defined in the objects Array.
// For every Column, it checks if data objects define a value for it.
// Arguments:
//   - sheet: the Sheet Object where the data will be written
//   - objects: an Array of Objects, each of which contains data for a row
//   - optHeadersRange: a Range of cells where the column headers are defined. This
//     defaults to the entire first row in sheet.
//   - optFirstDataRowIndex: index of the first row where data should be written. This
//     defaults to the row immediately below the headers.

function setRowData(sheet, object, optHeadersRange) {
  Logger.log("Writing at row index "+object.__rowIndex_);
  Logger.log(object);
  setRowsData(sheet, [object], optHeadersRange, object.__row_);
}

function setRowsData(sheet, objects, optHeadersRange, optFirstDataRowIndex) {
  var headersRange = optHeadersRange || sheet.getRange(1, 1, 1, sheet.getMaxColumns());
  var firstDataRowIndex = optFirstDataRowIndex || headersRange.getRowIndex() + 1;
  Logger.log("fdri="+firstDataRowIndex);
  var headers = normalizeHeaders(headersRange.getValues()[0]);

  var data = [];
  for (var i = 0; i < objects.length; ++i) {
    var values = []
    for (j = 0; j < headers.length; ++j) {
      var header = headers[j];
      // If the header is non-empty and the object value is 0...
      if ((header.length > 0) && (objects[i][header] == 0)) {
        values.push(0);
      }
      // If the header is empty or the object value is empty...
      else if ((!(header.length > 0)) || (objects[i][header]=='')) {
        values.push('');
      }
      else {
        values.push(objects[i][header]);
      }
    }
    data.push(values);
  }
  
  var destinationRange = sheet.getRange(firstDataRowIndex, headersRange.getColumnIndex(),
                                        objects.length, headers.length);
  destinationRange.setValues(data);
}

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//       This argument is optional and it defaults to all the cells except those in the first row
//       or all the cells below columnHeadersRowIndex (if defined).
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  
  Logger.log("getRowsData");
  var headersIndex = columnHeadersRowIndex || range ? range.getRowIndex() - 1 : 1;
  var dataRange = range ||
    sheet.getRange(headersIndex + 1, 1, sheet.getMaxRows() - headersIndex, sheet.getMaxColumns());
  var numColumns = dataRange.getLastColumn() - dataRange.getColumn() + 1;
  var headersRange = sheet.getRange(headersIndex, dataRange.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(dataRange.getValues(), normalizeHeaders(headers));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    // Increment 1 for the header row, and one more so that it matches the spreadsheet.
    object.__row_ = 2+i;
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Empty Strings are returned for all Strings that could not be successfully normalized.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {  
    keys.push(normalizeHeader(headers[i]));   
    Logger.log("string: "+headers[i]);
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  
  //Logger.log("header: "+key);
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}


// Configurable Values

var SPREADSHEET_KEY = "1a_joSO5eFezVsrxNDEb7HT1EqbvdV8EB6Z2zpu4914s";

var RESPONSES_SHEET = "Form Responses 1";
var LOG_SHEET = "Execution Log";
var SETTINGS_SHEET = "__Settings";
var CACHE_SETTINGS = false;
var SETTINGS_CACHE_TTL = 900;

// Constants
var BLANK_STATE = undefined;
var PENDING_STATE = "PENDING";
//var APPROVED_STATE = "APPROVED";
var COVERAGE_APPROVED_STATE = "COVERAGE APPROVED";
var MANAGER_APPROVED_STATE = "MANAGER APPROVED";
var DENIED_STATE = "DENIED";
var CANCELLED_STATE = "CANCELLED";
var EMAIL_REGEX = new RegExp("[a-zA-Z+0-9\._-]+@[a-zA-Z0-9\._-]+", 'i'); 

// Globals
var cache = JSONCacheService();
var SETTINGS = getSettings();
var eventId;

function JSONCacheService() {
  var _cache = CacheService.getPublicCache();
  var _key_prefix = "_json#";
  
  var get = function(k) {
    var payload = _cache.get(_key_prefix+k);
    if(payload !== undefined) {
      JSON.parse(payload);
    }
    return payload
  }
  
  var put = function(k, d, t) {
    _cache.put(_key_prefix+k, JSON.stringify(d), t);
  }
  
  return {
    'get': get,
    'put': put
  }
}

function getSettings() { 
  if(CACHE_SETTINGS) {
    var settings = cache.get("_settings");
  }
  
  if(settings == undefined) {
    var sheet = getSpreadsheet().getSheetByName(SETTINGS_SHEET);
    var values = sheet.getDataRange().getValues();
  
    var settings = {};
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      settings[row[0]] = row[1];
    }
    
    cache.put("_settings", settings, SETTINGS_CACHE_TTL);
  }
  return settings;
}

//generates a global unique identifier i think -rp
function Utils() {}
Utils.generateUUID = function() {
  var d = new Date().getTime();
  var uuid = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    var r = (d + Math.random()*16)%16 | 0;
    d = Math.floor(d/16);
    return (c=='x' ? r : (r&0x7|0x8)).toString(16);
  });
  return uuid;
}

 // i feel like this function might be important but im not sure -rp 
Utils.processTemplate = function(template, object) {
  // Do a cheap string replace on several template values.
  if(template == undefined) {
    return "NOT DEFINED";
  }
  for(var k in object) {
     var objectForPrint = object[k];
     if (object[k] instanceof Date) {  
      objectForPrint = object[k].toDateString();
      } else {  
      objectForPrint = object[k];
    }
    template = template.replace(SETTINGS.TEMPLATE_OPEN_TAG + k + SETTINGS.TEMPLATE_CLOSE_TAG, objectForPrint);
    
  }
  return template
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_KEY);
}

function SheetHandler(sheet) {
  var _sheet = sheet;
  var _data = getRowsData(_sheet);
  
  var _markPending = function(d) {
    
    d.state = PENDING_STATE;
    d.identifier = Utils.generateUUID();
    
    //Edits to change Two columns holding date object but in STANDARD TIME NO MORE MILITARY TIME
    d.standardStartTime = Utilities.formatDate(d.leaveStartDate, "PST", "EEE, MMM d, yyyy hh:mm a");
    d.standardEndTime = Utilities.formatDate(d.lastDayOfLeave, "PST", "EEE, MMM d, yyyy hh:mm a");
    
    coverage_email = d[SETTINGS.COVERAGE_EMAIL_COLUMN_NAME].match(EMAIL_REGEX);
    
    var scriptUri = ScriptApp.getService().getUrl();
    // hack some values on to the data just for email templates.

    d.accept_url = scriptUri + "?i=" + d.identifier + '&state=' + COVERAGE_APPROVED_STATE;

    d.reject_url = scriptUri + "?i=" + d.identifier + '&state=' + DENIED_STATE;
    d.rejectUrl = d.reject_url;
    d.acceptUrl = d.accept_url;
    
    d.approvalButton = 'Awaiting Coverage Response';
    d.denialButton = 'Awaiting Coverage Response';

    
    d.coverage_email = coverage_email

//sends pending email to coverage
    message = Utils.processTemplate(SETTINGS.PENDING_COVERAGE_EMAIL, d);
    subject = Utils.processTemplate(SETTINGS.PENDING_COVERAGE_EMAIL_SUBJECT, d);
    
    MailApp.sendEmail(coverage_email,subject,"",{ htmlBody: message });
    

    //create calendar event
    Logger.log(SETTINGS);
    if (SETTINGS.CREATE_CALENDAR_EVENT == 1) {
      Logger.log("Creating Calendar Event");
      _createCalendarEventForDataRow(d);
    }
    

    setRowData(_sheet, d);  //changes sheet
  }
  
  var _getDataByKey = function(k) {
    row = undefined;
    _data.forEach(function(d) {
      if (d.identifier == k) {
        row = d;
      }
    });
    return row;
  }
  
  var processSheet = function() {
    _data.forEach(function(d) {
      if(d.state == BLANK_STATE) {
        _markPending(d);
      }
    });
  }
  
  var _createCalendarEventForDataRow = function(d) {
    // default to the user's calendar, unless we figure out other wise.
    var calendarId = d.emailAddress;
    var guests = new Array();
    
    if (SETTINGS.WRITE_TO_GROUP_CALENDAR == 1) {
      // The group calendar owns the event, invite the user and the coverage person
      var calendarId = SETTINGS.GROUP_CALENDAR_ID;
      guests.push(d.emailAddress);
      guests.push(d[SETTINGS.COVERAGE_EMAIL_COLUMN_NAME]);
    }
    
    if (SETTINGS.INVITE_MANAGER_TO_EVENT == 1) {
      // Regardless of where the event is, invite the manager if requested.
      guests.push(d.actor);
    }
    
    var title = Utils.processTemplate(SETTINGS.CALENDAR_EVENT_TITLE, d);
    var description = Utils.processTemplate(SETTINGS.CALENDAR_EVENT_DESCRIPTION, d);
    var location = Utils.processTemplate(SETTINGS.CALENDAR_EVENT_LOCATION, d);
    
    // Create an event.
    var calendar = CalendarApp.getCalendarById(calendarId);
    Logger.log(calendarId);
    Logger.log(calendar.getName());
    var event = calendar.createEvent(title,new Date(d.leaveStartDate),new Date(d.lastDayOfLeave),{
      description: description,
      sendInvites: false,
      location: location,
      guests: guests.join()
    });
    Logger.log(event.getId());
    d.eventId = event.getId(); //added this line to store eventId to the google spreadsheet

    
    //set event color to yellow for pending event
    event.setColor(CalendarApp.EventColor.YELLOW);


  }
  

  //trying to make code for deleting calendar event -rp
  var _deleteCalendarEvent = function(d) {
    
    var calendarId = SETTINGS.GROUP_CALENDAR_ID;
    var calendar = CalendarApp.getCalendarById(calendarId);
    var event = calendar.getEventById(d.eventId);
    event.deleteEvent();
    
  }
  

  //function to change the event color
  var _modifyApprovedCalendarEvent = function(d) {
    var calendarId = SETTINGS.GROUP_CALENDAR_ID;
    var calendar = CalendarApp.getCalendarById(calendarId);
    var event = calendar.getEventById(d.eventId);
    event.setColor(CalendarApp.EventColor.BLUE);
  }
  
  
  //sends the pending emails to manager once coverage approves
  var approveByCoverageKey = function(k, user) {
    var d = _getDataByKey(k);
    d.state = COVERAGE_APPROVED_STATE;
    d.actor = user;
    
    //to check if the standard time columns needs to be re-initialized
    d.standardStartTime = Utilities.formatDate(d.leaveStartDate, "PST", "EEE, MMM d, yyyy hh:mm a");
    d.standardEndTime = Utilities.formatDate(d.lastDayOfLeave, "PST", "EEE, MMM d, yyyy hh:mm a");
    
    manager_email = d[SETTINGS.MANAGERS_EMAIL_COLUMN_NAME].match(EMAIL_REGEX);
    d.manager_email = manager_email;
    
    //added these two lines to get the cancel button to appear as a link!!!!
    var scriptUri = ScriptApp.getService().getUrl();
    d.approval_url = scriptUri + "?i=" + d.identifier + '&state=' + MANAGER_APPROVED_STATE; 
    d.deny_url = scriptUri + "?i=" + d.identifier + '&state=' + DENIED_STATE;
    
    d.approvalUrl = d.approval_url;
    d.denyUrl = d.deny_url;
  
    d.approvalButton = '=HYPERLINK(INDIRECT("R[0]C[-1]", FALSE),"Approve")';
    d.denialButton = '=HYPERLINK(INDIRECT("R[0]C[-1]", FALSE),"Deny")';
    
    //send pending manager email
    message = Utils.processTemplate(SETTINGS.PENDING_MANAGER_EMAIL, d);
    subject = Utils.processTemplate(SETTINGS.PENDING_MANAGER_EMAIL_SUBJECT, d); 
    
    MailApp.sendEmail(manager_email,subject,"",{ htmlBody: message });
    
    //send coverage approval email
    message = Utils.processTemplate(SETTINGS.COVERAGE_APPROVAL_EMAIL, d);
    subject = Utils.processTemplate(SETTINGS.COVERAGE_APPROVAL_EMAIL_SUBJECT, d); 
    
    MailApp.sendEmail(d.emailAddress,subject,"",{ htmlBody: message });
       
    setRowData(_sheet, d);
  }
  //end request approved email
  
  
  //sends the approved emails to employee once manager also confirms 
  var approveByManagerKey = function(k, user) {
    var d = _getDataByKey(k);
    d.state = MANAGER_APPROVED_STATE;
    d.actor = user;
    

    //to check if the standard time columns needs to be re-initialized
    d.standardStartTime = Utilities.formatDate(d.leaveStartDate, "PST", "EEE, MMM d, yyyy hh:mm a");
    d.standardEndTime = Utilities.formatDate(d.lastDayOfLeave, "PST", "EEE, MMM d, yyyy hh:mm a");
    
    //added these two lines to get the cancel button to appear as a link!!!!
    var scriptUri = ScriptApp.getService().getUrl();
    d.cancel_url = scriptUri + "?i=" + d.identifier + '&state=' + CANCELLED_STATE;
    
    d.approvalButton = 'Already Approved';
    d.denialButton = 'Already Approved';
    
    message = Utils.processTemplate(SETTINGS.MANAGER_APPROVAL_EMAIL, d);
    subject = Utils.processTemplate(SETTINGS.MANAGER_APPROVAL_EMAIL_SUBJECT, d);
    
    
    MailApp.sendEmail(d.emailAddress,subject,"",{ htmlBody: message });
    
    //send notification email to coverage person that manager has approved
    if(SETTINGS.SEND_APPROVAL_NOTICE_EMAIL == 1) {
    
      message = Utils.processTemplate(SETTINGS.APPROVAL_NOTICE_EMAIL, d);
      subject = Utils.processTemplate(SETTINGS.APPROVAL_NOTICE_EMAIL_SUBJECT, d);
      MailApp.sendEmail(d[SETTINGS.COVERAGE_EMAIL_COLUMN_NAME].match(EMAIL_REGEX), subject, "",{ htmlBody: message });
    }
    
    //change event color to blue
      _modifyApprovedCalendarEvent(d);

    
    setRowData(_sheet, d);
  }
  

  
  //sends the denial emails
  var denyByKey = function(k, user) {
    var d = _getDataByKey(k);
    d.state = DENIED_STATE;
    d.actor = user;
    //d.actor = d[SETTINGS.MANAGERS_EMAIL_COLUMN_NAME].match(EMAIL_REGEX);
    
    //to check if the standard time columns needs to be re-initialized
    d.standardStartTime = Utilities.formatDate(d.leaveStartDate, "PST", "EEE, MMM d, yyyy hh:mm a");
    d.standardEndTime = Utilities.formatDate(d.lastDayOfLeave, "PST", "EEE, MMM d, yyyy hh:mm a");
    

    d.approvalButton = 'Already Denied';
    d.denialButton = 'Already Denied';
    
    message = Utils.processTemplate(SETTINGS.USER_DENIED_EMAIL, d);
    subject = Utils.processTemplate(SETTINGS.USER_DENIED_EMAIL_SUBJECT, d);
    MailApp.sendEmail(d.emailAddress, subject, "",{ htmlBody: message });
    
     //delete calendar event
    _deleteCalendarEvent(d); 
    
    setRowData(_sheet, d);
  }
  
 
  
  //cancellation emails
  var cancelByKey = function(k, user) {
    var d = _getDataByKey(k);
    d.state = CANCELLED_STATE;
    d.actor = d[SETTINGS.MANAGERS_EMAIL_COLUMN_NAME].match(EMAIL_REGEX);
    
    d.standardStartTime = Utilities.formatDate(d.leaveStartDate, "PST", "EEE, MMM d, yyyy hh:mm a");
    d.standardEndTime = Utilities.formatDate(d.lastDayOfLeave, "PST", "EEE, MMM d, yyyy hh:mm a");
       
    d.approvalButton = 'Cancelled';
    d.denialButton = 'Cancelled';
    
    //send email to manager
    message = Utils.processTemplate(SETTINGS.USER_CANCELLED_EMAIL, d);
    subject = Utils.processTemplate(SETTINGS.USER_CANCELLED_EMAIL_SUBJECT, d);
    MailApp.sendEmail(d[SETTINGS.MANAGERS_EMAIL_COLUMN_NAME].match(EMAIL_REGEX), subject, "",{ htmlBody: message });
    
    //send email to coverage
    message = Utils.processTemplate(SETTINGS.USER_CANCELLED_EMAIL, d);
    subject = Utils.processTemplate(SETTINGS.USER_CANCELLED_EMAIL_SUBJECT, d);
    MailApp.sendEmail(d[SETTINGS.COVERAGE_EMAIL_COLUMN_NAME].match(EMAIL_REGEX), subject, "",{ htmlBody: message });
    
    //delete calendar event
    _deleteCalendarEvent(d); 

    setRowData(_sheet, d);
  }

  return {
    'processSheet': processSheet,
    //'approveByKey': approveByKey,
    'approveByCoverageKey': approveByCoverageKey,
    'approveByManagerKey': approveByManagerKey,
    'denyByKey': denyByKey,
    //'denyCoverageByKey': denyByCoverageKey,
    'cancelByKey': cancelByKey //added for cancel
  }
};

function authorize() {
  var mail = MailApp.getRemainingDailyQuota();
  var calendar = CalendarApp.getAllCalendars();
}

function mockGet() {
  var request = {
    'parameters': {
      'i': '39f383a1-6841-428e-9be5-c35a85243513',
      'state': APPROVED_STATE
    }
  };
  doGet(request);
}

function doGet(request) {
  var sheet = getSpreadsheet().getSheetByName(RESPONSES_SHEET);
  var handler = SheetHandler(sheet);
  
  var user = Session.getActiveUser().getEmail();
  
/*  if(request.parameters.state == APPROVED_STATE) {
    handler.approveByKey(request.parameters.i, user);
  }
  
  if(request.parameters.state == DENIED_STATE) {
    handler.denyByKey(request.parameters.i, user);
  } 
*/
  
  if(request.parameters.state == COVERAGE_APPROVED_STATE) {
    handler.approveByCoverageKey(request.parameters.i, user);    
    return HtmlService.createHtmlOutput(SETTINGS.COVERAGE_APPROVAL_PAGE_TEMPLATE);
  }
  
  if(request.parameters.state == MANAGER_APPROVED_STATE) {
    handler.approveByManagerKey(request.parameters.i, user);  
    return HtmlService.createHtmlOutput(SETTINGS.MANAGER_APPROVAL_PAGE_TEMPLATE);
  }
  
  if(request.parameters.state == DENIED_STATE) {
    handler.denyByKey(request.parameters.i, user);
    return HtmlService.createHtmlOutput(SETTINGS.DENIAL_PAGE_TEMPLATE);
  }

  if(request.parameters.state == CANCELLED_STATE) {  //added for cancel
    handler.cancelByKey(request.parameters.i, user);
    return HtmlService.createHtmlOutput(SETTINGS.CANCELLATION_PAGE_TEMPLATE);
  }
}

function onFormSubmit() {
  Logger.log("Good afternoon, Good Evening and Good night");
  var sheet = getSpreadsheet().getSheetByName(RESPONSES_SHEET);
  handler = SheetHandler(sheet);
  handler.processSheet();
  writeToLog();
  
};


function writeToLog(){
  
  // write the Drive file link to the Drive File Report Tab for safe keeping/logging purposes
  var reportSheet = getSpreadsheet().getSheetByName(LOG_SHEET);
  Logger.log("passed:" +reportSheet.getName());
  
  var dateForLogging = Utilities.formatDate(new Date(), "PST", "EEE, MMM d, yyyy hh:mm:ss a z");
  reportSheet.appendRow([dateForLogging,Session.getActiveUser().getEmail()]);
}
