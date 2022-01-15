// This script is partly based on the Quickstart example in the following link.
// ref: https://developers.google.com/apps-script/quickstart/forms

// ref: https://stackoverflow.com/a/29574068/1167750
let signUpResponseSpreadsheetId = '';
let signUpResponseSpreadsheetTab = 'Events';
let eventLocationsSpreadsheetTab = 'Event Locations';
let formFieldLabelsSpreadsheetTab = 'Form Field Labels';
let formSettingsSpreadsheetTab = 'Form Settings';
let calendarTimeZone = '';
let signUpCalendarId = '';
let signUpFormTitle = 'Calendar Event Form';

function doGet(e) {
  // ref: https://developers.google.com/apps-script/guides/html
  // ref: https://stackoverflow.com/a/67042695/1167750
  return HtmlService.createHtmlOutputFromFile('index.html')
    .setTitle(signUpFormTitle)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getLocations() {
  // ref: https://stackoverflow.com/q/21813699/1167750
  let spreadsheet = SpreadsheetApp.openById(signUpResponseSpreadsheetId);
  let sheet = spreadsheet.getSheetByName(eventLocationsSpreadsheetTab);
  let range = sheet.getDataRange();
  let values = range.getValues();
  let locations = [];
  for (let i = 1; i < values.length; i++) {
    locations.push(values[i][0]);
  }
  return locations;
}

function getFormFieldLabels() {
  // ref: https://stackoverflow.com/q/21813699/1167750
  let spreadsheet = SpreadsheetApp.openById(signUpResponseSpreadsheetId);
  let sheet = spreadsheet.getSheetByName(formFieldLabelsSpreadsheetTab);
  let range = sheet.getDataRange();
  let values = range.getValues();
  let formFieldLabels = [];
  for (let i = 1; i < values.length; i++) {    
    formFieldLabels.push(values[i][0]);
  }
  return formFieldLabels;
}

function getFormSettings() {
  // ref: https://stackoverflow.com/q/21813699/1167750
  let spreadsheet = SpreadsheetApp.openById(signUpResponseSpreadsheetId);
  let sheet = spreadsheet.getSheetByName(formSettingsSpreadsheetTab);
  let range = sheet.getDataRange();
  let values = range.getValues();
  let formSettings = {};
  for (let i = 1; i < values.length; i++) {    
    formSettings[values[i][0]] = values[i][1];
  }
  return formSettings;
}

function processForm(data) {
  let requiredFieldsFilledResult = requiredFieldsFilled(data);
  if (!requiredFieldsFilledResult.result) {
    return requiredFieldsFilledResult;
  }

  let validationResult = checkForValidEventTimeRange(data);
  if (!validationResult.result) {
    return validationResult;
  }

  let overlappingEventResult = checkForOverlappingEvent(data);
  if (!overlappingEventResult.result) {
    saveEntryToSpreadsheet(data);
    saveEntryToCalendar();
    
    let eventStartTimeHourPMConversion = convertDatetimeHourToPM(data.eventStartDatetime);
    let eventEndTimeHourPMConversion = convertDatetimeHourToPM(data.eventEndDatetime);

    overlappingEventResult.message = '<b>Success: Your event was successfully added!</b>'
      + '<br><b>Your Event Details:</b>'
      + '<br>Event Name: '
      + data.eventName
      + '<br>Event From: '
      + data.eventStartDatetime
      + eventStartTimeHourPMConversion
      + '<br>Event To: '
      + data.eventEndDatetime
      + eventEndTimeHourPMConversion
      + '<br>Event Location: '
      + data.eventLocation
      + '<br>Event Description: '
      + data.eventDescription;
  }

  return overlappingEventResult;
}

function requiredFieldsFilled(data) {
  let result = {
    'message': '',
    'result': false
  }

  if ((data.eventName.length > 0)
      && (data.eventStartDatetime.length > 0)
      && (data.eventEndDatetime.length > 0)
      && (data.eventStartDatetime != 'NaN:NaN')
      && (data.eventEndDatetime != 'NaN:NaN')) {  
    result.messsage = '<b>Success:</b> Required fields appear to be filled in.';
    result.result = true;
    return result;
  }

  result.message = '<b>Error:</b> Please make sure to fill in all required fields! (These are the fields with the * next to them!)';
  return result;
}

function checkForValidEventTimeRange(data) {
  let result = {
    'message': '',
    'result': true
  };

  let eventStartDatetimeFromForm = new Date(data.eventStartDatetime);
  let eventEndDatetimeFromForm = new Date(data.eventEndDatetime);
  let eventStartDateFromForm =
    Utilities.formatDate(eventStartDatetimeFromForm, calendarTimeZone, 'yyyy-MM-dd');
  let eventEndDateFromForm =
    Utilities.formatDate(eventEndDatetimeFromForm, calendarTimeZone, 'yyyy-MM-dd');
  let eventStartTimeFromForm =
    Utilities.formatDate(eventStartDatetimeFromForm, calendarTimeZone, 'HH:mm');
  let eventEndTimeFromForm =
    Utilities.formatDate(eventEndDatetimeFromForm, calendarTimeZone, 'HH:mm');

  if ((eventStartDateFromForm == eventEndDateFromForm)
      && (eventStartTimeFromForm > eventEndTimeFromForm)) {
    result.message = '<b>Error: Sorry, the event start time and end time are invalid!</b>'
      + '<br />Please make sure that the <b>event end time</b> is later than the '
      + '<b>event start time</b> if an event happens on the same day.';
    result.result = false;
  }

   if ((eventStartDateFromForm == eventEndDateFromForm)
      && (eventStartTimeFromForm == eventEndTimeFromForm)) {
    result.message = '<b>Error: Sorry, the event start time and end time are invalid!</b>'
      + '<br />Please make sure that the <b>event end time</b> is later than the '
      + '<b>event start time</b> if an event happens on the same day.';
    result.result = false;
  }

  return result;
}

function checkForOverlappingEvent(data) {
  let result = {
    'message': '',
    'result': false
  };

  // ref: https://stackoverflow.com/q/21813699/1167750
  let spreadsheet = SpreadsheetApp.openById(signUpResponseSpreadsheetId);
  let sheet = spreadsheet.getSheetByName(signUpResponseSpreadsheetTab);
  let range = sheet.getDataRange();
  let values = range.getValues();

  for (let i = 1; i < values.length; i++) {
    // ref: https://www.delftstack.com/howto/javascript/convert-string-to-date-in-javascript/    
    // ref: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Date
    let eventNameFromSpreadsheet = values[i][1];
    let eventStartDatetimeFromSpreadsheet = new Date(values[i][2]);
    let eventEndDatetimeFromSpreadsheet = new Date(values[i][3]);
    let eventDescriptionFromSpreadsheet = values[i][5];

    let eventStartDatetimeFromForm = new Date(data.eventStartDatetime);
    let eventEndDatetimeFromForm = new Date(data.eventEndDatetime);
   
    let eventStartDateFromSpreadsheet =
      Utilities.formatDate(eventStartDatetimeFromSpreadsheet, calendarTimeZone, 'yyyy-MM-dd');
    let eventEndDateFromSpreadsheet =
      Utilities.formatDate(eventEndDatetimeFromSpreadsheet, calendarTimeZone, 'yyyy-MM-dd');      
    let eventStartTimeFromSpreadsheet =
      Utilities.formatDate(eventStartDatetimeFromSpreadsheet, calendarTimeZone, 'HH:mm');      
    let eventEndTimeFromSpreadsheet =
      Utilities.formatDate(eventEndDatetimeFromSpreadsheet, calendarTimeZone, 'HH:mm');
    let eventLocationFromSpreadsheet = values[i][4];

    // Note: Needed to update the timezone in the "Apps Script" `appsscript.json` manifest file
    // in order for the dates from the HTML form to be brought in correctly.
    // ref: https://javascript.tutorialink.com/google-apps-script-returns-wrong-timezone/
    let eventStartDateFromForm =
      Utilities.formatDate(eventStartDatetimeFromForm, calendarTimeZone, 'yyyy-MM-dd');
    let eventEndDateFromForm =
      Utilities.formatDate(eventEndDatetimeFromForm, calendarTimeZone, 'yyyy-MM-dd');
    let eventStartTimeFromForm =
      Utilities.formatDate(eventStartDatetimeFromForm, calendarTimeZone, 'HH:mm');
    let eventEndTimeFromForm =
      Utilities.formatDate(eventEndDatetimeFromForm, calendarTimeZone, 'HH:mm');
    let eventLocationFromForm = data.eventLocation;
  
    if ((eventLocationFromSpreadsheet === eventLocationFromForm)
        && (
          (
            (eventStartDateFromForm >= eventStartDateFromSpreadsheet)
            && (eventEndDateFromForm <= eventEndDateFromSpreadsheet)
            && (eventStartTimeFromForm >= eventStartTimeFromSpreadsheet)
            && (eventEndTimeFromForm <= eventEndTimeFromSpreadsheet)
          )
          ||
          (
            (eventStartDateFromForm <= eventStartDateFromSpreadsheet)
            && (eventEndDateFromForm > eventEndDateFromSpreadsheet)            
          )
          ||
          (
            (eventStartDateFromForm <= eventStartDateFromSpreadsheet)
            && (eventEndDateFromForm >= eventEndDateFromSpreadsheet)
            && (eventStartTimeFromForm <= eventStartTimeFromSpreadsheet)
            && (eventEndTimeFromForm >= eventEndTimeFromSpreadsheet)
          )
          ||
          (
            (eventStartDateFromForm <= eventStartDateFromSpreadsheet)
            && (eventEndDateFromForm >= eventEndDateFromSpreadsheet)
            && (eventStartTimeFromForm <= eventStartTimeFromSpreadsheet)
            && (eventEndTimeFromForm > eventStartTimeFromSpreadsheet)
          )
          ||
          (
            (eventStartDateFromForm <= eventStartDateFromSpreadsheet)
            && (eventEndDateFromForm >= eventEndDateFromSpreadsheet)
            && (eventStartTimeFromForm < eventEndTimeFromSpreadsheet)
            && (eventEndTimeFromForm > eventStartTimeFromSpreadsheet)
          )
        )) {
      // then there is an overlapping event!
      // go back to form, if possible
      // give message to user, if possible
      let eventStartTimeHourPMConversion = convertHourToPM(eventStartTimeFromSpreadsheet);
      let eventEndTimeHourPMConversion = convertHourToPM(eventEndTimeFromSpreadsheet);

      result.message = '<b>Error: Found an event location conflict!</b>'
        + '<br>Sorry, please choose another <b>time</b> or <b>location</b> for your event!'
        + '<br><b>Details about the Other Event in Conflict</b>:'
        + '<br>Event Name: '
        + eventNameFromSpreadsheet
        + '<br>Event From: '
        + eventStartDateFromSpreadsheet
        + ' '
        + eventStartTimeFromSpreadsheet
        + eventStartTimeHourPMConversion
        + '<br>Event To: '
        + eventEndDateFromSpreadsheet
        + ' '
        + eventEndTimeFromSpreadsheet
        + eventEndTimeHourPMConversion
        + '<br>Event Location: '
        + eventLocationFromSpreadsheet
        + '<br>Event Description: '
        + eventDescriptionFromSpreadsheet;
        
      result.result = true;
      return result;
    }
  }
  return result;
}

function saveEntryToSpreadsheet(data) {
  // ref: https://stackoverflow.com/q/21813699/1167750
  let spreadsheet = SpreadsheetApp.openById(signUpResponseSpreadsheetId);
  let sheet = spreadsheet.getSheetByName(signUpResponseSpreadsheetTab);
  let range = sheet.getDataRange();
  let values = range.getValues();
  // ref: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Date/now
  let now = new Date();
  // ref: https://developers.google.com/google-ads/scripts/docs/features/dates
  let dateTimeStamp = Utilities.formatDate(now, calendarTimeZone, 'MM/dd/yyyy HH:mm:ss');

  // Save data to spreadsheet first
  // ref: https://spreadsheet.dev/write-multiple-rows-to-google-sheets-using-apps-script
  spreadsheet.appendRow([
    dateTimeStamp,
    data.eventName,
    data.eventStartDatetime,
    data.eventEndDatetime,
    data.eventLocation,
    data.eventDescription
  ]);
}

function saveEntryToCalendar() {
  // ref: https://stackoverflow.com/q/21813699/1167750
  let spreadsheet = SpreadsheetApp.openById(signUpResponseSpreadsheetId);
  let sheet = spreadsheet.getSheetByName(signUpResponseSpreadsheetTab);
  let range = sheet.getDataRange();
  let values = range.getValues();

  // Then create the calendar event
  let cal = CalendarApp.getCalendarById(signUpCalendarId);
  let session = values[values.length-1];
  let title = session[1];  
  let start = session[2];
  let end = session[3];
  let options = {
    description: session[5],
    location: session[4],
    sendInvites: false,
  };
  let event = cal.createEvent(title, start, end, options)
    .setGuestsCanSeeGuests(false);
  session[5] = event.getId();
}

function convertDatetimeHourToPM(datetimeStringInput) {
  let datetimePieces = datetimeStringInput.split(' ');
  let timePieces = datetimePieces[1].split(':');
  let hour = Number.parseInt(timePieces[0]);

  let hourPMValue = 0;
  if (hour > 12) {
    hourPMValue = hour - 12;
  }
  
  if (hourPMValue > 0) {
    return ' (' + hourPMValue + ':' + timePieces[1] + ' PM)';
  }

  return '';
}

function convertHourToPM(timeStringInput) {  
  let timePieces = timeStringInput.split(':');
  let hour = Number.parseInt(timePieces[0]);

  let hourPMValue = 0;
  if (hour > 12) {
    hourPMValue = hour - 12;
  }
  
  if (hourPMValue > 0) {
    return ' (' + hourPMValue + ':' + timePieces[1] + ' PM)';
  }

  return '';
}
