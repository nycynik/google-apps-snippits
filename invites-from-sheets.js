"use strict";

// -------------
// GUI Elements
// -------------

// triggered when sheet opens to add UI elements.
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('BTS Automation')
      .addItem('Create Calendar Invites', 'MakeEvents')
      //.addSeparator()
      .addToUi();
}

// UI events

function MakeEvents() {
  
  createCalendarEvent();

  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('Calendar events created!');
}

function SendInvites() {
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('Invites sent!');
}

// -------------

// add two cells together (one date, and one time) and form
// a javascript convertable date string.
function dateTimeToDate(date, time) {
  let d = new Date(date)
  let t = new Date(time)

  let hours = ('0' + (t.getHours())).slice(-2)
  let minutes = ('0' + (t.getMinutes())).slice(-2)
  //let ampm = d.getHours() >= 12 ? 'PM' : 'AM';
  let timeString = hours + ':' + minutes + ':00';

  let month = ('0' + (d.getMonth() + 1)).slice(-2)
  let day = ('0' + (d.getDate())).slice(-2)
  let dateString = '' + d.getFullYear() + '-' + month + '-' + day;

  var dateTimeString = dateString + 'T' + timeString;
  var output = new Date (dateTimeString);

  if(!isNaN(output.valueOf())) {
    return output;
  } else {
    throw new Error('Not a valid date-time');
  }   
}

function createCalendarEvent() {

  let sheet = SpreadsheetApp.getActiveSheet();
  let calendar = CalendarApp.getCalendarById('### CAL ID ###');

  let year = sheet.getRange(3,2).getValue();
  let guests = sheet.getRange(4,2).getValue(); 
  if (year == '' || guests == '') {
    throw new Error( "Year and Guests are manditory fields, please fill them in (B3 and E3 values are empty!)." );
  }

  let startRow = 8;  // First row of data to process - 2 exempts my header row
  let numRows = sheet.getLastRow();   // Number of rows to process
  let numColumns = sheet.getLastColumn();
 
  let dataRange = sheet.getRange(startRow, 1, numRows-1, numColumns);
  let data = dataRange.getValues();
 
  let complete = "Event Created";
 
  for (let i = 0; i < data.length; ++i) {

    let row = data[i];
    let eventDescription = row[2]; 
    let date = new Date(row[1]);  
    let wholeDayEvent = new Boolean(row[4])
    let startDate = dateTimeToDate(row[1], row[5]);  
    let endDate = dateTimeToDate(row[1], row[6]);  
    let eventStatus = row[0]; // Status
   
    // parse data
    let lines = eventDescription.split('\n')
    let lineCount = lines.length
    let title = lines[0]
    let body = ""
    for (let idx=2; idx < lineCount; idx++) {
      body += lines[idx] + '\r';
    }

    if (title != "" && eventStatus != complete) {

      let event;

      if (wholeDayEvent == false) {
        event = calendar.createEvent(title, startDate, endDate, {
          description: body + '\r' + date,
          guests: guests
        });
        Logger.log('Event ID: ' + event.getId() + ' Title: ' + title + ' start:' + startDate);
       
      } else {
        event = calendar.createAllDayEvent(title, date, {
          description: body + '\r' + date,
          guests: guests
        });
        Logger.log('Event ID: ' + event.getId() + ' Title: ' + title + ' On:' + date);
      }
      
      // show status on sheet
      let currentCell = sheet.getRange(startRow + i, 1);
      currentCell.setValue(complete);
    }
  }
}
