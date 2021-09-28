/**
 *  Create custom menu button to set time off
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Schedule Event')
    .addItem('Set Time Off', 'scheduleEvent')  // create set time off button under schedule event category
    .addToUi();
}

/**
 *   Notify approver and requester if someone submitted a request
 */
function notify() {

  Utilities.sleep(1000)  // sleep to ensure that the last row gets populated first before trying to grab it

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheets()[0];
  var requester = spreadsheet.getDataRange().getValues();
  var timeSubmitted = requester[sheet.getLastRow() - 1][0];
  var email = requester[sheet.getLastRow() - 1][1];
  var startDate = requester[sheet.getLastRow() - 1][2];
  var endDate = requester[sheet.getLastRow() - 1][3]
  var exempt = requester[sheet.getLastRow() - 1][4];
  var type = requester[sheet.getLastRow() - 1][5];
  var cover = requester[sheet.getLastRow() - 1][6];
  var details = requester[sheet.getLastRow() - 1][7];

  //  send confirmation email
  MailApp.sendEmail({
    to: email,
    subject: "Your leave request was submitted",
    htmlBody: "<HTML><BODY><P>" + "Your leave request was submitted at: " + timeSubmitted + ". Please look out for an email with an approval decision. <br><br>startDate: " + startDate + " <br>endDate: " + endDate + "<br>email: " + email + "<br>exempt: " + exempt + "<br>type: " + type + "<br>cover: " + cover + "<br>details: " + details + "<br><br>Thanks!" + "</P></BODY></HTML>",
  });

  // notify approver a new leave request was submitted
  MailApp.sendEmail({
    to: "[Enter email(s) to be notified in a comma separated list]",
    subject: "New Leave Request",
    htmlBody: "<HTML><BODY><P>" + "A new leave request was submitted at: " + timeSubmitted + ".<br><br>startDate: " + startDate + " <br>endDate: " + endDate + "<br>email: " + email + "<br>exempt: " + exempt + "<br>type: " + type + "<br>cover: " + cover + "<br>details: " + details + "<br><br>Please approve or deny their request with yes or no in the approved column.<br> [GOOGLE SHEET LINK] <br><br>Thanks!" + "</P></BODY></HTML>",
  });

}

/**
 *  Schedule Event for time off depending on approval status
 */
function scheduleEvent() {

  //  access active spreadsheet and calendar to schedule events to
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var eventCal = CalendarApp.getCalendarById("[CALENDAR ID]");

  let sourceSheet = spreadsheet.getSheets()[0];  //  take requests from pending
  let targetSheet = spreadsheet.getSheets()[1];  //  move processed requests to archive tab

  var cells = spreadsheet.getDataRange();  //  get full range of requests from the sheet
  var requests = cells.getValues();

  //  cycle through all requests in sheet
  // bottom-up parsing to account for changing index due to row deletions
  for (x = requests.length - 1; x >= 0; x--) {

    //  organize data to respective fields (make sure column numbers match form questions)
    var timeSubmitted = requests[x][0];
    var email = requests[x][1];
    var startDate = requests[x][2];
    var endDate = requests[x][3]
    var exempt = requests[x][4];
    var type = requests[x][5];
    var cover = requests[x][6];
    var details = requests[x][7];
    var approved = requests[x][8];

    if (approved.toLowerCase() == 'yes') {  // if approved, schedule calendar event

      //  move request to archive
      targetSheet.appendRow([timeSubmitted, email, startDate, endDate, exempt, type, cover, details, approved]);

      //  add event to calendar and let requester know through a calendar invite
      if (startDate.toString() == endDate.toString()) {
        event = eventCal.createAllDayEvent(email + ' Out of Office', startDate, {
          description: "email: " + email + "\nexempt: " + exempt + "\ntype: " + type + "\ncover: " + cover + "\ndetails: " + details,
          guests: email,
          sendInvites: true,
        });
      } else {
        event = eventCal.createAllDayEvent(email + ' Out of Office', startDate, endDate, {
          description: "email: " + email + "\nexempt: " + exempt + "\ntype: " + type + "\ncover: " + cover + "\ndetails: " + details,
          guests: email,
          sendInvites: true,
        });
      }

      //  notify requester leave request was approved through email
      MailApp.sendEmail({
        to: email,
        subject: "Approved Leave Request",
        htmlBody: "<HTML><BODY><P>" + "Your leave request for " + startDate + " to " + endDate + " was approved. Please contact your manager if you have any questions. <br><br>Thanks!" + "</P></BODY></HTML>",
      });

      sourceSheet.deleteRow(x + 2);  //  delete request off of pending sheet

    } else if (approved.toLowerCase() == 'no') {
      targetSheet.appendRow([timeSubmitted, email, startDate, endDate, exempt, type, cover, details, approved]);  //  move request to archive
      sourceSheet.deleteRow(x + 2);  //  delete request off of pending sheet

      //  notify requester leave request was denied
      MailApp.sendEmail({
        to: email,
        subject: "Denied Leave Request",
        htmlBody: "<HTML><BODY><P>" + "Your leave request was not approved. Please contact your manager if you have any questions. <br><br>Your leave request was submitted at: " + timeSubmitted + ".<br>startDate: " + startDate + " <br>endDate: " + endDate + "<br>email: " + email + "<br>exempt: " + exempt + "<br>type: " + type + "<br>cover: " + cover + "<br>details: " + details + "</P></BODY></HTML>",
      });
    }

  //  blank approved rows are left alone
  }


}
