/**
 * Sends an email based on a form response
 * Author: Nick Hotto
 * March 2014
 */
function sendEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Change to the "form responses" spreadsheet and sort Z to A
  var sheet = ss.getSheetByName("Requests");
  sheet.sort(4, false);
  // Get info needed for generated emails
  var emailRange = sheet.getRange("E2");
  var email = emailRange.getValue();
  var deptRange = sheet.getRange("F2");
  var dept = deptRange.getValue();
  var splitName = email.split(".");
  var firstName = splitName[0];
  var firstInitial = firstName.charAt(0).toUpperCase();
  var lastName = splitName[1];
  var lastInitial = lastName.charAt(0).toUpperCase();
  var projectNameRange = sheet.getRange("G2");
  var projectName = projectNameRange.getValue().toUpperCase();
  var dateRange = sheet.getRange("D2");
  var date = dateRange.getValue();
  var month = (date.getMonth()+1);
  var day = date.getDate();
  var hour = date.getHours();
  var minutes = date.getMinutes();
  var seconds = date.getSeconds();
  var jobNameCell = sheet.getRange("C2");
  jobNameCell.setValue(firstInitial +
                   lastInitial + '-' +
                   projectName + '-' +
                   month + '/' +
                   day + '-' +
                   hour +
                   minutes)
  if (dept == 'Process Engineering') {
    MailApp.sendEmail('diana.pincus@makerbot.com',
                      'New Machine Shop Request',
                      'New request from Process Engineering Department. ' +
                      'Link: http://goo.gl/rg8FaK' + '\r\n' + 'Job name: ' +
                      firstInitial +
                      lastInitial + '-' +
                      projectName + '-' +
                      month + '/' +
                      day + '-' +
                      hour +
                      minutes, {
      name: 'Machine Shop Request Notification',
      cc: 'chris.yarka@makerbot.com',
      bcc: 'nicholas.hotto@makerbot.com',
      noReply: true,
      replyTo: 'phil.koeing@makerbot.com'
    });
  } else if (dept == 'Mechanical Engineering') {
    MailApp.sendEmail('bill.buel@makerbot.com',
                      'New Machine Shop Request',
                      'New request from Mechanical Engineering Department. ' +
                      'Link: http://goo.gl/rg8FaK' + '\r\n' + 'Job name: ' +
                      firstInitial +
                      lastInitial + '-' +
                      projectName + '-' +
                      month + '/' +
                      day + '-' +
                      hour + ':' +
                      minutes + ':' +
                      seconds, {
      name: 'Machine Shop Request Notification',
      cc: 'chris.yarka@makerbot.com',
      noReply: true,
      replyTo: 'phil.koeing@makerbot.com'
    });
  } else if (dept == 'R&D') {
    MailApp.sendEmail('kate.nevin@makerbot.com',
                      'New Machine Shop Request',
                      'New request from R&D Department. ' +
                      'Link: http://goo.gl/rg8FaK' + '\r\n' + 'Job name: ' +
                      firstInitial +
                      lastInitial + '-' +
                      projectName + '-' +
                      month + '/' +
                      day + '-' +
                      hour + ':' +
                      minutes + ':' +
                      seconds, {
      name: 'Machine Shop Request Notification',
      cc: 'chris.yarka@makerbot.com',
      noReply: true,
      replyTo: 'phil.koeing@makerbot.com'
    });
  } else {
    MailApp.sendEmail('chris.yarka@makerbot.com',
                      'New Machine Shop Request',
                      'New request from Other Department. ' +
                      'Link: http://goo.gl/rg8FaK' + '\r\n' + 'Job name: ' +
                      firstInitial +
                      lastInitial + '-' +
                      projectName + '-' +
                      month + '/' +
                      day + '-' +
                      hour + ':' +
                      minutes + ':' +
                      seconds, {
      name: 'Machine Shop Request Notification',
      noReply: true,
      replyTo: 'phil.koeing@makerbot.com'
    });
  }
}
