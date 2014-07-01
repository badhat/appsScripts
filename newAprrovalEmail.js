/**
 * Sends an email based on a spreadsheet approval
 * Author: Nick Hotto
 * March 2014
 */
function sendApprovalNotification(event) {
  //  Set range length and create arrays
  "use strict";
  var rangeLength = 21, approvalCells = [], j = 2, i = 0, inList = false, approved = false, sheet = event.source, cell = event.range.getA1Notation(), sheetURL = 'http://goo.gl/LwR8jE', splitApproval;
  for (j = 2; j < rangeLength; j = j + 1) {
        "use strict";
        approvalCells.push("P" + j);
  }
  for (i = 0; i < approvalCells.length; i = i + 1) {
    if (cell === approvalCells[i]) {inList = true; }
    splitApproval = sheet.getRange(approvalCells[i]).getValue().split(" ")
    if (splitApproval[0] === 'Approved') {approved = true; }
    if (inList && approved) {
      MailApp.sendEmail('anthony.kim@makerbot.com',
                        'Supply Chain Request Approved',
                        'New Approval ' +
                        'Link: ' + sheetURL, {
                        name: 'Supply Chain Request Notification',
                        cc: 'nicholas.hotto@makerbot.com',
                        noReply: true
      });
    }
  }
}
