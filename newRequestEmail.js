/**
 * Sends an email based on a form response
 * Author: Nick Hotto
 * March 2014
 */
function sendEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Change to the "form responses" spreadsheet and sort Z to A
  var sheet = ss.getSheetByName("Requests");
  sheet.sort(1, false);
  // Get info needed for generated emails
  var emailRange = sheet.getRange("B2");
  var email = emailRange.getValue();
  var deptRange = sheet.getRange("C2");
  var dept = deptRange.getValue();
  var sheetName = 'Supply Chain'
  var sheetURL = 'http://goo.gl/LwR8jE'
  if (dept == 'Mfg. Engineering') {
    MailApp.sendEmail('diana.pincus@makerbot.com',
                      'New ' + sheetName + ' Request',
                      'New request from ' + dept + ' Department. ' +
                      'Link: ' + sheetURL, {
                      name: sheetName + ' Request Notification',
                      noReply: true
    });
  } else if (dept == 'Process Engineering') {
    MailApp.sendEmail('diana.pincus@makerbot.com',
                      'New ' + sheetName + ' Request',
                      'New request from ' + dept + ' Department. ' +
                      'Link: ' + sheetURL, {
                      name: sheetName + ' Request Notification',
                      noReply: true
    });
  } else if (dept == 'R&D') {
    MailApp.sendEmail('kate.nevin@makerbot.com',
                      'New ' + sheetName + ' Request',
                      'New request from ' + dept + ' Department. ' +
                      'Link: ' + sheetURL, {
                      name: sheetName + ' Request Notification',
                      noReply: true
    });
  } else if (dept == 'IT / Web') {
    MailApp.sendEmail('russell@makerbot.com',
                      'New ' + sheetName + ' Request',
                      'New request from ' + dept + ' Department. ' +
                      'Link: ' + sheetURL, {
                      name: sheetName + ' Request Notification',
                      noReply: true
    });
  } else if (dept == 'Test Engineering') {
    MailApp.sendEmail('james.gunipero@makerbot.com',
                      'New ' + sheetName + ' Request',
                      'New request from ' + dept + ' Department. ' +
                      'Link: ' + sheetURL, {
                      name: sheetName + ' Request Notification',
                      cc: 'tom.montagliano@makerbot.com',
                      noReply: true
    });
  } else if (dept == 'Quality Engineering') {
    MailApp.sendEmail('james.gunipero@makerbot.com',
                      'New ' + sheetName + ' Request',
                      'New request from ' + dept + ' Department. ' +
                      'Link: ' + sheetURL, {
                      name: sheetName + ' Request Notification',
                      noReply: true
    });
  } else if (dept == 'Facilities / OSHW') {
    MailApp.sendEmail('patrick.quinlivan@makerbot.com',
                      'New ' + sheetName + ' Request',
                      'New request from ' + dept + ' Department. ' +
                      'Link: ' + sheetURL, {
                      name: sheetName + ' Request Notification',
                      noReply: true
    });
  } else if (dept == 'BotFarm') {
    MailApp.sendEmail('mikeb@makerbot.com',
                      'New ' + sheetName + ' Request',
                      'New request from ' + dept + ' Department. ' +
                      'Link: ' + sheetURL, {
                      name: sheetName + ' Request Notification',
                      noReply: true
    });
  } else if (dept == 'Repair') {
    MailApp.sendEmail('brian.stamile@makerbot.com',
                      'New ' + sheetName + ' Request',
                      'New request from ' + dept + ' Department. ' +
                      'Link: ' + sheetURL, {
                      name: sheetName + ' Request Notification',
                      noReply: true
    });
  } else if (dept == 'Production') {
    MailApp.sendEmail('eshwar.ravichandran@makerbot.com',
                      'New ' + sheetName + ' Request',
                      'New request from ' + dept + ' Department. ' +
                      'Link: ' + sheetURL, {
                      name: sheetName + ' Request Notification',
                      noReply: true
    });
  } else if (dept == 'Software') {
    MailApp.sendEmail('kris.sandine@makerbot.com',
                      'New ' + sheetName + ' Request',
                      'New request from ' + dept + ' Department. ' +
                      'Link: ' + sheetURL, {
                      name: sheetName + ' Request Notification',
                      noReply: true
    });
  } else if (dept == 'Compliance') {
    MailApp.sendEmail('james.gunipero@makerbot.com',
                      'New ' + sheetName + ' Request',
                      'New request from ' + dept + ' Department. ' +
                      'Link: ' + sheetURL, {
                      name: sheetName + ' Request Notification',
                      noReply: true
    });
  } else {
    MailApp.sendEmail('nicholas.hotto@makerbot.com',
                      'New ' + sheetName + ' Request',
                      'New request from ' + dept + ' Department. ' +
                      'Link: ' + sheetURL, {
                      name: sheetName + ' Request Notification',
                      noReply: true
     });
  }
}
