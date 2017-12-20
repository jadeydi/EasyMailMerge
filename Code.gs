function onOpen() {
  SpreadsheetApp.getUi().createAddonMenu()
  .addItem('Start Mail Merge', 'showTemplatesModal')
  .addToUi();
}

function showTemplatesModal() {
  var htmlOutput = HtmlService
  .createTemplateFromFile("Page")
  .evaluate()
  .setWidth(200)
  .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Select Template");
}

function getMessages() {
  var drafts = GmailApp.getDrafts();
  var msgs = [];
  for (var i=0; i<drafts.length; i++){
    var msg = drafts[i].getMessage();
    msgs.push({"msgId": msg.getId(), "subject": msg.getSubject()});
  }
  return msgs
}

function sendEmail(id, withRecipient) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var msg = GmailApp.getMessageById(id);
  var options = {
    weekday: "long", year: "numeric", month: "short",
    day: "numeric", hour: "2-digit", minute: "2-digit"
  }
  for (var i = 1; i < data.length; i++) {
    var email = data[i][0]
    var body = msg.getBody();
    var j = 2;
    if (withRecipient) {
      body = body.replace(/{{ name }}/gi, data[i][1]);
      j = 3;
    }
    if (data[i].length > 2) {
      j = data[i].length;
    }
    if (!validateEmail(email)) {
      sheet.getRange(i+1, j).setBackground("#f9bcbc").setValue("Failure").setNote("Date: " + new Date());
      continue
    }
    sheet.getRange(i+1, j).setBackground("#b6d7a8").setValue("Success").setNote("Date: " + new Date());
    GmailApp.sendEmail(email, msg.getSubject(),"", {"htmlBody": body});
  }
}

function validateEmail(email) {
  var re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
  return re.test(email.toLowerCase());
}
