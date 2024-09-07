function getConfig() {
  return {
    docTemplateId: '',
    destinationFolderId: '',
    sheetName: 'Form Responses 1',
    newDocName: 'מילוי טופס - {{שם מלא}} - {{תאריך}}',
    urlColName: 'url',
    menuName: 'אוטומציות',
    menuItemName: 'יצירת מסמך משורה בטבלה',
    isSendingMail: true,
    mailRecipients: ''
  }
}

function createNewGoogleDocs() {

  const config = getConfig()

  const googleDocTemplate = DriveApp.getFileById(config.docTemplateId);
  const destinationFolder = DriveApp.getFolderById(config.destinationFolderId);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);

  const rows = sheet.getDataRange().getValues();
  const headers = rows.shift()

  rows.forEach((row, rowIndex) => {

    // adding 1 because its a range not an array and 2 to row index to consider headers shifted
    const urlCell = sheet.getRange(rowIndex + 2, headers.indexOf(config.urlColName) + 1)

    // return if value (doc) exists
    if (urlCell.getValue()) {
      return
    }

    // get columns values
    const colValueByHeader = {}
    row.forEach((col, colIndex) => {
      colValueByHeader[headers[colIndex]] = col
    })

    // get the new doc name with the corresponding values
    const regex = /\{\{(.*?)\}\}/g;
    let newDocName = config.newDocName
    let match;
    while ((match = regex.exec(newDocName)) !== null) {
      newDocName = newDocName.replace(match[0], colValueByHeader[match[1]])
    }

    // copy the template with the new name
    const copy = googleDocTemplate.makeCopy(newDocName, destinationFolder);
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();

    // go through the headers and check if the body has this pattern
    row.forEach((col, colIndex) => {
      let header = headers[colIndex];
      let search = '{{' + header + '}}';
      let searchPattern = '(?i)' + search;  // Add the RE2 case-insensitive flag (?i) to the pattern
      body.replaceText(searchPattern, col);
    });

    // close the doc and update url value
    doc.saveAndClose();
    const newDocUrl = doc.getUrl()
    urlCell.setValue(newDocUrl);

    if (config.isSendingMail) {
      let mailSubject = 'מסמך חדש נוצר מטופס'
      let mailbody =''
      mailbody += '<p dir="rtl">'
      mailbody += `${colValueByHeader['שם מלא']} מילא את הטופס,`
      mailbody += '<br>'
      mailbody += 'הנה הלינק למסמך:'
      mailbody += '<br>'
      mailbody += newDocUrl
      mailbody += '<br>'
      mailbody += 'בברכה,'
      mailbody += '<br>'
      mailbody += 'האוטומציה :)'
      mailbody += '</p>'

      sendEmailNotification(mailSubject, mailbody)
    }

  });

}

function onOpen() {
  const config = getConfig()
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu(config.menuName);
  menu.addItem(config.menuItemName, 'createNewGoogleDocs');
  menu.addToUi();
}

function sendEmailNotification(subject, body) {
  const config = getConfig()
  const options = {
    htmlBody: body
  }
  const emailApp = GmailApp.sendEmail(config.mailRecipients, subject, "", options)
}
