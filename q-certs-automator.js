// Replace with your document IDs before running the script:


// 1. Create a GDoc certificate template and put the ID here
var TEMPLATE_ID = '1jWnX6F7xBs-S1QRyXyvZhcWBw229DChuY_QstiQVAdQ'

// 2. Create a folder for the certificates and put the ID here
var CERTIFICATES_FOLDER_ID = '14LMqhoWCLhG6mJEv0DhjTMOZtopFrDiy'

// 3. Put the ID for the GDoc with the email body here and add a subject line
var EMAIL_TEMPLATE_ID = '1oFT_hWvVOLDE7RJY659JSqOXaCLEw0QIjhJ08XMdHvY'
var EMAIL_SUBJECT = 'Q Award Certificate (A gift from "Robot Stuff")'



/**
 * This function adds a menu to your spreadsheet. Only needs to be run once.
 */

function onOpen() {
  SpreadsheetApp
    .getUi()
    .createMenu('Robot Stuff')
    .addItem('Create pdf', 'createPdf')
    .addItem('Send pdf', 'sendPdf')
    .addToUi()
} 


/**
 * This function adds the recipient first name to the email body. This needs to be updated if new template fields are added.
 */

function getEmailBody(firstName) {
  var templateCopy = DriveApp.getFileById(EMAIL_TEMPLATE_ID).makeCopy()
  var templateCopyId = templateCopy.getId()
  var openTemplateCopy = DocumentApp.openById(templateCopyId)
  var templateContent = openTemplateCopy.getBody()
  templateContent.replaceText('%FIRSTNAME%', firstName)
  var text = templateContent.getText()
  return text 
}


/**
 * This function will send the pdfs. Run it from your spreadsheet.
 */

function sendPdf() {
  var activeSheet = SpreadsheetApp.getActiveSheet()
  var numberOfColumns = activeSheet.getLastColumn()
  var headerRow = activeSheet.getRange(1, 1, 1, numberOfColumns).getValues()
  var activeRange = activeSheet.getActiveRange()
  var currentRowIndex = activeRange.getRowIndex()
  var lastRowIndex = activeRange.getLastRow()
  var pdfLinksColumn = headerRow[0].findIndex(item => item == 'PDF')+1
  var emailColumn = headerRow[0].findIndex(item => item == 'Email')+1
  var firstNameColumn = headerRow[0].findIndex(item => item == 'FIRSTNAME')+1
    
  for(; currentRowIndex <= lastRowIndex; currentRowIndex++ ) {
   var firstName = activeSheet.getRange(currentRowIndex, firstNameColumn, 1, 1).getValue()
   var EMAIL_BODY = getEmailBody(firstName)
   var recipient = activeSheet.getRange(currentRowIndex, emailColumn, 1, 1).getValue()
   var pdfLink = activeSheet.getRange(currentRowIndex, pdfLinksColumn, 1, 1).getValue()
   const regex = new RegExp(/[-\w]{25,}/)
   var pdfId = regex.exec(pdfLink)
   var pdf = DriveApp.getFileById(pdfId)
  
   MailApp.sendEmail(
     recipient, 
     EMAIL_SUBJECT, 
     EMAIL_BODY,
     {attachments: [pdf]})
  }    
}


/**
 * This function creates the pdfs. Run it from the menu in your spreadsheet.
 */

function createPdf() {
  var fileNameColumnName = 'File Name'
  var pdfLinkColumnName = 'PDF'
  var certificatesFolder = DriveApp.getFolderById(CERTIFICATES_FOLDER_ID)
  
  var ui = SpreadsheetApp.getUi()

  if (!TEMPLATE_ID) {
    ui.alert('You forgot to add your TEMPLATE_ID to the script!')
    return
  }
  var activeSheet = SpreadsheetApp.getActiveSheet()
  var numberOfColumns = activeSheet.getLastColumn()
  var headerRow = activeSheet.getRange(1, 1, 1, numberOfColumns).getValues()
  var activeRange = activeSheet.getActiveRange()
  var currentRowIndex = activeRange.getRowIndex()
  var lastRowIndex = activeRange.getLastRow()
  
  for(; currentRowIndex <= lastRowIndex; currentRowIndex++ ) {
    var templateCopy = DriveApp.getFileById(TEMPLATE_ID).makeCopy()
    var templateCopyId = templateCopy.getId()
    var openTemplateCopy = DocumentApp.openById(templateCopyId)
    var templateContent = openTemplateCopy.getActiveSection()
    var currentRowValue = activeSheet.getRange(currentRowIndex, 1, 1, numberOfColumns).getValues()
    var headerValue
    var activeCell
    var fileName
      
    for (var columnIndex = 0; columnIndex < headerRow[0].length; columnIndex++) { 
      var columnName = headerRow[0][columnIndex]
      var columnValue = currentRowValue[0][columnIndex]
      templateContent.replaceText('%' + columnName + '%', columnValue)
      
      if (columnName === fileNameColumnName) {
        fileName = columnValue       
      } 
    }
    
    openTemplateCopy.saveAndClose()
    var pdf = certificatesFolder.createFile(templateCopy.getAs('application/pdf'))
    pdf.setName(fileName)
    var pdfLink = pdf.getUrl()    
    var pdfLinksColumn = headerRow[0].findIndex(item => item == 'PDF')+1
    var linkCell = activeSheet.getRange(currentRowIndex, pdfLinksColumn, 1, 1)
    linkCell.setValue(pdfLink)
    templateCopy.setTrashed(true)
   }
  
} 
