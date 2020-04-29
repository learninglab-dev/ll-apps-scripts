// Add your own file IDs here before running

const EMAIL_LIST_SHEET = '1JMpAkjHG9of7f7kE9TEE_z692pE5ezKYwj_7ZFG-UCA'
const EMAILS_SHEET_NAME = 'Sheet1'
const FILE_TO_SHARE = '1RPQ4-Rg7W4z7KOkskIX1BhrwBlTsqrXzvqH1TM7b7z0'

function createUserList(sheetId) {
  const emailsSpreadsheet = SpreadsheetApp.openById(sheetId)
  const sheet = emailsSpreadsheet.getSheetByName(EMAILS_SHEET_NAME)
  const numCols = sheet.getLastColumn()
  const headerRow = sheet.getRange(1, 1, 1, numCols).getValues()
  let emailsCol
  headerRow[0].forEach((label, i) => {
    if (label === 'EMAILS') {
      emailsCol = i+1
    }
  })
  const colVals = (sheet.getRange(2, emailsCol, sheet.getLastRow(), 1).getValues()).flat()
  const emails = colVals.filter(val => val.indexOf('@') >= 0)
  return emails
}

function shareFile() {
  const emailList = createUserList(EMAIL_LIST_SHEET)
  const fileToShare = DriveApp.getFileById(FILE_TO_SHARE)
  Logger.log(fileToShare)
  Logger.log(emailList)
  fileToShare.addEditors(emailList)
  //fileToShare.addViewers(emailList)
}
