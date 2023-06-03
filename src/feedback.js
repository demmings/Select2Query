// Original code from https://github.com/jamiewilson/form-to-google-sheets
// Updated for 2021 and ES6 standards

//  Used in GsSQL sheet to Get/Post feedback to Select2Query comment web page.

const sheetName = 'Sheet1'
const scriptProp = PropertiesService.getScriptProperties()

function initialSetup () {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  scriptProp.setProperty('key', activeSpreadsheet.getId())
}

function doPost (e) {
  const lock = LockService.getScriptLock()
  lock.tryLock(10000)

  try {
    Logger.log("Starting to record comment...");
    const doc = SpreadsheetApp.openById(scriptProp.getProperty('key'));
    const sheet = doc.getSheetByName(sheetName);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const nextRow = sheet.getLastRow() + 1;

    const newRow = headers.map(function(header) {
      if (header === 'Date') {
        return new Date();
      }
      else if (e.parameter[header].startsWith("=")) {
         return `"${e.parameter[header]}"`;
      }
      else {
          return e.parameter[header];
      }
    });

    const newCommentRange = sheet.getRange(nextRow, 1, 1, newRow.length);
    newCommentRange.setNumberFormat('@STRING@');
    newCommentRange.setValues([newRow]);
    
    return ContentService
      .createTextOutput(JSON.stringify({ 'result': 'success', 'row': nextRow }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  catch (e) {
    return ContentService
      .createTextOutput(JSON.stringify({ 'result': 'error', 'error': e }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  let project = typeof e.parameter !== 'undefined' && typeof e.parameter.project !== 'undefined' ? e.parameter.project : "";

  var content = getSheetData(project);
  var contentObject = {GoogleSheetData: content}
  return ContentService.createTextOutput(JSON.stringify(contentObject) ).setMimeType(ContentService.MimeType.JSON); 
}

function testGetSheetData() {
  const testData = getSheetData("");
}

function getSheetData(project)  { 
  Logger.log("Project to filter: " + project);
  const ss= SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('Sheet1'); 
  const dataRange = dataSheet.getDataRange();
  let dataValues = dataRange.getValues();  

  const projectColumn = dataValues[0].indexOf("Project");
  const dateColumn = dataValues[0].indexOf("Date");

  let titleData = [dataValues.shift()];   //  Title row.
  if (projectColumn !== -1 && project.trim().length !== 0) {
    dataValues = dataValues.filter(row => row[projectColumn].toUpperCase() === project.toUpperCase());
  }

  if (dateColumn !== -1) {
    for (let row of dataValues) {
      if (row[dateColumn] instanceof Date) {
        row[dateColumn] = getFormattedDate(row[dateColumn]);
      } 
    }
  }

  dataValues = titleData.concat(dataValues);

  //  Remove project column

  for (let i = 0; i < dataValues.length; i++) {
    if (projectColumn !== -1) {
       let row = dataValues[i];
       row.splice(projectColumn, 1); 
    }
  }

  return dataValues;
}

function getFormattedDate(date) {
  var year = date.getFullYear();

  var month = (1 + date.getMonth()).toString();
  month = month.length > 1 ? month : '0' + month;

  var day = date.getDate().toString();
  day = day.length > 1 ? day : '0' + day;
  
  return month + '/' + day + '/' + year;
}
