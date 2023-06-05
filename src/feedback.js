// Original code from https://github.com/jamiewilson/form-to-google-sheets
// Updated for 2021 and ES6 standards

//  Used in GsSQL sheet to Get/Post feedback to Select2Query comment web page.

/*
    After running 'initialSetup()' and adding the trigger for 'doPost()'
    Publish the project as a web app

    Click on Publish > Deploy as web app….
    Set Project Version to New and put initial version in the input field below.
    Leave Execute the app as: set to Me(your@address.com).
    For Who has access to the app: select Anyone, even anonymous.
    Click Deploy.
    In the popup, copy the Current web app URL from the dialog.
    And click OK.

*/

const sheetName = 'Sheet1';
const scriptProp = PropertiesService.getScriptProperties();                      // skipcq:  JS-0125

/**
 * Run manually when setting up first time in Sheets.
 */
function initialSetup() {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();              // skipcq:  JS-0125
    scriptProp.setProperty('key', activeSpreadsheet.getId());
}

/**
 * Add a new project trigger
 *
 *  Click on Edit > Current project’s triggers.
 *  In the dialog click No triggers set up. Click here to add one now.
 *  In the dropdowns select doPost
 *  Set the events fields to From spreadsheet and On form submit
 *  Then click Save
 * @param {*} e 
 * @returns 
 */
function doPost(e) {
    const lock = LockService.getScriptLock();
    lock.tryLock(10000);

    try {
        Logger.log("Starting to record comment...");                        // skipcq:  JS-0125
        const doc = SpreadsheetApp.openById(scriptProp.getProperty('key'));
        const sheet = doc.getSheetByName(sheetName);

        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const nextRow = sheet.getLastRow() + 1;

        const newRow = headers.map(function (header) {
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

        // skipcq:  JS-0125
        return ContentService
            .createTextOutput(JSON.stringify({ 'result': 'success', 'row': nextRow }))
            .setMimeType(ContentService.MimeType.JSON);                       // skipcq:  JS-0125
    }

    catch (e) {
        return ContentService                                               // skipcq:  JS-0125
            .createTextOutput(JSON.stringify({ 'result': 'error', 'error': e }))
            .setMimeType(ContentService.MimeType.JSON);                       // skipcq:  JS-0125
    }

    finally {
        lock.releaseLock();
    }
}

/**
 * Returns sheet content as JSON.
 * @param {*} e 
 * @returns 
 */
function doGet(e) {
    let project = typeof e.parameter !== 'undefined' && typeof e.parameter.project !== 'undefined' ? e.parameter.project : "";

    const content = getSheetData(project);
    const contentObject = { GoogleSheetData: content }
    return ContentService.createTextOutput(JSON.stringify(contentObject)).setMimeType(ContentService.MimeType.JSON);     // skipcq:  JS-0125
}

/**
 * For testing inside Google Sheet
 */
function testGetSheetData() {
    const testData = getSheetData("");
}

/**
 * Get comments as 2D array
 * @param {String} project Will retrieve comments for a specific project name.
 * @returns {String[]}
 */
function getSheetData(project) {
    Logger.log("Project to filter: " + project);                      // skipcq:  JS-0125
    const ss = SpreadsheetApp.getActiveSpreadsheet();                  // skipcq:  JS-0125
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

/**
 * Format comment date for display.
 * @param {*} date 
 * @returns {String}
 */
function getFormattedDate(date) {
    const year = date.getFullYear();

    let month = (1 + date.getMonth()).toString();
    month = month.length > 1 ? month : `0${month}`;

    let day = date.getDate().toString();
    day = day.length > 1 ? day : `0${day}`;

    return `${month}/${day}/${year}`;
}
