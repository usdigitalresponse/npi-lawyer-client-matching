function onOpen() {
  var menuItems = [
    {name: 'Create Matches', functionName: 'performMatching'},
    {name: 'Email Lawyers', functionName: 'emailLawyers'}
  ];
  SpreadsheetApp.getActive().addMenu('ESP Actions', menuItems);
}

// Sort in reverse to delete from bottom up.
function sortDescending(firstVal, secondVal) {
  if (firstVal < secondVal) {
    return 1;
  }
  if (firstVal > secondVal) {
    return -1;
  }
  return 0;  
}

function onEdit(e) {
  try {
    const range = e.range;
    if (range.getSheet().getName() === 'Confirmations Raw') {
      const confirmationsRaw = new SheetClass('Confirmations Raw');
      const awaitingConfirmation = new SheetClass('Awaiting Confirmation');
      let rowsToDelete = [];
      const lastRow = range.getLastRow();
      const firstRow = lastRow - range.getHeight() + 1;
      for (let row = firstRow; row <= lastRow; row++) {
        const id = confirmationsRaw.getRowData(row)[0][confirmationsRaw.columnIndex('Case')];
        const rowNumber = awaitingConfirmation.lookupRowIndex('Attorney Name - Client Name', id) + 1;
        if (rowNumber > 0) {
          rowsToDelete.push(rowNumber);
        }
      }
      rowsToDelete.sort(sortDescending);
      for (rn of rowsToDelete) {
        awaitingConfirmation.sheet.deleteRow(rn);
      }
    }
  } catch(e) {
    showOKAlert('onEdit catch', e);
  }
}

function showOKAlert(header, body) {
  ui = SpreadsheetApp.getUi();
  ui.alert(header, body, ui.ButtonSet.OK);
}

var runningTests = false;
function showAlert(title, msg) {
  if (runningTests) {
    console.log(title + ': ' + msg);
  } else {
    try {
      showOKAlert(title, msg);
    } catch(err) {
      console.log(title + ': ' + msg);
    }
  }
}

const maxColumns = 200;
class SheetClass {
  constructor(name) {
    this.name = name;
    this.sheet = SpreadsheetApp.getActive().getSheetByName(name);
    this.findLastColumnHeader();
    let headerRange = this.sheet.getRange('A1:' + this.lastColumn + '1');
    this.headerData = headerRange.getValues();
  }
  removeEmptyCells(rowData) {
    let i;
    let lastCol = rowData.length - 1;
    for (i = lastCol; i >= 0; i--) {
      if (rowData[i] === '') {
        rowData.pop();
      } else {
        break;
      }
    }
  }
  findLastColumnHeader() {
    let rangeSpec = 'A1:' + this.columnLetterFromIndex(maxColumns - 1) + '1';
    let headerRange = this.sheet.getRange(rangeSpec);
    let headerData = headerRange.getValues()[0];
    this.removeEmptyCells(headerData);
    if (headerData.length === maxColumns) {
      showAlert('Warning',
                  'Sheet: "' + this.name + '" may have more than ' + maxColumns +
                    ' columns. Ignoring columns after: ' + maxColumns + '.')
    }
    this.headerData = headerData;
    this.lastColumn = this.columnLetterFromIndex(headerData.length - 1);
  }
  columnIndex(columnName) {
    let index = this.headerData[0].indexOf(columnName);
    if (index < 0) {
      let msg = 'Unknown column named: "' + columnName + '" in sheet: "' + this.name + '"?';
      showAlert('Error', msg);
      throw msg;
    }
    return index;
  }
  columnName(columnIndex) {
    if (columnIndex >= this.headerData[0].length) {
      let msg = 'Column index too big: "' + columnIndex + '" in sheet: "' + this.name + '"?';
      showAlert('Error', msg);
      throw msg;
    }
    return this.headerData[0][columnIndex];
  }
  getRowCount() {
    let values = this.sheet.getRange("A1:A").getValues();
    return values.filter(String).length;
  }
  getRowData(rowNumber) {
    let rangeSpec = 'A' + rowNumber + ':' + this.lastColumn + rowNumber;
    try {
      let range = this.sheet.getRange(rangeSpec);
      return range.getValues();
    } catch(err) {
      console.log('Exception', 'Sheet: "' + this.name + '", range: ' + rangeSpec);
      throw err;
    }
  }
  setRowData(rowNumber, data) {
    let range = this.sheet.getRange('A' + rowNumber + ':' + this.lastColumn + rowNumber);
    range.setValues(data);
  }
  columnLetterFromIndex(columnIdx) {
    let charCodeA = 'A'.charCodeAt(0);
    let higherOrderDigit = Math.floor(columnIdx / 26);
    let columnLetter = '';
    if (higherOrderDigit > 0) {
      columnLetter = String.fromCharCode(charCodeA + higherOrderDigit - 1);
    }
    columnLetter += String.fromCharCode(charCodeA + (columnIdx % 26));
    return columnLetter;
  }
  columnLetterFromName(columnName) {
    return this.columnLetterFromIndex(this.columnIndex(columnName));
  }
  setCellData(rowNumber, columnName, data) {
    let columnLetter = this.columnLetterFromName(columnName);
    let range = this.sheet.getRange(columnLetter + rowNumber + ':' + columnLetter + rowNumber);
    let arr = [[data]];
    range.setValues(arr);
  }
  sortSheet(sortColumn, ascendingVal) {
    let address = 'A2:' + this.lastColumn + this.getRowCount();
    let range = this.sheet.getRange(address);
    let columnIdx = this.columnIndex(sortColumn) + 1;
    range.sort({column: columnIdx, ascending: ascendingVal});
  }
  lookupRowIndex(columnName, keyValue) { // (0-based) 
    let columnLetter = this.columnLetterFromName(columnName);
    let values = this.sheet.getRange(columnLetter + '1:' + columnLetter).getValues();
    let i = 0;
    while (i < values.length) {
      if (values[i][0] === keyValue) {
        return i;
      }
      i++;
    }
    return -1; // throw '"' + keyValue + '" not found ' + ' in column: "' + columnName + '"';
  }
  clear() {
    let rowCount = this.getRowCount();
    if (rowCount > 1) {
      let address = 'A2:' + this.lastColumn + rowCount;
      let range = this.sheet.getRange(address);
      range.clear();
    }
  }
  log(data) {
    let d = new Date();
    data.unshift(d);
    this.sheet.appendRow(data);
    console.log(data);
  }
  cloneSheet(sourceId, sourceSheetName) {
    let sourceWorkbook = SpreadsheetApp.openById(sourceId);
    let sourceSheet = sourceWorkbook.getSheetByName(sourceSheetName);
    let fullRange = sourceSheet.getDataRange();
    let rangeSpec = fullRange.getA1Notation();
    let sData = fullRange.getValues();
    this.sheet.clear({contentsOnly: true});
    this.sheet.getRange(rangeSpec).setValues(sData);
  }
  copyFrom(sourceSheetName, sourceRange) {
    let sourceSheet = SpreadsheetApp.getActive().getSheetByName(sourceSheetName);
    let fullRange = sourceSheet.getRange(sourceRange);
    let sData = fullRange.getValues();
    this.clear();
    this.sheet.getRange(sourceRange).setValues(sData);
  }
  getAllRows() {
    let ret = [];
    for (let i = 0; i < this.getRowCount(); i++) {
      ret.push(this.getRowData(i + 1)[0]);
    }
    return ret;
  }
}

class SheetRowIterator {
  constructor(sheet) {
    this.sheet = sheet;
    this.lastIndex = this.sheet.getRowCount();
    this.nextIndex = 2;
  }
  getNextRow() {
    if (this.nextIndex > this.lastIndex) {
      return null;
    }
    return this.sheet.getRowData(this.nextIndex++)[0];
  }
}

var clients = new SheetClass('Clients Raw');
var logSheet = new SheetClass('Do NOT Edit - Log');
var lineSep = String.fromCharCode(10);

function compareByCourtDate(firstElement, secondElement) {
  let courtDateIndex = clients.columnIndex('Court Date' + lineSep + 'auto');
  let firstRow = clients.getRowData(firstElement);
  let secondRow = clients.getRowData(secondElement);
  let firstDate = firstRow[0][courtDateIndex];
  let secondDate = secondRow[0][courtDateIndex];
  if (firstDate < secondDate) {
    return -1;
  }
  if (firstDate > secondDate) {
    return 1;
  }
  return 0;  
}

class TheApp {
  constructor() {
    this.availabilityColHeader = 'How many cases can you take on this week?';
  }
  buildSortedClientArray(clients) {
    let indexArray = [];
    let confirmationIndex = clients.columnIndex('Clerk Confirmation' + lineSep + 'manual');
    let matchStatusIndex = clients.columnIndex('Match Status' + lineSep + ' auto - Pending, Confirmed, Denied' + lineSep + 'manual for Reassigned');
    let lastClientIndex = clients.getRowCount();
    let clientIndex;
    for (clientIndex = 2; clientIndex <= lastClientIndex; clientIndex++) {
      let clientData = clients.getRowData(clientIndex)[0];
      if (clientData[confirmationIndex]) {
        if (!clientData[matchStatusIndex]) {
          indexArray.push(clientIndex);
        }
      }
    }
    indexArray.sort(compareByCourtDate);
    return indexArray;
  }
  cleanUpAvailabilities(availabilities, attorneys) {
    availabilities.sortSheet('Timestamp', false);
    let lastAvailabilityIndex = availabilities.getRowCount();
    let availabilityColIndex = availabilities.columnIndex(this.availabilityColHeader);
    let availabilityIndex;
    let attorneyUUIDs = [];
    for (availabilityIndex = 2; availabilityIndex <= lastAvailabilityIndex; availabilityIndex++) {
      let availabilityData = availabilities.getRowData(availabilityIndex)[0];
      let uuid = availabilityData[availabilities.columnIndex('Name')];
      if (availabilityData[availabilityColIndex] > 0) {
        if (attorneyUUIDs.indexOf(uuid) >= 0) {
          availabilities.setCellData(availabilityIndex, this.availabilityColHeader, 0);
        } else {
          attorneyUUIDs.push(uuid);
        }
      }
      let typeIndex;
      let attorneyRowIndex = attorneys.lookupRowIndex('Name', uuid);
      if (attorneyRowIndex < 0) {
        console.log('Error', 'No row for attorney in Staff List: "' + uuid + '". Skipping it.');
      } else {
        let attorneyType = attorneys.getRowData(attorneyRowIndex + 1)[0][attorneys.columnIndex('Type')];
        availabilities.setCellData(availabilityIndex, 'Type', attorneyType);
        switch (attorneyType) {
          case 'Pro Bono Attorney':
            typeIndex = 1; break;
          case 'Law Student/Former Law Student':
            typeIndex = 2; break;
          case 'NPI Staff Attorney':
            typeIndex = 3; break;
          default:
            typeIndex = 4;
        }
        availabilities.setCellData(availabilityIndex, 'Type Rank', typeIndex);
      }
    }
  }
  updateStaff(attorneys) {
    let newStaffList = new SheetClass('Staff List');
    let nextStaffIndex = attorneys.getRowCount() + 1;
    let d = new Date();
    let newStaffIterator = new SheetRowIterator(newStaffList);
    let newStaffData;
    while (newStaffData = newStaffIterator.getNextRow()) {
      let name = newStaffData[newStaffList.columnIndex('First Name')] + ' ' + newStaffData[newStaffList.columnIndex('Last Name')]
      if (attorneys.lookupRowIndex('Name', name) == -1) {
        let newRow = [];
        newRow[attorneys.columnIndex('Timestamp')] = d;
        newRow[attorneys.columnIndex('FirstName')] = newStaffData[newStaffList.columnIndex('First Name')];
        newRow[attorneys.columnIndex('LastName')] = newStaffData[newStaffList.columnIndex('Last Name')];
        newRow[attorneys.columnIndex('Email')] = newStaffData[newStaffList.columnIndex('Email')];
        let tColumnHeader = 'What is your affiliation? Let us know if you are a pro bono attorney, NPI Staff member, or a current/former law student.';
        newRow[attorneys.columnIndex('Type')] = newStaffData[newStaffList.columnIndex(tColumnHeader)];
        let sColumnHeader = 'Do you speak Spanish? Selecting yes will allow us to match you with Spanish speaking clients.';
        newRow[attorneys.columnIndex('Spanish?')] = newStaffData[newStaffList.columnIndex(sColumnHeader)];
        newRow[attorneys.columnIndex('Name')] = name;
        attorneys.setRowData(nextStaffIndex++, [newRow]);
      }
    }
  }
  doMatching() {
    let sortedClientArray = this.buildSortedClientArray(clients);
    if (sortedClientArray.length === 0) {
      showAlert('Warning', 'No clients found with "Clerk Confirmation" set to "Yes" with a blank "Match Status".');
      return;
    }
    let availabilities = new SheetClass('Ranked Availability');
    let rawAvailabilities = new SheetClass('Availability Raw');
    let attorneys = new SheetClass('Staff List');

    availabilities.copyFrom('Availability Raw', 'A2:C' + rawAvailabilities.getRowCount());
    this.cleanUpAvailabilities(availabilities, attorneys);
    availabilities.sortSheet('Type Rank', true);
//    this.updateStaff(attorneys); // Until the Google Form for adding attorneys is enabled.
    attorneys.sortSheet('FirstName', true);
    let emailedMatches = new SheetClass('Emailed Matches');
    let matches = new SheetClass('Created Matches');
    matches.clear();

    let lastAvailabilitiesIndex = availabilities.getRowCount();
    let nextMatchIndex = 2;
    let availabilityIndex = 2;
    let d = new Date();
    let clientIndex;
    for (clientIndex = 0; clientIndex < sortedClientArray.length; clientIndex++) {
      if (availabilityIndex > lastAvailabilitiesIndex) { // Check if no one available at all.
        break;
      }
      let clientData = clients.getRowData(sortedClientArray[clientIndex])[0];
      let caseNumber = clientData[clients.columnIndex('Case Number' + lineSep + 'auto')];
      if (emailedMatches.lookupRowIndex('Case Number', caseNumber) != -1) {
        showAlert('Warning', 'Case: ' + caseNumber + " has already been emailed, skipping it.");
        continue;
      }
      let availabilityData = availabilities.getRowData(availabilityIndex)[0];
      let availabilityColIndex = availabilities.columnIndex(this.availabilityColHeader);
      while (availabilityData[availabilityColIndex] <= 0) {
        availabilityIndex++;
        if (availabilityIndex > lastAvailabilitiesIndex) {
          availabilityData = null;
          break;
        }
        availabilityData = availabilities.getRowData(availabilityIndex)[0];
      };
      if (!availabilityData) {
        break;
      }
      let attorneyName = availabilityData[availabilities.columnIndex('Name')];
      let attorneyData = attorneys.getRowData(attorneys.lookupRowIndex('Name', attorneyName) + 1)[0];
      let lawyerName = attorneyName.split(' ');

      let match = [];
      match[matches.columnIndex('Timestamp')] = d;
      match[matches.columnIndex('Lawyer First Name')] = lawyerName[0];
      match[matches.columnIndex('Lawyer Last Name')] = lawyerName[1];
      match[matches.columnIndex('Lawyer Email')] = attorneyData[attorneys.columnIndex('Email')];
      match[matches.columnIndex('Client First Name')] = clientData[clients.columnIndex('First' + lineSep + 'auto')];
      match[matches.columnIndex('Client Last Name')] = clientData[clients.columnIndex('Last' + lineSep + 'auto')];
      match[matches.columnIndex('Client Email')] = clientData[clients.columnIndex('Email' + lineSep + 'auto')];
      match[matches.columnIndex('Client UUID')] = clientData[clients.columnIndex('Unique ID' + lineSep + 'auto')];
      match[matches.columnIndex('Client Folder')] = clientData[clients.columnIndex('Folder' + lineSep + 'auto')];
      match[matches.columnIndex('Client Phone Number')] = clientData[clients.columnIndex('Phone' + lineSep + 'auto')];
      match[matches.columnIndex('Client Address')] = clientData[clients.columnIndex('Address'  + lineSep + 'auto')];
      match[matches.columnIndex('Landlord Name')] = clientData[clients.columnIndex('Landlord Name'  + lineSep + 'auto')];
      match[matches.columnIndex('Landlord Email')] = clientData[clients.columnIndex('Landlord Email'  + lineSep + 'auto')];
      match[matches.columnIndex('Landlord Phone Number')] = clientData[clients.columnIndex('Landlord Phone' + lineSep + 'auto')];
      match[matches.columnIndex('Landlord Address')] = clientData[clients.columnIndex('Landlord Address' + lineSep + 'auto')];
      match[matches.columnIndex('Case Number')] = clientData[clients.columnIndex('Case Number' + lineSep + 'auto')];
      match[matches.columnIndex('Next Court Date')] = clientData[clients.columnIndex('Court Date' + lineSep + 'auto')];
      match[matches.columnIndex('Match Status')] = '';
      match[matches.columnIndex('Pending Timestamp')] = '';
      matches.setRowData(nextMatchIndex, [match]);

      nextMatchIndex++;
      availabilities.setCellData(availabilityIndex, this.availabilityColHeader, --availabilityData[availabilityColIndex]);
    }
    nextMatchIndex -= 2;
    let leftOver = sortedClientArray.length - nextMatchIndex;
    let msg = 'Matched ' + nextMatchIndex + ' clients. ' + leftOver + ' clients not matched.';
    logSheet.log([msg]);
  }
  performMatching() {
    clients.cloneSheet('1vnUVqjwj-u6Wn2v4rhBZN5qvfic6Pa7prLMMLGElBzo', 'Client List')
    doMatching();
  }
  emailLawyers() {
    let d = new Date();
    let newCaseCount = 0;
    let emailedMatches = new SheetClass('Emailed Matches');
    let awaitingConfirmation = new SheetClass('Awaiting Confirmation');
    awaitingConfirmation.clear();
    let nextEmailMatchIndex = emailedMatches.getRowCount() + 1;
    let matches = new SheetClass('Created Matches');
    let matchIterator = new SheetRowIterator(matches);
    let matchData;
    while (matchData = matchIterator.getNextRow()) {
      let newCaseNumber = matchData[matches.columnIndex('Case Number')];
      if (emailedMatches.lookupRowIndex('Case Number', newCaseNumber) != -1) {
        showAlert("Case: " + newCaseNumber + ' already emailed. Skipping it.');
        continue;
      }
      matchData[matches.columnIndex('Timestamp')] = d;
      matchData[matches.columnIndex('Match Status')] = '';
      emailedMatches.setRowData(nextEmailMatchIndex, [matchData]);
      let confirmationData = [];
      confirmationData[awaitingConfirmation.columnIndex('Timestamp')] = d;
      confirmationData[awaitingConfirmation.columnIndex('Attorney Name - Client Name')] = 
              matchData[matches.columnIndex('Lawyer First Name')] + 
              ' ' +
              matchData[matches.columnIndex('Lawyer Last Name')] + 
              ' - ' +
              matchData[matches.columnIndex('Client First Name')] +
              ' ' +
              matchData[matches.columnIndex('Client Last Name')];
      awaitingConfirmation.setRowData(nextEmailMatchIndex, [confirmationData]);
      nextEmailMatchIndex++;
      newCaseCount++;
    }
    showAlert('Info', 'Emailed ' + newCaseCount + ' new cases.');
  }
}

theApp = new TheApp();
function performMatching() { theApp.performMatching(); }
function emailLawyers() { theApp.emailLawyers(); }
function doMatching() { theApp.doMatching(); }

// ----------------------- code for automated testing
// Google Javascript isn't ES6,so no support for 'super' keyword. Thus the 'has-a' relationship. :(
class BaseSheetClass {
  constructor(name) {
    this.subSheet = new SheetClass(name);
    this.lastColumn = this.subSheet.columnLetterFromIndex(maxColumns);
  }
  getRowCount() {
    return this.subSheet.getRowCount();
  }
  getRowData(rowNumber) {
    let rangeSpec = 'A' + rowNumber + ':' + this.lastColumn + rowNumber;
    let range = this.subSheet.sheet.getRange(rangeSpec);
    let ret = range.getValues();
    this.subSheet.removeEmptyCells(ret);
    return ret;
  }
  removeEmptyCells(sheet, rowData) {
    let i;
    let lastCol = rowData.length - 1;
    let len = sheet.headerData[0].length;
    for (i = lastCol; i >= 0 && rowData.length > len; i--) {
      if (rowData[i] === '') {
        rowData.pop();
      } else {
        break;
      }
    }
  }
}
class Tester {
  constructor() {
    try {
      this.testDataSheet = new BaseSheetClass('Test Data');
    } catch(e) {
      console.log(e.toString());
      this.testDataSheet = null;
    }
  }
  loadAt(startRowNum, sheetName, rowCount, iter) {
    iter.getNextRow(); // Skip header.
    let sheet = new SheetClass(sheetName);
    if (startRowNum === 2) {
      sheet.clear();
    }
    while (rowCount--) {
      let rowData = iter.getNextRow();
      this.testDataSheet.removeEmptyCells(sheet, rowData);
      sheet.setRowData(startRowNum++, [rowData]);
    }
  }
  append(sheetName, rowCount, iter) {
    let sheet = new SheetClass(sheetName);
    this.loadAt(sheet.getRowCount() + 1, sheetName, rowCount, iter);
  }
  compareArrays(sheet, expected, actual) {
    if (expected.length != actual.length) {
      console.log('Expected length: ' + expected.length + ' not equal actual: ' + actual.length);
      return;
    }
    for (let i = 0; i < expected.length; i++) {
      let columnName = sheet.columnName(i); 
      if (columnName !== 'Timestamp') {
        if (expected[i].toString() !== actual[i].toString()) {
          console.log('Expected value: ');
          console.log('"' + expected[i] + '"');
          console.log('is not equal to actual: ')
          console.log('"'+ actual[i] + '"');
          return;   
        }           
      }
    }
  }
  verify(sheetName, rowCount, iter) {
    iter.getNextRow(); // Skip header.
    let sheet = new SheetClass(sheetName);
    if (rowCount !== sheet.getRowCount() - 1) {
      console.log('Sheet: ' + sheetName + 'Expected row count: ' + rowCount + ' not equal actual: ' + sheet.getRowCount());
    } else {
      let rowNum = 2;
      while (rowCount--) {
        let expected = iter.getNextRow();
        this.testDataSheet.removeEmptyCells(sheet, expected)
        let actual = sheet.getRowData(rowNum++);
        this.compareArrays(sheet, expected, actual[0])
      }
    }
    console.log('Verified sheet named: "' + sheetName + '"');
  }
  runTests() {
    if (this.testDataSheet) {
      runningTests = true;
      let clearedSheetNames = [
        'Created Matches',
        'Emailed Matches',
        'Awaiting Confirmation',
        'Confirmations Raw',
        'Confirmed Matches',
        'Staff List',
        'Availability Raw',
        'Ranked Availability',
        'Clients Raw'
      ];
      for (const name of clearedSheetNames) {
        (new SheetClass(name)).clear();
      }
      let iter = new SheetRowIterator(this.testDataSheet);
      let testRowData;
      while (testRowData = iter.getNextRow()) {
        let action = testRowData[0];
        switch(action) {
          case 'load':
            this.loadAt(2, testRowData[1], testRowData[2], iter);
            break;
          case 'append':
            this.append(testRowData[1], testRowData[2], iter);
            break;
          case 'verify':
            this.verify(testRowData[1], testRowData[2], iter);
            break;
          case 'doMatching':
            theApp.doMatching();
            break;
          case 'emailLawyers':
            theApp.emailLawyers();
            break;
          default:
            showAlert("Warning", "Unknown action: " + action);
        }
      }
    }
    runningTests = false;
  }
}
function runTests() {
  tester = new Tester();
//  tester.runTests();
}
