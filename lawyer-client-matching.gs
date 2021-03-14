function onOpen() {
  var menuItems = [
    {name: 'Create Matches', functionName: 'performMatching'},
    {name: 'Email Lawyers', functionName: 'emailLawyers'}
  ];
  SpreadsheetApp.getActive().addMenu('ESP Actions', menuItems); // Not available when running in debug mode in script window.
}

function onEdit(e) {
    try {
      const range = e.range;
      if (range.getSheet().getName() === 'Confirmations Raw') {
        // TODO: Handle the rare case where more than one row is added at once.
        const lastRow = range.getLastRow();
        const confirmationsRaw = new SheetClass('Confirmations Raw', 0);
        const id = confirmationsRaw.getRowData(lastRow)[0][confirmationsRaw.columnIndex('Case')];
        const response = confirmationsRaw.getRowData(lastRow)[0][confirmationsRaw.columnIndex('Do you accept the case?')];
        if (response === 'Yes, I am available and have no conflict') {
          const awaitingConfirmation = new SheetClass('Awaiting Confirmation', 0);
          const rowNumber = awaitingConfirmation.lookupRowNumber('Attorney Name - Client Name', id) + 1;
          if (rowNumber > 0) {
            awaitingConfirmation.sheet.deleteRow(rowNumber);
          }
        }
      }
    } catch(e) {
      showOKAlert('catch', e);
    }
}

function showOKAlert(header, body) {
  ui = SpreadsheetApp.getUi();
  ui.alert(header, body, ui.ButtonSet.OK);
}

function showAlert(title, msg) {
  try {
    showOKAlert(title, msg); // Not available when running in debug mode in script window.
  } catch(err) {
    console.log(title + ': ' + msg);
  }
}

const maxColumns = 200;
class SheetClass {
  constructor(name, formulaColumns) {
    this.name = name;
    this.formulaColumns = formulaColumns;
    this.sheet = SpreadsheetApp.getActive().getSheetByName(name);
    this.findLastColumnHeader();
    let headerRange = this.sheet.getRange('A1:' + this.lastColumn + '1');
    this.headerData = headerRange.getValues();
  }
  findLastColumnHeader() {
    let rangeSpec = 'A1:' + this.columnLetterFromIndex(maxColumns - 1) + '1';
    let headerRange = this.sheet.getRange(rangeSpec);
    let headerData = headerRange.getValues()[0];
    let i;
    for (i = maxColumns - 1; i >= 0; i--) {
      if (headerData[i] === '') {
        headerData.pop();
      } else {
        break;
      }
    }
    if (headerData.length === maxColumns) {
      showAlert('Warning',
                  'Sheet: "' + this.name + '" may have more than ' + maxColumns +
                    ' columns. Ignoring columns after: ' + maxColumns + '.')
    }
    this.lastColumn = this.columnLetterFromIndex(headerData.length - 1 - this.formulaColumns);
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
  getRowCount() {
    let values = this.sheet.getRange("A1:A").getValues();
    return values.filter(String).length;
  }
  getRowData(rowNumber) {
    let rangeSpec = 'A' + rowNumber + ':' + this.lastColumn + rowNumber;
    let range = this.sheet.getRange(rangeSpec);
    return range.getValues();
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
  lookupRowNumber(columnName, keyValue) { // TODO: rowIndex, not number
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

var clients = new SheetClass('Clients Raw', 0);
var logSheet = new SheetClass('Do NOT Edit - Log', 0);
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
    // TODO: remove rows where this is not: 'Yes, I am available and have no conflict'
    // TODO because this may not happen much.
    // const response = confirmationsRaw.getRowData(lastRow)[0][confirmationsRaw.columnIndex('Do you accept the case?')];

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
      let attorneyRowNumber = attorneys.lookupRowNumber('Name', uuid);
      let attorneyType = attorneys.getRowData(attorneyRowNumber)[0][attorneys.columnIndex('Type')];
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
  updateStaff(attorneys) {
    let newStaffList = new SheetClass('Staff List', 1);
    let nextStaffIndex = attorneys.getRowCount() + 1;
    let d = new Date();
    let newStaffIterator = new SheetRowIterator(newStaffList);
    let newStaffData;
    while (newStaffData = newStaffIterator.getNextRow()) {
      let name = newStaffData[newStaffList.columnIndex('First Name')] + ' ' + newStaffData[newStaffList.columnIndex('Last Name')]
      if (attorneys.lookupRowNumber('Name', name) == -1) {
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
    let availabilities = new SheetClass('Ranked Availability', 0);
    let rawAvailabilities = new SheetClass('Availability Raw', 0);
    let attorneys = new SheetClass('Staff List', 1);

    availabilities.copyFrom('Availability Raw', 'A2:C' + rawAvailabilities.getRowCount());
    this.cleanUpAvailabilities(availabilities, attorneys);
    availabilities.sortSheet('Type Rank', true);
//    this.updateStaff(attorneys); // Until the Google Form for adding attorneys is enabled.
    attorneys.sortSheet('FirstName', true);
    let emailedMatches = new SheetClass('Emailed Matches', 0);
    let matches = new SheetClass('Created Matches', 0);
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
      if (emailedMatches.lookupRowNumber('Case Number', caseNumber) != -1) {
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
      let attorneyData = attorneys.getRowData(attorneys.lookupRowNumber('Name', attorneyName))[0];
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
    let emailedMatches = new SheetClass('Emailed Matches', 0);
    let awaitingConfirmation = new SheetClass('Awaiting Confirmation', 0);
    awaitingConfirmation.clear();
    let nextEmailMatchIndex = emailedMatches.getRowCount() + 1;
    let matches = new SheetClass('Created Matches', 0);
    let matchIterator = new SheetRowIterator(matches);
    let matchData;
    while (matchData = matchIterator.getNextRow()) {
      let newCaseNumber = matchData[matches.columnIndex('Case Number')];
      if (emailedMatches.lookupRowNumber('Case Number', newCaseNumber) != -1) {
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
const STAFF_INDEX = 0;
const CLIENT_INDEX = 1;
const AVAILABILITY_INDEX = 2;
class Tester {
  constructor() {
    this.testData = [
      /* week 1 */ [
        [
          [ '', 'Attorney1', 'Name', 'chris.keith@gmail.com', 'NPI Staff Attorney', '', '' ], 
          [ '', 'Attorney2', 'Name', 'chris.keith@gmail.com', 'Law Student/Former Law Student', '', '' ], 
          [ '', 'Attorney3', 'Name', 'chris.keith@gmail.com', 'Pro Bono Attorney', '', '' ], 
        ],
        [
          [ 'STEVE ', 'TESTING EVICTIONS', 'chris.keith@gmail.com', '1-4734W2', 'https://drive.google.com/drive/folders/1e6CgsjOpG3j5p4RlDpDDYMwTEhqMEqSl', 
            'XYZ Property Management', 'English', '', 'Initial Submission', '', 
            '', '1', '', '(dd) dd', 'ff, dd, dd, dd', 
            '(901) 867-5309', '', 'bla@bla.com', '2031057', 'Wed Mar 17 2021 00:00:00 GMT-0400 (Eastern Daylight Time)',
            'Yes', '', '', '', '',
            '', ''
          ]
        ],
        [ 
          ['Sun Mar 07 2021 01:02:18 GMT-0500 (Eastern Standard Time)', 'Attorney1 Name - Client1 Name', 'Yes, I am available and have no conflict'],
          ['Mon Mar 08 2021 01:02:18 GMT-0500 (Eastern Standard Time)', 'Attorney2 Name - Client2 Name', 'No...']
        ]
      ]
    ]
  }
  loadSheet(sheetName, formulaColumns, weekIndex, dataIndex) {
    let weekData = this.testData[weekIndex];
    let rowData = weekData[dataIndex];
    let sheet = new SheetClass(sheetName, formulaColumns);
    let rowNumber;
    if (weekIndex === 0) {
      rowNumber = 2;
      sheet.clear();
    } else {
      rowNumber = sheet.getRowCount() + 1;
    }
    for (let i = 0; i < rowData.length; i++) {
      sheet.setRowData(rowNumber++, [rowData[i]]);
    }
  }
  runTests() {
    for (let weekIndex = 0; weekIndex < this.testData.length; weekIndex++) {
      this.loadSheet('Staff List', 1, weekIndex, STAFF_INDEX);
      this.loadSheet('Clients Raw', 0, weekIndex, CLIENT_INDEX);
      this.loadSheet('Availability Raw', 0, weekIndex, AVAILABILITY_INDEX);
    }
  }
  getData(sheetName) {
    let sheet = new SheetClass(sheetName, 0);
    let sheetData = sheet.getAllRows();
    let str = '';
    let quote = '\'';
    let maxRows = Math.min(sheetData.length, 4);
    for (let i = 0; i < maxRows; i++) {
      if (i > 0) {
        str += ', ';
      }
      str += '[ ';
      let rowData = sheetData[i];
      for (let j = 0; j < rowData.length; j++) {
        let cellValue = rowData[j].toString(); 
        if (cellValue.includes('\'')) {
          if (cellValue.includes('"')) {
            console.log('Cannot handle:  ' + rowData[j]);
            continue; 
          } else {
            quote = '"';
          }
        } else {
          quote = '\'';
        }
        if (j > 0) {
          str += ', ';
        }
        str += quote + cellValue + quote;
      }
      str += ' ]';
    }
    return str;
  }
}
function runTests() {
  tester = new Tester();
//  console.log(tester.getData('Confirmations Raw'));
  tester.runTests();
}
