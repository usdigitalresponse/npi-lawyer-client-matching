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

class OnEditHandler {
  deleteAwaiting(confirmationsRaw, range) {
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
    for (let rn of rowsToDelete) {
      awaitingConfirmation.sheet.deleteRow(rn);
    }
  }
  findEmailedMatch(emailedMatches, attorneyClientId) {
    let names = attorneyClientId.split(' - ');
    let attorneyName = names[0];
    let clientName = names[1];
    let iter = new SheetRowIterator(emailedMatches);
    let rowData;
    while (rowData = iter.getNextRow()) {
      // Unclear where extra spaces are coming from.
      // This code introduces a slight possiblity for a mismatch if names are similar,
      // but we have to live with it.
      // Would be better (if ever possible) to use court case number for the key, which should be unique.
      if (attorneyName.startsWith(rowData[emailedMatches.columnIndex('Lawyer First Name')]) &&
          attorneyName.endsWith(rowData[emailedMatches.columnIndex('Lawyer Last Name')]) &&
          clientName.startsWith(rowData[emailedMatches.columnIndex('Client First Name')]) &&
          clientName.endsWith(rowData[emailedMatches.columnIndex('Client Last Name')])) {
        return iter.nextIndex - 1;
      }
    }
    let msg = '"' + attorneyClientId + '" not found in "Emailed Matches"';
    logger.logAndAlert('Error', msg);
    throw msg;
  }
  updateConfirmed(confirmationsRaw, range) {
    let emailedMatches = new SheetClass('Emailed Matches');
    let confirmedMatches = new SheetClass('Confirmed Matches');
    const colNames = [
      'Timestamp', 'Lawyer First Name', 'Lawyer Last Name',
      'Lawyer Email', 'Client First Name', 'Client Last Name', 'Client Email', 'Client UUID',
      'Client Folder', 'Client Phone Number',	'Client Address',	'Landlord Name',	'Landlord Email',
      'Landlord Phone Number', 'Landlord Address', 'Case Number', 'Next Court Date', 'Match Status'
    ]
    let rowNum = confirmedMatches.getRowCount() + 1;
    const lastRow = range.getLastRow();
    const firstRow = lastRow - range.getHeight() + 1;
    for (let row = firstRow; row <= lastRow; row++) {
      const rawRow = confirmationsRaw.getRowData(row)[0]; 
      const response = rawRow[confirmationsRaw.columnIndex('Do you accept the case?')];
      if (response === 'Yes, I am available and have no conflict') {
        const id = rawRow[confirmationsRaw.columnIndex('Case')];
        const rowNumber = this.findEmailedMatch(emailedMatches, id);
        let sourceData = emailedMatches.getRowData(rowNumber);
        let targetData = [];
        for (let colName of colNames) {
          targetData[confirmedMatches.columnIndex(colName)] = sourceData[0][emailedMatches.columnIndex(colName)];
        }
        targetData[confirmedMatches.columnIndex('Confimed/Denied Timestamp')] = (new Date()).toString();
        targetData[confirmedMatches.columnIndex('Attorney Name - Client Name')] = 'see other columns';
        targetData[confirmedMatches.columnIndex('Do you accept the case?')] = 'Yes, I am available and have no conflict';
        confirmedMatches.setRowData(rowNum++, [targetData]);
      }
    }
  }
  doEdit(range) {
    const confirmationsRaw = new SheetClass('Confirmations Raw');
    this.deleteAwaiting(confirmationsRaw, range);
    this.updateConfirmed(confirmationsRaw, range);
  }
  handleEdit(e) {
    try {
      const range = e.range;
      if (range.getSheet().getName() === 'Confirmations Raw') {
        this.doEdit(range);
      }
    } catch(e) {
      console.log('onEdit catch: ' + e);
      showOKAlert('onEdit catch', e);
    }
  }
  doTest() {
    const confirmationsRaw = new SheetClass('Confirmations Raw');
    this.doEdit(confirmationsRaw.sheet.getRange('A9:C9'));
  }
}
var onEditHandler = new OnEditHandler();

function onEdit(e) {
  onEditHandler.handleEdit(e);
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
      logger.logAndAlert('Warning', 'Sheet: "' + this.name + '" may have more than ' + maxColumns +
                    ' columns. Ignoring columns after: ' + maxColumns + '.');
    }
    this.headerData = headerData;
    this.lastColumn = this.columnLetterFromIndex(headerData.length - 1);
  }
  columnIndex(columnName) {
    let index = this.headerData[0].indexOf(columnName);
    if (index < 0) {
      let msg = 'No column named: "' + columnName + '" in sheet: "' + this.name + '"?';
      logger.logAndAlert('Error', msg);
      throw msg;
    }
    return index;
  }
  columnName(columnIndex) {
    if (columnIndex >= this.headerData[0].length) {
      let msg = 'Column index too big: "' + columnIndex + '" in sheet: "' + this.name + '"?';
      logger.logAndAlert('Error', msg);
      throw msg;
    }
    return this.headerData[0][columnIndex];
  }
  getRowCount() {
    let count = 0;
    // If this turns out to be a performance problem down the road,
    // use the length of a 'key' column instead.
    for (let colIndex = 0; colIndex < this.headerData[0].length; colIndex++) {
      let colLetter = this.columnLetterFromIndex(colIndex);
      let rangeSpec = colLetter + '1:' + colLetter;
      let values = this.sheet.getRange(rangeSpec).getValues();
      count = Math.max(count, values.filter(String).length);
    }
    return count;
  }
  getRowData(rowNumber) {
    let rangeSpec = 'A' + rowNumber + ':' + this.lastColumn + rowNumber;
    try {
      let range = this.sheet.getRange(rangeSpec);
      return range.getValues();
    } catch(err) {
      logger.writeLogLine(['Exception', 'Sheet: "' + this.name + '", range: ' + rangeSpec]);
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
    return -1;
  }
  clear() {
    let rowCount = this.getRowCount();
    if (rowCount > 1) {
      let address = 'A2:' + this.lastColumn + rowCount;
      let range = this.sheet.getRange(address);
      range.clear();
    }
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
  appendRow(data) {
    this.subSheet.sheet.appendRow(data);
  }
}

class Logger {
  constructor() {
    try {
      this.logSheet = new BaseSheetClass('Do NOT Edit - Log');
    } catch(err) {
      console.log('Logger constructor exception: ' + err);
      this.logSheet = null;
    }
  }
  writeLogLine(data) {
    if (this.logSheet) {
      let d = new Date();
      data.unshift(d);
      this.logSheet.appendRow(data);
    }
    console.log(data);
  }
  logAndAlert(title, msg) {
    showAlert(title, msg);
    logger.writeLogLine([title, msg]);
  }
}
var logger = new Logger();

var clients = new SheetClass('Clients Raw');
var lineSep = String.fromCharCode(10);

const UNKNOWN_COURT_DATE = 0;
function compareByCourtDate(firstElement, secondElement) {
  let courtDateIndex = clients.columnIndex('Court Date' + lineSep + 'auto');
  let firstRow = clients.getRowData(firstElement);
  let secondRow = clients.getRowData(secondElement);
  let firstDate = firstRow[0][courtDateIndex];
  let secondDate = secondRow[0][courtDateIndex];

  const MAX_DATE = new Date(8640000000000000);
  if (firstDate === UNKNOWN_COURT_DATE) {
    firstDate = MAX_DATE;
  }
  if (secondDate === UNKNOWN_COURT_DATE) {
    secondDate = MAX_DATE;
  }
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
    let programEligibilityIndex = clients.columnIndex('Program Eligibility ' + lineSep + 'auto');
    let applicationStatusIndex = clients.columnIndex('Rental Assistance Application Status' + lineSep + 'auto & manual');
    let courtDateIndex = clients.columnIndex('Court Date' + lineSep + 'auto');
    let today = new Date();
    let lastClientIndex = clients.getRowCount();
    let clientIndex;
    for (clientIndex = 2; clientIndex <= lastClientIndex; clientIndex++) {
      let clientData = clients.getRowData(clientIndex)[0];
      let nextCourtDate = clientData[courtDateIndex];
      let dateOK = (nextCourtDate >= today || nextCourtDate === UNKNOWN_COURT_DATE);
      if (dateOK &&
          clientData[confirmationIndex] === 'Yes' &&
          clientData[programEligibilityIndex] === 'Verified eligible' &&
          clientData[applicationStatusIndex] === 'Rental application accepted as complete') {
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
        logger.writeLogLine(['Warning', 'No row for attorney in Staff List: "' + uuid + '". Skipping it.']);
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
  getAvailablityIndex(availabilityIndex, lastAvailabilitiesIndex, availabilities) {
    if (availabilityIndex > lastAvailabilitiesIndex) { // Check if no one available at all.
      return -1;
    }
    let availabilityData = availabilities.getRowData(availabilityIndex)[0];
    let availabilityColIndex = availabilities.columnIndex(this.availabilityColHeader);
    while (availabilityData[availabilityColIndex] <= 0) {
      availabilityIndex++;
      if (availabilityIndex > lastAvailabilitiesIndex) {
        return -1;
      }
      availabilityData = availabilities.getRowData(availabilityIndex)[0];
    };
    return availabilityIndex;
  }
  clientCanMatch(clientIndex, sortedClientArray, emailedMatches, availabilities, availabilityData, attorneys) {
    let clientData = clients.getRowData(sortedClientArray[clientIndex])[0];
    let caseNumber = clientData[clients.columnIndex('Case Number' + lineSep + 'auto')];
    if (emailedMatches.lookupRowIndex('Case Number', caseNumber) != -1) {
      let msg = 'Case: ' + caseNumber + ' has already been emailed, skipping it.';
      logger.writeLogLine(msg);
      return false;
    }
    let attorneyName = availabilityData[availabilities.columnIndex('Name')];
    let rowIdx = attorneys.lookupRowIndex('Name', attorneyName);
    if (rowIdx === -1) {
      logger.writeLogLine('Unknown attorney name: "' + attorneyName + '". Skipping it.');
      return false;
    }
    return true;
  }
  createMatch(date, matches, clientData, attorneys, availabilityData, availabilities) {
    let match = [];
    let attorneyName = availabilityData[availabilities.columnIndex('Name')];
    let rowIdx = attorneys.lookupRowIndex('Name', attorneyName);
    let attorneyData = attorneys.getRowData(rowIdx + 1)[0];
    let lawyerNames = attorneyName.split(' ');

    match[matches.columnIndex('Timestamp')] = date;
    match[matches.columnIndex('Lawyer First Name')] = lawyerNames[0];
    match[matches.columnIndex('Lawyer Last Name')] = lawyerNames[1];
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
    let nextCourtDate = clientData[clients.columnIndex('Court Date' + lineSep + 'auto')];
    if (nextCourtDate === UNKNOWN_COURT_DATE) {
      nextCourtDate = 'Unknown';
    }
    match[matches.columnIndex('Next Court Date')] = nextCourtDate;
    match[matches.columnIndex('Match Status')] = '';
    match[matches.columnIndex('Pending Timestamp')] = '';
    return match;
  }
  setupAvailabilities(attorneys) {
    let availabilities = new SheetClass('Ranked Availability');
    let rawAvailabilities = new SheetClass('Availability Raw');
    availabilities.copyFrom('Availability Raw', 'A2:C' + rawAvailabilities.getRowCount());
    this.cleanUpAvailabilities(availabilities, attorneys);
    availabilities.sortSheet('Type Rank', true);
    return availabilities;
  }
  doMatching() {
    let sortedClientArray = this.buildSortedClientArray(clients);
    if (sortedClientArray.length === 0) {
      let msg = 'No clients found with "Clerk Confirmation" set to "Yes", ' +
                'blank "Match Status" and "Program Eligibility" set to "Verified eligible"';
      logger.logAndAlert('Warning', msg);
      return;
    }
    let attorneys = new SheetClass('Staff List');
    let availabilities = this.setupAvailabilities(attorneys);
//    this.updateStaff(attorneys); // Until the Google Form for adding attorneys is enabled.
//    attorneys.sortSheet('FirstName', true); // Until 'length of empty column' bug is verified.
    let emailedMatches = new SheetClass('Emailed Matches');
    let matches = new SheetClass('Created Matches');
    matches.clear();

    let lastAvailabilitiesIndex = availabilities.getRowCount();
    let nextMatchIndex = 2;
    let availabilityIndex = 2;
    let d = new Date();
    let clientIndex;
    for (clientIndex = 0; clientIndex < sortedClientArray.length; clientIndex++) {
      availabilityIndex = this.getAvailablityIndex(availabilityIndex, lastAvailabilitiesIndex, availabilities);
      if (availabilityIndex < 0) {
        break;
      }
      let availabilityData = availabilities.getRowData(availabilityIndex)[0];
      if (this.clientCanMatch(clientIndex, sortedClientArray, emailedMatches,
                              availabilities, availabilityData, attorneys)) {
        let clientData = clients.getRowData(sortedClientArray[clientIndex])[0];
        let match = this.createMatch(d, matches, clientData, attorneys, availabilityData, availabilities);
        matches.setRowData(nextMatchIndex, [match]);
        nextMatchIndex++;
        let availabilityColIndex = availabilities.columnIndex(this.availabilityColHeader);
        availabilities.setCellData(availabilityIndex, this.availabilityColHeader, --availabilityData[availabilityColIndex]);
      }
    }
    nextMatchIndex -= 2;
    let leftOver = sortedClientArray.length - nextMatchIndex;
    let msg = 'Matched ' + nextMatchIndex + ' clients. ' + leftOver + ' clients not matched.';
    logger.logAndAlert('Info', msg);
  }
  performMatching() {
    clients.cloneSheet('1vnUVqjwj-u6Wn2v4rhBZN5qvfic6Pa7prLMMLGElBzo', 'Client List');
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
        let msg = 'Case: ' + newCaseNumber + ' already emailed. Skipping it.';
        logger.writeLogLine(msg);
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
    logger.logAndAlert('Info', 'Emailed ' + newCaseCount + ' new cases.');
  }
}

theApp = new TheApp();
function performMatching() { theApp.performMatching(); }
function emailLawyers() { theApp.emailLawyers(); }
function doMatching() { theApp.doMatching(); }
