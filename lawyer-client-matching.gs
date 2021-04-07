function onOpen() {
  var menuItems = [
    {name: 'Create Matches', functionName: 'performMatching'},
    {name: 'Email Lawyers', functionName: 'emailLawyers'}
  ];
  SpreadsheetApp.getActive().addMenu('ESP Actions', menuItems);
}

var logger = new Logger();

var clients = new SheetClass('Clients Raw');
var lineSep = String.fromCharCode(10);

function isUnknownDate(dateInput) {
  const UNKNOWN_COURT_DATE = 0;
  const CUTOFF_COURT_DATE = new Date('1900-01-01T00:00:00');
  if (dateInput === '') {
    return true;
  }
  if (dateInput === UNKNOWN_COURT_DATE) {
    return true;
  }
  if (dateInput < CUTOFF_COURT_DATE) {
    return true;
  }
  return false;
}
function handleUnknownDate(dateInput) {
  const MAX_DATE = new Date(8640000000000000);
  if (isUnknownDate(dateInput)) {
    return MAX_DATE;
  }
  return dateInput;
}

function compareByCourtDate(firstElement, secondElement) {
  let courtDateIndex = clients.columnIndex(clientColumnMetadata.courtDateColName);
  let firstRow = clients.getRowData(firstElement);
  let secondRow = clients.getRowData(secondElement);
  let firstDate = firstRow[0][courtDateIndex];
  let secondDate = secondRow[0][courtDateIndex];
  firstDate = handleUnknownDate(firstDate);
  secondDate = handleUnknownDate(secondDate);
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
    let programEligibilityIndex = clients.columnIndex(clientColumnMetadata.programEligibilityColName);
    let applicationStatusIndex = clients.columnIndex(clientColumnMetadata.rentalApplicationStatusColName);
    let courtDateIndex = clients.columnIndex(clientColumnMetadata.courtDateColName);
    let bulkAgreementIndex = clients.columnIndex('Associated with Bulk Agreement?');
    let today = new Date();
    let lastClientIndex = clients.getRowCount();
    let clientIndex;
    for (clientIndex = 2; clientIndex <= lastClientIndex; clientIndex++) {
      let clientData = clients.getRowData(clientIndex)[0];
      let nextCourtDate = clientData[courtDateIndex];
      let dateOK = (nextCourtDate >= today || isUnknownDate(nextCourtDate));
      if (dateOK &&
          clientData[confirmationIndex] === 'Yes' &&
          clientData[programEligibilityIndex] === 'Verified eligible' &&
          clientData[applicationStatusIndex] === 'Rental application accepted as complete' &&
          clientData[bulkAgreementIndex] !== 'Yes') {
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
    let caseNumber = clientData[clients.columnIndex(clientColumnMetadata.caseNumberColName)];
    if (emailedMatches.lookupRowIndex('Case Number', caseNumber) != -1) {
      let msg = 'Case: ' + caseNumber + ' has already been emailed, skipping it.';
      logger.writeLogLine([msg]);
      return false;
    }
    let attorneyName = availabilityData[availabilities.columnIndex('Name')];
    let rowIdx = attorneys.lookupRowIndex('Name', attorneyName);
    if (rowIdx === -1) {
      logger.writeLogLine(['Unknown attorney name: "' + attorneyName + '". Skipping it.']);
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
    match[matches.columnIndex('Client First Name')] = clientData[clients.columnIndex(clientColumnMetadata.firstColName)];
    match[matches.columnIndex('Client Last Name')] = clientData[clients.columnIndex(clientColumnMetadata.lastColName)];
    match[matches.columnIndex('Client Email')] = clientData[clients.columnIndex(clientColumnMetadata.emailColName)];
    match[matches.columnIndex('Client UUID')] = clientData[clients.columnIndex(clientColumnMetadata.uniqueIdColName)];
    match[matches.columnIndex('Client Folder')] = clientData[clients.columnIndex(clientColumnMetadata.folderColName)];
    match[matches.columnIndex('Client Phone Number')] = clientData[clients.columnIndex(clientColumnMetadata.clientPhoneColName)];
    match[matches.columnIndex('Client Address')] = clientData[clients.columnIndex(clientColumnMetadata.clientAddressColName)];
    match[matches.columnIndex('Landlord Name')] = clientData[clients.columnIndex(clientColumnMetadata.landLordNameColName)];
    match[matches.columnIndex('Landlord Email')] = clientData[clients.columnIndex(clientColumnMetadata.landlordEmailColName)];
    match[matches.columnIndex('Landlord Phone Number')] = clientData[clients.columnIndex(clientColumnMetadata.landlordPhoneColName)]; 
    match[matches.columnIndex('Landlord Address')] = clientData[clients.columnIndex(clientColumnMetadata.landlordAddressColName)];
    match[matches.columnIndex('Case Number')] = clientData[clients.columnIndex(clientColumnMetadata.caseNumberColName)];
    let nextCourtDate = clientData[clients.columnIndex(clientColumnMetadata.courtDateColName)];
    if (isUnknownDate(nextCourtDate)) {
      nextCourtDate = 'Unknown';
    }
    match[matches.columnIndex('Next Court Date')] = nextCourtDate;
    match[matches.columnIndex('Match Status')] = '';
    match[matches.columnIndex('Pending Timestamp')] = '';
    return match;
  }
  setupAvailabilities(attorneys, emailedMatches) {
    let availabilities = new SheetClass('Ranked Availability');
    let rawAvailabilities = new SheetClass('Availability Raw');
        // Delete all rows in ‘Ranked Availability’. There may have been unused availabilities,
        // but they are from last week (or whenever the last ‘asking for confirmation’ emails went out).
    availabilities.clear();
        // Copy from ‘Availability Raw’ all rows timestamped since the most recent email went out.
        // Assumes emailedMatches rows stay in Timestamp order.
    let lastEmailed = emailedMatches.getRowData(emailedMatches.getRowCount());
    let lastEmailedDate = lastEmailed[0][emailedMatches.columnIndex('Timestamp')];
    let nextRowNumber = 2;
    let iter = new SheetRowIterator(rawAvailabilities);
    let raw;
    while (raw = iter.getNextRow()) {
      if (lastEmailedDate < raw[rawAvailabilities.columnIndex('Timestamp')]) {
        raw[availabilities.columnIndex('Type')] = '';
        raw[availabilities.columnIndex('Type Rank')] = '';
        availabilities.setRowData(nextRowNumber++, [raw]);
      }
    }
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
    let emailedMatches = new SheetClass('Emailed Matches');
    let availabilities = this.setupAvailabilities(attorneys, emailedMatches);
//    this.updateStaff(attorneys); // Until the Google Form for adding attorneys is enabled.
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
    clients.cloneSheet(clientColumnMetadata.currentVersion, 'Client List');
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
        logger.writeLogLine([msg]);
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
