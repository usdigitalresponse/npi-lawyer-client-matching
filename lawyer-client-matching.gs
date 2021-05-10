function onOpen() {
  var menuItems = [
    {name: 'Create Matches', functionName: 'performMatching'},
    {name: 'Email Lawyers', functionName: 'emailLawyers'}
  ];
  SpreadsheetApp.getActive().addMenu('ESP Actions', menuItems);
}

var logger = new Logger();

const UNKNOWN_COURT_DATE = 0; // NPI staffers (at some point in time) entered unknown court dates as zero.
function isUnknownDate(dateInput) {
  const CUTOFF_COURT_DATE = new Date('1900-01-01T00:00:00'); // UNKNOWN_COURT_DATE, when formatted as a date, shows a date in 1899.
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
function hackTime(sData) {
  let headerRow = sData[0];
  let nextCourtDateIndex = headerRow.indexOf(clientColumnMetadata.courtDateColName);
  if (nextCourtDateIndex === -1) {
    console.log('Unable to find column named: ' + clientColumnMetadata.courtDateColName + '. Court dates may be off.');
    return;
  }
  let uniqueIdIndex = headerRow.indexOf(clientColumnMetadata.uniqueIdColName);
  if (uniqueIdIndex === -1) {
    console.log('Unable to find column named: ' + clientColumnMetadata.uniqueIdColName + '. Court dates may be off.');
    return;
  }
  for (let rowIndex = 1; rowIndex < sData.length; rowIndex++) {
    if (!sData[rowIndex][uniqueIdIndex]) {
        // Empty dropdowns in a sheet return non-null data,
        // so use the 'key' column to determine actual number of rows.
      break;
    }
    let strangeDate = sData[rowIndex][nextCourtDateIndex];
    if (!isUnknownDate(strangeDate)) {
      try {
        strangeDate.setHours(12);
      } catch (err) {
        let rowNumber = rowIndex + 1;
        if (strangeDate !== '') {
          console.log('Bad date from eviction sheet at column "' + clientColumnMetadata.nextCourtDateColumn +
                '", row ' + rowNumber + ': "' + strangeDate + '"');
        }
      }
    }
    sData[rowIndex][nextCourtDateIndex] = strangeDate;
  }
}

var clients = null;
var lineSep = String.fromCharCode(10);

function handleUnknownDate(dateInput) {
  const MAX_DATE = new Date(8640000000000000);
  if (isUnknownDate(dateInput)) {
    return MAX_DATE;
  }
  return dateInput;
}

var clientRows = null; // Performance optimization: Read all client rows into memory to avoid lots of network calls.

function compareByCourtDate(firstElement, secondElement) {
  let courtDateIndex = clients.columnIndex(clientColumnMetadata.courtDateColName);
  let firstRow = clientRows[firstElement];
  let secondRow = clientRows[secondElement];
  let firstDate = firstRow[courtDateIndex];
  let secondDate = secondRow[courtDateIndex];
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

class CodeTimer {
  constructor(name) {
    this.name = name;
    this.start = new Date();
  }
  done(newName) {
    let interval = ((new Date()) - this.start) / 1000;
    console.log(this.name + ' took ' + interval + ' seconds.');
    this.name = newName;
    this.start = new Date();
  }
}

class AirTableReader {
  readFromTable() {
    let apiKey = '';
    let tableName = 'Eviction Cases';
    let header = [
      clientColumnMetadata.uniqueIdColName,
      clientColumnMetadata.firstColName,
      clientColumnMetadata.lastColName,
      clientColumnMetadata.emailColName,
      clientColumnMetadata.clientPhoneColName,
      clientColumnMetadata.clientAddressColName,
      clientColumnMetadata.folderColName,
      clientColumnMetadata.landLordNameColName,
      clientColumnMetadata.landlordEmailColName,
      clientColumnMetadata.landlordPhoneColName,
      clientColumnMetadata.landlordAddressColName,
      clientColumnMetadata.courtDateColName,
      clientColumnMetadata.caseNumberColName,
      clientColumnMetadata.clerkConfirmationColName,
      clientColumnMetadata.bulkAgreementColName,
      clientColumnMetadata.rentalApplicationStatusColName,
      clientColumnMetadata.programEligibilityColName,
      clientColumnMetadata.attorneyColName,
      clientColumnMetadata.diagnosticColName
    ];
    let records = [header];
    let recordOffset = 0;
    while (recordOffset !== null) {	
      let url = [
        'https://api.airtable.com/v0/', clientColumnMetadata.airtableBaseID, '/', encodeURIComponent(tableName),
        '?', 'api_key=', apiKey, '&view=', clientColumnMetadata.airtableViewID, '&offset=', recordOffset
      ].join('');
      let response = JSON.parse(UrlFetchApp.fetch(url, {'method' : 'GET'}));
      for (let value1 of response.records.values()) {
        let rowRecord = Array(header.length).fill("");
        for (let propt in value1.fields) {
          let i = header.indexOf(propt);
          if (i > -1) {
            switch(propt) {
              case clientColumnMetadata.clerkConfirmationColName: {
                if (value1.fields[propt]) {
                  rowRecord[i] = 'true';
                }
                break;
              }
              case clientColumnMetadata.courtDateColName: {
                rowRecord[i] = value1.fields[propt];
                break;
              }
              case clientColumnMetadata.attorneyColName: {
                rowRecord[i] = value1.fields[propt];
                break;
              }
              case clientColumnMetadata.landLordNameColName: { // TODO: Need to look it up from record number.
                rowRecord[i] = 'Landlord name to come.';
                break;
              }
              case clientColumnMetadata.bulkAgreementColName: {
                rowRecord[i] = 'Yes';
                break;
              }
              default: {
                rowRecord[i] = value1.fields[propt][0];
                break;
              }
            }
          }
        }
        if (rowRecord[header.indexOf(clientColumnMetadata.uniqueIdColName)] !== '') {
          records.push(rowRecord);
        }
      }
      Utilities.sleep(300);       // Don't trigger Airtable rate limiting.
      if (response.offset) {      // Airtable returns NULL if no more records.
        recordOffset = response.offset;
      } else {
        recordOffset = null;
      }
    }
    return records;
  }
}

class TheApp {
  constructor() {
    this.availabilityColHeader = 'How many cases can you take on this week?';
  }
  clientsColumnIndex(columnName) {
    let index;
    try {
      index = clients.columnIndex(columnName);
    } catch(e) {
      e += ' Is it in https://docs.google.com/spreadsheets/d/' + clientColumnMetadata.currentVersion + '?';
      throw e;
    }
    return index;
  }
  buildSortedClientArray(clients) {
    let t = new CodeTimer('build client array');
    let indexArray = [];
    let confirmationIndex = this.clientsColumnIndex(clientColumnMetadata.clerkConfirmationColName);
    let programEligibilityIndex = this.clientsColumnIndex(clientColumnMetadata.programEligibilityColName);
    let applicationStatusIndex = this.clientsColumnIndex(clientColumnMetadata.rentalApplicationStatusColName);
    let courtDateIndex = this.clientsColumnIndex(clientColumnMetadata.courtDateColName);
    let bulkAgreementIndex = this.clientsColumnIndex(clientColumnMetadata.bulkAgreementColName);
    let attorneyIndex = this.clientsColumnIndex(clientColumnMetadata.attorneyColName);
    let diagnosticIndex = this.clientsColumnIndex(clientColumnMetadata.diagnosticColName);
    let today = new Date();
    clientRows = clients.getAllDataRows();
    let clientIndex;
    for (clientIndex = 0; clientIndex < clientRows.length; clientIndex++) {
      let clientData = clientRows[clientIndex];
      let nextCourtDate = clientData[courtDateIndex];
      let dateOK = (nextCourtDate >= today || isUnknownDate(nextCourtDate));
      let confirmed = clientData[confirmationIndex] !== '';
      let eligible = clientData[programEligibilityIndex] === 'Verified eligible';
      let complete = clientData[applicationStatusIndex] === 'Rental application accepted as complete';
      let notBulk = clientData[bulkAgreementIndex] === ''; 
      let notAssigned = clientData[attorneyIndex] === '';
      if (dateOK && confirmed && eligible && complete && notBulk && notAssigned) {
        indexArray.push(clientIndex);
      } else {
        let diagnostic = 'unknown';
        if (!dateOK) {
          diagnostic = 'dateOK';
        } else if (!confirmed) {
          diagnostic = 'confirmed';
        } else if (!eligible) {
          diagnostic = 'eligible';
        } else if (!complete) {
          diagnostic = 'complete';
        } else if (!notBulk) {
          diagnostic = 'notBulk';
        } else if (!notAssigned) {
          diagnostic = 'notAssigned';
        }
        clientRows[clientIndex][diagnosticIndex] = diagnostic;
      }
    }
    clients.setMultipleRows(2, clientRows);
    t.done('sort');
    indexArray.sort(compareByCourtDate);
    t.done('end');
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
    let clientData = clientRows[sortedClientArray[clientIndex]];
    let caseNumber = clientData[this.clientsColumnIndex(clientColumnMetadata.caseNumberColName)];
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
    match[matches.columnIndex('Client UUID')] = clientData[this.clientsColumnIndex(clientColumnMetadata.uniqueIdColName)];
    match[matches.columnIndex('Client Folder')] = clientData[this.clientsColumnIndex(clientColumnMetadata.folderColName)];
    this.copyFromClientList(match, matches, clientData);
    match[matches.columnIndex('Match Status')] = '';
    match[matches.columnIndex('Pending Timestamp')] = '';
    return match;
  }
  setupAvailabilities(attorneys, emailedMatches) {
    let availabilities = new SheetClass('Ranked Availability');
    let rawAvailabilities = new SheetClass('Availability Raw');
        // Delete all rows in ‘Ranked Availability’. There may have been unused availabilities,
        // but they are from last week (or whenever the last ‘asking for confirmation’ emails went out).
    availabilities.clearData('Name');
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
  copyFromClientList(targetData, targetSheet, clientData) {
    targetData[targetSheet.columnIndex('Client First Name')] = clientData[this.clientsColumnIndex(clientColumnMetadata.firstColName)];
    targetData[targetSheet.columnIndex('Client Last Name')] = clientData[this.clientsColumnIndex(clientColumnMetadata.lastColName)];
    targetData[targetSheet.columnIndex('Client Email')] = clientData[this.clientsColumnIndex(clientColumnMetadata.emailColName)];
    targetData[targetSheet.columnIndex('Client Phone Number')] = clientData[this.clientsColumnIndex(clientColumnMetadata.clientPhoneColName)];
    targetData[targetSheet.columnIndex('Client Address')] = clientData[this.clientsColumnIndex(clientColumnMetadata.clientAddressColName)];
    targetData[targetSheet.columnIndex('Landlord Name')] = clientData[this.clientsColumnIndex(clientColumnMetadata.landLordNameColName)];
    targetData[targetSheet.columnIndex('Landlord Email')] = clientData[this.clientsColumnIndex(clientColumnMetadata.landlordEmailColName)];
    targetData[targetSheet.columnIndex('Landlord Phone Number')] = clientData[this.clientsColumnIndex(clientColumnMetadata.landlordPhoneColName)]; 
    targetData[targetSheet.columnIndex('Landlord Address')] = clientData[this.clientsColumnIndex(clientColumnMetadata.landlordAddressColName)];
    targetData[targetSheet.columnIndex('Case Number')] = clientData[this.clientsColumnIndex(clientColumnMetadata.caseNumberColName)];
    let nextCourtDate = clientData[this.clientsColumnIndex(clientColumnMetadata.courtDateColName)];
    if (isUnknownDate(nextCourtDate)) {
      nextCourtDate = 'Unknown';
    }
    targetData[targetSheet.columnIndex('Next Court Date')] = nextCourtDate;
  }
  createHotList(clientIndex, sortedClientArray) {
    let columnHeaders = [
      'Tenant UID', 'Case Number',	'Next Court Date', 'Client First Name', 'Client Last Name',
      'Client Email', 'Client Phone Number', 'Client Address', 'Landlord Name', 'Landlord Address',
      'Landlord Email', 'Landlord Phone Number'	
    ];
    let hotList = new SheetClass('Hot List', null, columnHeaders);
    let rowsData = [];
    let sourceColIndex = this.clientsColumnIndex(clientColumnMetadata.uniqueIdColName);
    let targetColIndex = hotList.columnIndex('Tenant UID');
    for (; clientIndex < sortedClientArray.length; clientIndex++) {
      let client = [];
      let clientData = clientRows[sortedClientArray[clientIndex]];
      client[targetColIndex] = clientData[sourceColIndex];
      this.copyFromClientList(client, hotList, clientData);
      rowsData.push(client);
    }
    hotList.setMultipleRows(2, rowsData);
  }
  doMatching() {
    let t1 = new CodeTimer('new SheetClass');
    clients = new SheetClass('Clients Raw');
    if (clientColumnMetadata.airtableBaseID) {
      clients.load((new AirTableReader().readFromTable()));
    } else {
      clients.cloneSheet(clientColumnMetadata.currentVersion, 'Client List', hackTime);
    }
    t1.done('buildSortedClientArray');
    let sortedClientArray = this.buildSortedClientArray(clients);
    t1.done('pre-match');

    if (sortedClientArray.length === 0) {
      let totalClients = clients.getRowCount() - 1;
      let msg = 'No clients found that can be matched (out of ' + totalClients  + ' total).';
      logger.logAndAlert('Warning', msg);
      return;
    }
    let attorneys = new SheetClass('Staff List');
    let emailedMatches = new SheetClass('Emailed Matches');
    let availabilities = this.setupAvailabilities(attorneys, emailedMatches);
//    this.updateStaff(attorneys); // Until the Google Form for adding attorneys is enabled.
    let matches = new SheetClass('Created Matches');
    matches.clearData('Case Number');

    let lastAvailabilitiesIndex = availabilities.getRowCount();
    if (lastAvailabilitiesIndex < 2) {
      let msg = 'No attorneys found to match to.';
      logger.logAndAlert('Warning', msg);
      return;
    }
    let nextMatchIndex = 2;
    let availabilityIndex = 2;
    let d = new Date();
    let clientIndex;
    t1.done('match');
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
    t1.done('hotlist');
    this.createHotList(clientIndex, sortedClientArray);
    nextMatchIndex -= 2;
    let leftOver = sortedClientArray.length - nextMatchIndex;
    let msg = 'Matched ' + nextMatchIndex + ' clients. ' + leftOver + ' clients not matched.';
    t1.done('end');
    logger.logAndAlert('Info', msg);
    this.sendStatusEmail(msg);
  }
  sendStatusEmail(msg) {
    MailApp.sendEmail({
      to: 'christopher@mscera.org, usdr@mscera.org, steve@npimemphis.org',
      subject: msg,
      htmlBody: '.'
    });
  }
  performMatching() {
    try {
      doMatching();
    } catch(err) {
      logger.logAndAlert('performMatching: catch: ', err);
    }
  }
  doEmailLawyers() {
    let d = new Date();
    let newCaseCount = 0;
    let emailedMatches = new SheetClass('Emailed Matches');
    let awaitingConfirmation = new SheetClass('Awaiting Confirmation');
    awaitingConfirmation.clearData('Attorney Name - Client Name');
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
    let msg = 'Emailed ' + newCaseCount + ' new cases.';
    logger.logAndAlert('Info', msg);
    this.sendStatusEmail(msg);
  }
  emailLawyers() {
    try {
      this.doEmailLawyers();
    } catch(err) {
      logger.logAndAlert('emailLawyers: catch: ', err);
    }
  }
}

theApp = new TheApp();
function performMatching() { theApp.performMatching(); }
function emailLawyers() { theApp.emailLawyers(); }
function doMatching() { theApp.doMatching(); }
function doAll() { performMatching(); emailLawyers(); }

/* Uncomment and run only *once* after creating (or copying) Google Sheet.
function createTrigger() {
  try {
    ScriptApp.newTrigger("doAll")
      .timeBased()
      .atHour(12)
      .onWeekDay(ScriptApp.WeekDay.FRIDAY)
      .inTimezone("America/Chicago")
      .create();
  } catch(err) {
    (new Logger()).logAndAlert('function askForAvailability: catch: ', err);
  }
}
*/
