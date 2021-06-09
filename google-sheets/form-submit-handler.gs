class OnSubmitHandler {
  constructor() {
    this.logger = new Logger('Do NOT Edit - Log', FormApp.getActiveForm().getDestinationId());
  }
  writeLogLine(messageString) {
    if (this.logger) {
      this.logger.writeLogLine(['OnSubmitHandler', messageString]);
    } else {
      console.log('OnSubmitHandler' + messageString);
    }
  }
  deleteAwaitingConfirmation(attorneyClientId) {
    let awaitingConfirmation = new SheetClass('Awaiting Confirmation', FormApp.getActiveForm().getDestinationId());
    let rowNumber = awaitingConfirmation.lookupRowIndex('Attorney Name - Client Name', attorneyClientId) + 1;
    if (rowNumber > 1) {
      awaitingConfirmation.sheet.deleteRow(rowNumber);
    }
  }
  findEmailedMatch(emailedMatches, attorneyClientId) {
    let names = attorneyClientId.split(' - ');
    let attorneyName = names[0].trim();
    let clientName = names[1].trim();
    let iter = new SheetRowIterator(emailedMatches);
    let rowData;
    while (rowData = iter.getNextRow()) {
      // Unclear where extra spaces are coming from.
      // This code introduces a slight possiblity for a mismatch if names are similar,
      // but we have to live with it.
      // Would be better (if ever possible) to use court case number for the key, which should be unique.
      if (attorneyName.startsWith(rowData[emailedMatches.columnIndex('Lawyer First Name')].trim()) &&
          attorneyName.endsWith(rowData[emailedMatches.columnIndex('Lawyer Last Name')].trim()) &&
          clientName.startsWith(rowData[emailedMatches.columnIndex('Client First Name')].trim()) &&
          clientName.endsWith(rowData[emailedMatches.columnIndex('Client Last Name')].trim())) {
        return iter.nextIndex - 1;
      }
    }
    throw '"' + attorneyClientId + '" not found in "Emailed Matches"';
  }
  updateConfirmed(attorneyClientId, answer) {
    const emailedMatches = new SheetClass('Emailed Matches', FormApp.getActiveForm().getDestinationId());
    const rowNumber = this.findEmailedMatch(emailedMatches, attorneyClientId);
    const confirmedMatches = new SheetClass('Confirmed Matches', FormApp.getActiveForm().getDestinationId());
    const colNames = [
      'Timestamp', 'Lawyer First Name', 'Lawyer Last Name',
      'Lawyer Email', 'Client First Name', 'Client Last Name', 'Client Email', 'Client UUID',
      'Client Folder', 'Client Phone Number', 'Client Address', 'Landlord Name',  'Landlord Email',
      'Landlord Phone Number', 'Landlord Address', 'Case Number', 'Next Court Date', 'Match Status'
    ]
    const sourceData = emailedMatches.getRowData(rowNumber)[0];
    let targetData = [];
    for (let colName of colNames) {
      targetData[confirmedMatches.columnIndex(colName)] = sourceData[emailedMatches.columnIndex(colName)];
    }
    targetData[confirmedMatches.columnIndex('Timestamp')] = (new Date()).toString();
    targetData[confirmedMatches.columnIndex('Confimed/Denied Timestamp')] = targetData[confirmedMatches.columnIndex('Timestamp')];
    targetData[confirmedMatches.columnIndex('Attorney Name - Client Name')] = attorneyClientId;
    targetData[confirmedMatches.columnIndex('Do you accept the case?')] = answer;
    confirmedMatches.setRowData(confirmedMatches.getRowCount() + 1, [targetData]);
  }
  doUpdate(caseId, answer) {
    this.updateConfirmed(caseId, answer);
    this.deleteAwaitingConfirmation(caseId);
  }
  handleSubmit(e) {
    try {
      let caseId = '';
      let answer = ''
      let itemResponses = e.response.getItemResponses();
      for (var j = 0; j < itemResponses.length; j++) {
        var itemResponse = itemResponses[j];
        switch (itemResponse.getItem().getTitle()) {
          case 'Case':
            { caseId = itemResponse.getResponse(); break; }
          case 'Do you accept the case?':
            { answer = itemResponse.getResponse(); break; }
          default:
            { this.writeLogLine('Unknown itemResponse.getItem().getTitle(): ' + itemResponse.getItem().getTitle()); }
        }
      }
      this.doUpdate(caseId, answer);
    } catch(e) {
      this.writeLogLine('handleSubmit catch: ' + e);
    }
  }
}
function onSubmitForm(e) {
      // Works ONLY as long as there is only one trigger.
  let onSubmitHandler = new OnSubmitHandler();
  onSubmitHandler.handleSubmit(e);
}
