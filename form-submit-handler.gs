class OnSubmitHandler {
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
  updateConfirmed(range) {
    const confirmationsRaw = new SheetClass('Confirmations Raw');
    let emailedMatches = new SheetClass('Emailed Matches');
    let confirmedMatches = new SheetClass('Confirmed Matches');
    const colNames = [
      'Timestamp', 'Lawyer First Name', 'Lawyer Last Name',
      'Lawyer Email', 'Client First Name', 'Client Last Name', 'Client Email', 'Client UUID',
      'Client Folder', 'Client Phone Number', 'Client Address', 'Landlord Name',  'Landlord Email',
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
  handleSubmit(e) {
    try {
      const range = e.range;
      if (range.getSheet().getName() === 'Confirmations Raw') {
        this.updateConfirmed(range);
      }
    } catch(e) {
      console.log('handleSubmit catch: ' + e);
    }
  }
}
var onSubmitHandler = new OnSubmitHandler();
function onSubmitForm(e) { onSubmitHandler.handleSubmit(e); }
