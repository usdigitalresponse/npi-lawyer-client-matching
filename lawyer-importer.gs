class LawyerImporter {
  updateBooleans(rows) {
    let headerRow = rows[0];
    let malsIndex = headerRow.indexOf('MALS?');
    let spanishIndex = headerRow.indexOf('Spanish?');
    for (let rowIndex = 1; rowIndex < rows.length; rowIndex++) {
      if (rows[rowIndex][malsIndex]) {
        rows[rowIndex][malsIndex] = 'Yes';
      } else {
        rows[rowIndex][malsIndex] = 'No';
      }
      if (rows[rowIndex][spanishIndex]) {
        rows[rowIndex][spanishIndex] = 'Yes';
      } else {
        rows[rowIndex][spanishIndex] = 'No';
      }
    }
  }
  importAttorneys() {
    let apiKey = '';
    let tableName = 'Attorneys - ERA';
    let header = [
      'Timestamp',
      'FirstName',
      'LastName',
      'Attorney Email',
      'Type',
      'Spanish?',
      'MALS?',
      'Attorney Name'
    ];
    let rows = (new AirTableImporter().readFromTable(
                    apiKey, clientColumnMetadata.airtableBaseID,
                    'viwoZkM0piORNJIkc', tableName, header, 'Attorney Name'));
    this.updateBooleans(rows);  
    rows[0] = [
      'Timestamp',
      'FirstName',
      'LastName',
      'Email',
      'Type',
      'Spanish?',
      'MALS?',
      'Name'
    ];
    let staffList = new SheetClass('Staff List'); 
    staffList.clearData();
    staffList.setMultipleRows(1, rows);
  }
}

function doTest() {
  (new LawyerImporter()).importAttorneys();
}