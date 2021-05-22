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
  import() {
    let apiKey = '';
    let baseId = 'appbhKzcwhje8zJKH';
    let viewId = 'viwM9EapkyV7vAOl9';
    let tableName = 'Attorneys';
    let header = [
      'Timestamp',
      'FirstName',
      'LastName',
      'Email',
      'Type',
      'Spanish?',
      'MALS?',
      'Name'
    ];
    let rows = (new AirTableImporter().readFromTable(
                      apiKey, baseId, viewId, tableName, header, 'Email'));
    this.updateBooleans(rows);  
    (new SheetClass('Staff List')).setMultipleRows(1, rows);
  }
}

function doTest() {
  (new LawyerImporter()).import();
}