// ----------------------- code for automated testing
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
