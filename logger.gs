// Google Javascript isn't ES6,so no support for 'super' keyword. Thus the 'has-a' relationship. :(
class BaseSheetClass {
  constructor(name, workbookId) {
    this.subSheet = new SheetClass(name, workbookId);
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
  constructor(sheetName, workbookId) {
    try {
      if (!sheetName) {
        sheetName = 'Do NOT Edit - Log';
      }
      this.logSheet = new BaseSheetClass(sheetName, workbookId);
    } catch(err) {
      console.log('Logger constructor exception: ' + err);
      this.logSheet = null;
    }
  }
  showAlert(title, msg) {
    try {
      let ui = SpreadsheetApp.getUi();
      ui.alert(title, msg, ui.ButtonSet.OK);
    } catch(err) {
      this.writeLogLine(['showAlert catch', err]);
      console.log(title + ': ' + msg);
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
    this.writeLogLine([title, msg]);
    this.showAlert(title, msg);
  }
}
