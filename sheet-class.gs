const maxColumns = 200;
class SheetClass {
  constructor(name, workbookId) {
    this.name = name;
    if (workbookId) {
      this.sheet = SpreadsheetApp.openById(workbookId).getSheetByName(name);
    } else {
      this.sheet = SpreadsheetApp.getActive().getSheetByName(name);
    }
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
  columnIndexFromLetter(colId) {
    let highOrderVal = 0;
    let lowOrderIndex = 0;
    if (colId.length > 1) {
      highOrderVal = colId.charCodeAt(0) - 'A'.charCodeAt(0);
      lowOrderIndex = 1;
    }
    return highOrderVal + colId.charCodeAt(lowOrderIndex) - 'A'.charCodeAt(0);
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
  hackTime(sData) {
    if (this.name === 'Clients Raw') {
      let nextCourtDateIndex = this.columnIndexFromLetter('T');
      let uniqueIdIndex = this.columnIndexFromLetter('D');
      for (let rowIndex = 1; rowIndex < sData.length; rowIndex++) {
        if (!sData[rowIndex][uniqueIdIndex]) {
            // Empty dropdowns in a sheet return non-null data,
            // so use the 'key' column to determine actual number of rows.
          break; 
        }
        let strangeDate = sData[rowIndex][nextCourtDateIndex];
        try {
          strangeDate.setHours(12);
        } catch (err) {
          let rowNumber = rowIndex + 1;
          if (strangeDate !== '') {
            console.log('Bad date at column "T", row ' + rowNumber + ': "' + strangeDate + '"');
          }
        }
        sData[rowIndex][nextCourtDateIndex] = strangeDate;
      }
    }
  }
  cloneSheet(sourceId, sourceSheetName) {
    let sourceWorkbook = SpreadsheetApp.openById(sourceId);
    let sourceSheet = sourceWorkbook.getSheetByName(sourceSheetName);
    let fullRange = sourceSheet.getDataRange();
    let rangeSpec = fullRange.getA1Notation();
    let sData = fullRange.getValues();
    this.sheet.clear({contentsOnly: true});
    this.hackTime(sData);
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