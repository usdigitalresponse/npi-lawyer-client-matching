class GetMap {
  getMap(sheetName, keyColumnName, valueColumnName) {
    let sheet = new SheetClass(sheetName);
    let theMap = new Map();
    let rows = sheet.getAllDataRows();
    let keyIndex = sheet.columnIndex(keyColumnName); 
    let valueIndex = sheet.columnIndex(valueColumnName); 
    for (let i = 0; i < rows.length; i++) {
      let k = rows[i][keyIndex];
      let v = rows[i][valueIndex];
      theMap.set(k, v);
    }
    return theMap;
  }
}
function checkForBlank() {
  for (let name of ['Emailed Matches', 'Confirmed Matches']) {
    console.log(name, ': total rows: ' + (new SheetClass(name)).getRowCount());
  }
}