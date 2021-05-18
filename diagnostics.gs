function getMap(sheetName, keyColumnName, valueColumnName) {
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

function checkAttorneyEmails() {
  let attorneyNames = getMap('Staff List', 'Name', 'Name');
  let assignedAttorneys = getMap('Clients Raw', 'Attorney', 'UID');
  let unknownAttorneys = {};
  for (let name of assignedAttorneys.keys()) {
    if (!attorneyNames[name]) {
      unknownAttorneys[name] = assignedAttorneys.get(name);
    }
  }
  console.log(unknownAttorneys);
}
