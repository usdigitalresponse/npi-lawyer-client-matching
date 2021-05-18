function getMap(sheetName, columnName) {
  let sheet = new SheetClass(sheetName);
  let theMap = new Map();
  let rows = sheet.getAllDataRows();
  let nameIndex = sheet.columnIndex(columnName); 
  for (let i = 0; i < rows.length; i++) {
    let n = rows[i][nameIndex];
    theMap.set(n, n);
  }
  return theMap;
}

function checkAttorneyEmails() {
  let attorneyNames = getMap('Staff List', 'Name');
  let assignedAttorneys = getMap('Clients Raw', 'Attorney');
  let unknownAttorneys = [];
  for (let name of assignedAttorneys.values()) {
    if (!attorneyNames[name]) {
      unknownAttorneys.push(name);
    }
  }
  console.log(unknownAttorneys);
}
