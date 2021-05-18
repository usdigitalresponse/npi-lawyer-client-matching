class AirTableImporter {
  getValue(value1) {
      if (Array.isArray(value1)) {
        return this.getValue(value1[0]);
      }
      return value1;
  }
  readFromTable(apiKey, baseId, viewId, tableName, header, keyColumnName) {
    let records = [header];
    let recordOffset = 0;
    while (recordOffset !== null) {	
      let url = [
        'https://api.airtable.com/v0/', baseId, '/', encodeURIComponent(tableName),
        '?', 'api_key=', apiKey, '&view=', viewId, '&offset=', recordOffset
      ].join('');
      let response = JSON.parse(UrlFetchApp.fetch(url, {'method' : 'GET'}));
      for (let value1 of response.records.values()) {
        let rowRecord = Array(header.length).fill("");
        for (let propt in value1.fields) {
          let i = header.indexOf(propt);
          if (i > -1) {
            rowRecord[i] = this.getValue(value1.fields[propt]);
          }
        }
        if (rowRecord[header.indexOf(keyColumnName)] !== '') {
          records.push(rowRecord);
        }
      }
      Utilities.sleep(300);       // Don't trigger Airtable rate limiting.
      if (response.offset) {      // Airtable returns NULL if no more records.
        recordOffset = response.offset;
      } else {
        recordOffset = null;
      }
    }
    return records;
  }
}
