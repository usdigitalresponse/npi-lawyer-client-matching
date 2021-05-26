function triggerEmail() {
  /*
  Timestamp	Lawyer First Name	Lawyer Last Name	Lawyer Email	Client First Name	Client Last Name	Client Email	
  Client UUID	Client Folder	Client Phone Number	Client Address	
  Landlord Name	Landlord Email	Landlord Phone Number	Landlord Address	
  Case Number	Next Court Date	Match Status	Pending Timestamp																													
  */
  let options = { year: 'numeric', month: 'numeric', day: 'numeric',
                  hour: '2-digit', minute: '2-digit', second: '2-digit' }
  let d = (new Date()).toLocaleDateString("en-US", options)
  d = d.substr(0, d.length - 3)
  d = d.replace(',', '')
  let record = [
    d,
    'Chris',
    'Keith',
    'chris.keith@gmail.com',
    'client FN',
    'client LN',
    'chris.keith@gmail.com',
    '3-Test ID',
    'https://drive.google.com/drive/folders/test',
    '111 111 1111',
    'client address',
    'some landlord',
    'landlord@fubar.com',
    '222 222 2222',
    'landlord address',
    '9999999',
    'unknown',
    '',
    ''
  ]
  let s = new SheetClass('Emailed Matches')
  let nextRow = s.getRowCount() + 1
  s.setRowData(nextRow, [record])
}
