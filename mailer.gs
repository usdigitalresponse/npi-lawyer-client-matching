class Mailer {
  constructor() {
    this.availabilityBody = '<p><span style="font-weight: 400;">Hello,</span></p>' +
'<p><span style="font-weight: 400;">You are receiving this email because you volunteered your time to support the Eviction Settlement Program, providing legal support on eviction cases in Shelby County. If you are available to take on a new case next week please let us know by </span><strong>filling out</strong><a href="https://docs.google.com/forms/d/1z_Bfddz4XgTsUXIgyW_GCcsyQD0PH2lMkI8Fkyt8AH0/viewform?edit_requested=true"> <strong>this form</strong></a><strong> by the end of the day Thursday</strong><span style="font-weight: 400;">.</span></p>' +
'<p><strong>Background:</strong></p>' +
'<ul>' +
'<li style="font-weight: 400;" aria-level="1"><span style="font-weight: 400;">This is a weekly reminder, asking you to submit your availability for the following week</span><span style="font-weight: 400;">.</span></li>' +
'<li style="font-weight: 400;" aria-level="1"><span style="font-weight: 400;">Volunteers should expect each case to take approximately 5 hours a week, and last approximately two weeks.</span></li>' +
'<li style="font-weight: 400;" aria-level="1"><span style="font-weight: 400;">If you have availability, you may receive an email pairing you with a potential client. Confirming availability </span><em><span style="font-weight: 400;">does not</span></em><span style="font-weight: 400;"> bind you to a case - you will have the opportunity to do a conflict check and confirm your availability before accepting a case. If you choose to accept a case, you will be given access to the necessary client details to proceed.</span></li>' +
'</ul>' +
'<p><strong>What you need to do:</strong></p>' +
'<ul>' +
'<li style="font-weight: 400;" aria-level="1"><span style="font-weight: 400;">Eviction cases move quickly. The sooner you submit availability with this online form, the faster we can find a potential match.</span></li>' +
'<li style="font-weight: 400;" aria-level="1"><span style="font-weight: 400;">If you are not available this week, you do not have to fill out the form. We&rsquo;ll email every week to check in on availability to take on new clients.</span></li>' +
'<li style="font-weight: 400;" aria-level="1"><span style="font-weight: 400;">Once you receive a case assignment, please review the details and confirm that you are able to take on the case.</span></li>' +
'</ul>' +
'<p><span style="font-weight: 400;">Thank you for your valuable work and support in helping preserve our neighbors and communities in Shelby County.</span></p>' +
'<p><strong>Neighborhood Preservation Inc.</strong></p>' +
'<p><a href="http://npimemphis.org/"><span style="font-weight: 400;">npimemphis.org</span></a></p>' +
'<p><a href="https://www.facebook.com/MemphisFightsBlight/"><span style="font-weight: 400;">Facebook</span></a><span style="font-weight: 400;"> |</span><a href="https://www.instagram.com/npimemphis/"> <span style="font-weight: 400;">Instagram</span></a><span style="font-weight: 400;"> |</span><a href="https://twitter.com/NPIMemphis"> <span style="font-weight: 400;">Twitter</span></a></p>' +
'<p><em><span style="font-weight: 400;">Please note, esp@npimemphis.org is an automated messaging address, and responses directly to this email address will not be read. If you have questions about this email, would like to update your volunteer contact information, please reach out to </span></em><em><span style="font-weight: 400;">Steve Barlow</span></em><em><span style="font-weight: 400;">.</span></em></p>'
  }
  doMail() {
    let newStaffList = new SheetClass('Staff List');
    let arrayOfAddresses = newStaffList.getColumnData('Email');
    let flatArray = [];
    for (let i = 0; i < arrayOfAddresses.length; i++) {
      flatArray.push(arrayOfAddresses[i][0]);
    }
    const CHUNK_SIZE = 45; // leave room for up to 5 'to:' and 'cc:' addresses.
    for (let i = 0; i < flatArray.length; i += CHUNK_SIZE) {
      let temparray = flatArray.slice(i, i + CHUNK_SIZE);
      MailApp.sendEmail({
        to: 'usdr@mscera.org',
        bcc: temparray.join(','),
        subject: 'Eviction Settlement Program (ESP) - Are you available to volunteer next week?',
        htmlBody: this.availabilityBody
      });
    }
    let msg = 'Emailed ' + flatArray.length + ' attorneys asking for their availability';
    let statusEmailAddresses = 'christopher@mscera.org, usdr@mscera.org, steve@npimemphis.org';
    MailApp.sendEmail({
      to: statusEmailAddresses,
      subject: msg,
      htmlBody: this.availabilityBody
    });
    (new Logger()).writeLogLine([msg]);
  }
}

function askForAvailability() {
  try {
    (new Mailer()).doMail();
  } catch(err) {
    (new Logger()).logAndAlert('function askForAvailability: catch: ', err);
  }
}
/* Uncomment and run only *once* after creating (or copying) Google Sheet.
function createTrigger() {
  try {
    ScriptApp.newTrigger("askForAvailability")
      .timeBased()
      .atHour(8)
      .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
      .inTimezone("America/Chicago")
      .create();
  } catch(err) {
    (new Logger()).logAndAlert('function askForAvailability: catch: ', err);
  }
}
*/
