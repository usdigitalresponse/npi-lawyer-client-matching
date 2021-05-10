class ClientColumnMetadata {
  constructor() {
    let era_liveDatabaseBaseId = 'appYN8z5f60xC0XRE';
    let era_liveDatabaseViewID = 'viwCrwktzudsRGwzG';
    this.airtableBaseID = era_liveDatabaseBaseId;
    this.airtableViewID = era_liveDatabaseViewID;
    this.courtDateColName = 'Confirmed Court Date';
    this.caseNumberColName = 'Confirmed Case #';
    this.uniqueIdColName = 'UID';
    this.clerkConfirmationColName = 'Clerk Confirmation';
    this.matchStatusColName = 'Match Status';
    this.bulkAgreementColName = 'Bulk agreement entries';
    this.rentalApplicationStatusColName = 'Rental Assistance Application Status';
    this.programEligibilityColName = 'Program Eligibility';
    this.landlordEmailColName = 'Landlord Email';
    this.landLordNameColName = 'Landlord Name';
    this.landlordPhoneColName = 'Landlord Phone';
    this.landlordAddressColName = 'Landlord Address';
    this.clientAddressColName = 'Tenant Address';
    this.folderColName = 'Folder';
    this.firstColName = 'First';
    this.lastColName = 'Last';
    this.emailColName = 'Email';
    this.clientPhoneColName = 'Phone';
    this.attorneyColName = 'Attorney';
    this.diagnosticColName = 'Diagnostic';

    this.currentVersion = null; // TODO: remove code that copies from other sheet.
  }
}
 
var clientColumnMetadata = new ClientColumnMetadata();
