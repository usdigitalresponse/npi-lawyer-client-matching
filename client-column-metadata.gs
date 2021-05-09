class ClientColumnMetadata {
  constructor() {
    const lineSep = String.fromCharCode(10);
    let deprecatedBaseID = 'app0jpIprz0I1fOmP';
    let deprecatedViewID = 'viwEXNtNLXdDgVn4R';
    let era_liveDatabaseBaseId = 'appYN8z5f60xC0XRE';
    let era_liveDatabaseViewID = 'viwCrwktzudsRGwzG';
    this.airtableBaseID = null;
    this.airtableViewID = null;
    if (this.airtableBaseID) {
      this.currentVersion = null;
      this.courtDateColName = 'Confirmed Court Date';
      this.caseNumberColName = 'Confirmed Case Number';
      this.uniqueIdColName = 'UID';
      this.landlordPaymentStatus = 'Landlord Payment Status';
      this.clerkConfirmationColName = 'Clerk Confirmation';
      this.matchStatusColName = 'Match Status';
      this.bulkAgreementColName = 'Bulk agreement';
    } else {
      this.currentVersion = '1npa0evM4ifsKzEYUXjgVOiy9dN0AAT2Shmov9bGsAJk'; // v2
      this.courtDateColName = 'Confirmed Court Date' + lineSep + 'manual';
      this.caseNumberColName = 'Eviction Case Number';
      this.uniqueIdColName = 'Unique ID';
      this.landlordPaymentStatus = null; // 'Landlord Payment Status' + lineSep + 'manual';
      this.clerkConfirmationColName = 'Clerk Confirmation' + lineSep + 'manual';
      this.matchStatusColName = 'Match Status' + lineSep + ' auto - Pending, Confirmed, Denied' + lineSep + 'manual for Reassigned';
      this.bulkAgreementColName = 'Associated with Bulk Agreement?';
    }
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
  }
}
 
var clientColumnMetadata = new ClientColumnMetadata();
