class ClientColumnMetadata {
  constructor() {
    const lineSep = String.fromCharCode(10);
    this.currentVersion = '1npa0evM4ifsKzEYUXjgVOiy9dN0AAT2Shmov9bGsAJk'; // v2
    this.rentalApplicationStatusColName = 'Rental Assistance Application Status';
    this.courtDateColName = 'Confirmed Court Date' + lineSep + 'manual';
    this.programEligibilityColName = 'Program Eligibility';
    this.caseNumberColName = 'Eviction Case Number';
    this.firstColName = 'First';
    this.lastColName = 'Last';
    this.emailColName = 'Email';
    this.uniqueIdColName = 'Unique ID';
    this.clientPhoneColName = 'Phone';
    this.folderColName = 'Folder';
    this.clientAddressColName = 'Tenant Address';
    this.landLordNameColName = 'Landlord Name';
    this.landlordEmailColName = 'Landlord Email';
    this.landlordPhoneColName = 'Landlord Phone';
    this.landlordAddressColName = 'Landlord Address';
    this.landlordPaymentStatus = null; // TODO: 'Landlord Payment Status' + lineSep + 'manual';
    this.clerkConfirmationColName = 'Clerk Confirmation' + lineSep + 'manual';
    this.matchStatusColName = 'Match Status' + lineSep + ' auto - Pending, Confirmed, Denied' + lineSep + 'manual for Reassigned';
    this.bulkAgreementColName = 'Associated with Bulk Agreement?';
  }
}

var clientColumnMetadata = new ClientColumnMetadata();
