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
    this.uniqueIdColName = 'Tenant UID';
    this.clientPhoneColName = 'Phone';
    this.folderColName = 'Folder';
    this.clientAddressColName = 'Tenant Address';
    this.landLordNameColName = 'Landlord Name';
    this.landlordEmailColName = 'Landlord Email';
    this.landlordPhoneColName = 'Landlord Phone';
    this.landlordAddressColName = 'Landlord Address';
    this.landlordPaymentStatus = null; // 'Landlord Payment Status' + lineSep + 'manual';
  }
}

var clientColumnMetadata = new ClientColumnMetadata();
