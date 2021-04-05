class ClientColumnMetadata {
  constructor() {
    const lineSep = String.fromCharCode(10);
    this.v1 = '1vnUVqjwj-u6Wn2v4rhBZN5qvfic6Pa7prLMMLGElBzo';
    this.v2 = '1npa0evM4ifsKzEYUXjgVOiy9dN0AAT2Shmov9bGsAJk';
    // Must manually copy/paste headers from source client sheet into 'Client Raw' tab when changing this.
    this.currentVersion = this.v1;
    if (this.currentVersion === this.v1) {
      this.nextCourtDateColumn = 'T';
      this.uniqueIdColumn = 'D';
      this.rentalApplicationStatusColName = 'Rental Assistance Application Status' + lineSep + 'auto & manual';
      this.courtDateColName = 'Court Date' + lineSep + 'auto';
      this.programEligibilityColName = 'Program Eligibility ' + lineSep + 'auto';
      this.caseNumberColName = 'Case Number' + lineSep + 'auto';
      this.firstColName = 'First' + lineSep + 'auto';
      this.lastColName = 'First' + lineSep + 'auto';
      this.emailColName = 'Email' + lineSep + 'auto';
      this.uniqueIdColName = 'Unique ID' + lineSep + 'auto';
      this.clientPhoneColName = 'Phone' + lineSep + 'auto';
      this.folderColName = 'Folder' + lineSep + 'auto';
      this.clientAddressColName = 'Address'  + lineSep + 'auto';
      this.landLordNameColName = 'Folder' + lineSep + 'auto';
      this.landlordEmailColName = 'Landlord Email'  + lineSep + 'auto';
      this.landlordPhoneColName = 'Landlord Phone'  + lineSep + 'auto';
      this.landlordAddressColName = 'Landlord Address'  + lineSep + 'auto';
    } else {
      this.nextCourtDateColumn = 'AN';
      this.uniqueIdColumn = 'A';
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
      this.landLordNameColName = 'Folder';
      this.landlordEmailColName = 'Landlord Email';
      this.landlordPhoneColName = 'Landlord Phone';
      this.landlordAddressColName = 'Landlord Address';
    }
  }
}

var clientColumnMetadata = new ClientColumnMetadata();
