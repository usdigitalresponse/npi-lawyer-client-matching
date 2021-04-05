class ClientColumnMetadata {
  constructor() {
    const lineSep = String.fromCharCode(10);
    this.v1 = '1vnUVqjwj-u6Wn2v4rhBZN5qvfic6Pa7prLMMLGElBzo';
    this.v2 = '1npa0evM4ifsKzEYUXjgVOiy9dN0AAT2Shmov9bGsAJk';
    this.currentVersion = this.v1;
    if (this.currentVersion === this.v1) {
      this.nextCourtDateColumn = 'T';
      this.uniqueIdColumn = 'D';
      this.rentalApplicationStatusColName = 'Rental Assistance Application Status' + lineSep + 'auto & manual';
      this.courtDateColName = 'Court Date' + lineSep + 'auto';
      this.programEligibilityColName = 'Program Eligibility ' + lineSep + 'auto';
    } else {
      this.nextCourtDateColumn = 'AN';
      this.uniqueIdColumn = 'A';
      this.rentalApplicationStatusColName = 'Rental Assistance Application Status';
      this.courtDateColName = 'Confirmed Court Date' + lineSep + 'manual';
      throw 'this.programEligibilityColName = UNKNOWN';
    }
  }
}

var clientColumnMetadata = new ClientColumnMetadata();
