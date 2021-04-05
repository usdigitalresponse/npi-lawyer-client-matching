class ClientColumnMetadata {
  constructor() {
    this.v1 = '1vnUVqjwj-u6Wn2v4rhBZN5qvfic6Pa7prLMMLGElBzo';
    this.v2 = '1npa0evM4ifsKzEYUXjgVOiy9dN0AAT2Shmov9bGsAJk';
    this.currentVersion = this.v1;
    if (this.currentVersion === this.v1) {
      this.nextCourtDateColumn = 'T';
      this.uniqueIdColumn = 'D';
    } else {
      this.nextCourtDateColumn = 'T';
      this.uniqueIdColumn = 'A';
    }
  }
}

var clientColumnMetadata = new ClientColumnMetadata();
