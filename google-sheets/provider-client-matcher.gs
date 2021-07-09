class ProviderClientMatcher {
  constructor() {
//    this.logger = new Logger();
  }
  loadProviderData(providerData) {
    let providersByService = {};
    let headers = providerData.headerData[0];
    for (let i = 2; i < headers.length; i++) {
      let serviceName = headers[i];
      providersByService[serviceName] = [];
    }
    let providerIter = new SheetRowIterator(providerData);
    let columnIndex = providerData.columnIndex('What is the name of the organization?');
    let services;
    while (services = providerIter.getNextRow()) {
      let providerName = services[columnIndex];
      for (let i = 2; i < headers.length; i++) {
        if (services[i] === 1) {
          let serviceName = headers[i];
          providersByService[serviceName].push(providerName);
        }
      }
    }
    return providersByService;
  }
  loadClientData(clientData) {
    let providersByService = {};
    return providersByService;
  }
  matchThem() {
    console.log('matchThem TBD');
  }
  doMatching() {
    const providerTabName = 'Services provided - categorized';
    const providerWorkbookId = '1BHlfgXgA-Ej3iRwirMAm7kipAGKKSr3gnD95ktyReXM';
    const clientTabName = 'Clients';
    let providerData = new SheetClass(providerTabName, providerWorkbookId);
    let clientData = new SheetClass(clientTabName);
    let providersByService = this.loadProviderData(providerData);
    this.loadClientData(clientData, providersByService);
    this.matchThem();
  }
}

function doMatching() {
    (new ProviderClientMatcher()).doMatching();
}
