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
    let servicesByClient = {};
    let headers = clientData.headerData[0];
    for (let i = 1; i < headers.length; i++) {
      let clientName = headers[i];
      servicesByClient[clientName] = [];
    }
    let serviceIter = new SheetRowIterator(clientData);
    let columnIndex = clientData.columnIndex('Service Name');

    let service;
    while (service = serviceIter.getNextRow()) {
      for (let i = 1; i < headers.length; i++) {
        let clientName = headers[i];
        if (service[i] === 1) {
          let serviceName = service[columnIndex];
          servicesByClient[clientName].push(serviceName);
        }
      }
    }
    return servicesByClient;
  }
  matchThem(providersByService, servicesByClient) {
    console.log(providersByService);
    console.log(servicesByClient);
    console.log('matchThem TBD');
  }
  doMatching() {
    const providerTabName = 'Services provided - categorized';
    const providerWorkbookId = '1BHlfgXgA-Ej3iRwirMAm7kipAGKKSr3gnD95ktyReXM';
    const clientTabName = 'Clients';
    let providerData = new SheetClass(providerTabName, providerWorkbookId);
    let clientData = new SheetClass(clientTabName);
    let providersByService = this.loadProviderData(providerData);
    let servicesByClient = this.loadClientData(clientData);
    this.matchThem(providersByService, servicesByClient);
  }
}

function doMatching() {
    (new ProviderClientMatcher()).doMatching();
}
