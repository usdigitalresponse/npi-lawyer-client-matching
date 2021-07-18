class ProviderClientMatcher {

  // Returns a object containing a list of services.
  // Under each service is a list of providers who can provide that service.
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

  // Returns a object containing a list of clients.
  // Under each client is a list of services that they need.
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

  // Returns a pseudo-random integer between zero and max - 1
  // (including zero and max - 1) 
  getRandomInt(max) {
    return Math.floor(Math.random() * max);
  }

  // Create rows containing client name, service name and provider.
  matchThem(providersByService, servicesByClient) {
    let matchDataRows = [];
    for (let client in servicesByClient) {
      for (let serviceName of servicesByClient[client]) {
        let matchRow = [client];
        matchRow.push(serviceName);
        let providers = providersByService[serviceName];

        // Pick a random provider for this service.
        let index = this.getRandomInt(providers.length);
        let provider = providers[index];
        if (!provider) {
          provider = '*** None ***'
        }
        matchRow.push(provider);
        matchDataRows.push(matchRow);
      }
    }
    let matches = new SheetClass('Matches');
    matches.clearData();
    matches.setMultipleRows(2, matchDataRows);
  }
  
  doMatching() {
    const providerTabName = 'Services provided - categorized';
    const providerWorkbookId = '1BHlfgXgA-Ej3iRwirMAm7kipAGKKSr3gnD95ktyReXM';
    const clientTabName = 'Clients';
    let providerData = new SheetClass(providerTabName, providerWorkbookId);
    let clientData = new SheetClass(clientTabName);
    let providersByService = this.loadProviderData(providerData);
    console.log(JSON.stringify(providersByService));
    let servicesByClient = this.loadClientData(clientData);
    this.matchThem(providersByService, servicesByClient);
  }
}

function doMatching() {
    (new ProviderClientMatcher()).doMatching();
}
