function insertSheet(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var ui = SpreadsheetApp.getUi();
  var response;
  if (sheet) {
    response = ui.alert('Sheet named ' + sheetName + ' already exists! Click OK to overwrite, CANCEL to abort.', ui.ButtonSet.OK_CANCEL);
    return response === ui.Button.OK ? sheet : false;
  }
  return SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
}

function getAssetOverview(assets) {
  var assetlist = {};
  var sortedlist = [];
  var sum = 0;  
  assets.forEach(function(item) {
    if (!assetlist[item.type]) {
      assetlist[item.type] = 1;
    } else {
      assetlist[item.type] += 1;
    }
    sum += 1;
  });
  for (var item in assetlist) {
    sortedlist.push([item, assetlist[item]]);
  }
  sortedlist = sortedlist.sort(function(a,b) {
    return b[1] - a[1];
  });
  return {
    sortedlist: sortedlist.length === 0 ? [['','']] : sortedlist,
    sum: sum
  }
}

function buildContainerSheet(containerObj) {
  var sheetName = containerObj.containerPublicId + '_container';
  var sheet = insertSheet(sheetName);
  sheet.setColumnWidth(1, 190);
  sheet.setColumnWidth(2, 340);
  
  var containerHeader = sheet.getRange(1,1,1,2);
  containerHeader.setValues([['Google Tag Manager documentation','']]);
  containerHeader.mergeAcross();
  containerHeader.setBackground('#1155cc');
  containerHeader.setFontWeight('bold');
  containerHeader.setHorizontalAlignment('center');
  containerHeader.setFontColor('white');
  
  var containerLabels = ['Container ID:', 'Container name:', 'Container notes:', 'Latest version ID:', 'Version name:', 'Version description:', 'Version created/published:', 'Link to container:', 'API path:'];
  
  var containerContent = sheet.getRange(2,1,containerLabels.length,2);
  containerContent.setValues([
    ['Container ID:', containerObj.containerPublicId],
    ['Container name:', containerObj.containerName],
    ['Container notes:', containerObj.containerNotes],
    ['Latest version ID:', containerObj.versionId],
    ['Version name:', containerObj.versionName],
    ['Version description:', containerObj.versionDescription],
    ['Version created/published:', containerObj.versionCreatedOrPublished],
    ['Link to container:', containerObj.containerLink],
    ['API path:', 'accounts/' + containerObj.accountId + '/containers/' + containerObj.containerId + '/versions/' + containerObj.versionId]
  ]);
  containerContent.setBackgrounds([
    ['white', 'white'],
    ['#e8ebf8', '#e8ebf8'],
    ['white', 'white'],
    ['#e8ebf8', '#e8ebf8'],
    ['white', 'white'],
    ['#e8ebf8', '#e8ebf8'],
    ['white', 'white'],
    ['#e8ebf8', '#e8ebf8'],
    ['white', 'white']    
  ]);
  containerContent.setNumberFormats([
    ['@', '@'],
    ['@', '@'],
    ['@', '@'],
    ['@', '@'],
    ['@', '@'],    
    ['@', '@'],
    ['@', 'dd/mm/yy at h:mm'],
    ['@', '@'],
    ['@', '@']
  ]);
  containerContent.setVerticalAlignment('top');
  
  var containerLabelCol = sheet.getRange(2,1,containerLabels.length,1);
  containerLabelCol.setFontWeight('bold');
  containerLabelCol.setHorizontalAlignment('right');
  
  var containerDataCol = sheet.getRange(2,2,containerLabels.length,1);
  containerDataCol.setHorizontalAlignment('left');

  var emptyCellFix = sheet.getRange(2,3,containerLabels.length,1);
  var emptyCells = [];
  for (var i = 0; i < containerLabels.length; i++) {
    emptyCells.push([' ']);
  }
  emptyCellFix.setValues(emptyCells);
  
  var overviewHeader = sheet.getRange(1,4,1,6);
  overviewHeader.setValues([['Overview of contents', '', '', '', '', '']]);
  overviewHeader.mergeAcross();
  overviewHeader.setBackground('#85200c');
  overviewHeader.setFontWeight('bold');
  overviewHeader.setHorizontalAlignment('center');
  overviewHeader.setFontColor('white');
  
  var overviewSubHeader = sheet.getRange(2,4,1,6);
  overviewSubHeader.setValues([['Tag type', 'Quantity', 'Trigger type', 'Quantity', 'Variable type', 'Quantity']]);
  overviewSubHeader.setFontWeight('bold');
  overviewSubHeader.setBackground('#e6d6d6');
  
  var tags = getAssetOverview(containerObj.tags);
  var tagsRange = sheet.getRange(3,4,tags.sortedlist.length,2);
  var tagsSum = tags.sum;
  tagsRange.setValues(tags.sortedlist);
  sheet.getRange(3,4,tags.sortedlist.length,1).setHorizontalAlignment('right');
  sheet.getRange(3,5,tags.sortedlist.length,1).setHorizontalAlignment('left');

  var triggers = getAssetOverview(containerObj.triggers);
  var triggersRange = sheet.getRange(3,6,triggers.sortedlist.length,2);
  var triggersSum = triggers.sum;
  triggersRange.setValues(triggers.sortedlist);
  sheet.getRange(3,6,triggers.sortedlist.length,1).setHorizontalAlignment('right');
  sheet.getRange(3,7,triggers.sortedlist.length,1).setHorizontalAlignment('left');

  var variables = getAssetOverview(containerObj.variables);
  var variablesRange = sheet.getRange(3,8,variables.sortedlist.length,2);
  var variablesSum = variables.sum;
  variablesRange.setValues(variables.sortedlist);
  sheet.getRange(3,8,variables.sortedlist.length,1).setHorizontalAlignment('right');
  sheet.getRange(3,9,variables.sortedlist.length,1).setHorizontalAlignment('left');
  
  var contentLength = Math.max(tags.sortedlist.length, variables.sortedlist.length, triggers.sortedlist.length);
  var totalRow = sheet.getRange(contentLength + 3, 4, 1, 6);
  totalRow.setValues([
    ['Total tags:', tagsSum, 'Total triggers:', triggersSum, 'Total variables:', variablesSum]
  ]);
  totalRow.setHorizontalAlignments([['right', 'left', 'right', 'left', 'right', 'left']]);
  totalRow.setFontWeight('bold');
  totalRow.setBackground('#e6d6d6');
}

function startProcess(aid, cid) {
  var latestVersionId = fetchLatestVersionId(aid, cid);
  if (latestVersionId === '0') { throw new Error('No latest version found!'); }
  var latestVersion = fetchLatestVersion(aid, cid, latestVersionId);
  var containerObj = {
    accountId: latestVersion.container.accountId,
    containerId: latestVersion.container.containerId,
    containerName: latestVersion.container.name,
    containerPublicId: latestVersion.container.publicId,
    containerNotes: latestVersion.container.notes || '',
    containerLink: latestVersion.container.tagManagerUrl,
    versionName: latestVersion.name,
    versionId: latestVersion.containerVersionId,
    versionDescription: latestVersion.description || '',
    versionCreatedOrPublished: new Date(parseInt(latestVersion.fingerprint)),
    tags: latestVersion.tag || [],
    variables: latestVersion.variable || [],
    triggers: latestVersion.trigger || []
  };
  buildContainerSheet(containerObj);
}

function fetchLatestVersion(aid, cid, vid) {
  var parent = 'accounts/' + aid + '/containers/' + cid + '/versions/' + vid;
  return TagManager.Accounts.Containers.Versions.get(parent);
}

function fetchLatestVersionId(aid, cid) {
  var parent = 'accounts/' + aid + '/containers/' + cid;
  return TagManager.Accounts.Containers.Version_headers.latest(parent, {
    fields: 'containerVersionId'
  }).containerVersionId;
}

function fetchAccounts() {
  return TagManager.Accounts.list({
    fields: 'account(accountId,name)'
  }).account;
}

function fetchContainers(aid) {
  var parent = 'accounts/' + aid;
  return TagManager.Accounts.Containers.list(parent, {
    fields: 'container(accountId,containerId,publicId,name)'
  }).container;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}

function openContainerSelector() {
  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createTemplateFromFile('ContainerSelector').evaluate().setWidth(400).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Container');
}

function onOpen() {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem('Build documentation', 'openContainerSelector');
  menu.addToUi();
}
