function startProcess(aid, cid) {
  var latestVersionId = fetchLatestVersionId(aid, cid);
  if (latestVersionId === '0') { throw new Error('No latest version found!'); }
  var latestVersion = fetchLatestVersion(aid, cid, latestVersionId);
  Logger.log(latestVersion);
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
