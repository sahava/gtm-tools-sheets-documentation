/* GTM API methods */

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
  var accounts = TagManager.Accounts.list({
    fields: 'account(accountId,name)'
  }).account;
  return accounts || [];
}

function fetchContainers(aid) {
  var parent = 'accounts/' + aid;
  var containers = TagManager.Accounts.Containers.list(parent, {
    fields: 'container(accountId,containerId,publicId,name)'
  }).container;
  return containers || [];
}

function createVersion(aid, cid, wsid) {
  return TagManager.Accounts.Containers.Workspaces.create_version({"name": "Created by GTM Tools Google Sheets add-on", "notes": "Created by GTM Tools Google Sheets add-on"}, 'accounts/' + aid + '/containers/' + cid + '/workspaces/' + wsid).containerVersion;
}

function getWorkspaces() {
  var apiPath = getApiPath();
  
  if (!apiPath) {
    return false;
  }
  
  return TagManager.Accounts.Containers.Workspaces.list(apiPath, {
    fields: 'workspace(name, workspaceId)'
  }).workspace;
}

function fetchContainersWithSelectedMarked(aid) {
  var containerSummary = fetchContainers(aid);
  var selectedContainerId = getContainerIdFromApiPath();
  containerSummary.forEach(function(cont) {
    cont.selected = cont.containerId === selectedContainerId;
  });
  return containerSummary;
}

function fetchAccountsWithSelectedMarked() {
  var accountSummary = fetchAccounts();
  var selectedAccountId = getAccountIdFromApiPath();
  accountSummary.forEach(function(acct) {
    acct.selected = acct.accountId === selectedAccountId;
  });
  return accountSummary;
}

function getContainerPublicIdFromSheetName() {
  var sheet = SpreadsheetApp.getActiveSheet().getName();
  var cid = sheet.match(/^GTM-[a-zA-Z0-9]{4,}/) || [];
  return cid.length ? cid[0] : 'N/A';
}

function getAccountIdFromApiPath() {
  var apiPath = getApiPath();
  return apiPath ? apiPath.split('/')[1] : '';
}

function getContainerIdFromApiPath() {
  var apiPath = getApiPath();
  return apiPath ? apiPath.split('/')[3] : '';
}

function getApiPath() {
  var sheet = SpreadsheetApp.getActiveSheet().getName();
  if (!/^GTM-[a-zA-Z0-9]{4,}_(container|tags|variables|triggers)$/.test(sheet)) {
    return false;
  }
  var containerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet.replace(/_.+$/,'_container'));
  var apiPath = containerSheet.getRange('B10').getValue().replace(/\/versions\/.*/, '');
  return apiPath;
}

function insertSheet(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var ui = SpreadsheetApp.getUi();
  var response;
  if (sheet) {
    response = ui.alert('Sheet named ' + sheetName + ' already exists! Click OK to overwrite, CANCEL to skip.', ui.ButtonSet.OK_CANCEL);
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

function buildRangesObject() {
  var namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
  var rangesObject = {};
  
  namedRanges.forEach(function(range) {
    var name = range.getName();
    if (/(_notes|_json)$/.test(name)) {
      var bareName = name.replace(/(_notes|_json)$/g, '');
      rangesObject[bareName] = rangesObject[bareName] || {};
      if (/_notes$/.test(name)) {
        rangesObject[bareName].notes = range.getRange();
      }
      if (/_json$/.test(name)) {    
        rangesObject[bareName].json = range.getRange();
      }
      rangesObject[bareName].accountId = name.split('_')[1];
      rangesObject[bareName].containerId = name.split('_')[2];   
    }
  });
  
  return rangesObject;
}

function updateSingleNote(noteToUpdate, wsid) {
  var json = noteToUpdate.json;
  json.notes = noteToUpdate.note;
  
  var path = 'accounts/' + noteToUpdate.json.accountId + '/containers/' + noteToUpdate.json.containerId + '/workspaces/' + wsid;

  if ('tagId' in json) { 
    return TagManager.Accounts.Containers.Workspaces.Tags.update(JSON.stringify(json), path + '/tags/' + json.tagId);
  }
  if ('triggerId' in json) {
    return TagManager.Accounts.Containers.Workspaces.Triggers.update(JSON.stringify(json), path + '/triggers/' + json.triggerId);
  }
  if ('variableId' in json) {
    return TagManager.Accounts.Containers.Workspaces.Variables.update(JSON.stringify(json), path + '/variables/' + json.variableId);
  }
}

function markChangedNotes() {
  var rangesObject = buildRangesObject();
  var count = 0;
  
  if (Object.keys(rangesObject).length === 0) {
    throw new Error('No valid documentation sheets found. Remember to run the <strong>Build documentation</strong> menu option first!');
  }
  
  for (var item in rangesObject) {
    var notes = rangesObject[item].notes.getValues();
    var json = rangesObject[item].json.getValues();
    notes.forEach(function(note, index) {
      var cell = rangesObject[item].notes.getCell(index + 1, 1);
      var jsonNote = JSON.parse(json[index]).notes || '';
      if (note[0] === jsonNote) {
        cell.setBackground('#fff');
      } else if (note[0] !== jsonNote) {
        cell.setBackground('#f6b26b');
        count++;
      }
    });
  } 
    
  return count;
}

function processNotes(action) {
  var rangesObject = buildRangesObject();
  var notesToUpdate = [];
  var selectedAccountId = getAccountIdFromApiPath();
  var selectedContainerId = getContainerIdFromApiPath();

  for (var item in rangesObject) {
    var notes = rangesObject[item].notes.getValues();
    var json = rangesObject[item].json.getValues();
    notes.forEach(function(note, index) {
      var cell = rangesObject[item].notes.getCell(index + 1, 1);
      var jsonNote = JSON.parse(json[index]).notes || '';
      if (note[0] === jsonNote) {
        cell.setBackground('#fff');
      } else if (note[0] !== jsonNote) {
        if (action === 'mark') {
          cell.setBackground('#f6b26b');
        }
        if (action === 'push' && selectedAccountId === rangesObject[item].accountId && selectedContainerId === rangesObject[item].containerId) {
          cell.setBackground('#fff');
          notesToUpdate.push({
            note: note[0],
            json: JSON.parse(json[index])
          });
        }
      }
    });
  }
  
  return notesToUpdate;
}

function formatTags(tags) {
  var data = [];
  tags.forEach(function(tag) {
    data.push([
      tag.name,
      tag.tagId,
      tag.type,
      tag.parentFolderId || '',
      new Date(parseInt(tag.fingerprint)),
      tag.firingTriggerId ? tag.firingTriggerId.join(',') : '',
      tag.blockingTriggerId ? tag.blockingTriggerId.join(',') : '',
      tag.setupTag ? tag.setupTag[0].tagName : '',
      tag.teardownTag ? tag.teardownTag[0].tagName : '',
      tag.notes || '',
      JSON.stringify(tag)
    ]);
  });
  return data;
}

function formatVariables(variables) {
  var data = [];
  variables.forEach(function(variable) {
    data.push([
      variable.name,
      variable.variableId,
      variable.type,
      variable.parentFolderId || '',
      new Date(parseInt(variable.fingerprint)),
      variable.notes || '',
      JSON.stringify(variable)
    ]);
  });
  return data;
}

function formatTriggers(triggers) {
  var data = [];
  triggers.forEach(function(trigger) {
    data.push([
      trigger.name,
      trigger.triggerId,
      trigger.type,
      trigger.parentFolderId || '',
      new Date(parseInt(trigger.fingerprint)),
      trigger.notes || '',
      JSON.stringify(trigger)
    ]);
  });
  return data;
}

function clearInvalidRanges() {
  var storedRanges = JSON.parse(PropertiesService.getUserProperties().getProperty('named_ranges')) || {};
  var storedRangesNames = Object.keys(storedRanges);
  
  var namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
  var namedRangesNames = namedRanges.map(function(a) { return a.getName(); });
  
  storedRangesNames.forEach(function(storedRangeName) {
    if (namedRangesNames.indexOf(storedRangeName) === -1) {
      SpreadsheetApp.getActiveSpreadsheet().removeNamedRange(storedRangeName);
      delete storedRanges[storedRangeName];
    }
  });
  
  PropertiesService.getUserProperties().setProperty('named_ranges', JSON.stringify(storedRanges));
}

function setNamedRanges(sheet,rangeName,notesIndex,jsonIndex,colLength) {
  var notesRange = sheet.getRange(3,notesIndex,colLength,1);
  var notesRangeName = rangeName + '_notes';
  SpreadsheetApp.getActiveSpreadsheet().setNamedRange(notesRangeName, SpreadsheetApp.getActiveSpreadsheet().getRange(sheet.getName() + '!' + notesRange.getA1Notation()));
  var jsonRange = sheet.getRange(3,jsonIndex,colLength,1);
  var jsonRangeName = rangeName + '_json';
  SpreadsheetApp.getActiveSpreadsheet().setNamedRange(jsonRangeName, SpreadsheetApp.getActiveSpreadsheet().getRange(sheet.getName() + '!' + jsonRange.getA1Notation()));
  
  var ranges = JSON.parse(PropertiesService.getUserProperties().getProperty('named_ranges')) || {};
  ranges[notesRangeName] = true;
  ranges[jsonRangeName] = true;
  PropertiesService.getUserProperties().setProperty('named_ranges', JSON.stringify(ranges));
}

function createHeaders(sheet, labels, title) {
  var headerRange = sheet.getRange(1,1,1,labels.length);
  headerRange.mergeAcross();
  headerRange.setValue(title);
  headerRange.setBackground('#1155cc');
  headerRange.setFontWeight('bold');
  headerRange.setFontColor('white');
  
  var labelsRange = sheet.getRange(2,1,1,labels.length);
  labelsRange.setValues([labels]);
  labelsRange.setFontWeight('bold');
}

function buildTriggerSheet(containerObj) {
  var sheetName = containerObj.containerPublicId + '_triggers';
  var sheet = insertSheet(sheetName);
  
  if (sheet === false) { return; }
  
  sheet.clear();
  
  var triggerLabels = ['Trigger name', 'Trigger ID', 'Trigger type', 'Folder ID', 'Last modified', 'Notes', 'JSON (do NOT edit!)'];

  createHeaders(sheet, triggerLabels, 'Triggers for container ' + containerObj.containerPublicId + ' (' + containerObj.containerName + ').');

  sheet.setColumnWidth(1, 305);
  sheet.setColumnWidth(2, 75);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 75);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 305);
  sheet.setColumnWidth(7, 130);
  
  var triggersObject = formatTriggers(containerObj.triggers);
  if (triggersObject.length) {
    var dataRange = sheet.getRange(3,1,triggersObject.length,triggerLabels.length);
    dataRange.setValues(triggersObject);
    dataRange.setBackground('#fff');
    
    var rangeName = 'triggers_' + containerObj.accountId + '_' + containerObj.containerId;
    setNamedRanges(sheet,rangeName,triggerLabels.indexOf('Notes') + 1,triggerLabels.indexOf('JSON (do NOT edit!)') + 1,triggersObject.length);
  
    var formats = triggersObject.map(function(a) {
      return ['@', '@', '@', '@', 'dd/mm/yy at h:mm', '@', '@'];
    });
    dataRange.setNumberFormats(formats);
    dataRange.setHorizontalAlignment('left');
  }
}

function buildVariableSheet(containerObj) {
  var sheetName = containerObj.containerPublicId + '_variables';
  var sheet = insertSheet(sheetName);
  
  if (sheet === false) { return; }
  
  sheet.clear();
  
  var variableLabels = ['Variable name', 'Variable ID', 'Variable type', 'Folder ID', 'Last modified', 'Notes', 'JSON (do NOT edit!)'];

  createHeaders(sheet, variableLabels, 'Variables for container ' + containerObj.containerPublicId + ' (' + containerObj.containerName + ').');

  sheet.setColumnWidth(1, 305);
  sheet.setColumnWidth(2, 75);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 75);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 305);
  sheet.setColumnWidth(7, 130);
  
  var variablesObject = formatVariables(containerObj.variables);
  if (variablesObject.length) {
    var dataRange = sheet.getRange(3,1,variablesObject.length,variableLabels.length);
    dataRange.setValues(variablesObject);
    dataRange.setBackground('#fff');
    
    var rangeName = 'variables_' + containerObj.accountId + '_' + containerObj.containerId;
    setNamedRanges(sheet,rangeName,variableLabels.indexOf('Notes') + 1,variableLabels.indexOf('JSON (do NOT edit!)') + 1,variablesObject.length);
  
    var formats = variablesObject.map(function(a) {
      return ['@', '@', '@', '@', 'dd/mm/yy at h:mm', '@', '@'];
    });
    dataRange.setNumberFormats(formats);
    dataRange.setHorizontalAlignment('left');
  }
}

function buildTagSheet(containerObj) {
  var sheetName = containerObj.containerPublicId + '_tags';
  var sheet = insertSheet(sheetName);
  
  if (sheet === false) { return; }
  
  sheet.clear();
  
  var tagLabels = ['Tag name', 'Tag ID', 'Tag type', 'Folder ID', 'Last modified', 'Firing trigger IDs', 'Exception trigger IDs', 'Setup tag', 'Cleanup tag', 'Notes', 'JSON (do NOT edit!)'];

  createHeaders(sheet, tagLabels, 'Tags for container ' + containerObj.containerPublicId + ' (' + containerObj.containerName + ').');

  sheet.setColumnWidth(1, 305);
  sheet.setColumnWidth(2, 75);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 75);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 150);
  sheet.setColumnWidth(8, 205);
  sheet.setColumnWidth(9, 205);
  sheet.setColumnWidth(10, 305);
  sheet.setColumnWidth(11, 130);
  
  var tagsObject = formatTags(containerObj.tags);
  if (tagsObject.length) {
    var dataRange = sheet.getRange(3,1,tagsObject.length,tagLabels.length);
    dataRange.setValues(tagsObject);
    dataRange.setBackground('#fff');

    var rangeName = 'tags_' + containerObj.accountId + '_' + containerObj.containerId;
    setNamedRanges(sheet,rangeName,tagLabels.indexOf('Notes') + 1,tagLabels.indexOf('JSON (do NOT edit!)') + 1,tagsObject.length);
  
    var formats = tagsObject.map(function(a) {
      return ['@', '@', '@', '@', 'dd/mm/yy at h:mm', '@', '@', '@', '@', '@', '@'];
    });
    dataRange.setNumberFormats(formats);
    dataRange.setHorizontalAlignment('left');
  }
}

function buildContainerSheet(containerObj) {
  var sheetName = containerObj.containerPublicId + '_container';
  var sheet = insertSheet(sheetName);
  
  if (sheet === false) { return; }
  
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
    [containerLabels[0], containerObj.containerPublicId],
    [containerLabels[1], containerObj.containerName],
    [containerLabels[2], containerObj.containerNotes],
    [containerLabels[3], containerObj.versionId],
    [containerLabels[4], containerObj.versionName],
    [containerLabels[5], containerObj.versionDescription],
    [containerLabels[6], containerObj.versionCreatedOrPublished],
    [containerLabels[7], containerObj.containerLink],
    [containerLabels[8], 'accounts/' + containerObj.accountId + '/containers/' + containerObj.containerId + '/versions/' + containerObj.versionId]
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
  
  var overviewHeader = sheet.getRange(1,4,1,8);
  overviewHeader.setValues([['Overview of contents', '', '', '', '', '', '', '']]);
  overviewHeader.mergeAcross();
  overviewHeader.setBackground('#85200c');
  overviewHeader.setFontWeight('bold');
  overviewHeader.setHorizontalAlignment('center');
  overviewHeader.setFontColor('white');
  
  var overviewSubHeader = sheet.getRange(2,4,1,8);
  overviewSubHeader.setValues([['Tag type', 'Quantity', 'Trigger type', 'Quantity', 'Variable type', 'Quantity', 'Folder ID', 'Folder name']]);
  overviewSubHeader.setHorizontalAlignments([['right','left','right','left','right','left', 'right', 'left']]);
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
  
  var folders = containerObj.folders.map(function(folder) {
    return [folder.folderId, folder.name];
  });
  if (folders.length) {
    var foldersRange = sheet.getRange(3,10,folders.length,2);
    foldersRange.setValues(folders);
  }
  
  var contentLength = Math.max(tags.sortedlist.length, variables.sortedlist.length, triggers.sortedlist.length, folders.length);
  var totalRow = sheet.getRange(contentLength + 3, 4, 1, 8);
  totalRow.setValues([
    ['Total tags:', tagsSum, 'Total triggers:', triggersSum, 'Total variables:', variablesSum, '', '']
  ]);
  totalRow.setHorizontalAlignments([['right', 'left', 'right', 'left', 'right', 'left', 'right', 'left']]);
  totalRow.setFontWeight('bold');
  totalRow.setBackground('#e6d6d6');
}

function startProcess(aid, cid) {
  var latestVersionId = fetchLatestVersionId(aid, cid);
  if (latestVersionId === '0') { throw new Error('You need to create or publish a version in the container before you can build its documentaiton!'); }
  var latestVersion = fetchLatestVersion(aid, cid, latestVersionId);
  var containerObj = {
    accountId: latestVersion.container.accountId,
    containerId: latestVersion.container.containerId,
    containerName: latestVersion.container.name,
    containerPublicId: latestVersion.container.publicId,
    containerNotes: latestVersion.container.notes || '',
    containerLink: latestVersion.container.tagManagerUrl,
    versionName: latestVersion.name || '',
    versionId: latestVersion.containerVersionId,
    versionDescription: latestVersion.description || '',
    versionCreatedOrPublished: new Date(parseInt(latestVersion.fingerprint)),
    tags: latestVersion.tag || [],
    variables: latestVersion.variable || [],
    triggers: latestVersion.trigger || [],
    folders: latestVersion.folder || []
  };
  buildContainerSheet(containerObj);
  buildTagSheet(containerObj);
  buildTriggerSheet(containerObj);
  buildVariableSheet(containerObj);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}

function openContainerSelector() {
  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createTemplateFromFile('ContainerSelector').evaluate().setWidth(400).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, 'Build documentation');
}

function openMarkChangesModal() {
  clearInvalidRanges();
  var ui = SpreadsheetApp.getUi();
  if (Object.keys(buildRangesObject()).length === 0) {
    ui.alert('No valid documentation sheets found! Run "Build documentation" if necessary.');
    return;
  }
  var html = HtmlService.createTemplateFromFile('MarkChangesModal').evaluate().setWidth(400).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Mark changes to Notes');
}

function openPushChangesModal() {
  clearInvalidRanges();
  var ui = SpreadsheetApp.getUi();
  if (getApiPath() === false) {
    ui.alert('You need to have a valid documentation sheet selected first! Run "Build documentation" if necessary.', ui.ButtonSet.OK);
    return;
  }
  var html = HtmlService.createTemplateFromFile('PushChangesModal').evaluate().setWidth(400).setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(html, 'Push changes to Notes');
}

function onOpen() {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem('Build documentation', 'openContainerSelector')
  menu.addItem('Mark changes to Notes', 'openMarkChangesModal')
  menu.addItem('Push changes to Notes', 'openPushChangesModal')
  menu.addToUi();
}
