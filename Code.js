function buildGetModels(spreadsheetId, sheetName, startCol, endCol, hasHeader, builder) {
  return function() {
    let ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    let row = getLastRow(sheet, startCol);
    let firstRow = hasHeader ? 2 : 1;
    if (firstRow > row) return [];
    let values = sheet.getRange(startCol+firstRow + ":" + endCol+(row)).getValues();
    let models = [];
    for (let i in values) {
      models[i] = builder(values[i], i + firstRow);
    }
    return models;
  }
}

function buildFindModelByKey(spreadsheetId, sheetName, startCol, endCol, builder) {
  return function(key) {
    let ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    let row = findKey(sheet, key, startCol);
    let values = sheet.getRange(startCol+row + ":" + endCol+row).getValues();
    let model = builder(values[0], row);
    return model;
  }
}

function buildFindModelByRow(spreadsheetId, sheetName, startCol, endCol, builder) {
  return function(row) {
    let ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    let values = sheet.getRange(startCol+row + ":" + endCol+row).getValues();
    let model = builder(values[0], row);
    return model;
  }
}

function buildSaveModel(spreadsheetId, sheetName, startCol, endCol, keyName, getModelValues, buildModel) {
  return function(model) {
    let ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    // flatten to model values for the record
    let values = [getModelValues(model)];

    // check if we are processing rich text values or not
    let isRichText = (values[0][0] && typeof values[0][0] === "object" && values[0][0].constructor.name == "Object");
    let keyValue = isRichText ? values[0][0].getText() : values[0][0];
    
    // grab the document lock for read and write consistency
    let lock = LockService.getDocumentLock();
    lock.waitLock(10000);

    //if this requires a generated key and the key value isn't set, generate the key
    keyValue = keyValue || !keyName ? keyValue : incrementKey(ss, keyName);
    model[keyName] = keyValue;
    // convert the key back to a rich text value if necessary
    values[0][0] = isRichText ? SpreadsheetApp.newRichTextValue().setText(keyValue).build() : keyValue;

    // try to find a record to update based on the key, otherwise we'll create a new record
    let row;
    try {
      row = findKey(sheet, keyValue, startCol);
    } catch (e) {
      row = getFirstEmptyRow(sheet, startCol);
    }

    if (model.row && model.row != row) {
      throw new Error(`The row of the model (${model.row}) did not match the row of the primary key (${values[0][0]}) of the model (${row})`);
    }

    // write the values for the record
    if (isRichText) {
      sheet.getRange(startCol+row + ":" + endCol+row).setRichTextValues(values);
    } else {
      sheet.getRange(startCol+row + ":" + endCol+row).setValues(values);
    }

    // and we are done
    SpreadsheetApp.flush();
    lock.releaseLock();
    model["row"] = row;
    return model;
  }
}

function buildBulkInsertModels(spreadsheetId, sheetName, startCol, endCol, keyName, getModels, getModelValues) {
  return function(models) {
    let ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    // flatten the models to a 2D array
    let values = []
    for (i in models) {
      values[i] = getModelValues(models[i]);
    }

    // check if we are processing rich text values or not
    //let isRichText = (typeof values[0][0] === "object" && values[0][0].constructor.name == "Object");
    //let keyValue = isRichText ? values[0][0].getText() : values[0][0];

    // get the lock - we need to do this before any reads to guarantee both read and write consistency
    let lock = LockService.getDocumentLock();
    lock.waitLock(10000);

    // get a map of the existing keys. in the sheet
    let existing = getModels();
    let existingKeys = {};
    for (i in existing) {
      let existingValues = getModelValues(existing[i]);
      existingKeys[existingValues[0]] = true;
    }

    // iterate over the new values for bulk insert to find any duplicates
    let duplicateKeys = ""
    for (i in values) {
      duplicateKeys = existingKeys[values[i][0]] ? duplicateKeys + " " + values[i][0] : duplicateKeys;
    }

    // if we found any duplicates raise an error
    if (duplicateKeys != "") {
      throw new Error("The bulk insert has duplicate keys:" + duplicateKeys);
    }

    // if we need to create the keys then create them here
    let lastKey = incrementKey(ss, keyName, values.length);
    for (let i in values) {
      values[i][0] = lastKey - values.length + i;
    }
    
    // we are good to progress so run the insert
    let row = getFirstEmptyRow(sheet, startCol);
    sheet.getRange(startCol+row + ":" + endCol+(row+values.length-1)).setValues(values);

    // and we are done
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

function buildBuildModel(keys, enricher) {
  return function(values, row) {
    let model = row ? {"row" : row} : {};
    for (let i in keys) model[keys[i]] = values[i];
    return enricher ? enricher(model) : model;
  }
}

function buildGetModelValues(keys) {
  return function(model) {
    let values = [];
    for (let i in keys) values[i] = model[keys[i]];
    return values;
  }
}

function buildGetRichModelValues(keys, richConverters) {
  return function(model) {
    let values = [];
    for (let i in keys) values[i] = richConverters[i](model[keys[i]]);
    return values;
  }
}

/**
 * Spreadsheet navigation
 */
function getFirstEmptyRow(sheet, col) {
  return findKey(sheet, "", col);
}

function getLastRow(sheet, col) {
  return getFirstEmptyRow(sheet, col) - 1;
}

function findKey(sheet, key, col) {
  let column = sheet.getRange(col+':'+col);
  let values = column.getValues(); // get all data in one call
  let ct = 0;
  while ( values[ct] && values[ct][0] != key && values[ct][0] != "") {
    ct++;
  }
  if (values[ct] && values[ct][0] == key) return ct + 1;
  throw new Error("Could not find '"+key+"'");
}

function incrementKey(ss, key, increment) {
  increment = increment ? increment : 1;
  let range = ss.getRangeByName(key);
  let values = range.getValues();
  values[0][0] = values[0][0] + increment;
  range.setValues(values);
  return values[0][0];
}