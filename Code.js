/**
 * The ExternalCalls_ object is here to abstract away our external calls to allow us to drive
 * unit tests with mock objects to validate our functionality is working as expected
 */
let ExternalCalls_ = {
  "getDocumentLock" : () => {
    return LockService.getDocumentLock();
  },
  "spreadsheetFlush" : () => {
    SpreadsheetApp.flush();
  },
  "newRichTextValue" : (value) => {
    return SpreadsheetApp.newRichTextValue().setText(value).build();
  }
};

function buildModel_(values, row, keys, enricher) {
  let model = row ? {"row" : row} : {};
  for (let i in keys) model[keys[i]] = values[i];
  return enricher ? enricher(model) : model;
}

function getModelValues_(model, keys, converters) {
    let values = [];
    for (let i in keys) values[i] = converters[i](model[keys[i]]);
    return values;
}

function getModels_(sheet, keys, startCol, endCol, hasHeader, enricher) {
  let row = findLastRow_(sheet, startCol);
  let firstRow = hasHeader ? 2 : 1;
  if (firstRow > row) return [];
  let values = sheet.getRange(startCol+firstRow + ":" + endCol+(row)).getValues();
  let models = [];
  for (let i in values) {
    models[i] = buildModel_(values[i], i + firstRow, keys, enricher);
  }
  return models;
}

function findModelByKey_(key, sheet, keys, startCol, endCol, enricher) {
  let row = findKey_(sheet, key, startCol);
  let values = sheet.getRange(startCol+row + ":" + endCol+row).getValues();
  let model = buildModel_(values[0], row, keys, enricher);
  return model;
}

function findModelByRow_(row, sheet, keys, startCol, endCol, enricher) {
  let values = sheet.getRange(startCol+row + ":" + endCol+row).getValues();
  let model = buildModel_(values[0], row, keys, enricher);
  return model;
}

function saveModel_(model, sheet, keys, startCol, endCol, sequence, converters) {
  // flatten to model values for the record
  let values = [getModelValues_(model, keys, converters)];

  // check if we are processing rich text values or not
  let keyValue = values[0][0].getText();
  
  // grab the document lock for read and write consistency
  let lock = ExternalCalls_.getDocumentLock();
  lock.waitLock(10000);

  //if this requires a generated key and the key value isn't set, generate the key
  keyValue = keyValue || !sequence ? keyValue : incrementKey_(sequence);
  model[sequence] = keyValue;
  // convert the key back to a rich text value
  values[0][0] = getRichText_(keyValue);

  // try to find a record to update based on the key, otherwise we'll create a new record
  let row;
  try {
    row = findKey_(sheet, keyValue, startCol);
  } catch (e) {
    row = getFirstEmptyRow_(sheet, startCol);
  }

  if (model.row && model.row != row) {
    throw new Error(`The row of the model (${model.row}) did not match the row of the primary key (${values[0][0]}) of the model (${row})`);
  }

  // write the values for the record
  sheet.getRange(startCol+row + ":" + endCol+row).setRichTextValues(values);
  
  // and we are done
  ExternalCalls_.spreadsheetFlush();
  lock.releaseLock();
  model["row"] = row;
  return model;
}

function bulkInsertModels_(models, sheet, keys, startCol, endCol, hasHeader, enricher, sequence, converters) {
  // flatten the models to a 2D array
  let values = []
  for (i in models) {
    values[i] = getModelValues_(models[i], keys, converters);
  }

  // get the lock - we need to do this before any reads to guarantee both read and write consistency
  let lock = ExternalCalls_.getDocumentLock();
  lock.waitLock(10000);

  // get a map of the existing keys. in the sheet
  let existing = getModels_(sheet, keys, startCol, endCol, hasHeader, enricher);
  let existingKeys = {};
  for (i in existing) {
    let existingValues = getModelValues_(existing[i], keys, converters);
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
  let lastKey = incrementKey_(sequence, values.length);
  for (let i in values) {
    values[i][0] = lastKey - values.length + i;
  }
  
  // we are good to progress so run the insert
  let row = getFirstEmptyRow_(sheet, startCol);
  sheet.getRange(startCol+row + ":" + endCol+(row+values.length-1)).setValues(values);

  // and we are done
  ExternalCalls_.spreadsheetFlush();
  lock.releaseLock();
}

function findLastRow_(sheet, col) {
  return getFirstEmptyRow_(sheet, col) - 1;
}

/**
 * Spreadsheet navigation
 */
function getFirstEmptyRow_(sheet, col) {
  return findKey_(sheet, "", col);
}

function getRichText_(value) {
  value = value ? value : "";
  return ExternalCalls_.newRichTextValue(value);
}

function findKey_(sheet, key, col) {
  let column = sheet.getRange(col+':'+col);
  let values = column.getValues(); // get all data in one call
  let ct = 0;
  while ( values[ct] && values[ct][0] != key && values[ct][0] != "") {
    ct++;
  }
  if (values[ct] && values[ct][0] == key) return ct + 1;
  throw new Error("Could not find '"+key+"'");
}

function incrementKey_(sequence, increment = 1) {
  let values = sequence.getValues();
  values[0][0] = values[0][0] + increment;
  sequence.setValues(values);
  return values[0][0];
}