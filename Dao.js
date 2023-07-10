/**
 * A Data Access Object wraps up a set of functions to allow easily interactivity across a model
 */
class Dao_ {
  constructor(sheet, keys, startCol, hasHeader, enricher, sequence, richConverters) {
    this.SHEET = sheet;
    this.KEYS = keys;
    this.START_COL = startCol;
    this.END_COL = calculateEndColumn_(startCol, keys.length);;
    this.HAS_HEADER = hasHeader;
    this.ENRICHER = enricher;
    this.SEQUENCE = sequence;
    let safeConverters = richConverters ? richConverters : [];
    for (let i in keys) safeConverters[i] = safeConverters[i] ? safeConverters[i] : getRichText_;
    this.CONVERTERS = safeConverters;
  }

  build(values, row) {
    //return buildModel_(values, row, this.KEYS, this.ENRICHER)
    let model = row ? {"row" : row} : {};
    for (let i in this.KEYS) model[this.KEYS[i]] = values[i];
    return this.ENRICHER ? this.ENRICHER(model) : model;
  }

  findAll() {
    //return getModels_(this.SHEET, this.KEYS, this.START_COL, this.END_COL, this.HAS_HEADER, this.ENRICHER);
    let row = this.findLastRow();
    let firstRow = this.HAS_HEADER ? 2 : 1;
    if (firstRow > row) return [];
    let values = this.SHEET.getRange(this.START_COL+firstRow + ":" + this.END_COL+(row)).getValues();
    let models = [];
    for (let i in values) {
      models[i] = this.build(values[i], firstRow + Number(i));
    }
    return models;
  }
  
  findByKey(key) {
    //return findModelByKey_(key, this.SHEET, this.KEYS, this.START_COL, this.END_COL, this.ENRICHER);
    let row = findKey_(this.SHEET, key, this.START_COL);
    return this.findByRow(row);
  }
  
  findByRow(row) {
    //return findModelByRow_(row, this.SHEET, this.KEYS, this.START_COL, this.END_COL, this.ENRICHER);
    let values = this.SHEET.getRange(this.START_COL+row + ":" + this.END_COL+row).getValues();
    return this.build(values[0], row);
  }
  
  save(model) {
    //return saveModel_(model, this.SHEET, this.KEYS, this.START_COL, this.END_COL, this.SEQUENCE, this.CONVERTERS);

    // flatten to model values for the record
    let values = [getModelValues_(model, this.KEYS, this.CONVERTERS)];

    // check if we are processing rich text values or not
    let keyValue = values[0][0].getText();
    
    // grab the document lock for read and write consistency
    let lock = ExternalCalls_.getDocumentLock();
    lock.waitLock(10000);

    //if this requires a generated key and the key value isn't set, generate the key
    keyValue = keyValue || !this.SEQUENCE ? keyValue : incrementKey_(this.SEQUENCE);
    model[this.KEYS[0]] = keyValue;
    // convert the key back to a rich text value
    values[0][0] = getRichText_(keyValue);

    // try to find a record to update based on the key, otherwise we'll create a new record
    let row;
    try {
      row = findKey_(this.SHEET, keyValue, this.START_COL);
    } catch (e) {
      row = getFirstEmptyRow_(this.SHEET, this.START_COL);
    }

    if (model.row && model.row != row) {
      throw new Error(`The row of the model (${model.row}) did not match the row of the primary key (${values[0][0]}) of the model (${row})`);
    }

    // write the values for the record
    this.SHEET.getRange(this.START_COL+row + ":" + this.END_COL+row).setRichTextValues(values);
    
    // and we are done
    ExternalCalls_.spreadsheetFlush();
    lock.releaseLock();
    model["row"] = row;
    return model;
  }
  
  bulkInsert(models) {
    //return bulkInsertModels_(models, this.SHEET, this.KEYS, this.START_COL, this.END_COL, this.HAS_HEADER, this.ENRICHER, this.SEQUENCE, this.CONVERTERS);

    // flatten the models to a 2D array
    let values = []
    for (i in models) {
      values[i] = getModelValues_(models[i], this.KEYS, this.CONVERTERS);
    }

    // get the lock - we need to do this before any reads to guarantee both read and write consistency
    let lock = ExternalCalls_.getDocumentLock();
    lock.waitLock(10000);

    // get a map of the existing keys. in the sheet
    let existing = this.findAll();
    let existingKeys = {};
    for (i in existing) {
      let existingValues = getModelValues_(existing[i], this.KEYS, this.CONVERTERS);
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
    let lastKey = incrementKey_(this.SEQUENCE, values.length);
    for (let i in values) {
      values[i][0] = lastKey - values.length + i;
    }
    
    // we are good to progress so run the insert
    let row = getFirstEmptyRow_(this.SHEET, this.START_COL);
    this.SHEET.getRange(this.START_COL+row + ":" + this.END_COL+(row+values.length-1)).setValues(values);

    // and we are done
    ExternalCalls_.spreadsheetFlush();
    lock.releaseLock();
  }
  
  findLastRow() {
    //return findLastRow_(this.SHEET, this.START_COL);
    return getFirstEmptyRow_(this.SHEET, this.START_COL) - 1;
  }
  
  search(terms) {
    let models = this.findAll();
    return runSearch(terms, models);
  }
}

/**
 * Create a new Data Access Object from the metadata provided.
 * @param {string} sheet - the sheet that contains the data for the Data Access Object.
 * @param {[string]} keys - the list of keys to use as the fields for the object. These must be in the order of the columns for the data model.
 * @param {string} startCol - the column in the spreadsheet where the data for the model starts.
 * @param {boolean} hasHeader - whether the data table defined by the startCol / endCol has a header row.
 * @param {function} enricher - a function that takes a model object as an only parameter and enriches it with other data based on it's existing fields.
 * @param {string} sequence - a named range for a single cell that contains a number to be used as the sequence for the data model.
 * @param {[function]} richTextConverters - an array of functions that can takes a value as an only parameter and returns a RichTextValue object. Defaults to text only.
 * @return {Dao} a data access object that encapsulates the data access functions for the metadata provided.
 */
function createDao(sheet, keys, startCol, hasHeader, enricher, sequence, richTextConverters) {
  return new Dao_(sheet, keys, startCol, hasHeader, enricher, sequence, richTextConverters)
}

/**
 * Create a new Data Access Object that infers the metadata from the data in the sheet. The first row in the sheet must be a header row. The start column
 * will be inferred to be the first column from the left that has a header value. The end column will be inferred to be column before the first column
 * after the start column that has no header value. Fields will be inferred to be the titles in the header row for each column changed to camel case.
 * @param {string} sheet - the sheet that contains the data for the Data Access Object.
 * @param {function} enricher - a function that takes a model object as an only parameter and enriches it with other data based on it's existing fields.
 * @param {string} sequence - a named range for a single cell that contains a number to be used as the sequence for the data model.
 * @param {[function]} richTextConverters - an array of functions that can takes a value as an only parameter and returns a RichTextValue object. Defaults to text only.
 * @return {Dao} a data access object that encapsulates the data access functions for the metadata provided. 
 */
function inferDao(sheet, enricher, richTextConverters, sequence, startCol = 'A') {
  //get the header row
  let row = sheet.getRange('1:1');
  let values = row.getValues(); // get all data in one call
  let metadata = inferMetadata_(values, startCol);
  return createDao(sheet, metadata.keys, metadata.startCol, true, enricher, sequence, richTextConverters)
}

function inferMetadata_(values, startCol) {
  //where do we start having header values
  let cols = getColumnReferences_();
  let start = cols.indexOf(startCol);
  start = start < 0 ? 0 : start;
  while (values[0][start] == '') start++;

  //where do we end having header values
  let end = start;
  while (values[0][end] != '' && values[0].length >= end) end++;
  //end++; //we will use this in an array.slice() function, which ignores the final element at end index
  
  //the metadata object to return
  let metadata = {};

  //work out the start column reference
  metadata["startCol"] = cols[start];

  //get the header values from the header row
  let keys = values[0].slice(start, end);

  //convert the header values to camel case keys 
  for (i in keys) keys[i] = toCamelCase_(keys[i]);
  metadata["keys"] = keys;

  //return the inferred metadata
  return metadata;
}

function toCamelCase_(str) {
  let words = str.trim().split(/\s+/);
  for (i in words) {
    words[i] = words[i].toLowerCase();
    if (i > 0) words[i] = words[i].charAt(0).toUpperCase() + words[i].slice(1);
  }
  return words.join('');
}

function calculateEndColumn_(startCol, length) {
  let cols = getColumnReferences_();
  let startIndex = cols.indexOf(startCol);
  let endIndex = startIndex + length - 1;
  if (startIndex == -1) throw new Error(`Invalid startCol '${startCol}' provided.`);
  if (endIndex > 701) throw new Error('The Model library only supports models that go up to column ZZ');
  return cols[endIndex]; 
}

function getColumnReferences_() {
  let cols = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
  for (let count = 0; count < 26; count++) {
    for (let i = 0; i < 26; i++) cols[cols.length] = `${cols[count]}${cols[i]}`;
  }
  return cols;
}