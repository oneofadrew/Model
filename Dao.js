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

  findAll() {
    return getModels_(this.SHEET, this.KEYS, this.START_COL, this.END_COL, this.HAS_HEADER, this.ENRICHER);
  }
  
  findByKey(key) {
    return findModelByKey_(key, this.SHEET, this.KEYS, this.START_COL, this.END_COL, this.ENRICHER);
  }
  
  findByRow(row) {
    return findModelByRow_(row, this.SHEET, this.KEYS, this.START_COL, this.END_COL, this.ENRICHER);
  }
  
  save(model) {
    return saveModel_(model, this.SHEET, this.KEYS, this.START_COL, this.END_COL, this.SEQUENCE, this.CONVERTERS);
  }
  
  bulkInsert(models) {
    return bulkInsertModels_(models, this.SHEET, this.KEYS, this.START_COL, this.END_COL, this.HAS_HEADER, this.ENRICHER, this.SEQUENCE, this.CONVERTERS);
  }
  
  findLastRow() {
    return findLastRow_(this.SHEET, this.START_COL);
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