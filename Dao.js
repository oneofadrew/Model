/**
 * A Data Access Object wraps up a set of functions to allow easily interactivity across a model
 * The Search class allows for new terms to be added to it afterwards
 * so it allows it to be passed between functions to build up a set of terms via composition.
 */
class Dao_ {
  constructor(spreadsheetId, sheetName, keys, startCol, hasHeader, enricher, keyName, richTextConverters) {
    this.findAll = buildGetModels_(spreadsheetId, sheetName, keys, startCol, hasHeader, enricher);
    this.findByKey = buildFindModelByKey_(spreadsheetId, sheetName, keys, startCol, enricher);
    this.findByRow = buildFindModelByRow_(spreadsheetId, sheetName, keys, startCol, enricher);
    this.save = buildSaveModel_(spreadsheetId, sheetName, keys, startCol, keyName, richTextConverters);
    this.bulkInsert = buildBulkInsertModels_(spreadsheetId, sheetName, keys, startCol, hasHeader, enricher, keyName, richTextConverters);
    this.findLastRow = buildFindLastRow_(spreadsheetId, sheetName, startCol);
    this.DATA = spreadsheetId;
    this.SHEET = sheetName;
    this.KEYS = keys;
    this.START_COL = startCol;
    this.KEY_NAME = keyName;
  }

  search(terms) {
    let models = this.findAll();
    return runSearch(terms, models);
  }
}

/**
 * Create a new Data Access Object from the metadata provided.
 * @param {string} spreadsheetId - the file ID of the spreadsheet to use. Will default to the active spreadsheet if not defined.
 * @param {string} sheetName - the sheet that contains the data for the Data Access Object.
 * @param {[string]} keys - the list of keys to use as the fields for the object. These must be in the order of the columns for the data model.
 * @param {string} startCol - the column in the spreadsheet where the data for the model starts.
 * @param {boolean} hasHeader - whether the data table defined by the startCol / endCol has a header row.
 * @param {function} enricher - a function that takes a model object as an only parameter and enriches it with other data based on it's existing fields.
 * @param {string} keyName - a named range for a single cell that contains a number to be used as the sequence for the data model.
 * @param {[function]} richTextConverters - an array of functions that can takes a value as an only parameter and returns a RichTextValue object. Defaults to text only.
 * @return {Dao} a data access object that encapsulates the data access functions for the metadata provided.
 */
function createDao(spreadsheetId, sheetName, keys, startCol, hasHeader, enricher, keyName, richTextConverters) {
  return new Dao_(spreadsheetId, sheetName, keys, startCol, hasHeader, enricher, keyName, richTextConverters)
}

/**
 * Create a new Data Access Object that infers the metadata from the data in the sheet. The first row in the sheet must be a header row. The start column
 * will be inferred to be the first column from the left that has a header value. The end column will be inferred to be column before the first column
 * after the start column that has no header value. Fields will be inferred to be the titles in the header row for each column changed to camel case.
 * @param {string} spreadsheetId - the file ID of the spreadsheet to use. Will default to the active spreadsheet if not defined.
 * @param {string} sheetName - the sheet that contains the data for the Data Access Object.
 * @param {function} enricher - a function that takes a model object as an only parameter and enriches it with other data based on it's existing fields.
 * @param {string} keyName - a named range for a single cell that contains a number to be used as the sequence for the data model.
 * @param {[function]} richTextConverters - an array of functions that can takes a value as an only parameter and returns a RichTextValue object. Defaults to text only.
 * @return {Dao} a data access object that encapsulates the data access functions for the metadata provided. 
 */
function inferDao(spreadsheetId, sheetName, enricher, keyName, richTextConverters, startCol = 'A') {
  //get the header row
  let sheet = ExternalCalls_.getSheetByName(spreadsheetId, sheetName);
  let row = sheet.getRange('1:1');
  let values = row.getValues(); // get all data in one call

  let metadata = inferMetadata_(values, startCol);
  return createDao(spreadsheetId, sheetName, metadata.keys, metadata.startCol, true, enricher, keyName, richTextConverters)
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