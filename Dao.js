/**
 * A Data Access Object wraps up a set of functions to allow easily interactivity across a model
 */
class Dao_ {
  constructor(sheet, keys, startCol, hasHeader, enricher, sequence, richConverters) {
    this.SHEET = sheet;
    this.KEYS = keys;
    this.START_COL = startCol;
    this.END_COL = calculateEndColumn_(startCol, keys.length);
    this.HAS_HEADER = hasHeader;
    this.ENRICHER = enricher;
    this.SEQUENCE = sequence;
    const converters = richConverters ? richConverters : {};
    this.CONVERTERS = keys.reduce((safeConverters, key) => {return {...safeConverters, ...{[key] : converters[key] ? converters[key] : getRichText_}}}, {});
  }

  build(values, row) {
    const model = this.KEYS.reduce((model, key, i) => Object.assign(model, {[key]: values[i]}), {"row": row});
    return this.ENRICHER ? this.ENRICHER(model) : model;
  }

  findAll() {
    const row = this.findLastRow();
    const firstRow = this.HAS_HEADER ? 2 : 1;
    if (firstRow > row) return [];
    const values = this.SHEET.getRange(this.START_COL+firstRow + ":" + this.END_COL+(row)).getValues();
    return values.map((value, i) => this.build(value, firstRow + Number(i)));
  }
  
  findByKey(key) {
    let row = findKey_(this.SHEET, key, this.START_COL);
    return this.findByRow(row);
  }
  
  findByRow(row) {
    let values = this.SHEET.getRange(this.START_COL+row + ":" + this.END_COL+row).getValues();
    return this.build(values[0], row);
  }
  
  save(model) {
    // flatten to model values for the record
    let values = [getModelValues_(model, this.KEYS, this.CONVERTERS)];

    // get the key value
    let keyValue = values[0][0].getText();
    
    // grab the document lock for read and write consistency
    let lock = LockService.getDocumentLock();
    lock.waitLock(10000);

    //if this requires a generated key and the key value isn't set, generate the key
    keyValue = keyValue || !this.SEQUENCE ? keyValue : incrementKey_(this.SEQUENCE);
    model[this.KEYS[0]] = keyValue;
    // convert the key back to a rich text value
    values[0][0] = this.CONVERTERS[this.KEYS[0]](keyValue);

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
    SpreadsheetApp.flush();
    lock.releaseLock();
    model["row"] = row;
    return this.ENRICHER ? this.ENRICHER(model) : model;
  }
  
  bulkSave(models) {
    // flatten the models to a 2D array
    const values = models.map(model => getModelValues_(model, this.KEYS, this.CONVERTERS));

    //todo - check for duplicate keys
    if (!this.SEQUENCE) {
      
    }

    // get the lock - we need to do this before any reads to guarantee both read and write consistency
    let lock = LockService.getDocumentLock();
    lock.waitLock(10000);

    // get a map of the existing keys. in the sheet
    const existing = this.findAll();
    const existingValues = existing.map(model => getModelValues_(model, this.KEYS, this.CONVERTERS));
    const rowForKey = existingValues.reduce((keys, list, i) => Object.assign(keys, {[list[0].getText()]: existing[i].row}), {});
    const newValues = values.filter(value => !rowForKey[value[0].getText()]);

    //this creates an array of objects that already exist (by key) in the sheet with the row assigned to the values in each array entry
    let updatedValues = values.map((value, i) => rowForKey[value[0].getText()] ? {"row": rowForKey[value[0].getText()], "values" : value} : {"newRecord":true}).filter(value => !value.newRecord);
    //and then sorts them by row
    updatedValues.sort((a, b) => {return a.row - b.row});

    //Time to update the existing models. While it's only one line of code, Using a record by record approach
    //is slow because it looks up the key and finds the row and does the validation for each model over again.
    //for even 30 models with less 10 fields I've timed this process as taking more than 10 seconds
    //updatedValues.forEach(record => this.save(record.model));

    //Instead the records have been sorted according to their row positions in the spreadsheet, so instead we can go
    //through and group them into contiguous sets of records with a starting row.
    let updatedRecordSets = [];
    let lastRow = -1;
    updatedValues.forEach(record => {
      let recordSet = record.row > lastRow + 1 ? {"row": record.row, "values": []} : updatedRecordSets.pop();
      recordSet.values.push(record.values);
      lastRow = record.row;
      updatedRecordSets.push(recordSet);
    })

    //Now that we have this new data structure we can for each one insert the record sets as a block of values to optimise our write times.
    updatedRecordSets.forEach(recordSet => {
      this.SHEET.getRange(`${this.START_COL}${recordSet.row}:${this.END_COL}${recordSet.row+recordSet.values.length-1}`).setRichTextValues(recordSet.values);
    })

    // if we need to create the keys then create them here
    if (this.SEQUENCE) {
      let lastKey = incrementKey_(this.SEQUENCE, newValues.length);
      newValues = newValues.map((modelValues, i) => [(lastKey - newValues.length + i)].concat(modelValues.slice(1)));
    }

    // save the new records
    if (newValues.length > 0) {
      let row = getFirstEmptyRow_(this.SHEET, this.START_COL);
      this.SHEET.getRange(`${this.START_COL}${row}:${this.END_COL}${row+newValues.length-1}`).setRichTextValues(newValues);
    }

    // and we are done
    SpreadsheetApp.flush();
    lock.releaseLock();
  }

  clear(safety = false) {
    if (safety) {
      const firstRow = this.HAS_HEADER ? 2 : 1;
      const lastRow = this.findLastRow() < firstRow ? firstRow : this.findLastRow();
      this.SHEET.getRange(`${this.START_COL}${firstRow}:${this.END_COL}${lastRow}`).clearContent();
    }
  }
  
  findLastRow() {
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
  return new Dao_(sheet, keys, startCol, hasHeader, enricher, sequence, richTextConverters);
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
  
  //the metadata object to return
  let metadata = {};

  //work out the start column reference
  metadata["startCol"] = cols[start];

  //get the header values from the header row
  let keys = values[0].slice(start, end);

  //convert the header values to camel case keys 
  metadata["keys"] = keys.map(key => toCamelCase_(key));

  //return the inferred metadata
  return metadata;
}

function toCamelCase_(str) {
  const words = str.trim().split(/\s+/);
  return words.map((word, i) => {
    if (i == 0) return word.toLowerCase();
    return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
  }).join('');
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
  return cols.concat(cols.map(col1 => cols.map(col2 => `${col1}${col2}`)).flat(1));
}