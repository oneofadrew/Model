/**
 * A Data Access Object wraps up a set of functions to allow easily interactivity across a model
 */
class Dao_ {
  constructor(sheet, keys, primaryKey, startCol=A, startRow=2, options) {
    const colRefs = getColumnReferences_();
    const safeOptions = options ? options : {};

    this.SHEET = sheet;
    this.KEYS = keys;

    this.START_COL = startCol;
    this.SCI = colRefs.indexOf(startCol);

    this.PK = primaryKey ? primaryKey : keys[0];
    this.PKI = keys.indexOf(this.PK);
    this.PK_COL = colRefs[this.SCI + this.PKI];

    this.START_ROW = startRow;

    this.KEY_COLS_MAP = getKeyColMap_(this.START_COL, this.KEYS);
    this.END_COL = calculateEndColumn_(startCol, keys.length);
    
    this.ENRICHER = safeOptions["enricher"];
    this.SEQUENCE = safeOptions["sequence"];

    const converters = safeOptions["richTextConverters"] ? safeOptions["richTextConverters"] : {};
    this.CONVERTERS = keys.reduce((safeConverters, key) => Object.assign(safeConverters, {[key]: converters[key] ? converters[key] : getRichText_}), {});

    const safeFormulas = safeOptions["formulas"] ? safeOptions["formulas"] : {};
    this.FORMULAS = Object.keys(safeFormulas).reduce((fObj, key) => {
        const f = processStringTemplate_(safeFormulas[key], this.KEY_COLS_MAP);
        const formula = f.substring(0,1) === "=" ? f : `=${f}`;
        return Object.assign(fObj, {[key] : formula})
      }, {}
    );
  }

  /**
   * Creates a model object from a row dataset. The row can also be provided to include in the object.
   */
  build(values, row) {
    const start = row ? {"row": row} : {};
    const model = this.KEYS.reduce((model, key, i) => Object.assign(model, {[key]: values[i]}), start);
    return this.ENRICHER ? this.ENRICHER(model) : model;
  }

  /**
   * 
   */
  findAll() {
    const row = this.findLastRow();
    if (this.START_ROW > row) return [];
    const values = this.SHEET.getRange(`${this.START_COL}${this.START_ROW}:${this.END_COL}${row}`).getValues();
    return values.map((value, i) => this.build(value, this.START_ROW + Number(i)));
  }
  
  /**
   * 
   */
  findByKey(key) {
    let row = findKey_(this.SHEET, key, this.PK_COL, this.START_ROW);
    return this.findByRow(row);
  }
  
  /**
   * 
   */
  findByRow(row) {
    let values = this.SHEET.getRange(`${this.START_COL}${row}:${this.END_COL}${row}`).getValues();
    if (!values[0][this.PKI]) throw new Error(`Could not find model at row ${row}`)
    return this.build(values[0], row);
  }
  
  /**
   * 
   */
  save(model) {
    // flatten to model values for the record
    let values = [getModelValues_(model, this.KEYS, this.CONVERTERS)];

    // get the key value
    let keyValue = values[0][this.PKI].getText();
    
    // grab the document lock for read and write consistency
    let lock = LockService.getDocumentLock();
    lock.waitLock(10000);

    //if this requires a generated key and the key value isn't set, generate the key
    keyValue = keyValue || !this.SEQUENCE ? keyValue : incrementKey_(this.SEQUENCE);
    model[this.KEYS[this.PKI]] = keyValue;
    // convert the key back to a rich text value
    values[0][this.PKI] = this.CONVERTERS[this.KEYS[this.PKI]](keyValue);

    // try to find a record to update based on the key, otherwise we'll create a new record
    let row;
    let willCreate = false;
    try {
      row = findKey_(this.SHEET, keyValue, this.START_COL, this.START_ROW);
    } catch (e) {
      willCreate = true;
      row = getFirstEmptyRow_(this.SHEET, this.START_COL, this.START_ROW);
    }

    if (model.row && willCreate) {
      throw new Error(`The model doesn't exist but has a value for its row property present (${model.row})`);
    }

    if (model.row && model.row !== row) {
      throw new Error(`The row of the model (${model.row}) did not match the row of the model with primary key '${values[0][this.PKI].getText()}' (${row})`);
    }

    // write the values for the record
    this.SHEET.getRange(this.START_COL+row + ":" + this.END_COL+row).setRichTextValues(values);
    // add any formulas
    const substitutes = buildSubstitutes_(row, this.START_ROW);
    Object.keys(this.FORMULAS).forEach((key) => {
      const formula = [processStringTemplate_(this.FORMULAS[key], substitutes)];
      const cell = this.SHEET.getRange(`${this.KEY_COLS_MAP[`[${key}]`]}${row}`);
      cell.setValue(formula);
    });
    
    // and we are done
    SpreadsheetApp.flush();
    lock.releaseLock();
    return this.findByRow(row);
  }
  
  /**
   * 
   */
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
      //create an array of row numbers for the recordSet.
      const rows = Array.from({length: recordSet.values.length}, (_, i) => i + recordSet.row);
      //set the rich text values for the recordSet
      this.SHEET.getRange(`${this.START_COL}${rows[0]}:${this.END_COL}${rows[rows.length-1]}`).setRichTextValues(recordSet.values);
      //add the formulas
      const substitutesList = rows.map(row => buildSubstitutes_(row, this.START_ROW));
      Object.keys(this.FORMULAS).forEach((key) => {
        const formulas = substitutesList.map(substitutes => [processStringTemplate_(this.FORMULAS[key], substitutes)]);
        const range = this.SHEET.getRange(`${this.KEY_COLS_MAP[`[${key}]`]}${rows[0]}:${this.KEY_COLS_MAP[`[${key}]`]}${rows[rows.length-1]}`);
        range.setValues(formulas);
      });
    })

    // if we need to create the keys then create them here
    if (this.SEQUENCE) {
      let lastKey = incrementKey_(this.SEQUENCE, newValues.length);
      newValues = newValues.map((modelValues, i) => [(lastKey - newValues.length + i)].concat(modelValues.slice(1)));
    }

    // save the new records
    if (newValues.length > 0) {
      const rows = Array.from({"length": newValues.length}, (_, i) => i + getFirstEmptyRow_(this.SHEET, this.START_COL, this.START_ROW));
      this.SHEET.getRange(`${this.START_COL}${rows[0]}:${this.END_COL}${rows[rows.length-1]}`).setRichTextValues(newValues);
      //add the formulas
      const substitutesList = rows.map(row => buildSubstitutes_(row, this.START_ROW));
      Object.keys(this.FORMULAS).forEach((key) => {
        const formulas = substitutesList.map(substitutes => [processStringTemplate_(this.FORMULAS[key], substitutes)]);
        const range = this.SHEET.getRange(`${this.KEY_COLS_MAP[`[${key}]`]}${rows[0]}:${this.KEY_COLS_MAP[`[${key}]`]}${rows[rows.length-1]}`);
        range.setValues(formulas);
      });
    }

    // and we are done
    SpreadsheetApp.flush();
    lock.releaseLock();
  }

  /**
   * 
   */
  clear() {
    // get the lock - we need to do this before any reads to guarantee both read and write consistency
    let lock = LockService.getDocumentLock();
    lock.waitLock(10000);

    const lastRow = this.findLastRow() < this.START_ROW ? this.START_ROW : this.findLastRow();
    this.SHEET.getRange(`${this.START_COL}${this.START_ROW}:${this.END_COL}${lastRow}`).clearContent();

    // and we are done
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
  
  /**
   * 
   */
  findLastRow() {
    return getFirstEmptyRow_(this.SHEET, this.PK_COL, this.START_ROW) - 1;
  }
  
  /**
   * 
   */
  search(terms) {
    let models = this.findAll();
    return runSearch(terms, models);
  }
}

/**
 * Create a new Data Access Object from the properties provided. 
 * @param {Sheet} sheet - the sheet that contains the data for the Data Access Object.
 * @param {[string]} keys - the list of keys to use as the field names in the model object.
 * @param {string} primaryKey - the field to use as the primary key for the Data Access Object (defaults to first key in the list).
 * @param {string} startCol - the column to start looking for field names from (usually "A").
 * @param {int} startRow - the row to use to start saving data. This should be the row after any header values if they exist (usually 2).
 * @param {{object}} options - extra configuration options, documented by function Model.buildOptions(...).
 * @return {Dao} a data access object that encapsulates the data access functions based on the properties provided.
 */
function createDao(sheet, keys, primaryKey, startCol, startRow, options) {
  return new Dao_(sheet, keys, primaryKey, startCol, startRow, options);
}

/**
 * Helper method to build the options.
 * It's possible to define formula fields in a model by adding the formula string in a map against the field name for use in every row. Placeholders
 * are surrounded by []. Valid placeholders are field names and [row], [firstRow], and [previousRow]. The field will be replaced with calculated values
 * when the model is returned/retrieved. Formulas can be complex and error prone due to the mental model associated with using them with a DAO. Where
 * your data is a function of the existing data within the object, consider using an enricher function instead.
 * @param {function} enricher - a function that takes a model object as an only parameter, enriches it with other data and then returns it for use.
 * @param {string} sequence - a named range for a single cell that contains a number that will be incremented as a sequenced ID for the data model.
 * @param {{function}} richTextConverters - an map of field names to functions that can takes a field value as an only parameter and returns a RichTextValue object.
 * @param {{string}} formulas - a map of field names to strings that define a sheet formula for use in all rows, for instance {"bill":"=[price][row]*[quantity][row]"}.
 * @return {object} a map of options for use in the createDao(...) and inferDao(...) functions.
 */
function buildOptions(enricher, sequence, richTextConverters, formulas) {
  return {
    "enricher": enricher,
    "sequence": sequence,
    "richTextConverters": richTextConverters,
    "formulas": formulas
  };
}

/**
 * Create a new Data Access Object that infers the metadata from the data in the sheet. The first row in the sheet must be a header row. The start column
 * will be inferred to be the first column from the left that has a header value. The end column will be inferred to be column before the first column
 * after the start column that has no header value. Fields will be inferred to be the titles in the header row for each column changed to camel case.
 * @param {Sheet} sheet - the sheet that contains the data for the Data Access Object.
 * @param {string} primaryKey - the field to use as the primary key for the Data Access Object (defaults to first key found).
 * @param {{object}} options - extra configuration options, documented by function Model.buildOptions(...).
 * @param {string} startCol - the column to start looking for field names from (defaults to "A").
 * @param {int} startRow - the row to use for field names (defaults to 1).
 * @return {Dao} a data access object that encapsulates the data access functions based on the inferred keys and start column.
 */
function inferDao(sheet, primaryKey, options, startCol="A", startRow=1) {
  const safeOptions = options ? options : {};
  const values = sheet.getRange(`${startRow}:${startRow}`).getValues();
  const metadata = inferMetadata_(values, startCol);
  const pk = primaryKey ? primaryKey : metadata.keys[0].slice(0);
  return createDao(sheet, metadata.keys, pk, metadata.startCol, startRow+1, safeOptions);
}

/**
 * Looks at a single row dataset of values and tries to infer the startCol and keys properties of a DAO.
 * The logic will start at the column reference provided and look left until it finds the first populated
 * cell. It will then continue until it finds the next empty cell.
 */
function inferMetadata_(values, col) {
  //where do we start having header values
  let cols = getColumnReferences_();
  let startCol = cols.indexOf(col);
  startCol = startCol < 0 ? 0 : startCol;

  //todo - handle start row as well
  while (values[0][startCol] === '') startCol++;

  //where do we end having header values
  let endCol = startCol;
  while (values[0][endCol] !== '' && values[0].length >= endCol) endCol++;
  
  //the metadata object to return
  let metadata = {};

  //work out the start column reference
  metadata["startCol"] = cols[startCol];

  //get the header values from the header row
  let keys = values[0].slice(startCol, endCol);

  //convert the header values to camel case keys 
  metadata["keys"] = keys.map(key => toCamelCase_(key));

  //return the inferred metadata
  return metadata;
}