/**
 * A Data Access Object wraps up a set of functions to allow easily interactivity across a model
 */
class Dao_ {
  constructor(sheet, keys, primaryKey, startCol=A, startRow=2, options) {
    DaoLogger.trace("keys:'%s', pk:'%s', startCol:'%s', startRow:'%s'", keys, primaryKey, startCol, startRow);
    const colRefs = getColumnReferences_();
    const safeOptions = options ? options : {};

    this.SHEET = sheet;
    this.KEYS = keys;

    this.START_COL = startCol;
    this.SCI = colRefs.indexOf(startCol);

    if (primaryKey && keys.indexOf(primaryKey) < 0) throw new Error(`Primary key '${primaryKey}' is not one of the keys to use: [${keys}]`);

    this.PK = primaryKey ? primaryKey : keys[0];
    this.PKI = keys.indexOf(this.PK);
    this.PK_COL = colRefs[this.SCI + this.PKI];
    this.PKCI = colRefs.indexOf(this.PK_COL);
    DaoLogger.trace("PK Column '%s' found at column index '%s'", this.PK_COL, this.SCI+this.PKI);

    this.START_ROW = startRow;

    this.KEY_COLS_MAP = getKeyColMap_(this.START_COL, this.KEYS);
    this.END_COL = calculateEndColumn_(startCol, keys.length);
    this.ECI = colRefs.indexOf(this.END_COL);
    
    this.ENRICHER = safeOptions["enricher"];
    this.SEQUENCE = safeOptions["sequence"];
    this.CONVERTERS = safeOptions["richTextConverters"] ? safeOptions["richTextConverters"] : {};
    Object.keys(this.CONVERTERS).forEach(key => {
      if (keys.indexOf(key)<0) throw new Error(`Key '${key}' in converters does not exist in keys [${keys}]`);
    });

    const safeFormulas = safeOptions["formulas"] ? safeOptions["formulas"] : {};
    const subs = Object.keys(this.KEY_COLS_MAP).reduce((obj, key) => { return Object.assign(obj, {[`[${key}]`]: this.KEY_COLS_MAP[key]})}, {});
    this.FORMULAS = Object.keys(safeFormulas).reduce((fObj, key) => {
        if (keys.indexOf(key)<0) throw new Error(`Key '${key}' in formulas does not exist in keys [${keys}]`);
        const f = processStringTemplate_(safeFormulas[key], subs);
        const formula = f.substring(0,1) === "=" ? f : `=${f}`;
        return Object.assign(fObj, {[key] : formula})
      }, {}
    );
  }

  /**
   * Creates a model object from a row dataset. The row can also be provided to include in the object.
   */
  build(values, row) {
    let model = row ? {"row": row} : {};
    for (let i in this.KEYS) model[this.KEYS[i]] = values[i];
    return this.ENRICHER ? this.ENRICHER(model) : model;
  }

  /**
   * Returns all of the model objects available from the sheet.
   */
  findAll() {
    const row = this.findLastRow();
    if (this.START_ROW > row) return [];
    const values = this.SHEET.getRange(`${this.START_COL}${this.START_ROW}:${this.END_COL}${row}`).getValues();
    let models = [];
    for (let i=0;i<values.length;i++) models[models.length] = this.build(values[i], (this.START_ROW + i));
    return models;
  }
  
  /**
   * Returns the model object from the sheet identified by the primary key.
   */
  findByKey(key) {
    let row = findKey_(this.SHEET, key, this.PKCI, this.START_ROW);
    return this.findByRow(row);
  }
  
  /**
   * Returns the model object from the specified row in the sheet.
   */
  findByRow(row) {
    let values = this.SHEET.getRange(`${this.START_COL}${row}:${this.END_COL}${row}`).getValues();
    if (!values[0][this.PKI]) throw new Error(`Could not find model at row ${row}`)
    return this.build(values[0], row);
  }
  
  /**
   * Saves the model object. This will update if it already exists (according to primary key)
   * and create if it doesn't.
   */
  save(model) {
    DaoLogger.trace(`Saving model`);
    // flatten to model values for the record
    let values = [getModelValues_(model, this.KEYS)];

    // get the key value
    let keyValue = values[0][this.PKI];
    DaoLogger.trace(`Saving model with key '%s'`, keyValue);
    
    // grab the document lock for read and write consistency
    DaoLogger.trace(`Grab the document lock`);
    let lock = LockService.getDocumentLock();
    lock.waitLock(10000);

    //if this requires a generated key and the key value isn't set, generate the key
    keyValue = keyValue || !this.SEQUENCE ? keyValue : incrementKey_(this.SEQUENCE);
    model[this.KEYS[this.PKI]] = keyValue;

    // try to find a record to update based on the key, otherwise we'll create a new record
    let row;
    let willCreate = false;
    try {
      DaoLogger.trace(`Looking for an existing row for this key`);
      row = findKey_(this.SHEET, keyValue, this.PKCI, this.START_ROW);
      DaoLogger.trace(`Found row %s`, row);
    } catch (e) {
      DaoLogger.trace(`Not found, creating a new one (error: %s).`, e.message);
      willCreate = true;
      row = getFirstEmptyRow_(this.SHEET);
      DaoLogger.trace(`Using first empty row: %s`, row);
    }

    DaoLogger.trace(`Validation on the PK to model contents.`);
    if (model.row && willCreate) {
      throw new Error(`The model doesn't exist but has a value for its row property present (${model.row})`);
    }

    if (model.row && model.row !== row) {
      throw new Error(`The row of the model (${model.row}) did not match the row of the model with primary key '${values[0][this.PKI]}' (${row})`);
    }

    // write the values for the record
    DaoLogger.trace(`Getting the record range for the model.`);
    const rangeRef = `${this.START_COL}${row}:${this.END_COL}${row}`;
    const range = this.SHEET.getRange(rangeRef);

    //add the formulas
    DaoLogger.trace(`Add in all the formulas.`);
    for (let key in this.FORMULAS) {
      DaoLogger.trace(`Processing formular '%s' for key '%s.`, this.FORMULAS[key], key);
      let kIndex = this.KEYS.indexOf(key);
      let subs = buildSubstitutes_(row, this.START_ROW);
      values[0][kIndex] = processStringTemplate_(this.FORMULAS[key], subs);
    }
    DaoLogger.trace(`Now save the values to the range.`);
    range.setValues(values);

    //todo - add any rich text
    DaoLogger.trace(`Add in any rich text.`);
    for (let key in this.CONVERTERS) {
      DaoLogger.trace(`Processing rich text converter for key '%s'.`, key);
      const kIndex = this.KEYS.indexOf(key);
      const richTextValue = this.CONVERTERS[key](values[0][kIndex]);
      DaoLogger.trace(`Getting the range for the rich text converter.`);
      const cell = this.SHEET.getRange(`${this.KEY_COLS_MAP[key]}${row}`);
      DaoLogger.trace(`Setting the rich text value.`);
      cell.setRichTextValue(richTextValue);
    }
    DaoLogger.trace(`Rich text all done.`);
    
    // and we are done
    DaoLogger.trace(`Flush and unlock.`);
    SpreadsheetApp.flush();
    lock.releaseLock();
    return this.findByRow(row);
  }
  
  /**
   * Saves the list of model object. This will update if it already exists (according to primary key)
   * and create if it doesn't. Bulk save optimises the operation by using set processing of objects
   * in contiguous rows where possible.
   */
  bulkSave(models) {
    DaoLogger.debug(`Starting the bulk save by flattening models to values.`);
    // flatten the models to a 2D array
    const values = models.map(model => getModelValues_(model, this.KEYS));

    //todo - check for duplicate keys
    if (!this.SEQUENCE) {}

    // get the lock - we need to do this before any reads to guarantee both read and write consistency
    DaoLogger.debug(`Grab the document lock.`);
    let lock = LockService.getDocumentLock();
    lock.waitLock(10000);
    
    //get the first empty row
    DaoLogger.debug(`Get the last row of the dataset.`);
    const firstEmptyRow = getFirstEmptyRow_(this.SHEET);

    //Get the existing keys from the sheet
    const rangeRef = `${this.PK_COL}${this.START_ROW}:${this.PK_COL}${firstEmptyRow}`;
    DaoLogger.debug(`Grabbing the values for the keys to the sheet from '%s'.`, rangeRef);
    const existingKeys = this.SHEET.getRange(rangeRef).getValues();

    //flatten the key from an array of arrays to an array
    DaoLogger.debug(`Process the keys to a single flat array.`);
    let primaryKeys = [];
    for (let i=0;i<existingKeys.length;i++) primaryKeys[primaryKeys.length] = existingKeys[i][0];

    //sort the models into those to update and those to create
    DaoLogger.debug(`Sort the models into those to be updated and those to be created.`);
    let updatedValues = [];
    let newValues = [];
    for (let i=0;i<values.length;i++) {
      if (primaryKeys.indexOf(values[i][this.PKI]) > -1) updatedValues[updatedValues.length] = values[i];
      else newValues[newValues.length] = values[i];
    }

    //create the update batches based on contiguous rows
    DaoLogger.debug(`Sort the models to be updated into contiguous batches.`);
    let updatedRecordSets = [];
    if (updatedValues.length > 0) {
      let lastRow = primaryKeys.indexOf(updatedValues[0][this.PKI])+this.START_ROW;
      let recordSet = {"row": lastRow, "values": [values[0]]};
      updatedRecordSets = [recordSet];
      for (let i=1;i<updatedValues.length;i++) {
        let row = primaryKeys.indexOf(updatedValues[i][this.PKI])+this.START_ROW;
        if (lastRow === row-1) {
          //the next record is in sequence to the last, add it to the current batch
          recordSet["values"][recordSet["values"].length] = updatedValues[i];
        } else {
          // there's a gap in the rows between the last value set and this one - we need to start a new record set with the values
          updatedRecordSets[updatedRecordSets.length] = recordSet;
          recordSet = {"row": row, "values": [updatedValues[i]]};
        }
        lastRow = row;
      }
    }

    //process each of the update record sets
    DaoLogger.debug(`Process each batch.`);
    for (let i=0;i<updatedRecordSets.length;i++) {
      DaoLogger.trace(`Processing batch %s with %s records - start by getting the range`, i, updatedRecordSets[i]["values".length]);
      const firstRow = updatedRecordSets[i]["row"];
      const rangeRef = `${this.START_COL}${firstRow}:${this.END_COL}${firstRow+updatedRecordSets[i]["values"].length-1}`;
      const range = this.SHEET.getRange(rangeRef);

      //add in the formulas
      DaoLogger.trace(`Add in the formulas to the values.`);
      for (let key in this.FORMULAS) {
        DaoLogger.trace(`Processing formula '%s' for key '%s.`, this.FORMULAS[key], key);
        let kIndex = this.KEYS.indexOf(key);
        for (let j=0;j<updatedRecordSets[i]["values"].length;j++){
          let subs = buildSubstitutes_(j+updatedRecordSets[i]["row"], this.START_ROW);
          updatedRecordSets[i]["values"][j][kIndex] = processStringTemplate_(this.FORMULAS[key], subs);
        }
      }
      DaoLogger.debug(`Save the batch.`);
      range.setValues(updatedRecordSets[i]["values"])

      //todo - add the rich text
      DaoLogger.trace(`Add in any rich text converters.`);
      for (let key in this.CONVERTERS) {
        DaoLogger.trace(`Processing rich text converter for key '%s'.`, key);
        const kIndex = this.KEYS.indexOf(key);
        let richTextValues = [];
        for (let j=0;j<updatedRecordSets[i]["values"].length;j++) richTextValues[richTextValues.length] = [this.CONVERTERS[key](updatedRecordSets[i]["values"][j][kIndex])];
        const range = this.SHEET.getRange(`${this.KEY_COLS_MAP[key]}${firstRow}:${this.KEY_COLS_MAP[key]}${firstRow+updatedRecordSets[i]["values"].length-1}`);
        DaoLogger.trace(`Set the rich text values for key '%s'.`);
        range.setRichTextValues(richTextValues);
      }
    }

    // if we need to create the keys then create them here
    if (this.SEQUENCE) {
      DaoLogger.debug(`Get the next sequence values for models to create.`);
      let lastKey = incrementKey_(this.SEQUENCE, newValues.length);
      newValues = newValues.map((modelValues, i) => [(lastKey - newValues.length + i)].concat(modelValues.slice(1)));
    }

    // save the new records
    if (newValues.length > 0) {
      DaoLogger.trace(`Saving '%s' new models - start by getting the range.`, newValues.length);
      const rangeRef = `${this.START_COL}${firstEmptyRow}:${this.END_COL}${firstEmptyRow+newValues.length-1}`;
      const newRange = this.SHEET.getRange(rangeRef);

      //add the formulas
      DaoLogger.trace(`Add in the formulas for the new values.`);
      for (let key in this.FORMULAS) {
        DaoLogger.trace(`Processing formula '%s' for key '%s'.`, this.FORMULAS[key], key);
        let kIndex = this.KEYS.indexOf(key);
        for (let i=0;i<newValues.length;i++){
          let subs = buildSubstitutes_(i+firstEmptyRow, this.START_ROW);
          newValues[i][kIndex] = processStringTemplate_(this.FORMULAS[key], subs);
        }
      }
      DaoLogger.debug(`Saving the new models.`);
      newRange.setValues(newValues);

      //todo - add any rich text
      DaoLogger.trace(`Look after any rich text converters.`);
      for (let key in this.CONVERTERS) {
        DaoLogger.trace(`Process converters for '%s'.`, key);
        const kIndex = this.KEYS.indexOf(key);
        let richTextValues = [];
        for (let i=0;i<newValues.length;i++) richTextValues[richTextValues.length] = [this.CONVERTERS[key](newValues[i][kIndex])];
        const range = this.SHEET.getRange(`${this.KEY_COLS_MAP[key]}${firstEmptyRow}:${this.KEY_COLS_MAP[key]}${firstEmptyRow+newValues.length-1}`);
        DaoLogger.trace(`Saving the rich text values.`);
        range.setRichTextValues(richTextValues);
      }
    }

    // and we are done
    DaoLogger.debug(`Flush and unlock.`);
    SpreadsheetApp.flush();
    lock.releaseLock();
  }

  /**
   * Wipes the sheet leaving the title row.
   */
  clear() {
    // get the lock - we need to do this before any reads to guarantee both read and write consistency
    let lock = LockService.getDocumentLock();
    lock.waitLock(10000);

    let titleRange;
    let titles;
    let formulas;
    if (this.START_ROW > 0) {
      DaoLogger.debug(`We don't start at row 1, so save everything above the start row.`);
      titleRange = this.SHEET.getRange(`1:${this.START_ROW-1}`);
      titles = titleRange.getValues();
      formulas = titleRange.getFormulas();
      for (let i=0;i<formulas.length;i++) for (let j=0;j<formulas[i].length;j++) titles[i][j] = formulas[i][j] ? formulas[i][j] : titles[i][j];
    }

    this.SHEET.getDataRange().clearContent();

    if (this.START_ROW > 0) titleRange.setValues(titles)

    // and we are done
    DaoLogger.debug(`Flush and unlock.`);
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
  
  /**
   * Returns the last row of the model object table.
   */
  findLastRow() {
    return getFirstEmptyRow_(this.SHEET) - 1;
  }
  
  /**
   * Convenience method to Run a search over all the objects in the sheet. Model.runSearch() can be
   * used to run a more targeted search on a specific list of models if desired.
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
  DaoLogger.trace("Running inferDao(sheet:'%s', pk:'%s', startCol:'%s', startRow:'%s')", sheet.getName(), primaryKey, startCol, startRow);
  DaoLogger.trace(options);
  const safeOptions = options ? options : {};
  const values = sheet.getRange(`${startRow}:${startRow}`).getValues();
  const metadata = inferMetadata_(values, startCol);
  const pk = primaryKey ? primaryKey : metadata.keys[0].slice(0);
  DaoLogger.trace(metadata);
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
  DaoLogger.debug("Start Col = '%s'", startCol);

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