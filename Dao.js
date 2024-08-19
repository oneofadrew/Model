//---------------------------------------------------------------------------------------
// Copyright ⓒ 2024 Drew Harding
// All rights reserved.
//---------------------------------------------------------------------------------------
/*
 * A Data Access Object wraps up a set of functions to allow easily interactivity across a model
 */
class Dao_ {
  constructor(sheet, keys, primaryKey, startCol, startRow, options) {
    DaoLogger.trace("keys:'%s', pk:'%s', startCol:'%s', startRow:'%s'", keys, primaryKey, startCol, startRow);
    const colRefs = getColumnReferences_();
    const safeOptions = options ? options : {};

    this.SHEET = sheet;
    this.KEYS = keys;

    this.START_COL = startCol ? startCol : "A";
    this.SCI = colRefs.indexOf(this.START_COL);

    if (primaryKey && keys.indexOf(primaryKey) < 0) throw new Error(`Primary key '${primaryKey}' is not one of the keys to use: [${keys}]`);

    this.PK = primaryKey ? primaryKey : keys[0];
    this.PKI = keys.indexOf(this.PK);
    this.PK_COL = colRefs[this.SCI + this.PKI];
    this.PKCI = colRefs.indexOf(this.PK_COL);
    DaoLogger.trace("PK Column '%s' found at column index '%s'", this.PK_COL, this.SCI+this.PKI);

    this.START_ROW = startRow ? startRow : 2;

    this.KEY_COLS_MAP = getKeyColMap_(this.START_COL, this.KEYS);
    this.END_COL = calculateEndColumn_(this.START_COL, keys.length);
    this.ECI = colRefs.indexOf(this.END_COL);
    
    this.ENRICHER = safeOptions["enricher"];

    this.SEQUENCES = safeOptions["sequences"] ? safeOptions["sequences"] : {};
    Object.keys(this.SEQUENCES).forEach(key => {
      if (keys.indexOf(key)<0) throw new Error(`Key '${key}' in sequences does not exist in keys [${keys}]`);
    });

    this.CONVERTERS = safeOptions["richTextConverters"] ? safeOptions["richTextConverters"] : {};
    Object.keys(this.CONVERTERS).forEach(key => {
      if (keys.indexOf(key)<0) throw new Error(`Key '${key}' in converters does not exist in keys [${keys}]`);
    });

    this.VALIDATIONS = safeOptions["dataValidations"] ? safeOptions["dataValidations"] : {};
    Object.keys(this.VALIDATIONS).forEach(key => {
      if (keys.indexOf(key)<0) throw new Error(`Key '${key}' in validations does not exist in keys [${keys}]`);
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
    
    this.UNIQUE_KEYS = safeOptions["uniqueKeys"] ? safeOptions["uniqueKeys"] : [];
    this.UNIQUE_KEYS.forEach(key => {
      if (keys.indexOf(key)<0) throw new Error(`Key '${key}' in unique keys does not exist in keys [${keys}]`);
    });
    //the primary key must be unique - add it to the unique keys if it's not there.
    if (this.UNIQUE_KEYS.indexOf(this.PK) < 0) this.UNIQUE_KEYS[this.UNIQUE_KEYS.length] = this.PK;
  }

  /*
   * Creates a model object from a row dataset. The row can also be provided to include in the object.
   */
  build(values, row) {
    let model = row ? {"row": row} : {};
    for (let i in this.KEYS) model[this.KEYS[i]] = values[i];
    return this.ENRICHER ? this.ENRICHER(model) : model;
  }

  /*
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
  
  /*
   * Returns the model object from the sheet identified by the primary key.
   */
  findByKey(key) {
    let row = findKey_(this.SHEET, key, this.PKCI, this.START_ROW);
    return this.findByRow(row);
  }
  
  /*
   * Returns the model object from the specified row in the sheet.
   */
  findByRow(row) {
    let values = this.SHEET.getRange(`${this.START_COL}${row}:${this.END_COL}${row}`).getValues();
    if (!values[0][this.PKI] && values[0][this.PKI] !== 0) throw new Error(`Could not find model at row ${row}`)
    return this.build(values[0], row);
  }
  
  /*
   * Saves the model object. This will update if it already exists (according to primary key)
   * and create if it doesn't.
   */
  save(model) {
    if (DaoLogger.isTraceEnabled()) {
      DaoLogger.trace("Saving model %s", JSON.stringify(model));
    }

    // grab the document lock for read and write consistency
    DaoLogger.trace(`Grab the document lock`);
    let lock = LockService.getDocumentLock();
    lock.waitLock(10000);

    // generate sequences if necessary
    DaoLogger.trace(`Processing sequences %s`, JSON.stringify(this.SEQUENCES));
    Object.keys(this.SEQUENCES).forEach(key => {
      DaoLogger.trace(`Increment the sequence for %s`, key);
      if (!model[key]) model[key] = incrementKey_(this.SEQUENCES[key]);
    });

    // flatten to model values for the record
    let values = [getModelValues_(model, this.KEYS)];

    // get the key value
    let keyValue = values[0][this.PKI];
    if (!keyValue) {throw new Error(`Could find a value for primary key ${this.PK} in the model ${JSON.stringify(model)}`)};
    DaoLogger.trace(`Saving model with key '%s'`, keyValue);

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

    DaoLogger.trace(`Add in any data validations.`);
    for (let key in this.VALIDATIONS) {
      DaoLogger.trace(`Getting the range for the data validation of key '%s'.`, key);
      const cell = this.SHEET.getRange(`${this.KEY_COLS_MAP[key]}${row}`);
      DaoLogger.trace(`Setting the data validation.`);
      cell.setDataValidation(this.VALIDATIONS[key]);
    }
    DaoLogger.trace(`Validations all done.`);

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
  
  /*
   * Saves the list of model object. This will update if it already exists (according to primary key)
   * and create if it doesn't. Bulk save optimises the operation by using set processing of objects
   * in contiguous rows where possible.
   */
  bulkSave(models) {
    DaoLogger.debug(`Starting the bulk save by flattening models to values.`);

    //check the ones to create don't already exist
    for (let model of models)
      if (!model[this.PK] && model["row"]) throw new Error(`The model doesn't exist but has a value for its row property present (${model["row"]})`);

    // flatten the models to a 2D array
    const values = models.map(model => getModelValues_(model, this.KEYS));

    // process all the sequence fields
    const sequenceFields = Object.keys(this.SEQUENCES);
    if (sequenceFields.length > 0) {
      // create the counter objects
      let sequenceCounts = {};

      for (let field of sequenceFields) sequenceCounts[field] = {"index": this.KEYS.indexOf(field), "count": 0};
      
      // count the number of sequences to generate for each of the fields
      for (let record of values)
        for (let field of sequenceFields)
          // if the record doesn't have the sequence field then count it for incrementing the sequence,
          if (!record[sequenceCounts[field]["index"]]) sequenceCounts[field]["count"]++;

      // generate the sequences
      for (let field of sequenceFields)
        if (sequenceCounts[field]["count"])
          //the next counter = the last value for the key after incrementing - count to increment + 1
          sequenceCounts[field]["next"] = incrementKey_(this.SEQUENCES[field], sequenceCounts[field]["count"]) - sequenceCounts[field]["count"] + 1;

      // populate the sequences in the records
      for (let record of values)
        for (let field of sequenceFields)
          // if the sequence field is empty add the next value into it and increment the next value
          if (!record[sequenceCounts[field]["index"]]) record[sequenceCounts[field]["index"]] = sequenceCounts[field]["next"]++;
    }

    //todo - check for duplicate keys

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
    //todo
    //for (let i=0;i<existingKeys.length;i++) primaryKeys[primaryKeys.length] = existingKeys[i][0].getTime ? existingKeys[i][0].getTime() : existingKeys[i][0];
    for (let i=0;i<existingKeys.length;i++) primaryKeys[primaryKeys.length] = existingKeys[i][0];

    /*
    for (let model of models) {
      if (model["row"] && model["row"] !== primaryKeys.indexOf(model[this.PK])+this.START_ROW)
        throw new Error(`The row of the model (${model["row"]}) did not match the row of the model with primary key '${model[this.PK]}' (${primaryKeys.indexOf(model[this.PK])+this.START_ROW})`);
    };
    */

    //sort the models into those to update and those to create
    DaoLogger.debug(`Sort the models into those to be updated and those to be created.`);
    let updatedValues = [];
    let newValues = [];
    //todo
    //let pkValue = values[0].length;
    for (let i=0;i<values.length;i++) {
      //todo
      //values[i][pkValue] = values[i][this.PKI] && values[i][this.PKI].getTime ? values[i][this.PKI].getTime() : values[i][this.PKI];
      //if (primaryKeys.indexOf(values[i][pkValue]) > -1) updatedValues[updatedValues.length] = values[i];
      if (primaryKeys.indexOf(values[i][this.PKI]) > -1) updatedValues[updatedValues.length] = values[i];
      else newValues[newValues.length] = values[i];
    }

    // check for new records that have a row attribute
    let newRecordsByKey = {};
    for (let newRecord of newValues) if (newRecord[this.PKI]) newRecordsByKey[newRecord[this.PKI]] = newRecord;

    for (let model of models) {
      const newRecord = model[this.PK] ? newRecordsByKey[model[this.PK]] : null;
      if ((newRecord && model["row"]) || (!model[this.PK] && model["row"])) throw new Error(`The model doesn't exist but has a value for its row property present (${model["row"]})`);
    }

    /*
    for (let newRecord of newValues) {
      const key = newRecord[this.PKI];
      const model = modelsByKey[key];
      if (model && model["row"]) throw new Error(`The model doesn't exist but has a value for its row property present (${model["row"]})`);
    }
    */

    // check for updated records where the wrong doesn't match the current row
    let modelsByKey = {};
    for (let model of models) if (model[this.PK]) modelsByKey[model[this.PK]] = model;

    for (let updatedRecord of updatedValues) {
      const key = updatedRecord[this.PKI];
      const row = primaryKeys.indexOf(key) + this.START_ROW;
      const model = modelsByKey[key];
      if (model["row"] && model["row"] !== row) throw new Error(`The row of the model (${model["row"]}) did not match the row of the model with primary key '${key}' (${row})`);
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
      updatedRecordSets[updatedRecordSets.length] = recordSet;
    }

    //process each of the update record sets
    DaoLogger.debug(`Process each batch.`);
    for (let i=0;i<updatedRecordSets.length;i++) {
      DaoLogger.trace(`Processing batch %s with %s records - start by getting the range`, i, updatedRecordSets[i]["values".length]);
      const firstRow = updatedRecordSets[i]["row"];

      DaoLogger.trace(`Add in any data validations.`);
      for (let key in this.VALIDATIONS) {
        DaoLogger.trace(`Set the data validations for key '%s'.`, key);
        const range = this.SHEET.getRange(`${this.KEY_COLS_MAP[key]}${firstRow}:${this.KEY_COLS_MAP[key]}${firstRow+updatedRecordSets[i]["values"].length-1}`);
        range.setDataValidation(this.VALIDATIONS[key]);
      }

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

    // save the new records
    if (newValues.length > 0) {
      DaoLogger.trace(`Saving '%s' new models - start by getting the range.`, newValues.length);

      DaoLogger.trace(`Add in any data validations.`);
      for (let key in this.VALIDATIONS) {
        DaoLogger.trace(`Set the data validations for key '%s'.`, key);
        const range = this.SHEET.getRange(`${this.KEY_COLS_MAP[key]}${firstEmptyRow}:${this.KEY_COLS_MAP[key]}${firstEmptyRow+newValues.length-1}`);
        range.setDataValidation(this.VALIDATIONS[key]);
      }
      
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

  /*
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

    this.SHEET.getDataRange().clear({"contentsOnly": true, "formatsOnly": true, "validationsOnly": true});

    if (this.START_ROW > 0) titleRange.setValues(titles)

    // and we are done
    DaoLogger.debug(`Flush and unlock.`);
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
  
  /*
   * Returns the last row of the model object table.
   */
  findLastRow() {
    return getFirstEmptyRow_(this.SHEET) - 1;
  }
  
  /*
   * Convenience method to Run a search over all the objects in the sheet. Model.runSearch() can be
   * used to run a more targeted search on a specific list of models if desired.
   */
  search(terms) {
    let models = this.findAll();
    return runSearch(terms, models);
  }
}