//---------------------------------------------------------------------------------------
// Copyright ⓒ 2024 Drew Harding
// All rights reserved.
//---------------------------------------------------------------------------------------
// Model - a library for turning a google sheet into a simple data table for app script
// Copyright ⓒ 2022 Drew Harding
// All rights reserved.
//
// Script ID: 1pD0LxUmm1NHDrz9fdIikCCOZ_eVS_LP-qgsR5EXReXjh8CLj2xs_I7jF
// GitHub Repo: https://github.com/oneofadrew/Model
//---------------------------------------------------------------------------------------

const ModelLogger = Log.newLog("model.util");
const DaoLogger = Log.newLog("model.dao");
const SearchLogger = Log.newLog("model.search");

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
 * @param {{Range}} sequences - a map of field names to named ranges that each contains a single cell with a number that will be incremented as a sequence for field.
 * @param {{function}} richTextConverters - an map of field names to functions that can takes a field value as an only parameter and returns a RichTextValue object.
 * @param {{string}} formulas - a map of field names to strings that define a sheet formula for use in all rows, for instance {"bill":"=[price][row]*[quantity][row]"}.
 * @param {{DataValidation}} dataValidations - a map of field names to DataValidations that apply to the field.
 * @param {[string]} uniqueKeys - an array of field names that should remain unique across every instance of the model.
 * @param {number} maxLength - the maximum number of rows the table can be. This created a bounded table.
 * @return {object} a map of options for use in the createDao(...) and inferDao(...) functions.
 */
function buildOptions(enricher, sequences, richTextConverters, formulas, dataValidations, uniqueKeys, maxLength) {
  return {
    "enricher": enricher,
    "sequences": sequences,
    "richTextConverters": richTextConverters,
    "formulas": formulas,
    "dataValidations": dataValidations,
    "uniqueKeys": uniqueKeys,
    "maxLength": maxLength
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
  ModelLogger.trace("Running inferDao(sheet:'%s', pk:'%s', startCol:'%s', startRow:'%s')", sheet.getName(), primaryKey, startCol, startRow);
  if (ModelLogger.isTraceEnabled()) ModelLogger.trace("Using options %s", JSON.stringify(options));
  const safeOptions = options ? options : {};
  const values = sheet.getRange(`${startRow}:${startRow}`).getValues();
  const metadata = inferMetadata_(values, startCol);
  const pk = primaryKey ? primaryKey : metadata.keys[0].slice(0);
  ModelLogger.trace(metadata);
  return createDao(sheet, metadata.keys, pk, metadata.startCol, startRow+1, safeOptions);
}

/**
 * Allows for configuration of the Log library.
 * Script ID: 13RAf81luI1DJwKXIeWvK2daYsTN2Rnl2IE1oY_j156tEnNaVaXdRlg9O
 * Code: https://github.com/oneofadrew/Log
 */
function configureLog(config, dumpConfig) {
  Log.setConfig(config);
  if (dumpConfig) Log.dumpConfig();
}

/*
 * Rich Text Converters
 */

/**
 * This gets a RichTextConverter that has a link to a URL. The converter will be called every
 * time there is a write to the spreadsheet for a value in this field to write the value with
 * a link to the calculated URL from calcUrlFn.
 * @param {Function} calcUrlFn - the function that calculates the URL for a value. It takes a single parameter that is the value for the cell.
 * @return {RichTextConverter} A converter to rich text value with the calculated URL as a link
 */
function getUrlConverter(calcUrlFn) {
  return (value) => {
    const safeValue = (value === null || value === undefined) ? "" : value;
    const url = calcUrlFn(value);
    ModelLogger.debug(`URL for '%s' calculated to be %s`, value, url);
    return SpreadsheetApp.newRichTextValue().setText(safeValue).setLinkUrl(url).build();
  }
}

/*
 * Helper functions
 */

/*
 * Looks at a single row dataset of values and tries to infer the startCol and keys properties of a DAO.
 * The logic will start at the column reference provided and look left until it finds the first populated
 * cell. It will then continue until it finds the next empty cell.
 */
function inferMetadata_(values, col) {
  //where do we start having header values
  let cols = getColumnReferences_();
  let startCol = cols.indexOf(col);
  startCol = startCol < 0 ? 0 : startCol;
  ModelLogger.debug("Start Col = '%s'", startCol);

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

/*
 * Converts the model object into an array of values in the order of the defined keys.
 */
function getModelValues_(model, keys) {
  let values = [];
  for (let i=0;i<keys.length;i++) values[values.length] = model[keys[i]];
  return values;
}

/*
 * Looks for a unique value below the cell at the column and row provided.
 */
function findKey_(sheet, key, col, row) {
  ModelLogger.trace(`Running findKey_(sheet:'%s', key:'%s', col:'%s', row:'%s')`, sheet.getName(), key, col, row);
  //account for difference between index and range row values
  row = row - 1;
  ModelLogger.trace("Getting all the values");
  const values = sheet.getDataRange().getValues();
  ModelLogger.trace("All the values have been retrieved");

  let first = -1;
  let total = 0;
  ModelLogger.trace("Looking for key '%s'...", key);
  for (let i=row;i<values.length;i++) {
    if (values[i][col] === key || (values[i][col].getTime && values[i][col].getTime() === key.getTime())) {
      total++;
      if (first<0) first = i+1;
    }
  }
  ModelLogger.trace("Found %s records for key '%s' at row '%s'", total, key, first);
  if (total === 1) return first;
  else throw new Error(`Could not find record with key'${key}'`)
}

/*
 * Takes a named range of a single cell and adds the increment and returns the value. The new value
 * is saved back to the cell of named range.
 */
function incrementKey_(sequence, increment = 1) {
  const values = sequence.getValues();
  const newValue = values[0][0] + increment;
  ModelLogger.debug(`Incrementing sequence from '%s' to '%s'`, values[0][0], newValue);
  sequence.setValues([[newValue]]);
  return newValue;
}

/*
 * Creates a map of substitutes for the row values in a formula.
 */
function buildSubstitutes_(row, startRow) {
  return {
    "[firstRow]":startRow,
    "[previousRow]":row-1,
    "[row]":row,
  };
}

/*
 * Process a string with a set of substitute replacements. This allows a string act as a template
 * which can be used in multiple contexts based on the substitutes provided.
 */
function processStringTemplate_(template, substitutes) {
  return Object.keys(substitutes).reduce((tmpl, k) => {return tmpl.replaceAll(k, substitutes[k])}, template);
}

/*
 * Converts a provided string to standard camel case, with the first letter in lowerCase.
 */
function toCamelCase_(str) {
  const clean = str.replaceAll("'", "");
  const words = clean.trim().split(/[^a-zA-Z\d]+/);
  return words.map((word, i) => {
    if (i === 0) return word.toLowerCase();
    return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
  }).join('');
}

/*
 * Gets the columns for a set of keys to be used as substitution in formula templates.
 */
function getKeyColMap_(startCol, keys) {
  const allCols = getColumnReferences_();
  const start = allCols.indexOf(startCol);
  const cols = allCols.slice(start, start + keys.length);
  return keys.reduce((refsByKey, key, i) => {return Object.assign(refsByKey, {[key]: cols[i]})}, {});
}

/*
 * Calculates the last column based on the first column and a length.
 */
function calculateEndColumn_(startCol, length) {
  let cols = getColumnReferences_();
  let startIndex = cols.indexOf(startCol);
  let endIndex = startIndex + length - 1;
  if (startIndex === -1) throw new Error(`Invalid startCol '${startCol}' provided.`);
  if (endIndex > 701) throw new Error('The Model library only supports models that go up to column ZZ');
  return cols[endIndex]; 
}

/*
 * Gets an array of every column reference from "A" through to "ZZ".
 */
function getColumnReferences_() {
  let cols = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
  return cols.concat(cols.map(col1 => cols.map(col2 => `${col1}${col2}`)).flat(1));
}