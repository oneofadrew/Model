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
Log.setConfig({"default": {"level":"INFO"}});

/**
 * Allows for configuration of the Log library.
 * Script ID: 13RAf81luI1DJwKXIeWvK2daYsTN2Rnl2IE1oY_j156tEnNaVaXdRlg9O
 * Code: https://github.com/oneofadrew/Log
 */
function configureLog(config) {
  Log.setConfig(config);
  Log.dumpConfig();
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

/**
 * Converts the model object into an array of values in the order of the defined keys.
 */
function getModelValues_(model, keys) {
  let values = [];
  for (let i=0;i<keys.length;i++) values[values.length] = model[keys[i]];
  return values;
}

/**
 * Get's the first empty row below the cell at the column and row provided.
 */
function getFirstEmptyRow_(sheet) {
  return sheet.getDataRange().getLastRow() + 1;
}

/**
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
    if (values[i][col] === key) {
      total++;
      if (first<0) first = i+1;
    }
  }
  ModelLogger.trace("Found %s records for key '%s' at row '%s'", total, key, first);
  if (total === 1) return first;
  else throw new Error(`Could not find '${key}'`)
}

/**
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

/**
 * Creates a map of substitutes for the row values in a formula.
 */
function buildSubstitutes_(row, startRow) {
  return {
    "[firstRow]":startRow,
    "[previousRow]":row-1,
    "[row]":row,
  };
}

/**
 * Process a string with a set of substitute replacements. This allows a string act as a template
 * which can be used in multiple contexts based on the substitutes provided.
 */
function processStringTemplate_(template, substitutes) {
  return Object.keys(substitutes).reduce((tmpl, k) => {return tmpl.replaceAll(k, substitutes[k])}, template);
}

/**
 * Converts a provided string to standard camel case, with the first letter in lowerCase.
 */
function toCamelCase_(str) {
  const words = str.trim().split(/\s+/);
  return words.map((word, i) => {
    if (i === 0) return word.toLowerCase();
    return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
  }).join('');
}

/**
 * Gets the columns for a set of keys to be used as substitution in formula templates.
 */
function getKeyColMap_(startCol, keys) {
  const allCols = getColumnReferences_();
  const start = allCols.indexOf(startCol);
  const cols = allCols.slice(start, start + keys.length);
  return keys.reduce((refsByKey, key, i) => {return Object.assign(refsByKey, {[key]: cols[i]})}, {});
}

/**
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

/**
 * Gets an array of every column reference from "A" through to "ZZ".
 */
function getColumnReferences_() {
  let cols = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
  return cols.concat(cols.map(col1 => cols.map(col2 => `${col1}${col2}`)).flat(1));
}