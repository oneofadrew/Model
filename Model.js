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
    return SpreadsheetApp.newRichTextValue().setText(safeValue).setLinkUrl(url).build();
  }
}

/** The default converter - will be used for any key that doesn't have a converter provided */
function getRichText_(value) {
  const safeValue = (value === null || value === undefined) ? "" : value;
  return SpreadsheetApp.newRichTextValue().setText(safeValue).build();
}

/*
 * Helper functions
 */

/**
 * 
 */
function getModelValues_(model, keys, converters) {
  return keys.map(key => converters[key](model[key]));
}

/**
 * Get's the first empty row below the cell at the column and row provided.
 */
function getFirstEmptyRow_(sheet, col, row) {
  return findKey_(sheet, "", col, row);
}

/**
 * Looks for a unique value below the cell at the column and row provided.
 */
function findKey_(sheet, key, col, row) {
  const column = sheet.getRange(`${col}${row}:${col}`);
  const values = column.getValues(); // get all data in one call
  const keysByPos = values.map((value, i) => [i+row, value[0]]);
  const pos = keysByPos.filter(value => value[1] === key);

  //if the key is falsey we looking for the first empty row here
  //so could be multiple records from the filter - just return
  //the first position found
  if (!key) return pos[0][0];

  //we are looking for a specific key, check there is only one
  //and return its position
  if (pos.length === 1) return pos[0][0];

  //no key found - throw an error
  throw new Error(`Could not find '${key}'`);
}

/**
 * Takes a named range of a single cell and adds the increment and returns the value. The new value
 * is saved back to the cell of named range.
 */
function incrementKey_(sequence, increment = 1) {
  const values = sequence.getValues();
  const newValue = values[0][0] + increment;
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
  return keys.reduce((refsByKey, key, i) => {return Object.assign(refsByKey, {[`[${key}]`]: cols[i]})}, {});
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