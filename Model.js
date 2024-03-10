/**
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

/**
 * Helper functions
 */
function getModelValues_(model, keys, converters) {
  return keys.map(key => converters[key](model[key]));
}

function getFirstEmptyRow_(sheet, col, row) {
  return findKey_(sheet, "", col, row);
}

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

function incrementKey_(sequence, increment = 1) {
  const values = sequence.getValues();
  const newValue = values[0][0] + increment;
  sequence.setValues([[newValue]]);
  return newValue;
}