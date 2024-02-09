/**
 * The ExternalCalls_ object is here to abstract away our external calls to allow us to drive
 * unit tests with mock objects to validate our functionality is working as expected
 */
let ExternalCalls_ = {
  "getDocumentLock" : () => {
    return LockService.getDocumentLock();
  },
  "spreadsheetFlush" : () => {
    SpreadsheetApp.flush();
  },
  "newRichTextValue" : (value) => {
    return SpreadsheetApp.newRichTextValue().setText(value).build();
  },
  "newUrlConverter" : (calcUrlFn) => {
    return (value) => {
      const safeValue = (value == null || value == undefined) ? "" : value;
      const url = calcUrlFn(value);
      return SpreadsheetApp.newRichTextValue().setText(safeValue).setLinkUrl(url).build();
    }
  }
};

/**
 * Rich Text Converters
 */

/**
 * This gets a RichTextConverter that has a link to a URL. The converter will be called every
 * time there is a write to the spreadsheet for a value in this field to write the value with
 * a link to the calculated URL from calcUrlFn.
 * @param {Function} calcUrlFn - the function that calculates the URL for a value
 * @return {RichTextConverter} A converter to rich text value with the calculated URL as a link
 */
function getUrlConverter(calcUrlFn) {
  return ExternalCalls_.newUrlConverter(calcUrlFn);
}

/** The default converter - doesn't need to be called by the  */
function getRichText_(value) {
  const safeValue = (value == null || value == undefined) ? "" : value;
  return ExternalCalls_.newRichTextValue(safeValue);
}

/**
 * Helper functions
 */
function getModelValues_(model, keys, converters) {
  return keys.map(key => converters[key](model[key]));
}

function getFirstEmptyRow_(sheet, col) {
  return findKey_(sheet, "", col);
}

function findKey_(sheet, key, col) {
  const column = sheet.getRange(col+':'+col);
  const values = column.getValues(); // get all data in one call
  const keysByPos = values.map((value, i) => [i, value[0]]);
  const pos = keysByPos.filter(value => value[1] == key);

  //if the key is falsey we looking for the first empty row here
  //so could be multiple records from the filter - just return
  //the first position found
  if (!key) return pos[0][0] + 1;

  //we are looking for a specific key, check there is only one
  //and return its position
  if (pos.length == 1) return pos[0][0] + 1;

  //no key found - throw an error
  throw new Error(`Could not find '${key}'`);
}

function incrementKey_(sequence, increment = 1) {
  const values = sequence.getValues();
  const newValue = values[0][0] + increment;
  sequence.setValues([[newValue]]);
  return newValue;
}