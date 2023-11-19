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
 * 
 * cause an exception, but also won't be returned
 * @param {Function} calcUrlFn - the function that calculates the URL for a value
 * @return {RichTextConverter} The rich text value for the 
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
  Logger.log(keys);
  Logger.log(converters);
  Logger.log(model);
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

  //no key found - throw and error
  throw new Error(`Could not find '${key}'`);
}

function test() {
  let key = null;
  let pos = testFn(key);
  Logger.log(`${key} at position ${pos}`);
}

function testFn(key) {
  let values = [["key1"],["key2"],["key3"],["value4"],["key5"], [], []];
  
  const keysByPos = values.map((value, i) => [i, value[0]]);
  const pos = keysByPos.filter(value => value[1] == key);
  //we are looking for the first empty row here
  if (!key) return pos[0][0] + 1;
  if (pos.length == 1) return pos[0][0] + 1;
  throw new Error(`Could not find '${key}'`);
}

function incrementKey_(sequence, increment = 1) {
  const values = sequence.getValues();
  const newValue = values[0][0] + increment;
  sequence.setValues([[newValue]]);
  return newValue;
}