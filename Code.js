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
  }
};

/**
 * Helper functions
 */
function getModelValues_(model, keys, converters) {
    return keys.map((key, i) => converters[i](model[key]));
}

function getFirstEmptyRow_(sheet, col) {
  return findKey_(sheet, "", col);
}

function getRichText_(value) {
  value = value ? value : "";
  return ExternalCalls_.newRichTextValue(value);
}

function findKey_(sheet, key, col) {
  let column = sheet.getRange(col+':'+col);
  let values = column.getValues(); // get all data in one call
  let ct = 0;
  while ( values[ct] && values[ct][0] != key && values[ct][0] != "") {
    ct++;
  }
  if (values[ct] && values[ct][0] == key) return ct + 1;
  throw new Error("Could not find '"+key+"'");
}

function incrementKey_(sequence, increment = 1) {
  const values = sequence.getValues();
  const newValue = values[0][0] + increment;
  sequence.setValues([[newValue]]);
  return newValue;
}