function runTests() {
  let suite = Test.newTestSuite("All Tests")
    .addSuite(getBuilderSuite_())
    .addSuite(getSearchSuite_())
    .addSuite(getDaoSuite_());
  suite.run();
}

/* ----------------------------------------------------------------------------------------- */

function getBuilderSuite_() {
  let suite = Test.newTestSuite("Builders")
    .addTest(testGetCellColMap_)
    .addTest(testCalculateEndCol_);
  return suite;
}

function testGetCellColMap_() {
  const keys = ["key1", "key2", "key3", "key4", "key5", "key6"];
  const mapKeys = ["[key1]", "[key2]", "[key3]", "[key4]", "[key5]", "[key6]"];
  const result1 = ["A", "B", "C", "D", "E", "F"];
  const result2 = ["F", "G", "H", "I", "J", "K"];

  cellColMap1 = getCellColMap_("A", keys);
  mapKeys.forEach((key, i) => Test.isEqual(cellColMap1[key], result1[i]));
  cellColMap2 = getCellColMap_("F", keys);
  mapKeys.forEach((key, i) => Test.isEqual(cellColMap2[key], result2[i]));
}

function testCalculateEndCol_() {
  Test.isEqual(calculateEndColumn_('A', 5), 'E');
  Test.isEqual(calculateEndColumn_('AA', 5), 'AE');
  Test.isEqual(calculateEndColumn_('EA', 5), 'EE');
  Test.isEqual(calculateEndColumn_('ZA', 5), 'ZE');
  Test.isEqual(calculateEndColumn_('A', 31), 'AE');
  Test.isEqual(calculateEndColumn_('AA', 31), 'BE');
  Test.isEqual(calculateEndColumn_('EA', 31), 'FE');
  Test.isEqual(calculateEndColumn_('YA', 31), 'ZE');
  Test.isEqual(calculateEndColumn_('A', 702), 'ZZ');
  Test.willFail(()=>{calculateEndColumn_('AAA', 1)});
  Test.willFail(()=>{calculateEndColumn_('A', 703)});
  Test.willFail(()=>{calculateEndColumn_('B', 702)});
}

/* ----------------------------------------------------------------------------------------- */

function getSearchSuite_() {
  let suite = Test.newTestSuite("Search")
    .addTest(testRunSearch_);
  return suite;
}

function testRunSearch_() {
  let models = [
    {"id":1,"key":"one","active":true},
    {"id":2,"key":"two","active":false},
    {"id":3,"key":"two","active":true},
    {"id":4,"key":"three","active":true}
  ]
  let search = newSearch().where("active", true);
  let found = runSearch(search, models);
  Test.isEqual(found.length, 3);
  Test.isEqual(found[0].id, 1);
  Test.isEqual(found[1].id, 3);
  Test.isEqual(found[2].id, 4);

  search.and("key", "two");
  found = runSearch(search, models);
  Test.isEqual(found.length, 1);
  Test.isEqual(found[0].id, 3);

  search.and("foo", "bar");
  found = runSearch(search, models);
  Test.isEqual(found.length, 0);
}

/* ----------------------------------------------------------------------------------------- */

function getDaoSuite_() {
  let suite = Test.newTestSuite("Dao")
    .addTest(testCreateDao_)
    .addTest(testBuildNotEnrichedWithRow_)
    .addTest(testBuildNotEnrichedNoRow_)
    .addTest(testBuildEnrichedWithRow_)
    .addTest(testBuildEnrichedNoRow_)
    .addTest(testInferMetadata_)
    .addTest(testBuildSubstitutes_)
    .addTest(testBuildFormula_)
    .addTest(testToCamelCase_);
  return suite;
}

function testCreateDao_() {
  const keys = ["key1", "key2", "key3", "key4"];
  const primaryKey = "key2";
  const sheet = "mySheet";
  const enricher = (m) => {return m;};
  const sequence = "mySequence";
  const converter = () => {return true;};
  const converters = {"key4": converter};
  const formulas = {"key2" : "=[key1][row]+[key3][row]", "key4" : "$A$1 + [key4][previousRow] + [key3][row]"};
  const options = buildOptions(enricher, sequence, converters, formulas, false);
  const dao = createDao(sheet, keys, primaryKey, "G", 3, options);
  Test.isEqual(dao.SHEET, sheet);
  Test.isEqual(dao.KEYS, keys);
  Test.isEqual(dao.PK, "key2");
  Test.isEqual(dao.PKI, 1);
  Test.isEqual(dao.PK_COL, "H");
  Test.isEqual(dao.START_COL, "G");
  Test.isEqual(dao.SCI, 6);
  Test.isEqual(dao.START_ROW, 3);
  Test.isEqual(dao.KEY_COLS_MAP, {"[key1]": "G", "[key2]": "H", "[key3]": "I", "[key4]": "J"});
  Test.isEqual(dao.END_COL, "J");
  Test.isEqual(dao.ENRICHER, enricher);
  Test.isEqual(dao.SEQUENCE, sequence);
  Test.isEqual(Object.keys(dao.CONVERTERS).length, keys.length);
  Test.isEqual(dao.CONVERTERS["key4"], converter);
  Test.isTrue(dao.CONVERTERS["key4"]());
  Test.isEqual(Object.keys(dao.FORMULAS).length, 2);
  Test.isEqual(dao.FORMULAS["key2"], "=G[row]+I[row]");
  Test.isEqual(dao.FORMULAS["key4"], "=$A$1 + J[previousRow] + I[row]");
}

function testBuildNotEnrichedWithRow_() {
  const keys = ["key1", "key2", "key3", "key4"];
  const primaryKey = "key1";
  const values = ["value1", "value2", "value3", "value4"];
  const sheet = "mySheet";
  const dao = createDao(sheet, keys, primaryKey, "A", 1);
  const model = dao.build(values, 5);
  Test.isEqual(model["key1"], "value1");
  Test.isEqual(model["key2"], "value2");
  Test.isEqual(model["key3"], "value3");
  Test.isEqual(model["key4"], "value4");
  Test.isEqual(model["row"], 5);
}

function testBuildNotEnrichedNoRow_() {
  const keys = ["key1", "key2", "key3", "key4"];
  const primaryKey = "key1";
  const values = ["value1", "value2", "value3", "value4"];
  const sheet = "mySheet";
  const dao = createDao(sheet, keys, primaryKey, "A", 1);
  const model = dao.build(values);
  Test.isEqual(model["key1"], "value1");
  Test.isEqual(model["key2"], "value2");
  Test.isEqual(model["key3"], "value3");
  Test.isEqual(model["key4"], "value4");
  Test.isEmpty(model["row"]);
}

function testBuildEnrichedWithRow_() {
  const keys = ["key1", "key2", "key3", "key4"];
  const primaryKey = "key1";
  const values = ["value1", "value2", "value3", "value4"];
  const sheet = "mySheet";
  const enricher = (m) => {
    m["key3"] = "newValue3";
    m["key5"] = "value5";
    return m;
  };
  const dao = createDao(sheet, keys, primaryKey, "A", 1, {"enricher": enricher});
  const model = dao.build(values, 6);
  Test.isEqual(model["key1"], "value1");
  Test.isEqual(model["key2"], "value2");
  Test.isEqual(model["key3"], "newValue3");
  Test.isEqual(model["key4"], "value4");
  Test.isEqual(model["key5"], "value5");
  Test.isEqual(model["row"], 6);
}

function testBuildEnrichedNoRow_() {
  const keys = ["key1", "key2", "key3", "key4"];
  const primaryKey = "key1";
  const values = ["value1", "value2", "value3", "value4"];
  const sheet = "mySheet";
  const enricher = (m) => {
    m["key3"] = "newValue3";
    m["key5"] = "value5";
    return m;
  };
  const dao = createDao(sheet, keys, primaryKey, "A", 1, {"enricher": enricher});
  const model = dao.build(values);
  Test.isEqual(model["key1"], "value1");
  Test.isEqual(model["key2"], "value2");
  Test.isEqual(model["key3"], "newValue3");
  Test.isEqual(model["key4"], "value4");
  Test.isEqual(model["key5"], "value5");
  Test.isEmpty(model["row"]);
}

function testInferMetadata_() {
  let header = [['Email', 'First Name', 'Last Name', 'Mobile Number', 'Client Folder', 'Invoice Folder', 'Client Information Form', 'Consent Form', 'DASS 21', 'Charge Amount']];
  metadata = inferMetadata_(header);
  Test.isEqual(metadata.startCol, 'A');
  Test.isEqual(metadata.keys, ["email", "firstName", "lastName", "mobileNumber",  "clientFolder", "invoiceFolder", "clientInformationForm", "consentForm", "dass21","chargeAmount"]);
  metadata = inferMetadata_(header, 'E');
  Test.isEqual(metadata.startCol, 'E');
  Test.isEqual(metadata.keys, ["clientFolder", "invoiceFolder", "clientInformationForm", "consentForm", "dass21","chargeAmount"]);

  header = [['', 'Email', 'First Name', 'Last Name', '', 'Mobile Number']];
  metadata = inferMetadata_(header);
  Test.isEqual(metadata.startCol, 'B');
  Test.isEqual(metadata.keys, ["email", "firstName", "lastName"]);
  metadata = inferMetadata_(header, 'E');
  Test.isEqual(metadata.startCol, 'F');
  Test.isEqual(metadata.keys, ["mobileNumber"]);
  metadata = inferMetadata_(header, 'F');
  Test.isEqual(metadata.startCol, 'F');
  Test.isEqual(metadata.keys, ["mobileNumber"]);
}

function testBuildSubstitutes_() {
  let subs1 = buildSubstitutes_(6, 2);
  Test.isEqual(subs1["[firstRow]"], 2);
  Test.isEqual(subs1["[previousRow]"], 5);
  Test.isEqual(subs1["[row]"], 6);

  let subs2 = buildSubstitutes_(100, 1);
  Test.isEqual(subs2["[firstRow]"], 1);
  Test.isEqual(subs2["[previousRow]"], 99);
  Test.isEqual(subs2["[row]"], 100);
  
  let subs3 = buildSubstitutes_(23, 12);
  Test.isEqual(subs3["[firstRow]"], 12);
  Test.isEqual(subs3["[previousRow]"], 22);
  Test.isEqual(subs3["[row]"], 23);
}

function testBuildFormula_() {
  const substitutes = buildSubstitutes_(10, 2);

  const fTemplate1 = '=row(A[row]) + col(A[previousRow])';
  const formula1 = buildFormula_(fTemplate1, substitutes);
  Test.isEqual(formula1, '=row(A10) + col(A9)');

  const fTemplate2 = '=sum(A[firstRow]:A[previousRow])';
  const formula2 = buildFormula_(fTemplate2, substitutes);
  Test.isEqual(formula2, '=sum(A2:A9)');
}

function testToCamelCase_() {
  Test.isEqual(toCamelCase_("EquipmentClass name"), 'equipmentclassName');
  Test.isEqual(toCamelCase_("Equipment className"), 'equipmentClassname');
  Test.isEqual(toCamelCase_("equipment class name"), 'equipmentClassName');
  Test.isEqual(toCamelCase_("Equipment Class Name"), 'equipmentClassName');
  Test.isEqual(toCamelCase_("DASS 21"), 'dass21');
}

/* ----------------------------------------------------------------------------------------- */


function test() {
  const keyColsMap = {"[key1]" : "A", "[key2]" : "B", "[key3]" : "C", "[key4]" : "D", "[key5]" : "E"};
  const formulas = {"key2" : "=[key1][row]+[key3][row]", "key4" : "$G$1 + [key4][previousRow] + [key3][row]"};
  const mappedFormulas = Object.keys(formulas).reduce((fObj, fKey) => {
      const f = buildFormula_(formulas[fKey], keyColsMap)[0];
      const formula = f.substring(0,1) === "=" ? f : `=${f}`;
      return Object.assign(fObj, {[fKey] : formula})
    }, {}
  );

  const row = 23;

  const substitutes = buildSubstitutes_(row, 2);
  Object.keys(mappedFormulas).forEach((key) => {
    const formula = buildFormula_(mappedFormulas[key], substitutes);
    const cell = `${keyColsMap[key]}${row}`;
    Logger.log(`Putting formula ${formula} into cell ${cell}`);
  });

  const rows = [23, 24, 25, 26, 27];
  const substitutesList = rows.map(row => buildSubstitutes_(row, 2));
  Object.keys(mappedFormulas).forEach((key) => {
    const formulas = substitutesList.map(substitutes => buildFormula_(mappedFormulas[key], substitutes));
    const range = `${keyColsMap[key]}${rows[0]}:${keyColsMap[key]}${rows[rows.length-1]}:`;
    Logger.log(`Putting formulas ${formulas} into range ${range}`);
  });
}