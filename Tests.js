let externalCallsMock_ = Test.newMock();
let originalExtCallsObj_ = ExternalCalls_;

function setUp_() {
  // reset the external calls mock before each test
  externalCallsMock_ = Test.newMock();
}
function tearDown_() {
  //put the original external calls object back in place (just in case)
  ExternalCalls_ = originalExtCallsObj_;
}

function runUnitTests_() {
  let suite = Test.newTestSuite("All Tests")
    .addSetUp(setUp_)
    .addTearDown(tearDown_)
    .addSuite(getBuilderSuite_())
    .addSuite(getSearchSuite_())
    .addSuite(getDaoSuite_());
  suite.run();
}

/* ----------------------------------------------------------------------------------------- */

function getBuilderSuite_() {
  let suite = Test.newTestSuite("Builders")
    .addTest(testCalculateEndCol_);
  return suite;
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
    .addTest(testInferMetadata_)
    .addTest(testToCamelCase_);
  return suite;
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

function testToCamelCase_() {
  Test.isEqual(toCamelCase_("EquipmentClass name"), 'equipmentclassName');
  Test.isEqual(toCamelCase_("Equipment className"), 'equipmentClassname');
  Test.isEqual(toCamelCase_("equipment class name"), 'equipmentClassName');
  Test.isEqual(toCamelCase_("Equipment Class Name"), 'equipmentClassName');
  Test.isEqual(toCamelCase_("DASS 21"), 'dass21');
}

/* ----------------------------------------------------------------------------------------- */
