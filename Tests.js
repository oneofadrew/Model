
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

function runTests_() {
  let suite = Test.newTestSuite("All Tests")
    .addSetUp(setUp_)
    .addTearDown(tearDown_)
    .addSuite(getSearchSuite_())
  suite.run();
}

function getSearchSuite_() {
  let suite = Test.newTestSuite("Search")
    .addTest(testRunSearch_);
  return suite;
}

function testRunSearch_() {
  let models = [
    {"key":"one","active":true},
    {"key":"two","active":true},
    {"key":"three","active":false},
    {"key":"four","active":true}
  ]
  let search = newSearch().with("active", true);
  found = runSearch(search, models);
  Test.isEqual(found.length, 3);
  Test.isEqual(found[0].key, "one");
  Test.isEqual(found[1].key, "two");
  Test.isEqual(found[2].key, "four");

  search.and("key", "two");
  found = runSearch(search, models);
  Test.isEqual(found.length, 1);
  Test.isEqual(found[0].key, "two");
}