
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