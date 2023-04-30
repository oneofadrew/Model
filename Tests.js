
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
    .addTearDown(tearDown)
  suite.run();
}
