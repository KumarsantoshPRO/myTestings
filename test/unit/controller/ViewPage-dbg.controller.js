sap.ui.define(["my/app/generatebill/controller/View.controller"], function (__Controller) {
  "use strict";

  function _interopRequireDefault(obj) {
    return obj && obj.__esModule && typeof obj.default !== "undefined" ? obj.default : obj;
  }
  /*global QUnit*/
  const Controller = _interopRequireDefault(__Controller);
  QUnit.module("View Controller");
  QUnit.test("I should test the View controller", function (assert) {
    const oAppController = new Controller("View");
    oAppController.onInit();
    assert.ok(oAppController);
  });
});
//# sourceMappingURL=ViewPage-dbg.controller.js.map
