sap.ui.define(
    [
        "./BaseController"
    ],
    function(BaseController) {
      "use strict";
  
      return BaseController.extend("warehousemanagement.controller.App", {
        onInit: function() {
          debugger;
        },
        onItemSelect: function(oEvent){
          var sKey = oEvent.getParameter('item').getKey();
          debugger;
          this.getRouter().navTo(sKey);
        }
      });
    }
  );
  