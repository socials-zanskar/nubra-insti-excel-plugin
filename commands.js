(function () {
  "use strict";

  function noopAction(event) {
    if (event && typeof event.completed === "function") {
      event.completed();
    }
  }

  if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
    Office.actions.associate("noopAction", noopAction);
  }
})();
