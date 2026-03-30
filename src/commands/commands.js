/* global Office */

Office.onReady(() => {});

function openTaskpane(event) {
  Office.addin.showAsTaskpane();
  event.completed();
}

window.openTaskpane = openTaskpane;
