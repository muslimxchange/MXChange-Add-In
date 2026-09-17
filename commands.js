/* MXChange Excel add-in — ribbon command host (no commands beyond the task pane button) */
/* global Office */
Office.onReady(function () {});
Office.actions.associate("action", function (event) { event.completed(); });
