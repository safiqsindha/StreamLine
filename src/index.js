/**
 * Streamline - PowerPoint Add-in Entry Point
 * Initializes the Office Add-in and wires up the task pane controller.
 */

require("./ui/taskpane.css");

const { TaskPaneController } = require("./ui/taskpaneController");

// Wait for Office JS to be ready
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    const controller = new TaskPaneController();
    controller.init();
    console.log("Streamline add-in initialized.");
  } else {
    document.getElementById("status-bar").classList.remove("hidden");
    document.getElementById("status-bar").classList.add("error");
    document.getElementById("status-message").textContent =
      "Streamline requires PowerPoint. Please open this add-in in PowerPoint.";
  }
});
