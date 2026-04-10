/**
 * Streamline Command Runtime Entry
 *
 * This is the webpack entry point for the hidden function-file bundle that
 * Copilot for PowerPoint and the ribbon commands invoke. It registers every
 * command with Office.actions.associate() as soon as Office JS is ready.
 * No UI is rendered here - commands.html loads this bundle into a hidden
 * iframe and the functions run invisibly in the background.
 */

const { registerCommands } = require("./functionCommands");

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(() => {
    registerCommands();
  });
}

// Export for tests / manual invocation
module.exports = { registerCommands };
