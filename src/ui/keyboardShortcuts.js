/**
 * Streamline Keyboard Shortcuts
 * Binds keyboard shortcuts to controller actions.
 * Shortcuts only fire when the task pane has focus (not when typing in inputs).
 */

const IS_MAC = typeof navigator !== "undefined" &&
  /Mac|iPhone|iPad|iPod/.test(navigator.platform || navigator.userAgent || "");

const SHORTCUTS = [
  {
    key: "i",
    modifiers: ["ctrl"],
    action: "import",
    description: "Import Excel file",
    display: "Ctrl+I",
  },
  {
    key: "m",
    modifiers: ["ctrl"],
    action: "importMpp",
    description: "Import MS Project XML",
    display: "Ctrl+M",
  },
  {
    key: "v",
    modifiers: ["ctrl", "shift"],
    action: "paste",
    description: "Paste from clipboard",
    display: "Ctrl+Shift+V",
  },
  {
    key: "r",
    modifiers: ["ctrl"],
    action: "refresh",
    description: "Refresh timeline",
    display: "Ctrl+R",
  },
  {
    key: "s",
    modifiers: ["ctrl", "shift"],
    action: "exportPng",
    description: "Export as PNG",
    display: "Ctrl+Shift+S",
  },
  {
    key: "p",
    modifiers: ["ctrl", "shift"],
    action: "exportPdf",
    description: "Export as PDF",
    display: "Ctrl+Shift+P",
  },
  {
    key: "j",
    modifiers: ["ctrl", "shift"],
    action: "exportJpg",
    description: "Export as JPG",
    display: "Ctrl+Shift+J",
  },
  {
    key: "e",
    modifiers: ["ctrl", "shift"],
    action: "exportMpp",
    description: "Export as MS Project XML",
    display: "Ctrl+Shift+E",
  },
  {
    key: "1",
    modifiers: ["ctrl"],
    action: "tabImport",
    description: "Switch to Import tab",
    display: "Ctrl+1",
  },
  {
    key: "2",
    modifiers: ["ctrl"],
    action: "tabEditor",
    description: "Switch to Editor tab",
    display: "Ctrl+2",
  },
  {
    key: "3",
    modifiers: ["ctrl"],
    action: "tabStyle",
    description: "Switch to Style tab",
    display: "Ctrl+3",
  },
  {
    key: "4",
    modifiers: ["ctrl"],
    action: "tabSettings",
    description: "Switch to Settings tab",
    display: "Ctrl+4",
  },
  {
    key: "Delete",
    modifiers: [],
    action: "deleteShape",
    description: "Delete selected shape (when shape selected)",
    display: "Delete",
  },
  {
    key: "Enter",
    modifiers: ["ctrl"],
    action: "applyShapeEdit",
    description: "Apply shape edits",
    display: "Ctrl+Enter",
  },
  {
    key: "/",
    modifiers: ["ctrl"],
    action: "showShortcuts",
    description: "Show keyboard shortcuts",
    display: "Ctrl+/",
  },
  {
    key: "n",
    modifiers: ["ctrl"],
    action: "newRow",
    description: "Add new row in data editor",
    display: "Ctrl+N",
  },
];

class KeyboardShortcutManager {
  constructor(controller) {
    this.controller = controller;
    this.enabled = true;
    this._handler = null;
  }

  start() {
    if (this._handler) return;
    this._handler = (e) => this._onKeyDown(e);
    document.addEventListener("keydown", this._handler);
  }

  stop() {
    if (this._handler) {
      document.removeEventListener("keydown", this._handler);
      this._handler = null;
    }
  }

  _onKeyDown(e) {
    if (!this.enabled) return;

    // Skip if typing in an input field (unless it's a global shortcut like Ctrl+/)
    const target = e.target;
    const isTyping = target && (
      target.tagName === "INPUT" ||
      target.tagName === "TEXTAREA" ||
      target.tagName === "SELECT" ||
      target.isContentEditable
    );

    for (const shortcut of SHORTCUTS) {
      if (!this._matchesShortcut(e, shortcut)) continue;

      // Delete key and text shortcuts inside inputs should not trigger global actions
      if (isTyping && shortcut.action !== "showShortcuts") {
        // Exception: Ctrl+Enter in shape edit fields should still apply
        if (shortcut.action === "applyShapeEdit" && this._isShapeEditField(target)) {
          // Let it through
        } else {
          continue;
        }
      }

      e.preventDefault();
      e.stopPropagation();
      this._dispatch(shortcut.action);
      return;
    }
  }

  _isShapeEditField(el) {
    if (!el) return false;
    return el.id && el.id.startsWith("edit-shape-");
  }

  _matchesShortcut(e, shortcut) {
    // Check key
    if (e.key.toLowerCase() !== shortcut.key.toLowerCase()) return false;

    // Check modifiers
    // On macOS, only Ctrl (not Cmd) triggers our shortcuts so we don't
    // hijack system Cmd+R (reload), Cmd+P (print), Cmd+N (new window), etc.
    // On other platforms either Ctrl or Cmd works.
    const hasCtrl = IS_MAC ? e.ctrlKey : (e.ctrlKey || e.metaKey);
    const hasShift = e.shiftKey;
    const hasAlt = e.altKey;

    const wantCtrl = shortcut.modifiers.includes("ctrl");
    const wantShift = shortcut.modifiers.includes("shift");
    const wantAlt = shortcut.modifiers.includes("alt");

    return hasCtrl === wantCtrl && hasShift === wantShift && hasAlt === wantAlt;
  }

  _dispatch(action) {
    if (!this.controller || typeof this.controller.handleKeyboardShortcut !== "function") {
      return;
    }
    this.controller.handleKeyboardShortcut(action);
  }

  /**
   * Get all shortcuts as a display-ready list.
   */
  getShortcutList() {
    return SHORTCUTS.map((s) => ({
      keys: IS_MAC ? this._displayForMac(s.display) : s.display,
      description: s.description,
    }));
  }

  _displayForMac(display) {
    return display
      .replace(/Ctrl/g, "⌃")
      .replace(/Shift/g, "⇧")
      .replace(/Alt/g, "⌥");
  }
}

module.exports = { KeyboardShortcutManager, SHORTCUTS };
