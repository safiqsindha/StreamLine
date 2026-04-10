/**
 * Streamline MS Project Parser
 * Parses MS Project XML export format (.xml) into Streamline task rows.
 * MS Project can export to XML via File > Save As > XML.
 * Binary .mpp is proprietary; we support the interoperable XML format.
 */

/**
 * Parse MS Project XML string into rows compatible with Streamline data model.
 * @param {string} xmlString - Raw XML content from an MS Project XML export
 * @returns {Object} { rows, projectName, calendarName }
 */
function parseMppXml(xmlString) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlString, "text/xml");

  const parseError = doc.querySelector("parsererror");
  if (parseError) {
    throw new Error("Invalid XML file: " + parseError.textContent.substring(0, 100));
  }

  const ns = doc.documentElement.namespaceURI || "";
  const nsPrefix = ns ? `{${ns}}` : "";

  // Helper to query with or without namespace
  function query(parent, tagName) {
    let result = parent.getElementsByTagName(tagName);
    if (result.length === 0 && ns) {
      result = parent.getElementsByTagNameNS(ns, tagName);
    }
    return result;
  }

  function getText(parent, tagName) {
    const els = query(parent, tagName);
    return els.length > 0 ? els[0].textContent.trim() : "";
  }

  function getDate(parent, tagName) {
    const text = getText(parent, tagName);
    if (!text) return null;
    const d = new Date(text);
    return isNaN(d.getTime()) ? null : d;
  }

  function getInt(parent, tagName) {
    const text = getText(parent, tagName);
    const n = parseInt(text, 10);
    return isNaN(n) ? null : n;
  }

  const projectName = getText(doc.documentElement, "Name") || "Untitled Project";
  const calendarName = getText(doc.documentElement, "CalendarUID") || "";

  // Build resource map (UID -> Name) for swim lane assignment
  const resourceMap = new Map();
  const resources = query(doc.documentElement, "Resource");
  for (const res of resources) {
    const uid = getText(res, "UID");
    const name = getText(res, "Name");
    if (uid && name) resourceMap.set(uid, name);
  }

  // Build assignment map (TaskUID -> ResourceName)
  const assignmentMap = new Map();
  const assignments = query(doc.documentElement, "Assignment");
  for (const asn of assignments) {
    const taskUID = getText(asn, "TaskUID");
    const resUID = getText(asn, "ResourceUID");
    if (taskUID && resUID && resourceMap.has(resUID)) {
      assignmentMap.set(taskUID, resourceMap.get(resUID));
    }
  }

  // Parse tasks
  const taskElements = query(doc.documentElement, "Task");
  const taskMap = new Map(); // UID -> { name, outlineLevel, ... }
  const rows = [];

  // First pass: build UID->name map for dependency resolution
  for (const taskEl of taskElements) {
    const uid = getText(taskEl, "UID");
    const name = getText(taskEl, "Name");
    if (uid) taskMap.set(uid, name);
  }

  // Track outline hierarchy for swim lane derivation
  let currentSwimLane = "General";

  for (const taskEl of taskElements) {
    const uid = getText(taskEl, "UID");
    const name = getText(taskEl, "Name");
    const outlineLevel = getInt(taskEl, "OutlineLevel") || 0;
    const isSummary = getText(taskEl, "Summary") === "1";
    const milestone = getText(taskEl, "Milestone") === "1";
    const startDate = getDate(taskEl, "Start");
    const endDate = getDate(taskEl, "Finish");
    const percentComplete = getInt(taskEl, "PercentComplete");
    const baselineStart = getDate(taskEl, "BaselineStart");
    const baselineFinish = getDate(taskEl, "BaselineFinish");

    // Skip project summary (UID 0 or outline level 0)
    if (uid === "0" || (!name && outlineLevel === 0)) continue;

    // Summary tasks at outline level 1 become swim lanes
    if (isSummary && outlineLevel <= 1) {
      currentSwimLane = name || "General";
      continue; // Don't add summary tasks as separate rows
    }

    // Sub-summaries at deeper levels become sub-swimlanes
    let subSwimLane = null;
    if (isSummary && outlineLevel === 2) {
      subSwimLane = name;
      continue;
    }

    // Parse predecessor links
    const predLinks = query(taskEl, "PredecessorLink");
    const depParts = [];
    for (const link of predLinks) {
      const predUID = getText(link, "PredecessorUID");
      const linkType = getInt(link, "Type"); // 0=FF, 1=FS, 2=SF, 3=SS
      const lagDuration = getText(link, "LinkLag"); // in tenths of minutes
      const predName = taskMap.get(predUID) || "";

      if (!predName) continue;

      const typeMap = { 0: "FF", 1: "FS", 2: "SF", 3: "SS" };
      const depTypeStr = typeMap[linkType] || "FS";

      // LinkLag is in tenths of minutes — convert to days
      let lagDays = 0;
      if (lagDuration) {
        const tenthsOfMinutes = parseInt(lagDuration, 10);
        if (!isNaN(tenthsOfMinutes)) {
          lagDays = Math.round(tenthsOfMinutes / (10 * 60 * 24));
        }
      }

      let depStr = predName;
      if (depTypeStr !== "FS" || lagDays !== 0) {
        depStr += ` [${depTypeStr}`;
        if (lagDays !== 0) depStr += `${lagDays > 0 ? "+" : ""}${lagDays}d`;
        depStr += "]";
      }
      depParts.push(depStr);
    }

    // Determine status from percent complete
    let status = null;
    if (percentComplete !== null) {
      if (percentComplete >= 100) status = "Complete";
      else if (percentComplete > 0) status = "On Track";
    }

    // Assigned resource becomes owner
    const owner = assignmentMap.get(uid) || null;

    rows.push({
      swimLane: currentSwimLane,
      subSwimLane: subSwimLane,
      taskName: name,
      type: milestone ? "Milestone" : "Task",
      startDate: startDate,
      endDate: milestone ? null : endDate,
      plannedStartDate: baselineStart,
      plannedEndDate: baselineFinish,
      percentComplete: percentComplete,
      dependency: depParts.join(", "),
      status: status,
      owner: owner,
      notes: "",
      milestoneShape: null,
    });
  }

  return { rows, projectName, calendarName };
}

/**
 * Detect if a file is likely an MS Project XML export.
 * @param {string} xmlString
 * @returns {boolean}
 */
function isMppXml(xmlString) {
  return xmlString.includes("<Project") && xmlString.includes("<Task") &&
    (xmlString.includes("schemas.microsoft.com/project") || xmlString.includes("<UID>"));
}

/**
 * Detect if a file is a binary .mpp file by inspecting the first bytes.
 * MPP is an OLE2 compound document — signature: D0 CF 11 E0 A1 B1 1A E1
 * @param {ArrayBuffer} arrayBuffer
 * @returns {boolean}
 */
function isMppBinary(arrayBuffer) {
  if (!arrayBuffer || arrayBuffer.byteLength < 8) return false;
  const view = new Uint8Array(arrayBuffer, 0, 8);
  return view[0] === 0xD0 && view[1] === 0xCF && view[2] === 0x11 && view[3] === 0xE0 &&
         view[4] === 0xA1 && view[5] === 0xB1 && view[6] === 0x1A && view[7] === 0xE1;
}

/**
 * Attempt to extract project data from a binary .mpp file.
 * Native .mpp is a proprietary OLE compound document with Microsoft's binary
 * streams — fully parsing requires reverse-engineered format knowledge.
 *
 * This function does a best-effort extraction of ASCII task names from the
 * binary, then returns a helpful error pointing the user toward XML export.
 *
 * @param {ArrayBuffer} arrayBuffer
 * @returns {Object} { detected: true, canParse: false, message, extractedNames }
 */
function inspectMppBinary(arrayBuffer) {
  if (!isMppBinary(arrayBuffer)) {
    return { detected: false };
  }

  // Best-effort ASCII extraction from the binary streams
  const view = new Uint8Array(arrayBuffer);
  const extractedNames = [];
  const seen = new Set();
  let current = "";

  for (let i = 0; i < view.length; i++) {
    const byte = view[i];
    // Printable ASCII or common punctuation
    if (byte >= 32 && byte <= 126) {
      current += String.fromCharCode(byte);
    } else {
      if (current.length >= 6 && current.length <= 80 && !seen.has(current)) {
        // Filter likely task names (contain letters, not all caps internal noise)
        if (/^[A-Za-z][A-Za-z0-9 .,:\-_/()&]+$/.test(current) &&
            current.match(/[a-z]/) && !current.startsWith("MSProject")) {
          seen.add(current);
          extractedNames.push(current);
        }
      }
      current = "";
    }
  }

  return {
    detected: true,
    canParse: false,
    extractedNames: extractedNames.slice(0, 20),
    message: "Binary .mpp files are in Microsoft's proprietary format which cannot be parsed in a browser. " +
             "In MS Project, use File > Export > Save Project as File > XML Format to create an .xml file that Streamline can import directly.",
  };
}

/**
 * Universal MS Project import: detects format and dispatches.
 * @param {ArrayBuffer|string} data - Either binary ArrayBuffer or XML string
 * @returns {Object} { rows, projectName } on success, or throws with helpful error
 */
function parseMppFile(data) {
  if (data instanceof ArrayBuffer) {
    if (isMppBinary(data)) {
      const info = inspectMppBinary(data);
      const err = new Error(info.message);
      err.mppBinaryDetected = true;
      err.extractedNames = info.extractedNames;
      throw err;
    }
    // Try as text
    const text = new TextDecoder("utf-8").decode(data);
    if (isMppXml(text)) return parseMppXml(text);
    throw new Error("File format not recognized. Expected MS Project XML export (.xml).");
  }

  if (typeof data === "string") {
    if (isMppXml(data)) return parseMppXml(data);
    throw new Error("Not a valid MS Project XML file. Use File > Export > Save Project as File > XML Format in MS Project.");
  }

  throw new Error("Invalid input to MS Project parser.");
}

module.exports = { parseMppXml, isMppXml, isMppBinary, inspectMppBinary, parseMppFile };
