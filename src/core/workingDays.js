/**
 * Streamline Working Days Module
 * Handles working/non-working day configuration, weekend detection,
 * and working-day duration calculations.
 */

// Day-of-week: 0=Sunday, 1=Monday, ..., 6=Saturday
const DAY_NAMES = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

/**
 * Default working days configuration.
 */
const DEFAULT_WORKING_DAYS = {
  // Array of 7 booleans: [Sun, Mon, Tue, Wed, Thu, Fri, Sat]
  days: [false, true, true, true, true, true, false],
  // Custom holidays (array of Date objects)
  holidays: [],
  // Whether to highlight weekends on the timescale
  highlightWeekends: true,
  // Color for weekend shading (null = use template default)
  weekendColor: null,
  // Whether to skip weekends in duration calculations
  skipWeekendsInDuration: false,
};

/**
 * Check if a date falls on a working day.
 * @param {Date} date
 * @param {Object} config - Working days config
 * @returns {boolean}
 */
function isWorkingDay(date, config = DEFAULT_WORKING_DAYS) {
  if (!date) return false;
  const dow = date.getDay();
  if (!config.days[dow]) return false;

  // Check holidays
  if (config.holidays && config.holidays.length > 0) {
    const dateStr = date.toISOString().slice(0, 10);
    for (const holiday of config.holidays) {
      const holStr = holiday instanceof Date ? holiday.toISOString().slice(0, 10) : holiday;
      if (holStr === dateStr) return false;
    }
  }

  return true;
}

/**
 * Check if a date falls on a weekend (Saturday or Sunday).
 */
function isWeekend(date) {
  if (!date) return false;
  const dow = date.getDay();
  return dow === 0 || dow === 6;
}

/**
 * Calculate working-day duration between two dates.
 * Skips non-working days per the config.
 */
function getWorkingDays(startDate, endDate, config = DEFAULT_WORKING_DAYS) {
  if (!startDate || !endDate) return 0;

  let count = 0;
  const current = new Date(startDate);
  current.setHours(0, 0, 0, 0);
  const end = new Date(endDate);
  end.setHours(0, 0, 0, 0);

  while (current <= end) {
    if (isWorkingDay(current, config)) count++;
    current.setDate(current.getDate() + 1);
  }

  return count;
}

/**
 * Add a number of working days to a date.
 * Skips non-working days.
 */
function addWorkingDays(startDate, daysToAdd, config = DEFAULT_WORKING_DAYS) {
  const result = new Date(startDate);
  result.setHours(0, 0, 0, 0);

  if (daysToAdd === 0) return result;

  const direction = daysToAdd > 0 ? 1 : -1;
  let remaining = Math.abs(daysToAdd);

  while (remaining > 0) {
    result.setDate(result.getDate() + direction);
    if (isWorkingDay(result, config)) remaining--;
  }

  return result;
}

/**
 * Generate weekend shading regions for a date range.
 * Returns array of { startDate, endDate, type } regions for rendering.
 */
function getWeekendRegions(minDate, maxDate, config = DEFAULT_WORKING_DAYS) {
  if (!config.highlightWeekends) return [];

  const regions = [];
  const current = new Date(minDate);
  current.setHours(0, 0, 0, 0);
  const end = new Date(maxDate);
  end.setHours(23, 59, 59, 999);

  let regionStart = null;

  while (current <= end) {
    const isNonWorking = !isWorkingDay(current, config);

    if (isNonWorking && !regionStart) {
      regionStart = new Date(current);
    } else if (!isNonWorking && regionStart) {
      const regionEnd = new Date(current);
      regions.push({
        type: "weekendRegion",
        startDate: regionStart,
        endDate: regionEnd,
      });
      regionStart = null;
    }

    current.setDate(current.getDate() + 1);
  }

  // Close any open region
  if (regionStart) {
    regions.push({
      type: "weekendRegion",
      startDate: regionStart,
      endDate: new Date(current),
    });
  }

  return regions;
}

/**
 * Parse working days from a string format like "Mon-Fri" or "Sun,Mon,Tue".
 */
function parseWorkingDays(str) {
  const days = [false, false, false, false, false, false, false];
  if (!str) return days;

  const short = { sun: 0, mon: 1, tue: 2, wed: 3, thu: 4, fri: 5, sat: 6 };
  const normalized = str.toLowerCase().trim();

  // Handle "Mon-Fri" format
  const rangeMatch = normalized.match(/^(\w{3})-(\w{3})$/);
  if (rangeMatch) {
    const start = short[rangeMatch[1]];
    const end = short[rangeMatch[2]];
    if (start !== undefined && end !== undefined) {
      if (start <= end) {
        for (let i = start; i <= end; i++) days[i] = true;
      } else {
        // Wraparound (e.g. Fri-Mon)
        for (let i = start; i < 7; i++) days[i] = true;
        for (let i = 0; i <= end; i++) days[i] = true;
      }
      return days;
    }
  }

  // Handle comma-separated (e.g. "Mon,Tue,Wed")
  const parts = normalized.split(/[,\s]+/);
  for (const part of parts) {
    const d = short[part.substring(0, 3)];
    if (d !== undefined) days[d] = true;
  }

  return days;
}

/**
 * Preset working day configurations.
 */
const WORKING_DAY_PRESETS = {
  standard: {
    name: "Standard (Mon-Fri)",
    days: [false, true, true, true, true, true, false],
  },
  sixDay: {
    name: "Six-day (Mon-Sat)",
    days: [false, true, true, true, true, true, true],
  },
  sevenDay: {
    name: "Seven-day (All)",
    days: [true, true, true, true, true, true, true],
  },
  middleEast: {
    name: "Middle East (Sun-Thu)",
    days: [true, true, true, true, true, false, false],
  },
  fourDay: {
    name: "Four-day (Mon-Thu)",
    days: [false, true, true, true, true, false, false],
  },
};

module.exports = {
  DEFAULT_WORKING_DAYS,
  DAY_NAMES,
  WORKING_DAY_PRESETS,
  isWorkingDay,
  isWeekend,
  getWorkingDays,
  addWorkingDays,
  getWeekendRegions,
  parseWorkingDays,
};
