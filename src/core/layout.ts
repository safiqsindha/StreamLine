// The single layout authority. The grid, the SVG timeline and every export
// read positions from here. No other module computes x or y.
//
//   x = (date − rangeStart) × pixelsPerDay
//   y = headerHeight + rowIndex × rowHeight

import { dayNum, weekday } from './dates'

export type Zoom = 'day' | 'week' | 'month'

export const ROW_HEIGHT = 36
export const HEADER_HEIGHT = 44
export const PIXELS_PER_DAY: Record<Zoom, number> = { day: 28, week: 10, month: 4 }

/** Space reserved after the last date so end-of-bar labels stay on canvas. */
const TRAILING_LABEL_PX = 180
/** Space reserved before the first date for hidden-dependency markers. */
const LEADING_PAD_PX = 72

export interface DateRange {
  start: number // day number, inclusive
  end: number // day number, inclusive
}

export interface Layout {
  zoom: Zoom
  ppd: number
  range: DateRange
  rowCount: number
  width: number
  height: number
  /** Left edge of a date's day cell. */
  x(date: string | number): number
  /** Top edge of a row. */
  y(rowIndex: number): number
  /** Horizontal extent of a bar covering start..finish inclusive. */
  span(start: string, finish: string): { x: number; width: number }
  /** Centre x of a single-day marker such as a milestone. */
  cx(date: string): number
}

/**
 * Build the range from every date that will be drawn (tasks and, when comparing, baseline dates).
 * Starts on the Monday on or before the earliest date minus padding.
 */
export function computeRange(dates: Iterable<string | null>, zoom: Zoom): DateRange | null {
  let lo = Infinity
  let hi = -Infinity
  for (const d of dates) {
    if (!d) continue
    const n = dayNum(d)
    if (n < lo) lo = n
    if (n > hi) hi = n
  }
  if (!Number.isFinite(lo)) return null
  lo -= Math.max(3, Math.ceil(LEADING_PAD_PX / PIXELS_PER_DAY[zoom]))
  lo -= (weekday(lo) + 6) % 7
  hi += Math.ceil(TRAILING_LABEL_PX / PIXELS_PER_DAY[zoom])
  return { start: lo, end: hi }
}

export function createLayout(range: DateRange, zoom: Zoom, rowCount: number): Layout {
  const ppd = PIXELS_PER_DAY[zoom]
  const toDay = (d: string | number) => (typeof d === 'number' ? d : dayNum(d))
  const x = (d: string | number) => (toDay(d) - range.start) * ppd
  return {
    zoom,
    ppd,
    range,
    rowCount,
    width: (range.end - range.start + 1) * ppd,
    height: HEADER_HEIGHT + rowCount * ROW_HEIGHT,
    x,
    y: (i) => HEADER_HEIGHT + i * ROW_HEIGHT,
    span: (s, f) => ({ x: x(s), width: Math.max((toDay(f) - toDay(s) + 1) * ppd, 4) }),
    cx: (d) => x(d) + ppd / 2,
  }
}
