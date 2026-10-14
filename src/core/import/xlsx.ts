import ExcelJS from 'exceljs'
import { fallbackTaskId } from '../ids'
import { isIsoDate, isoOf } from '../dates'
import type { Criticality, Dependency, DependencyType, Task, Workstream } from '../model'
import { ImportError, type ColumnMapping, type ImportIssue, type MappedField, type ParsedSchedule } from './types'

const SYNONYMS: Record<MappedField, string[]> = {
  id: ['id', 'taskid', 'wbsid', 'uniqueid', 'wbs', 'activityid'],
  name: ['task', 'taskname', 'name', 'activity', 'activityname'],
  workstream: ['workstream', 'lane', 'swimlane', 'group', 'team', 'phase'],
  owner: ['owner', 'resource', 'assignedto', 'responsible'],
  start: ['start', 'startdate', 'begin'],
  finish: ['finish', 'end', 'finishdate', 'enddate', 'due'],
  milestone: ['milestone', 'type'],
  critical: ['critical', 'iscritical', 'criticalpath'],
  float: ['float', 'totalfloat', 'totalslack', 'slack'],
  parent: ['parent', 'parentid'],
  summary: ['summary'],
  predecessors: ['predecessors', 'dependson', 'dependencies'],
}

const FIELDS = Object.keys(SYNONYMS) as MappedField[]
const HEADER_SCAN_ROWS = 10

const MSG_NO_DATES =
  'This file has no recognizable date columns. Streamline needs a start and finish column for each task.'
const MSG_NO_NAME =
  'This file has no recognizable task name column. Streamline needs a column such as "Task" or "Task Name" that names each task.'
const MSG_UNREADABLE =
  'Streamline could not read this file. Open it in Excel, choose File > Save As, save it as an Excel Workbook (.xlsx), and try again.'
const MSG_NO_ROWS =
  'No usable tasks were found in this file. Check that each task row has a name and valid dates, then try again.'

const norm = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, '')

function plain(v: ExcelJS.CellValue | undefined): unknown {
  if (v === null || v === undefined) return null
  if (v instanceof Date || typeof v !== 'object') return v
  const o = v as unknown as Record<string, unknown>
  if ('result' in o) return plain(o.result as ExcelJS.CellValue)
  if ('richText' in o && Array.isArray(o.richText)) return (o.richText as Array<{ text: string }>).map((t) => t.text).join('')
  if ('text' in o) return plain(o.text as ExcelJS.CellValue)
  return null
}

function textOf(v: unknown): string {
  if (v === null || v === undefined) return ''
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? '' : v.toISOString().slice(0, 10)
  return String(v).trim()
}

const MONTHS: Record<string, number> = {
  jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12,
}

function build(y: number, m: number, d: number): string | null {
  const iso = `${String(y).padStart(4, '0')}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`
  return isIsoDate(iso) ? iso : null
}

function parseDateText(s: string): string | null {
  let m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[T\s].*)?$/.exec(s)
  if (m) return build(+m[1], +m[2], +m[3])
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+\d.*)?$/.exec(s)
  if (m) {
    let mo = +m[1]
    let d = +m[2]
    if (mo > 12 && d <= 12) [mo, d] = [d, mo]
    return build(+m[3], mo, d)
  }
  m = /^(\d{1,2})[-\s]+([A-Za-z]{3,9})\.?[-,\s]+(\d{4})$/.exec(s)
  if (m) {
    const mo = MONTHS[m[2].slice(0, 3).toLowerCase()]
    return mo ? build(+m[3], mo, +m[1]) : null
  }
  m = /^([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s+(\d{4})$/.exec(s)
  if (m) {
    const mo = MONTHS[m[1].slice(0, 3).toLowerCase()]
    return mo ? build(+m[3], mo, +m[2]) : null
  }
  return null
}

/** A slash date such as 03/04/2026 that reads differently as month/day and day/month. */
function isAmbiguousSlashDate(v: unknown): boolean {
  const m = typeof v === 'string' ? /^(\d{1,2})\/(\d{1,2})\/\d{4}/.exec(v.trim()) : null
  return !!m && +m[1] <= 12 && +m[2] <= 12 && m[1] !== m[2]
}

/** Returns an ISO date, null for empty, or undefined when the value is present but unreadable. */
function parseDate(v: unknown): string | null | undefined {
  if (v === null || v === undefined) return null
  if (v instanceof Date) {
    if (Number.isNaN(v.getTime())) return undefined
    const iso = isoOf(Math.floor(v.getTime() / 86_400_000))
    return isIsoDate(iso) ? iso : undefined
  }
  if (typeof v === 'number') {
    // Excel serial day number in a cell not formatted as a date.
    if (v < 20_000 || v > 80_000) return undefined
    return isoOf(Math.floor(Date.UTC(1899, 11, 30) / 86_400_000) + Math.floor(v))
  }
  const s = String(v).trim()
  if (!s) return null
  return parseDateText(s) ?? undefined
}

function parseBool(s: string): boolean | null {
  const t = s.trim().toLowerCase()
  if (['yes', 'true', '1', 'y'].includes(t)) return true
  if (['no', 'false', '0', 'n'].includes(t)) return false
  return null
}

function parseCritical(s: string): Criticality {
  const t = s.trim().toLowerCase()
  if (t === 'critical') return 'critical'
  const b = parseBool(t)
  if (b === true) return 'critical'
  if (b === false) return 'not-critical'
  return 'unavailable'
}

function parseFloatDays(v: unknown): number | null | undefined {
  if (v === null || v === undefined) return null
  if (typeof v === 'number') return Number.isFinite(v) ? v : undefined
  const s = String(v).trim()
  if (!s) return null
  const m = /^(-?\d+(?:\.\d+)?)\s*(?:d|days?)?$/i.exec(s)
  return m ? Number(m[1]) : undefined
}

interface PredToken {
  ref: string
  type: DependencyType
  lagDays: number | null
}

function parsePredecessors(text: string, knownIds: Set<string>): PredToken[] {
  const out: PredToken[] = []
  for (const raw of text.split(/[,;]/)) {
    const tok = raw.trim()
    if (!tok) continue
    if (knownIds.has(tok)) {
      out.push({ ref: tok, type: 'unknown', lagDays: null })
      continue
    }
    const m = /^(.+?)\s*(FS|SS|FF|SF)?\s*(?:([+-])\s*(\d+(?:\.\d+)?)\s*(?:d|days?)?)?$/i.exec(tok)
    if (!m) {
      out.push({ ref: tok, type: 'unknown', lagDays: null })
      continue
    }
    const type = (m[2]?.toUpperCase() as DependencyType | undefined) ?? 'unknown'
    const lagDays = m[4] === undefined ? null : Number(m[4]) * (m[3] === '-' ? -1 : 1)
    out.push({ ref: m[1].trim(), type, lagDays })
  }
  return out
}

function autoMap(headers: string[]): ColumnMapping {
  const mapping: ColumnMapping = {}
  const used = new Set<string>()
  for (const field of FIELDS) {
    const hit = headers.find((h) => !used.has(h) && SYNONYMS[field].includes(norm(h)))
    if (hit !== undefined) {
      mapping[field] = hit
      used.add(hit)
    }
  }
  return mapping
}

function resolveMapping(given: ColumnMapping, headers: string[]): ColumnMapping {
  const out: ColumnMapping = {}
  for (const field of FIELDS) {
    const want = given[field]
    if (!want) continue
    const hit = headers.find((h) => h === want) ?? headers.find((h) => h.trim() === want.trim())
    if (hit !== undefined) out[field] = hit
  }
  return out
}

function rowHeaders(row: ExcelJS.Row): Array<{ col: number; text: string }> {
  const out: Array<{ col: number; text: string }> = []
  row.eachCell({ includeEmpty: false }, (cell, col) => {
    const t = textOf(plain(cell.value))
    if (t) out.push({ col, text: t })
  })
  return out
}

interface HeaderHit {
  sheet: ExcelJS.Worksheet
  rowNumber: number
  headers: Array<{ col: number; text: string }>
  score: number
}

function findHeader(wb: ExcelJS.Workbook, given?: ColumnMapping): HeaderHit | null {
  let best: HeaderHit | null = null
  for (const sheet of wb.worksheets) {
    const last = Math.min(sheet.rowCount, HEADER_SCAN_ROWS)
    for (let r = 1; r <= last; r++) {
      const headers = rowHeaders(sheet.getRow(r))
      const names = headers.map((h) => h.text)
      const mapped = given ? resolveMapping(given, names) : autoMap(names)
      const score = Object.keys(mapped).length
      if (score > 0 && (!best || score > best.score)) best = { sheet, rowNumber: r, headers, score }
    }
  }
  return best
}

export async function parseXlsx(bytes: Uint8Array, filename: string, mapping?: ColumnMapping): Promise<ParsedSchedule> {
  const wb = new ExcelJS.Workbook()
  try {
    await wb.xlsx.load(bytes as unknown as Parameters<typeof wb.xlsx.load>[0])
  } catch {
    throw new ImportError(MSG_UNREADABLE)
  }

  const hit = findHeader(wb, mapping)
  if (!hit) throw new ImportError(MSG_NO_DATES)

  const detectedColumns = hit.headers.map((h) => h.text)
  const used = mapping ? resolveMapping(mapping, detectedColumns) : autoMap(detectedColumns)
  if (!used.start && !used.finish) throw new ImportError(MSG_NO_DATES)
  if (!used.name) throw new ImportError(MSG_NO_NAME)

  const colOf = (field: MappedField): number | null => {
    const h = used[field]
    return h === undefined ? null : hit.headers.find((x) => x.text === h)!.col
  }
  const cols = Object.fromEntries(FIELDS.map((f) => [f, colOf(f)])) as Record<MappedField, number | null>
  const mappedHeaders = new Set(Object.values(used))
  const unmapped = detectedColumns.filter((h) => !mappedHeaders.has(h))

  const issues: ImportIssue[] = []
  const skip = (row: number, message: string) => issues.push({ row, message, severity: 'skipped' })
  const warn = (row: number, message: string) => issues.push({ row, message, severity: 'warning' })

  const get = (row: ExcelJS.Row, field: MappedField): unknown => {
    const c = cols[field]
    return c === null ? null : plain(row.getCell(c).value)
  }
  const label = (field: MappedField) => used[field] ?? field

  interface Pending {
    task: Task
    row: number
    predText: string
  }
  const pending: Pending[] = []
  const seenIds = new Set<string>()
  let ambiguousDate: string | null = null
  const occurrences = new Map<string, number>()

  for (let r = hit.rowNumber + 1; r <= hit.sheet.rowCount; r++) {
    const row = hit.sheet.getRow(r)
    if (rowHeaders(row).length === 0) continue

    const name = textOf(get(row, 'name'))
    if (!name) {
      skip(r, `Row ${r} has no task name. That row was skipped.`)
      continue
    }
    const workstream = textOf(get(row, 'workstream')) || null

    // Counted before any validity skip so a task's fallback ID does not depend on its neighbours.
    let fallback = ''
    if (cols.id === null) {
      const key = `${workstream ?? ''}␟${name}`.toLowerCase()
      const n = occurrences.get(key) ?? 0
      occurrences.set(key, n + 1)
      fallback = fallbackTaskId([workstream, name], n)
    }

    const dates: Record<'start' | 'finish', string | null> = { start: null, finish: null }
    for (const f of ['start', 'finish'] as const) {
      const raw = get(row, f)
      if (ambiguousDate === null && isAmbiguousSlashDate(raw)) ambiguousDate = textOf(raw)
      const d = parseDate(raw)
      if (d === undefined) {
        warn(
          r,
          `Row ${r}: the "${label(f)}" column has a value Streamline could not read as a date ("${textOf(raw)}"). It is shown as Unavailable.`,
        )
        dates[f] = null
      } else {
        dates[f] = d
      }
    }
    if (dates.start && dates.finish && dates.finish < dates.start) {
      skip(r, `Row ${r} has a finish date before its start date. That row was skipped.`)
      continue
    }

    let id = fallback
    if (cols.id !== null) {
      id = textOf(get(row, 'id'))
      if (!id) {
        skip(r, `Row ${r} has no value in the "${label('id')}" column. That row was skipped.`)
        continue
      }
      if (seenIds.has(id)) {
        skip(r, `Row ${r} repeats the ID "${id}", which an earlier row already uses. That row was skipped.`)
        continue
      }
    }
    seenIds.add(id)

    let floatDays: number | null = null
    if (cols.float !== null) {
      const raw = get(row, 'float')
      const f = parseFloatDays(raw)
      if (f === undefined) warn(r, `Row ${r}: the "${label('float')}" column has a value Streamline could not read as a number of days ("${textOf(raw)}"). Float was left empty.`)
      else floatDays = f
    }

    let milestone = false
    if (cols.milestone !== null) {
      const t = textOf(get(row, 'milestone')).toLowerCase()
      milestone = t === 'milestone' || parseBool(t) === true
    }

    pending.push({
      row: r,
      predText: textOf(get(row, 'predecessors')),
      task: {
        id,
        name,
        workstream,
        owner: textOf(get(row, 'owner')) || null,
        start: dates.start,
        finish: dates.finish,
        milestone,
        critical: cols.critical === null ? 'unavailable' : parseCritical(textOf(get(row, 'critical'))),
        floatDays,
        parentId: textOf(get(row, 'parent')) || null,
        isSummary: cols.summary === null ? false : parseBool(textOf(get(row, 'summary'))) === true,
        deps: [],
      },
    })
  }

  if (pending.length === 0) throw new ImportError(MSG_NO_ROWS)

  for (const p of pending) {
    if (!p.predText) continue
    if (cols.id === null) {
      warn(
        p.row,
        `Row ${p.row} lists predecessors, but the file has no ID column, so Streamline cannot tell which tasks they refer to. The dependencies on that row were dropped.`,
      )
      continue
    }
    const deps: Dependency[] = []
    for (const t of parsePredecessors(p.predText, seenIds)) {
      if (t.ref === p.task.id) {
        warn(p.row, `Row ${p.row} lists itself as a predecessor. That dependency was dropped.`)
      } else if (!seenIds.has(t.ref)) {
        warn(
          p.row,
          `Row ${p.row} lists predecessor "${t.ref}", but no imported task has that ID. That dependency was dropped.`,
        )
      } else if (!deps.some((d) => d.predecessorId === t.ref)) {
        deps.push({ predecessorId: t.ref, type: t.type, lagDays: t.lagDays })
      }
    }
    p.task.deps = deps
  }

  const tasks = pending.map((p) => p.task)
  const seenWs = new Set<string>()
  const workstreams: Workstream[] = []
  for (const t of tasks) {
    if (t.workstream && !seenWs.has(t.workstream)) {
      seenWs.add(t.workstream)
      workstreams.push({ id: t.workstream, name: t.workstream })
    }
  }

  if (ambiguousDate !== null) {
    issues.push({
      message: `Some dates, such as "${ambiguousDate}", can be read as month/day or day/month. Streamline read them as month/day. If that is wrong, store the dates as Excel dates or as YYYY-MM-DD and import again.`,
      severity: 'warning',
    })
  }

  return {
    kind: 'xlsx',
    filename,
    name: filename.replace(/\.[^./\\]+$/, ''),
    workstreams,
    tasks,
    detectedColumns,
    mapping: used,
    unmapped,
    issues,
  }
}
