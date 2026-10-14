import type { IncomingMessage, ServerResponse } from 'node:http'
import { StoreError, type Store } from '../core/store'
import {
  emptyPresentation,
  toMeta,
  type Baseline,
  type Emphasis,
  type Presentation,
  type Schedule,
  type Task,
  type TaskPresentation,
} from '../core/model'
import { newId } from '../core/ids'
import { parseFile } from '../core/import'
import { reconcile } from '../core/import/reconcile'
import { ImportError, type ColumnMapping, type ImportPreview, type ParsedSchedule } from '../core/import/types'
import { resolveView, VIEW_IDS, type ViewId, type ViewState } from '../core/views'
import { computeRange, createLayout, HEADER_HEIGHT, ROW_HEIGHT, type Zoom } from '../core/layout'
import type { Criticality } from '../core/model'
import { isSafeId } from './fileStore'
import { openApiDocument } from './openapi'

export type ParseFn = (bytes: Uint8Array, filename: string, mapping?: ColumnMapping) => Promise<ParsedSchedule>

export interface AppOptions {
  parse?: ParseFn
  maxBodyBytes?: number
  previewTtlMs?: number
}

const MAX_BODY = 25 * 1024 * 1024
const PREVIEW_TTL = 30 * 60 * 1000
const SOURCE_FIELDS_MESSAGE = 'Source fields are read-only. Change them in the source file and re-import.'

const SOURCE_FIELD_KEYS = new Set([
  'id', 'name', 'workstream', 'owner', 'start', 'finish', 'milestone', 'critical', 'floatDays', 'float', 'parentId', 'parent',
  'isSummary', 'summary', 'deps', 'predecessors', 'dependencies', 'workstreams',
])

class HttpError extends Error {
  constructor(readonly status: number, message: string) {
    super(message)
  }
}

const bad = (message: string) => new HttpError(400, message)
const notFound = (message: string) => new HttpError(404, message)

type Query = URLSearchParams
interface Ctx {
  req: IncomingMessage
  res: ServerResponse
  params: string[]
  query: Query
}
type Handler = (c: Ctx) => Promise<void>
interface Route {
  method: string
  pattern: RegExp
  handler: Handler
}

function send(res: ServerResponse, status: number, body?: unknown, headers: Record<string, string> = {}): void {
  if (status === 204 || body === undefined) {
    res.writeHead(status, { 'Cache-Control': 'no-store', ...headers })
    res.end()
    return
  }
  const text = JSON.stringify(body)
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': Buffer.byteLength(text),
    'Cache-Control': 'no-store',
    ...headers,
  })
  res.end(text)
}

function isObject(v: unknown): v is Record<string, unknown> {
  return typeof v === 'object' && v !== null && !Array.isArray(v)
}

async function readBody(req: IncomingMessage, limit: number): Promise<Buffer> {
  const declared = Number(req.headers['content-length'])
  if (Number.isFinite(declared) && declared > limit) {
    req.resume()
    throw new HttpError(413, 'That file is too large. Streamline accepts files up to 25 MB.')
  }
  const chunks: Buffer[] = []
  let size = 0
  for await (const chunk of req) {
    size += (chunk as Buffer).length
    if (size > limit) {
      req.resume()
      throw new HttpError(413, 'That file is too large. Streamline accepts files up to 25 MB.')
    }
    chunks.push(chunk as Buffer)
  }
  return Buffer.concat(chunks)
}

async function readJson(req: IncomingMessage, limit: number, optional = false): Promise<Record<string, unknown>> {
  const raw = (await readBody(req, limit)).toString('utf8').trim()
  if (!raw) {
    if (optional) return {}
    throw bad('The request needs a JSON body.')
  }
  let parsed: unknown
  try {
    parsed = JSON.parse(raw)
  } catch {
    throw bad('The request body is not valid JSON.')
  }
  if (!isObject(parsed)) throw bad('The request body must be a JSON object.')
  return parsed
}

function header(req: IncomingMessage, name: string): string | undefined {
  const v = req.headers[name]
  return Array.isArray(v) ? v[0] : v
}

function decode(value: string): string {
  try {
    return decodeURIComponent(value)
  } catch {
    return value
  }
}

function checkId(id: string, what: string): string {
  if (!isSafeId(id)) throw bad(`That ${what} ID is not valid.`)
  return id
}

function cleanName(value: unknown, what: string): string {
  if (typeof value !== 'string' || !value.trim()) throw bad(`Give the ${what} a name.`)
  const name = value.trim()
  if (name.length > 200) throw bad(`The ${what} name is too long. Use 200 characters or fewer.`)
  return name
}

function intParam(q: Query, key: string): number | undefined {
  const v = q.get(key)
  if (v === null || v === '') return undefined
  if (!/^\d+$/.test(v) || Number(v) < 1) throw bad(`"${key}" must be a whole number of 1 or more.`)
  return Number(v)
}

function listParam(q: Query, key: string): string[] | undefined {
  const v = q.get(key)
  if (v === null || v === '') return undefined
  return v.split(',').filter((s) => s !== '')
}

const ISO_DAY = /^\d{4}-\d{2}-\d{2}$/
const EMPHASES: Emphasis[] = ['none', 'highlight', 'muted']
const CRITICALITIES: Criticality[] = ['critical', 'not-critical', 'unavailable']
const ZOOMS: Zoom[] = ['day', 'week', 'month']
const EXPORT_FORMATS = ['png', 'svg', 'pdf', 'csv', 'json']

function parseViewState(q: Query): ViewState {
  const view = (q.get('view') ?? 'detailed') as ViewId
  if (!VIEW_IDS.includes(view)) throw bad(`"view" must be one of: ${VIEW_IDS.join(', ')}.`)
  const state: ViewState = { view }
  const page = intParam(q, 'page')
  const pageSize = intParam(q, 'pageSize')
  if (page !== undefined) state.page = page
  if (pageSize !== undefined) state.pageSize = pageSize
  const collapsed = listParam(q, 'collapsed')
  if (collapsed) state.collapsed = collapsed
  const focus = q.get('focusGroup')
  if (focus !== null) state.focusGroup = focus

  const filters: NonNullable<ViewState['filters']> = {}
  const text = q.get('text')
  if (text) filters.text = text
  const groups = listParam(q, 'groups')
  if (groups) filters.groups = groups
  const crit = listParam(q, 'criticality')
  if (crit) {
    for (const c of crit) if (!CRITICALITIES.includes(c as Criticality)) throw bad(`"criticality" values must be: ${CRITICALITIES.join(', ')}.`)
    filters.criticality = crit as Criticality[]
  }
  const os = q.get('overlapStart')
  const of = q.get('overlapFinish')
  if (os || of) {
    if (!os || !of || !ISO_DAY.test(os) || !ISO_DAY.test(of)) throw bad('"overlapStart" and "overlapFinish" must both be given as YYYY-MM-DD dates.')
    filters.overlap = { start: os, finish: of }
  }
  if (Object.keys(filters).length) state.filters = filters
  return state
}

function summarize(parsed: ParsedSchedule): ImportPreview['summary'] {
  let lo: string | null = null
  let hi: string | null = null
  for (const t of parsed.tasks) {
    for (const d of [t.start, t.finish]) {
      if (!d) continue
      if (lo === null || d < lo) lo = d
      if (hi === null || d > hi) hi = d
    }
  }
  return {
    rowCount: parsed.tasks.length,
    laneCount: parsed.workstreams.length,
    dateRange: lo !== null && hi !== null ? { start: lo, finish: hi } : null,
    hasCriticality: parsed.tasks.some((t) => t.critical !== 'unavailable'),
    hasFloat: parsed.tasks.some((t) => t.floatDays !== null),
  }
}

function csvCell(v: string | number | boolean | null): string {
  if (v === null) return ''
  const s = String(v)
  return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s
}

const CSV_COLUMNS = ['id', 'name', 'workstream', 'owner', 'start', 'finish', 'milestone', 'critical', 'float_days', 'parent_id', 'is_summary', 'predecessors']

/** Empty cell means the source file did not provide the value. Predecessors are `id:TYPE:lagDays` joined by `;`. */
function toCsv(tasks: Task[]): string {
  const lines = [CSV_COLUMNS.join(',')]
  for (const t of tasks) {
    const preds = t.deps.map((d) => `${d.predecessorId}:${d.type}:${d.lagDays ?? ''}`).join(';')
    lines.push(
      [t.id, t.name, t.workstream, t.owner, t.start, t.finish, t.milestone, t.critical, t.floatDays, t.parentId, t.isSummary, preds || null]
        .map(csvCell)
        .join(','),
    )
  }
  return lines.join('\r\n') + '\r\n'
}

function fileSafe(name: string): string {
  const s = name.replace(/[^A-Za-z0-9._-]+/g, '_').replace(/^_+|_+$/g, '')
  return s || 'schedule'
}

function applyTaskPatch(current: TaskPresentation | undefined, patch: unknown, taskId: string): TaskPresentation | null {
  if (!isObject(patch)) throw bad(`The override for task "${taskId}" must be an object or null.`)
  const next: Record<string, unknown> = { ...(current ?? {}) }
  for (const [key, value] of Object.entries(patch)) {
    if (SOURCE_FIELD_KEYS.has(key)) throw bad(SOURCE_FIELDS_MESSAGE)
    const clear = value === null || value === ''
    switch (key) {
      case 'displayLabel':
      case 'displayGroup':
        if (clear) delete next[key]
        else if (typeof value === 'string') next[key] = value
        else throw bad(`"${key}" must be text.`)
        break
      case 'order':
        if (value === null) delete next[key]
        else if (typeof value === 'number' && Number.isFinite(value)) next[key] = value
        else throw bad('"order" must be a number.')
        break
      case 'emphasis':
        if (clear) delete next[key]
        else if (EMPHASES.includes(value as Emphasis)) next[key] = value
        else throw bad(`"emphasis" must be one of: ${EMPHASES.join(', ')}.`)
        break
      default:
        throw bad(`"${key}" is not a presentation field.`)
    }
  }
  return Object.keys(next).length ? (next as TaskPresentation) : null
}

export function createApp(store: Store, opts: AppOptions = {}): (req: IncomingMessage, res: ServerResponse) => void {
  const parse = opts.parse ?? parseFile
  const maxBody = opts.maxBodyBytes ?? MAX_BODY
  const ttl = opts.previewTtlMs ?? PREVIEW_TTL
  const previews = new Map<string, { preview: ImportPreview; replaceId: string | null; expires: number }>()

  const sweep = () => {
    const now = Date.now()
    for (const [id, p] of previews) if (p.expires <= now) previews.delete(id)
  }

  async function loadSchedule(id: string): Promise<Schedule> {
    const s = await store.getSchedule(checkId(id, 'schedule'))
    if (!s) throw notFound('That schedule was not found. It may have been deleted.')
    return s
  }

  const routes: Route[] = []
  const route = (method: string, path: string, handler: Handler) => {
    routes.push({ method, pattern: new RegExp(`^${path.replace(/\{[^}]+\}/g, '([^/]+)')}$`), handler })
  }

  route('GET', '/api/openapi.json', async ({ res }) => send(res, 200, openApiDocument))

  route('GET', '/api/schedules', async ({ res }) => send(res, 200, await store.listSchedules()))

  route('POST', '/api/schedules/import/preview', async ({ req, res }) => {
    const rawName = header(req, 'x-filename')
    if (!rawName || !decode(rawName).trim()) throw bad('Add the file name in the X-Filename header so Streamline knows how to read the file.')
    const filename = decode(rawName).trim()

    let mapping: ColumnMapping | undefined
    const rawMapping = header(req, 'x-mapping')
    if (rawMapping) {
      let m: unknown
      try {
        m = JSON.parse(decode(rawMapping))
      } catch {
        throw bad('The column mapping is not valid JSON.')
      }
      if (!isObject(m) || !Object.values(m).every((v) => typeof v === 'string')) throw bad('The column mapping must map each field to a column name.')
      mapping = m as ColumnMapping
    }

    const replaceRaw = header(req, 'x-replace-schedule-id')
    const replaceId = replaceRaw ? checkId(replaceRaw, 'schedule') : null
    const existing = replaceId ? await loadSchedule(replaceId) : null

    const bytes = await readBody(req, maxBody)
    if (bytes.length === 0) throw bad('The file is empty.')

    let parsed: ParsedSchedule
    try {
      parsed = await parse(new Uint8Array(bytes), filename, mapping)
    } catch (e) {
      if (e instanceof ImportError) throw new HttpError(422, e.message)
      throw new HttpError(500, 'Streamline could not read this file. Check that it is a valid schedule export and try again.')
    }

    const preview: ImportPreview = { previewId: newId('imp'), parsed, summary: summarize(parsed) }
    if (existing) {
      preview.reconciliation = reconcile(existing, parsed, await store.getPresentation(existing.id))
    }
    sweep()
    previews.set(preview.previewId, { preview, replaceId, expires: Date.now() + ttl })
    send(res, 200, preview)
  })

  route('POST', '/api/schedules/import/accept', async ({ req, res }) => {
    const body = await readJson(req, maxBody)
    if (typeof body.previewId !== 'string') throw bad('Include the previewId from the import preview.')
    sweep()
    const entry = previews.get(body.previewId)
    if (!entry) throw new HttpError(410, 'This import preview has expired. Choose the file again to continue.')
    const name = body.name === undefined || body.name === null ? undefined : cleanName(body.name, 'schedule')
    const { parsed } = entry.preview
    const { replaceId } = entry
    const now = new Date().toISOString()
    const source = { kind: parsed.kind, filename: parsed.filename, importedAt: now }

    let schedule: Schedule
    if (replaceId) {
      const existing = await loadSchedule(replaceId)
      schedule = { ...existing, name: name ?? existing.name, source, workstreams: parsed.workstreams, tasks: parsed.tasks }
    } else {
      schedule = { id: newId('sch'), name: name ?? parsed.name, createdAt: now, source, workstreams: parsed.workstreams, tasks: parsed.tasks }
    }
    await store.putSchedule(schedule)
    previews.delete(body.previewId)
    send(res, replaceId ? 200 : 201, schedule)
  })

  route('GET', '/api/schedules/{id}', async ({ res, params }) => send(res, 200, await loadSchedule(params[0])))

  // Rename changes only the schedule record's name, which is a user label. Source tasks and dates are untouched.
  route('PATCH', '/api/schedules/{id}', async ({ req, res, params }) => {
    const s = await loadSchedule(params[0])
    const body = await readJson(req, maxBody)
    const updated = { ...s, name: cleanName(body.name, 'schedule') }
    await store.putSchedule(updated)
    send(res, 200, toMeta(updated))
  })

  route('DELETE', '/api/schedules/{id}', async ({ res, params }) => {
    await loadSchedule(params[0])
    await store.deleteSchedule(params[0])
    send(res, 204)
  })

  route('GET', '/api/schedules/{id}/view', async ({ res, params, query }) => {
    const schedule = await loadSchedule(params[0])
    const state = parseViewState(query)
    const zoomRaw = query.get('zoom') ?? 'week'
    if (!ZOOMS.includes(zoomRaw as Zoom)) throw bad(`"zoom" must be one of: ${ZOOMS.join(', ')}.`)
    const zoom = zoomRaw as Zoom

    const compare = query.get('compare')
    let baseline: Baseline | undefined
    if (compare) {
      checkId(compare, 'baseline')
      baseline = (await store.listBaselines(schedule.id)).find((b) => b.id === compare)
      if (!baseline) throw notFound('That baseline was not found.')
    }

    const presentation = (await store.getPresentation(schedule.id)) ?? emptyPresentation(schedule.id)
    let resolved
    try {
      resolved = resolveView(schedule, presentation, state)
    } catch {
      throw new HttpError(500, 'Streamline could not build this view.')
    }

    const byId = new Map(schedule.tasks.map((t) => [t.id, t]))
    const baseById = new Map((baseline?.tasks ?? []).map((t) => [t.id, t]))
    const dates: Array<string | null> = []
    const addTask = (id: string) => {
      const t = byId.get(id)
      if (t) dates.push(t.start, t.finish)
      const b = baseById.get(id)
      if (b) dates.push(b.start, b.finish)
    }
    for (const row of resolved.rows) {
      if (row.kind === 'task') addTask(row.taskId)
      else if (row.kind === 'milestones') row.taskIds.forEach(addTask)
      else dates.push(row.start, row.finish)
    }
    const range = computeRange(dates, zoom)
    const layout = range
      ? (({ ppd, width, height }) => ({ zoom, ppd, range, width, height }))(createLayout(range, zoom, resolved.rows.length))
      : { zoom, ppd: createLayout({ start: 0, end: 0 }, zoom, 0).ppd, range: null, width: 0, height: HEADER_HEIGHT + resolved.rows.length * ROW_HEIGHT }
    send(res, 200, { resolved, layout })
  })

  route('GET', '/api/schedules/{id}/presentation', async ({ res, params }) => {
    const s = await loadSchedule(params[0])
    send(res, 200, (await store.getPresentation(s.id)) ?? emptyPresentation(s.id))
  })

  route('PATCH', '/api/schedules/{id}/presentation', async ({ req, res, params }) => {
    const schedule = await loadSchedule(params[0])
    const body = await readJson(req, maxBody)
    const current = (await store.getPresentation(schedule.id)) ?? emptyPresentation(schedule.id)
    const next: Presentation = { ...current, tasks: { ...current.tasks } }

    for (const key of Object.keys(body)) {
      if (['displayName', 'groupOrder', 'theme', 'tasks'].includes(key)) continue
      throw bad(SOURCE_FIELD_KEYS.has(key) ? SOURCE_FIELDS_MESSAGE : `"${key}" is not a presentation field.`)
    }

    for (const key of ['displayName', 'theme'] as const) {
      if (!(key in body)) continue
      const v = body[key]
      if (v === null || v === '') delete next[key]
      else if (typeof v === 'string') next[key] = v
      else throw bad(`"${key}" must be text.`)
    }
    if ('groupOrder' in body) {
      const v = body.groupOrder
      if (v === null) delete next.groupOrder
      else if (Array.isArray(v) && v.every((x) => typeof x === 'string')) next.groupOrder = v as string[]
      else throw bad('"groupOrder" must be a list of group names.')
    }
    if ('tasks' in body) {
      if (!isObject(body.tasks)) throw bad('"tasks" must map task IDs to overrides.')
      const known = new Set(schedule.tasks.map((t) => t.id))
      for (const [taskId, patch] of Object.entries(body.tasks)) {
        if (patch === null) {
          delete next.tasks[taskId]
          continue
        }
        if (!known.has(taskId)) throw bad(`Task "${taskId}" is not in this schedule, so it cannot have a presentation override.`)
        const merged = applyTaskPatch(next.tasks[taskId], patch, taskId)
        if (merged) next.tasks[taskId] = merged
        else delete next.tasks[taskId]
      }
    }
    next.updatedAt = new Date().toISOString()
    await store.putPresentation(schedule.id, next)
    send(res, 200, next)
  })

  route('POST', '/api/schedules/{id}/presentation/reset', async ({ req, res, params }) => {
    const schedule = await loadSchedule(params[0])
    const body = await readJson(req, maxBody, true)
    const current = (await store.getPresentation(schedule.id)) ?? emptyPresentation(schedule.id)
    let tasks: Presentation['tasks'] = {}
    if (body.taskIds !== undefined) {
      if (!Array.isArray(body.taskIds) || !body.taskIds.every((x) => typeof x === 'string')) throw bad('"taskIds" must be a list of task IDs.')
      const drop = new Set(body.taskIds as string[])
      tasks = Object.fromEntries(Object.entries(current.tasks).filter(([id]) => !drop.has(id)))
    }
    const next: Presentation = { ...current, tasks, updatedAt: new Date().toISOString() }
    await store.putPresentation(schedule.id, next)
    send(res, 200, next)
  })

  route('GET', '/api/schedules/{id}/baselines', async ({ res, params }) => {
    const s = await loadSchedule(params[0])
    send(res, 200, await store.listBaselines(s.id))
  })

  route('POST', '/api/schedules/{id}/baselines', async ({ req, res, params }) => {
    const schedule = await loadSchedule(params[0])
    const body = await readJson(req, maxBody)
    const baseline: Baseline = {
      id: newId('bl'),
      scheduleId: schedule.id,
      name: cleanName(body.name, 'baseline'),
      createdAt: new Date().toISOString(),
      tasks: schedule.tasks.map((t) => ({ id: t.id, name: t.name, start: t.start, finish: t.finish, milestone: t.milestone })),
    }
    await store.putBaseline(schedule.id, baseline)
    send(res, 201, baseline)
  })

  route('DELETE', '/api/schedules/{id}/baselines/{baselineId}', async ({ res, params }) => {
    const schedule = await loadSchedule(params[0])
    const baselineId = checkId(params[1], 'baseline')
    if (!(await store.listBaselines(schedule.id)).some((b) => b.id === baselineId)) throw notFound('That baseline was not found.')
    await store.deleteBaseline(schedule.id, baselineId)
    send(res, 204)
  })

  route('POST', '/api/schedules/{id}/export', async ({ req, res, params }) => {
    const schedule = await loadSchedule(params[0])
    const body = await readJson(req, maxBody)
    const format = body.format
    if (typeof format !== 'string' || !EXPORT_FORMATS.includes(format)) throw bad(`"format" must be one of: ${EXPORT_FORMATS.join(', ')}.`)
    if (format === 'png' || format === 'svg' || format === 'pdf') throw new HttpError(501, 'Not available yet')
    const base = fileSafe(schedule.name)
    const text = format === 'csv' ? toCsv(schedule.tasks) : JSON.stringify(schedule, null, 2)
    const filename = `${base}.${format}`
    res.writeHead(200, {
      'Content-Type': format === 'csv' ? 'text/csv; charset=utf-8' : 'application/json; charset=utf-8',
      'Content-Disposition': `attachment; filename="${filename}"`,
      'Content-Length': Buffer.byteLength(text),
      'Cache-Control': 'no-store',
    })
    res.end(text)
  })

  async function handle(req: IncomingMessage, res: ServerResponse): Promise<void> {
    const url = new URL(req.url ?? '/', 'http://localhost')
    const method = req.method ?? 'GET'
    let path = url.pathname
    if (path.length > 1 && path.endsWith('/')) path = path.slice(0, -1)

    let allowed = false
    for (const r of routes) {
      const m = r.pattern.exec(path)
      if (!m) continue
      allowed = true
      if (r.method !== method) continue
      let params: string[]
      try {
        params = m.slice(1).map((p) => decodeURIComponent(p))
      } catch {
        throw bad('That address contains invalid characters.')
      }
      await r.handler({ req, res, params, query: url.searchParams })
      return
    }
    if (allowed) throw new HttpError(405, 'That method is not supported for this address.')
    throw notFound('There is nothing at this address.')
  }

  return (req, res) => {
    handle(req, res).catch((e: unknown) => {
      if (res.headersSent) {
        res.destroy()
        return
      }
      if (e instanceof HttpError) return send(res, e.status, { error: { message: e.message } })
      if (e instanceof StoreError) return send(res, 500, { error: { message: e.message } })
      send(res, 500, { error: { message: 'Something went wrong on the server. Try again.' } })
    })
  }
}
