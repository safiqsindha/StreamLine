// The web app's only data access. Every call goes to the local API, the same
// endpoints agents use; the browser never touches storage directly.

import type { ImportPreview, ColumnMapping } from '../core/import/types'
import type { DateRange, Zoom } from '../core/layout'
import type { Baseline, Presentation, Schedule, ScheduleMeta, TaskPresentation } from '../core/model'
import type { ResolvedView, ViewState } from '../core/views'

export class ApiError extends Error {
  constructor(
    message: string,
    readonly status: number,
  ) {
    super(message)
    this.name = 'ApiError'
  }
}

export interface ViewResponse {
  resolved: ResolvedView
  layout: { zoom: Zoom; ppd: number; range: DateRange | null; width: number; height: number }
}

export type ExportFormat = 'png' | 'svg' | 'pdf' | 'csv' | 'json'

const enc = encodeURIComponent

async function request<T>(method: string, path: string, init: { json?: unknown; body?: BodyInit; headers?: Record<string, string> } = {}): Promise<T> {
  const headers: Record<string, string> = { ...init.headers }
  let body = init.body
  if (init.json !== undefined) {
    headers['Content-Type'] = 'application/json'
    body = JSON.stringify(init.json)
  }
  let res: Response
  try {
    res = await fetch(path, { method, headers, body })
  } catch {
    throw new ApiError('Streamline could not reach its local service. Start it with "npm start" and reload.', 0)
  }
  if (res.status === 204) return undefined as T
  const text = await res.text()
  let data: unknown = null
  try {
    data = text ? JSON.parse(text) : null
  } catch {
    // Non-JSON error bodies fall through to the generic message below.
  }
  if (!res.ok) {
    const message = (data as { error?: { message?: string } } | null)?.error?.message ?? `Request failed (${res.status}).`
    throw new ApiError(message, res.status)
  }
  return data as T
}

function viewQuery(state: ViewState, zoom: Zoom, compare?: string | null): string {
  const q = new URLSearchParams({ view: state.view, zoom })
  if (state.page) q.set('page', String(state.page))
  if (state.pageSize) q.set('pageSize', String(state.pageSize))
  if (state.collapsed?.length) q.set('collapsed', state.collapsed.join(','))
  if (state.focusGroup !== undefined) q.set('focusGroup', state.focusGroup)
  const f = state.filters
  if (f?.text) q.set('text', f.text)
  if (f?.groups?.length) q.set('groups', f.groups.join(','))
  if (f?.criticality?.length) q.set('criticality', f.criticality.join(','))
  if (f?.overlap) {
    q.set('overlapStart', f.overlap.start)
    q.set('overlapFinish', f.overlap.finish)
  }
  if (compare) q.set('compare', compare)
  return q.toString()
}

export const api = {
  listSchedules: () => request<ScheduleMeta[]>('GET', '/api/schedules'),
  getSchedule: (id: string) => request<Schedule>('GET', `/api/schedules/${enc(id)}`),
  renameSchedule: (id: string, name: string) => request<ScheduleMeta>('PATCH', `/api/schedules/${enc(id)}`, { json: { name } }),
  deleteSchedule: (id: string) => request<void>('DELETE', `/api/schedules/${enc(id)}`),

  previewImport: (file: File, opts: { mapping?: ColumnMapping; replaceId?: string } = {}) => {
    const headers: Record<string, string> = { 'X-Filename': enc(file.name), 'Content-Type': 'application/octet-stream' }
    if (opts.mapping) headers['X-Mapping'] = enc(JSON.stringify(opts.mapping))
    if (opts.replaceId) headers['X-Replace-Schedule-Id'] = opts.replaceId
    return request<ImportPreview>('POST', '/api/schedules/import/preview', { body: file, headers })
  },
  acceptImport: (previewId: string, name?: string) => request<Schedule>('POST', '/api/schedules/import/accept', { json: { previewId, name } }),

  getView: (id: string, state: ViewState, zoom: Zoom, compare?: string | null) =>
    request<ViewResponse>('GET', `/api/schedules/${enc(id)}/view?${viewQuery(state, zoom, compare)}`),

  getPresentation: (id: string) => request<Presentation>('GET', `/api/schedules/${enc(id)}/presentation`),
  patchPresentation: (
    id: string,
    patch: { groupOrder?: string[] | null; tasks?: Record<string, Partial<Record<keyof TaskPresentation, unknown>> | null> },
  ) => request<Presentation>('PATCH', `/api/schedules/${enc(id)}/presentation`, { json: patch }),
  resetPresentation: (id: string, taskIds?: string[]) =>
    request<Presentation>('POST', `/api/schedules/${enc(id)}/presentation/reset`, { json: taskIds ? { taskIds } : {} }),

  listBaselines: (id: string) => request<Baseline[]>('GET', `/api/schedules/${enc(id)}/baselines`),
  saveBaseline: (id: string, name: string) => request<Baseline>('POST', `/api/schedules/${enc(id)}/baselines`, { json: { name } }),
  deleteBaseline: (id: string, baselineId: string) => request<void>('DELETE', `/api/schedules/${enc(id)}/baselines/${enc(baselineId)}`),

  /** Returns the exported file. The server renders it from the same layout as the screen. */
  exportFile: async (id: string, body: Record<string, unknown>): Promise<{ blob: Blob; filename: string }> => {
    let res: Response
    try {
      res = await fetch(`/api/schedules/${enc(id)}/export`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) })
    } catch {
      throw new ApiError('Streamline could not reach its local service.', 0)
    }
    if (!res.ok) {
      const data = await res.json().catch(() => null)
      throw new ApiError(data?.error?.message ?? `Export failed (${res.status}).`, res.status)
    }
    const cd = res.headers.get('Content-Disposition') ?? ''
    const filename = /filename="([^"]+)"/.exec(cd)?.[1] ?? `export.${body.format}`
    return { blob: await res.blob(), filename }
  },

  openApi: () => request<OpenApiDoc>('GET', '/api/openapi.json'),
}

export interface OpenApiOperation {
  summary?: string
  operationId?: string
  'x-example'?: { request?: unknown; response?: unknown }
}
export interface OpenApiDoc {
  paths: Record<string, Record<string, OpenApiOperation>>
}
