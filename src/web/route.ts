// URL hash carries page, schedule and view so people and agents can deep-link:
//   #page=schedules&schedule=sch_…&view=critical

import { VIEW_IDS, type ViewId } from '../core/views'

export type PageId = 'schedules' | 'views' | 'baselines' | 'exports' | 'api'
export const PAGE_IDS: PageId[] = ['schedules', 'views', 'baselines', 'exports', 'api']

export interface Route {
  page: PageId
  /** Open schedule. Absent on the schedules page means the landing list. */
  schedule: string | null
  view: ViewId
  /** Workstreams drilldown target (exact display group key). */
  group?: string
}

export const DEFAULT_ROUTE: Route = { page: 'schedules', schedule: null, view: 'detailed' }

export function readRoute(hash: string): Route {
  const route: Route = { ...DEFAULT_ROUTE }
  try {
    const h = new URLSearchParams(hash.replace(/^#/, ''))
    const page = h.get('page') as PageId | null
    if (page && PAGE_IDS.includes(page)) route.page = page
    const view = h.get('view') as ViewId | null
    if (view && VIEW_IDS.includes(view)) route.view = view
    const schedule = h.get('schedule')
    if (schedule) route.schedule = schedule
    const group = h.get('group')
    if (group !== null && route.view === 'workstreams') route.group = group
  } catch {
    // Malformed hash: fall back to the landing page.
  }
  return route
}

export function routeHash(r: Route): string {
  const h = new URLSearchParams({ page: r.page })
  if (r.schedule) h.set('schedule', r.schedule)
  h.set('view', r.view)
  if (r.group !== undefined && r.view === 'workstreams') h.set('group', r.group)
  return `#${h.toString()}`
}
