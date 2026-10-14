import { useEffect, useMemo, useRef, useState, type KeyboardEvent } from 'react'
import { dayNum } from '../../core/dates'
import { createLayout, ROW_HEIGHT, type Zoom } from '../../core/layout'
import type { Baseline, Criticality, Presentation, Schedule } from '../../core/model'
import { renderTimeline, type ColorToken } from '../../core/render/timeline'
import { downstreamOf, VIEW_IDS, type Filters, type ViewId, type ViewState } from '../../core/views'
import { api, type ViewResponse } from '../api'
import { localTodayISO } from '../format'
import { S } from '../strings'
import { Drawer, type PresentationEdit } from './Drawer'
import { columnsFor, Grid, gridWidth } from './Grid'
import { Loading } from './ui'

export interface WorkspaceState {
  selectedId: string | null
  showDeps: boolean
  compareId: string
  zoom: Zoom
  collapsed: string[]
  filters: Filters
  page: number
}

export const INITIAL_WORKSPACE: WorkspaceState = { selectedId: null, showDeps: true, compareId: '', zoom: 'week', collapsed: [], filters: {}, page: 1 }

export function hasFilters(f: Filters): boolean {
  return !!(f.text || f.groups?.length || f.criticality?.length || f.overlap)
}

/** Screen colors come from the live theme so the renderer's output matches exports exactly. */
function themeColor(): (t: ColorToken) => string {
  const cs = getComputedStyle(document.documentElement)
  const cache = new Map<string, string>()
  return (t) => {
    let v = cache.get(t)
    if (v === undefined) {
      v = cs.getPropertyValue(`--${t}`).trim()
      cache.set(t, v)
    }
    return v
  }
}

interface Props {
  schedule: Schedule
  presentation: Presentation
  baselines: Baseline[]
  view: ViewId
  focusGroup: string | undefined
  state: WorkspaceState
  setState: (fn: (s: WorkspaceState) => WorkspaceState) => void
  onView: (v: ViewId, group?: string) => void
  onPresentation: (p: Presentation) => void
  /** Bumped by the parent to ask for a scroll to today. */
  todayTick: number
  toast: (msg: string, err?: boolean) => void
}

export function Workspace({ schedule, presentation, baselines, view, focusGroup, state, setState, onView, onPresentation, todayTick, toast }: Props) {
  const [resp, setResp] = useState<ViewResponse | null>(null)
  const [loading, setLoading] = useState(true)
  const [filtersOpen, setFiltersOpen] = useState(false)
  const chartRef = useRef<HTMLDivElement>(null)
  const gridRef = useRef<HTMLDivElement>(null)
  const today = useMemo(localTodayISO, [])

  const viewState: ViewState = useMemo(
    () => ({ view, filters: state.filters, collapsed: state.collapsed, focusGroup: view === 'workstreams' ? focusGroup : undefined, page: state.page }),
    [view, state.filters, state.collapsed, focusGroup, state.page],
  )

  useEffect(() => {
    let live = true
    setLoading(true)
    api
      .getView(schedule.id, viewState, state.zoom, state.compareId || null)
      .then((r) => live && setResp(r))
      .catch((e) => live && toast((e as Error).message, true))
      .finally(() => live && setLoading(false))
    return () => {
      live = false
    }
  }, [schedule, presentation, viewState, state.zoom, state.compareId, toast])

  const baseline = baselines.find((b) => b.id === state.compareId) ?? null
  const hasFloat = schedule.tasks.some((t) => t.floatDays !== null)
  const hasCriticality = schedule.tasks.some((t) => t.critical !== 'unavailable')
  const cols = columnsFor(view, hasFloat)

  const layout = useMemo(() => (resp?.layout.range ? createLayout(resp.layout.range, state.zoom, resp.resolved.rows.length) : null), [resp, state.zoom])

  const svg = useMemo(() => {
    if (!resp || !layout) return ''
    return renderTimeline({
      schedule,
      presentation,
      resolved: resp.resolved,
      layout,
      selectedId: state.selectedId,
      showDeps: state.showDeps,
      baseline,
      today,
      color: themeColor(),
      font: { text: getComputedStyle(document.documentElement).getPropertyValue('--font').trim(), mono: getComputedStyle(document.documentElement).getPropertyValue('--mono').trim() },
    })
  }, [resp, layout, schedule, presentation, state.selectedId, state.showDeps, baseline, today])

  const downstreamCount = useMemo(() => {
    const cache = new Map<string, number>()
    return (id: string) => {
      let n = cache.get(id)
      if (n === undefined) {
        n = downstreamOf(schedule, id).size
        cache.set(id, n)
      }
      return n
    }
  }, [schedule])

  const rows = resp?.resolved.rows ?? []
  const taskRowIndex = useMemo(() => rows.flatMap((r, i) => (r.kind === 'task' ? [i] : [])), [rows])

  const select = (id: string | null) => setState((s) => ({ ...s, selectedId: id }))

  /** Bring a row's grid cell and its bar into the shared chart viewport. */
  const reveal = (rowIndex: number, taskId: string) => {
    const chart = chartRef.current
    const btn = gridRef.current?.querySelector<HTMLElement>(`[data-row="${rowIndex}"]`)
    btn?.focus({ preventScroll: true })
    if (!chart || !layout) return
    const y = layout.y(rowIndex)
    const headH = layout.y(0)
    if (y - headH < chart.scrollTop) chart.scrollTop = y - headH
    else if (y + ROW_HEIGHT > chart.scrollTop + chart.clientHeight) chart.scrollTop = y + ROW_HEIGHT - chart.clientHeight
    const t = schedule.tasks.find((x) => x.id === taskId)
    const date = t?.start ?? t?.finish
    if (date) {
      const gw = gridWidth(cols)
      const x = gw + layout.x(date)
      const visibleLeft = chart.scrollLeft + gw
      const visibleRight = chart.scrollLeft + chart.clientWidth
      if (x < visibleLeft || x > visibleRight - 40) chart.scrollLeft = Math.max(0, x - gw - 80)
    }
  }

  const onKeyDown = (e: KeyboardEvent<HTMLDivElement>) => {
    if (!taskRowIndex.length) return
    if (!['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(e.key)) return
    if ((e.target as HTMLElement).closest('input,select,textarea')) return
    e.preventDefault()
    const cur = taskRowIndex.findIndex((i) => {
      const r = rows[i]
      return r.kind === 'task' && r.taskId === state.selectedId
    })
    let next = cur
    if (e.key === 'ArrowDown') next = cur < 0 ? 0 : Math.min(cur + 1, taskRowIndex.length - 1)
    if (e.key === 'ArrowUp') next = cur < 0 ? 0 : Math.max(cur - 1, 0)
    if (e.key === 'Home') next = 0
    if (e.key === 'End') next = taskRowIndex.length - 1
    const ri = taskRowIndex[next]
    const r = rows[ri]
    if (r.kind !== 'task') return
    select(r.taskId)
    requestAnimationFrame(() => reveal(ri, r.taskId))
  }

  useEffect(() => {
    if (!todayTick || !layout || !chartRef.current) return
    const r = layout.range
    const td = dayNum(today)
    if (td < r.start || td > r.end) {
      toast(S.states.todayOutside)
      return
    }
    const chart = chartRef.current
    chart.scrollLeft = Math.max(0, layout.cx(today) - (chart.clientWidth - gridWidth(cols)) / 2)
  }, [todayTick]) // eslint-disable-line react-hooks/exhaustive-deps

  const onTimelineClick = (e: React.MouseEvent) => {
    const el = e.target as Element
    const bar = el.closest('g.bar') as SVGGElement | null
    if (bar?.dataset.taskId) return select(bar.dataset.taskId)
    const sum = el.closest('g.summary') as SVGGElement | null
    if (sum?.dataset.groupKey !== undefined) onView('workstreams', sum.dataset.groupKey)
  }

  const savePresentation = async (taskId: string, edit: PresentationEdit) => {
    const order = edit.order.trim() === '' ? null : Number(edit.order)
    if (order !== null && !Number.isFinite(order)) return toast(S.drawer.order, true)
    try {
      const p = await api.patchPresentation(schedule.id, {
        tasks: { [taskId]: { displayLabel: edit.displayLabel.trim() || null, displayGroup: edit.displayGroup.trim() || null, order, emphasis: edit.emphasis === 'none' ? null : edit.emphasis } },
      })
      onPresentation(p)
      toast(S.drawer.saved)
    } catch (e) {
      toast((e as Error).message, true)
    }
  }
  const resetPresentation = async (taskId: string) => {
    try {
      onPresentation(await api.resetPresentation(schedule.id, [taskId]))
      toast(S.drawer.resetDone)
    } catch (e) {
      toast((e as Error).message, true)
    }
  }

  const resolved = resp?.resolved
  const selectedTask = state.selectedId ? schedule.tasks.find((t) => t.id === state.selectedId) ?? null : null
  const clearFilters = () => setState((s) => ({ ...s, filters: {}, page: 1 }))

  const empty = (() => {
    if (!resolved) return null
    switch (resolved.emptyReason) {
      case 'no-tasks':
        return <EmptyState title={S.empty.noTasks} />
      case 'no-criticality-data':
        return <EmptyState title={S.empty.noCriticalityTitle} body={S.empty.noCriticalityBody} />
      case 'no-critical-tasks':
        return <EmptyState title={S.empty.noCriticalTitle} body={S.empty.noCriticalBody} />
      case 'unknown-group':
        return (
          <EmptyState title={S.empty.unknownGroupTitle} body={S.empty.unknownGroupBody(focusGroup ?? '')}>
            <button className="btn" onClick={() => onView('workstreams')}>
              {S.empty.backToAll}
            </button>
          </EmptyState>
        )
      case 'no-matching-tasks':
        return (
          <EmptyState title={S.empty.noMatchTitle} body={S.empty.noMatchBody}>
            <button className="btn primary" onClick={clearFilters}>
              {S.toolbar.clearFilters}
            </button>
          </EmptyState>
        )
    }
    return null
  })()

  return (
    <section className="page" aria-label={S.views[view].label}>
      <div className="toolbar">
        <div className="tabs" role="tablist" aria-label={S.viewsPage.tabsLabel}>
          {VIEW_IDS.map((v) => (
            <button key={v} className="tab" role="tab" aria-selected={v === view} onClick={() => onView(v)}>
              {S.views[v].label}
            </button>
          ))}
        </div>
        <div className="spacer" />
        <div className="controls">
          <button className="btn small" aria-expanded={filtersOpen} aria-controls="filter-panel" onClick={() => setFiltersOpen((o) => !o)}>
            {S.toolbar.filters}
            {hasFilters(state.filters) ? ' •' : ''}
          </button>
          <label>
            <input
              type="checkbox"
              checked={state.showDeps || view === 'dependencies'}
              disabled={view === 'executive' || view === 'dependencies'}
              onChange={(e) => setState((s) => ({ ...s, showDeps: e.target.checked }))}
            />
            {S.toolbar.deps}
          </label>
          <label>
            {S.toolbar.compareTo}
            <select value={state.compareId} onChange={(e) => setState((s) => ({ ...s, compareId: e.target.value }))}>
              <option value="">{S.toolbar.noBaseline}</option>
              {baselines.map((b) => (
                <option key={b.id} value={b.id}>
                  {b.name}
                </option>
              ))}
            </select>
          </label>
          <div className="seg" role="group" aria-label={S.toolbar.zoomLabel}>
            {(['day', 'week', 'month'] as Zoom[]).map((z) => (
              <button key={z} aria-pressed={state.zoom === z} onClick={() => setState((s) => ({ ...s, zoom: z }))}>
                {z === 'day' ? S.toolbar.zoomDay : z === 'week' ? S.toolbar.zoomWeek : S.toolbar.zoomMonth}
              </button>
            ))}
          </div>
        </div>
      </div>

      {filtersOpen && (
        <FilterPanel
          filters={state.filters}
          groups={resolved?.groups ?? []}
          onApply={(f) => {
            setState((s) => ({ ...s, filters: f, page: 1 }))
            setFiltersOpen(false)
          }}
          onClear={() => {
            clearFilters()
            setFiltersOpen(false)
          }}
          onClose={() => setFiltersOpen(false)}
        />
      )}

      <div className="content">
        {loading && !resp ? (
          <div className="chart-wrap">
            <Loading />
          </div>
        ) : empty ? (
          <div className="chart-wrap">{empty}</div>
        ) : resolved ? (
          <div className="chart-wrap" aria-busy={loading}>
            <div className="chart" ref={chartRef} onKeyDown={onKeyDown}>
              <div className="grid-pane" ref={gridRef}>
                <Grid
                  schedule={schedule}
                  presentation={presentation}
                  resolved={resolved}
                  cols={cols}
                  selectedId={state.selectedId}
                  downstreamCount={downstreamCount}
                  onSelect={(id) => select(id)}
                  onToggleGroup={(key) =>
                    setState((s) => ({ ...s, collapsed: s.collapsed.includes(key) ? s.collapsed.filter((k) => k !== key) : [...s.collapsed, key] }))
                  }
                  onOpenGroup={(key) => onView('workstreams', key)}
                />
              </div>
              <div className="tl-pane" onClick={onTimelineClick} dangerouslySetInnerHTML={{ __html: svg }} />
            </div>
            {selectedTask && (
              <Drawer
                schedule={schedule}
                presentation={presentation}
                task={selectedTask}
                groups={resolved.groups}
                baseline={baseline}
                onClose={() => select(null)}
                onSelect={(id) => select(id)}
                onSave={savePresentation}
                onReset={resetPresentation}
              />
            )}
          </div>
        ) : null}

        {resolved && !empty && (
          <>
            {view === 'workstreams' && focusGroup !== undefined && (
              <div>
                <button className="btn small" onClick={() => onView('workstreams')}>
                  {S.empty.backToAll}
                </button>
              </div>
            )}
            <Legend view={view} hasCriticality={hasCriticality} baseline={baseline} selected={!!state.selectedId} groups={resolved.groups} />
            <div className="reductions" role="status">
              <span>{S.reductions.showing(rows.filter((r) => r.kind === 'task').length + rows.filter((r) => r.kind === 'milestones').reduce((n, r) => n + (r.kind === 'milestones' ? r.taskIds.length : 0), 0))}</span>
              {resolved.reductions.byView > 0 && <span className="muted">· {S.reductions.byView(resolved.reductions.byView)}</span>}
              {resolved.reductions.byFilters > 0 && <span className="muted">· {S.reductions.byFilters(resolved.reductions.byFilters)}</span>}
              {resolved.reductions.inCollapsedGroups > 0 && <span className="muted">· {S.reductions.collapsed(resolved.reductions.inCollapsedGroups)}</span>}
              {resolved.reductions.onOtherPages > 0 && <span className="muted">· {S.reductions.otherPages(resolved.reductions.onOtherPages)}</span>}
              {hasFilters(state.filters) && (
                <button className="btn small" onClick={clearFilters}>
                  {S.toolbar.clearFilters}
                </button>
              )}
              <span className="spacer" />
              {resolved.pageCount > 1 && (
                <span className="sim-row" style={{ alignItems: 'center' }}>
                  <button className="btn small" disabled={resolved.page <= 1} onClick={() => setState((s) => ({ ...s, page: resolved.page - 1 }))}>
                    {S.toolbar.prevPage}
                  </button>
                  <span className="mono">{S.toolbar.pageOf(resolved.page, resolved.pageCount)}</span>
                  <button className="btn small" disabled={resolved.page >= resolved.pageCount} onClick={() => setState((s) => ({ ...s, page: resolved.page + 1 }))}>
                    {S.toolbar.nextPage}
                  </button>
                </span>
              )}
            </div>
          </>
        )}
      </div>
    </section>
  )
}

function EmptyState({ title, body, children }: { title: string; body?: string; children?: React.ReactNode }) {
  return (
    <div className="empty-state" role="status">
      <h2>{title}</h2>
      {body && <p>{body}</p>}
      {children}
    </div>
  )
}

function Legend({
  view,
  hasCriticality,
  baseline,
  selected,
  groups,
}: {
  view: ViewId
  hasCriticality: boolean
  baseline: Baseline | null
  selected: boolean
  groups: Array<{ key: string; label: string }>
}) {
  return (
    <div className="legend">
      {view === 'workstreams' ? (
        groups.map((g, i) => (
          <span key={g.key}>
            <i className="sw" style={{ background: `var(--lane-${(i % 6) + 1})` }} />
            {g.label}
          </span>
        ))
      ) : (
        <>
          {hasCriticality ? (
            <span>
              <i className="sw" style={{ background: 'var(--critical)' }} />
              {S.legend.critical}
            </span>
          ) : (
            <span className="na">{S.legend.noCriticality}</span>
          )}
          <span>
            <i className="sw" style={{ background: 'var(--task)' }} />
            {S.legend.task}
          </span>
        </>
      )}
      {view === 'dependencies' && (
        <span>
          <i className="sw" style={{ background: 'var(--highlight)' }} />
          {S.legend.downstream}
        </span>
      )}
      <span>
        <i className="sw" style={{ width: 9, height: 9, transform: 'rotate(45deg)', background: 'var(--milestone)' }} />
        {S.legend.milestone}
      </span>
      {view !== 'executive' && (
        <span>
          <i className="sw" style={{ height: 0, borderTop: '2px dotted var(--dep-unknown)', borderRadius: 0 }} />
          {S.legend.depUnknown}
        </span>
      )}
      {baseline && (
        <span>
          <i className="sw" style={{ height: 4, background: 'var(--baseline)' }} />
          {S.legend.baseline(baseline.name)}
        </span>
      )}
      <span>
        <i className="sw" style={{ width: 0, height: 12, borderLeft: '2px dashed var(--today)', borderRadius: 0 }} />
        {S.legend.today}
      </span>
      {view === 'dependencies' && !selected && <span>{S.legend.tipDeps}</span>}
      {view !== 'executive' && view !== 'dependencies' && <span>{S.legend.tipSelect}</span>}
    </div>
  )
}

function FilterPanel({
  filters,
  groups,
  onApply,
  onClear,
  onClose,
}: {
  filters: Filters
  groups: Array<{ key: string; label: string }>
  onApply: (f: Filters) => void
  onClear: () => void
  onClose: () => void
}) {
  const [text, setText] = useState(filters.text ?? '')
  const [group, setGroup] = useState(filters.groups?.[0] ?? '\u0000')
  const [crit, setCrit] = useState<Criticality | ''>(filters.criticality?.[0] ?? '')
  const [from, setFrom] = useState(filters.overlap?.start ?? '')
  const [to, setTo] = useState(filters.overlap?.finish ?? '')

  const apply = () => {
    const f: Filters = {}
    if (text.trim()) f.text = text.trim()
    if (group !== '\u0000') f.groups = [group]
    if (crit) f.criticality = [crit]
    if (from && to) f.overlap = { start: from <= to ? from : to, finish: from <= to ? to : from }
    onApply(f)
  }

  return (
    <form
      id="filter-panel"
      className="toolbar"
      aria-label={S.toolbar.filters}
      onSubmit={(e) => {
        e.preventDefault()
        apply()
      }}
      onKeyDown={(e) => {
        if (e.key === 'Escape') onClose()
      }}
    >
      <div className="controls">
        <label>
          {S.toolbar.filterText}
          <input type="search" value={text} placeholder={S.toolbar.filterTextPlaceholder} onChange={(e) => setText(e.target.value)} />
        </label>
        <label>
          {S.toolbar.filterGroup}
          <select value={group} onChange={(e) => setGroup(e.target.value)}>
            <option value={'\u0000'}>{S.toolbar.filterAllGroups}</option>
            {groups.map((g) => (
              <option key={g.key} value={g.key}>
                {g.label}
              </option>
            ))}
          </select>
        </label>
        <label>
          {S.toolbar.filterCriticality}
          <select value={crit} onChange={(e) => setCrit(e.target.value as Criticality | '')}>
            <option value="">{S.toolbar.filterAnyCriticality}</option>
            <option value="critical">{S.toolbar.critCritical}</option>
            <option value="not-critical">{S.toolbar.critNotCritical}</option>
            <option value="unavailable">{S.toolbar.critUnavailable}</option>
          </select>
        </label>
        <label>
          {S.toolbar.overlapFrom}
          <input type="date" value={from} onChange={(e) => setFrom(e.target.value)} />
        </label>
        <label>
          {S.toolbar.overlapTo}
          <input type="date" value={to} onChange={(e) => setTo(e.target.value)} />
        </label>
      </div>
      <div className="spacer" />
      <div className="sim-row">
        <button className="btn small" type="button" onClick={onClear}>
          {S.toolbar.clearFilters}
        </button>
        <button className="btn small" type="button" onClick={onClose}>
          {S.toolbar.closeFilters}
        </button>
        <button className="btn small primary" type="submit">
          {S.toolbar.applyFilters}
        </button>
      </div>
    </form>
  )
}
