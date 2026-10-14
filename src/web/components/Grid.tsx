import type { Presentation, Schedule, Task } from '../../core/model'
import { spanDays } from '../../core/dates'
import type { ResolvedView, Row, ViewId } from '../../core/views'
import { displayName, hasLabelOverride, shortDate } from '../format'
import { S } from '../../core/strings'
import { Chevron, Unavailable } from './ui'

export type ColKey = 'item' | 'task' | 'owner' | 'ws' | 'start' | 'finish' | 'dur' | 'float' | 'down'

export const COLS: Record<ColKey, { label: string; w: number }> = {
  item: { label: S.cols.item, w: 230 },
  task: { label: S.cols.task, w: 220 },
  owner: { label: S.cols.owner, w: 88 },
  ws: { label: S.cols.ws, w: 100 },
  start: { label: S.cols.start, w: 66 },
  finish: { label: S.cols.finish, w: 66 },
  dur: { label: S.cols.dur, w: 48 },
  float: { label: S.cols.float, w: 54 },
  down: { label: S.cols.down, w: 92 },
}

const VIEW_COLS: Record<ViewId, ColKey[]> = {
  executive: ['item', 'start', 'finish'],
  detailed: ['task', 'owner', 'start', 'finish', 'dur', 'float'],
  dependencies: ['task', 'start', 'finish', 'down'],
  workstreams: ['task', 'owner', 'start', 'finish', 'dur'],
  critical: ['task', 'ws', 'start', 'finish', 'dur', 'float'],
}

/** Columns for a view. Float is omitted entirely when the source provides none. */
export function columnsFor(view: ViewId, hasFloat: boolean): ColKey[] {
  return VIEW_COLS[view].filter((c) => c !== 'float' || hasFloat)
}

export function gridWidth(cols: ColKey[]): number {
  return cols.reduce((s, k) => s + COLS[k].w, 0)
}

interface GridProps {
  schedule: Schedule
  presentation: Presentation
  resolved: ResolvedView
  cols: ColKey[]
  selectedId: string | null
  downstreamCount: (id: string) => number
  onSelect: (id: string) => void
  onToggleGroup: (key: string) => void
  onOpenGroup: (key: string) => void
}

export function Grid({ schedule, presentation, resolved, cols, selectedId, downstreamCount, onSelect, onToggleGroup, onOpenGroup }: GridProps) {
  const byId = new Map(schedule.tasks.map((t) => [t.id, t]))
  const groupLabel = new Map(resolved.groups.map((g) => [g.key, g.label]))
  const template = cols.map((k) => `${COLS[k].w}px`).join(' ')
  const style = { gridTemplateColumns: template }

  const taskCell = (key: ColKey, t: Task, depth: number) => {
    switch (key) {
      case 'item':
      case 'task': {
        const label = displayName(t, presentation)
        return (
          <span className="name" style={{ paddingLeft: 24 + depth * 12 }}>
            <span className="label" title={hasLabelOverride(t, presentation) ? `${label} — ${S.drawer.sourceName}: ${t.name}` : t.name}>
              {label}
              {hasLabelOverride(t, presentation) && <span className="sr-only-text">{` (${S.drawer.sourceName}: ${t.name})`}</span>}
            </span>
            {t.critical === 'critical' && <span className="pill crit">{S.crit.critical}</span>}
          </span>
        )
      }
      case 'owner':
        return <span className="muted">{t.owner ?? <Unavailable />}</span>
      case 'ws': {
        const key = presentation.tasks[t.id]?.displayGroup ?? t.workstream ?? ''
        return <span className="muted">{groupLabel.get(key) ?? (key || <Unavailable />)}</span>
      }
      case 'start':
        return <span className="mono">{t.start ? shortDate(t.start) : <Unavailable />}</span>
      case 'finish':
        return <span className="mono">{t.finish ? shortDate(t.finish) : <Unavailable />}</span>
      case 'dur': {
        if (t.milestone) return <span className="mono muted">0d</span>
        const d = spanDays(t.start, t.finish)
        return <span className="mono muted">{d === null ? <Unavailable /> : `${d}d`}</span>
      }
      case 'float': {
        const f = t.floatDays
        if (f === null) return <span className="mono"><Unavailable /></span>
        return <span className={`mono ${f < 0 ? 'neg' : f === 0 ? 'crit-text' : 'muted'}`}>{`${f}d`}</span>
      }
      case 'down':
        return <span className="mono muted">{downstreamCount(t.id)}</span>
    }
  }

  const spanCells = (key: ColKey, r: { start: string | null; finish: string | null }) => {
    if (key === 'start') return <span className="mono muted">{r.start ? shortDate(r.start) : ''}</span>
    if (key === 'finish') return <span className="mono muted">{r.finish ? shortDate(r.finish) : ''}</span>
    return <span />
  }

  const rowEl = (r: Row, i: number) => {
    if (r.kind === 'task') {
      const t = byId.get(r.taskId)
      if (!t) return null
      const sel = t.id === selectedId
      return (
        <button
          key={`t-${t.id}`}
          className={`grow ${sel ? 'selected' : ''}`}
          style={style}
          data-row={i}
          aria-pressed={sel}
          aria-label={`${displayName(t, presentation)}${hasLabelOverride(t, presentation) ? ` (${S.drawer.sourceName}: ${t.name})` : ''}`}
          onClick={() => onSelect(t.id)}
        >
          {cols.map((k) => (
            <span key={k}>{taskCell(k, t, r.depth)}</span>
          ))}
        </button>
      )
    }
    if (r.kind === 'group') {
      return (
        <button key={`g-${r.key}-${i}`} className="grow group" style={style} data-row={i} aria-expanded={!r.collapsed} onClick={() => onToggleGroup(r.key)}>
          {cols.map((k, n) =>
            n === 0 ? (
              <span key={k} className="name">
                <Chevron closed={r.collapsed} />
                {r.label} <span className="muted count">{r.count}</span>
              </span>
            ) : (
              <span key={k}>{spanCells(k, r)}</span>
            ),
          )}
        </button>
      )
    }
    if (r.kind === 'summary') {
      return (
        <button key={`s-${r.key}`} className="grow summary" style={style} data-row={i} title={r.label} onClick={() => onOpenGroup(r.key)}>
          {cols.map((k, n) =>
            n === 0 ? (
              <span key={k} className="name">
                <b>{r.label}</b>
                {r.critical === 'critical' && <span className="pill crit">{S.crit.critical}</span>}
              </span>
            ) : (
              <span key={k}>{spanCells(k, r)}</span>
            ),
          )}
        </button>
      )
    }
    const ms = r.taskIds.map((id) => byId.get(id)).filter((t): t is Task => !!t)
    const first = ms.reduce<string | null>((m, t) => (t.start && (!m || t.start < m) ? t.start : m), null)
    const last = ms.reduce<string | null>((m, t) => (t.start && (!m || t.start > m) ? t.start : m), null)
    return (
      <div key="milestones" className="grow group" style={{ ...style, textTransform: 'none', letterSpacing: 0 }} data-row={i}>
        {cols.map((k, n) =>
          n === 0 ? (
            <span key={k} className="name">
              <b>{r.label}</b>
            </span>
          ) : (
            <span key={k}>{spanCells(k, { start: first, finish: last })}</span>
          ),
        )}
      </div>
    )
  }

  return (
    <div style={{ width: gridWidth(cols) }}>
      <div className="ghead" style={style}>
        {cols.map((k) => (
          <span key={k}>{COLS[k].label}</span>
        ))}
      </div>
      {resolved.rows.map(rowEl)}
    </div>
  )
}
