import { useEffect, useState } from 'react'
import { dayNum, spanDays } from '../../core/dates'
import type { Baseline, Emphasis, Presentation, Schedule, Task } from '../../core/model'
import { downstreamOf, successorsOf } from '../../core/views'
import { displayName, hasLabelOverride, shortDate } from '../format'
import { S } from '../strings'
import { CloseIcon, Unavailable } from './ui'

export interface PresentationEdit {
  displayLabel: string
  displayGroup: string
  order: string
  emphasis: Emphasis
}

export function Drawer({
  schedule,
  presentation,
  task,
  groups,
  baseline,
  onClose,
  onSelect,
  onSave,
  onReset,
}: {
  schedule: Schedule
  presentation: Presentation
  task: Task
  groups: Array<{ key: string; label: string }>
  baseline: Baseline | null
  onClose: () => void
  onSelect: (id: string) => void
  onSave: (taskId: string, edit: PresentationEdit) => Promise<void>
  onReset: (taskId: string) => Promise<void>
}) {
  const p = presentation.tasks[task.id] ?? {}
  const [edit, setEdit] = useState<PresentationEdit>(() => fromPresentation())
  const [busy, setBusy] = useState(false)

  function fromPresentation(): PresentationEdit {
    return { displayLabel: p.displayLabel ?? '', displayGroup: p.displayGroup ?? '', order: p.order === undefined ? '' : String(p.order), emphasis: p.emphasis ?? 'none' }
  }
  // Reload the form when the selection or the stored presentation changes.
  useEffect(() => setEdit(fromPresentation()), [task.id, presentation.updatedAt]) // eslint-disable-line react-hooks/exhaustive-deps

  const byId = new Map(schedule.tasks.map((t) => [t.id, t]))
  const ws = schedule.workstreams.find((w) => w.id === task.workstream)
  const preds = task.deps.map((d) => ({ dep: d, task: byId.get(d.predecessorId) })).filter((x): x is { dep: typeof x.dep; task: Task } => !!x.task)
  const succs = successorsOf(schedule, task.id)
    .map((id) => byId.get(id))
    .filter((t): t is Task => !!t)
  const down = downstreamOf(schedule, task.id).size
  const bt = baseline?.tasks.find((b) => b.id === task.id)
  const variance = bt?.finish && task.finish ? dayNum(task.finish) - dayNum(bt.finish) : null
  const dur = task.milestone ? null : spanDays(task.start, task.finish)
  const shown = displayName(task, presentation)

  const chip = (t: Task, suffix?: string) => (
    <button key={t.id + (suffix ?? '')} className="chip" type="button" onClick={() => onSelect(t.id)} title={suffix ? `${displayName(t, presentation)} · ${suffix}` : undefined}>
      {displayName(t, presentation)}
      {suffix && <span className="muted"> · {suffix}</span>}
    </button>
  )

  return (
    <aside className="drawer" aria-label={S.drawer.label}>
      <header>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div className="muted" style={{ fontSize: 'var(--text-sm)', marginBottom: 2 }}>
            {ws?.name ?? task.workstream ?? S.unavailable} · <span className="mono">{task.id}</span>
          </div>
          <h2>
            {shown}
            {hasLabelOverride(task, presentation) && <span className="sr-only-text">{` (${S.drawer.sourceName}: ${task.name})`}</span>}
          </h2>
        </div>
        <button className="icon-btn" type="button" aria-label={S.drawer.close} onClick={onClose}>
          <CloseIcon />
        </button>
      </header>

      <form
        className="body"
        onSubmit={async (e) => {
          e.preventDefault()
          setBusy(true)
          await onSave(task.id, edit).finally(() => setBusy(false))
        }}
      >
        <h3 className="section-title">{S.drawer.presentation}</h3>
        <p className="note">{S.drawer.presentationNote}</p>
        <label className="field">
          {S.drawer.displayLabel}
          <input value={edit.displayLabel} maxLength={200} placeholder={task.name} onChange={(e) => setEdit({ ...edit, displayLabel: e.target.value })} />
          <span className="note">{S.drawer.displayLabelHint}</span>
        </label>
        <div className="field-row">
          <label className="field">
            {S.drawer.displayGroup}
            <input
              list="display-groups"
              value={edit.displayGroup}
              maxLength={120}
              placeholder={S.drawer.displayGroupSource}
              onChange={(e) => setEdit({ ...edit, displayGroup: e.target.value })}
            />
            <datalist id="display-groups">
              {groups.map((g) => (
                <option key={g.key} value={g.key}>
                  {g.label}
                </option>
              ))}
            </datalist>
          </label>
          <label className="field">
            {S.drawer.order}
            <input type="number" step="1" value={edit.order} onChange={(e) => setEdit({ ...edit, order: e.target.value })} />
          </label>
        </div>
        <label className="field">
          {S.drawer.emphasis}
          <select value={edit.emphasis} onChange={(e) => setEdit({ ...edit, emphasis: e.target.value as Emphasis })}>
            <option value="none">{S.drawer.emphasisNone}</option>
            <option value="highlight">{S.drawer.emphasisHighlight}</option>
            <option value="muted">{S.drawer.emphasisMuted}</option>
          </select>
        </label>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          <button className="btn primary" type="submit" disabled={busy}>
            {S.drawer.save}
          </button>
          <button className="btn" type="button" disabled={busy} onClick={() => onReset(task.id)}>
            {S.drawer.reset}
          </button>
        </div>

        <section className="readonly" aria-labelledby="source-facts">
          <h3 className="section-title" id="source-facts">
            {S.drawer.sourceFacts}
          </h3>
          <p className="note">{S.drawer.sourceNote}</p>
          <dl className="facts">
            <dt>{S.drawer.sourceName}</dt>
            <dd>{task.name}</dd>
            <dt>{S.drawer.id}</dt>
            <dd className="mono">{task.id}</dd>
            <dt>{S.drawer.workstream}</dt>
            <dd>{ws?.name ?? task.workstream ?? <Unavailable />}</dd>
            <dt>{S.drawer.owner}</dt>
            <dd>{task.owner ?? <Unavailable />}</dd>
            <dt>{S.drawer.start}</dt>
            <dd className="mono">{task.start ? shortDate(task.start) : <Unavailable />}</dd>
            <dt>{S.drawer.finish}</dt>
            <dd className="mono">{task.finish ? shortDate(task.finish) : <Unavailable />}</dd>
            <dt>{S.drawer.duration}</dt>
            <dd className="mono">{task.milestone ? S.drawer.milestone : dur === null ? <Unavailable /> : S.drawer.durationDays(dur)}</dd>
            <dt>{S.drawer.criticality}</dt>
            <dd className={task.critical === 'critical' ? 'crit-text' : task.critical === 'unavailable' ? 'na' : undefined}>{S.crit[task.critical]}</dd>
            <dt>{S.drawer.float}</dt>
            <dd className="mono">{task.floatDays === null ? <Unavailable /> : S.drawer.floatDays(task.floatDays)}</dd>
            {task.isSummary && (
              <>
                <dt>{S.drawer.summary}</dt>
                <dd>{S.drawer.yes}</dd>
              </>
            )}
            <dt>{S.drawer.downstream}</dt>
            <dd>{S.drawer.downstreamValue(down)}</dd>
            {bt && (
              <>
                <dt>{S.drawer.vsBaseline}</dt>
                <dd className={`mono ${variance !== null && variance > 0 ? 'neg' : ''}`}>
                  {variance === null ? <Unavailable /> : variance === 0 ? S.drawer.onBaseline : S.drawer.variance(variance)}
                </dd>
              </>
            )}
          </dl>
          <div className="field">
            {S.drawer.waitsOn}
            <div className="linklist">
              {preds.length ? preds.map(({ dep, task: t }) => chip(t, `${S.drawer.depType[dep.type]}, ${S.drawer.lag(dep.lagDays)}`)) : <span className="muted">{S.drawer.nothing}</span>}
            </div>
          </div>
          <div className="field">
            {S.drawer.drives}
            <div className="linklist">{succs.length ? succs.map((t) => chip(t)) : <span className="muted">{S.drawer.nothing}</span>}</div>
          </div>
        </section>
      </form>
    </aside>
  )
}
