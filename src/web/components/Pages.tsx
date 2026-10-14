import { useEffect, useState } from 'react'
import type { Baseline } from '../../core/model'
import { VIEW_IDS, type ViewId } from '../../core/views'
import { api, type ExportFormat, type OpenApiDoc } from '../api'
import { shortDate } from '../format'
import { S } from '../strings'
import { Loading } from './ui'

export function ViewsPage({ onOpen }: { onOpen: (v: ViewId) => void }) {
  return (
    <section className="page" aria-label={S.viewsPage.title}>
      <div className="content">
        <div className="panel">
          <h2>{S.viewsPage.title}</h2>
          <p>{S.viewsPage.body}</p>
          <div className="cards">
            {VIEW_IDS.map((v) => (
              <div key={v} className="card">
                <h3>{S.views[v].label}</h3>
                <div className="who">{S.viewsPage.forWho(S.views[v].who)}</div>
                <p>{S.views[v].desc}</p>
                <div>
                  <button className="btn small" onClick={() => onOpen(v)}>
                    {S.viewsPage.open}
                  </button>
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </section>
  )
}

export function BaselinesPage({
  baselines,
  compareId,
  onSave,
  onCompare,
  onDelete,
}: {
  baselines: Baseline[]
  compareId: string
  onSave: () => void
  onCompare: (id: string) => void
  onDelete: (b: Baseline) => void
}) {
  const endOf = (b: Baseline) => b.tasks.reduce<string | null>((m, t) => (t.finish && (!m || t.finish > m) ? t.finish : m), null)
  return (
    <section className="page" aria-label={S.baselines.title}>
      <div className="content">
        <div className="panel">
          <div className="panel-head" style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 10 }}>
            <div style={{ flex: '1 1 260px' }}>
              <h2>{S.baselines.title}</h2>
              <p>{S.baselines.body}</p>
            </div>
            <button className="btn primary" onClick={onSave}>
              {S.baselines.save}
            </button>
          </div>
          <div className="table-box">
            <table>
              <thead>
                <tr>
                  <th>{S.baselines.colName}</th>
                  <th>{S.baselines.colSaved}</th>
                  <th>{S.baselines.colTasks}</th>
                  <th>{S.baselines.colEnd}</th>
                  <th>
                    <span className="sr-only-text">{S.baselines.compare}</span>
                  </th>
                </tr>
              </thead>
              <tbody>
                {baselines.length ? (
                  baselines.map((b) => (
                    <tr key={b.id}>
                      <td>
                        <b>{b.name}</b>
                      </td>
                      <td>{new Date(b.createdAt).toLocaleString('en-GB', { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' })}</td>
                      <td className="mono">{b.tasks.length}</td>
                      <td className="mono">{shortDate(endOf(b))}</td>
                      <td>
                        <div className="sim-row">
                          <button className="btn small" aria-pressed={compareId === b.id} onClick={() => onCompare(b.id)}>
                            {compareId === b.id ? S.baselines.comparing : S.baselines.compare}
                          </button>
                          <button className="btn small" onClick={() => onDelete(b)}>
                            {S.baselines.delete}
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan={5} className="empty">
                      {S.baselines.none}
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </section>
  )
}

export const EXPORT_FORMATS: ExportFormat[] = ['png', 'svg', 'pdf', 'csv', 'json']

export function ExportsPage({ onExport }: { onExport: (f: ExportFormat) => void }) {
  const label: Record<ExportFormat, string> = { png: S.exportsPage.png, svg: S.exportsPage.svg, pdf: S.exportsPage.pdf, csv: S.exportsPage.csv, json: S.exportsPage.json }
  return (
    <section className="page" aria-label={S.exportsPage.title}>
      <div className="content">
        <div className="panel">
          <h2>{S.exportsPage.title}</h2>
          <p>{S.exportsPage.body}</p>
          <div className="sim-row">
            {EXPORT_FORMATS.map((f) => (
              <button key={f} className="btn" onClick={() => onExport(f)}>
                {label[f]}
              </button>
            ))}
          </div>
        </div>
      </div>
    </section>
  )
}

/** Rendered from the OpenAPI document the server publishes, so it cannot drift from the real API. */
export function ApiPage() {
  const [doc, setDoc] = useState<OpenApiDoc | null>(null)
  const [error, setError] = useState(false)
  useEffect(() => {
    api.openApi().then(setDoc, () => setError(true))
  }, [])

  const rows = doc
    ? Object.entries(doc.paths).flatMap(([path, ops]) =>
        Object.entries(ops)
          .filter(([m]) => ['get', 'post', 'put', 'patch', 'delete'].includes(m))
          .map(([method, op]) => ({ path, method: method.toUpperCase(), op })),
      )
    : []

  const example = (v: unknown) => (v === undefined ? '' : typeof v === 'string' ? v : JSON.stringify(v))

  return (
    <section className="page" aria-label={S.nav.api}>
      <div className="content">
        <div className="panel">
          <h2>{S.api.title}</h2>
          <p>{S.api.body}</p>
          {error ? (
            <p className="empty">{S.api.loadFailed}</p>
          ) : !doc ? (
            <Loading />
          ) : (
            <div className="table-box">
              <table>
                <thead>
                  <tr>
                    <th>{S.api.colEndpoint}</th>
                    <th>{S.api.colWhat}</th>
                    <th>{S.api.colExample}</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map(({ path, method, op }) => (
                    <tr key={`${method} ${path}`}>
                      <td style={{ whiteSpace: 'nowrap' }}>
                        <span className="method">{method}</span> <span className="mono">{path}</span>
                      </td>
                      <td>{op.summary}</td>
                      <td className="mono" style={{ maxWidth: 360, overflowWrap: 'anywhere' }}>
                        {example(op['x-example']?.request)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    </section>
  )
}
