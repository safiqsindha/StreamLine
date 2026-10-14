import { useRef, useState } from 'react'
import { fmtShort } from '../../core/dates'
import type { ColumnMapping, ImportPreview, MappedField } from '../../core/import/types'
import type { Schedule } from '../../core/model'
import { api } from '../api'
import { S } from '../strings'
import { Loading, Modal } from './ui'

const FIELDS: MappedField[] = ['id', 'name', 'workstream', 'owner', 'start', 'finish', 'milestone', 'critical', 'float', 'parent', 'summary', 'predecessors']

export function ImportDialog({
  open,
  replace,
  onClose,
  onAccepted,
}: {
  open: boolean
  /** When set, the import replaces this schedule after reconciliation. */
  replace: { id: string; name: string } | null
  onClose: () => void
  onAccepted: (s: Schedule, replaced: boolean) => void
}) {
  const [file, setFile] = useState<File | null>(null)
  const [preview, setPreview] = useState<ImportPreview | null>(null)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [name, setName] = useState('')
  const inputRef = useRef<HTMLInputElement>(null)

  const reset = () => {
    setFile(null)
    setPreview(null)
    setBusy(false)
    setError(null)
    setName('')
  }
  const close = () => {
    reset()
    onClose()
  }

  const load = async (f: File, mapping?: ColumnMapping) => {
    setBusy(true)
    setError(null)
    try {
      const p = await api.previewImport(f, { mapping, replaceId: replace?.id })
      setPreview(p)
      if (!mapping) setName(replace?.name ?? p.parsed.name)
    } catch (e) {
      setPreview(null)
      setError((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  const accept = async () => {
    if (!preview) return
    setBusy(true)
    try {
      const s = await api.acceptImport(preview.previewId, name.trim() || undefined)
      reset()
      onAccepted(s, !!replace)
    } catch (e) {
      setError((e as Error).message)
      setBusy(false)
    }
  }

  const setField = (field: MappedField, column: string) => {
    if (!file || !preview) return
    const mapping: ColumnMapping = { ...preview.parsed.mapping }
    if (column) mapping[field] = column
    else delete mapping[field]
    void load(file, mapping)
  }

  const title = replace ? S.importer.reimportTitle(replace.name) : S.importer.title
  const p = preview
  const sum = p?.summary
  const rec = p?.reconciliation
  const isSheet = p?.parsed.kind === 'xlsx'

  return (
    <Modal open={open} onClose={close} label={title} wide>
      <div className="dlg-body">
        <h2>{title}</h2>
        <label className="field">
          {S.importer.chooseFile}
          <input
            ref={inputRef}
            type="file"
            accept=".xlsx,.xml"
            onChange={(e) => {
              const f = e.target.files?.[0] ?? null
              setFile(f)
              setPreview(null)
              if (f) void load(f)
            }}
          />
          <span className="note">{S.importer.fileHint}</span>
        </label>

        {busy && <Loading label={S.importer.parsing} />}
        {error && (
          <p role="alert" className="pill bad" style={{ padding: '8px 10px', fontWeight: 500, fontSize: 'var(--text-md)' }}>
            {error}
          </p>
        )}

        {p && sum && !busy && (
          <>
            <h3 className="section-title">{S.importer.preview}</h3>
            <div className="stat-grid">
              <div className="stat">
                <span className="muted">{S.importer.rows}</span>
                <b>{sum.rowCount}</b>
              </div>
              <div className="stat">
                <span className="muted">{S.importer.lanes}</span>
                <b>{sum.laneCount}</b>
              </div>
              <div className="stat">
                <span className="muted">{S.importer.dateRange}</span>
                <b className="mono">{sum.dateRange ? `${fmtShort(sum.dateRange.start)} – ${fmtShort(sum.dateRange.finish)}` : S.unavailable}</b>
              </div>
              <div className="stat">
                <span className="muted">{S.importer.criticality}</span>
                <b>{sum.hasCriticality ? S.importer.provided : S.importer.notProvided}</b>
              </div>
              <div className="stat">
                <span className="muted">{S.importer.float}</span>
                <b>{sum.hasFloat ? S.importer.provided : S.importer.notProvided}</b>
              </div>
            </div>

            {isSheet ? (
              <div className="field">
                <span>{S.importer.mapping}</span>
                <span className="note">{S.importer.mappingHint}</span>
                <div className="field-row">
                  {FIELDS.map((f) => (
                    <label key={f} className="field">
                      {S.importer.fields[f]}
                      <select value={p.parsed.mapping[f] ?? ''} onChange={(e) => setField(f, e.target.value)}>
                        <option value="">{S.importer.notMapped}</option>
                        {p.parsed.detectedColumns.map((c) => (
                          <option key={c} value={c}>
                            {c}
                          </option>
                        ))}
                      </select>
                    </label>
                  ))}
                </div>
              </div>
            ) : (
              <p className="note">
                {S.importer.detected}: {p.parsed.detectedColumns.join(', ')}
              </p>
            )}

            {p.parsed.unmapped.length > 0 && (
              <p className="note">
                {S.importer.unmapped}: {p.parsed.unmapped.join(', ')}
              </p>
            )}

            <div className="field">
              <span>{S.importer.issues}</span>
              {p.parsed.issues.length ? (
                <ul className="issue-list">
                  {p.parsed.issues.map((i, n) => (
                    <li key={n}>{i.message}</li>
                  ))}
                </ul>
              ) : (
                <span className="note">{S.importer.noIssues}</span>
              )}
            </div>

            {rec && (
              <div className="field">
                <span>{S.importer.changes}</span>
                <ul className="issue-list">
                  <li>{S.importer.added(rec.added.length)}</li>
                  <li>{S.importer.removed(rec.removed.length)}</li>
                  <li>{S.importer.changed(rec.changed.length)}</li>
                  {rec.orphanedOverrides.length > 0 && <li>{S.importer.orphans(rec.orphanedOverrides.length)}</li>}
                </ul>
              </div>
            )}

            {!replace && (
              <label className="field">
                {S.importer.nameLabel}
                <input value={name} maxLength={120} onChange={(e) => setName(e.target.value)} />
              </label>
            )}
          </>
        )}

        <div className="dlg-actions">
          <button className="btn" type="button" onClick={close}>
            {S.importer.cancel}
          </button>
          <button className="btn primary" type="button" disabled={!p || busy || p.summary.rowCount === 0} onClick={accept}>
            {S.importer.accept}
          </button>
        </div>
      </div>
    </Modal>
  )
}
