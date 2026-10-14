import { useEffect, useRef, useState } from 'react'
import { S } from '../../core/strings'
import { Modal } from './ui'

export interface Command {
  label: string
  hint: string
  run: () => void
}

export function Palette({ open, commands, onClose }: { open: boolean; commands: Command[]; onClose: () => void }) {
  const [q, setQ] = useState('')
  const [idx, setIdx] = useState(0)
  const listRef = useRef<HTMLUListElement>(null)

  useEffect(() => {
    if (open) {
      setQ('')
      setIdx(0)
    }
  }, [open])

  const query = q.trim().toLowerCase()
  const items = commands.filter((c) => !query || c.label.toLowerCase().includes(query) || c.hint.toLowerCase().includes(query))
  const cur = Math.min(idx, Math.max(0, items.length - 1))

  useEffect(() => {
    listRef.current?.querySelector('[aria-selected="true"]')?.scrollIntoView({ block: 'nearest' })
  }, [cur, q])

  const run = (i: number) => {
    const c = items[i]
    if (!c) return
    onClose()
    c.run()
  }

  return (
    <Modal open={open} onClose={onClose} label={S.palette.label}>
      <input
        className="pal-input"
        autoFocus
        value={q}
        placeholder={S.palette.placeholder}
        aria-label={S.palette.label}
        aria-controls="pal-list"
        aria-activedescendant={items.length ? `pal-${cur}` : undefined}
        role="combobox"
        aria-expanded="true"
        autoComplete="off"
        onChange={(e) => {
          setQ(e.target.value)
          setIdx(0)
        }}
        onKeyDown={(e) => {
          if (e.key === 'ArrowDown') {
            setIdx(Math.min(cur + 1, items.length - 1))
            e.preventDefault()
          } else if (e.key === 'ArrowUp') {
            setIdx(Math.max(cur - 1, 0))
            e.preventDefault()
          } else if (e.key === 'Enter') {
            run(cur)
            e.preventDefault()
          }
        }}
      />
      <ul className="pal-list" id="pal-list" role="listbox" ref={listRef}>
        {items.length ? (
          items.map((c, i) => (
            <li key={`${c.hint}-${c.label}-${i}`} id={`pal-${i}`} role="option" aria-selected={i === cur} onClick={() => run(i)}>
              <span>{c.label}</span>
              <span className="muted">{c.hint}</span>
            </li>
          ))
        ) : (
          <li aria-disabled="true">
            <span className="muted">{S.palette.noMatches}</span>
          </li>
        )}
      </ul>
    </Modal>
  )
}
