// Reads the theme tokens from theme.css so exports resolve the same colors as the screen.

import type { ColorToken } from './timeline'

/** Custom properties declared in the first `:root { ... }` block, keyed without the leading `--`. */
export function parseThemeTokens(css: string): Record<string, string> {
  const stripped = css.replace(/\/\*[\s\S]*?\*\//g, '')
  const open = stripped.search(/:root\s*\{/)
  if (open < 0) return {}
  const start = stripped.indexOf('{', open) + 1
  const end = stripped.indexOf('}', start)
  const body = stripped.slice(start, end < 0 ? undefined : end)
  const tokens: Record<string, string> = {}
  for (const m of body.matchAll(/--([\w-]+)\s*:\s*([^;]+);/g)) tokens[m[1]] = m[2].trim()
  return tokens
}

export function tokenResolver(tokens: Record<string, string>): (t: ColorToken) => string {
  return (t) => {
    const v = tokens[t]
    if (v === undefined || v === '') throw new Error(`Theme token --${t} is not defined`)
    return v
  }
}
