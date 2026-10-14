import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Resvg } from '@resvg/resvg-js'
import { parseThemeTokens, tokenResolver } from '../core/render/theme'
import type { ColorToken } from '../core/render/timeline'
import { imageToPdf } from './pdf'

const root = join(dirname(fileURLToPath(import.meta.url)), '..', '..')

export function readThemeCss(): string {
  return readFileSync(join(root, 'src', 'web', 'theme.css'), 'utf8')
}

export interface ExportTheme {
  color: (token: ColorToken) => string
  font: { text: string; mono: string }
}

export function exportTheme(css: string): ExportTheme {
  const tokens = parseThemeTokens(css)
  const text = tokens.font
  const mono = tokens.mono
  if (!text || !mono) throw new Error('Theme tokens --font and --mono are required for exports.')
  return { color: tokenResolver(tokens), font: { text, mono } }
}

const SCALE = 2

function rasterize(svg: string) {
  const resvg = new Resvg(svg, {
    background: '#ffffff',
    fitTo: { mode: 'zoom', value: SCALE },
    font: { loadSystemFonts: true },
  })
  return resvg.render()
}

export function svgDocument(svg: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?>\n${svg}\n`
}

export function svgToPng(svg: string): Buffer {
  return rasterize(svg).asPng()
}

export function svgToPdf(svg: string): Buffer {
  const img = rasterize(svg)
  const rgba = img.pixels
  const rgb = Buffer.alloc(img.width * img.height * 3)
  for (let i = 0, o = 0; i < rgba.length; i += 4, o += 3) {
    const a = rgba[i + 3] / 255
    rgb[o] = Math.round(rgba[i] * a + 255 * (1 - a))
    rgb[o + 1] = Math.round(rgba[i + 1] * a + 255 * (1 - a))
    rgb[o + 2] = Math.round(rgba[i + 2] * a + 255 * (1 - a))
  }
  const w = img.width / SCALE
  const h = img.height / SCALE
  return imageToPdf(rgb, img.width, img.height, w * 0.75, h * 0.75)
}
