import { ImportError, type ColumnMapping, type ParsedSchedule } from './types'

/** Parse an uploaded file by extension. Throws ImportError with a user-facing message. */
export async function parseFile(bytes: Uint8Array, filename: string, mapping?: ColumnMapping): Promise<ParsedSchedule> {
  const lower = filename.toLowerCase()
  if (lower.endsWith('.xlsx')) {
    const { parseXlsx } = await import('./xlsx')
    return parseXlsx(bytes, filename, mapping)
  }
  if (lower.endsWith('.xml')) {
    const { parseMspdi } = await import('./mspdi')
    return parseMspdi(new TextDecoder('utf-8').decode(bytes), filename)
  }
  throw new ImportError('Streamline can import Excel (.xlsx) or Microsoft Project XML (.xml) files.')
}
