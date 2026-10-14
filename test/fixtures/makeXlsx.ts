import ExcelJS from 'exceljs'
import { mkdir, writeFile } from 'node:fs/promises'
import { join } from 'node:path'

export type Cell = string | number | boolean | Date | null

export interface MakeOptions {
  sheetName?: string
  /** Rows written above the header row (e.g. a title block). */
  preamble?: Cell[][]
}

/** Build an .xlsx file in memory. Date values become real Excel date cells. */
export async function makeXlsx(header: string[], rows: Cell[][], opts: MakeOptions = {}): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet(opts.sheetName ?? 'Schedule')
  for (const line of [...(opts.preamble ?? []), header, ...rows]) {
    const row = ws.addRow(line)
    row.eachCell((cell) => {
      if (cell.value instanceof Date) cell.numFmt = 'yyyy-mm-dd'
    })
  }
  return new Uint8Array(await wb.xlsx.writeBuffer())
}

/** A date cell at midnight UTC, as Excel stores it. */
export const d = (iso: string): Date => new Date(`${iso}T00:00:00Z`)

const HEADER = ['Task ID', 'Task Name', 'Workstream', 'Owner', 'Start', 'Finish', 'Milestone', 'Predecessors', 'Critical', 'Total Float']

// id, name, lane, owner, start, finish, milestone, predecessors, critical, float
const SAMPLE: Cell[][] = [
  ['D-01', 'Requirements baseline', 'Design', 'Avery', d('2026-01-05'), d('2026-01-23'), false, null, 'Yes', 0],
  ['D-02', 'Architecture review', 'Design', 'Avery', d('2026-01-26'), d('2026-02-13'), false, 'D-01FS', 'Yes', 0],
  ['D-03', 'Mechanical enclosure design', 'Design', 'Blake', d('2026-02-02'), d('2026-03-13'), false, 'D-01SS+5d', 'No', 6],
  ['D-04', 'Power board schematic', 'Design', 'Casey', d('2026-02-16'), d('2026-03-27'), false, 'D-02', null, 4],
  ['D-05', 'Firmware interface spec', 'Design', 'Devon', d('2026-02-16'), d('2026-03-20'), false, 'D-02FS', 'No', 9],
  ['D-06', 'Design freeze', 'Design', 'Avery', d('2026-03-30'), d('2026-03-30'), true, 'D-03, D-04, D-05', 'Yes', 0],
  ['D-07', 'Thermal analysis', 'Design', 'Blake', d('2026-03-02'), d('2026-03-27'), false, 'D-03SS+10d', 'No', 3],
  ['B-01', 'Long-lead parts order', 'Build', 'Emery', d('2026-03-30'), d('2026-04-10'), false, 'D-06', 'Yes', 0],
  ['B-02', 'PCB fabrication', 'Build', 'Casey', d('2026-04-13'), d('2026-05-08'), false, 'B-01FS', 'Yes', 0],
  ['B-03', 'Enclosure tooling', 'Build', 'Blake', d('2026-04-06'), d('2026-05-22'), false, 'D-06FS+1d', 'No', 8],
  ['B-04', 'Board assembly', 'Build', 'Emery', d('2026-05-11'), d('2026-05-29'), false, 'B-02', 'Yes', 0],
  ['B-05', 'Firmware alpha', 'Build', 'Devon', d('2026-04-06'), d('2026-06-05'), false, 'D-05; D-06', null, 5],
  ['B-06', 'Cable harness build', 'Build', 'Frankie', d('2026-05-04'), d('2026-05-29'), false, 'B-01SS+14d', 'No', 7],
  ['B-07', 'Unit integration', 'Build', 'Emery', d('2026-06-01'), d('2026-06-19'), false, 'B-04, B-03FF-2', 'Yes', 0],
  ['B-08', 'Build complete review', 'Build', 'Emery', d('2026-06-19'), d('2026-06-19'), false, 'B-07, B-05', 'Yes', 0],
  ['V-01', 'Test plan approval', 'Validation', 'Gray', d('2026-04-20'), d('2026-05-15'), false, 'D-06', 'No', 12],
  ['V-02', 'Test fixture build', 'Validation', 'Gray', d('2026-05-18'), d('2026-06-12'), false, 'V-01, B-02SF+3 days', 'No', 5],
  ['V-03', 'Functional test', 'Validation', 'Harper', d('2026-06-22'), d('2026-07-10'), false, 'B-08, V-02', 'Yes', 0],
  ['V-04', 'Environmental test', 'Validation', 'Harper', d('2026-07-13'), d('2026-08-07'), false, 'V-03FS', 'Yes', 0],
  ['V-05', 'EMC pre-scan', 'Validation', 'Gray', d('2026-07-13'), d('2026-07-31'), false, 'V-03SS', null, 10],
  ['V-06', 'Reliability soak', 'Validation', 'Harper', d('2026-07-13'), d('2026-09-04'), false, 'V-03', 'No', 6],
  ['V-07', 'Defect triage and fixes', 'Validation', 'Devon', d('2026-08-10'), d('2026-08-28'), false, 'V-04, V-05', 'Yes', 0],
  ['V-08', 'Regression test', 'Validation', 'Harper', d('2026-08-31'), d('2026-09-11'), false, 'V-07', 'Yes', 0],
  ['V-09', 'Validation report', 'Validation', 'Gray', d('2026-09-14'), d('2026-09-25'), false, 'V-08, V-06', null, 0],
  ['V-10', 'Validation exit', 'Validation', 'Gray', d('2026-09-25'), d('2026-09-25'), true, 'V-09', 'Yes', 0],
]

/** Write the synthetic sample program to `<dir>/evt-program.xlsx`. */
export async function writeSampleFixtures(dir: string): Promise<string> {
  await mkdir(dir, { recursive: true })
  const path = join(dir, 'evt-program.xlsx')
  await writeFile(path, await makeXlsx(HEADER, SAMPLE))
  return path
}
