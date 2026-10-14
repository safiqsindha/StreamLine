import { deflateSync } from 'node:zlib'

/** One-page PDF 1.4 holding a single full-page RGB image. `rgb` is width*height*3 raw bytes. */
export function imageToPdf(rgb: Buffer, width: number, height: number, pageWidthPt: number, pageHeightPt: number): Buffer {
  if (rgb.length !== width * height * 3) throw new Error('Pixel data does not match the image size.')
  const pw = Number(pageWidthPt.toFixed(2))
  const ph = Number(pageHeightPt.toFixed(2))
  const pixels = deflateSync(rgb)
  const content = Buffer.from(`q\n${pw} 0 0 ${ph} 0 0 cm\n/Im0 Do\nQ\n`, 'latin1')

  const stream = (dict: string, data: Buffer) => [Buffer.from(`<< ${dict} /Length ${data.length} >>\nstream\n`, 'latin1'), data, Buffer.from('\nendstream', 'latin1')]
  const objects: Buffer[][] = [
    [Buffer.from('<< /Type /Catalog /Pages 2 0 R >>', 'latin1')],
    [Buffer.from('<< /Type /Pages /Kids [3 0 R] /Count 1 >>', 'latin1')],
    [Buffer.from(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pw} ${ph}] /Resources << /XObject << /Im0 5 0 R >> >> /Contents 4 0 R >>`, 'latin1')],
    stream('', content),
    stream(`/Type /XObject /Subtype /Image /Width ${width} /Height ${height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode`, pixels),
  ]

  const parts: Buffer[] = [Buffer.from('%PDF-1.4\n%\xE2\xE3\xCF\xD3\n', 'latin1')]
  const offsets: number[] = []
  let pos = parts[0].length
  objects.forEach((body, i) => {
    offsets.push(pos)
    const chunk = Buffer.concat([Buffer.from(`${i + 1} 0 obj\n`, 'latin1'), ...body, Buffer.from('\nendobj\n', 'latin1')])
    parts.push(chunk)
    pos += chunk.length
  })

  const xref = [`xref\n0 ${objects.length + 1}\n`, '0000000000 65535 f \n', ...offsets.map((o) => `${String(o).padStart(10, '0')} 00000 n \n`)].join('')
  parts.push(Buffer.from(`${xref}trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${pos}\n%%EOF\n`, 'latin1'))
  return Buffer.concat(parts)
}
