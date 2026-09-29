import { Document, Packer, Paragraph, Table, TableCell, TableRow, WidthType, TextRun, ShadingType, VerticalAlign, AlignmentType } from 'docx'
import type { BadgeTemplate, Person } from '../types'

export async function exportDocx(people: Person[], t: BadgeTemplate) {
  const rows: TableRow[] = []
  for (let i = 0; i < people.length; i += 4) {
    const cells: TableCell[] = []
    for (let j = 0; j < 4; j++) {
      const p = people[i + j]
      if (!p) {
        cells.push(new TableCell({ children: [new Paragraph('')] }))
        continue
      }
      const field = (l: string, v: string) => new Paragraph({
        children: [new TextRun({ text: l, bold: true, size: 13 }), new TextRun({ text: ' ' + (v || '—'), size: 13 })],
      })
      cells.push(new TableCell({
        verticalAlign: VerticalAlign.CENTER,
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '⚠', bold: true, size: 32 })] }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            shading: { type: ShadingType.CLEAR, fill: t.primary.slice(1) },
            children: [new TextRun({ text: t.title + '\n' + t.subtitle, bold: true, color: 'FFFFFF', size: 16 })],
          }),
          field('NOME:', p.name),
          field('MATRÍCULA:', p.registration),
          field('FUNÇÃO:', p.role),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: t.logoText, bold: true, color: t.primary.slice(1), size: 17 })],
          }),
        ],
      }))
    }
    rows.push(new TableRow({ children: cells }))
  }

  const d = new Document({
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 300, bottom: 300, left: 300, right: 300 } } },
      children: [new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows })],
    }],
  })

  const blob = await Packer.toBlob(d)
  const u = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = u
  a.download = 'crachas-cracha-fast.docx'
  a.click()
  URL.revokeObjectURL(u)
}