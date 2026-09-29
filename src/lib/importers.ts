import * as XLSX from 'xlsx'
import Papa from 'papaparse'
import type { Person } from '../types'

const norm = (s: unknown) =>
  String(s ?? '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[ºª]/g, '')
    .replace(/[^a-z0-9]+/g, ' ')

function findValue(row: Record<string, unknown>, aliases: string[]): string {
  const entries = Object.entries(row)
  const normalizedAliases = aliases.map(a => norm(a))

  const exact = entries.find(([key]) => normalizedAliases.includes(norm(key)))
  if (exact) return String(exact[1] ?? '').trim()

  const partial = entries.find(([key]) => {
    const header = norm(key)
    return normalizedAliases.some(alias => header.includes(alias) || alias.includes(header))
  })
  return partial ? String(partial[1] ?? '').trim() : ''
}

function mapRow(row: Record<string, unknown>): Person | null {
  const name = findValue(row, ['nome', 'nome completo', 'name'])
  if (!name) return null

  return {
    id: crypto.randomUUID(),
    name,
    registration: findValue(row, [
      'matricula',
      'matrícula',
      'mat',
      'registro',
      'registro funcional',
      'registration',
    ]),
    role: findValue(row, [
      'funcao',
      'função',
      'cargo',
      'funcao cargo',
      'role',
    ]),
    phone: findValue(row, ['telefone', 'celular', 'phone']),
    company: findValue(row, ['empresa', 'company']),
    manager: findValue(row, ['gestor', 'supervisor', 'manager']),
    trainingDate: findValue(row, ['treinamento', 'data treinamento', 'training']),
  }
}

export async function importSpreadsheet(file: File): Promise<Person[]> {
  const ext = file.name.toLowerCase().split('.').pop()

  if (ext === 'csv') {
    const parsed = Papa.parse<Record<string, unknown>>(await file.text(), {
      header: true,
      skipEmptyLines: true,
      transformHeader: h => h.trim(),
    })
    return parsed.data.map(mapRow).filter(Boolean) as Person[]
  }

  const workbook = XLSX.read(await file.arrayBuffer(), {
    type: 'array',
    cellDates: true,
  })

  const sheet = workbook.Sheets[workbook.SheetNames[0]]
  if (!sheet) throw new Error('Nenhuma planilha encontrada.')

  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {
    defval: '',
    raw: false,
  })

  return rows.map(mapRow).filter(Boolean) as Person[]
}