import * as XLSX from 'xlsx'
import type { CategoryPlant, Commercial, DailyProduction, InventoryPoint, Operation, PlantMonth, ProductMonth, SheetInfo, WorkbookData } from './types'

export const monthOrder = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']

export const normalize = (value: unknown) => String(value ?? '').trim().toLocaleLowerCase('es').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
export const cleanNumber = (value: unknown) => Number(String(value ?? 0).replace(/[^0-9.-]/g, '')) || 0

function getCell(row: Record<string, unknown>, name: string) {
  const key = Object.keys(row).find(item => normalize(item) === normalize(name))
  return key ? row[key] : ''
}

function asText(value: unknown) {
  return String(value ?? '').trim()
}

function dateLabel(value: unknown) {
  if (value instanceof Date) return value.toISOString().slice(0, 10)
  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value)
    if (parsed) return `${parsed.y}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`
  }
  return asText(value)
}

function objects(workbook: XLSX.WorkBook, name: string) {
  return XLSX.utils.sheet_to_json<Record<string, unknown>>(workbook.Sheets[name], { defval: '', raw: true })
}

function matrix(workbook: XLSX.WorkBook, name: string) {
  return XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[name], { header: 1, defval: null, raw: true })
}

function parseOperations(workbook: XLSX.WorkBook): Operation[] {
  return objects(workbook, 'Data').map((row, index) => ({
    id: `op-${index}-${asText(getCell(row, 'Orden'))}`,
    planta: asText(getCell(row, 'Planta')),
    regional: asText(getCell(row, 'Regional')),
    categoria: asText(getCell(row, 'Catergoría')),
    empaque: asText(getCell(row, 'Empaque')),
    claseEnvase: asText(getCell(row, 'Clase Envase')),
    mes: normalize(getCell(row, 'Mes')),
    fecha: dateLabel(getCell(row, 'Fe.Cont')),
    material: asText(getCell(row, 'Material')),
    descripcion: asText(getCell(row, 'Descripción Material')),
    hectolitros: cleanNumber(getCell(row, 'Hectolitros')),
    unidades: cleanNumber(getCell(row, 'Unidades')),
    orden: asText(getCell(row, 'Orden')),
    recurso: asText(getCell(row, 'Recurso')),
  })).filter(row => row.planta && row.material)
}

function parseCommercial(workbook: XLSX.WorkBook): Commercial[] {
  return objects(workbook, 'Producción').map((row, index) => ({
    id: `commercial-${index}-${asText(getCell(row, 'Llave'))}`,
    regional: asText(getCell(row, 'Regional')),
    mes: normalize(getCell(row, 'Mes')),
    material: asText(getCell(row, 'Material')),
    descripcion: asText(getCell(row, 'Descripción Material')),
    produccion: cleanNumber(getCell(row, 'Producción')),
    inventario: cleanNumber(getCell(row, 'Inv Inicial')),
    pronostico: cleanNumber(getCell(row, 'Pronóstico')),
    sop: cleanNumber(getCell(row, 'S&OP')),
    venta: cleanNumber(getCell(row, 'Venta')),
    variacionFrc: cleanNumber(getCell(row, 'Variación FRC')),
    variacionSop: cleanNumber(getCell(row, 'Variación S&OP')),
    variacionVenta: cleanNumber(getCell(row, 'Variación Venta')),
    diasCobertura: cleanNumber(getCell(row, 'DC')),
  })).filter(row => row.material && row.regional)
}

function parsePlantMonth(workbook: XLSX.WorkBook): PlantMonth[] {
  const rows = matrix(workbook, 'TD Plantas Mes')
  let currentPlant = ''
  return rows.slice(5).map(row => {
    const rawPlant = asText(row[0])
    if (rawPlant && !normalize(rawPlant).startsWith('total')) currentPlant = rawPlant
    const categoria = asText(row[1])
    if (!categoria || normalize(rawPlant).startsWith('total')) return null
    const values = Object.fromEntries(monthOrder.slice(0, 8).map((month, index) => [month, cleanNumber(row[index + 2])]))
    return { planta: currentPlant, categoria, values, total: cleanNumber(row[10]) }
  }).filter((row): row is PlantMonth => row !== null)
}

function parseCategories(workbook: XLSX.WorkBook): CategoryPlant[] {
  return matrix(workbook, 'Prod Categorías').slice(5).map(row => ({
    planta: asText(row[0]), aguas: cleanNumber(row[1]), cervezas: cleanNumber(row[2]), maltas: cleanNumber(row[3]), exportacion: cleanNumber(row[4]), total: cleanNumber(row[5]),
  })).filter(row => row.planta && normalize(row.planta) !== 'total general')
}

function parseProductMonth(workbook: XLSX.WorkBook) {
  const rows = matrix(workbook, 'Prod Mes')
  const material = asText(rows[0]?.[1])
  const data: ProductMonth[] = rows.slice(4).map(row => {
    const planta = asText(row[0])
    if (!planta || normalize(planta) === 'total general') return null
    return {
      planta, descripcion: asText(row[1]), empaque: asText(row[2]),
      values: Object.fromEntries(monthOrder.slice(0, 8).map((month, index) => [month, cleanNumber(row[index + 3])])),
      total: cleanNumber(row[11]),
    }
  }).filter((row): row is ProductMonth => row !== null)
  return { material, data }
}

function parseDaily(workbook: XLSX.WorkBook) {
  const rows = matrix(workbook, 'Prod Día')
  const header = rows[4] || []
  const values = rows[5] || []
  const daily: DailyProduction[] = header.slice(4).map((date, index) => ({ fecha: dateLabel(date), valor: cleanNumber(values[index + 4]) })).filter(item => item.fecha && item.valor)
  return { daily, description: asText(values[3]) }
}

function parseInventory(workbook: XLSX.WorkBook): InventoryPoint[] {
  const rows = matrix(workbook, 'Inv Inic')
  return rows.slice(1).map(row => ({ regional: asText(row[1]), mes: normalize(row[2]), material: asText(row[3]), valor: cleanNumber(row[4]) })).filter(row => row.regional && row.mes && row.material)
}

const roles: Record<string, string> = {
  Data: 'Detalle de órdenes y ejecución',
  'Prod Mes': 'Seguimiento mensual por material',
  'Prod Día': 'Detalle diario del material seleccionado',
  'Prod Categorías': 'Distribución por planta y categoría',
  'TD Plantas Mes': 'Producción mensual por planta',
  Producción: 'Producción, pronóstico, S&OP y venta',
  'Inv Inic': 'Inventario inicial y bases comerciales',
}

export function readWorkbook(buffer: ArrayBuffer): WorkbookData {
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true })
  const required = Object.keys(roles)
  const missing = required.filter(name => !workbook.Sheets[name])
  if (missing.length) throw new Error(`Faltan hojas requeridas: ${missing.join(', ')}`)
  const productMonth = parseProductMonth(workbook)
  const daily = parseDaily(workbook)
  const sheets: SheetInfo[] = workbook.SheetNames.map(name => ({
    name,
    rows: matrix(workbook, name).filter(row => row.some(value => value !== null && value !== '')).length,
    role: roles[name] || 'Información complementaria',
  }))
  return {
    operations: parseOperations(workbook),
    commercial: parseCommercial(workbook),
    plantMonth: parsePlantMonth(workbook),
    categoriesByPlant: parseCategories(workbook),
    productMonth: productMonth.data,
    productMonthMaterial: productMonth.material,
    daily: daily.daily,
    dailyDescription: daily.description,
    inventory: parseInventory(workbook),
    sheets,
  }
}
