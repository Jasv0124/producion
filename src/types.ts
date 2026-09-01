export type Operation = {
  id: string
  planta: string
  regional: string
  categoria: string
  empaque: string
  claseEnvase: string
  mes: string
  fecha: string
  material: string
  descripcion: string
  hectolitros: number
  unidades: number
  orden: string
  recurso: string
}

export type Commercial = {
  id: string
  regional: string
  mes: string
  material: string
  descripcion: string
  produccion: number
  inventario: number
  pronostico: number
  sop: number
  venta: number
  variacionFrc: number
  variacionSop: number
  variacionVenta: number
  diasCobertura: number
  local?: boolean
}

export type PlantMonth = {
  planta: string
  categoria: string
  values: Record<string, number>
  total: number
}

export type CategoryPlant = {
  planta: string
  aguas: number
  cervezas: number
  maltas: number
  exportacion: number
  total: number
}

export type ProductMonth = {
  planta: string
  descripcion: string
  empaque: string
  values: Record<string, number>
  total: number
}

export type DailyProduction = {
  fecha: string
  valor: number
}

export type InventoryPoint = {
  regional: string
  mes: string
  material: string
  valor: number
}

export type SheetInfo = {
  name: string
  rows: number
  role: string
}

export type WorkbookData = {
  operations: Operation[]
  commercial: Commercial[]
  plantMonth: PlantMonth[]
  categoriesByPlant: CategoryPlant[]
  productMonth: ProductMonth[]
  productMonthMaterial: string
  daily: DailyProduction[]
  dailyDescription: string
  inventory: InventoryPoint[]
  sheets: SheetInfo[]
}

export type ManualRecord = {
  id: string
  planta: string
  regional: string
  categoria: string
  empaque: string
  mes: string
  material: string
  descripcion: string
  produccion: number
  unidades: number
  inventario: number
  pronostico: number
  sop: number
  venta: number
  local: true
}
