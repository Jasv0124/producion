import { ChangeEvent, FormEvent, useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import {
  Area, AreaChart, Bar, BarChart, CartesianGrid, Cell, Legend, Line, LineChart,
  Pie, PieChart, ResponsiveContainer, Tooltip, XAxis, YAxis,
} from 'recharts'
import { cleanNumber, monthOrder, normalize, readWorkbook } from './excel'
import type { Commercial, ManualRecord, Operation, WorkbookData } from './types'

type View = 'resumen' | 'operacion' | 'comercial' | 'fuentes'
type ManualForm = Omit<ManualRecord, 'id' | 'local' | 'produccion' | 'unidades' | 'inventario' | 'pronostico' | 'sop' | 'venta'> & {
  produccion: string; unidades: string; inventario: string; pronostico: string; sop: string; venta: string
}

const palette = ['#25c99a', '#f1b543', '#5787f6', '#a57bea', '#ed786b', '#35afbd']
const numberFormat = new Intl.NumberFormat('es-CO', { maximumFractionDigits: 0 })
const compactFormat = new Intl.NumberFormat('es-CO', { notation: 'compact', maximumFractionDigits: 1 })
const storageKey = 'produccion-registros-completos-v2'
const emptyForm: ManualForm = { planta: '', regional: '', categoria: 'Cervezas', empaque: 'Botellas', mes: 'agosto', material: '', descripcion: '', produccion: '', unidades: '', inventario: '', pronostico: '', sop: '', venta: '' }
const tooltipStyle = { background: 'rgba(10, 38, 32, .96)', border: '1px solid rgba(70, 220, 176, .22)', borderRadius: 10, color: '#fff', fontSize: 11 }

function unique(values: string[]) { return [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es')) }
function titleCase(value: string) { return value ? value[0].toUpperCase() + value.slice(1) : value }

function MetricCard({ icon, title, value, suffix, note, tone, delay }: { icon: string; title: string; value: number; suffix: string; note: string; tone: string; delay: number }) {
  return <article className="metric-card reveal" style={{ animationDelay: `${delay}ms` }}><div className={`metric-icon ${tone}`}>{icon}</div><div className="metric-copy"><span>{title}</span><strong>{numberFormat.format(value)} <small>{suffix}</small></strong><p>{note}</p></div><div className="metric-scan" /></article>
}

function PanelTitle({ eyebrow, title, description, tag }: { eyebrow: string; title: string; description: string; tag?: string }) {
  return <div className="panel-head"><div><p className="eyebrow">{eyebrow}</p><h2>{title}</h2><p>{description}</p></div>{tag && <span className="unit-pill">{tag}</span>}</div>
}

export default function App() {
  const [data, setData] = useState<WorkbookData | null>(null)
  const [manualRecords, setManualRecords] = useState<ManualRecord[]>(() => {
    try { return JSON.parse(localStorage.getItem(storageKey) || '[]') as ManualRecord[] } catch { return [] }
  })
  const [status, setStatus] = useState('Sincronizando las siete hojas del libro…')
  const [view, setView] = useState<View>('resumen')
  const [planta, setPlanta] = useState('Todas')
  const [regional, setRegional] = useState('Todas')
  const [categoria, setCategoria] = useState('Todas')
  const [empaque, setEmpaque] = useState('Todos')
  const [month, setMonth] = useState('Todos')
  const [search, setSearch] = useState('')
  const [form, setForm] = useState<ManualForm>(emptyForm)
  const [showForm, setShowForm] = useState(false)

  useEffect(() => {
    const load = async () => {
      try {
        const response = await fetch(new URL('../6.1 Ejecución Producción.xlsx', import.meta.url))
        if (!response.ok) throw new Error('No se pudo abrir el archivo base.')
        const workbookData = readWorkbook(await response.arrayBuffer())
        setData(workbookData)
        setStatus(`${workbookData.sheets.length} hojas · ${numberFormat.format(workbookData.operations.length)} movimientos sincronizados`)
      } catch (error) {
        setStatus(error instanceof Error ? error.message : 'Carga el Excel para iniciar')
      }
    }
    void load()
  }, [])

  const manualOperations = useMemo<Operation[]>(() => manualRecords.map(record => ({
    id: record.id, planta: record.planta, regional: record.regional, categoria: record.categoria, empaque: record.empaque,
    claseEnvase: '', mes: record.mes, fecha: '', material: record.material, descripcion: record.descripcion,
    hectolitros: record.produccion, unidades: record.unidades, orden: 'Registro web', recurso: 'Manual',
  })), [manualRecords])
  const operations = useMemo(() => [...manualOperations, ...(data?.operations || [])], [manualOperations, data])
  const commercials = useMemo<Commercial[]>(() => [
    ...manualRecords.map(record => ({
      id: record.id, regional: record.regional, mes: record.mes, material: record.material, descripcion: record.descripcion,
      produccion: record.produccion, inventario: record.inventario, pronostico: record.pronostico, sop: record.sop, venta: record.venta,
      variacionFrc: record.produccion + record.inventario - record.pronostico, variacionSop: record.produccion + record.inventario - record.sop,
      variacionVenta: record.produccion + record.inventario - record.venta, diasCobertura: 0, local: true,
    })), ...(data?.commercial || []),
  ], [manualRecords, data])

  const plants = useMemo(() => unique(operations.map(row => row.planta)), [operations])
  const regions = useMemo(() => unique(operations.map(row => row.regional)), [operations])
  const categories = useMemo(() => unique(operations.map(row => row.categoria)), [operations])
  const packages = useMemo(() => unique(operations.map(row => row.empaque)), [operations])
  const availableMonths = useMemo(() => monthOrder.filter(item => operations.some(row => row.mes === item)), [operations])

  const filteredOperations = useMemo(() => operations.filter(row => {
    const term = normalize(search)
    return (planta === 'Todas' || row.planta === planta) && (regional === 'Todas' || row.regional === regional) &&
      (categoria === 'Todas' || row.categoria === categoria) && (empaque === 'Todos' || row.empaque === empaque) &&
      (month === 'Todos' || row.mes === month) && (!term || normalize(`${row.material} ${row.descripcion} ${row.orden}`).includes(term))
  }), [operations, planta, regional, categoria, empaque, month, search])
  const filteredCommercial = useMemo(() => commercials.filter(row => {
    const term = normalize(search)
    return (regional === 'Todas' || row.regional === regional) && (month === 'Todos' || row.mes === month) && (!term || normalize(`${row.material} ${row.descripcion}`).includes(term))
  }), [commercials, regional, month, search])

  const totals = useMemo(() => filteredOperations.reduce((sum, row) => ({ hl: sum.hl + row.hectolitros, units: sum.units + row.unidades }), { hl: 0, units: 0 }), [filteredOperations])
  const orderCount = useMemo(() => new Set(filteredOperations.map(row => row.orden).filter(Boolean)).size, [filteredOperations])
  const materialCount = useMemo(() => new Set(filteredOperations.map(row => row.material)).size, [filteredOperations])

  const monthlyData = useMemo(() => monthOrder.map(item => {
    const rows = filteredOperations.filter(row => row.mes === item)
    return {
      month: titleCase(item.slice(0, 3)), total: rows.reduce((sum, row) => sum + row.hectolitros, 0),
      Cervezas: rows.filter(row => normalize(row.categoria) === 'cervezas').reduce((sum, row) => sum + row.hectolitros, 0),
      Maltas: rows.filter(row => normalize(row.categoria) === 'maltas').reduce((sum, row) => sum + row.hectolitros, 0),
      Aguas: rows.filter(row => normalize(row.categoria) === 'aguas').reduce((sum, row) => sum + row.hectolitros, 0),
    }
  }).filter(row => row.total > 0), [filteredOperations])

  const plantCategoryData = useMemo(() => {
    if (!data) return []
    return data.categoriesByPlant.filter(row => planta === 'Todas' || row.planta === planta)
  }, [data, planta])
  const packageData = useMemo(() => packages.map(name => ({ name, value: filteredOperations.filter(row => row.empaque === name).reduce((sum, row) => sum + row.hectolitros, 0) })).filter(row => row.value > 0), [filteredOperations, packages])
  const topMaterials = useMemo(() => {
    const grouped = filteredOperations.reduce<Record<string, { material: string; descripcion: string; produccion: number }>>((acc, row) => {
      if (!acc[row.material]) acc[row.material] = { material: row.material, descripcion: row.descripcion, produccion: 0 }
      acc[row.material].produccion += row.hectolitros; return acc
    }, {})
    return Object.values(grouped).sort((a, b) => b.produccion - a.produccion).slice(0, 10)
  }, [filteredOperations])
  const productMonthData = useMemo(() => monthOrder.slice(0, 8).map(item => {
    const point: Record<string, string | number> = { month: titleCase(item.slice(0, 3)) }
    data?.productMonth.forEach(row => { point[row.planta] = row.values[item] || 0 })
    return point
  }), [data])
  const commercialChart = useMemo(() => {
    const grouped = filteredCommercial.reduce<Record<string, { material: string; produccion: number; venta: number; pronostico: number }>>((acc, row) => {
      if (!acc[row.material]) acc[row.material] = { material: row.material, produccion: 0, venta: 0, pronostico: 0 }
      acc[row.material].produccion += row.produccion; acc[row.material].venta += row.venta; acc[row.material].pronostico += row.pronostico; return acc
    }, {})
    return Object.values(grouped).sort((a, b) => b.produccion - a.produccion).slice(0, 9)
  }, [filteredCommercial])
  const inventoryData = useMemo(() => monthOrder.map(item => ({
    month: titleCase(item.slice(0, 3)),
    inventario: (data?.inventory || []).filter(row => row.mes === item && (regional === 'Todas' || row.regional === regional)).reduce((sum, row) => sum + row.valor, 0),
  })).filter(row => row.inventario > 0), [data, regional])
  const commercialTotals = useMemo(() => filteredCommercial.reduce((sum, row) => ({ produccion: sum.produccion + row.produccion, venta: sum.venta + row.venta, pronostico: sum.pronostico + row.pronostico }), { produccion: 0, venta: 0, pronostico: 0 }), [filteredCommercial])

  const upload = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]; if (!file) return
    try { const loaded = readWorkbook(await file.arrayBuffer()); setData(loaded); setStatus(`${loaded.sheets.length} hojas sincronizadas desde ${file.name}`) }
    catch (error) { setStatus(error instanceof Error ? error.message : 'No fue posible leer el archivo') }
    event.target.value = ''
  }
  const addRecord = (event: FormEvent) => {
    event.preventDefault()
    const record: ManualRecord = { id: crypto.randomUUID(), planta: form.planta.trim(), regional: form.regional.trim(), categoria: form.categoria, empaque: form.empaque, mes: normalize(form.mes), material: form.material.trim(), descripcion: form.descripcion.trim(), produccion: cleanNumber(form.produccion), unidades: cleanNumber(form.unidades), inventario: cleanNumber(form.inventario), pronostico: cleanNumber(form.pronostico), sop: cleanNumber(form.sop), venta: cleanNumber(form.venta), local: true }
    const next = [record, ...manualRecords]; setManualRecords(next); localStorage.setItem(storageKey, JSON.stringify(next)); setForm(emptyForm); setShowForm(false); setStatus('Nueva producción incorporada al tablero')
  }
  const exportRecords = () => {
    const rows = operations.map(row => ({ Planta: row.planta, Regional: row.regional, Categoría: row.categoria, Empaque: row.empaque, Mes: row.mes, Material: row.material, Descripción: row.descripcion, Hectolitros: row.hectolitros, Unidades: row.unidades, Orden: row.orden }))
    const workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), 'Producción consolidada'); XLSX.writeFile(workbook, 'produccion-consolidada.xlsx')
  }
  const setField = (key: keyof ManualForm, value: string) => setForm(current => ({ ...current, [key]: value }))

  const navItems: Array<{ id: View; icon: string; label: string }> = [
    { id: 'resumen', icon: '⌁', label: 'Resumen' }, { id: 'operacion', icon: '↗', label: 'Operación' },
    { id: 'comercial', icon: '◎', label: 'Comercial' }, { id: 'fuentes', icon: '▤', label: 'Fuentes' },
  ]

  return <div className="app-shell">
    <aside className="sidebar"><div className="brand"><span>PX</span><div><strong>Producción</strong><small>Control integral</small></div></div><nav aria-label="Secciones del tablero">{navItems.map(item => <button key={item.id} className={view === item.id ? 'active' : ''} onClick={() => setView(item.id)}><span>{item.icon}</span>{item.label}</button>)}</nav><div className="data-source"><span className="live-dot" /><div><strong>{data ? 'Libro conectado' : 'Conectando'}</strong><small>{data ? `${data.sheets.length} hojas activas` : 'Excel de producción'}</small></div></div></aside>

    <main>
      <header className="topbar reveal"><div><p className="eyebrow">CENTRO DE OPERACIONES · 2026</p><h1>{view === 'resumen' ? 'Ejecución de producción' : view === 'operacion' ? 'Detalle operativo' : view === 'comercial' ? 'Planeación comercial' : 'Mapa de información'}</h1><p className="subtitle">{view === 'resumen' ? 'Una lectura consolidada de plantas, categorías y volúmenes.' : view === 'operacion' ? 'Órdenes, materiales y comportamiento diario.' : view === 'comercial' ? 'Producción frente a inventario, pronóstico, S&OP y venta.' : 'Las siete hojas del Excel organizadas por función.'}</p></div><div className="header-actions"><label className="button secondary">↑ Cargar Excel<input type="file" accept=".xlsx,.xls" onChange={upload} /></label>{data && <button className="button secondary" onClick={exportRecords}>↓ Exportar</button>}<button className="button primary" onClick={() => setShowForm(true)}>＋ Nueva producción</button></div></header>

      <section className="filterbar reveal"><div className="status"><span className="live-dot" />{status}</div><div className="filters"><label className="search"><span>⌕</span><input aria-label="Buscar" placeholder="Material, producto u orden" value={search} onChange={event => setSearch(event.target.value)} /></label><Filter label="Planta" value={planta} options={['Todas', ...plants]} onChange={setPlanta} /><Filter label="Regional" value={regional} options={['Todas', ...regions]} onChange={setRegional} /><Filter label="Categoría" value={categoria} options={['Todas', ...categories]} onChange={setCategoria} /><Filter label="Empaque" value={empaque} options={['Todos', ...packages]} onChange={setEmpaque} /><Filter label="Mes" value={month} options={['Todos', ...availableMonths]} onChange={setMonth} capitalize /></div></section>

      {showForm && <EntryModal form={form} plants={plants} regions={regions} categories={categories} packages={packages} onChange={setField} onClose={() => setShowForm(false)} onSubmit={addRecord} />}

      {!data ? <section className="empty-state"><div className="loader-ring" /><h2>Organizando la información</h2><p>Estamos leyendo las hojas, relaciones y tablas dinámicas del libro.</p></section> : <>
        {view === 'resumen' && <>
          <section className="metrics"><MetricCard icon="HL" title="Producción acumulada" value={totals.hl} suffix="HL" note="Base detallada de ejecución" tone="green" delay={50} /><MetricCard icon="UN" title="Unidades producidas" value={totals.units} suffix="UN" note={`${numberFormat.format(orderCount)} órdenes únicas`} tone="blue" delay={100} /><MetricCard icon="MT" title="Materiales activos" value={materialCount} suffix="SKU" note={`${new Set(filteredOperations.map(row => row.planta)).size} plantas en la selección`} tone="violet" delay={150} /><MetricCard icon="RG" title="Registros analizados" value={filteredOperations.length} suffix="FILAS" note="Incluye altas desde la aplicación" tone="amber" delay={200} /></section>
          <section className="dashboard-grid overview-grid"><article className="panel reveal"><PanelTitle eyebrow="TENDENCIA GENERAL" title="Producción mensual por categoría" description="Cálculo filtrable desde la hoja Data; coincide con la tabla de plantas por mes." tag="Hectolitros" /><div className="chart-wrap tall"><ResponsiveContainer width="100%" height="100%"><AreaChart data={monthlyData} margin={{ top: 15, right: 10, left: 0, bottom: 0 }}><defs><linearGradient id="beer" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#25c99a" stopOpacity={.45}/><stop offset="100%" stopColor="#25c99a" stopOpacity={.03}/></linearGradient></defs><CartesianGrid strokeDasharray="3 5" vertical={false} stroke="#e3ebe7" /><XAxis dataKey="month" axisLine={false} tickLine={false} tick={{ fill: '#71807d', fontSize: 11 }} /><YAxis axisLine={false} tickLine={false} tick={{ fill: '#879590', fontSize: 10 }} tickFormatter={value => compactFormat.format(value)} width={48} /><Tooltip contentStyle={tooltipStyle} formatter={value => `${numberFormat.format(Number(value) || 0)} HL`} /><Legend iconType="circle" iconSize={7} wrapperStyle={{ fontSize: 11 }} /><Area type="monotone" dataKey="Cervezas" stackId="1" stroke="#20ae84" fill="url(#beer)" strokeWidth={2.5} animationDuration={1000} /><Area type="monotone" dataKey="Maltas" stackId="1" stroke="#f1b543" fill="#f7d98d" fillOpacity={.5} animationDuration={1200} /><Area type="monotone" dataKey="Aguas" stackId="1" stroke="#5787f6" fill="#b8ccff" fillOpacity={.45} animationDuration={1400} /></AreaChart></ResponsiveContainer></div></article>
          <article className="panel reveal" style={{ animationDelay: '100ms' }}><PanelTitle eyebrow="MEZCLA DE PRODUCCIÓN" title="Distribución por empaque" description="Participación sobre la selección activa." /><div className="donut-wrap"><ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={packageData} dataKey="value" nameKey="name" innerRadius="66%" outerRadius="88%" paddingAngle={3} cornerRadius={6} animationDuration={1200}>{packageData.map((item, index) => <Cell key={item.name} fill={palette[index % palette.length]} />)}</Pie><Tooltip contentStyle={tooltipStyle} formatter={value => `${numberFormat.format(Number(value) || 0)} HL`} /></PieChart></ResponsiveContainer><div className="donut-center"><strong>{compactFormat.format(totals.hl)}</strong><span>HL seleccionados</span></div></div><div className="region-list">{packageData.map((item, index) => <div key={item.name}><span><i style={{ background: palette[index % palette.length] }} />{item.name}</span><strong>{totals.hl ? Math.round(item.value / totals.hl * 100) : 0}%</strong></div>)}</div></article></section>
          <article className="panel reveal section-panel"><PanelTitle eyebrow="PROD CATEGORÍAS" title="Producción por planta y categoría" description="Valores consolidados tal como aparecen en la hoja dinámica del Excel." tag="Enero – agosto" /><div className="chart-wrap plant-chart"><ResponsiveContainer width="100%" height="100%"><BarChart data={plantCategoryData} layout="vertical" margin={{ top: 10, right: 22, left: 18, bottom: 0 }}><CartesianGrid strokeDasharray="3 5" horizontal={false} stroke="#e6ece9" /><XAxis type="number" axisLine={false} tickLine={false} tickFormatter={value => compactFormat.format(value)} tick={{ fill: '#879590', fontSize: 10 }} /><YAxis type="category" dataKey="planta" axisLine={false} tickLine={false} width={90} tick={{ fill: '#526963', fontSize: 11 }} /><Tooltip contentStyle={tooltipStyle} formatter={value => `${numberFormat.format(Number(value) || 0)} HL`} /><Legend iconType="circle" iconSize={7} wrapperStyle={{ fontSize: 11 }} /><Bar dataKey="cervezas" name="Cervezas" stackId="a" fill="#25c99a" radius={[0, 5, 5, 0]} /><Bar dataKey="maltas" name="Maltas" stackId="a" fill="#f1b543" /><Bar dataKey="aguas" name="Aguas" stackId="a" fill="#5787f6" /><Bar dataKey="exportacion" name="Exportación" stackId="a" fill="#a57bea" /></BarChart></ResponsiveContainer></div></article>
          <PlantMonthTable rows={data.plantMonth} exportTotal={data.categoriesByPlant.reduce((sum, row) => sum + row.exportacion, 0)} />
        </>}

        {view === 'operacion' && <>
          <section className="dashboard-grid"><article className="panel reveal"><PanelTitle eyebrow="PROD DÍA" title={`Ejecución diaria · ${data.dailyDescription || 'Material seleccionado'}`} description="Serie conservada desde la vista diaria del Excel." tag="Agosto" /><div className="chart-wrap"><ResponsiveContainer width="100%" height="100%"><LineChart data={data.daily} margin={{ top: 18, right: 15, left: 0, bottom: 0 }}><CartesianGrid strokeDasharray="3 5" vertical={false} stroke="#e3ebe7" /><XAxis dataKey="fecha" axisLine={false} tickLine={false} tickFormatter={value => String(value).slice(8)} tick={{ fill: '#71807d', fontSize: 10 }} /><YAxis axisLine={false} tickLine={false} tickFormatter={value => compactFormat.format(value)} width={45} tick={{ fill: '#879590', fontSize: 10 }} /><Tooltip contentStyle={tooltipStyle} formatter={value => `${numberFormat.format(Number(value) || 0)} HL`} /><Line type="monotone" dataKey="valor" name="Producción" stroke="#25c99a" strokeWidth={3} dot={{ r: 3, fill: '#25c99a', strokeWidth: 0 }} activeDot={{ r: 5 }} animationDuration={1200} /></LineChart></ResponsiveContainer></div></article>
          <article className="panel reveal" style={{ animationDelay: '80ms' }}><PanelTitle eyebrow="PROD MES" title={`Material ${data.productMonthMaterial}`} description={data.productMonth[0]?.descripcion || 'Seguimiento mensual por planta'} /><div className="chart-wrap"><ResponsiveContainer width="100%" height="100%"><LineChart data={productMonthData} margin={{ top: 18, right: 12, left: 0, bottom: 0 }}><CartesianGrid strokeDasharray="3 5" vertical={false} stroke="#e3ebe7" /><XAxis dataKey="month" axisLine={false} tickLine={false} tick={{ fill: '#71807d', fontSize: 10 }} /><YAxis axisLine={false} tickLine={false} tickFormatter={value => compactFormat.format(value)} width={40} tick={{ fill: '#879590', fontSize: 10 }} /><Tooltip contentStyle={tooltipStyle} formatter={value => `${numberFormat.format(Number(value) || 0)} HL`} />{data.productMonth.map((row, index) => <Line key={row.planta} type="monotone" dataKey={row.planta} stroke={palette[index % palette.length]} strokeWidth={2} dot={false} animationDuration={900 + index * 120} />)}</LineChart></ResponsiveContainer></div><div className="mini-legend">{data.productMonth.map((row, index) => <span key={row.planta}><i style={{ background: palette[index % palette.length] }} />{row.planta}</span>)}</div></article></section>
          <article className="panel reveal section-panel"><PanelTitle eyebrow="DATA · MATERIALES" title="Materiales con mayor producción" description="Ranking calculado sobre los movimientos y filtros activos." tag="Top 10" /><div className="chart-wrap plant-chart"><ResponsiveContainer width="100%" height="100%"><BarChart data={topMaterials} layout="vertical" margin={{ top: 8, right: 20, left: 10, bottom: 0 }}><CartesianGrid strokeDasharray="3 5" horizontal={false} stroke="#e6ece9" /><XAxis type="number" axisLine={false} tickLine={false} tickFormatter={value => compactFormat.format(value)} tick={{ fill: '#879590', fontSize: 10 }} /><YAxis type="category" dataKey="material" axisLine={false} tickLine={false} width={70} tick={{ fill: '#526963', fontSize: 10 }} /><Tooltip contentStyle={tooltipStyle} formatter={value => `${numberFormat.format(Number(value) || 0)} HL`} /><Bar dataKey="produccion" name="Producción" fill="#25c99a" radius={[0, 5, 5, 0]} maxBarSize={20} animationDuration={1000} /></BarChart></ResponsiveContainer></div></article>
          <OperationTable rows={filteredOperations} />
        </>}

        {view === 'comercial' && <>
          <section className="metrics"><MetricCard icon="PR" title="Producción comercial" value={commercialTotals.produccion} suffix="HL" note="Hoja Producción" tone="green" delay={50} /><MetricCard icon="VT" title="Venta consolidada" value={commercialTotals.venta} suffix="HL" note={`${commercialTotals.produccion ? Math.round(commercialTotals.venta / commercialTotals.produccion * 100) : 0}% de producción`} tone="blue" delay={100} /><MetricCard icon="FC" title="Pronóstico" value={commercialTotals.pronostico} suffix="HL" note="Demanda estimada" tone="violet" delay={150} /><MetricCard icon="%" title="Cumplimiento" value={commercialTotals.pronostico ? Math.round(commercialTotals.produccion / commercialTotals.pronostico * 100) : 0} suffix="%" note="Producción frente a FRC" tone="amber" delay={200} /></section>
          <section className="dashboard-grid"><article className="panel reveal"><PanelTitle eyebrow="PRODUCCIÓN" title="Producción, venta y pronóstico" description="Comparativo por material desde la hoja comercial." tag="Top 9" /><div className="chart-wrap tall"><ResponsiveContainer width="100%" height="100%"><BarChart data={commercialChart} margin={{ top: 18, right: 10, left: 0, bottom: 0 }} barGap={3}><CartesianGrid strokeDasharray="3 5" vertical={false} stroke="#e3ebe7" /><XAxis dataKey="material" axisLine={false} tickLine={false} tick={{ fill: '#71807d', fontSize: 10 }} /><YAxis axisLine={false} tickLine={false} tickFormatter={value => compactFormat.format(value)} width={45} tick={{ fill: '#879590', fontSize: 10 }} /><Tooltip contentStyle={tooltipStyle} formatter={value => `${numberFormat.format(Number(value) || 0)} HL`} /><Legend iconType="circle" iconSize={7} wrapperStyle={{ fontSize: 11 }} /><Bar dataKey="produccion" name="Producción" fill="#25c99a" radius={[5, 5, 0, 0]} maxBarSize={24} /><Bar dataKey="venta" name="Venta" fill="#91d5c1" radius={[5, 5, 0, 0]} maxBarSize={24} /><Line dataKey="pronostico" name="Pronóstico" type="monotone" stroke="#f1b543" strokeWidth={3} dot={{ r: 3, strokeWidth: 0, fill: '#f1b543' }} /></BarChart></ResponsiveContainer></div></article>
          <article className="panel reveal" style={{ animationDelay: '80ms' }}><PanelTitle eyebrow="INV INIC" title="Inventario inicial por mes" description="Acumulado desde la base de inventarios." tag="HL" /><div className="chart-wrap tall"><ResponsiveContainer width="100%" height="100%"><AreaChart data={inventoryData} margin={{ top: 18, right: 12, left: 0, bottom: 0 }}><defs><linearGradient id="inventory" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#5787f6" stopOpacity={.4}/><stop offset="100%" stopColor="#5787f6" stopOpacity={.02}/></linearGradient></defs><CartesianGrid strokeDasharray="3 5" vertical={false} stroke="#e3ebe7" /><XAxis dataKey="month" axisLine={false} tickLine={false} tick={{ fill: '#71807d', fontSize: 10 }} /><YAxis axisLine={false} tickLine={false} tickFormatter={value => compactFormat.format(value)} width={42} tick={{ fill: '#879590', fontSize: 10 }} /><Tooltip contentStyle={tooltipStyle} formatter={value => `${numberFormat.format(Number(value) || 0)} HL`} /><Area type="monotone" dataKey="inventario" name="Inventario" stroke="#5787f6" fill="url(#inventory)" strokeWidth={3} animationDuration={1200} /></AreaChart></ResponsiveContainer></div></article></section>
          <CommercialTable rows={filteredCommercial} />
        </>}

        {view === 'fuentes' && <Sources data={data} />}
      </>}
    </main>
  </div>
}

function Filter({ label, value, options, onChange, capitalize = false }: { label: string; value: string; options: string[]; onChange: (value: string) => void; capitalize?: boolean }) {
  return <label><span>{label}</span><select value={value} onChange={event => onChange(event.target.value)}>{options.map(option => <option key={option} value={option}>{capitalize ? titleCase(option) : option}</option>)}</select></label>
}

function EntryModal({ form, plants, regions, categories, packages, onChange, onClose, onSubmit }: { form: ManualForm; plants: string[]; regions: string[]; categories: string[]; packages: string[]; onChange: (key: keyof ManualForm, value: string) => void; onClose: () => void; onSubmit: (event: FormEvent) => void }) {
  const textField = (key: keyof ManualForm, label: string, type = 'text') => <label>{label}<input required={['planta', 'regional', 'material', 'descripcion'].includes(key)} type={type} value={form[key]} onChange={event => onChange(key, event.target.value)} /></label>
  const selectField = (key: keyof ManualForm, label: string, options: string[]) => <label>{label}<select value={form[key]} onChange={event => onChange(key, event.target.value)}>{options.map(option => <option key={option}>{option}</option>)}</select></label>
  return <div className="modal-backdrop" onMouseDown={event => event.target === event.currentTarget && onClose()}><form className="entry-modal" onSubmit={onSubmit}><div className="form-title"><div><p className="eyebrow">NUEVO REGISTRO</p><h2>Registrar producción completa</h2></div><button type="button" className="close" onClick={onClose}>×</button></div><div className="form-grid">{textField('planta', 'Planta')}{textField('regional', 'Regional')}{selectField('categoria', 'Categoría', categories.length ? categories : ['Cervezas', 'Maltas', 'Aguas'])}{selectField('empaque', 'Empaque', packages.length ? packages : ['Botellas', 'Latas', 'Pet', 'Barril'])}{selectField('mes', 'Mes', monthOrder)}{textField('material', 'Material')}{textField('descripcion', 'Descripción')}{textField('produccion', 'Producción (HL)', 'number')}{textField('unidades', 'Unidades', 'number')}{textField('inventario', 'Inventario inicial (HL)', 'number')}{textField('pronostico', 'Pronóstico (HL)', 'number')}{textField('sop', 'S&OP (HL)', 'number')}{textField('venta', 'Venta (HL)', 'number')}</div><div className="form-actions"><button type="button" className="button secondary" onClick={onClose}>Cancelar</button><button className="button primary">Guardar producción</button></div><datalist id="plants">{plants.map(item => <option key={item}>{item}</option>)}</datalist><datalist id="regions">{regions.map(item => <option key={item}>{item}</option>)}</datalist></form></div>
}

function PlantMonthTable({ rows, exportTotal }: { rows: WorkbookData['plantMonth']; exportTotal: number }) {
  return <article className="panel table-panel reveal"><PanelTitle eyebrow="TD PLANTAS MES" title="Matriz mensual por planta y categoría" description="Valores tomados directamente de la tabla dinámica mostrada en el Excel." tag={`${rows.length} series`} /><div className="reconcile-note"><span>i</span><p>Esta tabla no incluye <strong>{numberFormat.format(exportTotal)} HL de Exportación</strong>. Ese volumen está separado en “Prod Categorías” y sí forma parte del total general calculado desde “Data”.</p></div><div className="table-scroll"><table><thead><tr><th>Planta</th><th>Categoría</th>{monthOrder.slice(0, 8).map(month => <th className="number" key={month}>{month.slice(0, 3)}</th>)}<th className="number">Total</th></tr></thead><tbody>{rows.map(row => <tr key={`${row.planta}-${row.categoria}`}><td className="strong">{row.planta}</td><td>{row.categoria}</td>{monthOrder.slice(0, 8).map(month => <td className="number" key={month}>{row.values[month] ? numberFormat.format(row.values[month]) : '—'}</td>)}<td className="number strong">{numberFormat.format(row.total)}</td></tr>)}</tbody></table></div></article>
}

function OperationTable({ rows }: { rows: Operation[] }) {
  return <article className="panel table-panel reveal"><PanelTitle eyebrow="DATA" title="Movimientos de ejecución" description="Detalle de órdenes desde la primera hoja del libro." tag={`${numberFormat.format(rows.length)} filas`} /><div className="table-scroll"><table><thead><tr><th>Orden</th><th>Material</th><th>Descripción</th><th>Planta</th><th>Regional</th><th>Categoría</th><th>Empaque</th><th>Mes</th><th className="number">Hectolitros</th></tr></thead><tbody>{rows.slice(0, 30).map(row => <tr key={row.id}><td><span className="code-pill">{row.orden}</span></td><td>{row.material}</td><td className="description-cell">{row.descripcion}</td><td>{row.planta}</td><td>{row.regional}</td><td>{row.categoria}</td><td>{row.empaque}</td><td className="capitalize">{row.mes}</td><td className="number strong">{numberFormat.format(row.hectolitros)}</td></tr>)}</tbody></table></div></article>
}

function CommercialTable({ rows }: { rows: Commercial[] }) {
  return <article className="panel table-panel reveal"><PanelTitle eyebrow="PRODUCCIÓN" title="Detalle comercial" description="Inventario, pronóstico, S&OP, venta y variaciones por material." tag={`${numberFormat.format(rows.length)} filas`} /><div className="table-scroll"><table><thead><tr><th>Material</th><th>Descripción</th><th>Regional</th><th>Mes</th><th className="number">Producción</th><th className="number">Inv. inicial</th><th className="number">Pronóstico</th><th className="number">S&OP</th><th className="number">Venta</th><th>Estado</th></tr></thead><tbody>{rows.slice(0, 30).map(row => { const rate = row.pronostico ? Math.round(row.produccion / row.pronostico * 100) : 0; return <tr key={row.id}><td><span className="code-pill">{row.material}</span></td><td className="description-cell">{row.descripcion}</td><td>{row.regional}</td><td className="capitalize">{row.mes}</td><td className="number strong">{numberFormat.format(row.produccion)}</td><td className="number">{numberFormat.format(row.inventario)}</td><td className="number">{numberFormat.format(row.pronostico)}</td><td className="number">{numberFormat.format(row.sop)}</td><td className="number">{numberFormat.format(row.venta)}</td><td><span className={`state ${rate >= 100 ? 'ok' : rate >= 80 ? 'warning' : 'low'}`}>{rate >= 100 ? 'Cumplido' : rate >= 80 ? 'En curso' : 'Bajo'} · {rate}%</span></td></tr> })}</tbody></table></div></article>
}

function Sources({ data }: { data: WorkbookData }) {
  const stats = [
    ['Movimientos', data.operations.length], ['Registros comerciales', data.commercial.length],
    ['Series planta/mes', data.plantMonth.length], ['Inventarios', data.inventory.length],
  ] as const
  return <><section className="source-summary reveal"><div><p className="eyebrow">LIBRO COMPLETO</p><h2>Todo el Excel está conectado</h2><p>La primera hoja alimenta los cálculos filtrables; las seis vistas restantes conservan la organización y métricas especializadas del archivo.</p></div><div className="source-stats">{stats.map(([label, value]) => <div key={label}><strong>{numberFormat.format(value)}</strong><span>{label}</span></div>)}</div></section><section className="source-grid">{data.sheets.map((sheet, index) => <article className="source-card reveal" style={{ animationDelay: `${index * 55}ms` }} key={sheet.name}><div className="sheet-number">{String(index + 1).padStart(2, '0')}</div><div><h3>{sheet.name}</h3><p>{sheet.role}</p><span>{numberFormat.format(sheet.rows)} filas con contenido</span></div><i className="source-status">Activo</i></article>)}</section></>
}
