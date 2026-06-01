import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react'
import * as XLSX from 'xlsx'
import {
  Bar,
  BarChart,
  CartesianGrid,
  Legend,
  Line,
  LineChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts'

const API_BASE     = 'http://127.0.0.1:8000/spread'
const API_PORTFOLIO = 'http://127.0.0.1:8000/portfolio'
const API_POSICION  = 'http://127.0.0.1:8000/portfolio/posicion'
const API_SIMULATE  = 'http://127.0.0.1:8000/simulate'
const API_PPP       = 'http://127.0.0.1:8000/contratos/ppp'

const MESES_ES = [
  'Enero',
  'Febrero',
  'Marzo',
  'Abril',
  'Mayo',
  'Junio',
  'Julio',
  'Agosto',
  'Septiembre',
  'Octubre',
  'Noviembre',
  'Diciembre',
] as const

const DIAS_ES = [
  'Domingo',
  'Lunes',
  'Martes',
  'Miércoles',
  'Jueves',
  'Viernes',
  'Sábado',
] as const

const CLASE_SELECT =
  'rounded-lg border border-gray-700 bg-gray-950 px-4 py-2.5 text-gray-100 outline-none transition focus:border-emerald-500/50 focus:ring-2 focus:ring-emerald-500/30'

const COLORES_MES_COMPARATIVO = [
  { text: 'text-blue-400', bg: 'bg-blue-500/15', ring: 'ring-blue-500/40' },
  { text: 'text-emerald-400', bg: 'bg-emerald-500/15', ring: 'ring-emerald-500/40' },
  { text: 'text-amber-400', bg: 'bg-amber-500/15', ring: 'ring-amber-500/40' },
  { text: 'text-orange-400', bg: 'bg-orange-500/15', ring: 'ring-orange-500/40' },
] as const

const ENSO_COMPLETO = [
  { key: 'nino_2010q1',  label: 'El Niño Fuerte (2010 Q1)',           fenomeno: 'nino',    intensidad: 'Fuerte',      pb_desde: '2010-01-01', pb_hasta: '2010-03-31' },
  { key: 'nina_2010q3',  label: 'La Niña Fuerte (2010 Q3)',           fenomeno: 'nina',    intensidad: 'Fuerte',      pb_desde: '2010-07-01', pb_hasta: '2010-09-30' },
  { key: 'nina_2011',    label: 'La Niña Moderada (2011)',            fenomeno: 'nina',    intensidad: 'Moderada',    pb_desde: '2011-01-01', pb_hasta: '2011-12-31' },
  { key: 'nino_2015',    label: 'El Niño Muy Fuerte (2015-2016)',     fenomeno: 'nino',    intensidad: 'Muy fuerte',  pb_desde: '2015-01-01', pb_hasta: '2016-05-31' },
  { key: 'nina_2020',    label: 'La Niña Persistente (2020-2022)',    fenomeno: 'nina',    intensidad: 'Persistente', pb_desde: '2020-09-01', pb_hasta: '2022-03-31' },
  { key: 'nino_2023',    label: 'El Niño Fuerte (2023-2024)',         fenomeno: 'nino',    intensidad: 'Fuerte',      pb_desde: '2023-06-01', pb_hasta: '2024-05-31' },
  { key: 'neutral_2025', label: 'Neutral / Transición (2025)',        fenomeno: 'neutral', intensidad: 'Débil',       pb_desde: '2025-01-01', pb_hasta: '2025-12-31' },
] as const

type EnsoKey = typeof ENSO_COMPLETO[number]['key']

const PRECIO_MOCK_COMPRA_R  = 320
const PRECIO_MOCK_COMPRA_NR = 290
const PRECIO_MOCK_VENTA     = 380


type VistaAnalisis = 'dia' | 'mes' | 'comparativo'
type DashboardTab = 'spread' | 'portafolio' | 'simulador'
type SimTipo = 'compra' | 'venta'
type SimMercado = 'regulado' | 'no_regulado' | 'ambos'
type WizFuentePB = 'historico' | 'enso' | 'proyectado'
type SimPerfilTipo = 'plano' | 'bloques' | 'solar' | 'excel'
type SimRecomendacion = 'verde' | 'amarillo' | 'rojo'

interface SimBloque { horaInicio: number; horaFin: number; mwhMes: number }

interface FilaSpread {
  fecha: string
  hora: number
  precio_bolsa: number
  spread: number
}

interface ResumenSpread {
  spread_promedio: number
  spread_minimo: number
  spread_maximo: number
  horas_negativas: number
  porcentaje_negativo: number
}

interface SpreadResponse {
  resumen: ResumenSpread
  datos: FilaSpread[]
  total_filas: number
}

interface FilaHoraDia {
  hora: number
  precio_bolsa: number | null
  precio_contrato: number
  spread: number | null
}

interface FilaHoraMes {
  hora: number
  precio_bolsa: number | null
  precio_contrato: number
  spread: number | null
  percentiles: { p10: number; p50: number; p90: number } | null
}

interface FilaPortfolio {
  fecha: string
  hora: number
  tipo_dia?: string
  compra_r: number
  compra_nr: number
  venta: number
  demanda_r: number    // NUEVO
  demanda_nr: number   // NUEVO
  posicion_neta: number
  costo_bolsa: number
}

/** Un elemento de resumen_mensual devuelto por /portfolio/posicion. */
interface FilaPosicionMensual {
  mes: string          // "YYYY-MM"
  compra_r_mwh: number
  compra_nr_mwh: number
  venta_mwh: number
  demanda_r_mwh?: number   // NUEVO
  demanda_nr_mwh?: number  // NUEVO
  posicion_neta_mwh: number
  dias: number         // días reales del mes procesados
}

/** Distribución de tipos de día devuelta por /portfolio/posicion. */
interface DistribucionTipoDia {
  ordinarios: number
  sabados: number
  domingos: number
  festivos: number
}

interface PortfolioResumen {
  posicion_neta_total_kwh: number
  demanda_r_total_kwh: number     // NUEVO
  demanda_nr_total_kwh: number    // NUEVO
  costo_bolsa_total_cop: number
  hora_pico_compra: number | null
  hora_pico_venta: number | null
}

interface SimResumen {
  posicion_neta_total_mwh: number
  costo_bolsa_total_mcop: number
  hora_pico_compra: number
  hora_pico_venta: number
}

interface SimPerfilHora {
  hora: number
  posicion_antes_mwh: number
  posicion_despues_mwh: number
}

interface SimMes {
  mes: string
  pos_actual_mwh: number
  pos_nueva_mwh: number
  diferencia_mwh: number
  costo_actual_mcop: number
  costo_nuevo_mcop: number
  ahorro_mcop: number
}

interface SimResultado {
  resumen_antes: SimResumen
  resumen_despues: SimResumen
  recomendacion: SimRecomendacion
  delta_costo_mcop: number
  perfil_horario: SimPerfilHora[]
  por_mes: SimMes[]
  ppp_contratos?: PPPResumen | null
}

/** PPP de una categoría de operación (compra R, compra NR o venta). */
interface PPPCategoria {
  ppp: number | null
  tipo: 'Indexado' | 'Proyectado' | 'Sin datos'
}

/** Resumen de PPP real por categoría devuelto por /contratos/ppp. */
interface PPPResumen {
  compra_r:      PPPCategoria
  compra_nr:     PPPCategoria
  venta:         PPPCategoria
  pld_excluidos: number
  contratos_pc:  number
}

function formatNumero(valor: number, decimales = 2): string {
  return valor.toFixed(decimales)
}

function parseFecha(fecha: string): Date | null {
  const texto = fecha.trim()
  const iso = texto.match(/^(\d{4})-(\d{2})-(\d{2})/)
  if (iso) {
    const anio = Number(iso[1])
    const mes = Number(iso[2]) - 1
    const dia = Number(iso[3])
    const dt = new Date(anio, mes, dia)
    if (
      dt.getFullYear() === anio &&
      dt.getMonth() === mes &&
      dt.getDate() === dia
    ) {
      return dt
    }
  }
  const dt = new Date(texto)
  return Number.isNaN(dt.getTime()) ? null : dt
}

function normalizarFecha(fecha: string): string {
  const dt = parseFecha(fecha)
  if (!dt) return fecha.slice(0, 10)
  const anio = dt.getFullYear()
  const mes = String(dt.getMonth() + 1).padStart(2, '0')
  const dia = String(dt.getDate()).padStart(2, '0')
  return `${anio}-${mes}-${dia}`
}

function mesDeFecha(fecha: string): string {
  const normalizada = normalizarFecha(fecha)
  return normalizada.slice(0, 7)
}

function formatMes(ym: string): string {
  const [anio, mes] = ym.split('-')
  const indice = parseInt(mes, 10) - 1
  if (indice < 0 || indice > 11) return ym
  return `${MESES_ES[indice]} ${anio}`
}

function formatFechaLegible(fecha: string): string {
  const dt = parseFecha(fecha)
  if (!dt) return fecha
  const diaSemana = DIAS_ES[dt.getDay()]
  const dia = dt.getDate()
  const mes = MESES_ES[dt.getMonth()]
  const anio = dt.getFullYear()
  return `${diaSemana} ${dia} ${mes} ${anio}`
}

function percentil(valores: number[], p: number): number {
  if (valores.length === 0) return NaN
  const ordenados = [...valores].sort((a, b) => a - b)
  if (ordenados.length === 1) return ordenados[0]
  const indice = (p / 100) * (ordenados.length - 1)
  const inferior = Math.floor(indice)
  const superior = Math.ceil(indice)
  if (inferior === superior) return ordenados[inferior]
  return (
    ordenados[inferior] +
    (ordenados[superior] - ordenados[inferior]) * (indice - inferior)
  )
}

function fechasUnicas(datos: FilaSpread[]): string[] {
  const fechas = new Set<string>()
  for (const fila of datos) {
    fechas.add(normalizarFecha(fila.fecha))
  }
  return [...fechas].sort()
}

function mesesUnicos(datos: FilaSpread[]): string[] {
  const meses = new Set<string>()
  for (const fila of datos) {
    meses.add(mesDeFecha(fila.fecha))
  }
  return [...meses].sort()
}

function filasPorHoraDelDia(
  datos: FilaSpread[],
  fecha: string,
  precioContrato: number,
): FilaHoraDia[] {
  const fechaNorm = normalizarFecha(fecha)
  const delDia = datos.filter((d) => normalizarFecha(d.fecha) === fechaNorm)
  return Array.from({ length: 24 }, (_, i) => {
    const hora = i + 1
    const fila = delDia.find((d) => d.hora === hora)
    return {
      hora,
      precio_bolsa: fila?.precio_bolsa ?? null,
      precio_contrato: precioContrato,
      spread: fila?.spread ?? null,
    }
  })
}

function filasPromedioMes(
  datos: FilaSpread[],
  mes: string,
  precioContrato: number,
): FilaHoraMes[] {
  const delMes = datos.filter((d) => mesDeFecha(d.fecha) === mes)
  return Array.from({ length: 24 }, (_, i) => {
    const hora = i + 1
    const filas = delMes.filter((d) => d.hora === hora)
    if (filas.length === 0) {
      return {
        hora,
        precio_bolsa: null,
        precio_contrato: precioContrato,
        spread: null,
        percentiles: null,
      }
    }
    const pbs = filas.map((d) => d.precio_bolsa)
    const spreads = filas.map((d) => d.spread)
    return {
      hora,
      precio_bolsa: pbs.reduce((s, v) => s + v, 0) / pbs.length,
      precio_contrato: precioContrato,
      spread: spreads.reduce((s, v) => s + v, 0) / spreads.length,
      percentiles: {
        p10: percentil(pbs, 10),
        p50: percentil(pbs, 50),
        p90: percentil(pbs, 90),
      },
    }
  })
}

function pbPromedioPorHoraMes(
  datos: FilaSpread[],
  mes: string,
): { hora: number; pb: number | null }[] {
  const delMes = datos.filter((d) => mesDeFecha(d.fecha) === mes)
  return Array.from({ length: 24 }, (_, i) => {
    const hora = i + 1
    const filas = delMes.filter((d) => d.hora === hora)
    if (filas.length === 0) return { hora, pb: null }
    const avg = filas.reduce((s, d) => s + d.precio_bolsa, 0) / filas.length
    return { hora, pb: avg }
  })
}


function formatMiles(valor: number, decimales = 2): string {
  return valor.toLocaleString('en-US', {
    minimumFractionDigits: 0,
    maximumFractionDigits: decimales,
  })
}

function _dateToIso(d: Date): string {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
}

function _fechaPascua(anio: number): Date {
  const a = anio % 19, b = Math.floor(anio / 100), c = anio % 100
  const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25)
  const g = Math.floor((b - f + 1) / 3)
  const h = (19 * a + b - d - g + 15) % 30
  const i = Math.floor(c / 4), k = c % 4
  const l = (32 + 2 * e + 2 * i - h - k) % 7
  const m = Math.floor((a + 11 * h + 22 * l) / 451)
  const mes = Math.floor((h + l - 7 * m + 114) / 31)
  const dia = ((h + l - 7 * m + 114) % 31) + 1
  return new Date(anio, mes - 1, dia)
}

function _addDias(d: Date, n: number): Date {
  const r = new Date(d); r.setDate(r.getDate() + n); return r
}

function _siguienteLunes(d: Date): Date {
  const r = new Date(d); r.setDate(r.getDate() + (8 - r.getDay()) % 7); return r
}

const _cacheFestivos = new Map<number, Set<string>>()

function festivosColombia(anio: number): Set<string> {
  if (_cacheFestivos.has(anio)) return _cacheFestivos.get(anio)!
  const pascua = _fechaPascua(anio)
  const f = new Set<string>()
  const fijos: [number, number][] = [[1,1],[5,1],[7,20],[8,7],[12,8],[12,25]]
  for (const [mes, dia] of fijos)
    f.add(`${anio}-${String(mes).padStart(2,'0')}-${String(dia).padStart(2,'0')}`)
  const emiliani = [
    new Date(anio,0,6),  new Date(anio,2,19), new Date(anio,5,29),
    new Date(anio,7,15), new Date(anio,9,12), new Date(anio,10,1), new Date(anio,10,11),
  ]
  for (const base of emiliani) f.add(_dateToIso(_siguienteLunes(base)))
  f.add(_dateToIso(_addDias(pascua, -3)))
  f.add(_dateToIso(_addDias(pascua, -2)))
  f.add(_dateToIso(_siguienteLunes(_addDias(pascua, 43))))
  f.add(_dateToIso(_siguienteLunes(_addDias(pascua, 64))))
  f.add(_dateToIso(_siguienteLunes(_addDias(pascua, 71))))
  _cacheFestivos.set(anio, f)
  return f
}

function tipoDiaDesdefecha(fechaStr: string): string {
  const dt = parseFecha(fechaStr)
  if (!dt) return 'Ordinario'
  if (festivosColombia(dt.getFullYear()).has(normalizarFecha(fechaStr))) return 'Festivo'
  const dow = dt.getDay()
  if (dow === 0) return 'Domingo'
  if (dow === 6) return 'Sábado'
  return 'Ordinario'
}

function parseNumero(valor: unknown): number {
  if (typeof valor === 'number') return Number.isFinite(valor) ? valor : 0
  if (typeof valor === 'string') {
    const n = Number(valor)
    return Number.isFinite(n) ? n : 0
  }
  return 0
}

function parseFilaPortfolio(valor: unknown): FilaPortfolio | null {
  if (!valor || typeof valor !== 'object') return null
  const fila = valor as Record<string, unknown>
  const hora = parseNumero(fila.hora)
  if (hora < 1 || hora > 24) return null
  const compraR  = parseNumero(fila.compra_r_kwh)
  const compraNr = parseNumero(fila.compra_nr_kwh)
  const venta    = parseNumero(fila.venta_kwh)
  const demandaR  = parseNumero(fila.demanda_r_kwh)
  const demandaNr = parseNumero(fila.demanda_nr_kwh)
  const posicionNeta =
    fila.posicion_neta_kwh != null
      ? parseNumero(fila.posicion_neta_kwh)
      : compraR + compraNr - venta - demandaR - demandaNr
  return {
    fecha: String(fila.fecha ?? ''),
    hora,
    tipo_dia:
      typeof fila.tipo_dia === 'string' && fila.tipo_dia.trim() !== ''
        ? fila.tipo_dia
        : undefined,
    compra_r:     compraR,
    compra_nr:    compraNr,
    venta,
    demanda_r:    demandaR,
    demanda_nr:   demandaNr,
    posicion_neta: posicionNeta,
    costo_bolsa:  parseNumero(fila.costo_bolsa_cop),
  }
}

function claseSpread(
  spread: number | null,
  posicion: 'vendedor' | 'comprador',
): string {
  if (spread === null) return 'text-gray-500'
  if (posicion === 'vendedor') {
    return spread < 0 ? 'text-red-400' : 'text-emerald-400'
  }
  return spread > 0 ? 'text-red-400' : 'text-emerald-400'
}

function App() {
  const [tabActiva, setTabActiva] = useState<DashboardTab>('spread')
  const [posicion, setPosicion] = useState<'vendedor' | 'comprador'>('vendedor')
  const [precioContrato, setPrecioContrato] = useState(350)
  const [fechaInicio, setFechaInicio] = useState('2024-01-01')
  const [fechaFin, setFechaFin] = useState(
    () => new Date().toISOString().split('T')[0],
  )
  const [cargando, setCargando] = useState(true)
  const [recalculando, setRecalculando] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [resumen, setResumen] = useState<ResumenSpread | null>(null)
  const [datos, setDatos] = useState<FilaSpread[]>([])

  const [vistaAnalisis, setVistaAnalisis] = useState<VistaAnalisis>('dia')
  const [fechaSeleccionada, setFechaSeleccionada] = useState('')
  const [mesSeleccionado, setMesSeleccionado] = useState('')
  const [mesesComparativo, setMesesComparativo] = useState<string[]>([])
  const [portfolioFechaInicio, setPortfolioFechaInicio] = useState('2026-01-01')
  const [portfolioFechaFin, setPortfolioFechaFin] = useState(
    () => new Date().toISOString().split('T')[0],
  )
  const [portfolioCargando, setPortfolioCargando] = useState(false)
  const [portfolioError, setPortfolioError] = useState<string | null>(null)
  const [portfolioDatos, setPortfolioDatos] = useState<FilaPortfolio[]>([])
  const [portfolioFiltroMes, setPortfolioFiltroMes] = useState('')
  const [portfolioFiltroTipoDia, setPortfolioFiltroTipoDia] = useState('todos')
  // Controla que el portafolio se cargue automáticamente sólo la primera vez
  const portfolioCargadoRef = useRef(false)

  // ── Wizard ─────────────────────────────────────────────────────────────────
  const [wizPaso, setWizPaso] = useState(1)

  // Paso 1 — Período
  const [wizPeriodoInicio, setWizPeriodoInicio] = useState('2026-01-01')
  const [wizPeriodoFin, setWizPeriodoFin]       = useState('2026-12-31')
  const [wizPosicionMensual,   setWizPosicionMensual]   = useState<FilaPosicionMensual[]>([])
  const [wizTotalDias,         setWizTotalDias]         = useState<number>(0)
  const [wizDistTipoDia,       setWizDistTipoDia]       = useState<DistribucionTipoDia | null>(null)
  const [wizPortfolioCargando, setWizPortfolioCargando] = useState(false)
  const [wizPortfolioError,    setWizPortfolioError]    = useState<string | null>(null)

  // Paso 2 — Escenario PB
  const [wizFuentePB, setWizFuentePB] = useState<WizFuentePB>('enso')
  const [wizEnsoKey,  setWizEnsoKey]  = useState<EnsoKey>('nino_2023')
  const [wizPBDesde,  setWizPBDesde]  = useState('2023-06-01')
  const [wizPBHasta,  setWizPBHasta]  = useState('2024-05-31')
  const [wizPBDatos,    setWizPBDatos]    = useState<FilaSpread[]>([])
  const [wizPBCargando, setWizPBCargando] = useState(false)
  const [wizPBError,    setWizPBError]    = useState<string | null>(null)

  // Paso 3 — Nuevo contrato
  const [wizSimTipo,        setWizSimTipo]        = useState<SimTipo>('compra')
  const [wizSimContraparte, setWizSimContraparte] = useState('')
  const [wizSimPrecio,      setWizSimPrecio]      = useState(350)
  const [wizSimPerfilTipo,  setWizSimPerfilTipo]  = useState<SimPerfilTipo>('plano')
  const [wizSimEnergiaKwh,  setWizSimEnergiaKwh]  = useState(1000)
  const [wizSimBloques,     setWizSimBloques]     = useState<SimBloque[]>([{ horaInicio: 8, horaFin: 17, mwhMes: 1000 }])
  const [wizSimExcel12x24,  setWizSimExcel12x24]  = useState<number[][] | null>(null)
  const [wizSimExcelNombre, setWizSimExcelNombre] = useState('')
  const [wizContratoInicio, setWizContratoInicio] = useState('2026-01-01')
  const [wizContratoFin,    setWizContratoFin]    = useState('2026-12-31')
  const [wizSimTipoMercado, setWizSimTipoMercado] = useState<SimMercado>('regulado')
  const wizExcelInputRef = useRef<HTMLInputElement>(null)

  // Paso 4 — Resultado
  const [wizResultado,   setWizResultado]   = useState<SimResultado | null>(null)
  const [wizSimCargando, setWizSimCargando] = useState(false)
  const [wizSimError,    setWizSimError]    = useState<string | null>(null)

  // PPP real de contratos PC (cargado en paralelo con la posición)
  const [wizPPPResumen,  setWizPPPResumen]  = useState<PPPResumen | null>(null)
  const [portfolioPPP,   setPortfolioPPP]   = useState<PPPResumen | null>(null)

  const datosConSpread = useMemo(() => {
    return datos.map((d) => ({
      ...d,
      spread:
        posicion === 'vendedor'
          ? precioContrato - d.precio_bolsa
          : d.precio_bolsa - precioContrato,
    }))
  }, [datos, posicion, precioContrato])

  const resumenLocal = useMemo(() => {
    if (datosConSpread.length === 0) return null
    const spreads = datosConSpread.map((d) => d.spread)
    const spread_promedio =
      spreads.reduce((s, v) => s + v, 0) / spreads.length
    const horasCriticas =
      posicion === 'vendedor'
        ? spreads.filter((s) => s < 0).length
        : spreads.filter((s) => s > 0).length
    return {
      spread_promedio,
      spread_minimo: Math.min(...spreads),
      spread_maximo: Math.max(...spreads),
      horas_criticas: horasCriticas,
      porcentaje_critico: (horasCriticas / spreads.length) * 100,
    }
  }, [datosConSpread, posicion])

  const fechasDisponibles = useMemo(
    () => fechasUnicas(datosConSpread),
    [datosConSpread],
  )
  const mesesDisponibles = useMemo(
    () => mesesUnicos(datosConSpread),
    [datosConSpread],
  )

  useEffect(() => {
    if (fechasDisponibles.length === 0) return
    setFechaSeleccionada((prev) =>
      prev && fechasDisponibles.includes(prev)
        ? prev
        : fechasDisponibles[fechasDisponibles.length - 1],
    )
  }, [fechasDisponibles])

  useEffect(() => {
    if (mesesDisponibles.length === 0) return
    setMesSeleccionado((prev) =>
      prev && mesesDisponibles.includes(prev)
        ? prev
        : mesesDisponibles[mesesDisponibles.length - 1],
    )
    setMesesComparativo((prev) => {
      const validos = prev.filter((m) => mesesDisponibles.includes(m))
      if (validos.length > 0) return validos
      const ultimos = mesesDisponibles.slice(-2)
      return ultimos.length > 0 ? ultimos : [mesesDisponibles[0]]
    })
  }, [mesesDisponibles])

  async function cargarSpread(
    precio: number,
    inicio: string,
    fin: string,
    opciones?: { inicial?: boolean; signal?: AbortSignal },
  ) {
    const esInicial = opciones?.inicial ?? false

    try {
      if (esInicial) {
        setCargando(true)
      } else {
        setRecalculando(true)
      }
      setError(null)

      const respuesta = await fetch(
        `${API_BASE}?precio_contrato=${precio}&fecha_inicio=${inicio}&fecha_fin=${fin}`,
        { signal: opciones?.signal },
      )

      if (!respuesta.ok) {
        const cuerpo = await respuesta.json().catch(() => null)
        const detalle =
          cuerpo && typeof cuerpo.detail === 'string'
            ? cuerpo.detail
            : `Error ${respuesta.status}`
        throw new Error(detalle)
      }

      const json: SpreadResponse = await respuesta.json()
      setResumen(json.resumen)
      setDatos(json.datos)
    } catch (err) {
      if (err instanceof Error && err.name === 'AbortError') return
      setError(
        err instanceof Error ? err.message : 'No se pudieron cargar los datos',
      )
    } finally {
      if (esInicial) {
        setCargando(false)
      } else {
        setRecalculando(false)
      }
    }
  }

  async function cargarPortafolio(
    inicio: string,
    fin: string,
    signal?: AbortSignal,
  ) {
    try {
      setPortfolioCargando(true)
      setPortfolioError(null)

      // Carga paralela: posición horaria y PPP real de contratos
      const [respPos, respPPP] = await Promise.allSettled([
        fetch(`${API_POSICION}?fecha_inicio=${inicio}&fecha_fin=${fin}`, { signal }),
        fetch(`${API_PPP}?start_date=${inicio}&end_date=${fin}`, { signal }),
      ])

      // Posición horaria (requerida)
      if (respPos.status === 'fulfilled') {
        const respuesta = respPos.value
        if (!respuesta.ok) {
          const cuerpo = await respuesta.json().catch(() => null)
          const detalle = cuerpo && typeof cuerpo.detail === 'string' ? cuerpo.detail : `Error ${respuesta.status}`
          throw new Error(detalle)
        }
        const json = await respuesta.json()
        const candidatos = Array.isArray(json?.datos) ? json.datos : Array.isArray(json) ? json : []
        setPortfolioDatos(candidatos.map(parseFilaPortfolio).filter(Boolean) as FilaPortfolio[])
      } else {
        if ((respPos.reason as Error)?.name === 'AbortError') return
        throw new Error((respPos.reason as Error)?.message ?? 'Error cargando portafolio')
      }

      // PPP real (opcional — fallback a null si falla)
      if (respPPP.status === 'fulfilled' && respPPP.value.ok) {
        const json: PPPResumen = await respPPP.value.json().catch(() => null)
        setPortfolioPPP(json ?? null)
      } else {
        setPortfolioPPP(null)
      }
    } catch (err) {
      if (err instanceof Error && err.name === 'AbortError') return
      setPortfolioError(err instanceof Error ? err.message : 'No se pudieron cargar los datos')
      setPortfolioDatos([])
      setPortfolioPPP(null)
    } finally {
      setPortfolioCargando(false)
    }
  }

  useEffect(() => {
    const controller = new AbortController()
    cargarSpread(350, '2024-01-01', new Date().toISOString().split('T')[0], {
      inicial: true,
      signal: controller.signal,
    })
    return () => controller.abort()
  }, [])

  function handleCalcular() {
    cargarSpread(precioContrato, fechaInicio, fechaFin)
  }

  function handleCalcularPortafolio() {
    cargarPortafolio(portfolioFechaInicio, portfolioFechaFin)
  }

  // Carga automática del portafolio la primera vez que se abre esa tab
  useEffect(() => {
    if (tabActiva !== 'portafolio') return
    if (portfolioCargadoRef.current) return
    portfolioCargadoRef.current = true
    const controller = new AbortController()
    cargarPortafolio(portfolioFechaInicio, portfolioFechaFin, controller.signal)
    return () => controller.abort()
  }, [tabActiva])

  // Sync ENSO → fechas PB
  useEffect(() => {
    if (wizFuentePB !== 'enso') return
    const enso = ENSO_COMPLETO.find(e => e.key === wizEnsoKey)
    if (enso) { setWizPBDesde(enso.pb_desde); setWizPBHasta(enso.pb_hasta) }
  }, [wizFuentePB, wizEnsoKey])

  async function wizCargarPosicion() {
    setWizPortfolioCargando(true); setWizPortfolioError(null)
    try {
      // Carga paralela: posición mensual de contratos y PPP real
      const [respPos, respPPP] = await Promise.allSettled([
        fetch(`${API_POSICION}?fecha_inicio=${wizPeriodoInicio}&fecha_fin=${wizPeriodoFin}`),
        fetch(`${API_PPP}?start_date=${wizPeriodoInicio}&end_date=${wizPeriodoFin}`),
      ])

      // Posición (requerida — lanza si falla)
      if (respPos.status === 'fulfilled') {
        const resp = respPos.value
        if (!resp.ok) { const c = await resp.json().catch(() => null); throw new Error(c?.detail ?? `Error ${resp.status}`) }
        const json = await resp.json()
        setWizPosicionMensual(Array.isArray(json?.resumen_mensual) ? json.resumen_mensual : [])
        setWizTotalDias(typeof json?.total_dias === 'number' ? json.total_dias : 0)
        setWizDistTipoDia(json?.distribucion_tipo_dia ?? null)
      } else {
        throw new Error((respPos.reason as Error)?.message ?? 'Error cargando posición')
      }

      // PPP real (opcional — si falla, el frontend usa precios mock como fallback)
      if (respPPP.status === 'fulfilled' && respPPP.value.ok) {
        const json: PPPResumen = await respPPP.value.json().catch(() => null)
        setWizPPPResumen(json ?? null)
      } else {
        console.warn('[PPP] No se obtuvo PPP real de contratos. Se usarán precios mock como referencia.')
        setWizPPPResumen(null)
      }
    } catch (err) {
      setWizPortfolioError(err instanceof Error ? err.message : 'Error')
      setWizPosicionMensual([]); setWizTotalDias(0); setWizDistTipoDia(null); setWizPPPResumen(null)
    } finally {
      setWizPortfolioCargando(false)
    }
  }

  async function wizCargarPB() {
    if (wizFuentePB === 'historico') {
      const d1 = new Date(wizPBDesde), d2 = new Date(wizPBHasta)
      const meses = (d2.getFullYear() - d1.getFullYear()) * 12 + (d2.getMonth() - d1.getMonth()) + 1
      if (meses > 36) { setWizPBError(`El rango tiene ${meses} meses. Máximo 36 meses.`); return }
      if (meses < 1) { setWizPBError('El período debe tener al menos 1 mes.'); return }
    }
    setWizPBCargando(true); setWizPBError(null)
    try {
      const resp = await fetch(`${API_BASE}?precio_contrato=0&fecha_inicio=${wizPBDesde}&fecha_fin=${wizPBHasta}`)
      if (!resp.ok) { const c = await resp.json().catch(() => null); throw new Error(c?.detail ?? `Error ${resp.status}`) }
      const json = await resp.json()
      setWizPBDatos(Array.isArray(json?.datos) ? json.datos : [])
    } catch (err) { setWizPBError(err instanceof Error ? err.message : 'Error'); setWizPBDatos([]) }
    finally { setWizPBCargando(false) }
  }

  async function wizSimular() {
    setWizSimCargando(true); setWizSimError(null)
    type SimBody = { tipo: string; contraparte: string; precio_cop_kwh: number; pb_desde: string; pb_hasta: string; contrato_inicio: string; contrato_fin: string; tipo_mercado: string; perfil_horario: string; energia_mensual_kwh?: number; bloques?: { hora_ini: number; hora_fin: number; mwh_mes: number }[]; perfil_excel_12x24?: number[][] }
    const body: SimBody = { tipo: wizSimTipo, contraparte: wizSimContraparte, precio_cop_kwh: wizSimPrecio, pb_desde: wizPBDesde, pb_hasta: wizPBHasta, contrato_inicio: wizContratoInicio, contrato_fin: wizContratoFin, tipo_mercado: wizSimTipoMercado, perfil_horario: wizSimPerfilTipo === 'excel' ? 'excel' : wizSimPerfilTipo }
    if (wizSimPerfilTipo === 'plano' || wizSimPerfilTipo === 'solar') body.energia_mensual_kwh = wizSimEnergiaKwh
    else if (wizSimPerfilTipo === 'bloques') body.bloques = wizSimBloques.map(b => ({ hora_ini: b.horaInicio, hora_fin: b.horaFin, mwh_mes: b.mwhMes }))
    else if (wizSimPerfilTipo === 'excel' && wizSimExcel12x24) body.perfil_excel_12x24 = wizSimExcel12x24
    try {
      const resp = await fetch(API_SIMULATE, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) })
      if (!resp.ok) { const c = await resp.json().catch(() => null); throw new Error(c?.detail ?? `Error ${resp.status}`) }
      const json: SimResultado = await resp.json()
      setWizResultado(json); setWizPaso(4)
    } catch (err) { setWizSimError(err instanceof Error ? err.message : 'Error al simular'); setWizResultado(null) }
    finally { setWizSimCargando(false) }
  }

  function wizReset() { setWizPaso(1); setWizPosicionMensual([]); setWizTotalDias(0); setWizDistTipoDia(null); setWizPortfolioError(null); setWizPBDatos([]); setWizPBError(null); setWizResultado(null); setWizSimError(null); setWizSimExcel12x24(null); setWizSimExcelNombre(''); setWizPPPResumen(null) }

  function wizDescargarJSON() {
    if (!wizResultado) return
    const payload = { periodo: { inicio: wizPeriodoInicio, fin: wizPeriodoFin }, escenario_pb: { fuente: wizFuentePB, desde: wizPBDesde, hasta: wizPBHasta }, contrato: { tipo: wizSimTipo, contraparte: wizSimContraparte, precio_cop_kwh: wizSimPrecio, tipo_mercado: wizSimTipoMercado, perfil: wizSimPerfilTipo, inicio: wizContratoInicio, fin: wizContratoFin }, resultado: wizResultado }
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a'); a.href = url; a.download = `simulacion_${new Date().toISOString().slice(0, 10)}.json`; a.click(); URL.revokeObjectURL(url)
  }

  function wizHandleExcelUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]; if (!file) return
    setWizSimExcelNombre(file.name)
    const reader = new FileReader()
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target?.result, { type: 'binary' })
        const ws = wb.Sheets[wb.SheetNames[0]]
        const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 }) as unknown[][]
        const dataRows = rows.filter(r => Array.isArray(r) && r.length >= 24 && typeof r[0] === 'number')
        if (dataRows.length < 1) { setWizSimError('Archivo sin filas numéricas válidas.'); return }
        const matrix: number[][] = dataRows.slice(0, 12).map(row => Array.from({ length: 24 }, (_, i) => parseFloat(String((row as unknown[])[i] ?? 0)) || 0))
        while (matrix.length < 12) matrix.push([...matrix[matrix.length - 1]])
        setWizSimExcel12x24(matrix); setWizSimError(null)
      } catch { setWizSimError('Error al leer el archivo Excel.') }
    }
    reader.readAsBinaryString(file); e.target.value = ''
  }
  function wizAddBloque() { if (wizSimBloques.length >= 3) return; setWizSimBloques(prev => [...prev, { horaInicio: 18, horaFin: 22, mwhMes: 500 }]) }
  function wizRemoveBloque(idx: number) { setWizSimBloques(prev => prev.filter((_, i) => i !== idx)) }
  function wizUpdateBloque(idx: number, campo: keyof SimBloque, valor: number) {
    setWizSimBloques(prev => prev.map((b, i) => i === idx ? { ...b, [campo]: valor, ...(campo === 'horaInicio' && valor > b.horaFin ? { horaFin: valor } : {}) } : b))
  }

  function toggleMesComparativo(mes: string) {
    setMesesComparativo((prev) => {
      if (prev.includes(mes)) return prev.filter((m) => m !== mes)
      if (prev.length >= 4) return prev
      return [...prev, mes].sort()
    })
  }

  const wizResumenMensual = useMemo(() => {
    // Precios PPP reales con fallback a mock cuando no está disponible
    const pppCompraR  = wizPPPResumen?.compra_r?.ppp  ?? PRECIO_MOCK_COMPRA_R
    const pppCompraNr = wizPPPResumen?.compra_nr?.ppp ?? PRECIO_MOCK_COMPRA_NR
    const pppVenta    = wizPPPResumen?.venta?.ppp     ?? PRECIO_MOCK_VENTA
    return wizPosicionMensual.map(m => ({
      mes:         m.mes,
      // Campos en kWh para la tabla del Paso 1
      compraRkwh:  m.compra_r_mwh  * 1_000,
      compraNrkwh: m.compra_nr_mwh * 1_000,
      ventaKwh:    m.venta_mwh     * 1_000,
      demandaRkwh: (m.demanda_r_mwh  ?? 0) * 1_000,
      demandaNrkwh:(m.demanda_nr_mwh ?? 0) * 1_000,
      posNetaKwh:  m.posicion_neta_mwh * 1_000,
      // Campos en MWh para cálculos del Paso 2 (NO cambiar estos)
      compraRmwh:  m.compra_r_mwh,
      compraNrmwh: m.compra_nr_mwh,
      ventaMwh:    m.venta_mwh,
      posNetaMwh:  m.posicion_neta_mwh,
      // COP con PPP real (o mock como fallback)
      copCompraR:  m.compra_r_mwh  * 1_000 * pppCompraR  / 1_000_000,
      copCompraNr: m.compra_nr_mwh * 1_000 * pppCompraNr / 1_000_000,
      copVenta:    m.venta_mwh     * 1_000 * pppVenta    / 1_000_000,
      tipoPos:     m.posicion_neta_mwh > 0 ? 'Comprando en bolsa' : 'Vendiendo en bolsa',
    }))
  }, [wizPosicionMensual, wizPPPResumen])

  const wizPBPromedioGlobal = useMemo(() => {
    if (wizPBDatos.length === 0) return null
    return wizPBDatos.reduce((s, d) => s + d.precio_bolsa, 0) / wizPBDatos.length
  }, [wizPBDatos])

  const wizPBResumenMensual = useMemo(() => {
    if (wizPBDatos.length === 0) return []
    const mapaPB = new Map<string, number[]>()
    for (const d of wizPBDatos) { const mes = normalizarFecha(d.fecha).slice(0, 7); const arr = mapaPB.get(mes) ?? []; arr.push(d.precio_bolsa); mapaPB.set(mes, arr) }
    const invAnio = new Date(wizPeriodoInicio).getFullYear()
    const mapaPort = new Map<string, number>()
    for (const m of wizResumenMensual) mapaPort.set(m.mes, m.posNetaMwh)
    return [...mapaPB.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([pbMes, pbs]) => {
      const pbProm = pbs.reduce((s, v) => s + v, 0) / pbs.length
      const portMes = `${invAnio}-${pbMes.slice(5)}`
      const posNetaMwh = mapaPort.get(portMes) ?? 0
      // Usar PPP real de compra R si disponible, mock como fallback
      const pppRef = wizPPPResumen?.compra_r?.ppp ?? PRECIO_MOCK_COMPRA_R
      return { mes: pbMes, pbProm, posNetaMwh, transaccionMcop: posNetaMwh * 1000 * pbProm / 1_000_000, spreadVsPB: pppRef - pbProm }
    })
  }, [wizPBDatos, wizResumenMensual, wizPeriodoInicio, wizPPPResumen])

  const wizPBPerfilHorario = useMemo(() => {
    if (wizPBDatos.length === 0) return []
    return Array.from({ length: 24 }, (_, i) => {
      const hora = i + 1
      const pbs = wizPBDatos.filter(d => d.hora === hora).map(d => d.precio_bolsa)
      if (pbs.length === 0) return { hora, p10: null as number | null, p50: null as number | null, p90: null as number | null, promedio: null as number | null }
      return { hora, p10: percentil(pbs, 10), p50: percentil(pbs, 50), p90: percentil(pbs, 90), promedio: pbs.reduce((s, v) => s + v, 0) / pbs.length }
    })
  }, [wizPBDatos])

  const datosPorHora = useMemo(
    () =>
      Array.from({ length: 24 }, (_, i) => {
        const hora = i + 1
        const filas = datosConSpread.filter((d) => d.hora === hora)
        const avg_pb =
          filas.length > 0
            ? filas.reduce((s, d) => s + d.precio_bolsa, 0) / filas.length
            : 0
        const avg_spread =
          filas.length > 0
            ? filas.reduce((s, d) => s + d.spread, 0) / filas.length
            : 0
        return {
          hora,
          precio_bolsa: +avg_pb.toFixed(2),
          spread: +avg_spread.toFixed(2),
          precio_contrato: precioContrato,
        }
      }),
    [datosConSpread, precioContrato],
  )

  const filasDia = useMemo(
    () =>
      filasPorHoraDelDia(datosConSpread, fechaSeleccionada, precioContrato),
    [datosConSpread, fechaSeleccionada, precioContrato],
  )

  const filasMes = useMemo(
    () => filasPromedioMes(datosConSpread, mesSeleccionado, precioContrato),
    [datosConSpread, mesSeleccionado, precioContrato],
  )

  const resumenDia = useMemo(() => {
    const conDatos = filasDia.filter((f) => f.precio_bolsa !== null)
    if (conDatos.length === 0) return null
    const pico = conDatos.reduce((a, b) =>
      (a.precio_bolsa ?? 0) >= (b.precio_bolsa ?? 0) ? a : b,
    )
    const valle = conDatos.reduce((a, b) =>
      (a.precio_bolsa ?? Infinity) <= (b.precio_bolsa ?? Infinity) ? a : b,
    )
    const horasCriticas = conDatos.filter((f) => {
      if (f.spread === null) return false
      return posicion === 'vendedor' ? f.spread < 0 : f.spread > 0
    }).length
    return { pico, valle, horasCriticas }
  }, [filasDia, posicion])

  const resumenMes = useMemo(() => {
    const delMes = datosConSpread.filter(
      (d) => mesDeFecha(d.fecha) === mesSeleccionado,
    )
    if (delMes.length === 0) return null
    const spreadPromedio =
      delMes.reduce((s, d) => s + d.spread, 0) / delMes.length
    const criticas =
      posicion === 'vendedor'
        ? delMes.filter((d) => d.spread < 0).length
        : delMes.filter((d) => d.spread > 0).length
    const pctCritico = (criticas / delMes.length) * 100
    const filasHora = filasMes.filter((f) => f.precio_bolsa !== null)
    const horaMasCara =
      filasHora.length > 0
        ? filasHora.reduce((a, b) =>
            (a.precio_bolsa ?? 0) >= (b.precio_bolsa ?? 0) ? a : b,
          )
        : null
    return { spreadPromedio, pctCritico, horaMasCara }
  }, [datosConSpread, mesSeleccionado, filasMes, posicion])

  const hayDatosFecha =
    fechaSeleccionada !== '' &&
    datosConSpread.some(
      (d) => normalizarFecha(d.fecha) === fechaSeleccionada,
    )

  const hayDatosMes =
    mesSeleccionado !== '' &&
    datosConSpread.some((d) => mesDeFecha(d.fecha) === mesSeleccionado)

  const seriesComparativo = useMemo(
    () =>
      mesesComparativo.map((mes) =>
        pbPromedioPorHoraMes(datosConSpread, mes),
      ),
    [datosConSpread, mesesComparativo],
  )

  const labelPromedio =
    posicion === 'vendedor' ? 'Spread promedio' : 'Exposición promedio'
  const labelHorasCriticas =
    posicion === 'vendedor' ? 'Horas en riesgo' : 'Horas caras'
  const labelHorasDia =
    posicion === 'vendedor' ? 'Horas en riesgo del día' : 'Horas caras del día'
  const labelPctMes =
    posicion === 'vendedor' ? '% horas en riesgo' : '% horas caras'
  const labelSpreadMes =
    posicion === 'vendedor'
      ? 'Spread promedio del mes'
      : 'Exposición promedio del mes'

  const portfolioResumen = useMemo<PortfolioResumen | null>(() => {
    if (portfolioDatos.length === 0) return null
    const posicion_neta_total_kwh = portfolioDatos.reduce((sum, f) => sum + f.posicion_neta, 0)
    const demanda_r_total_kwh    = portfolioDatos.reduce((sum, f) => sum + f.demanda_r,    0)
    const demanda_nr_total_kwh   = portfolioDatos.reduce((sum, f) => sum + f.demanda_nr,   0)
    const costo_bolsa_total_cop  = portfolioDatos.reduce((sum, f) => sum + f.costo_bolsa,  0)
    const porHora = Array.from({ length: 24 }, (_, i) => i + 1).map((hora) => {
      const filas = portfolioDatos.filter((f) => f.hora === hora)
      if (filas.length === 0) return { hora, compra: 0, venta: 0 }
      return {
        hora,
        compra: filas.reduce((s, f) => s + f.compra_r + f.compra_nr, 0) / filas.length,
        venta:  filas.reduce((s, f) => s + f.venta, 0) / filas.length,
      }
    })
    const horaPicoCompra = porHora.reduce((a, b) => (a.compra >= b.compra ? a : b))
    const horaPicoVenta  = porHora.reduce((a, b) => (a.venta  >= b.venta  ? a : b))
    return {
      posicion_neta_total_kwh,
      demanda_r_total_kwh,
      demanda_nr_total_kwh,
      costo_bolsa_total_cop,
      hora_pico_compra: horaPicoCompra.compra > 0 ? horaPicoCompra.hora : null,
      hora_pico_venta:  horaPicoVenta.venta   > 0 ? horaPicoVenta.hora  : null,
    }
  }, [portfolioDatos])

  const portfolioPorHora = useMemo(() => {
    return Array.from({ length: 24 }, (_, i) => {
      const hora = i + 1
      const filas = portfolioDatos.filter((f) => f.hora === hora)
      if (filas.length === 0) {
        return { hora, compra_r: 0, compra_nr: 0, venta: 0, demanda_r: 0, demanda_nr: 0, posicion_neta: 0 }
      }
      // Promedio por hora en kWh (sin dividir por 1000)
      const n = filas.length
      return {
        hora,
        compra_r:     filas.reduce((s, f) => s + f.compra_r,    0) / n,
        compra_nr:    filas.reduce((s, f) => s + f.compra_nr,   0) / n,
        venta:        filas.reduce((s, f) => s + f.venta,       0) / n,
        demanda_r:    filas.reduce((s, f) => s + f.demanda_r,   0) / n,
        demanda_nr:   filas.reduce((s, f) => s + f.demanda_nr,  0) / n,
        posicion_neta:filas.reduce((s, f) => s + f.posicion_neta, 0) / n,
      }
    })
  }, [portfolioDatos])

  const portfolioResumenDiario = useMemo(() => {
    const mapa = new Map<string, FilaPortfolio[]>()
    for (const fila of portfolioDatos) {
      const fecha = normalizarFecha(fila.fecha)
      const acumulado = mapa.get(fecha) ?? []
      acumulado.push(fila)
      mapa.set(fecha, acumulado)
    }
    return [...mapa.entries()]
      .map(([fecha, filas]) => {
        // Acumular en kWh (sin dividir por 1000)
        const compraR     = filas.reduce((s, f) => s + f.compra_r,     0)
        const compraNr    = filas.reduce((s, f) => s + f.compra_nr,    0)
        const venta       = filas.reduce((s, f) => s + f.venta,        0)
        const demandaR    = filas.reduce((s, f) => s + f.demanda_r,    0)
        const demandaNr   = filas.reduce((s, f) => s + f.demanda_nr,   0)
        const posicionNeta= filas.reduce((s, f) => s + f.posicion_neta, 0)
        const costoBolsa  = filas.reduce((s, f) => s + f.costo_bolsa,  0) / 1_000_000 // en M COP
        // Usar tipo_dia del dato (ya clasificado por el backend); fallback local
        const tipoDiaRaw  = filas[0]?.tipo_dia ?? ''
        const tipoDia = tipoDiaRaw === 'ordinario' ? 'Ordinario'
                      : tipoDiaRaw === 'sabado'    ? 'Sábado'
                      : tipoDiaRaw === 'domingo'   ? 'Domingo'
                      : tipoDiaRaw === 'festivo'   ? 'Festivo'
                      : tipoDiaDesdefecha(fecha)
        return { fecha, tipoDia, compraR, compraNr, venta, demandaR, demandaNr, posicionNeta, costoBolsa }
      })
      .sort((a, b) => a.fecha.localeCompare(b.fecha))
  }, [portfolioDatos])

  const portfolioMesesDisponibles = useMemo(() => {
    const meses = new Set<string>()
    for (const fila of portfolioResumenDiario) meses.add(fila.fecha.slice(0, 7))
    return [...meses].sort()
  }, [portfolioResumenDiario])

  const portfolioResumenDiarioFiltrado = useMemo(() => {
    return portfolioResumenDiario.filter((fila) => {
      const mesOk = !portfolioFiltroMes || fila.fecha.startsWith(portfolioFiltroMes)
      const tipoDiaOk =
        portfolioFiltroTipoDia === 'todos' || fila.tipoDia === portfolioFiltroTipoDia
      return mesOk && tipoDiaOk
    })
  }, [portfolioResumenDiario, portfolioFiltroMes, portfolioFiltroTipoDia])

  const portfolioTotalesFiltrados = useMemo(() => {
    if (portfolioResumenDiarioFiltrado.length === 0) return null
    return portfolioResumenDiarioFiltrado.reduce(
      (acc, f) => ({
        compraR:      acc.compraR      + f.compraR,
        compraNr:     acc.compraNr     + f.compraNr,
        venta:        acc.venta        + f.venta,
        demandaR:     acc.demandaR     + f.demandaR,
        demandaNr:    acc.demandaNr    + f.demandaNr,
        posicionNeta: acc.posicionNeta + f.posicionNeta,
        costoBolsa:   acc.costoBolsa   + f.costoBolsa,
      }),
      { compraR: 0, compraNr: 0, venta: 0, demandaR: 0, demandaNr: 0, posicionNeta: 0, costoBolsa: 0 },
    )
  }, [portfolioResumenDiarioFiltrado])

  const perfilHorarioPorTipoDia = useMemo(() => {
    const filtradas = portfolioFiltroMes
      ? portfolioDatos.filter((f) => normalizarFecha(f.fecha).startsWith(portfolioFiltroMes))
      : portfolioDatos

    const init24 = () =>
      Array.from({ length: 24 }, () => ({ compra: [] as number[], venta: [] as number[] }))

    const acum = { ordinario: init24(), sabado: init24(), domfest: init24() }

    for (const fila of filtradas) {
      const tipoRaw = fila.tipo_dia ?? ''
      const tipo =
        tipoRaw === 'ordinario' ? 'Ordinario'
        : tipoRaw === 'sabado'  ? 'Sábado'
        : tipoRaw === 'domingo' ? 'Domingo'
        : tipoRaw === 'festivo' ? 'Festivo'
        : tipoDiaDesdefecha(fila.fecha)
      const balde =
        tipo === 'Sábado'
          ? acum.sabado
          : tipo === 'Domingo' || tipo === 'Festivo'
            ? acum.domfest
            : acum.ordinario
      const idx = fila.hora - 1
      if (idx < 0 || idx > 23) continue
      balde[idx].compra.push((fila.compra_r + fila.compra_nr) / 1000)
      balde[idx].venta.push(fila.venta / 1000)
    }

    const avg = (arr: number[]) =>
      arr.length > 0 ? arr.reduce((s, v) => s + v, 0) / arr.length : null

    return Array.from({ length: 24 }, (_, i) => ({
      hora: i + 1,
      ordinario_compra: avg(acum.ordinario[i].compra),
      ordinario_venta: avg(acum.ordinario[i].venta),
      sabado_compra: avg(acum.sabado[i].compra),
      sabado_venta: avg(acum.sabado[i].venta),
      domfest_compra: avg(acum.domfest[i].compra),
      domfest_venta: avg(acum.domfest[i].venta),
    }))
  }, [portfolioDatos, portfolioFiltroMes])

  useEffect(() => {
    if (portfolioMesesDisponibles.length === 0) return
    setPortfolioFiltroMes((prev) =>
      prev && portfolioMesesDisponibles.includes(prev)
        ? prev
        : portfolioMesesDisponibles[portfolioMesesDisponibles.length - 1],
    )
  }, [portfolioMesesDisponibles])

  if (cargando) {
    return (
      <div className="flex min-h-screen items-center justify-center bg-gray-950">
        <div className="text-center">
          <div className="mx-auto mb-4 h-10 w-10 animate-spin rounded-full border-2 border-emerald-500 border-t-transparent" />
          <p className="text-sm text-gray-400">Cargando datos de spread…</p>
        </div>
      </div>
    )
  }

  if (error || !resumen) {
    return (
      <div className="flex min-h-screen items-center justify-center bg-gray-950 px-4">
        <div className="max-w-md rounded-lg border border-red-900/50 bg-gray-900 p-6 text-center">
          <p className="text-red-400">Error al cargar el dashboard</p>
          <p className="mt-2 text-sm text-gray-500">{error ?? 'Sin datos'}</p>
        </div>
      </div>
    )
  }

  return (
    <div className="min-h-screen bg-gray-950 text-gray-100">
      <header className="border-b border-gray-800 bg-black px-6 py-5">
        <div className="mx-auto flex max-w-7xl flex-wrap items-start justify-between gap-4">
          <div>
            <h1 className="text-2xl font-bold tracking-tight text-white">
              Spread Bolsa BIA
            </h1>
            <p className="mt-1 text-sm text-gray-400">
              Motor de análisis PB vs contratos
            </p>
          </div>
          <span className="inline-flex items-center gap-2 rounded-full bg-emerald-500/15 px-3 py-1 text-xs font-semibold uppercase tracking-wider text-emerald-400 ring-1 ring-emerald-500/40">
            <span className="h-2 w-2 animate-pulse rounded-full bg-emerald-400" />
            EN VIVO
          </span>
        </div>
      </header>

      <main className="mx-auto max-w-7xl space-y-8 px-6 py-8">
        <section className="rounded-xl border border-gray-800 bg-gray-900/60 p-2">
          <div className="inline-flex rounded-lg border border-gray-700 bg-gray-950 p-1">
            {(
              [
                ['spread', 'Mercado'],
                ['portafolio', 'Portafolio'],
                ['simulador', 'Simulador'],
              ] as const
            ).map(([id, label]) => (
              <button
                key={id}
                type="button"
                onClick={() => setTabActiva(id)}
                className={`rounded-md px-5 py-2 text-sm font-semibold tracking-wide transition ${
                  tabActiva === id
                    ? 'bg-emerald-600 text-white shadow-sm shadow-emerald-900/50'
                    : 'text-gray-400 hover:bg-gray-800 hover:text-gray-200'
                }`}
              >
                {label}
              </button>
            ))}
          </div>
        </section>

        {tabActiva === 'spread' && (
          <>
            <div>
              <h2 className="text-xl font-bold text-gray-100">Mercado</h2>
              <p className="mt-0.5 text-sm text-gray-500">
                Análisis de comportamiento del precio de bolsa
              </p>
            </div>
            <section className="rounded-xl border border-gray-800 bg-gray-900/60 p-6">
          <div className="flex flex-wrap items-end gap-4">
            <div className="flex min-w-[200px] flex-1 flex-col gap-2 sm:max-w-xs">
              <label
                htmlFor="precio-contrato"
                className="text-sm font-medium text-gray-300"
              >
                Precio de contrato (COP/kWh)
              </label>
              <input
                id="precio-contrato"
                type="number"
                min={0}
                value={precioContrato}
                onChange={(e) =>
                  setPrecioContrato(Math.max(0, Number(e.target.value) || 0))
                }
                className="w-full rounded-lg border border-gray-700 bg-gray-950 px-4 py-2.5 text-gray-100 outline-none ring-emerald-500/0 transition focus:border-emerald-500/50 focus:ring-2 focus:ring-emerald-500/30"
              />
            </div>
            <div className="flex flex-col gap-2">
              <span className="text-sm font-medium text-gray-300">
                Posición
              </span>
              <div className="inline-flex gap-1 rounded-full p-1">
                {(
                  [
                    ['vendedor', 'Vendedor'],
                    ['comprador', 'Comprador'],
                  ] as const
                ).map(([id, label]) => (
                  <button
                    key={id}
                    type="button"
                    onClick={() => setPosicion(id)}
                    className={`rounded-full px-4 py-2 text-sm font-semibold transition ${
                      posicion === id
                        ? 'bg-emerald-600 text-white'
                        : 'bg-gray-800 text-gray-300 hover:bg-gray-700'
                    }`}
                  >
                    {label}
                  </button>
                ))}
              </div>
            </div>
            <div className="flex min-w-[160px] flex-col gap-2">
              <label
                htmlFor="fecha-inicio"
                className="text-sm font-medium text-gray-300"
              >
                Desde
              </label>
              <input
                id="fecha-inicio"
                type="date"
                min="2010-01-01"
                value={fechaInicio}
                onChange={(e) => setFechaInicio(e.target.value)}
                className="w-full rounded-lg border border-gray-700 bg-gray-950 px-4 py-2.5 text-gray-100 outline-none transition [color-scheme:dark] focus:border-emerald-500/50 focus:ring-2 focus:ring-emerald-500/30"
              />
            </div>
            <div className="flex min-w-[160px] flex-col gap-2">
              <label
                htmlFor="fecha-fin"
                className="text-sm font-medium text-gray-300"
              >
                Hasta
              </label>
              <input
                id="fecha-fin"
                type="date"
                value={fechaFin}
                onChange={(e) => setFechaFin(e.target.value)}
                className="w-full rounded-lg border border-gray-700 bg-gray-950 px-4 py-2.5 text-gray-100 outline-none transition [color-scheme:dark] focus:border-emerald-500/50 focus:ring-2 focus:ring-emerald-500/30"
              />
            </div>
            <div className="flex items-center gap-3">
              <button
                type="button"
                onClick={handleCalcular}
                disabled={recalculando}
                className="rounded-lg bg-emerald-600 px-5 py-2.5 text-sm font-semibold text-white transition hover:bg-emerald-500 disabled:cursor-not-allowed disabled:opacity-60"
              >
                Calcular
              </button>
              {recalculando && (
                <span className="text-sm text-gray-400">Calculando...</span>
              )}
            </div>
          </div>
        </section>

        <section className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-4">
          <MetricCard
            label={labelPromedio}
            value={
              resumenLocal
                ? formatNumero(resumenLocal.spread_promedio)
                : '—'
            }
            accent="text-white"
          />
          <MetricCard
            label="Spread mínimo"
            value={
              resumenLocal ? formatNumero(resumenLocal.spread_minimo) : '—'
            }
            accent="text-sky-400"
          />
          <MetricCard
            label="Spread máximo"
            value={
              resumenLocal ? formatNumero(resumenLocal.spread_maximo) : '—'
            }
            accent="text-emerald-400"
          />
          <MetricCard
            label={labelHorasCriticas}
            value={
              resumenLocal ? `${resumenLocal.horas_criticas}` : '—'
            }
            subValue={
              resumenLocal
                ? `${formatNumero(resumenLocal.porcentaje_critico)}%`
                : undefined
            }
            accent="text-amber-400"
          />
        </section>

        <section className="rounded-xl border border-gray-800 bg-gray-900/60 p-6">
          <h2 className="mb-6 text-lg font-semibold text-gray-200">
            Perfil horario promedio — PB vs Contrato vs Spread
          </h2>
          <div className="h-80 w-full">
            <ResponsiveContainer width="100%" height="100%">
              <LineChart
                data={datosPorHora}
                margin={{ top: 8, right: 16, left: 0, bottom: 8 }}
              >
                <CartesianGrid stroke="#374151" strokeDasharray="3 3" />
                <XAxis
                  dataKey="hora"
                  type="number"
                  domain={[1, 24]}
                  ticks={Array.from({ length: 24 }, (_, i) => i + 1)}
                  stroke="#9ca3af"
                  tick={{ fill: '#9ca3af', fontSize: 12 }}
                  label={{
                    value: 'Hora',
                    position: 'insideBottom',
                    offset: -4,
                    fill: '#6b7280',
                  }}
                />
                <YAxis
                  stroke="#9ca3af"
                  tick={{ fill: '#9ca3af', fontSize: 12 }}
                />
                <Tooltip
                  contentStyle={{
                    backgroundColor: '#111827',
                    border: '1px solid #374151',
                    borderRadius: '8px',
                    color: '#f3f4f6',
                  }}
                  labelFormatter={(hora) => `Hora: ${hora}`}
                />
                <Legend wrapperStyle={{ color: '#9ca3af', paddingTop: 12 }} />
                <Line
                  type="monotone"
                  dataKey="precio_bolsa"
                  name="Precio Bolsa promedio"
                  stroke="#3b82f6"
                  strokeWidth={2}
                  dot={false}
                  activeDot={{ r: 4 }}
                />
                <Line
                  type="monotone"
                  dataKey="precio_contrato"
                  name="Precio Contrato"
                  stroke="#eab308"
                  strokeWidth={2}
                  strokeDasharray="6 4"
                  dot={false}
                  activeDot={{ r: 4 }}
                />
                <Line
                  type="monotone"
                  dataKey="spread"
                  name={labelPromedio}
                  stroke="#22c55e"
                  strokeWidth={2}
                  dot={false}
                  activeDot={{ r: 4 }}
                />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </section>

        {/* Análisis horario interactivo */}
        <section className="overflow-hidden rounded-xl border border-gray-800">
          <div className="border-b border-gray-800 bg-gray-900/80 px-6 py-4">
            <h2 className="text-lg font-semibold text-gray-200">
              Análisis horario
            </h2>
            <p className="mt-1 text-sm text-gray-500">
              Perfil por día, promedios mensuales y comparación estacional
            </p>
          </div>

          <div className="border-b border-gray-800 bg-gray-950/80 px-6 py-4">
            <div
              className="inline-flex rounded-lg border border-gray-700 bg-gray-900 p-1"
              role="tablist"
            >
              {(
                [
                  ['dia', 'Por Día'],
                  ['mes', 'Promedio Mes'],
                  ['comparativo', 'Comparativo'],
                ] as const
              ).map(([id, label]) => (
                <button
                  key={id}
                  type="button"
                  role="tab"
                  aria-selected={vistaAnalisis === id}
                  onClick={() => setVistaAnalisis(id)}
                  className={`rounded-md px-4 py-2 text-sm font-semibold tracking-wide transition ${
                    vistaAnalisis === id
                      ? 'bg-emerald-600 text-white shadow-sm shadow-emerald-900/50'
                      : 'text-gray-400 hover:bg-gray-800 hover:text-gray-200'
                  }`}
                >
                  {label}
                </button>
              ))}
            </div>
          </div>

          <div className="bg-gray-900/40 px-6 py-5">
            {vistaAnalisis === 'dia' && (
              <div className="space-y-5">
                <div className="flex flex-wrap items-end gap-4">
                  <div className="flex min-w-[240px] flex-col gap-2 sm:max-w-md">
                    <label
                      htmlFor="fecha-dia"
                      className="text-sm font-medium text-gray-300"
                    >
                      Fecha
                    </label>
                    <select
                      id="fecha-dia"
                      value={fechaSeleccionada}
                      onChange={(e) => setFechaSeleccionada(e.target.value)}
                      className={CLASE_SELECT}
                    >
                      {fechasDisponibles.map((f) => (
                        <option key={f} value={f}>
                          {formatFechaLegible(f)}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                {!hayDatosFecha ? (
                  <MensajeSinDatos>
                    No hay datos para la fecha{' '}
                    <span className="font-medium text-gray-300">
                      {formatFechaLegible(fechaSeleccionada)}
                    </span>
                    .
                  </MensajeSinDatos>
                ) : (
                  <>
                    <TablaAnalisis
                      encabezados={[
                        'Hora',
                        'Precio Bolsa',
                        'Precio Contrato',
                        'Spread',
                      ]}
                      filas={filasDia.map((f) => (
                        <tr
                          key={f.hora}
                          className="transition-colors hover:bg-gray-800/40"
                        >
                          <td className="px-4 py-2.5 text-gray-300">{f.hora}</td>
                          <td className="px-4 py-2.5 text-right tabular-nums text-gray-200">
                            {f.precio_bolsa !== null
                              ? formatNumero(f.precio_bolsa)
                              : '—'}
                          </td>
                          <td className="px-4 py-2.5 text-right tabular-nums text-amber-400/90">
                            {formatNumero(f.precio_contrato)}
                          </td>
                          <td
                            className={`px-4 py-2.5 text-right font-medium tabular-nums ${claseSpread(f.spread, posicion)}`}
                          >
                            {f.spread !== null ? formatNumero(f.spread) : '—'}
                          </td>
                        </tr>
                      ))}
                    />
                    {resumenDia && (
                      <ResumenBloque
                        items={[
                          {
                            label: 'Hora pico (máx PB)',
                            value: `H${resumenDia.pico.hora} — ${formatNumero(resumenDia.pico.precio_bolsa!)}`,
                          },
                          {
                            label: 'Hora valle (mín PB)',
                            value: `H${resumenDia.valle.hora} — ${formatNumero(resumenDia.valle.precio_bolsa!)}`,
                          },
                          {
                            label: labelHorasDia,
                            value: String(resumenDia.horasCriticas),
                            accent: 'text-amber-400',
                          },
                        ]}
                      />
                    )}
                  </>
                )}
              </div>
            )}

            {vistaAnalisis === 'mes' && (
              <div className="space-y-5">
                <div className="flex flex-col gap-2 sm:max-w-xs">
                  <label
                    htmlFor="mes-promedio"
                    className="text-sm font-medium text-gray-300"
                  >
                    Mes
                  </label>
                  <select
                    id="mes-promedio"
                    value={mesSeleccionado}
                    onChange={(e) => setMesSeleccionado(e.target.value)}
                    className={CLASE_SELECT}
                  >
                    {mesesDisponibles.map((m) => (
                      <option key={m} value={m}>
                        {formatMes(m)}
                      </option>
                    ))}
                  </select>
                </div>

                {!hayDatosMes ? (
                  <MensajeSinDatos>
                    No hay datos para{' '}
                    <span className="font-medium text-gray-300">
                      {formatMes(mesSeleccionado)}
                    </span>
                    .
                  </MensajeSinDatos>
                ) : (
                  <>
                    <TablaAnalisis
                      encabezados={[
                        'Hora',
                        'Precio Bolsa (prom.)',
                        'Precio Contrato',
                        'Spread (prom.)',
                        'P10 | P50 | P90',
                      ]}
                      filas={filasMes.map((f) => (
                        <tr
                          key={f.hora}
                          className="transition-colors hover:bg-gray-800/40"
                        >
                          <td className="px-4 py-2.5 text-gray-300">{f.hora}</td>
                          <td className="px-4 py-2.5 text-right tabular-nums text-gray-200">
                            {f.precio_bolsa !== null
                              ? formatNumero(f.precio_bolsa)
                              : '—'}
                          </td>
                          <td className="px-4 py-2.5 text-right tabular-nums text-amber-400/90">
                            {formatNumero(f.precio_contrato)}
                          </td>
                          <td
                            className={`px-4 py-2.5 text-right font-medium tabular-nums ${claseSpread(f.spread, posicion)}`}
                          >
                            {f.spread !== null ? formatNumero(f.spread) : '—'}
                          </td>
                          <td className="px-4 py-2.5 text-right font-mono text-xs tabular-nums text-sky-300/90">
                            {f.percentiles ? (
                              <>
                                {formatNumero(f.percentiles.p10)}
                                <span className="text-gray-600"> | </span>
                                {formatNumero(f.percentiles.p50)}
                                <span className="text-gray-600"> | </span>
                                {formatNumero(f.percentiles.p90)}
                              </>
                            ) : (
                              '—'
                            )}
                          </td>
                        </tr>
                      ))}
                    />
                    {resumenMes && (
                      <ResumenBloque
                        items={[
                          {
                            label: labelSpreadMes,
                            value: formatNumero(resumenMes.spreadPromedio),
                          },
                          {
                            label: labelPctMes,
                            value: `${formatNumero(resumenMes.pctCritico)}%`,
                            accent: 'text-amber-400',
                          },
                          {
                            label: 'Hora más cara del mes (prom. PB)',
                            value: resumenMes.horaMasCara
                              ? `H${resumenMes.horaMasCara.hora} — ${formatNumero(resumenMes.horaMasCara.precio_bolsa!)}`
                              : '—',
                          },
                        ]}
                      />
                    )}
                  </>
                )}
              </div>
            )}

            {vistaAnalisis === 'comparativo' && (
              <div className="space-y-5">
                <div>
                  <p className="mb-3 text-sm font-medium text-gray-300">
                    Meses a comparar (máx. 4)
                  </p>
                  <div className="flex flex-wrap gap-2">
                    {mesesDisponibles.map((mes) => {
                      const seleccionado = mesesComparativo.includes(mes)
                      const indice = mesesComparativo.indexOf(mes)
                      const color =
                        indice >= 0
                          ? COLORES_MES_COMPARATIVO[indice]
                          : null
                      const deshabilitado =
                        !seleccionado && mesesComparativo.length >= 4
                      return (
                        <button
                          key={mes}
                          type="button"
                          onClick={() => toggleMesComparativo(mes)}
                          disabled={deshabilitado}
                          className={`rounded-full px-3 py-1.5 text-sm font-semibold ring-1 transition ${
                            seleccionado && color
                              ? `${color.bg} ${color.text} ${color.ring}`
                              : deshabilitado
                                ? 'cursor-not-allowed bg-gray-900 text-gray-600 ring-gray-800'
                                : 'bg-gray-900 text-gray-400 ring-gray-700 hover:bg-gray-800 hover:text-gray-200'
                          }`}
                        >
                          {formatMes(mes)}
                        </button>
                      )
                    })}
                  </div>
                </div>

                {mesesComparativo.length === 0 ? (
                  <MensajeSinDatos>
                    Selecciona al menos un mes para comparar.
                  </MensajeSinDatos>
                ) : (
                  <TablaAnalisis
                    encabezados={[
                      'Hora',
                      ...mesesComparativo.map((m) => `PB ${formatMes(m)}`),
                    ]}
                    encabezadosClassName={mesesComparativo.map((_, i) =>
                      i < COLORES_MES_COMPARATIVO.length
                        ? COLORES_MES_COMPARATIVO[i].text
                        : '',
                    )}
                    filas={Array.from({ length: 24 }, (_, i) => {
                      const hora = i + 1
                      return (
                        <tr
                          key={hora}
                          className="transition-colors hover:bg-gray-800/40"
                        >
                          <td className="px-4 py-2.5 text-gray-300">{hora}</td>
                          {seriesComparativo.map((serie, j) => {
                            const celda = serie.find((r) => r.hora === hora)
                            return (
                              <td
                                key={`${hora}-${mesesComparativo[j]}`}
                                className={`px-4 py-2.5 text-right font-medium tabular-nums ${
                                  COLORES_MES_COMPARATIVO[j]?.text ??
                                  'text-gray-200'
                                }`}
                              >
                                {celda?.pb != null
                                  ? formatNumero(celda.pb)
                                  : '—'}
                              </td>
                            )
                          })}
                        </tr>
                      )
                    })}
                  />
                )}
              </div>
            )}
          </div>
            </section>
          </>
        )}

        {tabActiva === 'portafolio' && (
          <>
            <section className="rounded-xl border border-gray-800 bg-gray-900/60 p-6">
              <div className="flex flex-wrap items-end gap-4">
                <div className="flex min-w-[160px] flex-col gap-2">
                  <label
                    htmlFor="portfolio-fecha-inicio"
                    className="text-sm font-medium text-gray-300"
                  >
                    Desde
                  </label>
                  <input
                    id="portfolio-fecha-inicio"
                    type="date"
                    min="2010-01-01"
                    value={portfolioFechaInicio}
                    onChange={(e) => setPortfolioFechaInicio(e.target.value)}
                    className="w-full rounded-lg border border-gray-700 bg-gray-950 px-4 py-2.5 text-gray-100 outline-none transition [color-scheme:dark] focus:border-emerald-500/50 focus:ring-2 focus:ring-emerald-500/30"
                  />
                </div>
                <div className="flex min-w-[160px] flex-col gap-2">
                  <label
                    htmlFor="portfolio-fecha-fin"
                    className="text-sm font-medium text-gray-300"
                  >
                    Hasta
                  </label>
                  <input
                    id="portfolio-fecha-fin"
                    type="date"
                    value={portfolioFechaFin}
                    onChange={(e) => setPortfolioFechaFin(e.target.value)}
                    className="w-full rounded-lg border border-gray-700 bg-gray-950 px-4 py-2.5 text-gray-100 outline-none transition [color-scheme:dark] focus:border-emerald-500/50 focus:ring-2 focus:ring-emerald-500/30"
                  />
                </div>
                <div className="flex items-center gap-3">
                  <button
                    type="button"
                    onClick={handleCalcularPortafolio}
                    disabled={portfolioCargando}
                    className="rounded-lg bg-emerald-600 px-5 py-2.5 text-sm font-semibold text-white transition hover:bg-emerald-500 disabled:cursor-not-allowed disabled:opacity-60"
                  >
                    Calcular
                  </button>
                  {portfolioCargando && (
                    <span className="text-sm text-gray-400">Calculando...</span>
                  )}
                </div>
              </div>
              {portfolioError && (
                <p className="mt-4 text-sm text-red-400">{portfolioError}</p>
              )}
            </section>

            <section className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6">
              <MetricCard
                label={portfolioResumen
                  ? (portfolioResumen.posicion_neta_total_kwh > 0
                      ? 'Comprando en bolsa (kWh)'
                      : 'Vendiendo en bolsa (kWh)')
                  : 'Posición Neta (kWh)'}
                value={portfolioResumen ? formatMiles(portfolioResumen.posicion_neta_total_kwh) : '—'}
                accent={
                  portfolioResumen && portfolioResumen.posicion_neta_total_kwh > 0
                    ? 'text-red-400'
                    : 'text-emerald-400'
                }
              />
              <MetricCard
                label="Total Demanda R (kWh)"
                value={portfolioResumen ? formatMiles(portfolioResumen.demanda_r_total_kwh) : '—'}
                accent="text-orange-400"
              />
              <MetricCard
                label="Total Demanda NR (kWh)"
                value={portfolioResumen ? formatMiles(portfolioResumen.demanda_nr_total_kwh) : '—'}
                accent="text-purple-400"
              />
              <MetricCard
                label="Hora Pico Compra"
                value={portfolioResumen?.hora_pico_compra ? `H${portfolioResumen.hora_pico_compra}` : '—'}
                accent="text-sky-400"
              />
              <MetricCard
                label="Hora Pico Venta"
                value={portfolioResumen?.hora_pico_venta ? `H${portfolioResumen.hora_pico_venta}` : '—'}
                accent="text-rose-400"
              />
              <div className="rounded-xl border border-gray-800 bg-gray-900 p-5">
                <p className="text-xs font-medium uppercase tracking-wider text-gray-500">PPP Contratos PC</p>
                {portfolioPPP ? (
                  <div className="mt-2 space-y-2">
                    {(
                      [
                        { label: 'Compra R',  cat: portfolioPPP.compra_r  },
                        { label: 'Compra NR', cat: portfolioPPP.compra_nr },
                        { label: 'Venta',     cat: portfolioPPP.venta     },
                      ] as const
                    ).map(({ label, cat }) => (
                      <div key={label} className="flex items-center justify-between gap-2">
                        <span className="text-xs text-gray-500">{label}</span>
                        <div className="flex items-center gap-1.5">
                          {cat.ppp != null
                            ? <span className="text-sm font-bold tabular-nums text-white">{formatNumero(cat.ppp)} COP/kWh</span>
                            : <span className="text-sm text-gray-600">—</span>
                          }
                          <BadgeTipoPrecio tipo={cat.tipo} />
                        </div>
                      </div>
                    ))}
                    {portfolioPPP.pld_excluidos > 0 && (
                      <p className="pt-1 text-xs text-amber-600">
                        ⚠ {portfolioPPP.pld_excluidos} contrato(s) PLD excluidos (precio = bolsa)
                      </p>
                    )}
                  </div>
                ) : (
                  <p className="mt-2 text-sm text-amber-500/60">
                    {portfolioCargando ? 'Calculando…' : 'Sin datos de precios'}
                  </p>
                )}
              </div>
            </section>

            <section className="rounded-xl border border-gray-800 bg-gray-900/60 p-6">
              <h2 className="mb-6 text-lg font-semibold text-gray-200">
                Posición horaria del portafolio
              </h2>
              <div className="h-80 w-full">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={portfolioPorHora} margin={{ top: 8, right: 48, left: 0, bottom: 8 }}>
                    <CartesianGrid stroke="#374151" strokeDasharray="3 3" />
                    <XAxis dataKey="hora" stroke="#9ca3af" tick={{ fill: '#9ca3af', fontSize: 12 }} />
                    <YAxis yAxisId="left"  stroke="#9ca3af" tick={{ fill: '#9ca3af', fontSize: 12 }} />
                    <YAxis yAxisId="right" orientation="right" stroke="#eab308" tick={{ fill: '#eab308', fontSize: 12 }} />
                    <Tooltip
                      contentStyle={{ backgroundColor: '#111827', border: '1px solid #374151', borderRadius: '8px', color: '#f3f4f6' }}
                      labelFormatter={(hora) => `Hora: ${hora}`}
                      formatter={(value, name) => [`${formatMiles(Number(value))} kWh`, name]}
                    />
                    <Legend wrapperStyle={{ color: '#9ca3af', paddingTop: 12 }} />
                    <Bar yAxisId="left" dataKey="compra_r"  stackId="pos" name="Compra R"    fill="#3b82f6" />
                    <Bar yAxisId="left" dataKey="compra_nr" stackId="pos" name="Compra NR"   fill="#22c55e" />
                    <Bar yAxisId="left" dataKey="demanda_r"  name="Demanda R"  fill="#f97316" />
                    <Bar yAxisId="left" dataKey="demanda_nr" name="Demanda NR" fill="#a855f7" />
                    <Bar yAxisId="left" dataKey="venta"     name="Venta"       fill="#ef4444" />
                    <Line
                      yAxisId="right"
                      type="monotone"
                      dataKey="posicion_neta"
                      name="Posición Neta"
                      stroke="#eab308"
                      strokeWidth={2}
                      dot={false}
                    />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </section>

            <section className="rounded-xl border border-gray-800 bg-gray-900/60 p-6">
              <h2 className="mb-6 text-lg font-semibold text-gray-200">
                Perfil horario de compras por tipo de día
              </h2>
              <div className="h-80 w-full">
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart
                    data={perfilHorarioPorTipoDia}
                    margin={{ top: 8, right: 16, left: 0, bottom: 8 }}
                  >
                    <CartesianGrid stroke="#374151" strokeDasharray="3 3" />
                    <XAxis
                      dataKey="hora"
                      type="number"
                      domain={[1, 24]}
                      ticks={Array.from({ length: 24 }, (_, i) => i + 1)}
                      stroke="#9ca3af"
                      tick={{ fill: '#9ca3af', fontSize: 12 }}
                    />
                    <YAxis stroke="#9ca3af" tick={{ fill: '#9ca3af', fontSize: 12 }} />
                    <Tooltip
                      contentStyle={{
                        backgroundColor: '#111827',
                        border: '1px solid #374151',
                        borderRadius: '8px',
                        color: '#f3f4f6',
                      }}
                      labelFormatter={(hora) => `Hora: ${hora}`}
                      formatter={(value, name) => [
                        `${formatMiles(Number(value))} MWh`,
                        name,
                      ]}
                    />
                    <Legend wrapperStyle={{ color: '#9ca3af', paddingTop: 12 }} />
                    <Line
                      type="monotone"
                      dataKey="ordinario_compra"
                      name="Ordinario"
                      stroke="#60a5fa"
                      strokeWidth={2}
                      dot={false}
                      connectNulls
                    />
                    <Line
                      type="monotone"
                      dataKey="sabado_compra"
                      name="Sábado"
                      stroke="#22c55e"
                      strokeWidth={2}
                      dot={false}
                      connectNulls
                    />
                    <Line
                      type="monotone"
                      dataKey="domfest_compra"
                      name="Domingo/Festivo"
                      stroke="#eab308"
                      strokeWidth={2}
                      dot={false}
                      connectNulls
                    />
                  </LineChart>
                </ResponsiveContainer>
              </div>
            </section>

            <section className="rounded-xl border border-gray-800 bg-gray-900/60 p-6">
              <h2 className="mb-1 text-lg font-semibold text-gray-200">
                Perfil horario de ventas por tipo de día
              </h2>
              <p className="mb-5 text-xs text-gray-500">
                Contratos de venta a otros agentes — <code className="text-gray-400">venta_kwh</code> (inventario Olibia)
              </p>
              <div className="h-80 w-full">
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart
                    data={perfilHorarioPorTipoDia}
                    margin={{ top: 8, right: 16, left: 0, bottom: 8 }}
                  >
                    <CartesianGrid stroke="#374151" strokeDasharray="3 3" />
                    <XAxis
                      dataKey="hora"
                      type="number"
                      domain={[1, 24]}
                      ticks={Array.from({ length: 24 }, (_, i) => i + 1)}
                      stroke="#9ca3af"
                      tick={{ fill: '#9ca3af', fontSize: 12 }}
                    />
                    <YAxis stroke="#9ca3af" tick={{ fill: '#9ca3af', fontSize: 12 }} />
                    <Tooltip
                      contentStyle={{
                        backgroundColor: '#111827',
                        border: '1px solid #374151',
                        borderRadius: '8px',
                        color: '#f3f4f6',
                      }}
                      labelFormatter={(hora) => `Hora: ${hora}`}
                      formatter={(value, name) => [
                        `${formatMiles(Number(value))} MWh`,
                        name,
                      ]}
                    />
                    <Legend wrapperStyle={{ color: '#9ca3af', paddingTop: 12 }} />
                    <Line
                      type="monotone"
                      dataKey="ordinario_venta"
                      name="Ordinario"
                      stroke="#60a5fa"
                      strokeWidth={2}
                      dot={false}
                      connectNulls
                    />
                    <Line
                      type="monotone"
                      dataKey="sabado_venta"
                      name="Sábado"
                      stroke="#22c55e"
                      strokeWidth={2}
                      dot={false}
                      connectNulls
                    />
                    <Line
                      type="monotone"
                      dataKey="domfest_venta"
                      name="Domingo/Festivo"
                      stroke="#eab308"
                      strokeWidth={2}
                      dot={false}
                      connectNulls
                    />
                  </LineChart>
                </ResponsiveContainer>
              </div>
            </section>

            <section className="overflow-hidden rounded-xl border border-gray-800">
              <div className="border-b border-gray-800 bg-gray-900/80 px-6 py-4">
                <h2 className="text-lg font-semibold text-gray-200">
                  Resumen diario del portafolio
                </h2>
              </div>
              {portfolioResumenDiario.length > 0 && (
                <div className="flex flex-wrap items-end gap-3 border-b border-gray-800 bg-gray-950/60 px-6 py-3">
                  <div className="flex flex-col gap-1">
                    <label className="text-xs font-medium uppercase tracking-wider text-gray-500">
                      Mes
                    </label>
                    <select
                      value={portfolioFiltroMes}
                      onChange={(e) => setPortfolioFiltroMes(e.target.value)}
                      className={CLASE_SELECT}
                    >
                      {portfolioMesesDisponibles.map((m) => (
                        <option key={m} value={m}>
                          {formatMes(m)}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div className="flex flex-col gap-1">
                    <label className="text-xs font-medium uppercase tracking-wider text-gray-500">
                      Tipo de día
                    </label>
                    <select
                      value={portfolioFiltroTipoDia}
                      onChange={(e) => setPortfolioFiltroTipoDia(e.target.value)}
                      className={CLASE_SELECT}
                    >
                      <option value="todos">Todos</option>
                      <option value="Ordinario">Ordinario</option>
                      <option value="Sábado">Sábado</option>
                      <option value="Domingo">Domingo</option>
                      <option value="Festivo">Festivo</option>
                    </select>
                  </div>
                </div>
              )}
              <div className="bg-gray-900/40 px-6 py-5">
                {portfolioResumenDiario.length === 0 ? (
                  <MensajeSinDatos>
                    No hay datos disponibles para el rango seleccionado.
                  </MensajeSinDatos>
                ) : portfolioResumenDiarioFiltrado.length === 0 ? (
                  <MensajeSinDatos>
                    No hay días que coincidan con los filtros seleccionados.
                  </MensajeSinDatos>
                ) : (
                  <TablaAnalisis
                    encabezados={[
                      'Fecha',
                      'Tipo Día',
                      'Compra R (kWh)',
                      'Compra NR (kWh)',
                      'Venta (kWh)',
                      'Demanda R (kWh)',
                      'Demanda NR (kWh)',
                      'Posición Neta (kWh)',
                      'Costo Bolsa (M COP)',
                      'Tipo Posición',
                    ]}
                    filas={
                      <>
                        {portfolioResumenDiarioFiltrado.map((fila) => (
                          <tr key={fila.fecha} className="transition-colors hover:bg-gray-800/40">
                            <td className="px-4 py-2.5 text-gray-300">{formatFechaLegible(fila.fecha)}</td>
                            <td className="px-4 py-2.5 text-right text-gray-300">{fila.tipoDia}</td>
                            <td className="px-4 py-2.5 text-right tabular-nums text-sky-300">{formatMiles(fila.compraR)}</td>
                            <td className="px-4 py-2.5 text-right tabular-nums text-emerald-300">{formatMiles(fila.compraNr)}</td>
                            <td className="px-4 py-2.5 text-right tabular-nums text-rose-300">{formatMiles(fila.venta)}</td>
                            <td className="px-4 py-2.5 text-right tabular-nums text-orange-400">{formatMiles(fila.demandaR)}</td>
                            <td className="px-4 py-2.5 text-right tabular-nums text-purple-400">{formatMiles(fila.demandaNr)}</td>
                            <td className={`px-4 py-2.5 text-right font-medium tabular-nums ${fila.posicionNeta > 0 ? 'text-red-400' : 'text-emerald-400'}`}>
                              {formatMiles(fila.posicionNeta)}
                            </td>
                            <td className="px-4 py-2.5 text-right tabular-nums text-yellow-300">{formatMiles(fila.costoBolsa, 2)}</td>
                            <td className={`px-4 py-2.5 text-right text-xs font-semibold ${fila.posicionNeta > 0 ? 'text-red-400' : 'text-emerald-400'}`}>
                              {fila.posicionNeta > 0 ? 'Comprando en bolsa' : 'Vendiendo en bolsa'}
                            </td>
                          </tr>
                        ))}
                        {portfolioTotalesFiltrados && (() => {
                          const nComp = portfolioResumenDiarioFiltrado.filter(f => f.posicionNeta > 0).length
                          const nVend = portfolioResumenDiarioFiltrado.filter(f => f.posicionNeta <= 0).length
                          return (
                            <tr className="border-t-2 border-gray-600 bg-gray-700">
                              <td className="px-4 py-2.5 text-xs font-bold uppercase tracking-wider text-white">TOTAL PERÍODO</td>
                              <td className="px-4 py-2.5" />
                              <td className="px-4 py-2.5 text-right font-bold tabular-nums text-white">{formatMiles(portfolioTotalesFiltrados.compraR)}</td>
                              <td className="px-4 py-2.5 text-right font-bold tabular-nums text-white">{formatMiles(portfolioTotalesFiltrados.compraNr)}</td>
                              <td className="px-4 py-2.5 text-right font-bold tabular-nums text-white">{formatMiles(portfolioTotalesFiltrados.venta)}</td>
                              <td className="px-4 py-2.5 text-right font-bold tabular-nums text-white">{formatMiles(portfolioTotalesFiltrados.demandaR)}</td>
                              <td className="px-4 py-2.5 text-right font-bold tabular-nums text-white">{formatMiles(portfolioTotalesFiltrados.demandaNr)}</td>
                              <td className={`px-4 py-2.5 text-right font-bold tabular-nums ${portfolioTotalesFiltrados.posicionNeta > 0 ? 'text-red-300' : 'text-emerald-300'}`}>
                                {formatMiles(portfolioTotalesFiltrados.posicionNeta)}
                              </td>
                              <td className="px-4 py-2.5 text-right font-bold tabular-nums text-white">{formatMiles(portfolioTotalesFiltrados.costoBolsa, 2)}</td>
                              <td className="px-4 py-2.5 text-right text-xs font-bold text-white">
                                {nComp} días comp. / {nVend} días vend.
                              </td>
                            </tr>
                          )
                        })()}
                      </>
                    }
                  />
                )}
              </div>
            </section>
          </>
        )}

        {tabActiva === 'simulador' && (
          <>
            <div>
              <h2 className="text-xl font-bold text-gray-100">Simulador</h2>
              <p className="mt-0.5 text-sm text-gray-500">Evalúa el impacto de un nuevo contrato — sigue los 4 pasos</p>
            </div>

            <WizardProgressBar pasoActual={wizPaso} />

            {/* ══ PASO 1 — Período ══ */}
            {wizPaso === 1 && (
              <div className="space-y-5">
                <section className="rounded-xl border border-gray-800 bg-gray-900/60 p-6">
                  <p className="mb-4 text-xs font-semibold uppercase tracking-widest text-emerald-500">Paso 1 — ¿Qué período quieres simular?</p>
                  <div className="flex flex-wrap items-end gap-4">
                    <div className="flex flex-col gap-1.5">
                      <label className="text-xs text-gray-400">Inicio del período</label>
                      <input type="date" min="2026-01-01" value={wizPeriodoInicio} onChange={e => setWizPeriodoInicio(e.target.value)}
                        className="rounded-lg border border-gray-700 bg-gray-950 px-4 py-2.5 text-sm text-gray-100 outline-none [color-scheme:dark] focus:border-emerald-500/50 focus:ring-2 focus:ring-emerald-500/30" />
                    </div>
                    <div className="flex flex-col gap-1.5">
                      <label className="text-xs text-gray-400">Fin del período</label>
                      <input type="date" value={wizPeriodoFin} onChange={e => setWizPeriodoFin(e.target.value)}
                        className="rounded-lg border border-gray-700 bg-gray-950 px-4 py-2.5 text-sm text-gray-100 outline-none [color-scheme:dark] focus:border-emerald-500/50 focus:ring-2 focus:ring-emerald-500/30" />
                    </div>
                    <button type="button" onClick={wizCargarPosicion} disabled={wizPortfolioCargando}
                      className="rounded-lg bg-emerald-600 px-5 py-2.5 text-sm font-semibold text-white transition hover:bg-emerald-500 disabled:opacity-60 disabled:cursor-not-allowed">
                      {wizPortfolioCargando ? <span className="flex items-center gap-2"><span className="h-4 w-4 animate-spin rounded-full border-2 border-white border-t-transparent" />Cargando…</span> : '📊 Ver posición actual'}
                    </button>
                  </div>
                  {wizPortfolioError && <p className="mt-3 text-sm text-red-400">{wizPortfolioError}</p>}
                </section>

                {wizResumenMensual.length > 0 && (() => {
                  // Totales en kWh
                  const totR    = wizResumenMensual.reduce((s, m) => s + m.compraRkwh,   0)
                  const totNr   = wizResumenMensual.reduce((s, m) => s + m.compraNrkwh,  0)
                  const totV    = wizResumenMensual.reduce((s, m) => s + m.ventaKwh,     0)
                  const totDR   = wizResumenMensual.reduce((s, m) => s + m.demandaRkwh,  0)
                  const totDNr  = wizResumenMensual.reduce((s, m) => s + m.demandaNrkwh, 0)
                  const totP    = wizResumenMensual.reduce((s, m) => s + m.posNetaKwh,   0)
                  // COP mock en MWh para compatibilidad paso 2
                  const totCR   = wizResumenMensual.reduce((s, m) => s + m.copCompraR,   0)
                  const totCNr  = wizResumenMensual.reduce((s, m) => s + m.copCompraNr,  0)
                  const totCV   = wizResumenMensual.reduce((s, m) => s + m.copVenta,     0)
                  const nVend   = wizResumenMensual.filter(m => m.posNetaKwh <= 0).length
                  return (
                    <>
                      <div className="rounded-xl border border-gray-700 bg-gray-800/40 px-5 py-3">
                        <p className="text-xs font-semibold uppercase tracking-widest text-gray-400">Capítulo 1 — Posición actual en el período</p>
                        {wizPPPResumen ? (
                          <p className="mt-0.5 text-xs text-gray-500">
                            PPP Olibia — Compra R: {wizPPPResumen.compra_r.ppp != null ? `${formatNumero(wizPPPResumen.compra_r.ppp)} COP/kWh` : '—'}
                            {' · '}Compra NR: {wizPPPResumen.compra_nr.ppp != null ? `${formatNumero(wizPPPResumen.compra_nr.ppp)} COP/kWh` : '—'}
                            {' · '}Venta: {wizPPPResumen.venta.ppp != null ? `${formatNumero(wizPPPResumen.venta.ppp)} COP/kWh` : '—'}
                            {wizPPPResumen.pld_excluidos > 0 && ` · ${wizPPPResumen.pld_excluidos} PLD excluidos`}
                          </p>
                        ) : (
                          <p className="mt-0.5 text-xs text-gray-600">Precios mock (fallback): Compra R {PRECIO_MOCK_COMPRA_R} · Compra NR {PRECIO_MOCK_COMPRA_NR} · Venta {PRECIO_MOCK_VENTA} COP/kWh</p>
                        )}
                      </div>
                      <div className="grid grid-cols-2 gap-3 sm:grid-cols-4">
                        <MetricCard label="Compra total (kWh)" value={`${formatMiles(totR + totNr, 0)}`} subValue={`${formatMiles(totCR + totCNr, 1)} M COP`} accent="text-sky-400" />
                        <MetricCard label="Venta total (kWh)"  value={`${formatMiles(totV, 0)}`} subValue={`${formatMiles(totCV, 1)} M COP`} accent="text-rose-400" />
                        <MetricCard label="Posición neta (kWh)" value={`${formatMiles(totP, 0)}`} accent={totP <= 0 ? 'text-emerald-400' : 'text-red-400'} />
                        <MetricCard
                          label="Posición dominante"
                          value={nVend >= (wizResumenMensual.length - nVend) ? 'Vendiendo en bolsa' : 'Comprando en bolsa'}
                          subValue={`${nVend} vend. / ${wizResumenMensual.length - nVend} comp.`}
                          accent={nVend >= (wizResumenMensual.length - nVend) ? 'text-emerald-400' : 'text-red-400'}
                        />
                      </div>
                      <TablaAnalisis
                        encabezados={['Mes','Compra R (kWh)','Compra NR (kWh)','Venta (kWh)','Demanda R (kWh)','Demanda NR (kWh)','Pos. Neta (kWh)','Tipo Posición']}
                        filas={<>
                          {wizResumenMensual.map(m => (
                            <tr key={m.mes} className="transition-colors hover:bg-gray-800/40">
                              <td className="px-4 py-2.5 font-medium text-gray-200">{formatMes(m.mes)}</td>
                              <td className="px-4 py-2.5 text-right tabular-nums text-sky-300">{formatMiles(m.compraRkwh, 0)}</td>
                              <td className="px-4 py-2.5 text-right tabular-nums text-sky-300">{formatMiles(m.compraNrkwh, 0)}</td>
                              <td className="px-4 py-2.5 text-right tabular-nums text-rose-300">{formatMiles(m.ventaKwh, 0)}</td>
                              <td className="px-4 py-2.5 text-right tabular-nums text-orange-400">{formatMiles(m.demandaRkwh, 0)}</td>
                              <td className="px-4 py-2.5 text-right tabular-nums text-purple-400">{formatMiles(m.demandaNrkwh, 0)}</td>
                              <td className={`px-4 py-2.5 text-right font-semibold tabular-nums ${m.posNetaKwh <= 0 ? 'text-emerald-400' : 'text-red-400'}`}>{formatMiles(m.posNetaKwh, 0)}</td>
                              <td className={`px-4 py-2.5 text-right text-xs font-semibold ${m.posNetaKwh <= 0 ? 'text-emerald-400' : 'text-red-400'}`}>{m.tipoPos}</td>
                            </tr>
                          ))}
                          <tr className="border-t-2 border-gray-600 bg-gray-700">
                            <td className="px-4 py-2.5 text-xs font-bold uppercase tracking-wider text-white">TOTAL PERÍODO</td>
                            <td className="px-4 py-2.5 text-right font-bold tabular-nums text-white">{formatMiles(totR, 0)}</td>
                            <td className="px-4 py-2.5 text-right font-bold tabular-nums text-white">{formatMiles(totNr, 0)}</td>
                            <td className="px-4 py-2.5 text-right font-bold tabular-nums text-white">{formatMiles(totV, 0)}</td>
                            <td className="px-4 py-2.5 text-right font-bold tabular-nums text-white">{formatMiles(totDR, 0)}</td>
                            <td className="px-4 py-2.5 text-right font-bold tabular-nums text-white">{formatMiles(totDNr, 0)}</td>
                            <td className={`px-4 py-2.5 text-right font-bold tabular-nums ${totP <= 0 ? 'text-emerald-300' : 'text-red-300'}`}>{formatMiles(totP, 0)}</td>
                            <td className="px-4 py-2.5 text-right text-xs font-bold text-white">
                              {wizResumenMensual.length - nVend} meses comp. / {nVend} meses vend.
                            </td>
                          </tr>
                        </>}
                      />
                      <div className="rounded-xl border border-gray-800 bg-gray-900/60 p-5">
                        <p className="mb-4 text-sm font-medium text-gray-400">Compra vs Venta vs Demanda por mes (kWh)</p>
                        <div className="h-64">
                          <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={wizResumenMensual.map(m => ({ mes: m.mes, compra: Math.round(m.compraRkwh + m.compraNrkwh), venta: Math.round(m.ventaKwh), demanda: Math.round(m.demandaRkwh + m.demandaNrkwh) }))} margin={{ top: 4, right: 16, left: 0, bottom: 4 }}>
                              <CartesianGrid stroke="#374151" strokeDasharray="3 3" />
                              <XAxis dataKey="mes" stroke="#9ca3af" tick={{ fill: '#9ca3af', fontSize: 11 }} tickFormatter={v => formatMes(String(v)).slice(0,3)} />
                              <YAxis stroke="#9ca3af" tick={{ fill: '#9ca3af', fontSize: 11 }} />
                              <Tooltip contentStyle={{ backgroundColor: '#111827', border: '1px solid #374151', borderRadius: '8px', color: '#f3f4f6' }} labelFormatter={v => formatMes(String(v))} formatter={(v, n) => [`${formatMiles(Number(v), 0)} kWh`, n]} />
                              <Legend wrapperStyle={{ color: '#9ca3af', paddingTop: 8 }} />
                              <Bar dataKey="compra"  name="Compra"  fill="#3b82f6" radius={[3,3,0,0]} />
                              <Bar dataKey="venta"   name="Venta"   fill="#ef4444" radius={[3,3,0,0]} />
                              <Bar dataKey="demanda" name="Demanda" fill="#f97316" radius={[3,3,0,0]} />
                            </BarChart>
                          </ResponsiveContainer>
                        </div>
                      </div>
                      <div className="flex justify-end">
                        <button type="button" onClick={() => setWizPaso(2)} className="rounded-lg bg-emerald-600 px-6 py-2.5 text-sm font-semibold text-white transition hover:bg-emerald-500">
                          Siguiente → Escenario PB
                        </button>
                      </div>
                    </>
                  )
                })()}
              </div>
            )}

            {/* ══ PASO 2 — Escenario PB ══ */}
            {wizPaso === 2 && (
              <div className="space-y-5">
                <section className="rounded-xl border border-gray-800 bg-gray-900/60 p-6">
                  <p className="mb-4 text-xs font-semibold uppercase tracking-widest text-emerald-500">Paso 2 — ¿Con qué precio de bolsa vas a analizar?</p>
                  <div className="mb-5 inline-flex w-full rounded-lg border border-gray-700 bg-gray-950 p-1">
                    {(['historico', 'enso', 'proyectado'] as WizFuentePB[]).map(f => (
                      <button key={f} type="button" onClick={() => setWizFuentePB(f)}
                        className={`flex-1 rounded-md py-2 text-xs font-semibold transition ${wizFuentePB === f ? 'bg-emerald-600 text-white' : 'text-gray-400 hover:bg-gray-800'}`}>
                        {f === 'historico' ? '📅 PB Histórico' : f === 'enso' ? '🌡 Fenómeno ENSO' : '📈 PB Proyectado'}
                      </button>
                    ))}
                  </div>

                  {wizFuentePB === 'historico' && (
                    <div className="grid grid-cols-2 gap-3">
                      <div><label className="mb-1 block text-xs text-gray-500">Desde</label>
                        <input type="date" min="2010-01-01" value={wizPBDesde} onChange={e => setWizPBDesde(e.target.value)}
                          className="w-full rounded-lg border border-gray-700 bg-gray-950 px-3 py-2 text-sm text-gray-100 outline-none [color-scheme:dark] focus:border-emerald-500/50" />
                      </div>
                      <div><label className="mb-1 block text-xs text-gray-500">Hasta</label>
                        <input type="date" value={wizPBHasta} onChange={e => setWizPBHasta(e.target.value)}
                          className="w-full rounded-lg border border-gray-700 bg-gray-950 px-3 py-2 text-sm text-gray-100 outline-none [color-scheme:dark] focus:border-emerald-500/50" />
                      </div>
                      <p className="col-span-2 text-xs text-gray-600">Máximo 36 meses para evitar timeout.</p>
                    </div>
                  )}

                  {wizFuentePB === 'enso' && (
                    <div>
                      <div className="overflow-x-auto rounded-lg border border-gray-700">
                        <table className="w-full text-sm">
                          <thead>
                            <tr className="border-b border-gray-700 bg-gray-900/60 text-xs uppercase tracking-wider text-gray-500">
                              <th className="px-3 py-2.5 text-left">Período</th>
                              <th className="px-3 py-2.5 text-left">Fenómeno</th>
                              <th className="px-3 py-2.5 text-left">Intensidad</th>
                              <th className="px-3 py-2.5 text-center">Seleccionar</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-gray-800/60">
                            {ENSO_COMPLETO.map(e => {
                              const sel = wizEnsoKey === e.key
                              const clr = e.fenomeno === 'nino' ? 'text-orange-400' : e.fenomeno === 'nina' ? 'text-blue-400' : 'text-gray-400'
                              const bgSel = sel ? (e.fenomeno === 'nino' ? 'bg-orange-500/10' : e.fenomeno === 'nina' ? 'bg-blue-500/10' : 'bg-gray-500/10') : ''
                              return (
                                <tr key={e.key} onClick={() => setWizEnsoKey(e.key as EnsoKey)} className={`cursor-pointer transition-colors hover:bg-gray-800/40 ${bgSel}`}>
                                  <td className={`px-3 py-2.5 font-medium ${clr}`}>{e.label}</td>
                                  <td className={`px-3 py-2.5 text-xs font-semibold ${clr}`}>{e.fenomeno === 'nino' ? '🔴 El Niño' : e.fenomeno === 'nina' ? '🔵 La Niña' : '⚪ Neutral'}</td>
                                  <td className="px-3 py-2.5 text-xs text-gray-400">{e.intensidad}</td>
                                  <td className="px-3 py-2.5 text-center"><input type="radio" checked={sel} onChange={() => setWizEnsoKey(e.key as EnsoKey)} className="h-4 w-4 accent-emerald-500" /></td>
                                </tr>
                              )
                            })}
                          </tbody>
                        </table>
                      </div>
                      {(() => { const enso = ENSO_COMPLETO.find(e => e.key === wizEnsoKey); return enso ? (<p className="mt-3 text-xs text-gray-500">📌 Usando PB histórico del período <span className="font-semibold text-gray-300">{enso.pb_desde} → {enso.pb_hasta}</span> como proxy de escenario <span className="font-semibold text-gray-300">{enso.label}</span></p>) : null })()}
                    </div>
                  )}

                  {wizFuentePB === 'proyectado' && (
                    <div className="space-y-3">
                      <p className="text-xs text-amber-500">⚠ Por ahora se usará PB histórico como proxy proyectado. Configura el rango base:</p>
                      <div className="grid grid-cols-2 gap-3">
                        <div><label className="mb-1 block text-xs text-gray-500">Desde</label>
                          <input type="date" min="2010-01-01" value={wizPBDesde} onChange={e => setWizPBDesde(e.target.value)}
                            className="w-full rounded-lg border border-gray-700 bg-gray-950 px-3 py-2 text-sm text-gray-100 outline-none [color-scheme:dark] focus:border-emerald-500/50" />
                        </div>
                        <div><label className="mb-1 block text-xs text-gray-500">Hasta</label>
                          <input type="date" value={wizPBHasta} onChange={e => setWizPBHasta(e.target.value)}
                            className="w-full rounded-lg border border-gray-700 bg-gray-950 px-3 py-2 text-sm text-gray-100 outline-none [color-scheme:dark] focus:border-emerald-500/50" />
                        </div>
                      </div>
                    </div>
                  )}

                  <div className="mt-5 flex items-center gap-3">
                    <button type="button" onClick={wizCargarPB} disabled={wizPBCargando}
                      className="rounded-lg bg-emerald-600 px-5 py-2.5 text-sm font-semibold text-white transition hover:bg-emerald-500 disabled:opacity-60 disabled:cursor-not-allowed">
                      {wizPBCargando ? <span className="flex items-center gap-2"><span className="h-4 w-4 animate-spin rounded-full border-2 border-white border-t-transparent" />Calculando…</span> : '⚡ Calcular escenario bolsa'}
                    </button>
                    {wizPBError && <p className="text-sm text-red-400">{wizPBError}</p>}
                  </div>
                </section>

                {wizPBResumenMensual.length > 0 && wizPBPromedioGlobal !== null && (() => {
                  const totalTransaccionMcop = wizPBResumenMensual.reduce((s, m) => s + m.transaccionMcop, 0)
                  // Usar PPP real de Compra R si disponible; mock como fallback
                  const pppRef      = wizPPPResumen?.compra_r?.ppp ?? PRECIO_MOCK_COMPRA_R
                  const esPPPReal   = wizPPPResumen?.compra_r?.ppp != null
                  const tipoPPP     = wizPPPResumen?.compra_r?.tipo ?? 'Sin datos'
                  const spread      = pppRef - wizPBPromedioGlobal
                  const semColor    = spread > 5 ? 'text-emerald-400 border-emerald-500/40 bg-emerald-500/10' : spread > -5 ? 'text-amber-400 border-amber-500/40 bg-amber-500/10' : 'text-red-400 border-red-500/40 bg-red-500/10'
                  const semIcon     = spread > 5 ? '✓' : spread > -5 ? '⚠' : '✗'
                  const pppLabel    = esPPPReal ? `PPP Compra R ${formatNumero(pppRef)} COP/kWh` : `contratos mock ${pppRef} COP/kWh`
                  const semTexto    = spread > 5
                    ? `Spread favorable: ${pppLabel} vs PB promedio ${formatNumero(wizPBPromedioGlobal)} COP/kWh (+${formatNumero(spread)} COP/kWh).`
                    : spread > -5
                    ? `Spread marginal (${formatNumero(spread, 1)} COP/kWh). Evalúa con más detalle.`
                    : `Spread negativo: PB promedio ${formatNumero(wizPBPromedioGlobal)} COP/kWh supera ${pppLabel}.`
                  return (
                    <>
                      <div className="rounded-xl border border-gray-700 bg-gray-800/40 px-5 py-3">
                        <p className="text-xs font-semibold uppercase tracking-widest text-gray-400">Capítulo 2 — Costo / Ingreso en bolsa con este escenario</p>
                      </div>
                      <div className={`border px-5 py-3 ${semColor}`}>
                        <div className="flex items-center gap-3"><span className="text-xl font-bold">{semIcon}</span><p className="text-sm">{semTexto}</p></div>
                      </div>
                      <div className="grid grid-cols-2 gap-3 sm:grid-cols-4">
                        {/* PB promedio del escenario seleccionado */}
                        <MetricCard label="PB promedio escenario" value={`${formatNumero(wizPBPromedioGlobal)} COP/kWh`} accent="text-white" />
                        {/* PPP real de Compra R con badge de tipo; mock si no hay datos */}
                        {esPPPReal ? (
                          <div className="rounded-xl border border-gray-800 bg-gray-900 p-5">
                            <p className="text-xs font-medium uppercase tracking-wider text-gray-500">PPP Compra R</p>
                            <p className="mt-2 text-2xl font-bold tabular-nums text-white">{formatNumero(pppRef)}</p>
                            <p className="text-xs text-gray-500">COP/kWh</p>
                            <div className="mt-1"><BadgeTipoPrecio tipo={tipoPPP as 'Indexado' | 'Proyectado' | 'Sin datos'} /></div>
                          </div>
                        ) : (
                          <MetricCard label="PPP contratos (mock)" value={`${PRECIO_MOCK_COMPRA_R} COP/kWh`} accent="text-amber-400" />
                        )}
                        <MetricCard label="Spread promedio" value={`${spread >= 0 ? '+' : ''}${formatNumero(spread)} COP/kWh`} accent={spread >= 0 ? 'text-emerald-400' : 'text-red-400'} />
                        <MetricCard
                          label={totalTransaccionMcop > 0 ? 'Costo en bolsa (M COP)' : 'Ingreso en bolsa (M COP)'}
                          value={`${formatMiles(totalTransaccionMcop, 2)} M COP`}
                          accent={totalTransaccionMcop <= 0 ? 'text-emerald-400' : 'text-red-400'}
                        />
                      </div>
                      {/* Detalle de PPP por categoría y aviso PLD si hay contratos excluidos */}
                      {wizPPPResumen && (
                        <div className="rounded-xl border border-gray-800 bg-gray-900/40 px-5 py-4">
                          <p className="mb-3 text-xs font-semibold uppercase tracking-widest text-gray-500">PPP por categoría de contrato PC</p>
                          <div className="grid grid-cols-1 gap-2 sm:grid-cols-3">
                            {(
                              [
                                { label: 'Compra Regulada',    cat: wizPPPResumen.compra_r  },
                                { label: 'Compra No Regulada', cat: wizPPPResumen.compra_nr },
                                { label: 'Venta',              cat: wizPPPResumen.venta     },
                              ] as const
                            ).map(({ label, cat }) => (
                              <div key={label} className="flex items-center justify-between rounded-lg border border-gray-700/50 bg-gray-800/40 px-3 py-2">
                                <span className="text-xs text-gray-400">{label}</span>
                                <div className="flex items-center gap-1.5">
                                  {cat.ppp != null
                                    ? <span className="text-sm font-bold tabular-nums text-white">{formatNumero(cat.ppp)} COP/kWh</span>
                                    : <span className="text-sm text-gray-600">—</span>
                                  }
                                  <BadgeTipoPrecio tipo={cat.tipo} />
                                </div>
                              </div>
                            ))}
                          </div>
                          {wizPPPResumen.pld_excluidos > 0 && (
                            <p className="mt-3 text-xs text-amber-500">
                              ⚠ {wizPPPResumen.pld_excluidos} contrato(s) PLD excluidos del PPP — su precio es el precio de bolsa en cada hora.
                            </p>
                          )}
                        </div>
                      )}
                      <TablaAnalisis
                        encabezados={['Mes','PB Promedio (COP/kWh)','Posición Neta (MWh)','Trans. Bolsa (M COP)','Spread Contratos vs PB']}
                        filas={wizPBResumenMensual.map(m => (
                          <tr key={m.mes} className="transition-colors hover:bg-gray-800/40">
                            <td className="px-4 py-2.5 font-medium text-gray-200">{formatMes(m.mes)}</td>
                            <td className="px-4 py-2.5 text-right tabular-nums text-gray-200">{formatNumero(m.pbProm)}</td>
                            <td className={`px-4 py-2.5 text-right tabular-nums ${m.posNetaMwh <= 0 ? 'text-emerald-400' : 'text-red-400'}`}>{formatMiles(m.posNetaMwh, 0)}</td>
                            <td className={`px-4 py-2.5 text-right tabular-nums ${m.transaccionMcop <= 0 ? 'text-emerald-400' : 'text-red-400'}`}>{formatMiles(m.transaccionMcop, 2)}</td>
                            <td className={`px-4 py-2.5 text-right font-semibold tabular-nums ${m.spreadVsPB >= 0 ? 'text-emerald-400' : 'text-red-400'}`}>{m.spreadVsPB >= 0 ? '+' : ''}{formatNumero(m.spreadVsPB)}</td>
                          </tr>
                        ))}
                      />
                      <div className="rounded-xl border border-gray-800 bg-gray-900/60 p-5">
                        <p className="mb-1 text-sm font-medium text-gray-400">PB hora a hora — P10 / Promedio / P90</p>
                        <div className="h-64">
                          <ResponsiveContainer width="100%" height="100%">
                            <LineChart data={wizPBPerfilHorario} margin={{ top: 4, right: 16, left: 0, bottom: 4 }}>
                              <CartesianGrid stroke="#374151" strokeDasharray="3 3" />
                              <XAxis dataKey="hora" type="number" domain={[1,24]} ticks={Array.from({length:24},(_,i)=>i+1)} stroke="#9ca3af" tick={{ fill: '#9ca3af', fontSize: 11 }} />
                              <YAxis stroke="#9ca3af" tick={{ fill: '#9ca3af', fontSize: 11 }} />
                              <Tooltip contentStyle={{ backgroundColor: '#111827', border: '1px solid #374151', borderRadius: '8px', color: '#f3f4f6' }} labelFormatter={h => `Hora ${h}`} formatter={(v, n) => [v != null ? `${formatNumero(Number(v))} COP/kWh` : '—', n]} />
                              <Legend wrapperStyle={{ color: '#9ca3af', paddingTop: 8 }} />
                              <Line type="monotone" dataKey="p10"      name="P10"      stroke="#6b7280" strokeWidth={1} strokeDasharray="4 4" dot={false} connectNulls />
                              <Line type="monotone" dataKey="promedio" name="Promedio" stroke="#3b82f6" strokeWidth={2.5} dot={false} connectNulls />
                              <Line type="monotone" dataKey="p90"      name="P90"      stroke="#6b7280" strokeWidth={1} strokeDasharray="4 4" dot={false} connectNulls />
                            </LineChart>
                          </ResponsiveContainer>
                        </div>
                      </div>
                      <div className="flex justify-between">
                        <button type="button" onClick={() => setWizPaso(1)} className="rounded-lg border border-gray-700 px-5 py-2.5 text-sm font-semibold text-gray-300 transition hover:bg-gray-800">← Período</button>
                        <button type="button" onClick={() => setWizPaso(3)} className="rounded-lg bg-emerald-600 px-6 py-2.5 text-sm font-semibold text-white transition hover:bg-emerald-500">Siguiente → Nuevo contrato</button>
                      </div>
                    </>
                  )
                })()}
                {wizPBResumenMensual.length === 0 && (
                  <div className="flex justify-start">
                    <button type="button" onClick={() => setWizPaso(1)} className="rounded-lg border border-gray-700 px-5 py-2.5 text-sm font-semibold text-gray-300 transition hover:bg-gray-800">← Período</button>
                  </div>
                )}
              </div>
            )}

            {/* ══ PASO 3 — Nuevo contrato ══ */}
            {wizPaso === 3 && (
              <div className="grid grid-cols-1 gap-5 lg:grid-cols-5">
                <section className="rounded-xl border border-gray-800 bg-gray-900/60 p-5 lg:col-span-3">
                  <p className="mb-4 text-xs font-semibold uppercase tracking-widest text-blue-400">Paso 3 — ¿Qué contrato quieres simular?</p>

                  <div className="mb-4">
                    <label className="mb-1.5 block text-xs text-gray-500">Tipo de operación</label>
                    <div className="inline-flex w-full rounded-lg border border-gray-700 bg-gray-950 p-1">
                      {(['compra','venta'] as SimTipo[]).map(t => (
                        <button key={t} type="button" onClick={() => setWizSimTipo(t)}
                          className={`flex-1 rounded-md py-2 text-sm font-semibold capitalize transition ${wizSimTipo === t ? (t === 'compra' ? 'bg-blue-600 text-white' : 'bg-rose-600 text-white') : 'text-gray-400 hover:bg-gray-800'}`}>
                          {t === 'compra' ? '↓ Compra' : '↑ Venta'}
                        </button>
                      ))}
                    </div>
                  </div>

                  <div className="mb-3">
                    <label className="mb-1 block text-xs text-gray-500">Contraparte</label>
                    <input type="text" placeholder="Nombre de la contraparte" value={wizSimContraparte} onChange={e => setWizSimContraparte(e.target.value)}
                      className="w-full rounded-lg border border-gray-700 bg-gray-950 px-3 py-2 text-sm text-gray-100 placeholder-gray-600 outline-none focus:border-emerald-500/50 focus:ring-2 focus:ring-emerald-500/30" />
                  </div>

                  <div className="mb-4">
                    <label className="mb-1 block text-xs text-gray-500">Precio (COP/kWh)</label>
                    <input type="number" min={1} value={wizSimPrecio} onChange={e => setWizSimPrecio(Math.max(1, Number(e.target.value) || 1))}
                      className="w-full rounded-lg border border-gray-700 bg-gray-950 px-3 py-2 text-sm text-gray-100 outline-none focus:border-emerald-500/50 focus:ring-2 focus:ring-emerald-500/30" />
                  </div>

                  <div className="mb-4">
                    <label className="mb-2 block text-xs text-gray-500">Distribución horaria</label>
                    <div className="grid grid-cols-4 gap-1.5">
                      {(['plano','bloques','solar','excel'] as SimPerfilTipo[]).map(p => (
                        <button key={p} type="button" onClick={() => setWizSimPerfilTipo(p)}
                          className={`rounded-lg py-2 text-xs font-semibold capitalize transition ${wizSimPerfilTipo === p ? 'bg-gray-600 text-white ring-1 ring-gray-400' : 'bg-gray-800 text-gray-400 hover:bg-gray-700'}`}>
                          {p === 'plano' ? 'Plano 24h' : p === 'bloques' ? 'Bloques' : p === 'solar' ? 'Solar' : 'Excel'}
                        </button>
                      ))}
                    </div>
                    <div className="mt-3">
                      {(wizSimPerfilTipo === 'plano' || wizSimPerfilTipo === 'solar') && (
                        <div>
                          <label className="mb-1 block text-xs text-gray-500">Energía mensual (kWh/mes)</label>
                          <input type="number" min={1} value={wizSimEnergiaKwh} onChange={e => setWizSimEnergiaKwh(Math.max(1, Number(e.target.value) || 1))}
                            className="w-full rounded-lg border border-gray-700 bg-gray-950 px-3 py-2 text-sm text-gray-100 outline-none focus:border-emerald-500/50" />
                          {wizSimPerfilTipo === 'solar' && <p className="mt-1 text-xs text-gray-600">Curva solar: 0 en H1-H6 y H19-H24 · sube H7-H12 · baja H13-H18</p>}
                        </div>
                      )}
                      {wizSimPerfilTipo === 'bloques' && (
                        <div className="space-y-2">
                          {wizSimBloques.map((b, idx) => (
                            <div key={idx} className="flex items-end gap-2 rounded-lg border border-gray-700 bg-gray-950/60 p-3">
                              <div className="flex flex-col gap-1">
                                <label className="text-xs text-gray-600">H ini</label>
                                <select value={b.horaInicio} onChange={e => wizUpdateBloque(idx,'horaInicio',Number(e.target.value))} className="w-16 rounded border border-gray-700 bg-gray-900 px-2 py-1.5 text-xs text-gray-100 outline-none">
                                  {Array.from({length:24},(_,i)=>i+1).map(h=><option key={h} value={h}>H{h}</option>)}
                                </select>
                              </div>
                              <div className="flex flex-col gap-1">
                                <label className="text-xs text-gray-600">H fin</label>
                                <select value={b.horaFin} onChange={e => wizUpdateBloque(idx,'horaFin',Number(e.target.value))} className="w-16 rounded border border-gray-700 bg-gray-900 px-2 py-1.5 text-xs text-gray-100 outline-none">
                                  {Array.from({length:24},(_,i)=>i+1).filter(h=>h>=b.horaInicio).map(h=><option key={h} value={h}>H{h}</option>)}
                                </select>
                              </div>
                              <div className="flex flex-1 flex-col gap-1">
                                <label className="text-xs text-gray-600">MWh/mes</label>
                                <input type="number" min={1} value={b.mwhMes} onChange={e=>wizUpdateBloque(idx,'mwhMes',Math.max(1,Number(e.target.value)||1))} className="w-full rounded border border-gray-700 bg-gray-900 px-2 py-1.5 text-xs text-gray-100 outline-none" />
                              </div>
                              {wizSimBloques.length > 1 && <button type="button" onClick={()=>wizRemoveBloque(idx)} className="rounded px-1.5 py-1.5 text-gray-600 hover:text-red-400 transition">✕</button>}
                            </div>
                          ))}
                          {wizSimBloques.length < 3 && <button type="button" onClick={wizAddBloque} className="w-full rounded-lg border border-dashed border-gray-700 py-2 text-xs text-gray-500 hover:border-gray-500 hover:text-gray-300 transition">+ Agregar bloque</button>}
                          <p className="text-xs text-gray-600">Total: {formatMiles(wizSimBloques.reduce((s,b)=>s+b.mwhMes,0),0)} MWh/mes</p>
                        </div>
                      )}
                      {wizSimPerfilTipo === 'excel' && (
                        <div>
                          <input ref={wizExcelInputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={wizHandleExcelUpload} />
                          <button type="button" onClick={()=>wizExcelInputRef.current?.click()} className="w-full rounded-lg border border-dashed border-gray-600 py-3 text-sm text-gray-400 hover:border-emerald-500/50 hover:text-emerald-400 transition">📎 Cargar archivo Excel / CSV</button>
                          {wizSimExcelNombre && <p className="mt-1.5 flex items-center gap-1.5 text-xs text-emerald-400">✓ {wizSimExcelNombre} {wizSimExcel12x24 && <span className="text-gray-600">({wizSimExcel12x24.length} meses)</span>}</p>}
                        </div>
                      )}
                    </div>
                  </div>

                  <div className="mb-4 grid grid-cols-2 gap-3">
                    <div><label className="mb-1 block text-xs text-gray-500">Inicio vigencia</label>
                      <input type="date" value={wizContratoInicio} onChange={e=>setWizContratoInicio(e.target.value)} className="w-full rounded-lg border border-gray-700 bg-gray-950 px-3 py-2 text-sm text-gray-100 outline-none [color-scheme:dark] focus:border-emerald-500/50" />
                    </div>
                    <div><label className="mb-1 block text-xs text-gray-500">Fin vigencia</label>
                      <input type="date" value={wizContratoFin} onChange={e=>setWizContratoFin(e.target.value)} className="w-full rounded-lg border border-gray-700 bg-gray-950 px-3 py-2 text-sm text-gray-100 outline-none [color-scheme:dark] focus:border-emerald-500/50" />
                    </div>
                  </div>

                  <div className="mb-5">
                    <label className="mb-1 block text-xs text-gray-500">Tipo de mercado</label>
                    <select value={wizSimTipoMercado} onChange={e=>setWizSimTipoMercado(e.target.value as SimMercado)} className={CLASE_SELECT + ' w-full'}>
                      <option value="regulado">Regulado (Compra R)</option>
                      <option value="no_regulado">No Regulado (Compra NR)</option>
                      <option value="ambos">Ambos (50% R + 50% NR)</option>
                    </select>
                  </div>

                  <div className="flex items-center gap-3">
                    <button type="button" onClick={()=>setWizPaso(2)} className="rounded-lg border border-gray-700 px-5 py-2.5 text-sm font-semibold text-gray-300 transition hover:bg-gray-800">← Escenario PB</button>
                    <button type="button" onClick={wizSimular} disabled={wizSimCargando} className="flex-1 rounded-lg bg-emerald-600 py-2.5 text-sm font-bold text-white shadow transition hover:bg-emerald-500 disabled:opacity-60 disabled:cursor-not-allowed">
                      {wizSimCargando ? <span className="flex items-center justify-center gap-2"><span className="h-4 w-4 animate-spin rounded-full border-2 border-white border-t-transparent" />Simulando…</span> : '⚡ Simular impacto'}
                    </button>
                  </div>
                  {wizSimError && <div className="mt-3 rounded-lg border border-red-900/50 bg-red-950/30 px-4 py-3 text-sm text-red-400">{wizSimError}</div>}
                </section>

                <div className="space-y-4 lg:col-span-2">
                  {wizPBPromedioGlobal !== null && (
                    <div className="rounded-xl border border-gray-700 bg-gray-900/40 p-4">
                      <p className="mb-3 text-xs font-semibold uppercase tracking-widest text-gray-500">Escenario PB seleccionado</p>
                      <div className="space-y-1.5 text-sm text-gray-300">
                        <p>📅 <span className="font-semibold text-white">{wizPBDesde}</span> → <span className="font-semibold text-white">{wizPBHasta}</span></p>
                        <p>📊 PB promedio: <span className="font-semibold text-white">{formatNumero(wizPBPromedioGlobal)} COP/kWh</span></p>
                        {/* Spread: usa PPP real de Compra R si disponible, mock como fallback */}
                        {(() => {
                          const pppRef   = wizPPPResumen?.compra_r?.ppp ?? PRECIO_MOCK_COMPRA_R
                          const spreadPP = pppRef - wizPBPromedioGlobal
                          const esPPP    = wizPPPResumen?.compra_r?.ppp != null
                          return (
                            <div className="flex items-center gap-2">
                              <span className="text-gray-400">{esPPP ? 'Spread PPP vs PB:' : 'Spread mock vs PB:'}</span>
                              <span className={`font-semibold ${spreadPP >= 0 ? 'text-emerald-400' : 'text-red-400'}`}>
                                {spreadPP >= 0 ? '+' : ''}{formatNumero(spreadPP)} COP/kWh
                              </span>
                              {esPPP && <BadgeTipoPrecio tipo={wizPPPResumen!.compra_r.tipo} />}
                            </div>
                          )
                        })()}
                      </div>
                    </div>
                  )}
                  {wizResumenMensual.length > 0 && (
                    <div className="rounded-xl border border-gray-700 bg-gray-900/40 p-4">
                      <p className="mb-3 text-xs font-semibold uppercase tracking-widest text-gray-500">Período a simular</p>
                      <div className="space-y-1.5 text-sm text-gray-300">
                        <p>📆 <span className="font-semibold text-white">{wizPeriodoInicio}</span> → <span className="font-semibold text-white">{wizPeriodoFin}</span></p>
                        <p>🏭 Meses: <span className="font-semibold text-white">{wizResumenMensual.length}</span></p>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            )}

            {/* ══ PASO 4 — Resultado ══ */}
            {wizPaso === 4 && wizResultado && (() => {
              const { resumen_antes: ra, resumen_despues: rd, recomendacion, delta_costo_mcop, por_mes, perfil_horario } = wizResultado
              const pbProm = wizPBPromedioGlobal ?? 0
              const spreadNuevo = wizSimTipo === 'compra' ? pbProm - wizSimPrecio : wizSimPrecio - pbProm
              const semCfg: Record<SimRecomendacion, {border:string;color:string;icon:string;titulo:string}> = {
                verde:    { border: 'border-emerald-500/40 bg-emerald-500/10', color: 'text-emerald-400', icon: '✓', titulo: 'CONVIENE' },
                amarillo: { border: 'border-amber-500/40 bg-amber-500/10',     color: 'text-amber-400',   icon: '⚠', titulo: 'EVALUAR'  },
                rojo:     { border: 'border-red-500/40 bg-red-500/10',         color: 'text-red-400',     icon: '✗', titulo: 'NO CONVIENE en este escenario' },
              }
              const sem = semCfg[recomendacion]
              const ensoLabel = wizFuentePB === 'enso' ? (ENSO_COMPLETO.find(e => e.key === wizEnsoKey)?.label ?? 'ENSO') : `${wizPBDesde} → ${wizPBHasta}`
              const textoRec = recomendacion === 'verde'
                ? `El contrato de ${wizSimTipo} a ${wizSimPrecio} COP/kWh genera un spread favorable de +${formatNumero(Math.abs(spreadNuevo))} COP/kWh vs el PB del escenario ${ensoLabel}. Esto representa un ahorro/ingreso adicional de ${formatNumero(Math.abs(delta_costo_mcop), 2)} M COP en el período.`
                : recomendacion === 'amarillo'
                ? `El spread es marginal (${formatNumero(spreadNuevo, 1)} COP/kWh). Conviene si el PB se mantiene ${wizSimTipo === 'compra' ? 'por encima' : 'por debajo'} de ${wizSimPrecio} COP/kWh. El escenario muestra un PB promedio de ${formatNumero(pbProm)} COP/kWh.`
                : `Con el PB del escenario ${ensoLabel} (promedio ${formatNumero(pbProm)} COP/kWh), el contrato a ${wizSimPrecio} COP/kWh genera un spread negativo de ${formatNumero(spreadNuevo, 1)} COP/kWh. Sin embargo, si el PB ${wizSimTipo === 'compra' ? 'baja a menos de' : 'sube por encima de'} ${wizSimPrecio} COP/kWh, el contrato sería favorable.`
              const lineaColor = { verde: '#22c55e', amarillo: '#eab308', rojo: '#ef4444' }[recomendacion]
              return (
                <div className="space-y-5">
                  <div className="rounded-xl border border-gray-700 bg-gray-800/40 px-5 py-3">
                    <p className="text-xs font-semibold uppercase tracking-widest text-gray-400">Capítulo 3 — Impacto del nuevo contrato</p>
                  </div>
                  <div className="grid grid-cols-2 gap-3 sm:grid-cols-4">
                    <MetricCard label="Δ Posición Neta (MWh)" value={`${rd.posicion_neta_total_mwh - ra.posicion_neta_total_mwh >= 0 ? '+' : ''}${formatMiles(rd.posicion_neta_total_mwh - ra.posicion_neta_total_mwh, 0)}`} accent="text-white" />
                    <MetricCard label="Δ Costo Bolsa (M COP)" value={`${delta_costo_mcop >= 0 ? '+' : ''}${formatMiles(delta_costo_mcop, 2)}`} accent={delta_costo_mcop < 0 ? 'text-emerald-400' : 'text-red-400'} />
                    <MetricCard label="Spread nuevo vs PB" value={`${spreadNuevo >= 0 ? '+' : ''}${formatNumero(spreadNuevo)} COP/kWh`} accent={spreadNuevo >= 0 ? 'text-emerald-400' : 'text-red-400'} />
                    <MetricCard label="Energía contrato" value={`${formatMiles(wizSimPerfilTipo === 'bloques' ? wizSimBloques.reduce((s,b)=>s+b.mwhMes,0) : wizSimEnergiaKwh, 0)} kWh/mes`} accent="text-sky-400" />
                  </div>
                  <TablaAnalisis
                    encabezados={['Mes','Pos. Actual (MWh)','Pos. Nueva (MWh)','Δ (MWh)','Costo Actual (M COP)','Costo Nuevo (M COP)','Ahorro (M COP)']}
                    filas={<>
                      {por_mes.map(m => (
                        <tr key={m.mes} className="transition-colors hover:bg-gray-800/40">
                          <td className="px-4 py-2.5 font-medium text-gray-200">{m.mes}</td>
                          <td className={`px-4 py-2.5 text-right tabular-nums ${m.pos_actual_mwh > 0 ? 'text-red-400' : 'text-emerald-400'}`}>{formatMiles(m.pos_actual_mwh)}</td>
                          <td className={`px-4 py-2.5 text-right tabular-nums ${m.pos_nueva_mwh > 0 ? 'text-red-400' : 'text-emerald-400'}`}>{formatMiles(m.pos_nueva_mwh)}</td>
                          <td className={`px-4 py-2.5 text-right font-medium tabular-nums ${m.diferencia_mwh <= 0 ? 'text-emerald-400' : 'text-red-400'}`}>{m.diferencia_mwh > 0 ? '+' : ''}{formatMiles(m.diferencia_mwh)}</td>
                          <td className="px-4 py-2.5 text-right tabular-nums text-gray-300">{formatMiles(m.costo_actual_mcop)}</td>
                          <td className="px-4 py-2.5 text-right tabular-nums text-blue-300">{formatMiles(m.costo_nuevo_mcop)}</td>
                          <td className={`px-4 py-2.5 text-right font-semibold tabular-nums ${m.ahorro_mcop >= 0 ? 'text-emerald-400' : 'text-red-400'}`}>{m.ahorro_mcop > 0 ? '+' : ''}{formatMiles(m.ahorro_mcop)}</td>
                        </tr>
                      ))}
                      {(() => {
                        const totPosActual   = por_mes.reduce((s, m) => s + m.pos_actual_mwh,    0)
                        const totPosNueva    = por_mes.reduce((s, m) => s + m.pos_nueva_mwh,     0)
                        const totDiff        = por_mes.reduce((s, m) => s + m.diferencia_mwh,    0)
                        const totCostoActual = por_mes.reduce((s, m) => s + m.costo_actual_mcop, 0)
                        const totCostoNuevo  = por_mes.reduce((s, m) => s + m.costo_nuevo_mcop,  0)
                        const totAhorro      = por_mes.reduce((s, m) => s + m.ahorro_mcop,       0)
                        return (
                          <tr className="border-t-2 border-gray-600 bg-gray-700">
                            <td className="px-4 py-2.5 text-xs font-bold uppercase tracking-wider text-white">TOTAL PERÍODO</td>
                            <td className={`px-4 py-2.5 text-right font-bold tabular-nums ${totPosActual > 0 ? 'text-red-300' : 'text-emerald-300'}`}>{formatMiles(totPosActual)}</td>
                            <td className={`px-4 py-2.5 text-right font-bold tabular-nums ${totPosNueva > 0 ? 'text-red-300' : 'text-emerald-300'}`}>{formatMiles(totPosNueva)}</td>
                            <td className={`px-4 py-2.5 text-right font-bold tabular-nums ${totDiff <= 0 ? 'text-emerald-300' : 'text-red-300'}`}>{totDiff > 0 ? '+' : ''}{formatMiles(totDiff)}</td>
                            <td className="px-4 py-2.5 text-right font-bold tabular-nums text-white">{formatMiles(totCostoActual)}</td>
                            <td className="px-4 py-2.5 text-right font-bold tabular-nums text-white">{formatMiles(totCostoNuevo)}</td>
                            <td className={`px-4 py-2.5 text-right font-bold tabular-nums ${totAhorro >= 0 ? 'text-emerald-300' : 'text-red-300'}`}>{totAhorro > 0 ? '+' : ''}{formatMiles(totAhorro)}</td>
                          </tr>
                        )
                      })()}
                    </>}
                  />
                  <div className="rounded-xl border border-gray-800 bg-gray-900/60 p-5">
                    <p className="mb-1 text-sm font-medium text-gray-400">Posición Neta hora a hora — promedio del período</p>
                    <p className="mb-4 text-xs text-gray-600">Azul = posición actual · Punteada = con nuevo contrato</p>
                    <div className="h-64">
                      <ResponsiveContainer width="100%" height="100%">
                        <LineChart data={perfil_horario} margin={{ top: 4, right: 16, left: 0, bottom: 4 }}>
                          <CartesianGrid stroke="#374151" strokeDasharray="3 3" />
                          <XAxis dataKey="hora" type="number" domain={[1,24]} ticks={Array.from({length:24},(_,i)=>i+1)} stroke="#9ca3af" tick={{fill:'#9ca3af',fontSize:11}} />
                          <YAxis stroke="#9ca3af" tick={{fill:'#9ca3af',fontSize:11}} />
                          <Tooltip contentStyle={{backgroundColor:'#111827',border:'1px solid #374151',borderRadius:'8px',color:'#f3f4f6'}} labelFormatter={h=>`Hora ${h}`} formatter={(v,n)=>[`${formatMiles(Number(v),2)} MWh`,n]} />
                          <Legend wrapperStyle={{color:'#9ca3af',paddingTop:8}} />
                          <Line type="monotone" dataKey="posicion_antes_mwh"   name="Posición actual"      stroke="#3b82f6"   strokeWidth={2}   dot={false} />
                          <Line type="monotone" dataKey="posicion_despues_mwh" name="Con nuevo contrato"   stroke={lineaColor} strokeWidth={2.5} dot={false} strokeDasharray="6 3" />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>
                  </div>

                  <div className="rounded-xl border border-gray-700 bg-gray-800/40 px-5 py-3">
                    <p className="text-xs font-semibold uppercase tracking-widest text-gray-400">Capítulo 4 — Conclusión y recomendación</p>
                  </div>
                  {/* Narrativa de posición base de BIA */}
                  {(() => {
                    const posKwh = Math.abs(ra.posicion_neta_total_mwh * 1_000)
                    const costoIngreso = Math.abs(ra.costo_bolsa_total_mcop)
                    const esComprando = ra.posicion_neta_total_mwh > 0
                    const narrativa = esComprando
                      ? `BIA está comprando ${formatMiles(posKwh, 0)} kWh en bolsa a un costo estimado de ${formatMiles(costoIngreso, 2)} M COP con el escenario de PB seleccionado. Un contrato de compra a menos de ${formatNumero(pbProm)} COP/kWh reduciría este costo.`
                      : `BIA está vendiendo ${formatMiles(posKwh, 0)} kWh en bolsa generando un ingreso estimado de ${formatMiles(costoIngreso, 2)} M COP con el escenario de PB seleccionado. Un contrato de venta a más de ${formatNumero(pbProm)} COP/kWh aumentaría este ingreso.`
                    return (
                      <div className={`rounded-xl border px-5 py-4 ${esComprando ? 'border-red-500/30 bg-red-950/20' : 'border-emerald-500/30 bg-emerald-950/20'}`}>
                        <div className="flex items-start gap-3">
                          <span className={`shrink-0 text-lg font-bold ${esComprando ? 'text-red-400' : 'text-emerald-400'}`}>
                            {esComprando ? '↑' : '↓'}
                          </span>
                          <div>
                            <p className={`mb-1 text-xs font-semibold uppercase tracking-wider ${esComprando ? 'text-red-400' : 'text-emerald-400'}`}>
                              {esComprando ? 'Comprando en bolsa' : 'Vendiendo en bolsa'}
                            </p>
                            <p className="text-sm text-gray-300 leading-relaxed">{narrativa}</p>
                          </div>
                        </div>
                      </div>
                    )
                  })()}
                  <div className={`rounded-xl border px-5 py-4 ${sem.border}`}>
                    <div className="flex items-start gap-4">
                      <span className={`text-3xl font-bold ${sem.color}`}>{sem.icon}</span>
                      <div className="flex-1">
                        <p className={`text-lg font-bold ${sem.color}`}>{sem.titulo}</p>
                        <p className="mt-1 text-sm text-gray-300 leading-relaxed">{textoRec}</p>
                      </div>
                      <span className={`shrink-0 text-base font-semibold tabular-nums ${delta_costo_mcop < 0 ? 'text-emerald-400' : 'text-red-400'}`}>
                        {delta_costo_mcop < 0 ? '▼' : '▲'} {Math.abs(delta_costo_mcop).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})} M COP
                      </span>
                    </div>
                  </div>
                  <div className="rounded-xl border border-gray-700 bg-gray-950/60 px-5 py-4">
                    <div className="flex items-start gap-3">
                      <span className="shrink-0 text-blue-400">ℹ</span>
                      <div className="text-sm text-gray-300 leading-relaxed">
                        <p><span className="font-semibold text-white">Punto de equilibrio:</span>{' '}
                        el contrato es favorable cuando el PB{' '}
                        {wizSimTipo === 'compra' ? 'supera' : 'está por debajo de'}{' '}
                        <span className="font-semibold text-white">{wizSimPrecio} COP/kWh</span>.</p>
                        <p className="mt-1">El escenario analizado tiene PB promedio de{' '}
                        <span className="font-semibold text-white">{formatNumero(pbProm)} COP/kWh</span>,{' '}
                        {wizSimTipo === 'compra' ? (pbProm > wizSimPrecio ? 'por encima' : 'por debajo') : (pbProm < wizSimPrecio ? 'por debajo' : 'por encima')}{' '}
                        del punto de equilibrio.</p>
                      </div>
                    </div>
                  </div>
                  <div className="grid grid-cols-2 gap-3 lg:grid-cols-4">
                    <SimCompareCard label="Posición Neta (MWh)" antes={ra.posicion_neta_total_mwh} despues={rd.posicion_neta_total_mwh} formato={v=>formatMiles(v,0)} mejoraSiMenor={wizSimTipo==='venta'} />
                    <SimCompareCard label="Costo/Ingreso (M COP)" antes={ra.costo_bolsa_total_mcop} despues={rd.costo_bolsa_total_mcop} formato={v=>formatMiles(v,2)} mejoraSiMenor />
                    <SimCompareCard label="Hora pico compra" antes={ra.hora_pico_compra} despues={rd.hora_pico_compra} formato={v=>`H${v}`} soloInfo />
                    <SimCompareCard label="Hora pico venta"  antes={ra.hora_pico_venta}  despues={rd.hora_pico_venta}  formato={v=>`H${v}`} soloInfo />
                  </div>
                  <div className="flex flex-wrap items-center gap-3">
                    <button type="button" onClick={()=>setWizPaso(3)} className="rounded-lg border border-gray-700 px-5 py-2.5 text-sm font-semibold text-gray-300 transition hover:bg-gray-800">← Contrato</button>
                    <button type="button" onClick={wizDescargarJSON} className="rounded-lg border border-emerald-700 bg-emerald-950/40 px-5 py-2.5 text-sm font-semibold text-emerald-400 transition hover:bg-emerald-900/40">💾 Guardar simulación (JSON)</button>
                    <button type="button" onClick={wizReset} className="ml-auto rounded-lg bg-gray-700 px-5 py-2.5 text-sm font-semibold text-white transition hover:bg-gray-600">🔄 Nueva simulación</button>
                  </div>
                </div>
              )
            })()}
            {wizPaso === 4 && !wizResultado && (
              <div className="flex h-60 flex-col items-center justify-center rounded-xl border border-dashed border-gray-700">
                <p className="text-gray-500">Ejecuta la simulación en el Paso 3</p>
                <button type="button" onClick={()=>setWizPaso(3)} className="mt-3 text-sm text-emerald-400 hover:underline">← Volver al Paso 3</button>
              </div>
            )}
          </>
        )}
      </main>
    </div>
  )
}

function MensajeSinDatos({ children }: { children: ReactNode }) {
  return (
    <div className="rounded-lg border border-gray-800 bg-gray-950/60 px-4 py-8 text-center text-sm text-gray-500">
      {children}
    </div>
  )
}

/** Badge de color para indicar si el PPP es Indexado (verde), Proyectado (ámbar) o Sin datos (gris). */
function BadgeTipoPrecio({ tipo }: { tipo: 'Indexado' | 'Proyectado' | 'Sin datos' }) {
  const cfg: Record<string, string> = {
    'Indexado':   'bg-emerald-500/15 text-emerald-400 ring-emerald-500/40',
    'Proyectado': 'bg-amber-500/15   text-amber-400   ring-amber-500/40',
    'Sin datos':  'bg-gray-500/15    text-gray-400    ring-gray-500/40',
  }
  return (
    <span className={`inline-flex items-center rounded-full px-2 py-0.5 text-xs font-semibold ring-1 ${cfg[tipo] ?? cfg['Sin datos']}`}>
      {tipo}
    </span>
  )
}

function TablaAnalisis({
  encabezados,
  encabezadosClassName,
  filas,
}: {
  encabezados: string[]
  encabezadosClassName?: string[]
  filas: ReactNode
}) {
  return (
    <div className="overflow-x-auto rounded-lg border border-gray-800">
      <table className="w-full min-w-[480px] text-left text-sm">
        <thead>
          <tr className="border-b border-gray-800 bg-gray-900/60 text-xs uppercase tracking-wider text-gray-500">
            {encabezados.map((h, i) => (
              <th
                key={h}
                className={`px-4 py-3 font-medium ${
                  i === 0 ? '' : 'text-right'
                } ${encabezadosClassName?.[i] ?? ''}`}
              >
                {h}
              </th>
            ))}
          </tr>
        </thead>
        <tbody className="divide-y divide-gray-800/80">{filas}</tbody>
      </table>
    </div>
  )
}

function ResumenBloque({
  items,
}: {
  items: { label: string; value: string; accent?: string }[]
}) {
  return (
    <div className="grid grid-cols-1 gap-3 rounded-lg border border-gray-800 bg-gray-950/50 p-4 sm:grid-cols-3">
      {items.map((item) => (
        <div key={item.label}>
          <p className="text-xs uppercase tracking-wider text-gray-500">
            {item.label}
          </p>
          <p
            className={`mt-1 text-lg font-semibold tabular-nums ${item.accent ?? 'text-gray-100'}`}
          >
            {item.value}
          </p>
        </div>
      ))}
    </div>
  )
}

function MetricCard({
  label,
  value,
  subValue,
  accent = 'text-white',
}: {
  label: string
  value: string
  subValue?: string
  accent?: string
}) {
  return (
    <div className="rounded-xl border border-gray-800 bg-gray-900 p-5">
      <p className="text-xs font-medium uppercase tracking-wider text-gray-500">
        {label}
      </p>
      <p className={`mt-2 text-3xl font-bold tabular-nums ${accent}`}>
        {value}
        {subValue && (
          <span className="ml-2 text-lg font-semibold text-gray-400">
            ({subValue})
          </span>
        )}
      </p>
    </div>
  )
}

function SimCompareCard({
  label,
  antes,
  despues,
  formato,
  mejoraSiMenor = false,
  soloInfo = false,
}: {
  label: string
  antes: number
  despues: number
  formato: (v: number) => string
  mejoraSiMenor?: boolean
  soloInfo?: boolean
}) {
  const delta = despues - antes
  const sinCambio = delta === 0
  let colorDelta = 'text-gray-500'
  if (!soloInfo && !sinCambio) {
    const mejora = mejoraSiMenor ? delta < 0 : delta > 0
    colorDelta = mejora ? 'text-emerald-400' : 'text-red-400'
  }
  const signo = delta > 0 ? '+' : ''

  return (
    <div className="rounded-xl border border-gray-800 bg-gray-900 p-4">
      <p className="mb-3 text-xs font-medium uppercase tracking-wider text-gray-500">
        {label}
      </p>
      <div className="flex items-center gap-2">
        <span className="text-base font-semibold tabular-nums text-gray-400">
          {formato(antes)}
        </span>
        <span className="text-gray-600">→</span>
        <span className="text-lg font-bold tabular-nums text-gray-100">
          {formato(despues)}
        </span>
      </div>
      {!soloInfo && (
        <p className={`mt-1 text-sm font-medium tabular-nums ${colorDelta}`}>
          {sinCambio
            ? 'Sin cambio'
            : `${signo}${formato(delta)}`}
        </p>
      )}
    </div>
  )
}

function WizardProgressBar({ pasoActual }: { pasoActual: number }) {
  const pasos = [
    { n: 1, label: 'Período' },
    { n: 2, label: 'Escenario PB' },
    { n: 3, label: 'Nuevo contrato' },
    { n: 4, label: 'Resultado' },
  ]
  return (
    <nav className="overflow-hidden rounded-xl border border-gray-800 bg-gray-900/60 px-6 py-4">
      <ol className="flex items-center justify-between gap-2">
        {pasos.map((p, i) => {
          const activo     = p.n === pasoActual
          const completado = p.n < pasoActual
          return (
            <li key={p.n} className="flex flex-1 items-center">
              <div className="flex items-center gap-2">
                <span className={`flex h-7 w-7 shrink-0 items-center justify-center rounded-full text-xs font-bold transition ${completado ? 'bg-emerald-600 text-white' : activo ? 'bg-emerald-500 text-white shadow-md shadow-emerald-500/40' : 'bg-gray-800 text-gray-500'}`}>
                  {completado ? '✓' : p.n}
                </span>
                <span className={`text-sm font-medium whitespace-nowrap ${activo ? 'text-white' : completado ? 'text-emerald-400' : 'text-gray-500'}`}>
                  {p.label}
                </span>
              </div>
              {i < pasos.length - 1 && (
                <div className={`mx-3 h-px flex-1 transition ${p.n < pasoActual ? 'bg-emerald-600' : 'bg-gray-700'}`} />
              )}
            </li>
          )
        })}
      </ol>
    </nav>
  )
}

export default App
