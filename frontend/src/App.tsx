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

const API_BASE = 'http://127.0.0.1:8000/spread'
const API_PORTFOLIO = 'http://127.0.0.1:8000/portfolio'
const API_SIMULATE = 'http://127.0.0.1:8000/simulate'

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

const ENSO_OPCIONES = [
  { key: 'nino_2015',    label: 'El Niño Fuerte (2015-2016)',       inicio: '2015-03-01', fin: '2016-05-31' },
  { key: 'nina_2020',    label: 'La Niña Persistente (2020-2022)',   inicio: '2020-09-01', fin: '2022-03-31' },
  { key: 'nino_2023',    label: 'El Niño Fuerte (2023-2024)',        inicio: '2023-06-01', fin: '2024-05-31' },
  { key: 'neutral_2025', label: 'Neutral (2025)',                    inicio: '2025-01-01', fin: '2025-12-31' },
] as const

// Curva solar típica para Colombia (ecuatorial) — suma = 1.0
// H1-H6 y H19-H24 = 0; H7-H12 sube, H13-H18 baja
const SOLAR_PESOS_24H: number[] = [
  0,    0,    0,    0,    0,    0,     // H1-H6
  0.02, 0.06, 0.11, 0.15, 0.15, 0.14, // H7-H12
  0.12, 0.10, 0.08, 0.05, 0.02, 0,    // H13-H18
  0,    0,    0,    0,    0,    0,     // H19-H24
]

type VistaAnalisis = 'dia' | 'mes' | 'comparativo'
type DashboardTab = 'spread' | 'portafolio' | 'simulador'
type SimTipo = 'compra' | 'venta'
type SimMercado = 'regulado' | 'no_regulado' | 'ambos'
type SimFuentePB = 'historico' | 'enso'
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
  posicion_neta: number
  costo_bolsa: number
}

interface PortfolioResumen {
  posicion_neta_total_kwh: number
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

function formatNumeroMillonesCOP(valor: number): string {
  return formatNumero(valor / 1_000_000, 2)
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
  const compraR = parseNumero(fila.compra_r_kwh)
  const compraNr = parseNumero(fila.compra_nr_kwh)
  const venta = parseNumero(fila.venta_kwh)
  const posicionNeta =
    fila.posicion_neta_kwh != null
      ? parseNumero(fila.posicion_neta_kwh)
      : compraR + compraNr - venta
  return {
    fecha: String(fila.fecha ?? ''),
    hora,
    tipo_dia:
      typeof fila.tipo_dia === 'string' && fila.tipo_dia.trim() !== ''
        ? fila.tipo_dia
        : undefined,
    compra_r: compraR,
    compra_nr: compraNr,
    venta,
    posicion_neta: posicionNeta,
    costo_bolsa: parseNumero(fila.costo_bolsa_cop),
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
  const [fechaInicio, setFechaInicio] = useState('2026-01-01')
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

  // ── Simulador ──────────────────────────────────────────────────────────────
  const [simFuentePB, setSimFuentePB] = useState<SimFuentePB>('historico')
  const [simEnso, setSimEnso] = useState('nino_2023')
  const [simTipo, setSimTipo] = useState<SimTipo>('compra')
  const [simContraparte, setSimContraparte] = useState('')
  const [simEnergiaMwh, setSimEnergiaMwh] = useState(1000)
  const [simPrecio, setSimPrecio] = useState(350)
  // Estado A — Período del PB histórico (fuente de datos de bolsa)
  const [pbDesde, setPbDesde] = useState('2026-01-01')
  const [pbHasta, setPbHasta] = useState('2026-12-31')
  // Estado C — Vigencia del nuevo contrato simulado
  const [contratoInicio, setContratoInicio] = useState('2026-01-01')
  const [contratoFin, setContratoFin] = useState('2026-12-31')
  const [simTipoMercado, setSimTipoMercado] = useState<SimMercado>('regulado')
  const [simPerfilTipo, setSimPerfilTipo] = useState<SimPerfilTipo>('plano')
  const [simBloques, setSimBloques] = useState<SimBloque[]>([
    { horaInicio: 8, horaFin: 17, mwhMes: 1000 },
  ])
  const [simExcel12x24, setSimExcel12x24] = useState<number[][] | null>(null)
  const [simExcelNombre, setSimExcelNombre] = useState('')
  const [simResultado, setSimResultado] = useState<SimResultado | null>(null)
  const [simCargando, setSimCargando] = useState(false)
  const [simError, setSimError] = useState<string | null>(null)
  const simExcelInputRef = useRef<HTMLInputElement>(null)

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

      const respuesta = await fetch(
        `${API_PORTFOLIO}?fecha_inicio=${inicio}&fecha_fin=${fin}`,
        { signal },
      )

      if (!respuesta.ok) {
        const cuerpo = await respuesta.json().catch(() => null)
        const detalle =
          cuerpo && typeof cuerpo.detail === 'string'
            ? cuerpo.detail
            : `Error ${respuesta.status}`
        throw new Error(detalle)
      }

      const json = await respuesta.json()
      const candidatos = Array.isArray(json?.datos)
        ? json.datos
        : Array.isArray(json)
          ? json
          : []
      setPortfolioDatos(candidatos.map(parseFilaPortfolio).filter(Boolean) as FilaPortfolio[])
    } catch (err) {
      if (err instanceof Error && err.name === 'AbortError') return
      setPortfolioError(
        err instanceof Error ? err.message : 'No se pudieron cargar los datos',
      )
      setPortfolioDatos([])
    } finally {
      setPortfolioCargando(false)
    }
  }

  useEffect(() => {
    const controller = new AbortController()
    cargarSpread(350, '2026-01-01', new Date().toISOString().split('T')[0], {
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

  // Auto-fill del período de PB cuando se elige un escenario ENSO
  useEffect(() => {
    if (simFuentePB !== 'enso') return
    const enso = ENSO_OPCIONES.find((e) => e.key === simEnso)
    if (enso) {
      setPbDesde(enso.inicio)
      setPbHasta(enso.fin)
    }
  }, [simFuentePB, simEnso])

  function addBloque() {
    if (simBloques.length >= 3) return
    setSimBloques((prev) => [...prev, { horaInicio: 18, horaFin: 22, mwhMes: 500 }])
  }

  function removeBloque(idx: number) {
    setSimBloques((prev) => prev.filter((_, i) => i !== idx))
  }

  function updateBloque(idx: number, campo: keyof SimBloque, valor: number) {
    setSimBloques((prev) =>
      prev.map((b, i) =>
        i === idx
          ? {
              ...b,
              [campo]: valor,
              // Ensure horaFin >= horaInicio
              ...(campo === 'horaInicio' && valor > b.horaFin
                ? { horaFin: valor }
                : {}),
            }
          : b,
      ),
    )
  }

  function handleExcelUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (!file) return
    setSimExcelNombre(file.name)
    const reader = new FileReader()
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target?.result, { type: 'binary' })
        const ws = wb.Sheets[wb.SheetNames[0]]
        const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 }) as unknown[][]
        // Skip header row if first cell is text (e.g. "Mes", "H1", etc.)
        const dataRows = rows.filter(
          (r) => Array.isArray(r) && r.length >= 24 && typeof r[0] === 'number',
        )
        if (dataRows.length < 1) {
          setSimError(
            'El archivo no tiene filas numéricas válidas. Asegúrate de que filas = meses y columnas = H1..H24 con valores en kWh.',
          )
          return
        }
        const matrix: number[][] = dataRows.slice(0, 12).map((row) =>
          Array.from({ length: 24 }, (_, i) => parseFloat(String((row as unknown[])[i] ?? 0)) || 0),
        )
        while (matrix.length < 12) matrix.push([...matrix[matrix.length - 1]])
        setSimExcel12x24(matrix)
        setSimError(null)
      } catch {
        setSimError('Error al leer el archivo Excel. Verifica el formato.')
      }
    }
    reader.readAsBinaryString(file)
    // Reset input so same file can be re-uploaded
    e.target.value = ''
  }

  async function cargarSimulacion() {
    // ── Validate date ranges ──────────────────────────────────────────────────
    const diffMesesPB = (() => {
      const d1 = new Date(pbDesde)
      const d2 = new Date(pbHasta)
      return (d2.getFullYear() - d1.getFullYear()) * 12 + (d2.getMonth() - d1.getMonth()) + 1
    })()
    if (diffMesesPB < 1) {
      setSimError('El período de PB histórico debe tener al menos 1 mes.')
      return
    }
    if (diffMesesPB > 36) {
      setSimError(
        `El período de PB histórico tiene ${diffMesesPB} meses. ` +
        'Usa un rango de máximo 36 meses para evitar timeout — ' +
        'o selecciona un escenario ENSO predefinido.',
      )
      return
    }

    // ── Validate profile configuration ───────────────────────────────────────
    if (simPerfilTipo === 'excel' && !simExcel12x24) {
      setSimError('Carga un archivo Excel antes de simular.')
      return
    }
    if (
      (simPerfilTipo === 'plano' || simPerfilTipo === 'solar') &&
      simEnergiaMwh <= 0
    ) {
      setSimError('Ingresa una energía mensual mayor que 0.')
      return
    }

    try {
      setSimCargando(true)
      setSimError(null)

      // ── Build profile fields for the API ───────────────────────────────────
      // The backend handles all profile logic; the frontend just ships the
      // raw inputs (bloques list, excel matrix, energy amount).
      type SimBody = {
        tipo: string
        contraparte: string
        precio_cop_kwh: number
        pb_desde: string
        pb_hasta: string
        contrato_inicio: string
        contrato_fin: string
        tipo_mercado: string
        perfil_horario: string
        energia_mensual_mwh?: number
        bloques?: { hora_ini: number; hora_fin: number; mwh_mes: number }[]
        perfil_excel_12x24?: number[][]
      }

      const body: SimBody = {
        tipo: simTipo,
        contraparte: simContraparte,
        precio_cop_kwh: simPrecio,
        pb_desde: pbDesde,
        pb_hasta: pbHasta,
        contrato_inicio: contratoInicio,
        contrato_fin: contratoFin,
        tipo_mercado: simTipoMercado,
        perfil_horario: simPerfilTipo === 'excel' ? 'excel' : simPerfilTipo,
      }

      if (simPerfilTipo === 'plano' || simPerfilTipo === 'solar') {
        body.energia_mensual_mwh = simEnergiaMwh
      } else if (simPerfilTipo === 'bloques') {
        body.bloques = simBloques.map((b) => ({
          hora_ini: b.horaInicio,
          hora_fin: b.horaFin,
          mwh_mes: b.mwhMes,
        }))
      } else if (simPerfilTipo === 'excel' && simExcel12x24) {
        body.perfil_excel_12x24 = simExcel12x24
      }

      const respuesta = await fetch(API_SIMULATE, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      })
      if (!respuesta.ok) {
        const cuerpo = await respuesta.json().catch(() => null)
        throw new Error(
          typeof cuerpo?.detail === 'string'
            ? cuerpo.detail
            : `Error ${respuesta.status}`,
        )
      }
      const json: SimResultado = await respuesta.json()
      setSimResultado(json)
    } catch (err) {
      setSimError(err instanceof Error ? err.message : 'Error al simular')
      setSimResultado(null)
    } finally {
      setSimCargando(false)
    }
  }

  function toggleMesComparativo(mes: string) {
    setMesesComparativo((prev) => {
      if (prev.includes(mes)) return prev.filter((m) => m !== mes)
      if (prev.length >= 4) return prev
      return [...prev, mes].sort()
    })
  }

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
    const posicion_neta_total_kwh = portfolioDatos.reduce(
      (sum, f) => sum + f.posicion_neta,
      0,
    )
    const costo_bolsa_total_cop = portfolioDatos.reduce(
      (sum, f) => sum + f.costo_bolsa,
      0,
    )
    const porHora = Array.from({ length: 24 }, (_, i) => i + 1).map((hora) => {
      const filas = portfolioDatos.filter((f) => f.hora === hora)
      if (filas.length === 0) return { hora, compra: 0, venta: 0 }
      return {
        hora,
        compra: filas.reduce((s, f) => s + f.compra_r + f.compra_nr, 0) / filas.length,
        venta: filas.reduce((s, f) => s + f.venta, 0) / filas.length,
      }
    })
    const horaPicoCompra = porHora.reduce((a, b) => (a.compra >= b.compra ? a : b))
    const horaPicoVenta = porHora.reduce((a, b) => (a.venta >= b.venta ? a : b))
    return {
      posicion_neta_total_kwh,
      costo_bolsa_total_cop,
      hora_pico_compra: horaPicoCompra.compra > 0 ? horaPicoCompra.hora : null,
      hora_pico_venta: horaPicoVenta.venta > 0 ? horaPicoVenta.hora : null,
    }
  }, [portfolioDatos])

  const portfolioPorHora = useMemo(() => {
    return Array.from({ length: 24 }, (_, i) => {
      const hora = i + 1
      const filas = portfolioDatos.filter((f) => f.hora === hora)
      if (filas.length === 0) {
        return { hora, compra_r: 0, compra_nr: 0, venta: 0, posicion_neta: 0 }
      }
      return {
        hora,
        compra_r: filas.reduce((s, f) => s + f.compra_r, 0) / filas.length / 1000,
        compra_nr: filas.reduce((s, f) => s + f.compra_nr, 0) / filas.length / 1000,
        venta: filas.reduce((s, f) => s + f.venta, 0) / filas.length / 1000,
        posicion_neta: filas.reduce((s, f) => s + f.posicion_neta, 0) / filas.length / 1000,
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
        const compraR = filas.reduce((s, f) => s + f.compra_r, 0) / 1000
        const compraNr = filas.reduce((s, f) => s + f.compra_nr, 0) / 1000
        const venta = filas.reduce((s, f) => s + f.venta, 0) / 1000
        const posicionNeta = filas.reduce((s, f) => s + f.posicion_neta, 0) / 1000
        const costoBolsa = filas.reduce((s, f) => s + f.costo_bolsa, 0) / 1_000_000
        const tipoDia = tipoDiaDesdefecha(fecha)
        return { fecha, tipoDia, compraR, compraNr, venta, posicionNeta, costoBolsa }
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
        compraR: acc.compraR + f.compraR,
        compraNr: acc.compraNr + f.compraNr,
        venta: acc.venta + f.venta,
        posicionNeta: acc.posicionNeta + f.posicionNeta,
        costoBolsa: acc.costoBolsa + f.costoBolsa,
      }),
      { compraR: 0, compraNr: 0, venta: 0, posicionNeta: 0, costoBolsa: 0 },
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
      const tipo = tipoDiaDesdefecha(fila.fecha)
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

            <section className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-4 xl:grid-cols-5">
              <MetricCard
                label="Posición Neta (MWh)"
                value={
                  portfolioResumen
                    ? `${formatMiles(portfolioResumen.posicion_neta_total_kwh / 1000)}`
                    : '—'
                }
                accent={
                  portfolioResumen && portfolioResumen.posicion_neta_total_kwh <= 0
                    ? 'text-emerald-400'
                    : 'text-red-400'
                }
              />
              <MetricCard
                label="Costo/Ingreso Bolsa (M COP)"
                value={
                  portfolioResumen
                    ? `${formatMiles(portfolioResumen.costo_bolsa_total_cop / 1_000_000)}`
                    : '—'
                }
                accent={
                  portfolioResumen && portfolioResumen.costo_bolsa_total_cop <= 0
                    ? 'text-emerald-400'
                    : 'text-red-400'
                }
              />
              <MetricCard
                label="Hora Pico Compra"
                value={
                  portfolioResumen?.hora_pico_compra
                    ? `H${portfolioResumen.hora_pico_compra}`
                    : '—'
                }
                accent="text-sky-400"
              />
              <MetricCard
                label="Hora Pico Venta"
                value={
                  portfolioResumen?.hora_pico_venta
                    ? `H${portfolioResumen.hora_pico_venta}`
                    : '—'
                }
                accent="text-rose-400"
              />
              <div className="rounded-xl border border-dashed border-amber-500/30 bg-amber-500/5 p-5">
                <p className="text-xs font-medium uppercase tracking-wider text-amber-600">
                  PPP Contratos
                </p>
                <p className="mt-2 text-lg font-bold text-amber-500/60">Pendiente</p>
                <p className="mt-1 text-xs text-gray-600 leading-relaxed">
                  Requiere precios por contrato desde Olibia. Se calculará como
                  Σ(energía × precio) / Σ(energía).
                </p>
              </div>
            </section>

            <section className="rounded-xl border border-gray-800 bg-gray-900/60 p-6">
              <h2 className="mb-6 text-lg font-semibold text-gray-200">
                Posición horaria del portafolio
              </h2>
              <div className="h-80 w-full">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart
                    data={portfolioPorHora}
                    margin={{ top: 8, right: 48, left: 0, bottom: 8 }}
                  >
                    <CartesianGrid stroke="#374151" strokeDasharray="3 3" />
                    <XAxis
                      dataKey="hora"
                      stroke="#9ca3af"
                      tick={{ fill: '#9ca3af', fontSize: 12 }}
                    />
                    <YAxis
                      yAxisId="left"
                      stroke="#9ca3af"
                      tick={{ fill: '#9ca3af', fontSize: 12 }}
                    />
                    <YAxis
                      yAxisId="right"
                      orientation="right"
                      stroke="#eab308"
                      tick={{ fill: '#eab308', fontSize: 12 }}
                    />
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
                    <Bar
                      yAxisId="left"
                      dataKey="compra_r"
                      stackId="compras"
                      name="Compra R"
                      fill="#3b82f6"
                    />
                    <Bar
                      yAxisId="left"
                      dataKey="compra_nr"
                      stackId="compras"
                      name="Compra NR"
                      fill="#22c55e"
                    />
                    <Bar yAxisId="left" dataKey="venta" name="Venta" fill="#ef4444" />
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
                      'Compra R (MWh)',
                      'Compra NR (MWh)',
                      'Venta (MWh)',
                      'Posición Neta (MWh)',
                      'Costo Bolsa (M COP)',
                    ]}
                    filas={
                      <>
                        {portfolioResumenDiarioFiltrado.map((fila) => (
                          <tr
                            key={fila.fecha}
                            className="transition-colors hover:bg-gray-800/40"
                          >
                            <td className="px-4 py-2.5 text-gray-300">
                              {formatFechaLegible(fila.fecha)}
                            </td>
                            <td className="px-4 py-2.5 text-right text-gray-300">
                              {fila.tipoDia}
                            </td>
                            <td className="px-4 py-2.5 text-right tabular-nums text-sky-300">
                              {formatMiles(fila.compraR)}
                            </td>
                            <td className="px-4 py-2.5 text-right tabular-nums text-emerald-300">
                              {formatMiles(fila.compraNr)}
                            </td>
                            <td className="px-4 py-2.5 text-right tabular-nums text-rose-300">
                              {formatMiles(fila.venta)}
                            </td>
                            <td
                              className={`px-4 py-2.5 text-right font-medium tabular-nums ${
                                fila.posicionNeta <= 0 ? 'text-emerald-400' : 'text-red-400'
                              }`}
                            >
                              {formatMiles(fila.posicionNeta)}
                            </td>
                            <td className="px-4 py-2.5 text-right tabular-nums text-gray-200">
                              {formatMiles(fila.costoBolsa)}
                            </td>
                          </tr>
                        ))}
                        {portfolioTotalesFiltrados && (
                          <tr className="border-t-2 border-gray-700 bg-gray-800/60">
                            <td className="px-4 py-2.5 text-xs font-semibold uppercase tracking-wider text-gray-400">
                              TOTAL MES
                            </td>
                            <td className="px-4 py-2.5" />
                            <td className="px-4 py-2.5 text-right font-semibold tabular-nums text-sky-300">
                              {formatMiles(portfolioTotalesFiltrados.compraR)}
                            </td>
                            <td className="px-4 py-2.5 text-right font-semibold tabular-nums text-emerald-300">
                              {formatMiles(portfolioTotalesFiltrados.compraNr)}
                            </td>
                            <td className="px-4 py-2.5 text-right font-semibold tabular-nums text-rose-300">
                              {formatMiles(portfolioTotalesFiltrados.venta)}
                            </td>
                            <td
                              className={`px-4 py-2.5 text-right font-semibold tabular-nums ${
                                portfolioTotalesFiltrados.posicionNeta <= 0
                                  ? 'text-emerald-400'
                                  : 'text-red-400'
                              }`}
                            >
                              {formatMiles(portfolioTotalesFiltrados.posicionNeta)}
                            </td>
                            <td className="px-4 py-2.5 text-right font-semibold tabular-nums text-gray-200">
                              {formatMiles(portfolioTotalesFiltrados.costoBolsa)}
                            </td>
                          </tr>
                        )}
                      </>
                    }
                  />
                )}
              </div>
            </section>
          </>
        )}

        {/* ══════════════════════════════════════════════════════════════════
            TAB: SIMULADOR  (redesign v2)
        ══════════════════════════════════════════════════════════════════ */}
        {tabActiva === 'simulador' && (
          <>
            {/* Header */}
            <div>
              <h2 className="text-xl font-bold text-gray-100">Simulador</h2>
              <p className="mt-0.5 text-sm text-gray-500">
                Evalúa el impacto de un nuevo contrato sobre el portafolio Olibia
              </p>
            </div>

            {/* ── Two-column layout: form (left) + results (right) ─────── */}
            <div className="grid grid-cols-1 gap-6 lg:grid-cols-5">

              {/* ═══════════════════════════════════════════════════════════
                  LEFT PANEL — form
              ═══════════════════════════════════════════════════════════ */}
              <div className="space-y-4 lg:col-span-2">

                {/* ── PARTE A: Escenario de PB ─────────────────────────── */}
                <section className="rounded-xl border border-gray-800 bg-gray-900/60 p-5">
                  <p className="mb-4 text-xs font-semibold uppercase tracking-widest text-emerald-500">
                    A — Escenario de precio de bolsa
                  </p>

                  {/* Toggle Histórico / ENSO */}
                  <div className="mb-4 inline-flex w-full rounded-lg border border-gray-700 bg-gray-950 p-1">
                    {(['historico', 'enso'] as SimFuentePB[]).map((f) => (
                      <button
                        key={f}
                        type="button"
                        onClick={() => setSimFuentePB(f)}
                        className={`flex-1 rounded-md py-2 text-sm font-semibold transition ${
                          simFuentePB === f
                            ? 'bg-emerald-600 text-white shadow-sm'
                            : 'text-gray-400 hover:bg-gray-800 hover:text-gray-200'
                        }`}
                      >
                        {f === 'historico' ? '📅 PB Histórico' : '🌡 Fenómeno ENSO'}
                      </button>
                    ))}
                  </div>

                  {simFuentePB === 'historico' ? (
                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="mb-1 block text-xs text-gray-500">Desde</label>
                        <input
                          type="date"
                          min="2010-01-01"
                          value={pbDesde}
                          onChange={(e) => setPbDesde(e.target.value)}
                          className="w-full rounded-lg border border-gray-700 bg-gray-950 px-3 py-2 text-sm text-gray-100 outline-none transition [color-scheme:dark] focus:border-emerald-500/50 focus:ring-2 focus:ring-emerald-500/30"
                        />
                      </div>
                      <div>
                        <label className="mb-1 block text-xs text-gray-500">Hasta</label>
                        <input
                          type="date"
                          value={pbHasta}
                          onChange={(e) => setPbHasta(e.target.value)}
                          className="w-full rounded-lg border border-gray-700 bg-gray-950 px-3 py-2 text-sm text-gray-100 outline-none transition [color-scheme:dark] focus:border-emerald-500/50 focus:ring-2 focus:ring-emerald-500/30"
                        />
                      </div>
                    </div>
                  ) : (
                    <div>
                      <label className="mb-1 block text-xs text-gray-500">
                        Período de referencia
                      </label>
                      <select
                        value={simEnso}
                        onChange={(e) => setSimEnso(e.target.value)}
                        className={CLASE_SELECT + ' w-full'}
                      >
                        {ENSO_OPCIONES.map((e) => (
                          <option key={e.key} value={e.key}>
                            {e.label}
                          </option>
                        ))}
                      </select>
                      {(() => {
                        const enso = ENSO_OPCIONES.find((e) => e.key === simEnso)
                        return enso ? (
                          <p className="mt-2 text-xs text-gray-600">
                            📌 {enso.inicio} → {enso.fin}
                          </p>
                        ) : null
                      })()}
                    </div>
                  )}
                </section>

                {/* ── PARTE C: Nuevo contrato ──────────────────────────── */}
                <section className="rounded-xl border border-gray-800 bg-gray-900/60 p-5">
                  <p className="mb-4 text-xs font-semibold uppercase tracking-widest text-blue-400">
                    C — Nuevo contrato
                  </p>

                  {/* Tipo Compra / Venta */}
                  <div className="mb-4">
                    <label className="mb-1.5 block text-xs text-gray-500">
                      Tipo de operación
                    </label>
                    <div className="inline-flex w-full rounded-lg border border-gray-700 bg-gray-950 p-1">
                      {(['compra', 'venta'] as SimTipo[]).map((t) => (
                        <button
                          key={t}
                          type="button"
                          onClick={() => setSimTipo(t)}
                          className={`flex-1 rounded-md py-2 text-sm font-semibold capitalize transition ${
                            simTipo === t
                              ? t === 'compra'
                                ? 'bg-blue-600 text-white shadow-sm'
                                : 'bg-rose-600 text-white shadow-sm'
                              : 'text-gray-400 hover:bg-gray-800 hover:text-gray-200'
                          }`}
                        >
                          {t === 'compra' ? '↓ Compra' : '↑ Venta'}
                        </button>
                      ))}
                    </div>
                  </div>

                  {/* Contraparte + Precio */}
                  <div className="mb-3">
                    <label className="mb-1 block text-xs text-gray-500">Contraparte</label>
                    <input
                      type="text"
                      placeholder="Nombre de la contraparte"
                      value={simContraparte}
                      onChange={(e) => setSimContraparte(e.target.value)}
                      className="w-full rounded-lg border border-gray-700 bg-gray-950 px-3 py-2 text-sm text-gray-100 placeholder-gray-600 outline-none transition focus:border-emerald-500/50 focus:ring-2 focus:ring-emerald-500/30"
                    />
                  </div>
                  <div className="mb-4">
                    <label className="mb-1 block text-xs text-gray-500">
                      Precio (COP/kWh)
                    </label>
                    <input
                      type="number"
                      min={1}
                      value={simPrecio}
                      onChange={(e) =>
                        setSimPrecio(Math.max(1, Number(e.target.value) || 1))
                      }
                      className="w-full rounded-lg border border-gray-700 bg-gray-950 px-3 py-2 text-sm text-gray-100 outline-none transition focus:border-emerald-500/50 focus:ring-2 focus:ring-emerald-500/30"
                    />
                  </div>

                  {/* Perfil horario */}
                  <div className="mb-4">
                    <label className="mb-2 block text-xs text-gray-500">
                      Distribución horaria
                    </label>
                    <div className="grid grid-cols-4 gap-1.5">
                      {(
                        [
                          ['plano', 'Plano 24h'],
                          ['bloques', 'Bloques'],
                          ['solar', 'Solar'],
                          ['excel', 'Excel'],
                        ] as [SimPerfilTipo, string][]
                      ).map(([p, lbl]) => (
                        <button
                          key={p}
                          type="button"
                          onClick={() => setSimPerfilTipo(p)}
                          className={`rounded-lg py-2 text-xs font-semibold transition ${
                            simPerfilTipo === p
                              ? 'bg-gray-600 text-white ring-1 ring-gray-400'
                              : 'bg-gray-800 text-gray-400 hover:bg-gray-700 hover:text-gray-200'
                          }`}
                        >
                          {lbl}
                        </button>
                      ))}
                    </div>

                    {/* Profile-specific inputs */}
                    <div className="mt-3">
                      {(simPerfilTipo === 'plano' || simPerfilTipo === 'solar') && (
                        <div>
                          <label className="mb-1 block text-xs text-gray-500">
                            Energía mensual (MWh/mes)
                          </label>
                          <input
                            type="number"
                            min={1}
                            value={simEnergiaMwh}
                            onChange={(e) =>
                              setSimEnergiaMwh(Math.max(1, Number(e.target.value) || 1))
                            }
                            className="w-full rounded-lg border border-gray-700 bg-gray-950 px-3 py-2 text-sm text-gray-100 outline-none transition focus:border-emerald-500/50 focus:ring-2 focus:ring-emerald-500/30"
                          />
                          {simPerfilTipo === 'solar' && (
                            <p className="mt-1.5 text-xs text-gray-600">
                              Curva solar: 0 en H1-H6 y H19-H24 · sube H7-H12 · baja H13-H18
                            </p>
                          )}
                        </div>
                      )}

                      {simPerfilTipo === 'bloques' && (
                        <div className="space-y-2">
                          {simBloques.map((b, idx) => (
                            <div
                              key={idx}
                              className="flex items-end gap-2 rounded-lg border border-gray-700 bg-gray-950/60 p-3"
                            >
                              <div className="flex flex-col gap-1">
                                <label className="text-xs text-gray-600">H ini</label>
                                <select
                                  value={b.horaInicio}
                                  onChange={(e) =>
                                    updateBloque(idx, 'horaInicio', Number(e.target.value))
                                  }
                                  className="w-16 rounded border border-gray-700 bg-gray-900 px-2 py-1.5 text-xs text-gray-100 outline-none"
                                >
                                  {Array.from({ length: 24 }, (_, i) => i + 1).map((h) => (
                                    <option key={h} value={h}>
                                      H{h}
                                    </option>
                                  ))}
                                </select>
                              </div>
                              <div className="flex flex-col gap-1">
                                <label className="text-xs text-gray-600">H fin</label>
                                <select
                                  value={b.horaFin}
                                  onChange={(e) =>
                                    updateBloque(idx, 'horaFin', Number(e.target.value))
                                  }
                                  className="w-16 rounded border border-gray-700 bg-gray-900 px-2 py-1.5 text-xs text-gray-100 outline-none"
                                >
                                  {Array.from({ length: 24 }, (_, i) => i + 1)
                                    .filter((h) => h >= b.horaInicio)
                                    .map((h) => (
                                      <option key={h} value={h}>
                                        H{h}
                                      </option>
                                    ))}
                                </select>
                              </div>
                              <div className="flex flex-1 flex-col gap-1">
                                <label className="text-xs text-gray-600">MWh/mes</label>
                                <input
                                  type="number"
                                  min={1}
                                  value={b.mwhMes}
                                  onChange={(e) =>
                                    updateBloque(idx, 'mwhMes', Math.max(1, Number(e.target.value) || 1))
                                  }
                                  className="w-full rounded border border-gray-700 bg-gray-900 px-2 py-1.5 text-xs text-gray-100 outline-none"
                                />
                              </div>
                              {simBloques.length > 1 && (
                                <button
                                  type="button"
                                  onClick={() => removeBloque(idx)}
                                  className="rounded px-1.5 py-1.5 text-gray-600 hover:text-red-400 transition"
                                >
                                  ✕
                                </button>
                              )}
                            </div>
                          ))}
                          {simBloques.length < 3 && (
                            <button
                              type="button"
                              onClick={addBloque}
                              className="w-full rounded-lg border border-dashed border-gray-700 py-2 text-xs text-gray-500 transition hover:border-gray-500 hover:text-gray-300"
                            >
                              + Agregar bloque
                            </button>
                          )}
                          <p className="text-xs text-gray-600">
                            Total:{' '}
                            {formatMiles(
                              simBloques.reduce((s, b) => s + b.mwhMes, 0),
                              0,
                            )}{' '}
                            MWh/mes
                          </p>
                        </div>
                      )}

                      {simPerfilTipo === 'excel' && (
                        <div>
                          <input
                            ref={simExcelInputRef}
                            type="file"
                            accept=".xlsx,.xls,.csv"
                            className="hidden"
                            onChange={handleExcelUpload}
                          />
                          <button
                            type="button"
                            onClick={() => simExcelInputRef.current?.click()}
                            className="w-full rounded-lg border border-dashed border-gray-600 py-3 text-sm text-gray-400 transition hover:border-emerald-500/50 hover:text-emerald-400"
                          >
                            📎 Cargar archivo Excel / CSV
                          </button>
                          {simExcelNombre && (
                            <p className="mt-1.5 flex items-center gap-1.5 text-xs text-emerald-400">
                              <span>✓</span>
                              {simExcelNombre}
                              {simExcel12x24 && (
                                <span className="text-gray-600">
                                  ({simExcel12x24.length} meses)
                                </span>
                              )}
                            </p>
                          )}
                          <p className="mt-2 text-xs text-gray-600 leading-relaxed">
                            Plantilla: filas = meses (Ene–Dic), columnas = H1..H24, valores en kWh.
                            Exportar desde Excel como CSV o .xlsx.
                          </p>
                        </div>
                      )}
                    </div>
                  </div>

                  {/* Vigencia y mercado */}
                  <div className="mb-4 grid grid-cols-2 gap-3">
                    <div>
                      <label className="mb-1 block text-xs text-gray-500">Inicio vigencia</label>
                      <input
                        type="date"
                        value={contratoInicio}
                        onChange={(e) => setContratoInicio(e.target.value)}
                        className="w-full rounded-lg border border-gray-700 bg-gray-950 px-3 py-2 text-sm text-gray-100 outline-none transition [color-scheme:dark] focus:border-emerald-500/50 focus:ring-2 focus:ring-emerald-500/30"
                      />
                    </div>
                    <div>
                      <label className="mb-1 block text-xs text-gray-500">Fin vigencia</label>
                      <input
                        type="date"
                        value={contratoFin}
                        onChange={(e) => setContratoFin(e.target.value)}
                        className="w-full rounded-lg border border-gray-700 bg-gray-950 px-3 py-2 text-sm text-gray-100 outline-none transition [color-scheme:dark] focus:border-emerald-500/50 focus:ring-2 focus:ring-emerald-500/30"
                      />
                    </div>
                  </div>
                  <div className="mb-5">
                    <label className="mb-1 block text-xs text-gray-500">Tipo de mercado</label>
                    <select
                      value={simTipoMercado}
                      onChange={(e) => setSimTipoMercado(e.target.value as SimMercado)}
                      className={CLASE_SELECT + ' w-full'}
                    >
                      <option value="regulado">Regulado (Compra R)</option>
                      <option value="no_regulado">No Regulado (Compra NR)</option>
                      <option value="ambos">Ambos (50 % R + 50 % NR)</option>
                    </select>
                  </div>

                  {/* CTA */}
                  <button
                    type="button"
                    onClick={cargarSimulacion}
                    disabled={simCargando}
                    className="w-full rounded-lg bg-emerald-600 py-3 text-sm font-bold text-white shadow transition hover:bg-emerald-500 disabled:cursor-not-allowed disabled:opacity-60"
                  >
                    {simCargando ? (
                      <span className="flex items-center justify-center gap-2">
                        <span className="h-4 w-4 animate-spin rounded-full border-2 border-white border-t-transparent" />
                        Simulando…
                      </span>
                    ) : (
                      '⚡ Simular impacto'
                    )}
                  </button>
                  {simError && (
                    <div className="mt-3 rounded-lg border border-red-900/50 bg-red-950/30 px-4 py-3 text-sm text-red-400">
                      {simError}
                    </div>
                  )}
                </section>
              </div>

              {/* ═══════════════════════════════════════════════════════════
                  RIGHT PANEL — results
              ═══════════════════════════════════════════════════════════ */}
              <div className="space-y-4 lg:col-span-3">
                {!simResultado && !simCargando && (
                  <div className="flex h-80 flex-col items-center justify-center rounded-xl border border-dashed border-gray-700 bg-gray-900/30 text-center">
                    <span className="mb-3 text-3xl">⚡</span>
                    <p className="text-sm text-gray-400">Configura el escenario y el contrato</p>
                    <p className="mt-1 text-xs text-gray-600">
                      Pulsa «Simular impacto» para ver los resultados comparativos
                    </p>
                  </div>
                )}

                {simResultado && (() => {
                  const { resumen_antes: ra, resumen_despues: rd, recomendacion, delta_costo_mcop } = simResultado
                  const semCfg: Record<SimRecomendacion, { borderBg: string; color: string; icon: string; titulo: string; desc: string }> = {
                    verde: {
                      borderBg: 'border-emerald-500/40 bg-emerald-500/10',
                      color: 'text-emerald-400',
                      icon: '✓',
                      titulo: 'CONVIENE',
                      desc: 'El contrato mejora la posición y reduce el costo de bolsa.',
                    },
                    amarillo: {
                      borderBg: 'border-amber-500/40 bg-amber-500/10',
                      color: 'text-amber-400',
                      icon: '⚠',
                      titulo: 'EVALUAR',
                      desc: 'Impacto mixto: mejora algunos indicadores pero no todos.',
                    },
                    rojo: {
                      borderBg: 'border-red-500/40 bg-red-500/10',
                      color: 'text-red-400',
                      icon: '✗',
                      titulo: 'NO CONVIENE',
                      desc: 'El contrato empeora la posición y aumenta el costo de bolsa.',
                    },
                  }
                  const sem = semCfg[recomendacion]

                  return (
                    <>
                      {/* PARTE B — Posición base */}
                      <div className="rounded-xl border border-gray-800 bg-gray-900/60 p-4">
                        <p className="mb-3 text-xs font-semibold uppercase tracking-widest text-gray-500">
                          B — Posición base (sin contrato nuevo)
                        </p>
                        <div className="grid grid-cols-2 gap-3">
                          <div className="rounded-lg bg-gray-950/50 px-4 py-3">
                            <p className="text-xs text-gray-500">Posición Neta</p>
                            <p className={`mt-1 text-xl font-bold tabular-nums ${ra.posicion_neta_total_mwh <= 0 ? 'text-emerald-400' : 'text-red-400'}`}>
                              {formatMiles(ra.posicion_neta_total_mwh, 0)} <span className="text-sm font-normal text-gray-500">MWh</span>
                            </p>
                          </div>
                          <div className="rounded-lg bg-gray-950/50 px-4 py-3">
                            <p className="text-xs text-gray-500">Costo Bolsa</p>
                            <p className={`mt-1 text-xl font-bold tabular-nums ${ra.costo_bolsa_total_mcop <= 0 ? 'text-emerald-400' : 'text-red-400'}`}>
                              {formatMiles(ra.costo_bolsa_total_mcop, 2)} <span className="text-sm font-normal text-gray-500">M COP</span>
                            </p>
                          </div>
                        </div>
                      </div>

                      {/* Semáforo */}
                      <div className={`rounded-xl border px-5 py-4 ${sem.borderBg}`}>
                        <div className="flex items-center gap-3">
                          <span className={`text-2xl font-bold ${sem.color}`}>{sem.icon}</span>
                          <div className="flex-1">
                            <p className={`text-base font-bold ${sem.color}`}>{sem.titulo}</p>
                            <p className="text-sm text-gray-400">{sem.desc}</p>
                          </div>
                          <span className={`text-right text-sm font-semibold tabular-nums ${delta_costo_mcop < 0 ? 'text-emerald-400' : 'text-red-400'}`}>
                            {delta_costo_mcop < 0 ? '▼' : '▲'} {Math.abs(delta_costo_mcop).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} M COP
                          </span>
                        </div>
                      </div>

                      {/* PARTE D — KPI comparativas */}
                      <div className="grid grid-cols-2 gap-3">
                        <SimCompareCard
                          label="Posición Neta (MWh)"
                          antes={ra.posicion_neta_total_mwh}
                          despues={rd.posicion_neta_total_mwh}
                          formato={(v) => formatMiles(v, 0)}
                          mejoraSiMenor={simTipo === 'venta'}
                        />
                        <SimCompareCard
                          label="Costo/Ingreso Bolsa (M COP)"
                          antes={ra.costo_bolsa_total_mcop}
                          despues={rd.costo_bolsa_total_mcop}
                          formato={(v) => formatMiles(v, 2)}
                          mejoraSiMenor
                        />
                        <SimCompareCard
                          label="Hora pico compra"
                          antes={ra.hora_pico_compra}
                          despues={rd.hora_pico_compra}
                          formato={(v) => `H${v}`}
                          soloInfo
                        />
                        <SimCompareCard
                          label="Hora pico venta"
                          antes={ra.hora_pico_venta}
                          despues={rd.hora_pico_venta}
                          formato={(v) => `H${v}`}
                          soloInfo
                        />
                      </div>
                    </>
                  )
                })()}
              </div>
            </div>

            {/* ── Chart + Table (full width) ─────────────────────────────── */}
            {simResultado && (() => {
              const lineaColor = { verde: '#22c55e', amarillo: '#eab308', rojo: '#ef4444' }[simResultado.recomendacion]
              return (
                <>
                  <section className="rounded-xl border border-gray-800 bg-gray-900/60 p-6">
                    <h2 className="mb-1 text-lg font-semibold text-gray-200">
                      Posición Neta hora a hora — promedio del período
                    </h2>
                    <p className="mb-5 text-xs text-gray-500">
                      Posición actual (azul) vs. con el nuevo contrato (línea punteada)
                    </p>
                    <div className="h-72 w-full">
                      <ResponsiveContainer width="100%" height="100%">
                        <LineChart
                          data={simResultado.perfil_horario}
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
                            contentStyle={{ backgroundColor: '#111827', border: '1px solid #374151', borderRadius: '8px', color: '#f3f4f6' }}
                            labelFormatter={(h) => `Hora: ${h}`}
                            formatter={(v, n) => [`${formatMiles(Number(v), 2)} MWh`, n]}
                          />
                          <Legend wrapperStyle={{ color: '#9ca3af', paddingTop: 12 }} />
                          <Line type="monotone" dataKey="posicion_antes_mwh" name="Posición actual" stroke="#3b82f6" strokeWidth={2} dot={false} />
                          <Line type="monotone" dataKey="posicion_despues_mwh" name="Con nuevo contrato" stroke={lineaColor} strokeWidth={2.5} dot={false} strokeDasharray="6 3" />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>
                  </section>

                  <section className="overflow-hidden rounded-xl border border-gray-800">
                    <div className="border-b border-gray-800 bg-gray-900/80 px-6 py-4">
                      <h2 className="text-lg font-semibold text-gray-200">Resumen por mes</h2>
                    </div>
                    <div className="bg-gray-900/40 px-6 py-5">
                      <TablaAnalisis
                        encabezados={['Mes', 'Pos. Actual (MWh)', 'Pos. Nueva (MWh)', 'Diferencia (MWh)', 'Costo Actual (M COP)', 'Costo Nuevo (M COP)', 'Ahorro (M COP)']}
                        filas={
                          <>
                            {simResultado.por_mes.map((fila) => (
                              <tr key={fila.mes} className="transition-colors hover:bg-gray-800/40">
                                <td className="px-4 py-2.5 font-medium text-gray-200">{fila.mes}</td>
                                <td className="px-4 py-2.5 text-right tabular-nums text-gray-300">{formatMiles(fila.pos_actual_mwh)}</td>
                                <td className="px-4 py-2.5 text-right tabular-nums text-blue-300">{formatMiles(fila.pos_nueva_mwh)}</td>
                                <td className={`px-4 py-2.5 text-right font-medium tabular-nums ${fila.diferencia_mwh <= 0 ? 'text-emerald-400' : 'text-red-400'}`}>
                                  {fila.diferencia_mwh > 0 ? '+' : ''}{formatMiles(fila.diferencia_mwh)}
                                </td>
                                <td className="px-4 py-2.5 text-right tabular-nums text-gray-300">{formatMiles(fila.costo_actual_mcop)}</td>
                                <td className="px-4 py-2.5 text-right tabular-nums text-blue-300">{formatMiles(fila.costo_nuevo_mcop)}</td>
                                <td className={`px-4 py-2.5 text-right font-semibold tabular-nums ${fila.ahorro_mcop >= 0 ? 'text-emerald-400' : 'text-red-400'}`}>
                                  {fila.ahorro_mcop > 0 ? '+' : ''}{formatMiles(fila.ahorro_mcop)}
                                </td>
                              </tr>
                            ))}
                          </>
                        }
                      />
                    </div>
                  </section>
                </>
              )
            })()}
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

export default App
