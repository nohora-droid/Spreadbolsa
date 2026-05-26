import { useEffect, useMemo, useState, type ReactNode } from 'react'
import {
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

type VistaAnalisis = 'dia' | 'mes' | 'comparativo'

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

export default App
