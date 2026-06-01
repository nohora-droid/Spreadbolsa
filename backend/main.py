"""

API REST para el proyecto Spread Bolsa BIA.

"""



from __future__ import annotations



from datetime import date, datetime, timedelta



import os

from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from typing import List, Optional
import pandas as pd

from dotenv import load_dotenv

import concurrent.futures

from spread_engine import calcular_spread, cargar_pb_sql
from olibia_loader import (
    cargar_posicion_olibia,
    cargar_posicion_con_demanda,
    get_contracts,
    get_contract_hourly,
    get_precios_contratos,
)
from portfolio_engine import (
    calcular_posicion_neta,
    calcular_costo_bolsa,
    resumen_portafolio,
)
from simulation_engine import simular_contrato





# Carga variables de entorno si existe un .env (por si luego agregas configuraciones).

load_dotenv()



app = FastAPI()



# ── CORS ──────────────────────────────────────────────────────────────────────
# En desarrollo acepta localhost:5173 (Vite dev server).
# En producción Railway lee FRONTEND_URL desde las variables de entorno de Railway
# y lo añade a la lista de orígenes permitidos.
# Si FRONTEND_URL no está definida se usa "*" como fallback (solo en desarrollo).

_frontend_url = os.getenv("FRONTEND_URL", "")

_allowed_origins: list[str] = [
    "http://localhost:5173",   # Vite dev server local
    "http://127.0.0.1:5173",  # alternativa localhost
]

if _frontend_url:
    # Agrega el dominio de Vercel configurado en Railway
    _allowed_origins.append(_frontend_url)
else:
    # Sin FRONTEND_URL definida → acepta cualquier origen (solo útil en desarrollo)
    _allowed_origins = ["*"]

app.add_middleware(
    CORSMiddleware,
    allow_origins=_allowed_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)



# ID de la base de datos en Metabase para consultas SQL nativas (price_pb_hourly).

METABASE_DATABASE_PB = 2344


# ── Raíz — evita 404 en "/" que hace que Render/Railway apague el servicio ────

@app.get("/")
def root():
    """Endpoint raíz para plataformas que sondean '/' como healthcheck (Render, etc.)."""
    return {"status": "ok", "app": "Spread Bolsa BIA", "version": "1.0"}


# ── Healthcheck explícito — Railway lo sondea para verificar que el contenedor está vivo ─

@app.get("/health")
def health():
    """Endpoint de salud para Railway healthcheck."""
    return {"status": "ok"}






def _fecha_inicio_por_defecto() -> str:

    """Primer día del mes actual en formato YYYY-MM-DD."""

    hoy = date.today()

    return date(hoy.year, hoy.month, 1).isoformat()





def _fecha_fin_por_defecto() -> str:

    """Fecha de hoy en formato YYYY-MM-DD."""

    return date.today().isoformat()





def _validar_fecha_iso(fecha: str, nombre_parametro: str) -> None:

    """Comprueba que la fecha tenga formato YYYY-MM-DD."""

    try:

        datetime.strptime(fecha, "%Y-%m-%d")

    except ValueError as error:

        raise HTTPException(

            status_code=400,

            detail=f"{nombre_parametro} debe tener formato YYYY-MM-DD (ej. 2024-01-15).",

        ) from error





@app.get("/health")

def health():

    return {"status": "ok", "proyecto": "Spread Bolsa BIA"}





@app.get("/spread")

def spread(

    precio_contrato: float = Query(

        350.0,

        description="Precio fijo del contrato en COP/kWh para calcular el spread.",

    ),

    fecha_inicio: str = Query(

        default_factory=_fecha_inicio_por_defecto,

        description="Fecha mínima inclusive (YYYY-MM-DD). Por defecto: primer día del mes actual.",

    ),

    fecha_fin: str = Query(

        default_factory=_fecha_fin_por_defecto,

        description="Fecha máxima inclusive (YYYY-MM-DD). Por defecto: hoy.",

    ),

):

    """

    Calcula el spread horario a partir de precios de bolsa en Metabase (SQL nativo).



    Flujo: cargar PB vía SQL (price_pb_hourly) → calcular spread → devolver resumen y datos completos.

    """

    # Validar formato de fechas antes de consultar Metabase

    _validar_fecha_iso(fecha_inicio, "fecha_inicio")

    _validar_fecha_iso(fecha_fin, "fecha_fin")

    if fecha_inicio > fecha_fin:

        raise HTTPException(

            status_code=400,

            detail="fecha_inicio no puede ser posterior a fecha_fin.",

        )



    try:

        # 1. Cargar precios de bolsa desde SQL (formato largo: fecha, hora, precio_bolsa)

        df_pb = cargar_pb_sql(

            database_id=METABASE_DATABASE_PB,

            fecha_inicio=fecha_inicio,

            fecha_fin=fecha_fin,

        )



        if df_pb.empty:

            raise HTTPException(

                status_code=404,

                detail="No hay datos de precio de bolsa para el rango de fechas indicado.",

            )



        # 2. Calcular spread y métricas de resumen

        df_spread, resumen = calcular_spread(df_pb, precio_contrato)



        total_filas = len(df_spread)



        # 3. Serializar todas las filas del rango para el frontend

        columnas_respuesta = ["fecha", "hora", "precio_bolsa", "spread"]

        df_respuesta = df_spread[columnas_respuesta]



        # Convertir a tipos nativos de Python para serialización JSON

        datos = [

            {

                "fecha": str(fila["fecha"]),

                "hora": int(fila["hora"]),

                "precio_bolsa": float(fila["precio_bolsa"]),

                "spread": float(fila["spread"]),

            }

            for fila in df_respuesta.to_dict(orient="records")

        ]



        return {

            "resumen": resumen,

            "datos": datos,

            "total_filas": total_filas,

        }



    except HTTPException:

        # Re-lanzar errores HTTP ya definidos (validación, sin datos, etc.)

        raise



    except ValueError as error:

        # Errores de estructura de datos o configuración (.env, columnas, etc.)

        raise HTTPException(

            status_code=422,

            detail=f"Error al procesar los datos: {error}",

        ) from error



    except Exception as error:

        # Fallos de red, Metabase, SQL u otros imprevistos

        raise HTTPException(

            status_code=502,

            detail=f"No se pudieron obtener los datos de precio de bolsa: {error}",

        ) from error


@app.get("/portfolio/posicion")
def portfolio_posicion(
    fecha_inicio: str = Query(
        ...,
        description="Fecha inicio del período (YYYY-MM-DD, inclusive).",
    ),
    fecha_fin: str = Query(
        ...,
        description="Fecha fin del período (YYYY-MM-DD, inclusive).",
    ),
):
    """
    Calcula la posición energética del portafolio Olibia incluyendo demanda.

    Fuentes de datos:
      - Contratos: API Olibia Energy (compra R, compra NR, venta).
      - Demanda de clientes: Metabase cards 9440 (regulada) y 9439 (no regulada),
        promediadas por hora y tipo de día, asignadas según el tipo_dia real
        de cada fecha (ordinario / sábado / domingo / festivo).

    Fórmula posición neta:
      posicion_neta_kwh = compra_r + compra_nr − venta − demanda_r − demanda_nr

    Retorna (todo en kWh, sin precios):
      datos              : filas horarias con {fecha, hora, tipo_dia,
                           compra_r_kwh, compra_nr_kwh, venta_kwh,
                           demanda_r_kwh, demanda_nr_kwh, posicion_neta_kwh}
      resumen_mensual    : agregados mensuales en MWh + número de días
      total_dias         : total de días procesados
      distribucion_tipo_dia : conteo de días por tipo
    """
    _validar_fecha_iso(fecha_inicio, "fecha_inicio")
    _validar_fecha_iso(fecha_fin, "fecha_fin")
    if fecha_inicio > fecha_fin:
        raise HTTPException(
            status_code=400,
            detail="fecha_inicio no puede ser posterior a fecha_fin.",
        )

    # a) Cargar posición de contratos + demanda de clientes desde Olibia y Metabase.
    try:
        df_raw = cargar_posicion_con_demanda(fecha_inicio, fecha_fin)
    except Exception as error:
        raise HTTPException(
            status_code=500,
            detail=f"Error calculando posición del portafolio: {error}",
        ) from error

    if df_raw.empty:
        raise HTTPException(
            status_code=404,
            detail=(
                f"No hay datos de posición para el rango "
                f"{fecha_inicio} a {fecha_fin}."
            ),
        )

    # b) Renombrar columnas al esquema del endpoint (date → fecha, hour → hora).
    df = df_raw.rename(columns={"date": "fecha", "hour": "hora"})

    # c) Serializar filas horarias completas.
    datos = [
        {
            "fecha":           str(fila["fecha"]),
            "hora":            int(fila["hora"]),
            "tipo_dia":        str(fila["tipo_dia"]),
            "compra_r_kwh":    float(fila["compra_r_kwh"]),
            "compra_nr_kwh":   float(fila["compra_nr_kwh"]),
            "venta_kwh":       float(fila["venta_kwh"]),
            "demanda_r_kwh":   float(fila["demanda_r_kwh"]),
            "demanda_nr_kwh":  float(fila["demanda_nr_kwh"]),
            "posicion_neta_kwh": float(fila["posicion_neta_kwh"]),
        }
        for fila in df.to_dict(orient="records")
    ]

    # d) Resumen mensual: sumar kWh por mes y contar días únicos.
    df["mes"] = df["fecha"].str[:7]

    columnas_agg = [
        "compra_r_kwh", "compra_nr_kwh", "venta_kwh",
        "demanda_r_kwh", "demanda_nr_kwh", "posicion_neta_kwh",
    ]
    agg_kwh = (
        df.groupby("mes")[columnas_agg]
        .sum()
        .reset_index()
    )
    dias_mes = (
        df.groupby("mes")["fecha"]
        .nunique()
        .rename("dias")
        .reset_index()
    )
    resumen_df = agg_kwh.merge(dias_mes, on="mes")

    resumen_mensual = [
        {
            "mes":               str(fila["mes"]),
            "compra_r_mwh":      round(float(fila["compra_r_kwh"])       / 1_000, 3),
            "compra_nr_mwh":     round(float(fila["compra_nr_kwh"])      / 1_000, 3),
            "venta_mwh":         round(float(fila["venta_kwh"])          / 1_000, 3),
            "demanda_r_mwh":     round(float(fila["demanda_r_kwh"])      / 1_000, 3),
            "demanda_nr_mwh":    round(float(fila["demanda_nr_kwh"])     / 1_000, 3),
            "posicion_neta_mwh": round(float(fila["posicion_neta_kwh"])  / 1_000, 3),
            "dias":              int(fila["dias"]),
        }
        for fila in resumen_df.to_dict(orient="records")
    ]

    # e) Distribución de tipos de día (un conteo por fecha única).
    conteo = (
        df.drop_duplicates("fecha")["tipo_dia"]
        .value_counts()
        .to_dict()
    )
    distribucion_tipo_dia = {
        "ordinarios": int(conteo.get("ordinario", 0)),
        "sabados":    int(conteo.get("sabado",    0)),
        "domingos":   int(conteo.get("domingo",   0)),
        "festivos":   int(conteo.get("festivo",   0)),
    }

    return {
        "datos":                 datos,
        "resumen_mensual":       resumen_mensual,
        "total_dias":            int(df["fecha"].nunique()),
        "distribucion_tipo_dia": distribucion_tipo_dia,
    }


@app.get("/portfolio")
def portfolio(
    fecha_inicio: str = Query(
        default_factory=_fecha_inicio_por_defecto,
        description="Fecha mínima inclusive (YYYY-MM-DD). Por defecto: primer día del mes actual.",
    ),
    fecha_fin: str = Query(
        default_factory=_fecha_fin_por_defecto,
        description="Fecha máxima inclusive (YYYY-MM-DD). Por defecto: hoy.",
    ),
):
    """
    Calcula portafolio horario combinando inventario de contratos y precio de bolsa.
    """
    # Validar formato y coherencia de fechas antes de cualquier proceso.
    _validar_fecha_iso(fecha_inicio, "fecha_inicio")
    _validar_fecha_iso(fecha_fin, "fecha_fin")
    if fecha_inicio > fecha_fin:
        raise HTTPException(
            status_code=400,
            detail="fecha_inicio no puede ser posterior a fecha_fin.",
        )

    # a) Cargar posición del portafolio desde la API de Olibia (paralelo).
    try:
        df_inventario = cargar_posicion_olibia(fecha_inicio, fecha_fin)
    except Exception as error:
        raise HTTPException(
            status_code=502,
            detail=f"Error consultando API Olibia para posición del portafolio: {error}",
        ) from error

    # b) Cargar PB desde Metabase usando SQL nativo.
    try:
        df_pb = cargar_pb_sql(
            database_id=METABASE_DATABASE_PB,
            fecha_inicio=fecha_inicio,
            fecha_fin=fecha_fin,
        )
    except Exception as error:
        raise HTTPException(
            status_code=502,
            detail=f"Error consultando Metabase para PB: {error}",
        ) from error

    # 404 si no existen datos de PB en el rango consultado.
    if df_pb.empty:
        raise HTTPException(
            status_code=404,
            detail="No hay datos de PB para el rango de fechas indicado.",
        )

    # c) Calcular resultados por cada fecha única de PB.
    resultados_por_dia = []
    fechas_unicas = sorted(df_pb["fecha"].dropna().unique().tolist())

    for fecha in fechas_unicas:
        # Posición neta horaria del día (filtra el DataFrame de la API Olibia).
        df_posicion = calcular_posicion_neta(df_inventario, fecha)

        # Subconjunto PB del día para cruzar por hora.
        df_pb_dia = df_pb.loc[df_pb["fecha"] == fecha].copy()

        # Costo de bolsa por hora (posición neta * precio bolsa).
        df_costo_dia = calcular_costo_bolsa(df_posicion, df_pb_dia)
        df_costo_dia["fecha"] = fecha
        resultados_por_dia.append(df_costo_dia)

    # d) Concatenar resultados diarios en una sola tabla.
    if not resultados_por_dia:
        raise HTTPException(
            status_code=404,
            detail="No se pudieron construir resultados de portafolio para el rango indicado.",
        )
    df_portafolio = pd.concat(resultados_por_dia, ignore_index=True).sort_values(
        ["fecha", "hora"]
    ).reset_index(drop=True)

    # e) Calcular resumen agregado del portafolio.
    resumen = resumen_portafolio(df_portafolio, f"{fecha_inicio} a {fecha_fin}")

    # Construir salida con columnas solicitadas en formato serializable JSON.
    columnas_respuesta = [
        "fecha",
        "hora",
        "compra_r_kwh",
        "compra_nr_kwh",
        "venta_kwh",
        "posicion_neta_kwh",
        "precio_bolsa",
        "costo_bolsa_cop",
    ]
    df_respuesta = df_portafolio[columnas_respuesta]

    datos = [
        {
            "fecha": str(fila["fecha"]),
            "hora": int(fila["hora"]),
            "compra_r_kwh": float(fila["compra_r_kwh"]),
            "compra_nr_kwh": float(fila["compra_nr_kwh"]),
            "venta_kwh": float(fila["venta_kwh"]),
            "posicion_neta_kwh": float(fila["posicion_neta_kwh"]),
            "precio_bolsa": float(fila["precio_bolsa"]),
            "costo_bolsa_cop": float(fila["costo_bolsa_cop"]),
        }
        for fila in df_respuesta.to_dict(orient="records")
    ]

    return {
        "resumen": resumen,
        "datos": datos,
        "total_filas": len(df_respuesta),
    }


# ─────────────────────────────────────────────────────────────────────────────
# Simulador de nuevos contratos
# ─────────────────────────────────────────────────────────────────────────────

_PERFILES_VALIDOS = frozenset({
    "plano", "bloques", "solar", "excel",
    # aliases legacy por compatibilidad
    "custom", "excel_custom",
    "ordinario", "sabado", "festivo",
})


class BloqueHorario(BaseModel):
    hora_ini: int = Field(..., ge=1, le=24, description="Hora de inicio (1-24)")
    hora_fin: int = Field(..., ge=1, le=24, description="Hora de fin inclusiva (1-24)")
    mwh_mes: float = Field(..., gt=0, description="Energía del bloque en MWh/mes")


class ContratoSimulacion(BaseModel):
    tipo: str = Field("compra", description="'compra' o 'venta'")
    contraparte: str = Field("", description="Nombre libre de la contraparte")
    precio_cop_kwh: float = Field(..., gt=0, description="Precio del contrato en COP/kWh")
    # ── Período del escenario de precio de bolsa ──────────────────────────────
    pb_desde: str = Field(..., description="Inicio del período PB histórico YYYY-MM-DD")
    pb_hasta: str = Field(..., description="Fin del período PB histórico YYYY-MM-DD")
    # ── Vigencia del contrato simulado ────────────────────────────────────────
    contrato_inicio: str = Field(..., description="Inicio de vigencia del contrato YYYY-MM-DD")
    contrato_fin: str = Field(..., description="Fin de vigencia del contrato YYYY-MM-DD")
    tipo_mercado: str = Field("regulado", description="'regulado', 'no_regulado' o 'ambos'")
    perfil_horario: str = Field(
        "plano",
        description="'plano' | 'bloques' | 'solar' | 'excel'",
    )
    # Energía total en kWh/mes (requerida para plano y solar)
    energia_mensual_kwh: Optional[float] = Field(None, gt=0, description="Energía en kWh/mes")
    # Para perfil 'bloques': lista de bloques horarios
    bloques: Optional[List[BloqueHorario]] = Field(None, description="Bloques horarios para perfil 'bloques'")
    # Para perfil 'custom' (legacy): 24 pesos relativos
    perfil_pesos_24h: Optional[List[float]] = Field(None, description="24 pesos horarios normalizados")
    # Para perfil 'excel' / 'excel_custom' (legacy): 12 × 24 valores absolutos de kWh/mes por hora
    perfil_excel_12x24: Optional[List[List[float]]] = Field(None, description="Matriz 12×24 kWh/mes por hora")


@app.post("/simulate")
def simulate(contrato: ContratoSimulacion):
    """
    Simula el impacto de un nuevo contrato sobre el portafolio Olibia.

    Compara posición neta y costo de bolsa antes y después de incluir el contrato.
    Retorna: resumen_antes, resumen_despues, recomendacion (verde/amarillo/rojo),
    perfil_horario promedio (24h) y tabla por mes.
    """
    # Validar los cuatro rangos de fechas
    _validar_fecha_iso(contrato.pb_desde, "pb_desde")
    _validar_fecha_iso(contrato.pb_hasta, "pb_hasta")
    _validar_fecha_iso(contrato.contrato_inicio, "contrato_inicio")
    _validar_fecha_iso(contrato.contrato_fin, "contrato_fin")
    if contrato.pb_desde > contrato.pb_hasta:
        raise HTTPException(
            status_code=400,
            detail="pb_desde no puede ser posterior a pb_hasta.",
        )
    if contrato.contrato_inicio > contrato.contrato_fin:
        raise HTTPException(
            status_code=400,
            detail="contrato_inicio no puede ser posterior a contrato_fin.",
        )

    # Validar enumeraciones
    if contrato.tipo not in ("compra", "venta"):
        raise HTTPException(status_code=400, detail="tipo debe ser 'compra' o 'venta'.")
    if contrato.tipo_mercado not in ("regulado", "no_regulado", "ambos"):
        raise HTTPException(
            status_code=400,
            detail="tipo_mercado debe ser 'regulado', 'no_regulado' o 'ambos'.",
        )
    if contrato.perfil_horario not in _PERFILES_VALIDOS:
        raise HTTPException(
            status_code=400,
            detail=(
                f"perfil_horario '{contrato.perfil_horario}' no reconocido. "
                "Valores aceptados: 'plano', 'bloques', 'solar', 'excel'."
            ),
        )

    # Cargar posición del portafolio desde la API Olibia para el período del contrato
    try:
        inventario = cargar_posicion_olibia(
            contrato.contrato_inicio, contrato.contrato_fin
        )
    except Exception as error:
        raise HTTPException(
            status_code=502,
            detail=f"Error consultando API Olibia para posición del portafolio: {error}",
        ) from error

    # Cargar precios de bolsa para el período del ESCENARIO (pb_desde → pb_hasta)
    try:
        df_pb = cargar_pb_sql(
            database_id=METABASE_DATABASE_PB,
            fecha_inicio=contrato.pb_desde,
            fecha_fin=contrato.pb_hasta,
        )
    except Exception as error:
        raise HTTPException(
            status_code=502,
            detail=f"Error consultando Metabase para PB: {error}",
        ) from error

    if df_pb.empty:
        raise HTTPException(
            status_code=404,
            detail=(
                f"No hay datos de PB para el período del escenario "
                f"({contrato.pb_desde} a {contrato.pb_hasta})."
            ),
        )

    # Obtener PPP real de contratos PC para el período del contrato
    # (no bloquea si falla — se devuelve None y el frontend usa mock)
    ppp_resumen = None
    try:
        ppp_resumen = get_precios_contratos(
            contrato.contrato_inicio, contrato.contrato_fin
        )
    except Exception as ppp_error:
        print(
            f"[/simulate] No se pudo obtener PPP de contratos: {ppp_error}. "
            "Se omite del resultado."
        )

    # Ejecutar simulación
    try:
        # Convertir bloques Pydantic → dicts simples para el engine
        bloques_dict = (
            [b.model_dump() for b in contrato.bloques]
            if contrato.bloques
            else None
        )
        resultado = simular_contrato(
            inventario=inventario,
            df_pb=df_pb,
            tipo=contrato.tipo,
            precio_cop_kwh=contrato.precio_cop_kwh,
            pb_desde=contrato.pb_desde,
            pb_hasta=contrato.pb_hasta,
            contrato_inicio=contrato.contrato_inicio,
            contrato_fin=contrato.contrato_fin,
            tipo_mercado=contrato.tipo_mercado,
            perfil_horario=contrato.perfil_horario,
            energia_mensual_kwh=contrato.energia_mensual_kwh,
            bloques=bloques_dict,
            perfil_pesos_24h=contrato.perfil_pesos_24h,
            perfil_excel_12x24=contrato.perfil_excel_12x24,
            ppp_resumen=ppp_resumen,
        )
    except ValueError as error:
        raise HTTPException(
            status_code=422,
            detail=f"Error en simulación: {error}",
        ) from error
    except Exception as error:
        raise HTTPException(
            status_code=500,
            detail=f"Error interno en simulación: {error}",
        ) from error

    return resultado


# ─────────────────────────────────────────────────────────────────────────────
# Listado de contratos activos desde Olibia Energy
# ─────────────────────────────────────────────────────────────────────────────

@app.get("/contratos")
def listar_contratos(
    start_date: str = Query(
        ...,
        description="Fecha inicio del período (YYYY-MM-DD, inclusive).",
    ),
    end_date: str = Query(
        ...,
        description="Fecha fin del período (YYYY-MM-DD, inclusive).",
    ),
):
    """
    Retorna la lista de contratos validados de Olibia Energy activos en el
    período indicado, con métricas de energía y precio promedio ponderado.

    Para cada contrato de modalidad PC (Precio Constante) se calcula el PPP:
        PPP = Σ(qty_abs × fixed_price) / Σ(qty_abs)

    Para contratos PLD (precio = bolsa) el campo precio_promedio_cop_kwh
    es null (el precio depende de la bolsa en cada hora).
    """
    _validar_fecha_iso(start_date, "start_date")
    _validar_fecha_iso(end_date, "end_date")
    if start_date > end_date:
        raise HTTPException(
            status_code=400,
            detail="start_date no puede ser posterior a end_date.",
        )

    # a) Obtener lista de contratos validados desde Olibia
    try:
        todos = get_contracts()
    except Exception as error:
        raise HTTPException(
            status_code=502,
            detail=f"Error consultando contratos en Olibia: {error}",
        ) from error

    # Filtrar solo los activos en el período solicitado
    contratos_activos = [
        c for c in todos
        if c["start_date"] <= end_date and c["end_date"] >= start_date
    ]

    # b) Cargar datos horarios de cada contrato en paralelo para calcular métricas
    def _cargar_metricas(contrato: dict) -> dict:
        """
        Carga datos horarios de un contrato y calcula:
          - energia_total_kwh: suma de |quantity| en el período
          - precio_promedio_cop_kwh: PPP para contratos PC; None para PLD
        """
        try:
            df = get_contract_hourly(contrato["id"], start_date, end_date)

            energia_kwh = float(df["quantity"].abs().sum()) if not df.empty else 0.0

            precio_ppp = None
            if contrato.get("modalidad") == "PC" and not df.empty:
                df_pc = df[
                    df["fixed_price"].notna() & (df["quantity"].abs() > 0)
                ].copy()
                if not df_pc.empty:
                    total_qty = df_pc["quantity"].abs().sum()
                    total_cop = (df_pc["quantity"].abs() * df_pc["fixed_price"]).sum()
                    precio_ppp = round(float(total_cop / total_qty), 4) if total_qty > 0 else None

            return {
                "id":                      contrato["id"],
                "nombre":                  contrato["contract_number"],
                "operacion":               contrato["operation"],
                "mercado":                 contrato["market_type"],
                "modalidad":               contrato.get("modalidad", ""),
                "energia_total_kwh":       round(energia_kwh, 2),
                "precio_promedio_cop_kwh": precio_ppp,
                "vigencia":                f"{contrato['start_date']}..{contrato['end_date']}",
            }

        except Exception as exc:
            # Devolver el contrato con error anotado para no bloquear la respuesta
            print(f"[/contratos] Error cargando {contrato['contract_number']}: {exc}")
            return {
                "id":                      contrato["id"],
                "nombre":                  contrato["contract_number"],
                "operacion":               contrato["operation"],
                "mercado":                 contrato["market_type"],
                "modalidad":               contrato.get("modalidad", ""),
                "energia_total_kwh":       None,
                "precio_promedio_cop_kwh": None,
                "vigencia":                f"{contrato['start_date']}..{contrato['end_date']}",
                "error":                   str(exc),
            }

    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as pool:
        futuros = {pool.submit(_cargar_metricas, c): c for c in contratos_activos}
        resultados = [futuro.result() for futuro in concurrent.futures.as_completed(futuros)]

    # Ordenar por operación, mercado y nombre para presentación consistente
    resultados.sort(key=lambda x: (x["operacion"], x["mercado"], x["nombre"]))

    return {
        "contratos": resultados,
        "total":     len(resultados),
        "periodo":   f"{start_date} a {end_date}",
    }


# ─────────────────────────────────────────────────────────────────────────────
# PPP (Precio Promedio Ponderado) de contratos PC por categoría
# ─────────────────────────────────────────────────────────────────────────────

@app.get("/contratos/ppp")
def contratos_ppp(
    start_date: str = Query(
        ...,
        description="Fecha inicio del período (YYYY-MM-DD, inclusive).",
    ),
    end_date: str = Query(
        ...,
        description="Fecha fin del período (YYYY-MM-DD, inclusive).",
    ),
):
    """
    Retorna el PPP (Precio Promedio Ponderado) de contratos PC de Olibia
    por categoría de operación: compra regulada, compra no regulada y venta.

    Solo considera contratos con modalidad PC (Precio Constante); los contratos
    PLD se excluyen porque su precio efectivo es el de la bolsa.

    Fórmula PPP:  Σ(|energía| × fixed_price) / Σ(|energía|)

    Tipo del PPP:
      "Indexado"   → todos los datos son históricos confirmados (is_projected_data=False)
      "Proyectado" → algún dato es proyección futura (is_projected_data=True)
      "Sin datos"  → no hay contratos PC activos en esa categoría en el período

    Retorna:
        {
          "compra_r":      {"ppp": float | null, "tipo": str},
          "compra_nr":     {"ppp": float | null, "tipo": str},
          "venta":         {"ppp": float | null, "tipo": str},
          "pld_excluidos": int,
          "contratos_pc":  int,
        }
    """
    _validar_fecha_iso(start_date, "start_date")
    _validar_fecha_iso(end_date,   "end_date")
    if start_date > end_date:
        raise HTTPException(
            status_code=400,
            detail="start_date no puede ser posterior a end_date.",
        )

    try:
        return get_precios_contratos(start_date, end_date)
    except Exception as error:
        raise HTTPException(
            status_code=502,
            detail=f"Error consultando precios de contratos en Olibia: {error}",
        ) from error
