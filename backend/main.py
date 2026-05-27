"""

API REST para el proyecto Spread Bolsa BIA.

"""



from __future__ import annotations



from datetime import date, datetime
from pathlib import Path



from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from typing import List, Optional
import pandas as pd

from dotenv import load_dotenv

from spread_engine import calcular_spread, cargar_pb_sql
from portfolio_engine import (
    cargar_inventario,
    calcular_posicion_neta,
    calcular_costo_bolsa,
    resumen_portafolio,
)
from simulation_engine import simular_contrato





# Carga variables de entorno si existe un .env (por si luego agregas configuraciones).

load_dotenv()



app = FastAPI()



# Habilita CORS para todos los orígenes (útil para desarrollo).

app.add_middleware(

    CORSMiddleware,

    allow_origins=["*"],

    allow_credentials=True,

    allow_methods=["*"],

    allow_headers=["*"],

)



# ID de la base de datos en Metabase para consultas SQL nativas (price_pb_hourly).

METABASE_DATABASE_PB = 2344





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

    # a) Cargar inventario desde la carpeta raíz del proyecto (un nivel arriba de backend/).
    ruta_raiz_proyecto = Path(__file__).parent.parent
    try:
        inventario = cargar_inventario(ruta_raiz_proyecto)
    except Exception as error:
        raise HTTPException(
            status_code=422,
            detail=f"No fue posible cargar inventario: {error}",
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
        # Posición neta horaria según inventario y tipo de día.
        df_posicion = calcular_posicion_neta(inventario, fecha)

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

class ContratoSimulacion(BaseModel):
    tipo: str = Field("compra", description="'compra' o 'venta'")
    contraparte: str = Field("", description="Nombre libre de la contraparte")
    precio_cop_kwh: float = Field(..., gt=0, description="Precio del contrato en COP/kWh")
    fecha_inicio: str = Field(..., description="Inicio de vigencia YYYY-MM-DD")
    fecha_fin: str = Field(..., description="Fin de vigencia YYYY-MM-DD")
    tipo_mercado: str = Field("regulado", description="'regulado', 'no_regulado' o 'ambos'")
    perfil_horario: str = Field("plano", description="'plano'|'custom'|'excel_custom'|'ordinario'|'sabado'|'festivo'")
    # Energía total en MWh/mes (requerida para plano, solar, bloques)
    energia_mensual_mwh: Optional[float] = Field(None, gt=0, description="Energía en MWh/mes")
    # Para perfil solar/bloques: 24 pesos relativos (se normalizan a suma=1)
    perfil_pesos_24h: Optional[List[float]] = Field(None, description="24 pesos horarios normalizados")
    # Para perfil Excel: 12 × 24 valores absolutos de kWh/mes por hora
    perfil_excel_12x24: Optional[List[List[float]]] = Field(None, description="Matriz 12×24 kWh/mes por hora")


@app.post("/simulate")
def simulate(contrato: ContratoSimulacion):
    """
    Simula el impacto de un nuevo contrato sobre el portafolio Olibia.

    Compara posición neta y costo de bolsa antes y después de incluir el contrato.
    Retorna: resumen_antes, resumen_despues, recomendacion (verde/amarillo/rojo),
    perfil_horario promedio (24h) y tabla por mes.
    """
    # Validar fechas
    _validar_fecha_iso(contrato.fecha_inicio, "fecha_inicio")
    _validar_fecha_iso(contrato.fecha_fin, "fecha_fin")
    if contrato.fecha_inicio > contrato.fecha_fin:
        raise HTTPException(
            status_code=400,
            detail="fecha_inicio no puede ser posterior a fecha_fin.",
        )

    # Validar enumeraciones
    if contrato.tipo not in ("compra", "venta"):
        raise HTTPException(status_code=400, detail="tipo debe ser 'compra' o 'venta'.")
    if contrato.tipo_mercado not in ("regulado", "no_regulado", "ambos"):
        raise HTTPException(
            status_code=400,
            detail="tipo_mercado debe ser 'regulado', 'no_regulado' o 'ambos'.",
        )
    if contrato.perfil_horario not in ("plano", "ordinario", "sabado", "festivo"):
        raise HTTPException(
            status_code=400,
            detail="perfil_horario debe ser 'plano', 'ordinario', 'sabado' o 'festivo'.",
        )

    # Cargar inventario
    ruta_raiz = Path(__file__).parent.parent
    try:
        inventario = cargar_inventario(ruta_raiz)
    except Exception as error:
        raise HTTPException(
            status_code=422,
            detail=f"No fue posible cargar inventario: {error}",
        ) from error

    # Cargar precios de bolsa para el período de vigencia
    try:
        df_pb = cargar_pb_sql(
            database_id=METABASE_DATABASE_PB,
            fecha_inicio=contrato.fecha_inicio,
            fecha_fin=contrato.fecha_fin,
        )
    except Exception as error:
        raise HTTPException(
            status_code=502,
            detail=f"Error consultando Metabase para PB: {error}",
        ) from error

    if df_pb.empty:
        raise HTTPException(
            status_code=404,
            detail="No hay datos de PB para el rango de fechas del contrato.",
        )

    # Ejecutar simulación
    try:
        resultado = simular_contrato(
            inventario=inventario,
            df_pb=df_pb,
            tipo=contrato.tipo,
            precio_cop_kwh=contrato.precio_cop_kwh,
            fecha_inicio=contrato.fecha_inicio,
            fecha_fin=contrato.fecha_fin,
            tipo_mercado=contrato.tipo_mercado,
            perfil_horario=contrato.perfil_horario,
            energia_mensual_mwh=contrato.energia_mensual_mwh,
            perfil_pesos_24h=contrato.perfil_pesos_24h,
            perfil_excel_12x24=contrato.perfil_excel_12x24,
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
