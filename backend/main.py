"""

API REST para el proyecto Spread Bolsa BIA.

"""



from __future__ import annotations



from datetime import date, datetime



from fastapi import FastAPI, HTTPException, Query

from fastapi.middleware.cors import CORSMiddleware



from dotenv import load_dotenv



from spread_engine import calcular_spread, cargar_pb_sql





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


