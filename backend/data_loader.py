"""
Funciones para cargar datos desde archivos CSV, Excel y Metabase.
"""

from __future__ import annotations

import json
import os
import urllib.error
import urllib.request
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv


COLUMNAS_ESPERADAS = ["fecha", "hora", "precio_bolsa", "energia_kwh"]


def _validar_columnas(df: pd.DataFrame, ruta: str) -> None:
    """Valida que el DataFrame tenga las columnas esperadas."""
    faltantes = [c for c in COLUMNAS_ESPERADAS if c not in df.columns]
    if faltantes:
        raise ValueError(
            "El archivo no tiene las columnas esperadas. "
            f"Faltantes: {faltantes}. Ruta: {ruta}"
        )


def load_csv(ruta: str) -> pd.DataFrame:
    """
    Carga un CSV y devuelve un DataFrame con columnas:
    fecha, hora, precio_bolsa, energia_kwh
    """
    df = pd.read_csv(ruta)
    _validar_columnas(df, ruta)
    print(f"Se cargaron {len(df)} filas desde CSV: {ruta}")
    return df


def load_excel(ruta: str) -> pd.DataFrame:
    """
    Carga un Excel y devuelve un DataFrame con columnas:
    fecha, hora, precio_bolsa, energia_kwh
    """
    df = pd.read_excel(ruta)
    _validar_columnas(df, ruta)
    print(f"Se cargaron {len(df)} filas desde Excel: {ruta}")
    return df


def load_metabase_card(card_id: int | str, limit: int = 10000) -> pd.DataFrame:
    """
    Consulta una tarjeta de Metabase por API y devuelve un DataFrame.

    Lee METABASE_URL y METABASE_API_KEY desde el archivo .env del backend.

    Args:
        card_id: Identificador de la tarjeta en Metabase.
        limit: Máximo de filas a solicitar (default 10000).
    """
    # Ruta al .env en la misma carpeta que este módulo
    env_path = Path(__file__).resolve().parent / ".env"
    load_dotenv(env_path)

    metabase_url = os.getenv("METABASE_URL")
    metabase_api_key = os.getenv("METABASE_API_KEY")

    if not metabase_url or not metabase_api_key:
        mensaje = (
            "Faltan variables de entorno. Verifique METABASE_URL y "
            "METABASE_API_KEY en el archivo .env"
        )
        print(f"Error al cargar Metabase: {mensaje}")
        raise ValueError(mensaje)

    # URL del endpoint de consulta de la tarjeta (sin /json)
    url = f"{metabase_url.rstrip('/')}/api/card/{card_id}/query"
    headers = {
        "x-api-key": metabase_api_key,
        "Content-Type": "application/json",
    }
    # Cuerpo con parámetros vacíos y límite de filas en la consulta
    cuerpo_post = json.dumps(
        {
            "parameters": [],
            "constraints": {"max-results": limit},
        }
    ).encode("utf-8")

    try:
        # Petición POST con autenticación por API key
        solicitud = urllib.request.Request(
            url, data=cuerpo_post, headers=headers, method="POST"
        )
        with urllib.request.urlopen(solicitud, timeout=60) as respuesta:
            cuerpo = respuesta.read().decode("utf-8")
            datos = json.loads(cuerpo)

        # La respuesta trae filas y columnas dentro de la clave "data"
        bloque_datos = datos.get("data")
        if bloque_datos is None:
            mensaje = (
                "La respuesta de Metabase no contiene la clave 'data'. "
                f"Claves recibidas: {list(datos.keys())}"
            )
            print(f"Error al cargar Metabase (card_id={card_id}): {mensaje}")
            raise ValueError(mensaje)

        filas = bloque_datos.get("rows")
        columnas_meta = bloque_datos.get("cols")
        if filas is None or columnas_meta is None:
            mensaje = (
                "La clave 'data' no incluye 'rows' y/o 'cols'. "
                f"Claves en data: {list(bloque_datos.keys())}"
            )
            print(f"Error al cargar Metabase (card_id={card_id}): {mensaje}")
            raise ValueError(mensaje)

        # Nombres de columnas desde el campo "name" de cada col
        nombres_columnas = [col.get("name", f"col_{i}") for i, col in enumerate(columnas_meta)]
        df = pd.DataFrame(filas, columns=nombres_columnas)

        print(
            f"Se cargaron {len(df)} filas desde Metabase "
            f"(card_id={card_id}, limit={limit})"
        )
        return df

    except urllib.error.HTTPError as error:
        # Error devuelto por el servidor (4xx, 5xx)
        print(
            f"Error HTTP al consultar Metabase (card_id={card_id}): "
            f"{error.code} {error.reason}"
        )
        if error.fp:
            detalle = error.read().decode("utf-8", errors="replace")
            print(f"Detalle del servidor: {detalle}")
        raise

    except urllib.error.URLError as error:
        # Problemas de red o URL inválida
        print(f"Error de conexión con Metabase (card_id={card_id}): {error.reason}")
        raise

    except json.JSONDecodeError as error:
        # Respuesta que no es JSON válido
        print(f"Error al interpretar la respuesta JSON de Metabase: {error}")
        raise

    except ValueError:
        # Re-lanzar errores de estructura de respuesta ya impresos arriba
        raise

    except Exception as error:
        print(f"Error inesperado al cargar Metabase (card_id={card_id}): {error}")
        raise


def load_metabase_card_rango(
    card_id: int | str, fecha_inicio: str, fecha_fin: str
) -> pd.DataFrame:
    """
    Consulta una tarjeta de Metabase por API filtrando por rango de fechas y
    devuelve un DataFrame.

    Lee METABASE_URL y METABASE_API_KEY desde el archivo .env del backend.

    Args:
        card_id: Identificador de la tarjeta en Metabase.
        fecha_inicio: Fecha inicial (inclusive) en formato YYYY-MM-DD.
        fecha_fin: Fecha final (inclusive) en formato YYYY-MM-DD.
    """
    # Ruta al .env en la misma carpeta que este módulo
    env_path = Path(__file__).resolve().parent / ".env"
    load_dotenv(env_path)

    metabase_url = os.getenv("METABASE_URL")
    metabase_api_key = os.getenv("METABASE_API_KEY")

    if not metabase_url or not metabase_api_key:
        mensaje = (
            "Faltan variables de entorno. Verifique METABASE_URL y "
            "METABASE_API_KEY en el archivo .env"
        )
        print(f"Error al cargar Metabase: {mensaje}")
        raise ValueError(mensaje)

    # URL del endpoint de consulta de la tarjeta
    url = f"{metabase_url.rstrip('/')}/api/card/{card_id}/query"
    headers = {
        "x-api-key": metabase_api_key,
        "Content-Type": "application/json",
    }

    cuerpo_post = json.dumps(
        {
            "parameters": [
                {
                    "id": "823071f1-08c6-fd0e-d0e6-a03d37867f16",
                    "type": "string/=",
                    "target": ["dimension", ["template-tag", "version"]],
                    "value": "TxF",
                },
                {
                    "id": "b08dd8c9-cd7b-df8a-76d6-a19c9d0a1248",
                    "type": "date/range",
                    "target": ["dimension", ["template-tag", "date"]],
                    "value": f"{fecha_inicio}~{fecha_fin}",
                },
            ]
        }
    ).encode("utf-8")

    try:
        solicitud = urllib.request.Request(
            url, data=cuerpo_post, headers=headers, method="POST"
        )
        with urllib.request.urlopen(solicitud, timeout=60) as respuesta:
            cuerpo = respuesta.read().decode("utf-8")
            datos = json.loads(cuerpo)

        # La respuesta trae filas y columnas dentro de la clave "data"
        bloque_datos = datos.get("data")
        if bloque_datos is None:
            mensaje = (
                "La respuesta de Metabase no contiene la clave 'data'. "
                f"Claves recibidas: {list(datos.keys())}"
            )
            print(f"Error al cargar Metabase (card_id={card_id}): {mensaje}")
            raise ValueError(mensaje)

        filas = bloque_datos.get("rows")
        columnas_meta = bloque_datos.get("cols")
        if filas is None or columnas_meta is None:
            mensaje = (
                "La clave 'data' no incluye 'rows' y/o 'cols'. "
                f"Claves en data: {list(bloque_datos.keys())}"
            )
            print(f"Error al cargar Metabase (card_id={card_id}): {mensaje}")
            raise ValueError(mensaje)

        # Nombres de columnas desde el campo "name" de cada col
        nombres_columnas = [
            col.get("name", f"col_{i}") for i, col in enumerate(columnas_meta)
        ]
        df = pd.DataFrame(filas, columns=nombres_columnas)

        print(
            f"Se cargaron {len(df)} filas desde Metabase "
            f"(card_id={card_id}, rango={fecha_inicio} a {fecha_fin})"
        )
        return df

    except urllib.error.HTTPError as error:
        print(
            f"Error HTTP al consultar Metabase (card_id={card_id}): "
            f"{error.code} {error.reason}"
        )
        if error.fp:
            detalle = error.read().decode("utf-8", errors="replace")
            print(f"Detalle del servidor: {detalle}")
        raise

    except urllib.error.URLError as error:
        print(f"Error de conexión con Metabase (card_id={card_id}): {error.reason}")
        raise

    except json.JSONDecodeError as error:
        print(f"Error al interpretar la respuesta JSON de Metabase: {error}")
        raise

    except ValueError:
        raise

    except Exception as error:
        print(f"Error inesperado al cargar Metabase (card_id={card_id}): {error}")
        raise


def load_metabase_sql(
    sql: str, database_id: int | str, limit: int = 1_000_000
) -> pd.DataFrame:
    """
    Ejecuta una consulta SQL nativa en Metabase y devuelve un DataFrame.

    Lee METABASE_URL y METABASE_API_KEY desde el archivo .env del backend.

    Args:
        sql: Consulta SQL a ejecutar en la base de datos indicada.
        database_id: Identificador de la base de datos en Metabase.
        limit: Máximo de filas a solicitar a Metabase (default 1_000_000).
    """
    # Ruta al .env en la misma carpeta que este módulo
    env_path = Path(__file__).resolve().parent / ".env"
    load_dotenv(env_path)

    metabase_url = os.getenv("METABASE_URL")
    metabase_api_key = os.getenv("METABASE_API_KEY")

    if not metabase_url or not metabase_api_key:
        mensaje = (
            "Faltan variables de entorno. Verifique METABASE_URL y "
            "METABASE_API_KEY en el archivo .env"
        )
        print(f"Error al cargar Metabase SQL: {mensaje}")
        raise ValueError(mensaje)

    # Endpoint de consultas nativas (dataset)
    url = f"{metabase_url.rstrip('/')}/api/dataset"
    headers = {
        "x-api-key": metabase_api_key,
        "Content-Type": "application/json",
    }
    # Cuerpo con la consulta SQL y el id de base de datos
    cuerpo_post = json.dumps(
        {
            "database": database_id,
            "native": {"query": sql},
            "type": "native",
            "constraints": {"max-results": limit},
        }
    ).encode("utf-8")

    try:
        # Petición POST con autenticación por API key
        solicitud = urllib.request.Request(
            url, data=cuerpo_post, headers=headers, method="POST"
        )
        # timeout=60 s: evita que urllib quede colgado en conexiones lentas/caídas.
        with urllib.request.urlopen(solicitud, timeout=60) as respuesta:
            cuerpo = respuesta.read().decode("utf-8")
            datos = json.loads(cuerpo)

        # La respuesta trae filas y columnas dentro de la clave "data"
        bloque_datos = datos.get("data")
        if bloque_datos is None:
            mensaje = (
                "La respuesta de Metabase no contiene la clave 'data'. "
                f"Claves recibidas: {list(datos.keys())}"
            )
            print(
                f"Error al cargar Metabase SQL (database_id={database_id}): {mensaje}"
            )
            raise ValueError(mensaje)

        filas = bloque_datos.get("rows")
        columnas_meta = bloque_datos.get("cols")
        if filas is None or columnas_meta is None:
            mensaje = (
                "La clave 'data' no incluye 'rows' y/o 'cols'. "
                f"Claves en data: {list(bloque_datos.keys())}"
            )
            print(
                f"Error al cargar Metabase SQL (database_id={database_id}): {mensaje}"
            )
            raise ValueError(mensaje)

        # Nombres de columnas desde el campo "name" de cada col
        nombres_columnas = [
            col.get("name", f"col_{i}") for i, col in enumerate(columnas_meta)
        ]
        df = pd.DataFrame(filas, columns=nombres_columnas)

        print(
            f"Se cargaron {len(df)} filas desde Metabase SQL "
            f"(database_id={database_id}, limit={limit})"
        )
        return df

    except urllib.error.HTTPError as error:
        # Error devuelto por el servidor (4xx, 5xx)
        print(
            f"Error HTTP al consultar Metabase SQL (database_id={database_id}): "
            f"{error.code} {error.reason}"
        )
        if error.fp:
            detalle = error.read().decode("utf-8", errors="replace")
            print(f"Detalle del servidor: {detalle}")
        raise

    except urllib.error.URLError as error:
        # Problemas de red o URL inválida
        print(
            f"Error de conexión con Metabase SQL (database_id={database_id}): "
            f"{error.reason}"
        )
        raise

    except json.JSONDecodeError as error:
        # Respuesta que no es JSON válido
        print(f"Error al interpretar la respuesta JSON de Metabase SQL: {error}")
        raise

    except ValueError:
        # Re-lanzar errores de estructura de respuesta ya impresos arriba
        raise

    except Exception as error:
        print(
            f"Error inesperado al cargar Metabase SQL (database_id={database_id}): "
            f"{error}"
        )
        raise
