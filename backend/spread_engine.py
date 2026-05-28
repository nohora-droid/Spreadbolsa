"""
Motor de transformación y cálculo de spread para precios de bolsa (PB).
"""

from __future__ import annotations

from datetime import datetime
import time

import pandas as pd

from data_loader import load_metabase_card_rango, load_metabase_sql

# Columnas horarias esperadas en el formato ancho de Metabase
COLUMNAS_HORARIAS = [f"H{i}" for i in range(1, 25)]

# Columnas esperadas tras cargar PB desde SQL nativo
COLUMNAS_PB_SQL = {"file_date", "hour", "pb"}


def _generar_rangos_mensuales(fecha_inicio: str, fecha_fin: str) -> list[tuple[str, str]]:
    """
    Divide un rango de fechas en sub-rangos por mes calendario.

    El primer y último chunk respetan fecha_inicio y fecha_fin del rango total.
    Ejemplo: 2026-01-01 a 2026-05-21 → enero completo … mayo parcial hasta el día 21.
    """
    inicio = pd.Timestamp(fecha_inicio)
    fin = pd.Timestamp(fecha_fin)

    # Periodos mensuales (M) que cubren desde el mes de inicio hasta el de fin
    periodos = pd.period_range(inicio.to_period("M"), fin.to_period("M"), freq="M")

    rangos: list[tuple[str, str]] = []
    for periodo in periodos:
        # Límites del mes calendario (primer y último día)
        mes_inicio = periodo.start_time.normalize()
        mes_fin = periodo.end_time.normalize()

        # Recortar al rango solicitado por el usuario
        chunk_inicio = max(inicio, mes_inicio)
        chunk_fin = min(fin, mes_fin)

        rangos.append(
            (chunk_inicio.strftime("%Y-%m-%d"), chunk_fin.strftime("%Y-%m-%d"))
        )

    return rangos


def _generar_rangos_anuales(fecha_inicio: str, fecha_fin: str) -> list[tuple[str, str]]:
    """
    Divide un rango de fechas en sub-rangos por año calendario.

    El primer y último chunk respetan fecha_inicio y fecha_fin del rango total.
    Ejemplo: 2024-06-01 a 2026-02-10 → 2024 parcial, 2025 completo, 2026 parcial.
    """
    inicio = pd.Timestamp(fecha_inicio)
    fin = pd.Timestamp(fecha_fin)

    anios = range(int(inicio.year), int(fin.year) + 1)
    rangos: list[tuple[str, str]] = []

    for anio in anios:
        anio_inicio = pd.Timestamp(year=anio, month=1, day=1)
        anio_fin = pd.Timestamp(year=anio, month=12, day=31)

        chunk_inicio = max(inicio, anio_inicio)
        chunk_fin = min(fin, anio_fin)

        if chunk_inicio <= chunk_fin:
            rangos.append(
                (chunk_inicio.strftime("%Y-%m-%d"), chunk_fin.strftime("%Y-%m-%d"))
            )

    return rangos


def _sql_pb_rango(fecha_inicio: str, fecha_fin: str) -> str:
    """
    Arma la consulta SQL de PB para un sub-rango de fechas.

    Prioridad de versión por fecha:
      1. Tx2  — disponible desde mayo 2026 en adelante
      2. Tx1  — disponible desde mayo 2026 en adelante
      3. TxF  — cubre todo el histórico (2010-01-01 a 2026-04-30)
    El COALESCE garantiza que siempre se seleccione alguna versión,
    incluso para fechas históricas donde sólo existe TxF.
    """
    return f"""
SELECT p.file_date, p.hour, p.pb
FROM energy.price_pb_hourly p
INNER JOIN (
    SELECT
        file_date,
        COALESCE(
            MAX(CASE WHEN version_file = 'Tx2' THEN 'Tx2' END),
            MAX(CASE WHEN version_file = 'Tx1' THEN 'Tx1' END),
            MAX(version_file)
        ) AS version_preferida
    FROM energy.price_pb_hourly
    WHERE file_date BETWEEN '{fecha_inicio}' AND '{fecha_fin}'
    GROUP BY file_date
) v ON p.file_date = v.file_date
   AND p.version_file = v.version_preferida
WHERE p.file_date BETWEEN '{fecha_inicio}' AND '{fecha_fin}'
ORDER BY p.file_date, p.hour
"""


def cargar_pb_sql(
    database_id: int | str,
    fecha_inicio: str,
    fecha_fin: str,
) -> pd.DataFrame:
    """
    Carga precios de bolsa horarios desde energy.price_pb_hourly vía Metabase SQL.

    Consulta la tabla con prioridad de versión Tx2 sobre Tx1 por cada file_date.
    Metabase limita a 2000 filas por consulta; por eso el rango se divide en chunks
    mensuales, se consulta cada mes por separado y se concatenan los resultados.

    Devuelve un DataFrame en formato largo: fecha, hora, precio_bolsa.

    Args:
        database_id: Identificador de la base de datos en Metabase.
        fecha_inicio: Fecha mínima inclusive en formato YYYY-MM-DD.
        fecha_fin: Fecha máxima inclusive en formato YYYY-MM-DD.
    """
    # Validar formato de fechas antes de interpolar en el SQL
    for etiqueta, fecha in (("fecha_inicio", fecha_inicio), ("fecha_fin", fecha_fin)):
        try:
            datetime.strptime(fecha, "%Y-%m-%d")
        except ValueError as error:
            mensaje = (
                f"{etiqueta} debe tener formato YYYY-MM-DD (ej. 2024-01-15). "
                f"Valor recibido: {fecha!r}"
            )
            print(f"Error al cargar PB SQL: {mensaje}")
            raise ValueError(mensaje) from error

    if fecha_inicio > fecha_fin:
        mensaje = "fecha_inicio no puede ser posterior a fecha_fin."
        print(f"Error al cargar PB SQL: {mensaje}")
        raise ValueError(mensaje)

    # Lista de sub-rangos mensuales entre fecha_inicio y fecha_fin
    rangos_mensuales = _generar_rangos_mensuales(fecha_inicio, fecha_fin)
    print(
        f"Cargando PB en {len(rangos_mensuales)} chunk(s) mensual(es) "
        f"({fecha_inicio} a {fecha_fin}, database_id={database_id})."
    )

    chunks_exitosos: list[pd.DataFrame] = []
    chunks_fallidos = 0

    _PAUSA_ENTRE_CHUNKS = 0.15   # segundos entre llamadas — evita rate-limit de Metabase
    _MAX_REINTENTOS      = 2     # reintentos adicionales por chunk ante error transitorio

    for indice, (chunk_inicio, chunk_fin) in enumerate(rangos_mensuales, start=1):
        print(f"  Chunk {indice}/{len(rangos_mensuales)}: {chunk_inicio} a {chunk_fin}")

        # Pequeña pausa para no saturar Metabase con llamadas consecutivas rápidas
        if indice > 1:
            time.sleep(_PAUSA_ENTRE_CHUNKS)

        sql = _sql_pb_rango(chunk_inicio, chunk_fin)

        df_chunk = None
        for intento in range(1, _MAX_REINTENTOS + 2):   # hasta 3 intentos
            try:
                df_chunk = load_metabase_sql(sql, database_id)
                break   # éxito: salir del loop de reintentos
            except Exception as error:
                if intento <= _MAX_REINTENTOS:
                    espera = intento * 1.0   # backoff: 1 s, 2 s, …
                    print(
                        f"  Intento {intento} fallido para chunk {chunk_inicio}..{chunk_fin}: "
                        f"{error}. Reintentando en {espera:.1f} s…"
                    )
                    time.sleep(espera)
                else:
                    chunks_fallidos += 1
                    print(
                        f"  Chunk {chunk_inicio} a {chunk_fin}: todos los intentos fallaron "
                        f"({error}). Se continúa con el siguiente."
                    )

        if df_chunk is None or df_chunk.empty:
            if df_chunk is not None:
                print(f"  Chunk {chunk_inicio} a {chunk_fin}: sin filas.")
            continue

        # Verificar columnas del SELECT en este chunk
        faltantes = COLUMNAS_PB_SQL - set(df_chunk.columns)
        if faltantes:
            chunks_fallidos += 1
            mensaje = (
                f"Chunk {chunk_inicio} a {chunk_fin}: columnas esperadas faltantes "
                f"{sorted(faltantes)}. Recibidas: {list(df_chunk.columns)}"
            )
            print(f"  {mensaje}. Se continúa con el siguiente.")
            continue

        chunks_exitosos.append(df_chunk)
        print(f"  Chunk {chunk_inicio} a {chunk_fin}: {len(df_chunk)} filas.")

    if not chunks_exitosos:
        print(
            f"No se cargó ningún chunk de PB para el rango "
            f"{fecha_inicio} a {fecha_fin} "
            f"(database_id={database_id}, chunks fallidos: {chunks_fallidos})."
        )
        return pd.DataFrame(columns=["fecha", "hora", "precio_bolsa"])

    # Unir todos los meses en un solo DataFrame
    df = pd.concat(chunks_exitosos, ignore_index=True)

    # Renombrar al esquema usado por calcular_spread y transformar_pb
    df_resultado = df.rename(
        columns={"file_date": "fecha", "hour": "hora", "pb": "precio_bolsa"}
    )

    # Normalizar fecha a YYYY-MM-DD (puede venir como timestamp ISO desde Metabase)
    df_resultado["fecha"] = pd.to_datetime(
        df_resultado["fecha"], utc=True, errors="coerce"
    ).dt.strftime("%Y-%m-%d")

    # Tipos numéricos para hora y precio
    df_resultado["hora"] = pd.to_numeric(
        df_resultado["hora"], errors="coerce"
    ).astype(int)
    df_resultado["precio_bolsa"] = pd.to_numeric(
        df_resultado["precio_bolsa"], errors="coerce"
    )

    # Orden cronológico por fecha y hora en el resultado combinado
    df_resultado = df_resultado.sort_values(["fecha", "hora"]).reset_index(drop=True)
    df_resultado = df_resultado[["fecha", "hora", "precio_bolsa"]]

    print(
        f"Total cargado: {len(df_resultado)} filas de precio de bolsa desde SQL "
        f"({fecha_inicio} a {fecha_fin}, database_id={database_id}, "
        f"{len(chunks_exitosos)} chunk(s) OK, {chunks_fallidos} fallido(s))."
    )

    return df_resultado


def cargar_pb_historico(fecha_inicio: str, fecha_fin: str) -> pd.DataFrame:
    """
    Carga el histórico de precios de bolsa (PB) desde la card 1240 de Metabase
    usando un rango de fechas, dividido en chunks ANUALES.

    La card 1240 retorna formato ancho:
        file_date, version_file, H1, H2, ... H24

    Este método transforma el resultado a formato largo:
        fecha, hora (1-24), precio_bolsa
    """
    # Validar formato de fechas
    for etiqueta, fecha in (("fecha_inicio", fecha_inicio), ("fecha_fin", fecha_fin)):
        try:
            datetime.strptime(fecha, "%Y-%m-%d")
        except ValueError as error:
            mensaje = (
                f"{etiqueta} debe tener formato YYYY-MM-DD (ej. 2024-01-15). "
                f"Valor recibido: {fecha!r}"
            )
            print(f"Error al cargar PB histórico: {mensaje}")
            raise ValueError(mensaje) from error

    if fecha_inicio > fecha_fin:
        mensaje = "fecha_inicio no puede ser posterior a fecha_fin."
        print(f"Error al cargar PB histórico: {mensaje}")
        raise ValueError(mensaje)

    rangos_anuales = _generar_rangos_anuales(fecha_inicio, fecha_fin)
    print(
        f"Cargando PB histórico en {len(rangos_anuales)} chunk(s) anual(es) "
        f"({fecha_inicio} a {fecha_fin}, card_id=1240)."
    )

    chunks: list[pd.DataFrame] = []

    for indice, (chunk_inicio, chunk_fin) in enumerate(rangos_anuales, start=1):
        print(f"  Chunk {indice}/{len(rangos_anuales)}: {chunk_inicio} a {chunk_fin}")

        # Consultar la card 1240 con rango de fechas
        df_ancho = load_metabase_card_rango(1240, chunk_inicio, chunk_fin)

        if df_ancho.empty:
            print(f"  Chunk {chunk_inicio} a {chunk_fin}: sin filas.")
            continue

        # Validar columnas esperadas del formato ancho
        columnas_requeridas = {"file_date", "version_file", *COLUMNAS_HORARIAS}
        faltantes = columnas_requeridas - set(df_ancho.columns)
        if faltantes:
            raise ValueError(
                "La card 1240 no devolvió las columnas esperadas. "
                f"Faltantes: {sorted(faltantes)}. Recibidas: {list(df_ancho.columns)}"
            )

        # Transformación a formato largo (una fila por fecha-hora)
        df_trabajo = df_ancho.copy()
        df_trabajo["fecha"] = pd.to_datetime(
            df_trabajo["file_date"], utc=True, errors="coerce"
        ).dt.strftime("%Y-%m-%d")

        df_largo = df_trabajo.melt(
            id_vars=["fecha"],
            value_vars=COLUMNAS_HORARIAS,
            var_name="columna_hora",
            value_name="precio_bolsa",
        )

        df_largo["hora"] = df_largo["columna_hora"].str.extract(r"H(\d+)").astype(int)
        df_largo = df_largo.drop(columns=["columna_hora"])
        df_largo["precio_bolsa"] = pd.to_numeric(df_largo["precio_bolsa"], errors="coerce")

        chunks.append(df_largo[["fecha", "hora", "precio_bolsa"]])
        print(f"  Chunk {chunk_inicio} a {chunk_fin}: {len(df_largo)} filas (largo).")

    if not chunks:
        print(
            f"No se cargaron filas de PB histórico para el rango {fecha_inicio} a {fecha_fin}."
        )
        return pd.DataFrame(columns=["fecha", "hora", "precio_bolsa"])

    df_total = pd.concat(chunks, ignore_index=True)
    df_total = df_total.sort_values(["fecha", "hora"]).reset_index(drop=True)

    print(f"Total cargado PB histórico: {len(df_total)} filas.")
    return df_total


def transformar_pb(df: pd.DataFrame) -> pd.DataFrame:
    """
    Convierte el DataFrame ancho de Metabase a formato largo (tidy).

    Entrada esperada:
        - file_date: fecha ISO (ej. "2010-01-01T00:00:00Z")
        - H1 a H24: precio de bolsa por hora en COP/kWh

    Salida:
        - fecha: YYYY-MM-DD
        - hora: entero del 1 al 24
        - precio_bolsa: valor numérico
    """
    # Validar que existan las columnas necesarias
    if "file_date" not in df.columns:
        raise ValueError("El DataFrame debe incluir la columna 'file_date'.")

    faltantes = [c for c in COLUMNAS_HORARIAS if c not in df.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas horarias en el DataFrame: {faltantes}")

    # Copia para no modificar el DataFrame original
    df_trabajo = df.copy()

    # Extraer solo la fecha (sin hora) en formato YYYY-MM-DD
    df_trabajo["fecha"] = pd.to_datetime(
        df_trabajo["file_date"], utc=True, errors="coerce"
    ).dt.strftime("%Y-%m-%d")

    # Pasar de formato ancho a largo: una fila por cada hora del día
    df_largo = df_trabajo.melt(
        id_vars=["fecha"],
        value_vars=COLUMNAS_HORARIAS,
        var_name="columna_hora",
        value_name="precio_bolsa",
    )

    # Obtener el número de hora (1-24) desde el nombre de columna H1, H2, ...
    df_largo["hora"] = df_largo["columna_hora"].str.extract(r"H(\d+)").astype(int)
    df_largo = df_largo.drop(columns=["columna_hora"])

    # Asegurar tipo numérico en el precio
    df_largo["precio_bolsa"] = pd.to_numeric(df_largo["precio_bolsa"], errors="coerce")

    # Ordenar cronológicamente por fecha y hora
    df_largo = df_largo.sort_values(["fecha", "hora"]).reset_index(drop=True)

    # Dejar solo las columnas finales
    df_resultado = df_largo[["fecha", "hora", "precio_bolsa"]]

    # Verificación rápida en consola
    print("Primeras 5 filas del PB transformado:")
    print(df_resultado.head())

    return df_resultado


def calcular_spread(
    df_pb: pd.DataFrame, precio_contrato: float
) -> tuple[pd.DataFrame, dict[str, float | int]]:
    """
    Calcula el spread horario y un resumen estadístico.

    spread = precio_contrato - precio_bolsa

    Retorna el DataFrame con la columna 'spread' y un diccionario de resumen.
    """
    columnas_requeridas = {"fecha", "hora", "precio_bolsa"}
    faltantes = columnas_requeridas - set(df_pb.columns)
    if faltantes:
        raise ValueError(
            f"El DataFrame de PB no tiene las columnas esperadas. Faltantes: {sorted(faltantes)}"
        )

    df_spread = df_pb.copy()

    # Spread: diferencia entre precio fijo del contrato y precio de bolsa
    df_spread["spread"] = precio_contrato - df_spread["precio_bolsa"]

    total_horas = len(df_spread)
    horas_negativas = int((df_spread["spread"] < 0).sum())

    # Resumen de métricas del spread
    resumen: dict[str, float | int] = {
        "spread_promedio": float(df_spread["spread"].mean()),
        "spread_minimo": float(df_spread["spread"].min()),
        "spread_maximo": float(df_spread["spread"].max()),
        "horas_negativas": horas_negativas,
        "porcentaje_negativo": (
            float(horas_negativas / total_horas * 100) if total_horas > 0 else 0.0
        ),
    }

    print("Resumen del spread:")
    for clave, valor in resumen.items():
        if clave == "porcentaje_negativo":
            print(f"  {clave}: {valor:.2f}%")
        elif isinstance(valor, float):
            print(f"  {clave}: {valor:.4f}")
        else:
            print(f"  {clave}: {valor}")

    return df_spread, resumen
