"""
Cargador de datos de contratos y posición energética desde la API de Olibia Energy.

Reemplaza la carga de los 7 archivos Excel de inventario por llamadas REST a:
  https://integrations.bia.app/ms-olibia-energy/v1

Funciones exportadas:
  get_contracts()                 → lista de contratos validados
  get_contract_hourly(id, f, f)   → DataFrame horario de un contrato
  cargar_posicion_olibia(f, f)    → posición agregada (compra R/NR, venta)
  get_precios_contratos(f, f)     → precio promedio ponderado por hora (PC)
  cargar_posicion_con_demanda(f,f)→ posición + demanda real de Metabase

Todas las funciones leen OLIBIA_URL, OLIBIA_API_KEY, OLIBIA_USER_EMAIL y
OLIBIA_USER_ID desde el archivo .env del backend.
"""

from __future__ import annotations

import concurrent.futures
import json
import os
import urllib.error
import urllib.request
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

# ─── Variables de entorno ─────────────────────────────────────────────────────
_ENV_PATH = Path(__file__).resolve().parent / ".env"
load_dotenv(_ENV_PATH)

OLIBIA_URL        = os.getenv("OLIBIA_URL", "https://integrations.bia.app/ms-olibia-energy/v1")
OLIBIA_API_KEY    = os.getenv("OLIBIA_API_KEY", "")
OLIBIA_USER_EMAIL = os.getenv("OLIBIA_USER_EMAIL", "integrations@bia.app")
OLIBIA_USER_ID    = os.getenv("OLIBIA_USER_ID", "1")

# IDs de cards de Metabase para demanda horaria de clientes
_CARD_DEMANDA_R  = 9440   # Demanda regulada horaria
_CARD_DEMANDA_NR = 9439   # Demanda no regulada horaria

# Columnas horarias estándar H1..H24 para conversión ancho→largo
_COLUMNAS_H = [f"H{i}" for i in range(1, 25)]

# Hilos para carga paralela de contratos
_MAX_WORKERS = 8

# Timeout en segundos por petición HTTP a la API de Olibia
_TIMEOUT = 30


# ─── Utilidades internas ──────────────────────────────────────────────────────

def _headers() -> dict[str, str]:
    """Construye los headers requeridos por la API de Olibia."""
    return {
        "Api-key":      OLIBIA_API_KEY,
        "X-User-Email": OLIBIA_USER_EMAIL,
        "X-User-ID":    OLIBIA_USER_ID,
        "Content-Type": "application/json",
    }


def _get(path: str, params: dict | None = None) -> dict:
    """
    Realiza un GET a la API de Olibia y retorna el cuerpo JSON parseado.

    Parámetros:
        path   : Ruta relativa (ej. "/contracts" o "/contracts/{id}/hourly").
        params : Query params opcionales (ej. {"start_date": "2026-01-01"}).

    Lanza RuntimeError ante respuesta 4xx/5xx, o URLError ante fallo de red.
    """
    # Construir URL con query string si hay parámetros
    url = f"{OLIBIA_URL.rstrip('/')}{path}"
    if params:
        qs = "&".join(f"{k}={v}" for k, v in params.items())
        url = f"{url}?{qs}"

    solicitud = urllib.request.Request(url, headers=_headers(), method="GET")

    try:
        with urllib.request.urlopen(solicitud, timeout=_TIMEOUT) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        # Leer cuerpo del error para mostrar mensaje de la API
        detalle = exc.read().decode("utf-8", errors="replace") if exc.fp else ""
        raise RuntimeError(
            f"Olibia API HTTP {exc.code} en {path}: {detalle}"
        ) from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Olibia API conexión fallida en {path}: {exc.reason}") from exc


def _grilla_vacia(start_date: str, end_date: str) -> pd.DataFrame:
    """
    Crea una grilla date × hora (1-24) inicializada en cero para el rango dado.

    Columnas: date, hour, compra_r_kwh, compra_nr_kwh, venta_kwh, posicion_neta_kwh.
    Se usa cuando no hay contratos activos en el período o como base del merge.
    """
    fechas = pd.date_range(start=start_date, end=end_date, freq="D")
    filas = [
        {"date": f.strftime("%Y-%m-%d"), "hour": h}
        for f in fechas
        for h in range(1, 25)
    ]
    df = pd.DataFrame(filas) if filas else pd.DataFrame(columns=["date", "hour"])
    for col in ["compra_r_kwh", "compra_nr_kwh", "venta_kwh", "posicion_neta_kwh"]:
        df[col] = 0.0
    return df


def _ancho_a_largo(df_ancho: pd.DataFrame, col_nombre: str) -> pd.DataFrame:
    """
    Convierte un DataFrame de demanda formato ancho (H1..H24) a formato largo.

    Entrada esperada:  file_date | H1 | H2 | … | H24
    Salida:            date | hour | col_nombre

    Detecta automáticamente la columna de fecha (file_date, fecha o date).
    """
    if df_ancho.empty:
        return pd.DataFrame(columns=["date", "hour", col_nombre])

    # Columnas horarias que efectivamente están en el DataFrame
    cols_h = [c for c in _COLUMNAS_H if c in df_ancho.columns]
    if not cols_h:
        return pd.DataFrame(columns=["date", "hour", col_nombre])

    df = df_ancho.copy()

    # Detectar columna de fecha por nombre
    col_fecha = next(
        (c for c in df.columns if c.lower() in ("file_date", "fecha", "date")),
        df.columns[0],
    )
    df["date"] = pd.to_datetime(
        df[col_fecha], utc=True, errors="coerce"
    ).dt.strftime("%Y-%m-%d")

    # Pivotar de ancho a largo
    df_largo = df.melt(
        id_vars=["date"],
        value_vars=cols_h,
        var_name="col_hora",
        value_name=col_nombre,
    )
    df_largo["hour"]     = df_largo["col_hora"].str.extract(r"H(\d+)").astype(int)
    df_largo[col_nombre] = pd.to_numeric(df_largo[col_nombre], errors="coerce").fillna(0.0)

    return (
        df_largo[["date", "hour", col_nombre]]
        .sort_values(["date", "hour"])
        .reset_index(drop=True)
    )


# ═══════════════════════════════════════════════════════════════════════════════
# FUNCIÓN 1 — Listar contratos validados
# ═══════════════════════════════════════════════════════════════════════════════

def get_contracts() -> list[dict]:
    """
    Retorna la lista de contratos validados de Olibia Energy.

    Consulta GET /contracts y filtra únicamente los que tienen
    status = "validated". Extrae solo los campos necesarios para
    la lógica del portafolio.

    Campos retornados por contrato:
        id, contract_number, type, operation, market_type,
        modalidad, day_type, start_date (YYYY-MM-DD), end_date (YYYY-MM-DD), status

    Ejemplo:
        contratos = get_contracts()
        compras = [c for c in contratos if c["operation"] == "Compra"]
    """
    datos = _get("/contracts")
    items = datos.get("items", [])

    contratos_validados = []
    for c in items:
        if c.get("status") != "validated":
            continue
        contratos_validados.append({
            "id":              c["id"],
            "contract_number": c.get("contract_number", ""),
            "type":            c.get("type", ""),
            "operation":       c.get("operation", ""),     # "Compra" | "Venta"
            "market_type":     c.get("market_type", ""),   # "REGULADO" | "NO REGULADO"
            "modalidad":       c.get("modalidad", ""),     # "PC" | "PD CON TOPE"
            "day_type":        c.get("day_type", "TO"),    # "TO" | "DOFEORSA"
            "start_date":      (c.get("start_date") or "")[:10],
            "end_date":        (c.get("end_date")   or "")[:10],
            "status":          c.get("status", ""),
        })

    print(
        f"[olibia] Contratos validados: {len(contratos_validados)} "
        f"de {len(items)} total"
    )
    return contratos_validados


# ═══════════════════════════════════════════════════════════════════════════════
# FUNCIÓN 2 — Datos horarios de un contrato individual
# ═══════════════════════════════════════════════════════════════════════════════

def get_contract_hourly(
    contract_id: str,
    start_date: str,
    end_date: str,
) -> pd.DataFrame:
    """
    Retorna los datos horarios de energía de un contrato en el rango dado.

    Consulta GET /contracts/{id}/hourly?start_date=...&end_date=...

    Parámetros:
        contract_id : UUID del contrato (campo "id" de get_contracts()).
        start_date  : Fecha inicio YYYY-MM-DD (inclusive).
        end_date    : Fecha fin YYYY-MM-DD (inclusive).

    Retorna DataFrame con columnas:
        date              : YYYY-MM-DD
        hour              : entero 1-24
        day_type          : tipo de día reportado por la API ("TO", "OR", "SA", etc.)
        quantity          : kWh del período; NEGATIVO para contratos de Venta
        price             : COP/kWh precio base del contrato (None en PLD)
        fixed_price       : COP/kWh ajustado por IPP (None en PLD)
        projected_price   : COP/kWh proyectado (None en PLD)
        is_projected_data : True si la cantidad es proyectada, False si es real

    Nota PLD: Los contratos con modalidad "PD CON TOPE" no tienen precio
    pactado (precio = bolsa). Para ellos, price, fixed_price y projected_price
    serán None. La cantidad aún refleja la energía asignada.
    """
    datos = _get(
        f"/contracts/{contract_id}/hourly",
        params={"start_date": start_date, "end_date": end_date},
    )
    items = datos.get("items", [])

    if not items:
        return pd.DataFrame(columns=[
            "date", "hour", "day_type", "quantity",
            "price", "fixed_price", "projected_price", "is_projected_data",
        ])

    # Normalizar cada registro — tolera contratos PLD sin campo price
    filas = []
    for it in items:
        filas.append({
            "date":            str(it.get("date", "")),
            "hour":            int(it.get("hour", 0)),
            "day_type":        str(it.get("day_type", "TO")),
            "quantity":        float(it.get("quantity") or 0.0),
            # Campos de precio opcionales (ausentes en contratos PLD)
            "price":           (float(it["price"])           if it.get("price")           is not None else None),
            "fixed_price":     (float(it["fixed_price"])     if it.get("fixed_price")     is not None else None),
            "projected_price": (float(it["projected_price"]) if it.get("projected_price") is not None else None),
            "is_projected_data": bool(it.get("is_projected_data", False)),
        })

    df = pd.DataFrame(filas)
    # Normalizar fecha a YYYY-MM-DD
    df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.strftime("%Y-%m-%d")
    return df.sort_values(["date", "hour"]).reset_index(drop=True)


# ═══════════════════════════════════════════════════════════════════════════════
# FUNCIÓN 3 — Posición agregada del portafolio Olibia
# ═══════════════════════════════════════════════════════════════════════════════

def cargar_posicion_olibia(start_date: str, end_date: str) -> pd.DataFrame:
    """
    Carga y agrega la posición energética de Olibia para un período.

    Reemplaza la lógica de cargar_inventario() + calcular_posicion_periodo()
    que antes requería los 7 archivos Excel de inventario por tipo de día.
    La API ya entrega la cantidad correcta por fecha y hora real, por lo
    que no es necesario calcular festivos colombianos aquí.

    Proceso:
        1. Obtiene lista de contratos validados (get_contracts).
        2. Filtra contratos cuyo período se solapa con [start_date, end_date].
        3. Carga datos horarios de todos los contratos activos en PARALELO
           (hasta _MAX_WORKERS hilos simultáneos).
        4. Clasifica cada fragmento en compra_r, compra_nr o venta.
        5. Agrega por (date, hour) sumando quantities.
        6. Construye grilla completa date × hora y rellena con ceros
           donde no haya contratos activos.

    Parámetros:
        start_date : Fecha inicio YYYY-MM-DD (inclusive).
        end_date   : Fecha fin YYYY-MM-DD (inclusive).

    Retorna DataFrame con columnas:
        date              : YYYY-MM-DD
        hour              : entero 1-24
        compra_r_kwh      : kWh compra regulado (suma de contratos R)
        compra_nr_kwh     : kWh compra no regulado (suma de contratos NR)
        venta_kwh         : kWh venta (valor absoluto; positivo)
        posicion_neta_kwh : compra_r + compra_nr − venta

    Clasificación de contratos:
        operation="Compra" + market_type="REGULADO"    → compra_r_kwh
        operation="Compra" + market_type="NO REGULADO" → compra_nr_kwh
        operation="Venta"                              → venta_kwh (abs)
    """
    contratos = get_contracts()

    # Contratos cuyo período se intersecta con el rango solicitado
    contratos_activos = [
        c for c in contratos
        if c["start_date"] <= end_date and c["end_date"] >= start_date
    ]
    print(
        f"[olibia] Contratos activos en {start_date}..{end_date}: "
        f"{len(contratos_activos)} de {len(contratos)}"
    )

    if not contratos_activos:
        # Sin contratos en el período → posición cero
        return _grilla_vacia(start_date, end_date)

    # ── Carga paralela — un hilo por contrato ─────────────────────────────────
    fragmentos_compra_r:  list[pd.DataFrame] = []
    fragmentos_compra_nr: list[pd.DataFrame] = []
    fragmentos_venta:     list[pd.DataFrame] = []

    def _cargar_contrato(contrato: dict) -> tuple[dict, pd.DataFrame]:
        """Carga datos horarios de un contrato; retorna vacío ante error."""
        try:
            df = get_contract_hourly(contrato["id"], start_date, end_date)
            return contrato, df
        except Exception as exc:
            print(
                f"[olibia] Error cargando {contrato['contract_number']}: {exc}. "
                "Se omite este contrato de la posición."
            )
            return contrato, pd.DataFrame()

    with concurrent.futures.ThreadPoolExecutor(max_workers=_MAX_WORKERS) as pool:
        futuros = {pool.submit(_cargar_contrato, c): c for c in contratos_activos}
        for futuro in concurrent.futures.as_completed(futuros):
            contrato, df = futuro.result()
            if df.empty:
                continue

            operacion = contrato["operation"]    # "Compra" | "Venta"
            mercado   = contrato["market_type"]  # "REGULADO" | "NO REGULADO"

            # Solo necesitamos date, hour, quantity para la agregación
            df_mini = df[["date", "hour", "quantity"]].copy()

            if operacion == "Compra" and mercado == "REGULADO":
                fragmentos_compra_r.append(df_mini)

            elif operacion == "Compra" and mercado == "NO REGULADO":
                fragmentos_compra_nr.append(df_mini)

            elif operacion == "Venta":
                # Las cantidades de venta llegan NEGATIVAS desde la API;
                # almacenamos el valor absoluto para restar explícitamente
                df_abs = df_mini.copy()
                df_abs["quantity"] = df_abs["quantity"].abs()
                fragmentos_venta.append(df_abs)

    # ── Agregación por (date, hour) para cada categoría ───────────────────────
    def _agregar(fragmentos: list[pd.DataFrame], col: str) -> pd.DataFrame:
        """Concatena fragmentos del mismo tipo y suma quantity por (date, hour)."""
        if not fragmentos:
            return pd.DataFrame(columns=["date", "hour", col])
        return (
            pd.concat(fragmentos, ignore_index=True)
            .groupby(["date", "hour"])["quantity"]
            .sum()
            .reset_index()
            .rename(columns={"quantity": col})
        )

    df_r  = _agregar(fragmentos_compra_r,  "compra_r_kwh")
    df_nr = _agregar(fragmentos_compra_nr, "compra_nr_kwh")
    df_v  = _agregar(fragmentos_venta,     "venta_kwh")

    # ── Grilla base y merge con los tres tipos ────────────────────────────────
    df_pos = _grilla_vacia(start_date, end_date)

    for df_tipo in (df_r, df_nr, df_v):
        if not df_tipo.empty:
            # Eliminar la columna precargada de ceros antes del merge
            col = [c for c in df_tipo.columns if c not in ("date", "hour")][0]
            df_pos = df_pos.drop(columns=[col], errors="ignore")
            df_pos = df_pos.merge(df_tipo, on=["date", "hour"], how="left")

    # Rellenar NaN donde no hay contratos en esa hora
    for col in ["compra_r_kwh", "compra_nr_kwh", "venta_kwh"]:
        if col not in df_pos.columns:
            df_pos[col] = 0.0
        else:
            df_pos[col] = df_pos[col].fillna(0.0)

    # Posición neta = compra_r + compra_nr − venta
    df_pos["posicion_neta_kwh"] = (
        df_pos["compra_r_kwh"]
        + df_pos["compra_nr_kwh"]
        - df_pos["venta_kwh"]
    )

    df_pos = df_pos.sort_values(["date", "hour"]).reset_index(drop=True)
    print(
        f"[olibia] Posición lista: {len(df_pos)} filas "
        f"({df_pos['date'].nunique()} días × 24 horas)"
    )
    return df_pos


# ═══════════════════════════════════════════════════════════════════════════════
# FUNCIÓN 4 — Precio promedio ponderado de contratos PC
# ═══════════════════════════════════════════════════════════════════════════════

def get_precios_contratos(start_date: str, end_date: str) -> pd.DataFrame:
    """
    Calcula el precio promedio ponderado (PPP) por hora de contratos PC.

    Solo considera contratos con modalidad "PC" (Precio Constante), que
    tienen el campo fixed_price definido en la API. Los contratos PLD
    se excluyen porque su precio es el de la bolsa (no un precio pactado).

    Fórmula PPP por (date, hour, operation, market_type):
        PPP = Σ(|quantity| × fixed_price) / Σ(|quantity|)

    Parámetros:
        start_date : Fecha inicio YYYY-MM-DD.
        end_date   : Fecha fin YYYY-MM-DD.

    Retorna DataFrame con columnas:
        date                            : YYYY-MM-DD
        hour                            : entero 1-24
        operation                       : "Compra" | "Venta"
        market_type                     : "REGULADO" | "NO REGULADO"
        precio_promedio_ponderado_cop_kwh : PPP en COP/kWh (4 decimales)

    Los contratos PLD (sin fixed_price) no aportan al cálculo.
    Horas sin contratos PC activos no aparecen en el resultado.
    """
    contratos = get_contracts()

    # Filtrar solo contratos PC activos en el período solicitado
    contratos_pc = [
        c for c in contratos
        if c.get("modalidad") == "PC"
        and c["start_date"] <= end_date
        and c["end_date"] >= start_date
    ]
    print(f"[olibia] Contratos PC para precios: {len(contratos_pc)}")

    if not contratos_pc:
        return pd.DataFrame(columns=[
            "date", "hour", "operation", "market_type",
            "precio_promedio_ponderado_cop_kwh",
        ])

    fragmentos: list[pd.DataFrame] = []

    for contrato in contratos_pc:
        try:
            df = get_contract_hourly(contrato["id"], start_date, end_date)
            if df.empty:
                continue

            # Solo filas con precio definido y cantidad distinta de cero
            df_pc = df[df["fixed_price"].notna() & (df["quantity"].abs() > 0)].copy()
            if df_pc.empty:
                continue

            df_pc["operation"]    = contrato["operation"]
            df_pc["market_type"]  = contrato["market_type"]
            df_pc["qty_abs"]      = df_pc["quantity"].abs()
            df_pc["qty_x_precio"] = df_pc["qty_abs"] * df_pc["fixed_price"]

            fragmentos.append(
                df_pc[[
                    "date", "hour", "operation", "market_type",
                    "qty_abs", "qty_x_precio",
                ]]
            )

        except Exception as exc:
            print(
                f"[olibia] Error en precios de "
                f"{contrato['contract_number']}: {exc}"
            )

    if not fragmentos:
        return pd.DataFrame(columns=[
            "date", "hour", "operation", "market_type",
            "precio_promedio_ponderado_cop_kwh",
        ])

    df_all = pd.concat(fragmentos, ignore_index=True)

    # PPP por (date, hour, operation, market_type)
    agg = (
        df_all
        .groupby(["date", "hour", "operation", "market_type"])
        .agg(
            total_kwh =("qty_abs",      "sum"),
            total_cop =("qty_x_precio", "sum"),
        )
        .reset_index()
    )
    agg["precio_promedio_ponderado_cop_kwh"] = (
        agg["total_cop"] / agg["total_kwh"]
    ).round(4)

    return agg[[
        "date", "hour", "operation", "market_type",
        "precio_promedio_ponderado_cop_kwh",
    ]]


# ═══════════════════════════════════════════════════════════════════════════════
# FUNCIÓN 5 — Posición con demanda real de clientes
# ═══════════════════════════════════════════════════════════════════════════════

def cargar_posicion_con_demanda(start_date: str, end_date: str) -> pd.DataFrame:
    """
    Carga la posición de contratos Olibia más la demanda real de clientes.

    Combina dos fuentes:
        - Energía de contratos Olibia (compra R, compra NR, venta) via API.
        - Demanda regulada   (card Metabase 9440, formato ancho H1..H24).
        - Demanda no regulada(card Metabase 9439, formato ancho H1..H24).

    La demanda representa la energía que los clientes de BIA efectivamente
    consumieron; se trata como una "salida" adicional que reduce el excedente
    en bolsa.

    Fórmula posición neta total:
        posicion_neta_total = compra_r + compra_nr − venta
                              − demanda_r − demanda_nr

    Parámetros:
        start_date : Fecha inicio YYYY-MM-DD.
        end_date   : Fecha fin YYYY-MM-DD.

    Retorna DataFrame con columnas:
        date, hour,
        compra_r_kwh, compra_nr_kwh, venta_kwh,
        posicion_neta_kwh,           ← solo contratos (sin demanda)
        demanda_r_kwh, demanda_nr_kwh,
        posicion_neta_total_kwh      ← contratos menos demanda real

    Si no se puede cargar la demanda de Metabase, se usan ceros
    y se registra un warning en consola.
    """
    # 1. Posición de contratos desde la API de Olibia
    df_pos = cargar_posicion_olibia(start_date, end_date)

    # 2. Demanda desde Metabase (formato ancho H1..H24 → largo date/hour)
    try:
        from data_loader import load_metabase_card_rango

        # Card 9440: demanda regulada horaria
        df_dr_ancho  = load_metabase_card_rango(_CARD_DEMANDA_R,  start_date, end_date)
        # Card 9439: demanda no regulada horaria
        df_dnr_ancho = load_metabase_card_rango(_CARD_DEMANDA_NR, start_date, end_date)

        df_dr  = _ancho_a_largo(df_dr_ancho,  "demanda_r_kwh")
        df_dnr = _ancho_a_largo(df_dnr_ancho, "demanda_nr_kwh")

        if not df_dr.empty:
            df_pos = df_pos.merge(df_dr,  on=["date", "hour"], how="left")
        if not df_dnr.empty:
            df_pos = df_pos.merge(df_dnr, on=["date", "hour"], how="left")

        print(
            f"[olibia] Demanda cargada desde Metabase "
            f"({start_date} a {end_date})"
        )

    except Exception as exc:
        print(
            f"[olibia] No se pudo cargar demanda de Metabase: {exc}. "
            "Se usarán ceros para demanda_r_kwh y demanda_nr_kwh."
        )

    # Asegurar columnas de demanda — ceros donde Metabase no tenga datos
    for col in ("demanda_r_kwh", "demanda_nr_kwh"):
        if col not in df_pos.columns:
            df_pos[col] = 0.0
        else:
            df_pos[col] = df_pos[col].fillna(0.0)

    # 3. Posición neta total = contratos − demanda real
    df_pos["posicion_neta_total_kwh"] = (
        df_pos["posicion_neta_kwh"]
        - df_pos["demanda_r_kwh"]
        - df_pos["demanda_nr_kwh"]
    )

    return df_pos
