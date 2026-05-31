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
from datetime import date, datetime, timedelta
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


# ─── Festivos colombianos (para clasificar tipo de día de la demanda) ─────────
# Se usan en cargar_posicion_con_demanda para saber qué columna de la card
# de Metabase aplicar: demanda_*_ordinario | _sabado | _festivo.
# Domingo se agrupa con festivo (no hay columna separada en las cards).

def _siguiente_lunes(fecha_base: date) -> date:
    """Aplica Ley Emiliani: desplaza algunas festividades al lunes siguiente."""
    dias_hasta_lunes = (7 - fecha_base.weekday()) % 7
    return fecha_base + timedelta(days=dias_hasta_lunes)


def _fecha_pascua(anio: int) -> date:
    """Calcula el domingo de Pascua usando el algoritmo de Meeus (gregoriano)."""
    a = anio % 19
    b = anio // 100
    c = anio % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes = (h + l - 7 * m + 114) // 31
    dia = ((h + l - 7 * m + 114) % 31) + 1
    return date(anio, mes, dia)


def _festivos_colombia(anio: int) -> set[date]:
    """Retorna el conjunto de fechas festivas en Colombia para el año dado."""
    pascua = _fecha_pascua(anio)

    # Festivos de fecha fija
    festivos: set[date] = {
        date(anio,  1,  1),   # Año Nuevo
        date(anio,  5,  1),   # Día del Trabajo
        date(anio,  7, 20),   # Independencia
        date(anio,  8,  7),   # Batalla de Boyacá
        date(anio, 12,  8),   # Inmaculada Concepción
        date(anio, 12, 25),   # Navidad
    }

    # Festivos con traslado al lunes (Ley Emiliani)
    emiliani = [
        date(anio,  1,  6),   # Reyes Magos
        date(anio,  3, 19),   # San José
        date(anio,  6, 29),   # San Pedro y San Pablo
        date(anio,  8, 15),   # Asunción
        date(anio, 10, 12),   # Día de la Raza
        date(anio, 11,  1),   # Todos los Santos
        date(anio, 11, 11),   # Independencia de Cartagena
    ]
    festivos.update(_siguiente_lunes(f) for f in emiliani)

    # Festivos móviles respecto a Pascua
    festivos.add(pascua - timedelta(days=3))                         # Jueves Santo
    festivos.add(pascua - timedelta(days=2))                         # Viernes Santo
    festivos.add(_siguiente_lunes(pascua + timedelta(days=43)))      # Ascensión
    festivos.add(_siguiente_lunes(pascua + timedelta(days=64)))      # Corpus Christi
    festivos.add(_siguiente_lunes(pascua + timedelta(days=71)))      # Sagrado Corazón

    return festivos


def _tipo_dia(fecha_obj: datetime) -> str:
    """
    Clasifica una fecha como ordinario, sabado, domingo o festivo.

    Devuelve:
        "ordinario" → lunes a viernes no festivos
        "sabado"    → sábados
        "domingo"   → domingos (sin festivo)
        "festivo"   → festivos colombianos (cualquier día de la semana)
    """
    fecha_d = fecha_obj.date()
    if fecha_d in _festivos_colombia(fecha_obj.year):
        return "festivo"
    dia_semana = fecha_obj.weekday()   # lunes=0 … domingo=6
    if dia_semana == 5:
        return "sabado"
    if dia_semana == 6:
        return "domingo"
    return "ordinario"


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
        posicion_neta_kwh : (venta) − (compra_r + compra_nr)
                            Positivo → BIA debe comprar en bolsa.
                            Negativo → BIA tiene exceso para vender en bolsa.

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

    # Posición neta en bolsa = (venta) − (compra_r + compra_nr)
    # Positivo → BIA necesita comprar en bolsa (compromisos de venta > compras de contratos)
    # Negativo → BIA tiene exceso para vender en bolsa (compras > compromisos de venta)
    df_pos["posicion_neta_kwh"] = (
        df_pos["venta_kwh"]
        - df_pos["compra_r_kwh"]
        - df_pos["compra_nr_kwh"]
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

def get_precios_contratos(start_date: str, end_date: str) -> dict:
    """
    Calcula el PPP (Precio Promedio Ponderado) de contratos PC por categoría.

    Solo considera contratos con modalidad "PC" (Precio Constante), que tienen
    el campo fixed_price definido en la API. Los contratos PLD se excluyen
    porque su precio efectivo es el de la bolsa en cada hora (sin precio pactado).

    Tipo del PPP por categoría:
      "Indexado"   → todos los datos tienen is_projected_data=False
                     (precios históricos confirmados, ya ajustados por IPP)
      "Proyectado" → algún dato tiene is_projected_data=True
                     (hay al menos una hora futura proyectada)
      "Sin datos"  → no hay contratos PC activos en esa categoría en el período

    Fórmula PPP por categoría:
        PPP = Σ(|quantity| × fixed_price) / Σ(|quantity|)

    Parámetros:
        start_date : Fecha inicio YYYY-MM-DD (inclusive).
        end_date   : Fecha fin YYYY-MM-DD (inclusive).

    Retorna:
        {
            "compra_r":      {"ppp": float | None, "tipo": str},
            "compra_nr":     {"ppp": float | None, "tipo": str},
            "venta":         {"ppp": float | None, "tipo": str},
            "pld_excluidos": int,   # contratos PLD no incluidos en el cálculo
            "contratos_pc":  int,   # contratos PC incluidos en el cálculo
        }
    """
    contratos = get_contracts()

    # Separar contratos activos en el período en PC y PLD
    contratos_activos = [
        c for c in contratos
        if c["start_date"] <= end_date and c["end_date"] >= start_date
    ]
    contratos_pc = [c for c in contratos_activos if c.get("modalidad") == "PC"]
    n_pld        = len(contratos_activos) - len(contratos_pc)

    print(
        f"[olibia] PPP — contratos PC: {len(contratos_pc)}, "
        f"PLD excluidos: {n_pld}"
    )

    # Acumuladores por categoría de operación
    acum: dict[str, dict] = {
        "compra_r":  {"qty": 0.0, "cop": 0.0, "proyectado": False, "tiene_datos": False},
        "compra_nr": {"qty": 0.0, "cop": 0.0, "proyectado": False, "tiene_datos": False},
        "venta":     {"qty": 0.0, "cop": 0.0, "proyectado": False, "tiene_datos": False},
    }

    for contrato in contratos_pc:
        try:
            df = get_contract_hourly(contrato["id"], start_date, end_date)
            if df.empty:
                continue

            # Solo filas con precio fijo definido y cantidad distinta de cero
            df_pc = df[
                df["fixed_price"].notna() & (df["quantity"].abs() > 0)
            ].copy()
            if df_pc.empty:
                continue

            # Determinar categoría según operación y mercado
            operacion = contrato["operation"]    # "Compra" | "Venta"
            mercado   = contrato["market_type"]  # "REGULADO" | "NO REGULADO"
            if operacion == "Compra" and mercado == "REGULADO":
                clave = "compra_r"
            elif operacion == "Compra" and mercado == "NO REGULADO":
                clave = "compra_nr"
            else:
                clave = "venta"

            # Acumular energía y COP ponderados
            qty_abs = df_pc["quantity"].abs()
            acum[clave]["qty"] += float(qty_abs.sum())
            acum[clave]["cop"] += float((qty_abs * df_pc["fixed_price"]).sum())
            acum[clave]["tiene_datos"] = True

            # Si alguna fila tiene datos proyectados, marcar la categoría
            if df_pc["is_projected_data"].any():
                acum[clave]["proyectado"] = True

        except Exception as exc:
            print(
                f"[olibia] Error en PPP de {contrato['contract_number']}: {exc}"
            )

    def _categoria(clave: str) -> dict:
        """Calcula PPP y tipo para una categoría de operación."""
        a = acum[clave]
        if not a["tiene_datos"] or a["qty"] == 0.0:
            return {"ppp": None, "tipo": "Sin datos"}
        return {
            "ppp":  round(a["cop"] / a["qty"], 4),
            "tipo": "Proyectado" if a["proyectado"] else "Indexado",
        }

    resultado = {
        "compra_r":      _categoria("compra_r"),
        "compra_nr":     _categoria("compra_nr"),
        "venta":         _categoria("venta"),
        "pld_excluidos": n_pld,
        "contratos_pc":  len(contratos_pc),
    }

    print(
        f"[olibia] PPP — "
        f"Compra R: {resultado['compra_r']['ppp']} ({resultado['compra_r']['tipo']})  "
        f"Compra NR: {resultado['compra_nr']['ppp']} ({resultado['compra_nr']['tipo']})  "
        f"Venta: {resultado['venta']['ppp']} ({resultado['venta']['tipo']})"
    )
    return resultado


# ═══════════════════════════════════════════════════════════════════════════════
# FUNCIÓN 5 — Posición con demanda real de clientes
# ═══════════════════════════════════════════════════════════════════════════════

def cargar_posicion_con_demanda(start_date: str, end_date: str) -> pd.DataFrame:
    """
    Calcula la posición energética neta incorporando la demanda de usuarios finales.

    Combina la posición de contratos Olibia (API) con la demanda promedio
    horaria por tipo de día almacenada en las cards de Metabase.

    Estructura de las cards:
        Card 9440 — demanda regulada (una fila por hora 1-24):
            hour | demanda_r_ordinario | demanda_r_sabado | demanda_r_festivo
        Card 9439 — demanda no regulada (una fila por hora 1-24):
            hour | demanda_nr_ordinario | demanda_nr_sabado | demanda_nr_festivo

    Las cards contienen valores PROMEDIO en kWh por hora y tipo de día.
    Para asignar la demanda correcta a cada fila (fecha, hora) del período:
        1. Se clasifica la fecha: ordinario | sabado | domingo | festivo.
           (domingo se mapea a la columna "festivo" de las cards, que
           no tiene columna propia.)
        2. Se consulta la tabla de lookup:
               demanda_r  = card9440[tipo_dia][hora]
               demanda_nr = card9439[tipo_dia][hora]

    Fórmula de posición neta en bolsa:
        posicion_neta_kwh = (venta + demanda_r + demanda_nr)
                            − (compra_r + compra_nr)
        Positivo → BIA debe comprar en bolsa (compromisos > compras de contratos).
        Negativo → BIA tiene exceso para vender en bolsa (compras > compromisos).

    Parámetros:
        start_date : Fecha inicio YYYY-MM-DD (inclusive).
        end_date   : Fecha fin YYYY-MM-DD (inclusive).

    Retorna DataFrame con columnas:
        date, hour, tipo_dia,
        compra_r_kwh, compra_nr_kwh, venta_kwh,
        demanda_r_kwh, demanda_nr_kwh,
        posicion_neta_kwh   ← contratos menos demanda de usuarios finales

    Si Metabase no está disponible se asignan ceros a las columnas de
    demanda y se registra un aviso en consola.
    """
    # ── 1. Posición de contratos desde la API de Olibia ───────────────────────
    df_pos = cargar_posicion_olibia(start_date, end_date)

    # ── 2. Clasificar tipo_dia por cada fecha única del período ───────────────
    # ordinario / sabado / domingo / festivo según calendario colombiano.
    # Se calcula aquí porque determina qué columna de la card de Metabase usar.
    tipo_dia_map: dict[str, str] = {}
    for fecha_str in df_pos["date"].unique():
        try:
            tipo_dia_map[fecha_str] = _tipo_dia(datetime.strptime(fecha_str, "%Y-%m-%d"))
        except ValueError:
            tipo_dia_map[fecha_str] = "ordinario"

    df_pos["tipo_dia"] = df_pos["date"].map(tipo_dia_map)

    # ── 3. Cargar y aplicar demanda desde Metabase ────────────────────────────
    try:
        from data_loader import load_metabase_card

        # Card 9440: demanda regulada — filas por hora, columnas por tipo de día
        df_dr  = load_metabase_card(_CARD_DEMANDA_R)
        # Card 9439: demanda no regulada — misma estructura
        df_dnr = load_metabase_card(_CARD_DEMANDA_NR)

        # Normalizar columna de hora a entero
        df_dr["hour"]  = pd.to_numeric(df_dr["hour"],  errors="coerce").astype(int)
        df_dnr["hour"] = pd.to_numeric(df_dnr["hour"], errors="coerce").astype(int)

        # Normalizar columnas de kWh a float (algunos valores pueden venir como str)
        for col in ["demanda_r_ordinario", "demanda_r_sabado", "demanda_r_festivo"]:
            df_dr[col] = pd.to_numeric(df_dr[col], errors="coerce").fillna(0.0)
        for col in ["demanda_nr_ordinario", "demanda_nr_sabado", "demanda_nr_festivo"]:
            df_dnr[col] = pd.to_numeric(df_dnr[col], errors="coerce").fillna(0.0)

        # Construir tabla de lookup: (tipo_dia, hora) → kWh
        # domingo se mapea a la columna "festivo" (no existe columna separada)
        lookup_r: dict[tuple[str, int], float] = {}
        for _, fila in df_dr.iterrows():
            h = int(fila["hour"])
            lookup_r[("ordinario", h)] = float(fila["demanda_r_ordinario"])
            lookup_r[("sabado",    h)] = float(fila["demanda_r_sabado"])
            lookup_r[("domingo",   h)] = float(fila["demanda_r_festivo"])   # sin col propia → festivo
            lookup_r[("festivo",   h)] = float(fila["demanda_r_festivo"])

        lookup_nr: dict[tuple[str, int], float] = {}
        for _, fila in df_dnr.iterrows():
            h = int(fila["hour"])
            lookup_nr[("ordinario", h)] = float(fila["demanda_nr_ordinario"])
            lookup_nr[("sabado",    h)] = float(fila["demanda_nr_sabado"])
            lookup_nr[("domingo",   h)] = float(fila["demanda_nr_festivo"])  # sin col propia → festivo
            lookup_nr[("festivo",   h)] = float(fila["demanda_nr_festivo"])

        # Asignar demanda por fila según (tipo_dia, hora)
        tipo_dia_vec = df_pos["tipo_dia"].tolist()
        hora_vec     = df_pos["hour"].tolist()

        df_pos["demanda_r_kwh"] = [
            lookup_r.get((td, int(h)), 0.0)
            for td, h in zip(tipo_dia_vec, hora_vec)
        ]
        df_pos["demanda_nr_kwh"] = [
            lookup_nr.get((td, int(h)), 0.0)
            for td, h in zip(tipo_dia_vec, hora_vec)
        ]

        print(
            f"[olibia] Demanda cargada desde Metabase "
            f"(cards {_CARD_DEMANDA_R}/{_CARD_DEMANDA_NR}, "
            f"{start_date} a {end_date})"
        )

    except Exception as exc:
        print(
            f"[olibia] No se pudo cargar demanda de Metabase: {exc}. "
            "Se asignan ceros a demanda_r_kwh y demanda_nr_kwh."
        )
        df_pos["demanda_r_kwh"]  = 0.0
        df_pos["demanda_nr_kwh"] = 0.0

    # ── 4. Recalcular posicion_neta incluyendo demanda de usuarios finales ─────
    # Fórmula: (venta + demanda_r + demanda_nr) − (compra_r + compra_nr)
    #   venta      : energía comprometida a contraparte vía contratos de venta
    #   demanda_r  : consumo de clientes regulados
    #   demanda_nr : consumo de clientes no regulados
    #   compra_r   : energía comprada de generadores (mercado regulado)
    #   compra_nr  : energía comprada de generadores (mercado no regulado)
    # Positivo → compromisos superan compras → BIA debe comprar en bolsa.
    # Negativo → compras superan compromisos → BIA tiene exceso para vender en bolsa.
    df_pos["posicion_neta_kwh"] = (
        df_pos["venta_kwh"]
        + df_pos["demanda_r_kwh"]
        + df_pos["demanda_nr_kwh"]
        - df_pos["compra_r_kwh"]
        - df_pos["compra_nr_kwh"]
    )

    return df_pos
