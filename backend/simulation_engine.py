"""
Motor de simulación de impacto de nuevos contratos sobre el portafolio Olibia.
"""

from __future__ import annotations

import calendar
from datetime import datetime
from typing import Any

import pandas as pd

from portfolio_engine import (
    calcular_costo_bolsa,
    calcular_posicion_neta,
    _normalizar_serie_horaria,
    _periodo_desde_fecha,
)

MES_ABR = {
    1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr",  5: "May", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic",
}

_PERFIL_A_CLAVE: dict[str, str] = {
    "ordinario": "compra_or",
    "sabado":    "compra_sa",
    "festivo":   "compra_fe",
}


def _anio_inventario(inventario: dict[str, pd.DataFrame]) -> int:
    """
    Detecta el año del inventario leyendo la primera etiqueta de período.
    Ej. "Ene-26" → 2026. Retorna 2026 si no puede determinarlo.
    """
    for clave in ("compra_or", "compra_nr", "venta"):
        df = inventario.get(clave)
        if df is not None and not df.empty:
            periodo = str(df["Periodo"].iloc[0])          # ej. "Ene-26"
            partes = periodo.split("-")
            if len(partes) == 2 and partes[1].isdigit():
                return 2000 + int(partes[1])
    return 2026


def _proxy_fecha(pb_fecha_str: str, inv_anio: int) -> str:
    """
    Mapea una fecha histórica de PB a la fecha equivalente en el año del inventario.

    Esto permite usar precios de bolsa históricos (escenario) junto con el inventario
    del año vigente (ej. 2026). El tipo de día (festivo, sábado, etc.) se preserva
    usando la fecha original en `calcular_posicion_neta(fecha_tipo_dia=pb_fecha)`.

    Ej.: "2024-03-15" con inv_anio=2026 → "2026-03-15"
    Maneja febrero-29 en años no bisiestos recortando al 28.
    """
    obj = datetime.strptime(pb_fecha_str, "%Y-%m-%d")
    max_dia = calendar.monthrange(inv_anio, obj.month)[1]
    dia = min(obj.day, max_dia)
    return f"{inv_anio}-{obj.month:02d}-{dia:02d}"

# Curva solar típica para Colombia (ecuatorial) — suma = 1.0
# H1-H6 y H19-H24 = 0; campana gaussiana entre H7 y H18
SOLAR_PESOS_24H: list[float] = [
    0,    0,    0,    0,    0,    0,      # H1-H6
    0.02, 0.06, 0.11, 0.15, 0.15, 0.14,  # H7-H12
    0.12, 0.10, 0.08, 0.05, 0.02, 0,     # H13-H18
    0,    0,    0,    0,    0,    0,      # H19-H24
]


def _dias_en_mes(anio: int, mes: int) -> int:
    return calendar.monthrange(anio, mes)[1]


def _distribuir_energia_horaria(
    energia_mensual_kwh: float | None,
    fecha_str: str,
    perfil: str,
    inventario: dict[str, pd.DataFrame],
    bloques: list[dict] | None = None,
    perfil_pesos_24h: list[float] | None = None,
    perfil_excel_12x24: list[list[float]] | None = None,
) -> pd.Series:
    """
    Devuelve una Serie indexada 1..24 con kWh por hora para ese día.

    Prioridad de resolución:
      1. perfil == 'excel' / 'excel_custom'  → usa perfil_excel_12x24 (kWh/mes ÷ días)
      2. perfil == 'bloques'                 → construye pesos desde lista de bloques
      3. perfil == 'solar'                   → usa curva solar colombiana (SOLAR_PESOS_24H)
      4. perfil_pesos_24h provisto ('custom')→ normaliza y escala
      5. perfil == 'plano'                   → distribución uniforme
      6. perfil en inventario Olibia         → forma del inventario (ordinario/sabado/festivo)
      7. fallback                            → plano
    """
    fecha_obj = datetime.strptime(fecha_str, "%Y-%m-%d")
    dias = _dias_en_mes(fecha_obj.year, fecha_obj.month)
    kwh = energia_mensual_kwh or 0.0

    # ── 1. Excel personalizado (12 × 24 kWh/mes por hora) ────────────────────
    if perfil in ("excel", "excel_custom") and perfil_excel_12x24 is not None:
        mes_idx = fecha_obj.month - 1        # 0-11
        if mes_idx < len(perfil_excel_12x24):
            fila = list(perfil_excel_12x24[mes_idx])
            while len(fila) < 24:
                fila.append(0.0)
            return pd.Series(
                [float(v) / dias for v in fila[:24]],
                index=range(1, 25),
                dtype="float64",
            )
        # mes fuera del rango de la matriz → plano con energía provista
        return pd.Series([kwh / dias / 24.0] * 24, index=range(1, 25), dtype="float64")

    # ── 2. Bloques horarios definidos por el usuario ──────────────────────────
    if perfil == "bloques" and bloques:
        pesos = [0.0] * 24
        total_kwh = 0.0
        for b in bloques:
            h_ini = int(b.get("hora_ini", 1))
            h_fin = int(b.get("hora_fin", 24))
            mwh = float(b.get("mwh_mes", 0))
            kwh_bloque = mwh * 1_000.0
            n_horas = max(1, h_fin - h_ini + 1)
            kwh_por_hora = kwh_bloque / n_horas
            for h in range(max(1, h_ini), min(24, h_fin) + 1):
                pesos[h - 1] += kwh_por_hora / dias
            total_kwh += kwh_bloque
        if total_kwh <= 0:
            return pd.Series([0.0] * 24, index=range(1, 25), dtype="float64")
        return pd.Series(pesos, index=range(1, 25), dtype="float64")

    # ── 3. Curva solar colombiana ─────────────────────────────────────────────
    if perfil == "solar":
        energia_diaria = kwh / dias if kwh > 0 else 0.0
        total_pesos = sum(SOLAR_PESOS_24H)
        valores = [
            (p / total_pesos) * energia_diaria if total_pesos > 0 else 0.0
            for p in SOLAR_PESOS_24H
        ]
        return pd.Series(valores, index=range(1, 25), dtype="float64")

    # ── 4. Pesos personalizados explícitos (alias 'custom') ───────────────────
    if perfil_pesos_24h is not None and kwh > 0:
        pesos = list(perfil_pesos_24h)
        while len(pesos) < 24:
            pesos.append(0.0)
        pesos = pesos[:24]
        total = sum(pesos)
        energia_diaria = kwh / dias
        if total > 0:
            return pd.Series(
                [(p / total) * energia_diaria for p in pesos],
                index=range(1, 25),
                dtype="float64",
            )
        return pd.Series([energia_diaria / 24.0] * 24, index=range(1, 25), dtype="float64")

    # ── 5. Perfil plano ───────────────────────────────────────────────────────
    if perfil == "plano":
        return pd.Series([kwh / dias / 24.0] * 24, index=range(1, 25), dtype="float64")

    # ── 6. Perfil con forma del inventario Olibia ─────────────────────────────
    clave = _PERFIL_A_CLAVE.get(perfil)
    if clave and clave in inventario and kwh > 0:
        periodo = _periodo_desde_fecha(fecha_obj)
        filas = inventario[clave].loc[inventario[clave]["Periodo"] == periodo]
        if not filas.empty:
            try:
                shape = _normalizar_serie_horaria(filas, clave)
                total_shape = float(shape.sum())
                if total_shape > 0:
                    energia_diaria = kwh / dias
                    return (shape / total_shape) * energia_diaria
            except ValueError:
                pass

    # ── 7. Fallback plano ─────────────────────────────────────────────────────
    return pd.Series([kwh / dias / 24.0] * 24, index=range(1, 25), dtype="float64")


def simular_contrato(
    inventario: dict[str, pd.DataFrame],
    df_pb: pd.DataFrame,
    tipo: str,
    precio_cop_kwh: float,
    pb_desde: str,
    pb_hasta: str,
    contrato_inicio: str,
    contrato_fin: str,
    tipo_mercado: str,
    perfil_horario: str,
    energia_mensual_mwh: float | None = None,
    bloques: list[dict] | None = None,
    perfil_pesos_24h: list[float] | None = None,
    perfil_excel_12x24: list[list[float]] | None = None,
) -> dict[str, Any]:
    """
    Simula el impacto de un nuevo contrato sobre el portafolio actual usando
    precios de bolsa históricos como escenario.

    Flujo:
      - El PB ya viene cargado para el período pb_desde..pb_hasta.
      - Para cada fecha de PB (escenario histórico):
          · Se calcula la posición neta del inventario usando el mes equivalente
            del año del inventario (ej. Mar-24 PB → posición de Mar-26 inventario).
          · El tipo de día (ordinario/sábado/festivo) se obtiene de la fecha real de PB
            para preservar una mezcla realista de días laborables/festivos.
          · La energía del contrato se distribuye en las 24h según el perfil elegido.
      - contrato_inicio/contrato_fin se registran en el resumen pero no filtran PB.

    Returns dict con resumen_antes/despues, recomendacion, delta_costo_mcop,
    perfil_horario (24h promedio) y por_mes (tabla mensual).
    """
    energia_mensual_kwh = (energia_mensual_mwh or 0.0) * 1_000.0

    # Año del inventario para el mapeo de fechas históricas de PB
    inv_anio = _anio_inventario(inventario)

    # Todas las fechas de PB disponibles en el período del escenario
    fechas_pb = sorted(df_pb["fecha"].dropna().unique().tolist())
    fechas_escenario = [f for f in fechas_pb if pb_desde <= f <= pb_hasta]

    if not fechas_escenario:
        raise ValueError(
            f"No hay datos de PB en el período del escenario "
            f"({pb_desde} a {pb_hasta}). "
            "Verifica que las fechas del escenario de PB coincidan con datos disponibles."
        )

    resultados_antes:   list[pd.DataFrame] = []
    resultados_despues: list[pd.DataFrame] = []

    for pb_fecha in fechas_escenario:
        # Fecha proxy: mismo mes/día pero en el año del inventario
        # → permite buscar "Mar-26" aunque el PB sea de marzo 2024
        inv_fecha = _proxy_fecha(pb_fecha, inv_anio)

        # Posición actual: inventario del mes equivalente, tipo de día de la fecha real
        df_pos_antes = calcular_posicion_neta(
            inventario, inv_fecha, fecha_tipo_dia=pb_fecha
        )

        # Precio de bolsa: fecha histórica real (escenario)
        df_pb_dia = df_pb.loc[df_pb["fecha"] == pb_fecha].copy()

        df_costo_antes = calcular_costo_bolsa(df_pos_antes, df_pb_dia)
        df_costo_antes["fecha"] = pb_fecha

        # Distribución horaria del contrato usando la fecha proxy
        # (días-del-mes correctos para el mes, sin importar el año del PB)
        delta_kwh = _distribuir_energia_horaria(
            energia_mensual_kwh if energia_mensual_kwh > 0 else None,
            inv_fecha,
            perfil_horario,
            inventario,
            bloques=bloques,
            perfil_pesos_24h=perfil_pesos_24h,
            perfil_excel_12x24=perfil_excel_12x24,
        )
        delta_values = delta_kwh.values  # ndarray (24,)

        # Aplicar a la posición
        df_pos_nueva = df_pos_antes.copy()
        if tipo == "compra":
            if tipo_mercado == "regulado":
                df_pos_nueva["compra_r_kwh"]  += delta_values
            elif tipo_mercado == "no_regulado":
                df_pos_nueva["compra_nr_kwh"] += delta_values
            else:  # ambos
                df_pos_nueva["compra_r_kwh"]  += delta_values * 0.5
                df_pos_nueva["compra_nr_kwh"] += delta_values * 0.5
        else:  # venta
            df_pos_nueva["venta_kwh"] += delta_values

        df_pos_nueva["posicion_neta_kwh"] = (
            df_pos_nueva["compra_r_kwh"]
            + df_pos_nueva["compra_nr_kwh"]
            - df_pos_nueva["venta_kwh"]
        )

        df_costo_despues = calcular_costo_bolsa(df_pos_nueva, df_pb_dia)
        df_costo_despues["fecha"] = pb_fecha

        resultados_antes.append(df_costo_antes)
        resultados_despues.append(df_costo_despues)

    if not resultados_antes:
        raise ValueError("No se generaron resultados de simulación.")

    df_antes   = pd.concat(resultados_antes,   ignore_index=True).sort_values(["fecha", "hora"])
    df_despues = pd.concat(resultados_despues, ignore_index=True).sort_values(["fecha", "hora"])

    # ── Resúmenes globales ────────────────────────────────────────────────────
    def _resumen(df: pd.DataFrame) -> dict[str, float | int]:
        idx_max = int(df["posicion_neta_kwh"].idxmax())
        idx_min = int(df["posicion_neta_kwh"].idxmin())
        return {
            "posicion_neta_total_mwh": round(float(df["posicion_neta_kwh"].sum() / 1_000), 2),
            "costo_bolsa_total_mcop":  round(float(df["costo_bolsa_cop"].sum()  / 1_000_000), 2),
            "hora_pico_compra": int(df.loc[idx_max, "hora"]),
            "hora_pico_venta":  int(df.loc[idx_min, "hora"]),
        }

    resumen_antes   = _resumen(df_antes)
    resumen_despues = _resumen(df_despues)

    # ── Semáforo ──────────────────────────────────────────────────────────────
    delta_costo_mcop = round(
        float(resumen_despues["costo_bolsa_total_mcop"])
        - float(resumen_antes["costo_bolsa_total_mcop"]),
        2,
    )
    mejora_costo = delta_costo_mcop < 0
    delta_pos = (
        float(resumen_despues["posicion_neta_total_mwh"])
        - float(resumen_antes["posicion_neta_total_mwh"])
    )
    mejora_posicion = (delta_pos < 0) if tipo == "venta" else mejora_costo

    if mejora_costo and mejora_posicion:
        recomendacion = "verde"
    elif mejora_costo or mejora_posicion:
        recomendacion = "amarillo"
    else:
        recomendacion = "rojo"

    # ── Perfil horario promedio (24h) ─────────────────────────────────────────
    pa = df_antes.groupby("hora")["posicion_neta_kwh"].mean().reset_index()
    pd_ = df_despues.groupby("hora")["posicion_neta_kwh"].mean().reset_index()
    pm = pd.merge(pa, pd_, on="hora", suffixes=("_a", "_d"))
    perfil_horario_lista = [
        {
            "hora": int(r["hora"]),
            "posicion_antes_mwh":   round(float(r["posicion_neta_kwh_a"] / 1_000), 2),
            "posicion_despues_mwh": round(float(r["posicion_neta_kwh_d"] / 1_000), 2),
        }
        for _, r in pm.iterrows()
    ]

    # ── Tabla por mes ─────────────────────────────────────────────────────────
    df_antes["mes_key"]   = pd.to_datetime(df_antes["fecha"]).dt.to_period("M")
    df_despues["mes_key"] = pd.to_datetime(df_despues["fecha"]).dt.to_period("M")

    agg_a = df_antes.groupby("mes_key").agg(
        pos_kwh=("posicion_neta_kwh", "sum"), costo_cop=("costo_bolsa_cop", "sum")
    )
    agg_d = df_despues.groupby("mes_key").agg(
        pos_kwh=("posicion_neta_kwh", "sum"), costo_cop=("costo_bolsa_cop", "sum")
    )
    mm = agg_a.join(agg_d, lsuffix="_a", rsuffix="_d")
    por_mes = []
    for pi, row in mm.iterrows():
        lbl = f"{MES_ABR[pi.month]}-{str(pi.year)[-2:]}"
        pa_ = float(row["pos_kwh_a"] / 1_000)
        pd__ = float(row["pos_kwh_d"] / 1_000)
        ca = float(row["costo_cop_a"] / 1_000_000)
        cd = float(row["costo_cop_d"] / 1_000_000)
        por_mes.append({
            "mes": lbl,
            "pos_actual_mwh":  round(pa_, 2),
            "pos_nueva_mwh":   round(pd__, 2),
            "diferencia_mwh":  round(pd__ - pa_, 2),
            "costo_actual_mcop": round(ca, 2),
            "costo_nuevo_mcop":  round(cd, 2),
            "ahorro_mcop":       round(ca - cd, 2),
        })

    print(
        f"[simulate] tipo={tipo} mercado={tipo_mercado} perfil={perfil_horario} "
        f"energia={energia_mensual_mwh}MWh/mes "
        f"pb={pb_desde}..{pb_hasta} contrato={contrato_inicio}..{contrato_fin} "
        f"fechas_escenario={len(fechas_escenario)} inv_anio={inv_anio} "
        f"recomendacion={recomendacion} delta={delta_costo_mcop:+.2f}M COP"
    )

    return {
        "resumen_antes":    resumen_antes,
        "resumen_despues":  resumen_despues,
        "recomendacion":    recomendacion,
        "delta_costo_mcop": delta_costo_mcop,
        "perfil_horario":   perfil_horario_lista,
        "por_mes":          por_mes,
    }
