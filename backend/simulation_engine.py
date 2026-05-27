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


def _dias_en_mes(anio: int, mes: int) -> int:
    """Cantidad de días del mes."""
    return calendar.monthrange(anio, mes)[1]


def _distribuir_energia_horaria(
    energia_mensual_kwh: float,
    fecha_str: str,
    perfil: str,
    inventario: dict[str, pd.DataFrame],
) -> pd.Series:
    """
    Distribuye la energía mensual (kWh) en 24 valores horarios para un día específico.

    - perfil='plano':                    distribución uniforme entre las 24 horas.
    - perfil='ordinario'|'sabado'|'festivo': usa la forma del inventario correspondiente,
      normalizada para que la energía diaria sea energia_mensual_kwh / dias_en_mes.

    Retorna Serie indexada 1..24 con kWh para ese día.
    """
    fecha_obj = datetime.strptime(fecha_str, "%Y-%m-%d")
    dias = _dias_en_mes(fecha_obj.year, fecha_obj.month)
    energia_diaria_kwh = energia_mensual_kwh / dias

    # ── Perfil plano ─────────────────────────────────────────────────────────
    if perfil == "plano":
        return pd.Series(
            [energia_diaria_kwh / 24.0] * 24,
            index=range(1, 25),
            dtype="float64",
        )

    # ── Perfil con forma (usa inventario como curva de distribución) ─────────
    clave = _PERFIL_A_CLAVE.get(perfil)
    if clave and clave in inventario:
        periodo = _periodo_desde_fecha(fecha_obj)
        df_inv = inventario[clave]
        filas = df_inv.loc[df_inv["Periodo"] == periodo]
        if not filas.empty:
            try:
                shape = _normalizar_serie_horaria(filas, clave)
                total_shape = shape.sum()
                if total_shape > 0:
                    # Normalizar shape (suma = 1) y escalar a energía diaria
                    return (shape / total_shape) * energia_diaria_kwh
            except ValueError:
                pass  # Si falla, cae al fallback plano

    # Fallback: plano
    return pd.Series(
        [energia_diaria_kwh / 24.0] * 24,
        index=range(1, 25),
        dtype="float64",
    )


def simular_contrato(
    inventario: dict[str, pd.DataFrame],
    df_pb: pd.DataFrame,
    tipo: str,
    energia_mensual_mwh: float,
    precio_cop_kwh: float,
    fecha_inicio: str,
    fecha_fin: str,
    tipo_mercado: str,
    perfil_horario: str,
) -> dict[str, Any]:
    """
    Simula el impacto de un nuevo contrato sobre el portafolio actual.

    Para cada fecha con datos de PB dentro de la vigencia del contrato:
      1. Calcula posición neta y costo de bolsa actuales (sin contrato).
      2. Distribuye la energía del contrato en 24 horas según el perfil.
      3. Suma/resta la energía a la posición según tipo (compra/venta) y mercado.
      4. Recalcula posición neta y costo con el nuevo contrato incluido.

    Args:
        inventario:          Dict de DataFrames cargado por cargar_inventario().
        df_pb:               DataFrame con columnas [fecha, hora, precio_bolsa].
        tipo:                'compra' | 'venta'
        energia_mensual_mwh: Energía del contrato en MWh/mes.
        precio_cop_kwh:      Precio del contrato en COP/kWh (informativo; el costo
                             de bolsa usa el precio_bolsa real, no el contrato).
        fecha_inicio:        Inicio de vigencia YYYY-MM-DD.
        fecha_fin:           Fin de vigencia YYYY-MM-DD.
        tipo_mercado:        'regulado' | 'no_regulado' | 'ambos'
        perfil_horario:      'plano' | 'ordinario' | 'sabado' | 'festivo'

    Returns:
        dict con resumen_antes, resumen_despues, recomendacion, delta_costo_mcop,
        perfil_horario (lista 24 horas promedio) y por_mes (tabla mensual).
    """
    energia_mensual_kwh = energia_mensual_mwh * 1_000.0

    # Solo fechas dentro de la vigencia del contrato
    fechas_pb = sorted(df_pb["fecha"].dropna().unique().tolist())
    fechas_vigencia = [f for f in fechas_pb if fecha_inicio <= f <= fecha_fin]

    if not fechas_vigencia:
        raise ValueError(
            "No hay datos de precio de bolsa en el rango de vigencia del contrato "
            f"({fecha_inicio} a {fecha_fin})."
        )

    resultados_antes:   list[pd.DataFrame] = []
    resultados_despues: list[pd.DataFrame] = []

    for fecha in fechas_vigencia:
        # ── Posición actual (sin nuevo contrato) ─────────────────────────────
        df_pos_antes = calcular_posicion_neta(inventario, fecha)
        df_pb_dia = df_pb.loc[df_pb["fecha"] == fecha].copy()

        df_costo_antes = calcular_costo_bolsa(df_pos_antes, df_pb_dia)
        df_costo_antes["fecha"] = fecha

        # ── Energía del nuevo contrato para este día (24 valores en kWh) ─────
        delta_kwh = _distribuir_energia_horaria(
            energia_mensual_kwh, fecha, perfil_horario, inventario
        )
        delta_values = delta_kwh.values  # ndarray (24,)

        # ── Aplicar el contrato a la posición ────────────────────────────────
        df_pos_nueva = df_pos_antes.copy()

        if tipo == "compra":
            if tipo_mercado == "regulado":
                df_pos_nueva["compra_r_kwh"]  = df_pos_nueva["compra_r_kwh"]  + delta_values
            elif tipo_mercado == "no_regulado":
                df_pos_nueva["compra_nr_kwh"] = df_pos_nueva["compra_nr_kwh"] + delta_values
            else:  # ambos — split 50/50
                df_pos_nueva["compra_r_kwh"]  = df_pos_nueva["compra_r_kwh"]  + delta_values * 0.5
                df_pos_nueva["compra_nr_kwh"] = df_pos_nueva["compra_nr_kwh"] + delta_values * 0.5
        else:  # venta
            df_pos_nueva["venta_kwh"] = df_pos_nueva["venta_kwh"] + delta_values

        df_pos_nueva["posicion_neta_kwh"] = (
            df_pos_nueva["compra_r_kwh"]
            + df_pos_nueva["compra_nr_kwh"]
            - df_pos_nueva["venta_kwh"]
        )

        df_costo_despues = calcular_costo_bolsa(df_pos_nueva, df_pb_dia)
        df_costo_despues["fecha"] = fecha

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
    # Costo positivo = pago neto a bolsa; negativo = ingreso neto de bolsa.
    # Reducir el costo (delta < 0) siempre es favorable.
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
    # Venta: reduce posición (más vendido) → mejora si bolsa alta
    # Compra: sube posición (más comprado) → evaluamos solo por costo
    if tipo == "venta":
        mejora_posicion = delta_pos < 0
    else:
        mejora_posicion = mejora_costo

    if mejora_costo and mejora_posicion:
        recomendacion = "verde"
    elif mejora_costo or mejora_posicion:
        recomendacion = "amarillo"
    else:
        recomendacion = "rojo"

    # ── Perfil horario promedio (media aritmética por hora sobre el período) ──
    perfil_a_df = df_antes.groupby("hora")["posicion_neta_kwh"].mean().reset_index()
    perfil_d_df = df_despues.groupby("hora")["posicion_neta_kwh"].mean().reset_index()
    perfil_merge = pd.merge(
        perfil_a_df, perfil_d_df, on="hora", suffixes=("_antes", "_despues")
    )
    perfil_horario_lista = [
        {
            "hora": int(row["hora"]),
            "posicion_antes_mwh":   round(float(row["posicion_neta_kwh_antes"]   / 1_000), 2),
            "posicion_despues_mwh": round(float(row["posicion_neta_kwh_despues"] / 1_000), 2),
        }
        for _, row in perfil_merge.iterrows()
    ]

    # ── Tabla resumen por mes ─────────────────────────────────────────────────
    df_antes["mes_key"]   = pd.to_datetime(df_antes["fecha"]).dt.to_period("M")
    df_despues["mes_key"] = pd.to_datetime(df_despues["fecha"]).dt.to_period("M")

    agg_a = df_antes.groupby("mes_key").agg(
        pos_kwh=("posicion_neta_kwh", "sum"),
        costo_cop=("costo_bolsa_cop", "sum"),
    )
    agg_d = df_despues.groupby("mes_key").agg(
        pos_kwh=("posicion_neta_kwh", "sum"),
        costo_cop=("costo_bolsa_cop", "sum"),
    )
    merge_mes = agg_a.join(agg_d, lsuffix="_antes", rsuffix="_despues")

    por_mes = []
    for periodo_idx, row in merge_mes.iterrows():
        mes_label = f"{MES_ABR[periodo_idx.month]}-{str(periodo_idx.year)[-2:]}"
        pos_a   = float(row["pos_kwh_antes"]   / 1_000)
        pos_d   = float(row["pos_kwh_despues"] / 1_000)
        costo_a = float(row["costo_cop_antes"]   / 1_000_000)
        costo_d = float(row["costo_cop_despues"] / 1_000_000)
        por_mes.append({
            "mes":             mes_label,
            "pos_actual_mwh":  round(pos_a, 2),
            "pos_nueva_mwh":   round(pos_d, 2),
            "diferencia_mwh":  round(pos_d - pos_a, 2),
            "costo_actual_mcop": round(costo_a, 2),
            "costo_nuevo_mcop":  round(costo_d, 2),
            "ahorro_mcop":       round(costo_a - costo_d, 2),
        })

    print(
        f"[simulate] tipo={tipo} energia={energia_mensual_mwh}MWh/mes "
        f"mercado={tipo_mercado} perfil={perfil_horario} "
        f"fechas={len(fechas_vigencia)} recomendacion={recomendacion} "
        f"delta_costo={delta_costo_mcop:+.2f}M COP"
    )

    return {
        "resumen_antes":   resumen_antes,
        "resumen_despues": resumen_despues,
        "recomendacion":   recomendacion,
        "delta_costo_mcop": delta_costo_mcop,
        "perfil_horario":  perfil_horario_lista,
        "por_mes":         por_mes,
    }
