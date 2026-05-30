"""
Motor de simulación de impacto de nuevos contratos sobre el portafolio Olibia.

CONCEPTO CLAVE — el período del contrato y el escenario de PB son independientes:

  Escenario de PB (pb_desde..pb_hasta):
    Solo se usa para calcular un PERFIL HORARIO PROMEDIO de 24 valores.
    Para cada hora h: perfil_pb[h] = promedio de todos los PB en esa hora
    durante el período del escenario (ENSO, histórico o proyectado).
    Este vector se aplica igual a TODOS los meses del contrato.

  Período del contrato (contrato_inicio..contrato_fin):
    Define los meses que aparecen en la tabla de resultados.
    Se itera día a día para obtener la posición real del inventario Olibia.
    La tabla por_mes SIEMPRE muestra meses del contrato, NUNCA del escenario PB.
"""

from __future__ import annotations

import calendar
from datetime import date, datetime, timedelta
from typing import Any

import pandas as pd

from portfolio_engine import (
    calcular_costo_bolsa,
    calcular_posicion_neta,
    _normalizar_serie_horaria,
    _periodo_desde_fecha,
)

# Abreviaturas de meses en español para etiquetar la tabla por_mes.
MES_ABR = {
    1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr",  5: "May", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic",
}

# Curva solar típica para Colombia (ecuatorial) — pesos normalizados, suma = 1.0.
# H1-H6 y H19-H24 sin irradiación; campana entre H7 y H18.
SOLAR_PESOS_24H: list[float] = [
    0,    0,    0,    0,    0,    0,      # H1-H6
    0.02, 0.06, 0.11, 0.15, 0.15, 0.14,  # H7-H12
    0.12, 0.10, 0.08, 0.05, 0.02, 0,     # H13-H18
    0,    0,    0,    0,    0,    0,      # H19-H24
]


def _dias_en_mes(anio: int, mes: int) -> int:
    """Retorna el número de días del mes indicado."""
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
    Devuelve una Serie indexada 1..24 con los kWh asignados por hora a ese día.

    Orden de prioridad de resolución del perfil:
      1. 'excel' / 'excel_custom'  → matriz 12×24 kWh/mes dividida entre días del mes
      2. 'bloques'                 → kWh/h calculados desde la lista de bloques
      3. 'solar'                   → curva solar colombiana (SOLAR_PESOS_24H)
      4. perfil_pesos_24h provisto → normaliza y escala con la energía diaria
      5. 'plano'                   → distribución uniforme en las 24 horas
      6. Forma del inventario      → perfil de ordinario / sábado / festivo de Olibia
      7. Fallback                  → plano con la energía disponible
    """
    fecha_obj = datetime.strptime(fecha_str, "%Y-%m-%d")
    dias = _dias_en_mes(fecha_obj.year, fecha_obj.month)
    kwh  = energia_mensual_kwh or 0.0

    # ── 1. Matriz Excel 12 × 24 (kWh/mes por hora) ───────────────────────────
    if perfil in ("excel", "excel_custom") and perfil_excel_12x24 is not None:
        mes_idx = fecha_obj.month - 1           # índice 0-11
        if mes_idx < len(perfil_excel_12x24):
            fila = list(perfil_excel_12x24[mes_idx])
            while len(fila) < 24:
                fila.append(0.0)
            return pd.Series(
                [float(v) / dias for v in fila[:24]],
                index=range(1, 25),
                dtype="float64",
            )
        # Mes fuera del rango de la matriz → distribución plana
        return pd.Series([kwh / dias / 24.0] * 24, index=range(1, 25), dtype="float64")

    # ── 2. Bloques horarios definidos por el usuario ──────────────────────────
    if perfil == "bloques" and bloques:
        pesos: list[float] = [0.0] * 24
        total_kwh = 0.0
        for b in bloques:
            h_ini        = int(b.get("hora_ini", 1))
            h_fin        = int(b.get("hora_fin", 24))
            kwh_bloque   = float(b.get("mwh_mes", 0)) * 1_000.0
            n_horas      = max(1, h_fin - h_ini + 1)
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
        total_pesos    = sum(SOLAR_PESOS_24H)
        valores = [
            (p / total_pesos) * energia_diaria if total_pesos > 0 else 0.0
            for p in SOLAR_PESOS_24H
        ]
        return pd.Series(valores, index=range(1, 25), dtype="float64")

    # ── 4. Pesos horarios personalizados (24 valores explícitos) ─────────────
    if perfil_pesos_24h is not None and kwh > 0:
        pesos_list = list(perfil_pesos_24h)
        while len(pesos_list) < 24:
            pesos_list.append(0.0)
        pesos_list    = pesos_list[:24]
        total          = sum(pesos_list)
        energia_diaria = kwh / dias
        if total > 0:
            return pd.Series(
                [(p / total) * energia_diaria for p in pesos_list],
                index=range(1, 25),
                dtype="float64",
            )
        return pd.Series([energia_diaria / 24.0] * 24, index=range(1, 25), dtype="float64")

    # ── 5. Distribución plana (24h iguales) ───────────────────────────────────
    if perfil == "plano":
        return pd.Series([kwh / dias / 24.0] * 24, index=range(1, 25), dtype="float64")

    # ── 6. Forma del inventario Olibia (ordinario / sábado / festivo) ─────────
    _PERFIL_A_CLAVE: dict[str, str] = {
        "ordinario": "compra_or",
        "sabado":    "compra_sa",
        "festivo":   "compra_fe",
    }
    clave = _PERFIL_A_CLAVE.get(perfil)
    if clave and clave in inventario and kwh > 0:
        periodo = _periodo_desde_fecha(fecha_obj)
        filas   = inventario[clave].loc[inventario[clave]["Periodo"] == periodo]
        if not filas.empty:
            try:
                shape       = _normalizar_serie_horaria(filas, clave)
                total_shape = float(shape.sum())
                if total_shape > 0:
                    energia_diaria = kwh / dias
                    return (shape / total_shape) * energia_diaria
            except ValueError:
                pass

    # ── 7. Fallback: distribución plana ──────────────────────────────────────
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
    ppp_resumen: dict | None = None,
) -> dict[str, Any]:
    """
    Simula el impacto de un nuevo contrato sobre el portafolio Olibia.

    Parámetros
    ----------
    inventario          : inventario de contratos Olibia (7 DataFrames).
    df_pb               : precios de bolsa del período del ESCENARIO (pb_desde..pb_hasta).
    tipo                : 'compra' o 'venta'.
    precio_cop_kwh      : precio del nuevo contrato en COP/kWh.
    pb_desde / pb_hasta : rango del escenario de PB (ENSO, histórico, proyectado).
    contrato_inicio / contrato_fin : vigencia del contrato simulado.
    tipo_mercado        : 'regulado', 'no_regulado' o 'ambos'.
    perfil_horario      : 'plano', 'bloques', 'solar' o 'excel'.
    energia_mensual_mwh : energía mensual en MWh/mes (para perfiles plano/solar).
    bloques             : lista de dicts {hora_ini, hora_fin, mwh_mes}.
    perfil_pesos_24h    : 24 pesos relativos normalizados.
    perfil_excel_12x24  : matriz 12×24 de kWh/mes por hora.

    Retorna
    -------
    Dict con claves:
      resumen_antes, resumen_despues, recomendacion, delta_costo_mcop,
      perfil_horario (24h promedio), por_mes (tabla meses del contrato),
      perfil_pb_escenario (24 valores del vector PB usado).
    """
    energia_mensual_kwh = (energia_mensual_mwh or 0.0) * 1_000.0

    # ═══════════════════════════════════════════════════════════════════════════
    # PASO 1 — Calcular el perfil horario promedio del escenario de PB
    # ═══════════════════════════════════════════════════════════════════════════
    # Para cada hora h (1..24): promedio de todos los precios de esa hora
    # en el período pb_desde..pb_hasta (puede ser ENSO, histórico o proyectado).
    # Este vector de 24 valores se aplica igual a TODOS los meses del contrato.
    # Solo se usa como precio de referencia; no genera meses en la tabla.
    perfil_pb: dict[int, float] = {}
    for h in range(1, 25):
        precios_hora = df_pb.loc[df_pb["hora"] == h, "precio_bolsa"]
        perfil_pb[h] = float(precios_hora.mean()) if len(precios_hora) > 0 else 0.0

    if all(v == 0.0 for v in perfil_pb.values()):
        raise ValueError(
            f"El DataFrame de PB no contiene datos válidos para el escenario "
            f"({pb_desde} a {pb_hasta}). Verifica que la consulta devolvió datos."
        )

    # DataFrame sintético con el perfil promedio para reutilizar calcular_costo_bolsa.
    # Tiene exactamente las columnas que espera esa función: 'hora' y 'precio_bolsa'.
    df_pb_perfil = pd.DataFrame({
        "hora":         list(range(1, 25)),
        "precio_bolsa": [perfil_pb[h] for h in range(1, 25)],
    })

    pb_promedio_global = sum(perfil_pb.values()) / 24.0

    # Spread del contrato respecto al escenario de PB:
    #   COMPRA: PB - precio  → positivo = conveniente (compro más barato que bolsa)
    #   VENTA:  precio - PB  → positivo = conveniente (vendo más caro que bolsa)
    if tipo == "compra":
        spread_cop_kwh = pb_promedio_global - precio_cop_kwh
    else:
        spread_cop_kwh = precio_cop_kwh - pb_promedio_global

    # ═══════════════════════════════════════════════════════════════════════════
    # PASO 2 — Iterar sobre cada día real del período del CONTRATO
    # ═══════════════════════════════════════════════════════════════════════════
    # Los meses que aparecen en la tabla son SIEMPRE los del contrato.
    # NUNCA se muestran meses del escenario PB en la tabla de resultados.
    desde_contrato = datetime.strptime(contrato_inicio, "%Y-%m-%d").date()
    hasta_contrato = datetime.strptime(contrato_fin,    "%Y-%m-%d").date()

    resultados_antes:   list[pd.DataFrame] = []
    resultados_despues: list[pd.DataFrame] = []
    # Acumula la energía del nuevo contrato por mes (clave "YYYY-MM") en kWh
    delta_energia_mensual: dict[str, float] = {}

    dia_actual = desde_contrato
    while dia_actual <= hasta_contrato:
        fecha_str = dia_actual.isoformat()

        try:
            # ── 2a. Posición neta ACTUAL del inventario para este día ─────────
            # calcular_posicion_neta determina automáticamente el tipo_dia
            # (ordinario / sábado / domingo / festivo) y aplica:
            #   compra_r[h] = TO[mes][h] + inventario_tipo_dia[mes][h]
            #   compra_nr[h] = NR[mes][h]
            #   venta[h] = venta[mes][h]
            #   posicion_neta[h] = compra_r + compra_nr - venta
            df_pos_antes = calcular_posicion_neta(inventario, fecha_str)

            # ── 2b. Costo de bolsa ANTES del nuevo contrato ───────────────────
            # costo_bolsa[h] = posicion_neta[h] × perfil_pb[h]
            # Se usa el mismo perfil_pb para todos los días del contrato.
            df_costo_antes = calcular_costo_bolsa(df_pos_antes, df_pb_perfil.copy())
            df_costo_antes["fecha"] = fecha_str

            # ── 2c. Distribuir la energía del nuevo contrato en 24 horas ─────
            delta_kwh    = _distribuir_energia_horaria(
                energia_mensual_kwh if energia_mensual_kwh > 0 else None,
                fecha_str,
                perfil_horario,
                inventario,
                bloques=bloques,
                perfil_pesos_24h=perfil_pesos_24h,
                perfil_excel_12x24=perfil_excel_12x24,
            )
            delta_values = delta_kwh.values   # ndarray (24,)

            # Acumular energía diaria del contrato por mes (para ahorro real)
            mes_key_dia = fecha_str[:7]   # "YYYY-MM"
            delta_energia_mensual[mes_key_dia] = (
                delta_energia_mensual.get(mes_key_dia, 0.0)
                + float(delta_kwh.sum())
            )

            # ── 2d. Aplicar el nuevo contrato a la posición ───────────────────
            df_pos_nueva = df_pos_antes.copy()
            if tipo == "compra":
                if tipo_mercado == "regulado":
                    df_pos_nueva["compra_r_kwh"]  += delta_values
                elif tipo_mercado == "no_regulado":
                    df_pos_nueva["compra_nr_kwh"] += delta_values
                else:   # ambos: mitad regulado, mitad no regulado
                    df_pos_nueva["compra_r_kwh"]  += delta_values * 0.5
                    df_pos_nueva["compra_nr_kwh"] += delta_values * 0.5
            else:   # venta
                df_pos_nueva["venta_kwh"] += delta_values

            # Recalcular posición neta con el nuevo contrato incluido.
            df_pos_nueva["posicion_neta_kwh"] = (
                df_pos_nueva["compra_r_kwh"]
                + df_pos_nueva["compra_nr_kwh"]
                - df_pos_nueva["venta_kwh"]
            )

            # ── 2e. Costo de bolsa DESPUÉS, con el mismo perfil del escenario ─
            df_costo_despues = calcular_costo_bolsa(df_pos_nueva, df_pb_perfil.copy())
            df_costo_despues["fecha"] = fecha_str

            resultados_antes.append(df_costo_antes)
            resultados_despues.append(df_costo_despues)

        except (ValueError, KeyError):
            # Fecha fuera del período cubierto por los inventarios → se omite.
            pass

        dia_actual += timedelta(days=1)

    if not resultados_antes:
        raise ValueError(
            "No se generaron resultados de simulación. Verifica que el período "
            f"del contrato ({contrato_inicio} a {contrato_fin}) esté cubierto "
            "por los inventarios de Olibia."
        )

    df_antes = (
        pd.concat(resultados_antes,   ignore_index=True)
        .sort_values(["fecha", "hora"])
        .reset_index(drop=True)
    )
    df_despues = (
        pd.concat(resultados_despues, ignore_index=True)
        .sort_values(["fecha", "hora"])
        .reset_index(drop=True)
    )

    # ═══════════════════════════════════════════════════════════════════════════
    # PASO 3 — Resúmenes, semáforo y tablas de salida
    # ═══════════════════════════════════════════════════════════════════════════

    def _resumen(df: pd.DataFrame) -> dict[str, float | int]:
        """Calcula métricas agregadas de posición y costo."""
        idx_max = int(df["posicion_neta_kwh"].idxmax())
        idx_min = int(df["posicion_neta_kwh"].idxmin())
        return {
            "posicion_neta_total_mwh": round(float(df["posicion_neta_kwh"].sum()) / 1_000, 2),
            "costo_bolsa_total_mcop":  round(float(df["costo_bolsa_cop"].sum())   / 1_000_000, 2),
            "hora_pico_compra": int(df.loc[idx_max, "hora"]),
            "hora_pico_venta":  int(df.loc[idx_min, "hora"]),
        }

    resumen_antes   = _resumen(df_antes)
    resumen_despues = _resumen(df_despues)

    # ── Semáforo de recomendación (basado en spread vs PB del escenario) ─────
    # spread_cop_kwh ya fue calculado en el PASO 1:
    #   COMPRA: PB_promedio - precio_contrato  (> 0 → compro más barato que bolsa)
    #   VENTA:  precio_contrato - PB_promedio  (> 0 → vendo más caro que bolsa)
    # Umbrales: VERDE > +20 COP/kWh | AMARILLO entre -20 y +20 | ROJO < -20 COP/kWh
    if spread_cop_kwh > 20:
        recomendacion = "verde"
    elif spread_cop_kwh >= -20:
        recomendacion = "amarillo"
    else:
        recomendacion = "rojo"

    # delta_costo_mcop se mantiene como métrica informativa (diferencia de costos de bolsa)
    delta_costo_mcop = round(
        float(resumen_despues["costo_bolsa_total_mcop"])
        - float(resumen_antes["costo_bolsa_total_mcop"]),
        2,
    )

    # ── Perfil horario promedio (24h) para el gráfico comparativo ────────────
    agg_pa  = df_antes.groupby("hora")["posicion_neta_kwh"].mean().reset_index()
    agg_pd  = df_despues.groupby("hora")["posicion_neta_kwh"].mean().reset_index()
    pm      = pd.merge(agg_pa, agg_pd, on="hora", suffixes=("_a", "_d"))
    perfil_horario_lista = [
        {
            "hora":                 int(r["hora"]),
            "posicion_antes_mwh":   round(float(r["posicion_neta_kwh_a"]) / 1_000, 2),
            "posicion_despues_mwh": round(float(r["posicion_neta_kwh_d"]) / 1_000, 2),
        }
        for _, r in pm.iterrows()
    ]

    # ── Tabla por mes del período del CONTRATO ────────────────────────────────
    # Los meses aquí son SIEMPRE los del contrato (ej. Jul-26 a Dic-26).
    # NUNCA aparecen meses del escenario PB (ej. 2015 o 2023).
    df_antes["mes_key"]   = pd.to_datetime(df_antes["fecha"]).dt.to_period("M")
    df_despues["mes_key"] = pd.to_datetime(df_despues["fecha"]).dt.to_period("M")

    agg_a = df_antes.groupby("mes_key").agg(
        pos_kwh   = ("posicion_neta_kwh", "sum"),
        costo_cop = ("costo_bolsa_cop",   "sum"),
    )
    agg_d = df_despues.groupby("mes_key").agg(
        pos_kwh   = ("posicion_neta_kwh", "sum"),
        costo_cop = ("costo_bolsa_cop",   "sum"),
    )
    mm = agg_a.join(agg_d, lsuffix="_a", rsuffix="_d")

    por_mes = []
    for periodo_idx, row in mm.iterrows():
        etiqueta      = f"{MES_ABR[periodo_idx.month]}-{str(periodo_idx.year)[-2:]}"
        mes_str       = f"{periodo_idx.year}-{periodo_idx.month:02d}"   # "YYYY-MM"
        pos_antes_mwh = float(row["pos_kwh_a"])   / 1_000
        pos_nueva_mwh = float(row["pos_kwh_d"])   / 1_000
        costo_antes   = float(row["costo_cop_a"]) / 1_000_000
        costo_nuevo   = float(row["costo_cop_d"]) / 1_000_000

        # Ahorro real = spread × energía del contrato en ese mes
        # (spread > 0 → ahorro positivo; spread < 0 → costo adicional)
        energia_mes_kwh = delta_energia_mensual.get(mes_str, 0.0)
        ahorro_mes = round(spread_cop_kwh * energia_mes_kwh / 1_000_000, 2)

        por_mes.append({
            "mes":               etiqueta,
            "pos_actual_mwh":    round(pos_antes_mwh,                   2),
            "pos_nueva_mwh":     round(pos_nueva_mwh,                   2),
            "diferencia_mwh":    round(pos_nueva_mwh - pos_antes_mwh,   2),
            "costo_actual_mcop": round(costo_antes,                     2),
            "costo_nuevo_mcop":  round(costo_nuevo,                     2),
            "ahorro_mcop":       ahorro_mes,
        })

    # Texto del punto de equilibrio para mostrar en el frontend
    accion      = "compra" if tipo == "compra" else "venta"
    signo_str   = f"+{spread_cop_kwh:.1f}" if spread_cop_kwh >= 0 else f"{spread_cop_kwh:.1f}"
    breakeven = {
        "precio_contrato_cop_kwh":   round(precio_cop_kwh,      2),
        "pb_promedio_escenario_cop_kwh": round(pb_promedio_global, 2),
        "spread_cop_kwh":            round(spread_cop_kwh,       2),
        "descripcion": (
            f"El contrato es indiferente cuando PB = precio_contrato = "
            f"{precio_cop_kwh:.1f} COP/kWh. "
            f"Con PB promedio del escenario de {pb_promedio_global:.1f} COP/kWh, "
            f"el spread de {accion} es {signo_str} COP/kWh."
        ),
    }

    # ── Log informativo ───────────────────────────────────────────────────────
    print(
        f"[simulate] tipo={tipo} mercado={tipo_mercado} perfil={perfil_horario} "
        f"energia={energia_mensual_mwh}MWh/mes "
        f"escenario_pb={pb_desde}..{pb_hasta} "
        f"contrato={contrato_inicio}..{contrato_fin} "
        f"dias_procesados={df_antes['fecha'].nunique()} "
        f"pb_prom_global={pb_promedio_global:.1f} COP/kWh "
        f"spread={spread_cop_kwh:+.1f} COP/kWh "
        f"recomendacion={recomendacion} delta={delta_costo_mcop:+.2f}M COP"
    )

    return {
        # Métricas agregadas antes y después del contrato
        "resumen_antes":   resumen_antes,
        "resumen_despues": resumen_despues,
        # Semáforo y variación económica
        "recomendacion":    recomendacion,
        "delta_costo_mcop": delta_costo_mcop,
        # Spread del contrato vs PB del escenario
        "spread_cop_kwh":   round(spread_cop_kwh, 2),
        # Punto de equilibrio y descripción textual
        "breakeven": breakeven,
        # Gráfico horario comparativo (promedio de todos los días del contrato)
        "perfil_horario": perfil_horario_lista,
        # Tabla mensual — siempre meses del contrato, nunca del escenario PB
        "por_mes": por_mes,
        # Vector PB utilizado (informativo, para mostrar en el frontend)
        "perfil_pb_escenario": [
            {"hora": h, "pb_promedio": round(perfil_pb[h], 2)}
            for h in range(1, 25)
        ],
        # PPP real de contratos PC del período simulado (None si no disponible)
        "ppp_contratos": ppp_resumen,
    }
