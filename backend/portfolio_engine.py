"""
Motor de portafolio para inventarios de contratos de Olibia.
"""

from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd


# Columnas horarias esperadas en los archivos de inventario.
COLUMNAS_HORARIAS = [f"H{i}" for i in range(1, 25)]

# Mapa de abreviaturas de meses en espanol para formar "Ene-26", "Feb-26", etc.
MES_ABR = {
    1: "Ene",
    2: "Feb",
    3: "Mar",
    4: "Abr",
    5: "May",
    6: "Jun",
    7: "Jul",
    8: "Ago",
    9: "Sep",
    10: "Oct",
    11: "Nov",
    12: "Dic",
}


def _periodo_desde_fecha(fecha: datetime) -> str:
    """Convierte una fecha al formato de inventario: Ene-26, Feb-26, etc."""
    return f"{MES_ABR[fecha.month]}-{str(fecha.year)[-2:]}"


def _validar_columnas_inventario(df: pd.DataFrame, nombre_archivo: str) -> None:
    """Valida que existan Periodo y H1..H24 en el inventario."""
    columnas_requeridas = {"Periodo", *COLUMNAS_HORARIAS}
    faltantes = columnas_requeridas - set(df.columns)
    if faltantes:
        raise ValueError(
            f"El archivo '{nombre_archivo}' no tiene columnas requeridas. "
            f"Faltantes: {sorted(faltantes)}"
        )


def _normalizar_serie_horaria(df_mes: pd.DataFrame, etiqueta: str) -> pd.Series:
    """
    Convierte la fila mensual del inventario en una serie numérica de 24 horas.

    Retorna una Serie indexada por hora (1..24).
    """
    if df_mes.empty:
        raise ValueError(
            f"No hay fila para el periodo solicitado en el inventario '{etiqueta}'."
        )

    if len(df_mes) > 1:
        raise ValueError(
            f"El inventario '{etiqueta}' tiene {len(df_mes)} filas para el mismo periodo."
        )

    fila = df_mes.iloc[0]
    valores = pd.to_numeric(fila[COLUMNAS_HORARIAS], errors="coerce")

    if valores.isna().any():
        horas_invalidas = [
            int(columna[1:])
            for columna in valores.index[valores.isna()].tolist()
        ]
        raise ValueError(
            f"El inventario '{etiqueta}' tiene valores no numericos en horas: "
            f"{horas_invalidas}"
        )

    serie = pd.Series(valores.values, index=range(1, 25), dtype="float64")
    return serie


def _siguiente_lunes(fecha_base: date) -> date:
    """
    Aplica Ley Emiliani: mueve algunas festividades al lunes siguiente.
    """
    dias_hasta_lunes = (7 - fecha_base.weekday()) % 7
    return fecha_base + timedelta(days=dias_hasta_lunes)


def _fecha_pascua(anio: int) -> date:
    """
    Calcula el domingo de Pascua (algoritmo de Meeus para calendario gregoriano).
    """
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
    """
    Retorna el conjunto de fechas festivas en Colombia para un anio.
    """
    pascua = _fecha_pascua(anio)

    # Festivos fijos
    festivos: set[date] = {
        date(anio, 1, 1),   # Ano Nuevo
        date(anio, 5, 1),   # Dia del Trabajo
        date(anio, 7, 20),  # Independencia
        date(anio, 8, 7),   # Batalla de Boyaca
        date(anio, 12, 8),  # Inmaculada Concepcion
        date(anio, 12, 25),  # Navidad
    }

    # Festivos por Ley Emiliani (traslado al lunes)
    emiliani = [
        date(anio, 1, 6),   # Reyes Magos
        date(anio, 3, 19),  # San Jose
        date(anio, 6, 29),  # San Pedro y San Pablo
        date(anio, 8, 15),  # Asuncion
        date(anio, 10, 12),  # Dia de la Raza
        date(anio, 11, 1),  # Todos los Santos
        date(anio, 11, 11),  # Independencia de Cartagena
    ]
    festivos.update(_siguiente_lunes(fecha_base) for fecha_base in emiliani)

    # Festivos moviles respecto a Pascua
    festivos.add(pascua - timedelta(days=3))   # Jueves Santo
    festivos.add(pascua - timedelta(days=2))   # Viernes Santo
    festivos.add(_siguiente_lunes(pascua + timedelta(days=43)))  # Ascension
    festivos.add(_siguiente_lunes(pascua + timedelta(days=64)))  # Corpus Christi
    festivos.add(_siguiente_lunes(pascua + timedelta(days=71)))  # Sagrado Corazon

    return festivos


def _tipo_dia(fecha_obj: datetime) -> str:
    """
    Determina el tipo de dia para seleccionar inventario R:
    ordinario, sabado, domingo o festivo.
    """
    fecha_dia = fecha_obj.date()

    if fecha_dia in _festivos_colombia(fecha_obj.year):
        return "festivo"

    dia_semana = fecha_obj.weekday()  # lunes=0 ... domingo=6
    if dia_semana == 5:
        return "sabado"
    if dia_semana == 6:
        return "domingo"
    return "ordinario"


def cargar_inventario(ruta_base: str | Path) -> dict[str, pd.DataFrame]:
    """
    Carga los 7 archivos de inventario de contratos desde una carpeta base.

    Retorna un diccionario con llaves:
    compra_or, compra_to, compra_sa, compra_do, compra_fe, compra_nr, venta
    """
    base = Path(ruta_base)
    archivos = {
        "compra_or": "Inventario_Total_Ene-26_Dic-26 compra Ordinario R.xlsx",
        "compra_to": "Inventario_Total_Ene-26_Dic-26 compra R TO.xlsx",
        "compra_sa": "Inventario_Total_Ene-26_Dic-26 compra Sabado R.xlsx",
        "compra_do": "Inventario_Total_Ene-26_Dic-26 compra Domingo R.xlsx",
        "compra_fe": "Inventario_Total_Ene-26_Dic-26 compra Festivo R.xlsx",
        "compra_nr": "Inventario_Total_Ene-26_Dic-26 compra NR.xlsx",
        "venta": "Inventario_Total_Ene-26_Dic-26 venta.xlsx",
    }

    inventario: dict[str, pd.DataFrame] = {}
    for clave, nombre_archivo in archivos.items():
        ruta_archivo = base / nombre_archivo
        if not ruta_archivo.exists():
            raise FileNotFoundError(f"No se encontro el archivo: {ruta_archivo}")

        df = pd.read_excel(ruta_archivo)
        _validar_columnas_inventario(df, nombre_archivo)
        inventario[clave] = df

    return inventario


def calcular_posicion_neta(
    inventario: dict[str, pd.DataFrame],
    fecha: str,
    fecha_tipo_dia: str | None = None,
) -> pd.DataFrame:
    """
    Calcula la posicion neta horaria para una fecha puntual.

    Args:
        inventario: Diccionario de inventarios de contratos.
        fecha: Fecha para buscar el período en el inventario (YYYY-MM-DD).
               Determina el mes/año del inventario a usar (ej. "2026-03-15" → "Mar-26").
        fecha_tipo_dia: Fecha opcional para clasificar el tipo de día (ordinario/sábado/
               festivo). Si se omite se usa la misma que `fecha`. Permite pasar una fecha
               histórica real para obtener una mezcla realista de días laborables/festivos
               mientras se busca en el inventario del año vigente.

    Formula por hora:
        compra_r = TO + inventario R segun tipo de dia
        posicion_neta = compra_r + compra_nr - venta
    """
    fecha_obj = datetime.strptime(fecha, "%Y-%m-%d")
    periodo = _periodo_desde_fecha(fecha_obj)

    # Clasificar el día usando fecha_tipo_dia si se provee (p.ej. fecha histórica de PB)
    tipo_dia_obj = (
        datetime.strptime(fecha_tipo_dia, "%Y-%m-%d") if fecha_tipo_dia else fecha_obj
    )
    tipo_dia = _tipo_dia(tipo_dia_obj)

    if tipo_dia == "ordinario":
        clave_tipo_dia = "compra_or"
    elif tipo_dia == "sabado":
        clave_tipo_dia = "compra_sa"
    elif tipo_dia == "domingo":
        clave_tipo_dia = "compra_do"
    else:
        clave_tipo_dia = "compra_fe"

    claves_requeridas = {
        "compra_to",
        "compra_nr",
        "venta",
        clave_tipo_dia,
    }
    faltantes = claves_requeridas - set(inventario.keys())
    if faltantes:
        raise ValueError(f"Faltan llaves en inventario: {sorted(faltantes)}")

    # Filtrar por mes (Periodo) en cada inventario relevante.
    serie_to = _normalizar_serie_horaria(
        inventario["compra_to"].loc[inventario["compra_to"]["Periodo"] == periodo],
        "compra_to",
    )
    serie_r_tipo = _normalizar_serie_horaria(
        inventario[clave_tipo_dia].loc[inventario[clave_tipo_dia]["Periodo"] == periodo],
        clave_tipo_dia,
    )
    serie_nr = _normalizar_serie_horaria(
        inventario["compra_nr"].loc[inventario["compra_nr"]["Periodo"] == periodo],
        "compra_nr",
    )
    serie_venta = _normalizar_serie_horaria(
        inventario["venta"].loc[inventario["venta"]["Periodo"] == periodo],
        "venta",
    )

    compra_r = serie_to + serie_r_tipo
    posicion_neta = compra_r + serie_nr - serie_venta

    df_posicion = pd.DataFrame(
        {
            "hora": range(1, 25),
            "compra_r_kwh": compra_r.values,
            "compra_nr_kwh": serie_nr.values,
            "venta_kwh": serie_venta.values,
            "posicion_neta_kwh": posicion_neta.values,
        }
    )

    return df_posicion


def calcular_posicion_periodo(
    inventario: dict[str, pd.DataFrame],
    fecha_inicio: str,
    fecha_fin: str,
) -> pd.DataFrame:
    """
    Calcula la posición neta horaria para un rango completo de fechas.

    Para cada día real del período:
      1. Determina el tipo de día (ordinario / sábado / domingo / festivo).
      2. Obtiene la serie horaria de compra_r = TO[mes][hora] + tipo_día[mes][hora]
         donde tipo_día es el inventario R que corresponde al día.
      3. compra_nr = NR[mes][hora]
      4. venta = venta[mes][hora]
      5. posicion_neta = compra_r + compra_nr - venta

    Los días fuera del período cubierto por los inventarios se omiten
    silenciosamente (sin lanzar excepción).

    Parámetros
    ----------
    inventario   : diccionario devuelto por cargar_inventario().
    fecha_inicio : primer día del rango en formato YYYY-MM-DD (inclusive).
    fecha_fin    : último día del rango en formato YYYY-MM-DD (inclusive).

    Retorna
    -------
    DataFrame con columnas:
        fecha, hora, tipo_dia, compra_r_kwh, compra_nr_kwh,
        venta_kwh, posicion_neta_kwh
    Ordenado por (fecha, hora). Vacío → lanza ValueError.
    """
    desde = datetime.strptime(fecha_inicio, "%Y-%m-%d").date()
    hasta = datetime.strptime(fecha_fin, "%Y-%m-%d").date()

    fragmentos: list[pd.DataFrame] = []
    dia_actual = desde

    while dia_actual <= hasta:
        fecha_str = dia_actual.isoformat()

        # Calcular tipo_dia para anotarlo en el DataFrame resultante.
        tipo_dia = _tipo_dia(datetime(dia_actual.year, dia_actual.month, dia_actual.day))

        try:
            # calcular_posicion_neta ya aplica TO + tipo_día internamente.
            df_dia = calcular_posicion_neta(inventario, fecha_str)

            # Agregar columnas de contexto al inicio del DataFrame.
            df_dia.insert(0, "tipo_dia", tipo_dia)
            df_dia.insert(0, "fecha", fecha_str)

            fragmentos.append(df_dia)

        except (ValueError, KeyError):
            # Fecha fuera del período del inventario (ej. 2027 con Excel 2026).
            pass

        dia_actual += timedelta(days=1)

    if not fragmentos:
        raise ValueError(
            f"No hay datos de inventario para el rango "
            f"{fecha_inicio} a {fecha_fin}. "
            "Verifica que los archivos de inventario cubran ese período."
        )

    df_resultado = (
        pd.concat(fragmentos, ignore_index=True)
        .sort_values(["fecha", "hora"])
        .reset_index(drop=True)
    )
    return df_resultado


def calcular_costo_bolsa(df_posicion: pd.DataFrame, df_pb_dia: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula el costo/ingreso de bolsa por hora:
        costo_bolsa_cop = posicion_neta_kwh * precio_bolsa
    """
    columnas_posicion = {"hora", "posicion_neta_kwh"}
    faltantes_posicion = columnas_posicion - set(df_posicion.columns)
    if faltantes_posicion:
        raise ValueError(
            f"df_posicion no tiene columnas requeridas: {sorted(faltantes_posicion)}"
        )

    columnas_pb = {"hora", "precio_bolsa"}
    faltantes_pb = columnas_pb - set(df_pb_dia.columns)
    if faltantes_pb:
        raise ValueError(
            f"df_pb_dia no tiene columnas requeridas: {sorted(faltantes_pb)}"
        )

    df_merge = pd.merge(
        df_posicion.copy(),
        df_pb_dia[["hora", "precio_bolsa"]].copy(),
        on="hora",
        how="inner",
    )

    if len(df_merge) != 24:
        raise ValueError(
            "El cruce entre posicion y PB no genero 24 horas. "
            f"Filas obtenidas: {len(df_merge)}"
        )

    df_merge["costo_bolsa_cop"] = (
        pd.to_numeric(df_merge["posicion_neta_kwh"], errors="coerce")
        * pd.to_numeric(df_merge["precio_bolsa"], errors="coerce")
    )

    if df_merge["costo_bolsa_cop"].isna().any():
        raise ValueError("Se detectaron valores no numericos al calcular costo_bolsa_cop.")

    return df_merge


def resumen_portafolio(df_costo: pd.DataFrame, periodo: str) -> dict[str, float | int | str]:
    """
    Genera el resumen agregado del portafolio para un periodo.
    """
    columnas_requeridas = {
        "hora",
        "compra_r_kwh",
        "compra_nr_kwh",
        "venta_kwh",
        "posicion_neta_kwh",
        "costo_bolsa_cop",
    }
    faltantes = columnas_requeridas - set(df_costo.columns)
    if faltantes:
        raise ValueError(
            f"df_costo no tiene columnas requeridas para resumen: {sorted(faltantes)}"
        )

    df = df_costo.copy()

    hora_pico_compra = int(df.loc[df["posicion_neta_kwh"].idxmax(), "hora"])
    hora_pico_venta = int(df.loc[df["posicion_neta_kwh"].idxmin(), "hora"])

    resumen = {
        "periodo": periodo,
        "total_compra_r_kwh": float(df["compra_r_kwh"].sum()),
        "total_compra_nr_kwh": float(df["compra_nr_kwh"].sum()),
        "total_venta_kwh": float(df["venta_kwh"].sum()),
        "posicion_neta_total_kwh": float(df["posicion_neta_kwh"].sum()),
        "costo_bolsa_total_cop": float(df["costo_bolsa_cop"].sum()),
        "hora_pico_compra": hora_pico_compra,
        "hora_pico_venta": hora_pico_venta,
    }

    return resumen
