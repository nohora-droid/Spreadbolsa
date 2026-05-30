"""
Motor de portafolio para inventarios de contratos de Olibia.
"""

from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

from olibia_loader import cargar_posicion_olibia


# Columnas horarias esperadas en los archivos de inventario (legado Excel).
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


def cargar_inventario(ruta_base: str | Path) -> None:
    """
    OBSOLETA: La carga de inventario ahora se realiza vía API de Olibia Energy.

    Esta función se conserva únicamente por compatibilidad de firma; devuelve None.
    En su lugar usa cargar_posicion_olibia() de olibia_loader.
    """
    print(
        "[portfolio] AVISO: cargar_inventario() está obsoleta. "
        "La posición ahora se obtiene desde la API de Olibia vía olibia_loader."
    )
    return None


def calcular_posicion_neta(
    inventario: pd.DataFrame | dict | None,
    fecha: str,
    fecha_tipo_dia: str | None = None,
) -> pd.DataFrame:
    """
    Calcula la posicion neta horaria para una fecha puntual.

    Cuando inventario es un DataFrame proveniente de la API de Olibia
    (columnas: date, hour, compra_r_kwh, compra_nr_kwh, venta_kwh,
    posicion_neta_kwh), filtra las filas de esa fecha, renombra
    'hour' → 'hora' y devuelve las 24 filas con columnas estándar.

    Cuando inventario es None o un tipo no soportado, devuelve una
    tabla de 24 filas con ceros (comportamiento seguro para simulaciones).

    Args:
        inventario: DataFrame de posición (API Olibia), dict legado o None.
        fecha: Fecha en formato YYYY-MM-DD.
        fecha_tipo_dia: Ignorado; se conserva por compatibilidad de firma.
    """
    _COLUMNAS_POS = [
        "hora", "compra_r_kwh", "compra_nr_kwh", "venta_kwh", "posicion_neta_kwh"
    ]

    def _ceros() -> pd.DataFrame:
        return pd.DataFrame(
            {
                "hora":             range(1, 25),
                "compra_r_kwh":     0.0,
                "compra_nr_kwh":    0.0,
                "venta_kwh":        0.0,
                "posicion_neta_kwh": 0.0,
            }
        )

    # Camino API: inventario es DataFrame de cargar_posicion_olibia
    if isinstance(inventario, pd.DataFrame):
        # Acepta columna 'date' (API) o 'fecha' (ya renombrada)
        col_fecha = "date" if "date" in inventario.columns else "fecha"
        df_dia = inventario[inventario[col_fecha] == fecha].copy()

        if df_dia.empty:
            return _ceros()

        # Normalizar nombre de la columna de hora
        if "hour" in df_dia.columns:
            df_dia = df_dia.rename(columns={"hour": "hora"})

        # Garantizar que todas las columnas esperadas existan
        for col in _COLUMNAS_POS:
            if col not in df_dia.columns:
                df_dia[col] = 0.0

        return (
            df_dia[_COLUMNAS_POS]
            .sort_values("hora")
            .reset_index(drop=True)
        )

    # Camino legado / sin datos: devolver ceros
    return _ceros()


def calcular_posicion_periodo(
    inventario: pd.DataFrame | dict | None,
    fecha_inicio: str,
    fecha_fin: str,
) -> pd.DataFrame:
    """
    Calcula la posición neta horaria para un rango completo de fechas.

    Obtiene los datos directamente desde la API de Olibia Energy usando
    cargar_posicion_olibia() e inyecta la columna tipo_dia por cada fecha.

    El parámetro inventario se ignora; se conserva por compatibilidad de firma
    con los endpoints existentes que pasaban el dict de Excel.

    Parámetros
    ----------
    inventario   : ignorado (puede ser None, dict o DataFrame).
    fecha_inicio : primer día del rango en formato YYYY-MM-DD (inclusive).
    fecha_fin    : último día del rango en formato YYYY-MM-DD (inclusive).

    Retorna
    -------
    DataFrame con columnas:
        fecha, hora, tipo_dia, compra_r_kwh, compra_nr_kwh,
        venta_kwh, posicion_neta_kwh
    Ordenado por (fecha, hora). Vacío → lanza ValueError.
    """
    # Cargar posición desde la API de Olibia (paralelo, 8 workers)
    df = cargar_posicion_olibia(fecha_inicio, fecha_fin)

    if df.empty:
        raise ValueError(
            f"No hay datos de posición para el rango {fecha_inicio} a {fecha_fin}. "
            "Verifica la conexión con la API de Olibia y que los contratos "
            "cubran ese período."
        )

    # Renombrar columnas al esquema estándar
    df = df.rename(columns={"date": "fecha", "hour": "hora"})

    # Calcular tipo_dia por fecha única y mapearlo al DataFrame
    mapa_tipo_dia: dict[str, str] = {}
    for fecha_str in df["fecha"].unique():
        try:
            fecha_obj = datetime.strptime(fecha_str, "%Y-%m-%d")
            mapa_tipo_dia[fecha_str] = _tipo_dia(fecha_obj)
        except ValueError:
            mapa_tipo_dia[fecha_str] = "ordinario"

    df.insert(2, "tipo_dia", df["fecha"].map(mapa_tipo_dia))

    # Garantizar orden de columnas y que todas existan
    columnas_orden = [
        "fecha", "hora", "tipo_dia",
        "compra_r_kwh", "compra_nr_kwh", "venta_kwh", "posicion_neta_kwh",
    ]
    for col in columnas_orden:
        if col not in df.columns:
            df[col] = 0.0

    return (
        df[columnas_orden]
        .sort_values(["fecha", "hora"])
        .reset_index(drop=True)
    )


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
