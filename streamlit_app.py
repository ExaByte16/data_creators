# -*- coding: utf-8 -*-
"""
Oxynia Balance General – Streamlit App

This app reproduces exactly the transformation steps from the provided notebook,
adding a UI to upload the Excel export from SIIGO, input MES/ESTADO/AÑO/CENTRO DE COSTOS,
and download the two resulting Excel files:
- datos_balance_general.xlsx
- datos_estado_resultados.xlsx

Notes:
- The header is assumed to be on row 8 (0-indexed skiprows=7), identical to the notebook.
- Column names must match the SIIGO export used in the notebook.
- "Nombre tercero" is used to fill the TERCERO column; if it is missing, it will be created empty.
"""

from __future__ import annotations

from io import BytesIO, StringIO
from zipfile import ZipFile
from typing import Tuple

import pandas as pd
import streamlit as st

from report_generator import generate_informe, make_branding


APP_TITLE = "Oxynia Balance General"


# 1. Definición de los Mapeos según el PUC Colombia (idénticos al notebook)
CLASE_MAP = {
    "1": "ACTIVO",
    "2": "PASIVO",
    "3": "PATRIMONIO",
    "4": "INGRESOS",
    "5": "GASTOS",
    "6": "COSTOS DE VENTAS",
    "7": "COSTOS DE PRODUCCIÓN O DE OPERACIÓN",
    "8": "CUENTAS DE ORDEN DEUDORAS",
    "9": "CUENTAS DE ORDEN ACREEDORAS",
}

SUBGRUPO_MAP = {
    "11": "DISPONIBLE",
    "12": "INVERSIONES",
    "13": "DEUDORES",
    "14": "INVENTARIOS",
    "15": "PROPIEDADES, PLANTA Y EQUIPO",
    "16": "INTANGIBLES",
    "17": "DIFERIDOS",
    "18": "OTROS ACTIVOS",
    "19": "VALORIZACIONES",
    "21": "OBLIGACIONES FINANCIERAS",
    "22": "PROVEEDORES",
    "23": "CUENTAS POR PAGAR",
    "24": "IMPUESTOS, GRAVÁMENES Y TASAS",
    "25": "OBLIGACIONES LABORALES",
    "26": "PASIVOS ESTIMADOS Y PROVISIONES",
    "27": "DIFERIDOS",
    "28": "OTROS PASIVOS",
    "29": "BONOS Y PAPELES COMERCIALES",
    "31": "CAPITAL SOCIAL",
    "32": "SUPERÁVIT DE CAPITAL",
    "33": "RESERVAS",
    "34": "REVALORIZACIÓN DEL PATRIMONIO",
    "35": "DIVIDENDOS O PARTICIPACIONES DECRETADOS EN ACCIONES",
    "36": "RESULTADOS DEL EJERCICIO",
    "37": "RESULTADOS DE EJERCICIOS ANTERIORES",
    "38": "SUPERÁVIT POR VALORIZACIONES",
    "41": "OPERACIONALES",
    "42": "NO OPERACIONALES",
    "47": "AJUSTES POR INFLACIÓN",
    "51": "OPERACIONALES DE ADMINISTRACIÓN",
    "52": "OPERACIONALES DE VENTAS",
    "53": "NO OPERACIONALES",
    "54": "IMPUESTO DE RENTA Y COMPLEMENTARIOS",
    "59": "GANANCIAS Y PÉRDIDAS",
    "61": "COSTO DE VENTAS Y DE PRESTACIÓN DE SERVICIOS",
    "62": "COMPRAS",
    "71": "MATERIA PRIMA",
    "72": "MANO DE OBRA DIRECTA",
    "73": "COSTOS INDIRECTOS",
    "74": "CONTRATOS DE SERVICIOS",
    "81": "DERECHOS CONTINGENTES",
    "82": "DEUDORAS FISCALES",
    "83": "DEUDORAS DE CONTROL",
    "91": "RESPONSABILIDADES CONTINGENTES",
    "92": "ACREEDORAS FISCALES",
    "93": "ACREEDORAS DE CONTROL",
}


def determinar_grupo(row: pd.Series) -> str:
    """Función idéntica al notebook para determinar GRUPO (Corriente / No Corriente)."""
    clase_cod = row["Código str"][0]
    subgrupo_cod = row["Código str"][:2]

    if clase_cod == "1":  # ACTIVO
        if subgrupo_cod in ["11", "12", "13", "14"]:
            return "ACTIVO CORRIENTE"
        elif subgrupo_cod in ["15", "16", "17", "18", "19"]:
            return "ACTIVO NO CORRIENTE"
    elif clase_cod == "2":  # PASIVO
        if subgrupo_cod in ["21", "22", "23", "24", "25", "26", "28"]:
            return "PASIVO CORRIENTE"
        elif subgrupo_cod in ["27", "29"]:
            return "PASIVO NO CORRIENTE"
    elif clase_cod == "3":  # PATRIMONIO
        return "PATRIMONIO"  # El patrimonio por definición es no corriente
    elif clase_cod in ["4", "5", "6", "7"]:
        return "CUENTA DE RESULTADOS"
    elif clase_cod in ["8", "9"]:
        return "CUENTA DE ORDEN"
    return "NO CLASIFICADO"  # Para códigos que no encajen


def _filter_leaf_accounts(df: pd.DataFrame) -> pd.DataFrame:
    """
    Filtra el DataFrame para mantener solo las cuentas "hoja" (de mayor detalle).
    
    Una cuenta es "hoja" si su código NO es prefijo de otra cuenta en el dataset.
    Esto evita incluir subtotales junto con sus componentes detallados.
    
    Por ejemplo, si tenemos cuentas 1635, 163505, 163510:
    - 1635 NO es hoja (es prefijo de 163505 y 163510)
    - 163505 y 163510 SÍ son hojas
    
    Args:
        df: DataFrame con columna 'Código cuenta contable'
        
    Returns:
        DataFrame filtrado con solo cuentas hoja
    """
    if df.empty:
        return df
    
    # Obtener todos los códigos únicos como strings limpios
    df = df.copy()
    codigos = (
        df["Código cuenta contable"]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .unique()
    )
    codigos_set = set(codigos)
    
    def es_cuenta_hoja(codigo: str) -> bool:
        """Verifica si un código es hoja (no es prefijo de otro código)."""
        codigo_limpio = str(codigo).replace(".0", "")
        for otro_codigo in codigos_set:
            if otro_codigo != codigo_limpio and otro_codigo.startswith(codigo_limpio):
                return False
        return True
    
    # Crear máscara para filtrar solo cuentas hoja
    df["_codigo_limpio"] = (
        df["Código cuenta contable"]
        .astype(str)
        .str.replace(".0", "", regex=False)
    )
    df["_es_hoja"] = df["_codigo_limpio"].apply(es_cuenta_hoja)
    
    # Filtrar y limpiar columnas temporales
    df_filtered = df[df["_es_hoja"]].copy()
    df_filtered = df_filtered.drop(columns=["_codigo_limpio", "_es_hoja"])
    
    return df_filtered


def calcular_valor_cuenta(row: pd.Series) -> float:
    """
    Calcula el valor de una cuenta según su tipo (primer dígito del código).
    
    Fórmulas según el PUC Colombia:
    - Cuentas 1, 2, 3 (Balance): Saldo Final
    - Cuenta 4 (Ingresos): Crédito - Débito
    - Cuentas 5, 6, 7 (Gastos/Costos): Débito - Crédito
    - Cuentas 8, 9 (Orden): Saldo Final
    
    Args:
        row: Fila del DataFrame con columnas 'Código str', 'Saldo final',
             'Movimiento débito', 'Movimiento crédito'
             
    Returns:
        Valor calculado según el tipo de cuenta
    """
    primer_digito = row["Código str"][0] if row["Código str"] else "0"
    saldo_final = row.get("Saldo final", 0) or 0
    debito = row.get("Movimiento débito", 0) or 0
    credito = row.get("Movimiento crédito", 0) or 0
    
    if primer_digito in ["1", "2", "3", "8", "9"]:
        # Cuentas de Balance y Orden: usar Saldo Final
        return saldo_final
    elif primer_digito == "4":
        # Ingresos: Crédito - Débito (ingresos se registran en el crédito)
        return credito - debito
    elif primer_digito in ["5", "6", "7"]:
        # Gastos y Costos: Débito - Crédito (gastos se registran en el débito)
        return debito - credito
    else:
        # Fallback: usar Saldo Final
        return saldo_final


def read_siigo_excel(file_obj) -> pd.DataFrame:
    """Read the SIIGO Excel export assuming header on row 8 (header=7)."""
    excel = pd.ExcelFile(file_obj)
    df = excel.parse(header=7)
    return df


def read_datax_excel(file_obj) -> pd.DataFrame:
    """Read the Datax Balance de Comprobación Excel. Header is on row 1 (header=0)."""
    df = pd.read_excel(file_obj, header=0)
    return df


def normalize_datax_to_siigo(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normaliza un DataFrame de Datax al formato interno equivalente al de Siigo,
    permitiendo reutilizar todo el pipeline de procesamiento existente.

    Estructura Datax:
        cuenta, nom_cuenta, cod_benf, nom_benf, saldo_ini, debitos, creditos,
        saldo, nivel, tipo, orden, ter_benf

    Reglas de conversión:
    - tipo='T': subtotales intermedios → se descartan
    - tipo='D', cod_benf=NaN, ter_benf=True  → fila resumen por cuenta
                                                → Transaccional='No'
    - tipo='D', cod_benf≠NaN                 → fila detalle por tercero
                                                → Transaccional='Sí'
    - tipo='D', cod_benf=NaN, ter_benf=False → cuenta hoja sin terceros;
                                                se marca Transaccional='Sí'
                                                para que la lógica de cuentas
                                                huérfanas la detecte y la incluya
                                                correctamente en ambos modos.
    """
    df = df.copy()

    # 1. Eliminar filas de metadatos (cuenta vacía o NaN)
    df = df[df["cuenta"].notna() & (df["cuenta"].astype(str).str.strip() != "")].copy()

    # 2. Eliminar filas de subtotales intermedios (tipo='T')
    df = df[df["tipo"] == "D"].copy()

    # 3. Crear columna Transaccional equivalente a la de Siigo
    def _get_transaccional(row) -> str:
        if pd.notna(row["cod_benf"]):
            return "Sí"  # detalle por tercero
        elif bool(row.get("ter_benf", False)):
            return "No"  # fila resumen de cuenta con terceros
        else:
            return "Sí"  # cuenta hoja sin terceros → tratada como huérfana

    df["Transaccional"] = df.apply(_get_transaccional, axis=1)

    # 4. Renombrar columnas al formato interno del pipeline
    df = df.rename(columns={
        "cuenta":     "Código cuenta contable",
        "nom_cuenta": "Nombre cuenta contable",
        "nom_benf":   "Nombre tercero",
        "saldo_ini":  "Saldo inicial",
        "debitos":    "Movimiento débito",
        "creditos":   "Movimiento crédito",
        "saldo":      "Saldo final",
    })

    # 5. Agregar columnas dummy requeridas por el pipeline (se descartan más adelante)
    df["Sucursal"] = ""
    df["Identificación"] = ""

    return df


def validate_required_columns(df: pd.DataFrame) -> None:
    """Ensure the DataFrame contains required columns used by the notebook."""
    required_columns = {
        "Transaccional",
        "Saldo inicial",
        "Movimiento débito",
        "Movimiento crédito",
        "Sucursal",
        "Identificación",
        "Código cuenta contable",
        "Nombre cuenta contable",
        "Saldo final",
    }
    missing = [c for c in required_columns if c not in df.columns]
    if missing:
        raise ValueError(
            f"Faltan columnas requeridas en el archivo Excel: {', '.join(missing)}"
        )


def ensure_nombre_tercero(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure column 'Nombre tercero' exists; if not, create it empty to match notebook usage."""
    if "Nombre tercero" not in df.columns:
        df = df.copy()
        df["Nombre tercero"] = ""
    return df


def create_excel_download_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    """Convert a DataFrame to an in-memory Excel file and return its bytes."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output.getvalue()


def _find_orphan_transactional_accounts(df: pd.DataFrame) -> pd.DataFrame:
    """
    Encuentra cuentas "huérfanas": subcuentas que solo existen como Transaccional="Sí"
    y no tienen un resumen Transaccional="No".

    En SIIGO, la mayoría de subcuentas tienen esta estructura:
        510506 (Subcuenta, Trans=No)   ← subtotal/resumen
          └── 51050601 (Auxiliar, Trans=Sí)  ← detalle por tercero

    Pero algunas subcuentas solo tienen entradas Transaccional="Sí" sin resumen "No":
        510518 (Subcuenta, Trans=Sí)   ← sin resumen, solo detalle
          └── (sin hijos)

    Cuando filtramos por Transaccional="No", estas cuentas se pierden.
    Esta función las detecta y agrega sus valores para incluirlas.

    Criterio para considerar una cuenta "huérfana":
      1. Solo tiene filas Transaccional="Sí" (no tiene resumen "No")
      2. Es cuenta hoja (ningún otro código empieza con ella)
      3. Ningún ancestro (código padre) es hoja en el subconjunto Trans="No"
         → Esto evita incluir Auxiliares (8 dígitos) cuyo Subcuenta padre
           (6 dígitos) ya tiene un resumen "No" que cubre ese detalle.

    Args:
        df: DataFrame original completo (sin filtrar).

    Returns:
        DataFrame con las cuentas huérfanas agregadas (sumadas por cuenta).
    """
    # Limpiar códigos como strings
    df = df.copy()
    df["_cod_str"] = df["Código cuenta contable"].apply(
        lambda x: str(int(float(x))) if pd.notna(x) else ""
    )

    # Códigos que tienen al menos una fila Transaccional="No"
    codes_with_no = set(
        df.loc[df["Transaccional"] == "No", "_cod_str"].unique()
    )
    codes_with_no.discard("")

    # Códigos que SOLO tienen Transaccional="Sí" (nunca tienen "No")
    df_si = df[df["Transaccional"] == "Sí"].copy()
    codes_only_si = set(df_si["_cod_str"].unique()) - codes_with_no
    codes_only_si.discard("")

    if not codes_only_si:
        return pd.DataFrame()

    # Todos los códigos en el dataset completo
    all_codes = set(df["_cod_str"].unique())
    all_codes.discard("")

    # Paso 1: Identificar cuáles códigos Trans="No" son hojas dentro del subset "No"
    # (no tienen otro código "No" más largo que empiece por ellos)
    codes_no_leaves = set()
    for code in codes_with_no:
        is_leaf_in_no = not any(
            other != code and other.startswith(code) for other in codes_with_no
        )
        if is_leaf_in_no:
            codes_no_leaves.add(code)

    # Paso 2: Filtrar cuentas solo-Sí que son hojas en el dataset completo
    # Y que NO tienen un ancestro que sea hoja en el subset "No"
    orphan_leaf_codes = set()
    for code in codes_only_si:
        # ¿Es hoja en el dataset completo? (ningún otro código empieza con ella)
        is_leaf = not any(
            other != code and other.startswith(code) for other in all_codes
        )
        if not is_leaf:
            continue

        # ¿Tiene algún ancestro que sea hoja en el subset "No"?
        # Ej: 11050501 empieza con 110505 (que es hoja en "No") → NO es huérfana
        # Ej: 510518 NO empieza con ninguna hoja "No" → SÍ es huérfana
        has_no_leaf_ancestor = any(
            code.startswith(no_leaf) and code != no_leaf
            for no_leaf in codes_no_leaves
        )
        if has_no_leaf_ancestor:
            continue

        orphan_leaf_codes.add(code)

    if not orphan_leaf_codes:
        return pd.DataFrame()

    # Filtrar solo las filas Sí de esas cuentas huérfanas y agregar
    df_orphans = df_si[df_si["_cod_str"].isin(orphan_leaf_codes)].copy()

    df_aggregated = (
        df_orphans.groupby(
            ["Código cuenta contable", "Nombre cuenta contable"],
            as_index=False,
            dropna=False,
        )
        .agg({
            "Saldo final": "sum",
            "Movimiento débito": "sum",
            "Movimiento crédito": "sum",
        })
    )

    return df_aggregated


def process_dataframe(
    df: pd.DataFrame, mes: str, estado: str, anio: str, centro_costos: str, desglosar_por_tercero: bool = False
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, str]:
    """
    Reproduce el pipeline del notebook y retorna:
    - datos_balance_general (con columnas MES/AÑO/CENTRO DE COSTOS al inicio)
    - datos_estado_resultados (con columnas MES/ESTADO/AÑO/CENTRO DE COSTOS al inicio)
    - df_final (tabla intermedia mostrada por el notebook)
    - df_filtered_head (primeras filas del DataFrame filtrado por Transaccional=="No")
    - df_unique_accounts_head (primeras filas después de drop_duplicates)
    - df_unique_accounts_info_text (salida de df_unique_accounts.info())
    """
    # Display df.head() in UI outside this function.

    # Filter Transaccional:
    # - Cuando desglosamos por tercero: usar registros detallados (Transaccional="Sí")
    #   para evitar duplicados con subtotales
    # - Cuando NO desglosamos: usar subtotales ya agregados (Transaccional="No")
    #   PLUS cuentas huérfanas que solo existen como "Sí" (ver _find_orphan_transactional_accounts)
    if desglosar_por_tercero:
        df_filtered = df[df["Transaccional"] == "Sí"].copy()  # Solo registros detallados
    else:
        df_filtered = df[df["Transaccional"] == "No"].copy()  # Solo subtotales

        # Incluir cuentas hoja que solo existen como Transaccional="Sí"
        # (SIIGO no genera resumen "No" para ellas porque ya son el nivel más bajo)
        df_orphans = _find_orphan_transactional_accounts(df)
        if not df_orphans.empty:
            # Asegurar que las columnas coincidan antes de concatenar
            for col in df_filtered.columns:
                if col not in df_orphans.columns:
                    df_orphans[col] = pd.NA
            df_orphans = df_orphans[[c for c in df_filtered.columns if c in df_orphans.columns]]
            df_filtered = pd.concat([df_filtered, df_orphans], ignore_index=True)

    # Remove specified columns (keep Movimiento débito/crédito for value calculation)
    columns_to_drop = [
        "Saldo inicial",
        "Sucursal",
        "Identificación",
    ]
    df_filtered = df_filtered.drop(columns=columns_to_drop)

    # Agregar por cuenta (y opcionalmente por tercero) para conservar o resumir el detalle de terceros
    if desglosar_por_tercero:
        group_keys = [
            "Código cuenta contable",
            "Nombre cuenta contable",
            "Nombre tercero",
        ]
    else:
        group_keys = [
            "Código cuenta contable",
            "Nombre cuenta contable",
        ]

    # Agrupar y sumar valores, manteniendo las columnas clave
    # Incluimos débito y crédito para calcular el valor según tipo de cuenta
    df_unique_accounts = (
        df_filtered.groupby(group_keys, as_index=False, dropna=False)
        .agg({
            "Saldo final": "sum",
            "Movimiento débito": "sum",
            "Movimiento crédito": "sum",
        })
    )

    # Si no se desglosa por tercero, garantizamos la columna para el flujo posterior
    if "Nombre tercero" not in df_unique_accounts.columns:
        df_unique_accounts["Nombre tercero"] = ""

    # Filtrar filas con Código cuenta contable nulo o inválido antes de procesar
    df_unique_accounts = df_unique_accounts[
        df_unique_accounts["Código cuenta contable"].notna()
    ].copy()
    
    # Capture info() output
    info_buffer = StringIO()
    df_unique_accounts.info(buf=info_buffer)
    df_unique_accounts_info_text = info_buffer.getvalue()

    # Convertir el código a string para manipulación (manejar posibles decimales)
    df_unique_accounts = df_unique_accounts.copy()
    df_unique_accounts["Código str"] = (
        df_unique_accounts["Código cuenta contable"]
        .apply(lambda x: str(int(float(x))) if pd.notna(x) else "")
    )

    # Crear las columnas usando los mapeos y la función
    df_unique_accounts["CLASE"] = (
        df_unique_accounts["Código str"].str[0].map(CLASE_MAP).fillna("NO DEFINIDA")
    )
    df_unique_accounts["SUBGRUPO"] = (
        df_unique_accounts["Código str"].str[:2].map(SUBGRUPO_MAP).fillna("NO DEFINIDO")
    )
    df_unique_accounts["GRUPO"] = df_unique_accounts.apply(determinar_grupo, axis=1)

    # Crear las columnas 'Cuenta' y 'Valor' y 'TERCERO'
    df_unique_accounts["CUENTA"] = (
        df_unique_accounts["Código str"]
        + " - "
        + df_unique_accounts["Nombre cuenta contable"].fillna("")
    )
    # Calcular VALOR según tipo de cuenta (1-3: saldo final, 4: Cr-Db, 5-7: Db-Cr)
    df_unique_accounts["VALOR"] = df_unique_accounts.apply(calcular_valor_cuenta, axis=1)
    df_unique_accounts["TERCERO"] = df_unique_accounts["Nombre tercero"]

    # DataFrame final con el formato deseado
    columnas_finales = [
        "CLASE",
        "GRUPO",
        "SUBGRUPO",
        "CUENTA",
        "TERCERO",
        "VALOR",
        "Código cuenta contable",
    ]
    df_final = df_unique_accounts[columnas_finales].copy()

    # Filtrar el DataFrame
    df_final = df_final[
        (df_final["GRUPO"] != "NO CLASIFICADO") & (df_final["SUBGRUPO"] != "NO DEFINIDO")
    ].copy()

    # Filtrar para obtener solo las cuentas de mayor detalle (hojas del árbol contable)
    # Una cuenta es "hoja" si no es prefijo de otra cuenta en el dataset
    # Esto evita duplicar valores de subtotales con sus componentes
    df_final = _filter_leaf_accounts(df_final)

    # Convert to get first digit and split
    df_final = df_final.copy()
    df_final["First Digit"] = (
        df_final["Código cuenta contable"].astype(int).astype(str).str[0]
    )

    # Drop 'Código cuenta contable'
    df_final_wo_code = df_final.drop(columns=["Código cuenta contable"]).copy()

    # Split into Balance General (1,2,3) and Estado de Resultados (4-9)
    datos_balance_general = df_final_wo_code[
        df_final_wo_code["First Digit"].isin(["1", "2", "3"])
    ].copy()
    datos_estado_resultados = df_final_wo_code[
        df_final_wo_code["First Digit"].isin(["4", "5", "6", "7", "8", "9"])
    ].copy()

    # Drop the temporary 'First Digit'
    datos_balance_general = datos_balance_general.drop(columns=["First Digit"])
    datos_estado_resultados = datos_estado_resultados.drop(columns=["First Digit"])

    # Insert new columns at the beginning in the exact order inserted by the notebook
    # Balance General: MES, AÑO, CENTRO DE COSTOS
    datos_balance_general = datos_balance_general.copy()
    datos_balance_general.insert(0, "CENTRO DE COSTOS", centro_costos)
    datos_balance_general.insert(0, "AÑO", anio)
    datos_balance_general.insert(0, "MES", mes)

    # Estado de Resultados: MES, ESTADO, AÑO, CENTRO DE COSTOS
    datos_estado_resultados = datos_estado_resultados.copy()
    datos_estado_resultados.insert(0, "CENTRO DE COSTOS", centro_costos)
    datos_estado_resultados.insert(0, "AÑO", anio)
    datos_estado_resultados.insert(0, "ESTADO", estado)
    datos_estado_resultados.insert(0, "MES", mes)

    # Prepare heads for UI display to match notebook's displays
    df_filtered_head = df_filtered.head().copy()
    df_unique_accounts_head = df_unique_accounts.head().copy()

    return (
        datos_balance_general,
        datos_estado_resultados,
        df_final,
        df_filtered_head,
        df_unique_accounts_head,
        df_unique_accounts_info_text,
    )


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.caption(
        "Este app procesa y agrupa un export de un periodo de operación mensual "
        "(SIIGO o Datax), generando Balance General y Estado de Resultados en Excel."
    )

    # Selector de formato
    formato = st.radio(
        "Formato del archivo de entrada",
        options=["Siigo", "Datax"],
        horizontal=True,
    )

    with st.expander("Instrucciones", expanded=True):
        if formato == "Siigo":
            st.markdown("- Cargue el archivo Excel exportado de **SIIGO**.")
            st.markdown(
                "- El encabezado del archivo debe estar en la fila 8 (header=7)."
            )
        else:
            st.markdown(
                "- Cargue el archivo Excel de **Balance de Comprobación de Datax**."
            )
            st.markdown(
                "- El encabezado debe estar en la primera fila "
                "(columnas: `cuenta`, `nom_cuenta`, `cod_benf`, `nom_benf`, "
                "`saldo_ini`, `debitos`, `creditos`, `saldo`, `nivel`, `tipo`, "
                "`orden`, `ter_benf`)."
            )
        st.markdown("- Ingrese MES, ESTADO, AÑO y CENTRO DE COSTOS.")
        st.markdown(
            "- Descargue un ZIP con `datos_balance_general.xlsx` y "
            "`datos_estado_resultados.xlsx`."
        )
        st.markdown(
            "- Nota: Si no se puede determinar 'Nombre tercero', la columna "
            "TERCERO se llenará en blanco."
        )

    label_uploader = (
        "Suba el archivo Excel de SIIGO"
        if formato == "Siigo"
        else "Suba el archivo Excel de Datax (Balance de Comprobación)"
    )
    uploaded_file = st.file_uploader(label_uploader, type=["xlsx"])

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        mes = st.text_input("MES", value="")
    with col2:
        estado = st.text_input("ESTADO", value="")
    with col3:
        anio = st.text_input("AÑO", value="")
    with col4:
        centro_costos = st.text_input("CENTRO DE COSTOS", value="")

    desglosar_por_tercero = st.checkbox(
        "Desglosar por TERCERO en los reportes finales", value=True
    )

    st.divider()
    st.subheader("Branding e Informe Completo")
    st.caption(
        "Opcional: configure datos de branding para generar el INFORME completo "
        "con formato profesional, gráficas, colores y firmas."
    )

    generar_informe_flag = st.checkbox("Generar INFORME completo (Excel con formato)", value=False)

    branding_config = {}
    if generar_informe_flag:
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            branding_empresa = st.text_input("Nombre de la empresa", value="")
            branding_nit = st.text_input("NIT", value="")
            branding_rep_legal = st.text_input("Representante Legal", value="")
            branding_rep_cc = st.text_input("C.C. Representante", value="")
        with col_b2:
            branding_contador = st.text_input("Contador(a)", value="")
            branding_contador_tp = st.text_input("T.P. Contador(a)", value="")
            branding_contador_cc = st.text_input("C.C. Contador(a)", value="")
            branding_logo = st.file_uploader("Logo (PNG/JPG)", type=["png", "jpg", "jpeg"])

        logo_bytes = branding_logo.read() if branding_logo else None
        branding_config = make_branding(
            empresa=branding_empresa,
            nit=branding_nit,
            representante_legal=branding_rep_legal,
            representante_cc=branding_rep_cc,
            contador=branding_contador,
            contador_tp=branding_contador_tp,
            contador_cc=branding_contador_cc,
            logo_bytes=logo_bytes,
        )

    st.divider()
    process_clicked = st.button("Procesar y Generar Archivos")

    if process_clicked:
        if uploaded_file is None:
            st.error("Debe subir un archivo Excel (.xlsx).")
            st.stop()
        if not mes or not estado or not anio or not centro_costos:
            st.error("Por favor complete MES, ESTADO, AÑO y CENTRO DE COSTOS.")
            st.stop()

        try:
            # Leer y normalizar según el formato seleccionado
            if formato == "Siigo":
                raw_df = read_siigo_excel(uploaded_file)
            else:
                raw_df = read_datax_excel(uploaded_file)
                raw_df = normalize_datax_to_siigo(raw_df)

            # Vista previa (primeras filas del archivo ya procesado)
            st.subheader("Vista previa (primeras filas)")
            st.dataframe(raw_df.head())

            # Validar columnas y garantizar 'Nombre tercero'
            validate_required_columns(raw_df)
            raw_df = ensure_nombre_tercero(raw_df)

            # Procesar el DataFrame (lógica del notebook)
            (
                datos_balance_general,
                datos_estado_resultados,
                _df_final,
                _df_filtered_head,
                _df_unique_accounts_head,
                _df_unique_accounts_info_text,
            ) = process_dataframe(raw_df, mes, estado, anio, centro_costos, desglosar_por_tercero)

            # Construir bytes Excel
            bg_bytes = create_excel_download_bytes(
                datos_balance_general, sheet_name="datos_balance_general"
            )
            er_bytes = create_excel_download_bytes(
                datos_estado_resultados, sheet_name="datos_estado_resultados"
            )

            # Crear ZIP en memoria
            zip_buffer = BytesIO()
            with ZipFile(zip_buffer, mode="w") as zf:
                zf.writestr("datos_balance_general.xlsx", bg_bytes)
                zf.writestr("datos_estado_resultados.xlsx", er_bytes)
            zip_buffer.seek(0)

            st.download_button(
                label="Descargar resultados.zip (datos crudos)",
                data=zip_buffer.getvalue(),
                file_name="resultados.zip",
                mime="application/zip",
            )

            st.success("ZIP generado con ambos archivos Excel (datos crudos).")

            # --- INFORME COMPLETO ---
            if generar_informe_flag:
                with st.spinner("Generando INFORME completo con formato..."):
                    periodo_label = f"{mes} de {anio}" if mes and anio else f"{mes} {anio}"
                    informe_bytes = generate_informe(
                        df_balance=datos_balance_general,
                        df_er=datos_estado_resultados,
                        branding=branding_config,
                        periodo_actual=periodo_label,
                    )

                st.download_button(
                    label="Descargar INFORME completo (.xlsx)",
                    data=informe_bytes,
                    file_name="INFORME.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.success("INFORME completo generado con formato profesional.")

        except Exception as exc:  # noqa: BLE001
            st.error(f"Error procesando el archivo: {exc}")


if __name__ == "__main__":
    main()


