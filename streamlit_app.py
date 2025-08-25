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


APP_TITLE = "Oxnya Balance General"


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


def read_siigo_excel(file_obj) -> pd.DataFrame:
    """Read the SIIGO Excel export assuming header on row 8 (header=7)."""
    excel = pd.ExcelFile(file_obj)
    df = excel.parse(header=7)
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


def process_dataframe(
    df: pd.DataFrame, mes: str, estado: str, anio: str, centro_costos: str
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

    # Filter Transaccional == "No"
    df_filtered = df[df["Transaccional"] == "No"].copy()

    # Remove specified columns (exactly as notebook)
    columns_to_drop = [
        "Saldo inicial",
        "Movimiento débito",
        "Movimiento crédito",
        "Sucursal",
        "Identificación",
    ]
    df_filtered = df_filtered.drop(columns=columns_to_drop)

    # Drop duplicates based on "Código cuenta contable"
    df_unique_accounts = df_filtered.drop_duplicates(
        subset=["Código cuenta contable"]
    )

    # Capture info() output
    info_buffer = StringIO()
    df_unique_accounts.info(buf=info_buffer)
    df_unique_accounts_info_text = info_buffer.getvalue()

    # Convertir el código a string para manipulación
    df_unique_accounts = df_unique_accounts.copy()
    df_unique_accounts["Código str"] = (
        df_unique_accounts["Código cuenta contable"].astype(int).astype(str)
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
    df_unique_accounts["VALOR"] = df_unique_accounts["Saldo final"]
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

    # Filter by length of the 'Código cuenta contable' column (after stripping .0) length >= 6
    df_final = df_final[
        df_final["Código cuenta contable"].astype(str).str.replace(
            ".0", "", regex=False
        ).str.len()
        >= 6
    ].copy()

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
        "Este app procesa y agrupa un export de un periodo de operación mensual de SIIGO, "
        "generando Balance General y Estado de Resultados en Excel."
    )

    with st.expander("Instrucciones", expanded=True):
        st.markdown(
            "- Cargue el archivo Excel exportado de SIIGO."
        )
        st.markdown(
            "- El encabezado del archivo debe estar en la fila 8 (header=7), igual que en el notebook."
        )
        st.markdown(
            "- Ingrese MES, ESTADO, AÑO y CENTRO DE COSTOS."
        )
        st.markdown(
            "- Descargue un ZIP con 'datos_balance_general.xlsx' y 'datos_estado_resultados.xlsx'."
        )
        st.markdown(
            "- Nota: Si no se logra determinar 'Nombre tercero' en el archivo, la columna TERCERO se llenará en blanco."
        )

    uploaded_file = st.file_uploader("Suba el archivo Excel de SIIGO", type=["xlsx"])

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        mes = st.text_input("MES", value="")
    with col2:
        estado = st.text_input("ESTADO", value="")
    with col3:
        anio = st.text_input("AÑO", value="")
    with col4:
        centro_costos = st.text_input("CENTRO DE COSTOS", value="")

    process_clicked = st.button("Procesar y Generar Archivos")

    if process_clicked:
        if uploaded_file is None:
            st.error("Debe subir un archivo Excel (.xlsx).")
            st.stop()
        if not mes or not estado or not anio or not centro_costos:
            st.error("Por favor complete MES, ESTADO, AÑO y CENTRO DE COSTOS.")
            st.stop()

        try:
            # Read Excel
            raw_df = read_siigo_excel(uploaded_file)
            # Optional: minimal preview to confirm load (first rows only)
            st.subheader("Vista previa (primeras filas)")
            st.dataframe(raw_df.head())

            # Validate columns and ensure 'Nombre tercero'
            validate_required_columns(raw_df)
            raw_df = ensure_nombre_tercero(raw_df)

            # Process DataFrame (exact notebook logic)
            (
                datos_balance_general,
                datos_estado_resultados,
                _df_final,
                _df_filtered_head,
                _df_unique_accounts_head,
                _df_unique_accounts_info_text,
            ) = process_dataframe(raw_df, mes, estado, anio, centro_costos)

            # Build Excel bytes
            bg_bytes = create_excel_download_bytes(
                datos_balance_general, sheet_name="datos_balance_general"
            )
            er_bytes = create_excel_download_bytes(
                datos_estado_resultados, sheet_name="datos_estado_resultados"
            )

            # Create ZIP in-memory
            zip_buffer = BytesIO()
            with ZipFile(zip_buffer, mode="w") as zf:
                zf.writestr("datos_balance_general.xlsx", bg_bytes)
                zf.writestr("datos_estado_resultados.xlsx", er_bytes)
            zip_buffer.seek(0)

            st.download_button(
                label="Descargar resultados.zip",
                data=zip_buffer.getvalue(),
                file_name="resultados.zip",
                mime="application/zip",
            )

            st.success("ZIP generado con ambos archivos Excel.")

        except Exception as exc:  # noqa: BLE001
            st.error(f"Error procesando el archivo: {exc}")


if __name__ == "__main__":
    main()


