# -*- coding: utf-8 -*-
"""
Report Generator — Genera el INFORME financiero completo en Excel.

Replica la estructura del INFORME.xls profesional con:
  - EF: Estado de Situación Financiera (agregado por SUBGRUPO, comparativo)
  - ER: Estado de Resultado Integral (waterfall P&L)
  - Notas: Notas a los Estados Financieros (detalle por subcuenta)
  - Graficas: Gráficas de composición
"""

from __future__ import annotations

from io import BytesIO
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.drawing.image import Image as XlImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# NIIF presentation names for each SUBGRUPO (used in EF)
# ---------------------------------------------------------------------------
NIIF_NAMES = {
    "11": "Efectivo y Equivalentes al Efectivo",
    "12": "Inversiones",
    "13": "Deudores comerciales y otras cuentas por cobrar",
    "14": "Inventarios",
    "15": "Propiedades, Planta y Equipo",
    "16": "Intangibles",
    "17": "Diferidos",
    "18": "Otros Activos",
    "19": "Valorizaciones",
    "21": "Obligaciones financieras",
    "22": "Proveedores",
    "23": "Cuentas comerciales por pagar y otras cuentas por pagar",
    "24": "Impuestos corrientes por pagar",
    "25": "Obligaciones laborales",
    "26": "Pasivos estimados y provisiones",
    "27": "Ingresos recibidos por anticipado",
    "28": "Otros pasivos",
    "29": "Bonos y papeles comerciales",
    "31": "Capital emitido",
    "32": "Superávit de capital",
    "33": "Reservas",
    "34": "Revalorización del patrimonio",
    "35": "Dividendos o participaciones",
    "36": "Resultado integral del año",
    "37": "Resultados de ejercicios anteriores",
    "38": "Superávit por valorizaciones",
}

ER_NIIF_NAMES = {
    "41": "Venta de producto y prestación de servicios",
    "42": "Ingresos no ordinarios",
    "47": "Ajustes por inflación (ingresos)",
    "51": "Gastos de administración",
    "52": "Gastos de Venta",
    "53": "Gastos no ordinarios",
    "54": "Provisión Impuesto de Renta",
    "59": "Ganancias y pérdidas",
    "61": "Costo de ventas y prestación de servicios",
    "62": "Compras",
    "71": "Materia prima",
    "72": "Mano de obra directa",
    "73": "Costos indirectos",
    "74": "Contratos de servicios",
}

# Suffix C=Corriente, F=Fijo/No Corriente for the code display
SUBGRUPO_SUFFIX = {
    "11": "C", "12": "C", "13": "C", "14": "C",
    "15": "F", "16": "F", "17": "F", "18": "F", "19": "F",
    "21": "C", "22": "C", "23": "C", "24": "C", "25": "C", "26": "C", "28": "C",
    "27": "F", "29": "F",
    "31": "F", "32": "F", "33": "F", "34": "F", "35": "F", "36": "F", "37": "F", "38": "F",
}

# ---------------------------------------------------------------------------
# Colors matching the INFORME.xls teal/professional style
# ---------------------------------------------------------------------------
COL_TEAL_DARK = "1B6B6D"
COL_TEAL_MED = "3A9A9C"
COL_TEAL_LIGHT = "B5DFE0"
COL_TEAL_BG = "E0F2F2"
COL_WHITE = "FFFFFF"
COL_BLACK = "000000"
COL_DARK_TEXT = "1B1B1B"
COL_RED = "C00000"
COL_GRAY = "808080"
COL_LIGHT_GRAY = "F5F5F5"

# ---------------------------------------------------------------------------
# Reusable styles
# ---------------------------------------------------------------------------
THIN_BORDER = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
BOX_BORDER = Border(
    left=Side(style="medium", color=COL_BLACK),
    right=Side(style="medium", color=COL_BLACK),
    top=Side(style="medium", color=COL_BLACK),
    bottom=Side(style="medium", color=COL_BLACK),
)

FONT_TITLE = Font(name="Arial Narrow", size=12, bold=True, color=COL_BLACK)
FONT_SUBTITLE = Font(name="Arial Narrow", size=10, bold=True, color=COL_BLACK)
FONT_SMALL_ITALIC = Font(name="Arial Narrow", size=9, italic=True, color=COL_BLACK)
FONT_HEADER = Font(name="Arial Narrow", size=9, bold=True, color=COL_BLACK)
FONT_BODY = Font(name="Arial Narrow", size=9, color=COL_BLACK)
FONT_BODY_BOLD = Font(name="Arial Narrow", size=9, bold=True, color=COL_BLACK)
FONT_TOTAL_WHITE = Font(name="Arial Narrow", size=9, bold=True, color=COL_WHITE)
FONT_SECTION = Font(name="Arial Narrow", size=9, bold=True, color=COL_BLACK)
FONT_FIRMA = Font(name="Arial Narrow", size=9, color=COL_BLACK)
FONT_FIRMA_BOLD = Font(name="Arial Narrow", size=9, bold=True, color=COL_BLACK)
FONT_FILTRO = Font(name="Arial Narrow", size=8, color=COL_GRAY)

FILL_TEAL_DARK = PatternFill(start_color=COL_TEAL_DARK, end_color=COL_TEAL_DARK, fill_type="solid")
FILL_TEAL_MED = PatternFill(start_color=COL_TEAL_MED, end_color=COL_TEAL_MED, fill_type="solid")
FILL_TEAL_LIGHT = PatternFill(start_color=COL_TEAL_LIGHT, end_color=COL_TEAL_LIGHT, fill_type="solid")
FILL_TEAL_BG = PatternFill(start_color=COL_TEAL_BG, end_color=COL_TEAL_BG, fill_type="solid")
FILL_WHITE = PatternFill(start_color=COL_WHITE, end_color=COL_WHITE, fill_type="solid")
FILL_LIGHT_GRAY = PatternFill(start_color=COL_LIGHT_GRAY, end_color=COL_LIGHT_GRAY, fill_type="solid")

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

NUM_FMT = '#,##0'
NUM_FMT_NEG = '#,##0;(#,##0)'
PCT_FMT = '0.00%'


def _sc(cell, font=None, fill=None, alignment=None, border=None, number_format=None):
    """Style a cell."""
    if font: cell.font = font
    if fill: cell.fill = fill
    if alignment: cell.alignment = alignment
    if border: cell.border = border
    if number_format: cell.number_format = number_format


def _wc(ws, r, c, val, **kw):
    """Write + style a cell."""
    cell = ws.cell(row=r, column=c, value=val)
    _sc(cell, **kw)
    return cell


def _money(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    return float(v)


# ---------------------------------------------------------------------------
# Branding
# ---------------------------------------------------------------------------
def make_branding(
    empresa: str = "",
    nit: str = "",
    representante_legal: str = "",
    representante_cc: str = "",
    contador: str = "",
    contador_tp: str = "",
    contador_cc: str = "",
    logo_bytes: bytes | None = None,
) -> dict[str, Any]:
    return dict(
        empresa=empresa, nit=nit,
        representante_legal=representante_legal, representante_cc=representante_cc,
        contador=contador, contador_tp=contador_tp, contador_cc=contador_cc,
        logo_bytes=logo_bytes,
    )


# ---------------------------------------------------------------------------
# Aggregate data by SUBGRUPO (2-digit code)
# ---------------------------------------------------------------------------
def _aggregate_by_subgrupo(df: pd.DataFrame) -> pd.DataFrame:
    """
    Takes a df_balance or df_er (output from process_dataframe) and aggregates
    VALOR by the first 2 digits of the account code (SUBGRUPO level).
    """
    if df.empty:
        return pd.DataFrame(columns=["SUBGRUPO_COD", "SUBGRUPO", "CLASE", "GRUPO", "VALOR"])

    df = df.copy()
    # Extract 2-digit subgrupo code from the CUENTA column ("130505 - Nacionales" → "13")
    df["SUBGRUPO_COD"] = df["CUENTA"].astype(str).str.extract(r'^(\d{2})')[0]

    agg = df.groupby(["SUBGRUPO_COD", "SUBGRUPO", "CLASE", "GRUPO"], as_index=False, dropna=False).agg(
        VALOR=("VALOR", "sum")
    )
    return agg.sort_values("SUBGRUPO_COD")


# ---------------------------------------------------------------------------
# Header block
# ---------------------------------------------------------------------------
def _write_header(ws, branding, title, subtitle, start_row=1, merge_cols=8):
    row = start_row

    if branding.get("logo_bytes"):
        try:
            img = XlImage(BytesIO(branding["logo_bytes"]))
            img.width, img.height = 150, 50
            ws.add_image(img, f"A{row}")
        except Exception:
            pass
        row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=merge_cols)
    _wc(ws, row, 1, branding.get("empresa", ""), font=FONT_TITLE, alignment=ALIGN_CENTER)
    row += 1

    if branding.get("nit"):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=merge_cols)
        _wc(ws, row, 1, f"NIT. {branding['nit']}", font=FONT_SUBTITLE, alignment=ALIGN_CENTER)
        row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=merge_cols)
    _wc(ws, row, 1, title, font=FONT_SUBTITLE, alignment=ALIGN_CENTER)
    row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=merge_cols)
    _wc(ws, row, 1, subtitle, font=FONT_SMALL_ITALIC, alignment=ALIGN_CENTER)
    row += 2

    return row


# ---------------------------------------------------------------------------
# Signature block
# ---------------------------------------------------------------------------
def _write_signatures(ws, branding, start_row, col_left=1, col_right=6):
    row = start_row + 2
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    _wc(ws, row, 1, "Las notas adjuntas son parte integral de este estado financiero.",
        font=Font(name="Arial Narrow", size=8, italic=True, color=COL_GRAY), alignment=ALIGN_CENTER)
    row += 3

    # Rep Legal
    _wc(ws, row, col_left, branding.get("representante_legal", ""), font=FONT_FIRMA_BOLD)
    row += 1
    _wc(ws, row, col_left, "Representante Legal", font=FONT_FIRMA)
    row += 1
    if branding.get("representante_cc"):
        _wc(ws, row, col_left, f"C.C. {branding['representante_cc']}", font=FONT_FIRMA)

    # Contador
    cr = start_row + 5
    _wc(ws, cr, col_right, branding.get("contador", ""), font=FONT_FIRMA_BOLD)
    cr += 1
    _wc(ws, cr, col_right, "Contadora Pública" if branding.get("contador") else "", font=FONT_FIRMA)
    cr += 1
    if branding.get("contador_tp"):
        _wc(ws, cr, col_right, f"T.P. {branding['contador_tp']}", font=FONT_FIRMA)
        cr += 1
    if branding.get("contador_cc"):
        _wc(ws, cr, col_right, f"C.C {branding['contador_cc']}", font=FONT_FIRMA)


# ---------------------------------------------------------------------------
# Write a data row in the EF/ER format
# ---------------------------------------------------------------------------
def _ef_data_row(ws, row, code, name, nota, val_actual, pct_actual, val_ant, pct_ant, var_val, var_pct,
                 is_total=False, is_grand_total=False, is_section_header=False, filtro=True,
                 col_start=1):
    """Write one row in the EF or ER layout."""

    if is_grand_total:
        font = FONT_TOTAL_WHITE
        fill = FILL_TEAL_DARK
    elif is_total:
        font = FONT_BODY_BOLD
        fill = FILL_TEAL_LIGHT
    elif is_section_header:
        font = FONT_SECTION
        fill = FILL_WHITE
    else:
        font = FONT_BODY
        fill = FILL_WHITE

    neg_font_actual = Font(name="Arial Narrow", size=9, bold=is_total or is_grand_total,
                            color=COL_WHITE if is_grand_total else COL_BLACK)

    c = col_start
    # Col A: subgrupo code
    _wc(ws, row, c, code if code else "", font=font, fill=fill, border=THIN_BORDER, alignment=ALIGN_LEFT); c += 1
    # Col B: name
    _wc(ws, row, c, name, font=font, fill=fill, border=THIN_BORDER, alignment=ALIGN_LEFT); c += 1
    # Col C: nota
    _wc(ws, row, c, nota if nota else "", font=font, fill=fill, border=THIN_BORDER, alignment=ALIGN_CENTER); c += 1
    # Col D: valor actual
    _wc(ws, row, c, val_actual, font=neg_font_actual, fill=fill, border=THIN_BORDER,
        alignment=ALIGN_RIGHT, number_format=NUM_FMT_NEG); c += 1
    # Col E: % actual
    _wc(ws, row, c, pct_actual, font=neg_font_actual, fill=fill, border=THIN_BORDER,
        alignment=ALIGN_RIGHT, number_format=PCT_FMT); c += 1
    # Col F: valor anterior
    _wc(ws, row, c, val_ant, font=neg_font_actual, fill=fill, border=THIN_BORDER,
        alignment=ALIGN_RIGHT, number_format=NUM_FMT_NEG); c += 1
    # Col G: % anterior
    _wc(ws, row, c, pct_ant, font=neg_font_actual, fill=fill, border=THIN_BORDER,
        alignment=ALIGN_RIGHT, number_format=PCT_FMT); c += 1
    # Col H: variación
    _wc(ws, row, c, var_val, font=neg_font_actual, fill=fill, border=THIN_BORDER,
        alignment=ALIGN_RIGHT, number_format=NUM_FMT_NEG); c += 1
    # Col I: % variación
    _wc(ws, row, c, var_pct, font=neg_font_actual, fill=fill, border=THIN_BORDER,
        alignment=ALIGN_RIGHT, number_format=PCT_FMT); c += 1
    # Col J: FILTRO
    _wc(ws, row, c, "TRUE" if filtro else "FALSE", font=FONT_FILTRO, alignment=ALIGN_CENTER)


# ---------------------------------------------------------------------------
# EF — Estado de Situación Financiera
# ---------------------------------------------------------------------------
def _build_ef_sheet(wb, df_balance, branding, periodo_actual, periodo_anterior=None, df_balance_anterior=None):
    ws = wb.create_sheet("EF")
    merge_cols = 9

    row = _write_header(ws, branding, "ESTADO DE SITUACIÓN FINANCIERA",
                        "(Expresados en pesos colombianos)", merge_cols=merge_cols)

    # "Período terminado a"
    _wc(ws, row, 1, "Período terminado a", font=FONT_SMALL_ITALIC); row += 1

    # Column headers
    per_ant_label = periodo_anterior or "Período Anterior"
    headers = ["", "ACTIVO", "Nota", periodo_actual, "%", per_ant_label, "%", "Variación", "%", "FILTRO"]
    for i, h in enumerate(headers, 1):
        _wc(ws, row, i, h, font=FONT_HEADER, fill=FILL_TEAL_BG, border=THIN_BORDER, alignment=ALIGN_CENTER)
    row += 1

    agg = _aggregate_by_subgrupo(df_balance)
    agg_ant = _aggregate_by_subgrupo(df_balance_anterior) if df_balance_anterior is not None else None

    total_activo = agg.loc[agg["CLASE"] == "ACTIVO", "VALOR"].sum()
    if total_activo == 0:
        total_activo = 1

    total_activo_ant = 0
    if agg_ant is not None:
        total_activo_ant = agg_ant.loc[agg_ant["CLASE"] == "ACTIVO", "VALOR"].sum()

    nota_num = 1

    sections_ef = [
        ("ACTIVO", [
            ("CORRIENTE", "ACTIVO CORRIENTE",
             ["11", "12", "13", "14"]),
            ("NO CORRIENTE", "ACTIVO NO CORRIENTE",
             ["15", "16", "17", "18", "19"]),
        ]),
        ("PASIVO", [
            ("PASIVO CORRIENTE", "PASIVO CORRIENTE",
             ["21", "22", "23", "24", "25", "26", "28"]),
            ("PASIVO NO CORRIENTE", "PASIVO NO CORRIENTE",
             ["27", "29"]),
        ]),
        ("PATRIMONIO", [
            ("PATRIMONIO", "PATRIMONIO",
             ["31", "32", "33", "34", "35", "36", "37", "38"]),
        ]),
    ]

    grand_totals = {}

    for clase_label, groups in sections_ef:
        if clase_label == "PASIVO":
            row += 1  # blank before PASIVO

        clase_total_val = 0
        clase_total_ant = 0

        for section_title, grupo_name, subgrupo_codes in groups:
            # Section header (e.g., "CORRIENTE", "NO CORRIENTE")
            _ef_data_row(ws, row, "", section_title, "", "", "", "", "", "", "",
                         is_section_header=True, filtro=True)
            row += 1

            section_val = 0
            section_ant = 0

            for sg_code in subgrupo_codes:
                sg_data = agg[agg["SUBGRUPO_COD"] == sg_code]
                val = sg_data["VALOR"].sum() if not sg_data.empty else 0

                val_ant = 0
                if agg_ant is not None:
                    sg_ant_data = agg_ant[agg_ant["SUBGRUPO_COD"] == sg_code]
                    val_ant = sg_ant_data["VALOR"].sum() if not sg_ant_data.empty else 0

                if val == 0 and val_ant == 0:
                    continue

                section_val += val
                section_ant += val_ant
                var_val = val - val_ant
                pct_act = val / total_activo if total_activo else 0
                pct_ant_v = val_ant / total_activo_ant if total_activo_ant else 0
                var_pct = (var_val / abs(val_ant)) if val_ant != 0 else 0

                niif_name = NIIF_NAMES.get(sg_code, sg_code)
                suffix = SUBGRUPO_SUFFIX.get(sg_code, "")
                display_code = f"{sg_code} {sg_code}{suffix}"

                note_val = nota_num
                nota_num += 1

                _ef_data_row(ws, row, display_code, niif_name, note_val,
                             val, pct_act, val_ant, pct_ant_v, var_val, var_pct, filtro=True)
                row += 1

            # Section total
            var_sec = section_val - section_ant
            pct_sec = section_val / total_activo if total_activo else 0
            pct_sec_ant = section_ant / total_activo_ant if total_activo_ant else 0
            var_pct_sec = (var_sec / abs(section_ant)) if section_ant != 0 else 0

            _ef_data_row(ws, row, "", f"TOTAL {section_title}", "",
                         section_val, pct_sec, section_ant, pct_sec_ant, var_sec, var_pct_sec,
                         is_total=True, filtro=True)
            row += 1

            clase_total_val += section_val
            clase_total_ant += section_ant

        # Grand total for clase (TOTAL ACTIVO, TOTAL PASIVO, TOTAL PATRIMONIO)
        var_clase = clase_total_val - clase_total_ant
        pct_clase = clase_total_val / total_activo if total_activo else 0
        pct_clase_ant = clase_total_ant / total_activo_ant if total_activo_ant else 0
        var_pct_clase = (var_clase / abs(clase_total_ant)) if clase_total_ant != 0 else 0

        _ef_data_row(ws, row, "", f"TOTAL {clase_label}", "",
                     clase_total_val, pct_clase, clase_total_ant, pct_clase_ant, var_clase, var_pct_clase,
                     is_grand_total=True, filtro=True)
        row += 1

        grand_totals[clase_label] = (clase_total_val, clase_total_ant)

    # TOTAL PASIVO Y PATRIMONIO
    row += 1
    pasivo_val = grand_totals.get("PASIVO", (0, 0))[0] + grand_totals.get("PATRIMONIO", (0, 0))[0]
    pasivo_ant = grand_totals.get("PASIVO", (0, 0))[1] + grand_totals.get("PATRIMONIO", (0, 0))[1]
    var_pp = pasivo_val - pasivo_ant
    pct_pp = pasivo_val / total_activo if total_activo else 0
    pct_pp_ant = pasivo_ant / total_activo_ant if total_activo_ant else 0
    var_pct_pp = (var_pp / abs(pasivo_ant)) if pasivo_ant != 0 else 0

    _ef_data_row(ws, row, "", "TOTAL PASIVO Y PATRIMONIO", "",
                 pasivo_val, pct_pp, pasivo_ant, pct_pp_ant, var_pp, var_pct_pp,
                 is_grand_total=True, filtro=True)
    row += 2

    _write_signatures(ws, branding, row)

    # Column widths
    widths = {"A": 10, "B": 52, "C": 6, "D": 16, "E": 9, "F": 16, "G": 9, "H": 16, "I": 9, "J": 8}
    for col_letter, w in widths.items():
        ws.column_dimensions[col_letter].width = w

    return nota_num


# ---------------------------------------------------------------------------
# ER — Estado de Resultado Integral (P&L waterfall)
# ---------------------------------------------------------------------------
def _build_er_sheet(wb, df_er, branding, periodo_actual, nota_start=1,
                    periodo_anterior=None, df_er_anterior=None):
    ws = wb.create_sheet("ER")
    merge_cols = 9

    row = _write_header(ws, branding, "ESTADO DE RESULTADO INTEGRAL",
                        "(Expresados en pesos colombianos)", merge_cols=merge_cols)

    # Period label
    _wc(ws, row, 1, "Período comprendido entre", font=FONT_SMALL_ITALIC); row += 1
    _wc(ws, row, 1, f"01 de Enero y al fin del mes de : {periodo_actual}", font=FONT_SMALL_ITALIC); row += 1
    row += 1

    # Column headers
    per_ant_label = periodo_anterior or "Período Anterior"
    headers = ["Actividades Ordinarias", "", "Nota", periodo_actual, "%", per_ant_label, "%", "VARIACIÓN", "%", "FILTRO"]
    for i, h in enumerate(headers, 1):
        _wc(ws, row, i, h, font=FONT_HEADER, fill=FILL_TEAL_BG, border=THIN_BORDER, alignment=ALIGN_CENTER)
    row += 1

    agg = _aggregate_by_subgrupo(df_er)
    agg_ant = _aggregate_by_subgrupo(df_er_anterior) if df_er_anterior is not None else None

    def _get_val(sg_code):
        d = agg[agg["SUBGRUPO_COD"] == sg_code]
        return d["VALOR"].sum() if not d.empty else 0

    def _get_val_ant(sg_code):
        if agg_ant is None:
            return 0
        d = agg_ant[agg_ant["SUBGRUPO_COD"] == sg_code]
        return d["VALOR"].sum() if not d.empty else 0

    # Ingresos operacionales (41)
    ingresos_41 = _get_val("41")
    ingresos_41_ant = _get_val_ant("41")
    total_ingresos = abs(ingresos_41) if ingresos_41 != 0 else 1
    total_ingresos_ant = abs(ingresos_41_ant) if ingresos_41_ant != 0 else 1

    def _pct(val, base):
        return val / base if base != 0 else 0

    def _var_pct(val, val_ant):
        return (val - val_ant) / abs(val_ant) if val_ant != 0 else 0

    nota = nota_start

    def _row_item(label, val, val_ant, note=None, bold=False, grand=False, total_bg=False):
        nonlocal row, nota
        var = val - val_ant
        n = None
        if note:
            n = nota
            nota += 1

        if grand:
            _ef_data_row(ws, row, "", label, n, val, _pct(val, total_ingresos),
                         val_ant, _pct(val_ant, total_ingresos_ant),
                         var, _var_pct(val, val_ant), is_grand_total=True)
        elif total_bg:
            _ef_data_row(ws, row, "", label, n, val, _pct(val, total_ingresos),
                         val_ant, _pct(val_ant, total_ingresos_ant),
                         var, _var_pct(val, val_ant), is_total=True)
        else:
            font = FONT_BODY_BOLD if bold else FONT_BODY
            c = 1
            _wc(ws, row, c, "", font=font, fill=FILL_WHITE, border=THIN_BORDER); c += 1
            _wc(ws, row, c, label, font=font, fill=FILL_WHITE, border=THIN_BORDER, alignment=ALIGN_LEFT); c += 1
            _wc(ws, row, c, n if n else "", font=font, fill=FILL_WHITE, border=THIN_BORDER, alignment=ALIGN_CENTER); c += 1
            _wc(ws, row, c, val, font=font, fill=FILL_WHITE, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=NUM_FMT_NEG); c += 1
            _wc(ws, row, c, _pct(val, total_ingresos), font=font, fill=FILL_WHITE, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FMT); c += 1
            _wc(ws, row, c, val_ant, font=font, fill=FILL_WHITE, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=NUM_FMT_NEG); c += 1
            _wc(ws, row, c, _pct(val_ant, total_ingresos_ant), font=font, fill=FILL_WHITE, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FMT); c += 1
            _wc(ws, row, c, var, font=font, fill=FILL_WHITE, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=NUM_FMT_NEG); c += 1
            _wc(ws, row, c, _var_pct(val, val_ant), font=font, fill=FILL_WHITE, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FMT); c += 1
            _wc(ws, row, c, "TRUE", font=FONT_FILTRO)

        row += 1

    # --- P&L Waterfall ---

    # Ventas (41)
    _row_item("Venta de producto y prestación de servicios", ingresos_41, ingresos_41_ant, note=True)

    # Costos de ventas (61, 62, 71-74)
    costo_codes = ["61", "62", "71", "72", "73", "74"]
    costos = sum(_get_val(c) for c in costo_codes)
    costos_ant = sum(_get_val_ant(c) for c in costo_codes)
    _row_item("MENOS:  COSTO DE VENTAS", costos, costos_ant)

    # UTILIDAD BRUTA
    util_bruta = ingresos_41 - costos
    util_bruta_ant = ingresos_41_ant - costos_ant
    _row_item("UTILIDAD BRUTA", util_bruta, util_bruta_ant, total_bg=True)

    # Gastos de administración (51)
    gastos_admin = _get_val("51")
    gastos_admin_ant = _get_val_ant("51")
    _row_item("Gastos de administración", gastos_admin, gastos_admin_ant, note=True)

    # Gastos de venta (52)
    gastos_venta = _get_val("52")
    gastos_venta_ant = _get_val_ant("52")
    _row_item("Gastos de Venta", gastos_venta, gastos_venta_ant)

    # UTILIDAD OPERACIONAL
    util_oper = util_bruta - gastos_admin - gastos_venta
    util_oper_ant = util_bruta_ant - gastos_admin_ant - gastos_venta_ant
    _row_item("UTILIDAD OPERACIONAL", util_oper, util_oper_ant, grand=True)

    # Ingresos no ordinarios (42)
    ing_no_ord = _get_val("42")
    ing_no_ord_ant = _get_val_ant("42")
    _row_item("Ingresos no ordinarios", ing_no_ord, ing_no_ord_ant, note=True)

    # Gastos no ordinarios (53)
    gtos_no_ord = _get_val("53")
    gtos_no_ord_ant = _get_val_ant("53")
    _row_item("Gastos no ordinarios", gtos_no_ord, gtos_no_ord_ant, note=True)

    # UTILIDAD (PERDIDA) ANTES DE IMPUESTOS
    util_antes_imp = util_oper + ing_no_ord - gtos_no_ord
    util_antes_imp_ant = util_oper_ant + ing_no_ord_ant - gtos_no_ord_ant
    _row_item("UTILIDAD (PERDIDA) ANTES DE IMPUESTOS", util_antes_imp, util_antes_imp_ant, grand=True)

    # Provisión Impuesto de Renta (54)
    imp_renta = _get_val("54")
    imp_renta_ant = _get_val_ant("54")
    _row_item("Provisión Impuesto de Renta", imp_renta, imp_renta_ant)

    # UTILIDAD NETA
    util_neta = util_antes_imp - imp_renta
    util_neta_ant = util_antes_imp_ant - imp_renta_ant
    _row_item("UTILIDAD NETA", util_neta, util_neta_ant, grand=True)

    row += 1
    _write_signatures(ws, branding, row)

    widths = {"A": 8, "B": 48, "C": 6, "D": 16, "E": 9, "F": 16, "G": 9, "H": 16, "I": 9, "J": 8}
    for col_letter, w in widths.items():
        ws.column_dimensions[col_letter].width = w


# ---------------------------------------------------------------------------
# Notas
# ---------------------------------------------------------------------------
def _build_notas_sheet(wb, df_balance, df_er, branding, periodo_actual):
    ws = wb.create_sheet("Notas")

    row = _write_header(ws, branding, "NOTAS A LOS ESTADOS FINANCIEROS",
                        f"(Expresados en pesos colombianos) — {periodo_actual}", merge_cols=6)

    combined = pd.concat([df_balance, df_er], ignore_index=True)
    if combined.empty:
        _wc(ws, row, 1, "Sin datos para generar notas.", font=FONT_BODY)
        return

    combined["SUBGRUPO_COD"] = combined["CUENTA"].astype(str).str.extract(r'^(\d{2})')[0]

    nota_num = 1
    subgrupo_codes = sorted(combined["SUBGRUPO_COD"].dropna().unique())

    for sg_code in subgrupo_codes:
        sg_data = combined[combined["SUBGRUPO_COD"] == sg_code]
        if sg_data.empty:
            continue

        total_sub = sg_data["VALOR"].sum()
        if total_sub == 0 and len(sg_data) < 2:
            continue

        niif_name = NIIF_NAMES.get(sg_code, ER_NIIF_NAMES.get(sg_code, sg_code))

        # Nota header
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        _wc(ws, row, 1, f"NOTA {nota_num}    {niif_name}",
            font=FONT_SECTION, fill=FILL_TEAL_BG, border=THIN_BORDER)
        for c in range(2, 5):
            _sc(ws.cell(row=row, column=c), fill=FILL_TEAL_BG, border=THIN_BORDER)
        row += 1

        # Column headers
        for i, h in enumerate(["CTA", "Cuenta", "Tercero", periodo_actual], 1):
            _wc(ws, row, i, h, font=FONT_HEADER, fill=FILL_TEAL_LIGHT, border=THIN_BORDER, alignment=ALIGN_CENTER)
        row += 1

        for _, acct_row in sg_data.iterrows():
            val = _money(acct_row.get("VALOR", 0))
            cuenta_str = str(acct_row.get("CUENTA", ""))
            codigo = cuenta_str.split(" - ")[0].strip() if " - " in cuenta_str else ""
            nombre = cuenta_str.split(" - ")[1].strip() if " - " in cuenta_str else cuenta_str

            tercero = str(acct_row.get("TERCERO", "")) if pd.notna(acct_row.get("TERCERO")) else ""

            _wc(ws, row, 1, codigo, font=FONT_BODY, border=THIN_BORDER)
            _wc(ws, row, 2, nombre, font=FONT_BODY, border=THIN_BORDER, alignment=ALIGN_LEFT)
            _wc(ws, row, 3, tercero, font=FONT_BODY, border=THIN_BORDER, alignment=ALIGN_LEFT)
            _wc(ws, row, 4, val, font=FONT_BODY, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=NUM_FMT_NEG)
            row += 1

        _wc(ws, row, 1, "", fill=FILL_TEAL_LIGHT, border=THIN_BORDER)
        _wc(ws, row, 2, "Total", font=FONT_BODY_BOLD, fill=FILL_TEAL_LIGHT, border=THIN_BORDER)
        _wc(ws, row, 3, "", fill=FILL_TEAL_LIGHT, border=THIN_BORDER)
        _wc(ws, row, 4, total_sub, font=FONT_BODY_BOLD, fill=FILL_TEAL_LIGHT, border=THIN_BORDER,
            alignment=ALIGN_RIGHT, number_format=NUM_FMT_NEG)
        row += 2

        nota_num += 1

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["D"].width = 18


# ---------------------------------------------------------------------------
# Graficas
# ---------------------------------------------------------------------------
def _build_graficas_sheet(wb, df_balance, df_er, branding):
    ws = wb.create_sheet("Graficas")

    _wc(ws, 1, 1, branding.get("empresa", ""), font=FONT_TITLE)
    _wc(ws, 2, 1, "GRÁFICAS DE COMPOSICIÓN FINANCIERA", font=FONT_SUBTITLE)

    _build_pie(ws, df_balance, "ACTIVO", "Composición del Activo", 5, "A20", 0)
    _build_pie(ws, df_balance, "PASIVO", "Composición del Pasivo", 5, "J20", 5)
    _build_bar(ws, df_er, 5, "A38")

    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["F"].width = 25
    ws.column_dimensions["G"].width = 15


def _build_pie(ws, df, clase, title, data_row, anchor, col_off):
    data = df[df["CLASE"] == clase]
    if data.empty:
        return

    data = data.copy()
    data["SG"] = data["CUENTA"].astype(str).str.extract(r'^(\d{2})')[0]
    grouped = data.groupby("SG", as_index=False)["VALOR"].sum()
    grouped = grouped[grouped["VALOR"] != 0]
    if grouped.empty:
        return

    grouped["LABEL"] = grouped["SG"].map(lambda x: NIIF_NAMES.get(x, x))

    cl, cv = 1 + col_off, 2 + col_off
    _wc(ws, data_row, cl, "Concepto", font=FONT_HEADER, fill=FILL_TEAL_LIGHT, border=THIN_BORDER)
    _wc(ws, data_row, cv, "Valor", font=FONT_HEADER, fill=FILL_TEAL_LIGHT, border=THIN_BORDER)

    for i, (_, r) in enumerate(grouped.iterrows()):
        rr = data_row + 1 + i
        _wc(ws, rr, cl, r["LABEL"], font=FONT_BODY, border=THIN_BORDER)
        _wc(ws, rr, cv, abs(r["VALOR"]), font=FONT_BODY, border=THIN_BORDER, number_format=NUM_FMT)

    last = data_row + len(grouped)
    pie = PieChart()
    pie.title = title
    pie.style = 10
    pie.width, pie.height = 18, 14
    labels = Reference(ws, min_col=cl, min_row=data_row + 1, max_row=last)
    vals = Reference(ws, min_col=cv, min_row=data_row, max_row=last)
    pie.add_data(vals, titles_from_data=True)
    pie.set_categories(labels)
    ws.add_chart(pie, anchor)


def _build_bar(ws, df_er, data_row, anchor):
    if df_er.empty:
        return

    df_er = df_er.copy()
    df_er["SG"] = df_er["CUENTA"].astype(str).str.extract(r'^(\d{2})')[0]

    items = {}
    for sg, label in [("41", "Ingresos"), ("51", "Gastos Admin"), ("52", "Gastos Venta"),
                       ("61", "Costo Ventas"), ("53", "Gastos No Op")]:
        d = df_er[df_er["SG"] == sg]
        v = abs(d["VALOR"].sum()) if not d.empty else 0
        if v > 0:
            items[label] = v

    if not items:
        return

    col = 10
    _wc(ws, data_row, col, "Concepto", font=FONT_HEADER, fill=FILL_TEAL_LIGHT, border=THIN_BORDER)
    _wc(ws, data_row, col + 1, "Valor", font=FONT_HEADER, fill=FILL_TEAL_LIGHT, border=THIN_BORDER)

    for i, (label, val) in enumerate(items.items()):
        rr = data_row + 1 + i
        _wc(ws, rr, col, label, font=FONT_BODY, border=THIN_BORDER)
        _wc(ws, rr, col + 1, val, font=FONT_BODY, border=THIN_BORDER, number_format=NUM_FMT)

    last = data_row + len(items)
    bar = BarChart()
    bar.type = "col"
    bar.title = "Ingresos vs Gastos vs Costos"
    bar.style = 10
    bar.width, bar.height = 18, 14
    bar.y_axis.title = "Pesos ($)"
    cats = Reference(ws, min_col=col, min_row=data_row + 1, max_row=last)
    vals = Reference(ws, min_col=col + 1, min_row=data_row, max_row=last)
    bar.add_data(vals, titles_from_data=True)
    bar.set_categories(cats)
    ws.add_chart(bar, anchor)


# ===========================================================================
# PUBLIC API
# ===========================================================================
def generate_informe(
    df_balance: pd.DataFrame,
    df_er: pd.DataFrame,
    branding: dict,
    periodo_actual: str,
    periodo_anterior: str | None = None,
    df_balance_anterior: pd.DataFrame | None = None,
    df_er_anterior: pd.DataFrame | None = None,
) -> bytes:
    """
    Genera el INFORME completo como archivo Excel (.xlsx) en memoria,
    replicando el formato del INFORME.xls profesional.
    """
    wb = Workbook()
    wb.remove(wb.active)

    nota_next = _build_ef_sheet(wb, df_balance, branding, periodo_actual,
                                 periodo_anterior, df_balance_anterior)
    _build_er_sheet(wb, df_er, branding, periodo_actual, nota_start=nota_next,
                    periodo_anterior=periodo_anterior, df_er_anterior=df_er_anterior)
    _build_notas_sheet(wb, df_balance, df_er, branding, periodo_actual)
    _build_graficas_sheet(wb, df_balance, df_er, branding)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()
