"""Plantillas bancarias didacticas actualizables.

Cada perfil conserva rasgos estructurales observados en las referencias (orden
de resumen, paleta y tabla), pero omite sellos, codigos de barras y elementos de
autenticacion. Todas las paginas se identifican como documentos ficticios.
"""

from __future__ import annotations

import os
from dataclasses import dataclass

import fitz
import pandas as pd


PAGE_W, PAGE_H = 612.0, 792.0


@dataclass(frozen=True)
class Theme:
    key: str
    title: str
    primary: tuple
    secondary: tuple
    dark: tuple
    summary_layout: str


THEMES = {
    "AZTECA": Theme("AZTECA", "BANCO AZTECA - EJERCICIO DIDACTICO", (0.00, 0.39, 0.24), (0.93, 0.95, 0.93), (0.02, 0.20, 0.12), "vertical"),
    "BANAMEX": Theme("BANAMEX", "BANAMEX - EJERCICIO DIDACTICO", (0.82, 0.02, 0.20), (0.89, 0.94, 0.96), (0.00, 0.22, 0.28), "cards"),
    "BBVA 1": Theme("BBVA 1", "BBVA - EJERCICIO DIDACTICO / PERFIL 1", (0.02, 0.32, 0.62), (0.87, 0.94, 0.98), (0.00, 0.13, 0.30), "horizontal"),
    "BBVA 2": Theme("BBVA 2", "BBVA - EJERCICIO DIDACTICO / PERFIL 2", (0.00, 0.42, 0.65), (0.91, 0.96, 0.98), (0.00, 0.18, 0.34), "cards"),
    "SANTANDER": Theme("SANTANDER", "SANTANDER - EJERCICIO DIDACTICO", (0.88, 0.00, 0.10), (0.95, 0.95, 0.95), (0.22, 0.22, 0.22), "horizontal"),
    "GENERICO": Theme("GENERICO", "ESTADO DE CUENTA - EJERCICIO DIDACTICO", (0.20, 0.32, 0.40), (0.93, 0.95, 0.96), (0.12, 0.16, 0.18), "vertical"),
}


def _font_path(filename: str) -> str:
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fuentes", filename)
    if not os.path.exists(path):
        raise FileNotFoundError(f"Falta la fuente {filename} en la carpeta fuentes.")
    return path


def _money(value) -> str:
    try:
        return f"${float(value):,.2f}"
    except Exception:
        return ""


def _text(page, x, y, value, fontfile, fontname, size=8, color=(0, 0, 0)):
    page.insert_text((x, y), str(value), fontfile=fontfile, fontname=fontname, fontsize=size, color=color)


def _right(page, right, y, value, font, fontfile, fontname, size=8, color=(0, 0, 0)):
    value = str(value)
    _text(page, right - font.text_length(value, fontsize=size), y, value, fontfile, fontname, size, color)


def _wrap(text, width, font, size):
    lines, current = [], ""
    for word in str(text).split():
        trial = word if not current else current + " " + word
        if font.text_length(trial, fontsize=size) <= width:
            current = trial
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines or [""]


def _page_frame(page, theme, regular, bold, page_number):
    page.draw_rect((0, 0, PAGE_W, 16), color=theme.primary, fill=theme.primary)
    _text(page, 40, 42, theme.title, bold, "Bold", 14, theme.primary)
    _right(page, 570, 42, f"PAGINA {page_number}", fitz.Font(fontfile=regular), regular, "Regular", 7, theme.dark)
    page.draw_line((40, 50), (570, 50), color=theme.primary, width=1.2)
    page.draw_rect((40, 752, 570, 778), color=theme.primary, fill=theme.primary)
    _text(page, 190, 770, "DOCUMENTO FICTICIO - USO DIDACTICO", bold, "Bold", 10, (1, 1, 1))


def _summary(page, theme, identity, start, end, df, regular, bold, font):
    dep = pd.to_numeric(df["DEPOSITO"], errors="coerce").fillna(0)
    ret = pd.to_numeric(df["RETIRO"], errors="coerce").fillna(0)
    balances = pd.to_numeric(df["SALDO"], errors="coerce").dropna()
    opening = float(balances.iloc[0]) if len(balances) else 0.0
    closing = float(balances.iloc[-1]) if len(balances) else 0.0
    average = float(balances.iloc[1:].mean()) if len(balances) > 1 else closing
    _text(page, 42, 76, identity.get("NOMBRE", "CLIENTE FICTICIO"), bold, "Bold", 10, theme.dark)
    _text(page, 42, 91, identity.get("DOMICILIO", "DOMICILIO FICTICIO"), regular, "Regular", 7.5)
    labels = [("CUENTA FICTICIA", identity.get("CUENTA", "")), ("CLABE FICTICIA", identity.get("CLABE", "")), ("PERIODO", f"{start:%d/%m/%Y} - {end:%d/%m/%Y}"), ("MONEDA", "NACIONAL")]
    y = 73
    for label, value in labels:
        page.draw_rect((355, y - 10, 570, y + 4), color=theme.secondary, fill=theme.secondary)
        _text(page, 360, y, label, bold, "Bold", 7, theme.dark)
        _right(page, 565, y, value, font, regular, "Regular", 7)
        y += 16
    page.draw_rect((40, 150, 570, 171), color=theme.dark, fill=theme.dark)
    _text(page, 48, 165, "RESUMEN DEL PERIODO", bold, "Bold", 9, (1, 1, 1))
    rows = [("Saldo inicial", opening), ("Depositos", float(dep.sum())), ("Retiros", float(ret.sum())), ("Saldo final", closing), ("Saldo promedio", average)]
    if theme.summary_layout == "horizontal":
        widths = 106
        for i, (label, value) in enumerate(rows):
            x = 40 + i * widths
            page.draw_rect((x, 178, x + widths - 3, 230), color=theme.secondary, fill=theme.secondary)
            _text(page, x + 7, 195, label, bold, "Bold", 7, theme.dark)
            _right(page, x + widths - 10, 218, _money(value), font, regular, "Regular", 8.5)
    else:
        y = 190
        for i, (label, value) in enumerate(rows):
            fill = theme.secondary if i % 2 == 0 else (1, 1, 1)
            page.draw_rect((40, y - 13, 330, y + 4), color=fill, fill=fill)
            _text(page, 48, y, label, bold if label == "Saldo final" else regular, "Bold" if label == "Saldo final" else "Regular", 8.5)
            _right(page, 320, y, _money(value), font, regular, "Regular", 8.5)
            y += 20
    _text(page, 42, 300, "ACTIVIDAD", bold, "Bold", 10, theme.primary)
    _text(page, 42, 322, f"{int((dep > 0).sum())} ingresos por {_money(dep.sum())}", regular, "Regular", 9)
    _text(page, 42, 342, f"{int((ret > 0).sum())} egresos por {_money(ret.sum())}", regular, "Regular", 9)
    _text(page, 42, 382, "Identidad, cuentas, referencias y operaciones generadas exclusivamente para clase.", bold, "Bold", 8.5, theme.primary)


def generate_themed_pdf_bytes(df: pd.DataFrame, identity: dict, period_start, period_end, style: str) -> bytes:
    theme = THEMES.get(style, THEMES["GENERICO"])
    regular, bold = _font_path("MYRIADPRO-COND.OTF"), _font_path("MYRIADPRO-BOLDCOND.OTF")
    font = fitz.Font(fontfile=regular)
    doc = fitz.open()
    page = doc.new_page(width=PAGE_W, height=PAGE_H)
    _page_frame(page, theme, regular, bold, 1)
    _summary(page, theme, identity, period_start, period_end, df, regular, bold, font)

    page_number, page, y = 2, None, 0
    for row_index, (_, row) in enumerate(df.iterrows()):
        lines = _wrap(row.get("DESCRIPCION", ""), 260, font, 7.5)
        height = max(18, 10 * len(lines) + 5)
        if page is None or y + height > 730:
            page = doc.new_page(width=PAGE_W, height=PAGE_H)
            _page_frame(page, theme, regular, bold, page_number)
            page.draw_rect((40, 59, 570, 80), color=theme.dark, fill=theme.dark)
            for x, label in [(46, "FECHA"), (94, "CONCEPTO"), (369, "DEPOSITO"), (438, "RETIRO"), (511, "SALDO")]:
                _text(page, x, 73, label, bold, "Bold", 7.2, (1, 1, 1))
            y, page_number = 96, page_number + 1
        if row_index % 2 == 0:
            page.draw_rect((40, y - 11, 570, y + height - 10), color=theme.secondary, fill=theme.secondary)
        raw_date = row.get("FECHA")
        date_text = "" if pd.isna(raw_date) else pd.to_datetime(raw_date).strftime("%d/%m/%Y")
        _text(page, 46, y, date_text, regular, "Regular", 7.3)
        for i, line in enumerate(lines):
            _text(page, 94, y + i * 10, line, regular, "Regular", 7.3)
        deposit, withdrawal, balance = row.get("DEPOSITO"), row.get("RETIRO"), row.get("SALDO")
        _right(page, 428, y, _money(deposit) if not pd.isna(deposit) and float(deposit) else "", font, regular, "Regular", 7.3)
        _right(page, 495, y, _money(withdrawal) if not pd.isna(withdrawal) and float(withdrawal) else "", font, regular, "Regular", 7.3)
        _right(page, 566, y, "" if pd.isna(balance) else _money(balance), font, regular, "Regular", 7.3)
        y += height
    return doc.tobytes(deflate=True, garbage=4)
