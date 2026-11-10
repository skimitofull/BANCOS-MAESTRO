"""Extractor por coordenadas para estados Banamex MiCuenta.

La tabla MiCuenta separa internamente el día y el mes, y puede continuar un
movimiento en la página siguiente. Por eso no debe procesarse como texto lineal.
"""

from __future__ import annotations

import re
from datetime import datetime

import fitz
import pandas as pd

from excel_layouts import PHYSICAL_LINES_COLUMN, encode_physical_lines


BANK_ID = "BANAMEX MICUENTA"
JUN_JUL_BANK_ID = "BANAMEX MICUENTA (JUN-JUL)"
MONTHS = {
    "ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DIC": 12,
}


def is_banamex_jun_jul_layout(pdf_bytes: bytes) -> bool:
    """Detecta la retícula MiCuenta de 29 filas iniciales y 49 posteriores."""
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as document:
            for page in document:
                text = (page.get_text("text") or "").upper()
                if "DETALLE DE OPERACIONES" not in text or "CONCEPTO" not in text:
                    continue
                is_first_movement_page = "SALDO AL CORTE" in text
                minimum_y = 380.0 if is_first_movement_page else 145.0
                maximum_y = 730.0
                horizontal = set()
                for drawing in page.get_drawings():
                    for item in drawing.get("items", []):
                        if item[0] != "l":
                            continue
                        start, end = item[1], item[2]
                        if abs(float(start.y) - float(end.y)) > 0.2:
                            continue
                        if abs(float(end.x) - float(start.x)) < 400:
                            continue
                        y = round(float(start.y), 2)
                        if minimum_y <= y <= maximum_y:
                            horizontal.add(y)
                physical_rows = max(0, len(horizontal) - 1)
                return physical_rows == 29 if is_first_movement_page else physical_rows == 49
    except Exception:
        return False
    return False
MONEY_PATTERN = r"(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2}"


def is_banamex_micuenta(pdf_bytes: bytes) -> bool:
    with fitz.open(stream=pdf_bytes, filetype="pdf") as document:
        text = "\n".join((page.get_text("text") or "") for page in document[:2]).upper()
    return "MICUENTA" in text and ("BANAMEX" in text or "NÚMERO DE CONTRATO" in text or "N�MERO DE CONTRATO" in text)


def _clean(value) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    replacements = {
        "COMPA�IA": "COMPAÑIA",
        "INSTITUCI�N": "INSTITUCION",
        "M�XICO": "MEXICO",
    }
    for old, new in replacements.items():
        text = text.replace(old, new).replace(old.lower(), new.lower())
    text = re.sub(r"\bCOMPA.IA\b", "COMPAÑIA", text, flags=re.I)
    text = re.sub(r"\bEX\s+ENTAR\b", "EXENTAR", text, flags=re.I)
    text = re.sub(r"\bCOM\s+ISION\b", "COMISION", text, flags=re.I)
    return text


def _amount(value) -> float:
    try:
        return float(str(value).replace("$", "").replace(",", "").strip())
    except (TypeError, ValueError):
        return 0.0


def _summary_value(text: str, label_pattern: str):
    match = re.search(rf"{label_pattern}\s*\n\s*\$?({MONEY_PATTERN})", text, re.I)
    return _amount(match.group(1)) if match else None


def _statement_summary(document) -> dict:
    page_texts = [(document[index].get_text("text") or "") for index in range(min(3, len(document)))]
    text = "\n".join(page_texts)

    # Algunos estados MiCuenta traen una caratula resumida y, en la pagina
    # siguiente, el resumen contable que precede al detalle. Hemos encontrado
    # documentos en los que la caratula difiere por una bonificacion, mientras
    # que el segundo resumen si reconcilia con cada movimiento y con el saldo
    # final. Usar primero la pagina del detalle evita descartar el extractor
    # por coordenadas y caer al lector lineal, que puede fusionar movimientos.
    detail_text = next(
        (
            page_text for page_text in page_texts
            if re.search(r"Retiros/Otros cargos", page_text, re.I)
            and re.search(r"Detalle de Operaciones", page_text, re.I)
        ),
        "",
    )

    def preferred_value(label_pattern: str):
        return _summary_value(detail_text, label_pattern) if detail_text else _summary_value(text, label_pattern)

    summary = {
        "SALDO_INICIAL_REPORTADO": preferred_value(r"Saldo anterior"),
        "INGRESOS_REPORTADOS": preferred_value(r"Dep.sitos"),
        "EGRESOS_REPORTADOS": preferred_value(r"Otros cargos"),
        "SALDO_FINAL_REPORTADO": preferred_value(r"Saldo al [Cc]orte"),
        "SALDO_PROMEDIO_MENSUAL_REPORTADO": preferred_value(r"Saldo promedio mensual"),
    }
    count_text = detail_text or (document[1].get_text("text") if len(document) > 1 else text)
    deposits = re.search(r"\(\+\)\s*\n\s*(\d+)\s*\n\s*Dep.sitos", count_text, re.I)
    withdrawals = re.search(r"\(-\)\s*\n\s*(\d+)\s*\n\s*Retiros/Otros cargos", count_text, re.I)
    summary["INGRESOS_CANTIDAD_REPORTADA"] = int(deposits.group(1)) if deposits else None
    summary["EGRESOS_CANTIDAD_REPORTADA"] = int(withdrawals.group(1)) if withdrawals else None
    return summary


def _word_records(page):
    return [
        {"x0": float(word[0]), "y0": float(word[1]), "x1": float(word[2]), "y1": float(word[3]), "text": _clean(word[4])}
        for word in page.get_text("words", sort=True)
        if _clean(word[4])
    ]


def _date_starts(words, page_number):
    groups = []
    for word in words:
        if not (20 <= word["x0"] < 70 and 100 <= word["y0"] < 735):
            continue
        group = next((item for item in groups if abs(item[0] - word["y0"]) <= 2.5), None)
        if group is None:
            group = [word["y0"], []]
            groups.append(group)
        group[1].append(word)
    starts = []
    for y0, items in groups:
        tokens = [item["text"].upper() for item in sorted(items, key=lambda item: item["x0"])]
        day = next((token for token in tokens if re.fullmatch(r"\d{1,2}", token)), None)
        month = next((token[:3] for token in tokens if token[:3] in MONTHS), None)
        if day and month:
            starts.append({"page": page_number, "y0": float(y0), "day": int(day), "month": month})
    return sorted(starts, key=lambda item: item["y0"])


def _is_movement_table_page(words):
    """Evita confundir fechas de anexos posteriores con movimientos."""
    has_date = any(word["text"].upper() == "FECHA" and 20 <= word["x0"] < 70 for word in words)
    has_concept = any(word["text"].upper() == "CONCEPTO" and 75 <= word["x0"] < 150 for word in words)
    has_withdrawal = any(word["text"].upper().startswith("RETIRO") and 330 <= word["x0"] < 415 for word in words)
    has_balance = any(word["text"].upper() == "SALDO" and 505 <= word["x0"] < 580 for word in words)
    return has_date and has_concept and has_withdrawal and has_balance


def _date_in_period(day, month, period_start, period_end):
    years = sorted({period_start.year, period_end.year}) if period_start is not None and period_end is not None else [datetime.now().year]
    candidates = [pd.Timestamp(year=year, month=MONTHS[month], day=day) for year in years]
    if period_start is not None and period_end is not None:
        valid = [value for value in candidates if period_start.normalize() <= value <= period_end.normalize()]
        if valid:
            return valid[0]
    return min(candidates, key=lambda value: abs((value - (period_end or value)).days))


def _segment_words(words_by_page, start, following):
    selected = []
    for page_number in range(start["page"], following["page"] + 1):
        lower = start["y0"] - 1.0 if page_number == start["page"] else 125.0
        upper = following["y0"] - 1.0 if page_number == following["page"] else 735.0
        page_words = words_by_page.get(page_number, [])

        # Las tablas auxiliares de MiCuenta ("Resumen Operaciones Chequera"
        # y "Resumen Operaciones Tarjeta") usan las mismas columnas monetarias
        # que el detalle. Si se dejan dentro del último segmento, sus totales
        # sustituyen el importe y el saldo del último movimiento y provocan que
        # el extractor estricto caiga al lector lineal. Ese lector lineal puede
        # agrupar varias fechas en un solo movimiento, que fue la causa de los
        # PDF con fechas e importes desplazados.
        line_tokens = []
        for word in sorted(page_words, key=lambda item: (item["y0"], item["x0"])):
            group = next((item for item in reversed(line_tokens) if abs(item[0] - word["y0"]) <= 2.5), None)
            if group is None:
                group = [word["y0"], []]
                line_tokens.append(group)
            group[1].append(word["text"].upper())
        summary_y = next(
            (
                y0 for y0, tokens in line_tokens
                if "RESUMEN" in tokens
                and "OPERACIONES" in tokens
                and any(token in {"CHEQUERA", "TARJETA"} for token in tokens)
            ),
            None,
        )
        if summary_y is not None and summary_y > lower:
            upper = min(upper, summary_y - 1.0)

        for word in page_words:
            if lower <= word["y0"] < upper and 20 <= word["x0"] < 580:
                selected.append({**word, "page": page_number})
    return selected


def _column_amount(words, left, right):
    candidates = [word for word in words if left <= word["x0"] < right and re.fullmatch(MONEY_PATTERN, word["text"])]
    return _amount(candidates[-1]["text"]) if candidates else 0.0


def _concept_lines(words):
    candidates = [word for word in words if 75 <= word["x0"] < 330 and 125 <= word["y0"] < 735]
    grouped = []
    for word in sorted(candidates, key=lambda item: (item["page"], item["y0"], item["x0"])):
        group = next(
            (item for item in reversed(grouped) if item[0] == word["page"] and abs(item[1] - word["y0"]) <= 2.0),
            None,
        )
        if group is None:
            group = [word["page"], word["y0"], []]
            grouped.append(group)
        group[2].append((word["x0"], word["text"]))
    noise = {"CONCEPTO", "DETALLE DE OPERACIONES", "EN PESOS MONEDA NACIONAL"}
    lines = [" ".join(text for _, text in sorted(parts)) for _, _, parts in grouped]
    return [line for line in lines if line.upper() not in noise]


def _assert_reconciled(dataframe: pd.DataFrame, summary: dict):
    deposits = round(float(pd.to_numeric(dataframe["DEPOSITO"], errors="coerce").fillna(0.0).sum()), 2)
    withdrawals = round(float(pd.to_numeric(dataframe["RETIRO"], errors="coerce").fillna(0.0).sum()), 2)
    deposit_count = int((pd.to_numeric(dataframe["DEPOSITO"], errors="coerce").fillna(0.0) > 0.005).sum())
    withdrawal_count = int((pd.to_numeric(dataframe["RETIRO"], errors="coerce").fillna(0.0) > 0.005).sum())
    checks = [
        ("ingresos", deposits, summary.get("INGRESOS_REPORTADOS")),
        ("egresos", withdrawals, summary.get("EGRESOS_REPORTADOS")),
        ("cantidad de ingresos", deposit_count, summary.get("INGRESOS_CANTIDAD_REPORTADA")),
        ("cantidad de egresos", withdrawal_count, summary.get("EGRESOS_CANTIDAD_REPORTADA")),
    ]
    failures = [(label, actual, expected) for label, actual, expected in checks if expected is not None and abs(actual - expected) > 0.005]
    balances = pd.to_numeric(dataframe["SALDO"], errors="coerce").dropna()
    reported_final = summary.get("SALDO_FINAL_REPORTADO")
    if reported_final is not None and len(balances) and abs(float(balances.iloc[-1]) - reported_final) > 0.005:
        failures.append(("saldo final", round(float(balances.iloc[-1]), 2), reported_final))
    if failures:
        detail = "; ".join(f"{label}: extraído {actual:,.2f}, reportado {expected:,.2f}" for label, actual, expected in failures)
        raise ValueError(f"BANAMEX MICUENTA no pasó la conciliación del estado de cuenta: {detail}.")


def extract_banamex_micuenta_movements(pdf_bytes: bytes, period_start=None, period_end=None) -> pd.DataFrame:
    with fitz.open(stream=pdf_bytes, filetype="pdf") as document:
        all_words = {index: _word_records(page) for index, page in enumerate(document)}
        words_by_page = {index: words for index, words in all_words.items() if _is_movement_table_page(words)}
        starts = []
        for page_number, words in words_by_page.items():
            starts.extend(_date_starts(words, page_number))
        starts.sort(key=lambda item: (item["page"], item["y0"]))
        if not starts:
            return pd.DataFrame(columns=["FECHA", "DESCRIPCION", "DEPOSITO", "RETIRO", "SALDO", PHYSICAL_LINES_COLUMN])
        summary = _statement_summary(document)
        terminal = {"page": max(words_by_page), "y0": 735.0}
        rows = []
        for index, start in enumerate(starts):
            following = starts[index + 1] if index + 1 < len(starts) else terminal
            words = _segment_words(words_by_page, start, following)
            lines = _concept_lines(words)
            if not lines:
                continue
            withdrawal = _column_amount(words, 335, 415)
            deposit = _column_amount(words, 415, 505)
            balance = _column_amount(words, 505, 580)
            rows.append({
                "FECHA": _date_in_period(start["day"], start["month"], period_start, period_end),
                "DESCRIPCION": " ".join(lines),
                "DEPOSITO": deposit,
                "RETIRO": withdrawal,
                "SALDO": balance,
                PHYSICAL_LINES_COLUMN: encode_physical_lines(lines),
            })
    dataframe = pd.DataFrame(rows)
    _assert_reconciled(dataframe, summary)
    dataframe.attrs["statement_summary"] = summary
    dataframe.attrs["bank"] = BANK_ID
    return dataframe
