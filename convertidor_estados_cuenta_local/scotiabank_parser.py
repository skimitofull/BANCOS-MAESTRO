"""Extractor reutilizable de movimientos Scotiabank PDF -> tabla.

Este modulo no depende de Streamlit y puede trasladarse despues al proyecto
independiente de conversion PDF a Excel.
"""

from __future__ import annotations

import re
from datetime import datetime

import fitz
import pandas as pd

from excel_layouts import PHYSICAL_LINES_COLUMN, encode_physical_lines


MONTHS_ES = {
    1: "ENE", 2: "FEB", 3: "MAR", 4: "ABR", 5: "MAY", 6: "JUN",
    7: "JUL", 8: "AGO", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DIC",
}
MONTH_TO_NUM = {
    "ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DIC": 12,
    "JAN": 1, "APR": 4, "AUG": 8, "DEC": 12,
}
MONEY_RE = re.compile(r"(?<![A-Z0-9])(?:\$\s*)?(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2}(?![A-Z0-9])")


def _clean_text(value) -> str:
    if value is None:
        return ""
    text = str(value).replace("\r", " ").replace("\n", " ").strip()
    return re.sub(r"\s+", " ", text)


def _amount(value) -> float:
    text = str(value or "").replace("$", "").replace(",", "").replace(" ", "").strip()
    try:
        return float(text)
    except (TypeError, ValueError):
        return 0.0


def extract_scotiabank_movements(pdf_bytes: bytes) -> pd.DataFrame:
    """Extrae la tabla por coordenadas y agrega su saldo anterior calculado."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    all_text = " ".join(page.get_text("text") or "" for page in doc)
    years = re.findall(r"\b20\d{2}\b", all_text)
    year = int(max(set(years), key=years.count)) if years else datetime.now().year
    rows = []

    for page in doc:
        spans = []
        for block in page.get_text("dict").get("blocks", []):
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    text = _clean_text(span.get("text", ""))
                    if text:
                        spans.append((text, span["bbox"]))

        page_text = " ".join(text for text, _ in spans).upper()
        if "DETALLE DE TUS MOVIMIENTOS" not in page_text:
            continue

        starts = []
        for text, bbox in spans:
            normalized = " ".join(text.upper().split())
            if re.fullmatch(r"\d{2}\s+[A-ZÁÉÍÓÚ]{3}", normalized) and 45 <= bbox[0] <= 88 and 74 <= bbox[1] <= 742:
                starts.append((bbox[1], normalized))
        starts.sort()

        for index, (top, date_text) in enumerate(starts):
            next_top = starts[index + 1][0] if index + 1 < len(starts) else top + 73.2
            bottom = min(next_top, top + 73.2, 742.0)
            in_row = [(text, bbox) for text, bbox in spans if top - 0.6 <= bbox[1] < bottom - 0.6]

            def contiguous_column(left, right):
                candidates = sorted((bbox[1], bbox[0], text) for text, bbox in in_row if left <= bbox[0] < right)
                selected = []
                last_y = None
                for item_y, item_x, item_text in candidates:
                    if last_y is not None and item_y - last_y > 11.0:
                        break
                    selected.append((item_y, item_x, item_text))
                    last_y = item_y
                return " ".join(text for _, _, text in selected)

            def physical_concept_lines():
                candidates = sorted(
                    (bbox[1], bbox[0], text)
                    for text, bbox in in_row
                    if 89 <= bbox[0] < 376
                )
                grouped = []
                for item_y, item_x, item_text in candidates:
                    if not grouped or abs(item_y - grouped[-1][0]) > 1.5:
                        grouped.append([item_y, [(item_x, item_text)]])
                    else:
                        grouped[-1][1].append((item_x, item_text))
                return [
                    " ".join(text for _, text in sorted(parts))
                    for _, parts in grouped
                    if parts
                ]

            concept = contiguous_column(89, 250)
            reference = contiguous_column(250, 376)
            deposit_text = " ".join(text for text, bbox in in_row if 376 <= bbox[0] < 440)
            withdrawal_text = " ".join(text for text, bbox in in_row if 440 <= bbox[0] < 512)
            balance_text = " ".join(text for text, bbox in in_row if 512 <= bbox[0] < 590)
            day_text, month_text = date_text.split()
            month = MONTH_TO_NUM.get(month_text[:3], 1)

            def column_amount(value):
                match = MONEY_RE.search(value)
                return _amount(match.group(0)) if match else 0.0

            description = _clean_text(concept + (f" REF {reference}" if reference else ""))
            rows.append({
                "FECHA": f"{int(day_text):02d}-{MONTHS_ES[month]}-{str(year)[-2:]}",
                "DESCRIPCION": description,
                "DEPOSITO": column_amount(deposit_text),
                "RETIRO": column_amount(withdrawal_text),
                "SALDO": column_amount(balance_text),
                PHYSICAL_LINES_COLUMN: encode_physical_lines(physical_concept_lines() or [description]),
            })

    doc.close()
    df = pd.DataFrame(rows, columns=["FECHA", "DESCRIPCION", "DEPOSITO", "RETIRO", "SALDO", PHYSICAL_LINES_COLUMN])
    if not df.empty:
        first = df.iloc[0]
        opening = round(float(first["SALDO"]) - float(first["DEPOSITO"]) + float(first["RETIRO"]), 2)
        opening_row = pd.DataFrame([{
            "FECHA": first["FECHA"], "DESCRIPCION": "SALDO ANTERIOR",
            "DEPOSITO": 0.0, "RETIRO": 0.0, "SALDO": opening,
            PHYSICAL_LINES_COLUMN: encode_physical_lines(["SALDO ANTERIOR"]),
        }])
        df = pd.concat([opening_row, df], ignore_index=True)
    return df
