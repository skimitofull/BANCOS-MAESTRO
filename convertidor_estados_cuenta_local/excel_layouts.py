"""Modalidades de exportacion de movimientos a Excel.

La columna tecnica ``_LINEAS_DOCUMENTO`` conserva, cuando el extractor puede
obtenerlos, los renglones fisicos del PDF. Nunca se muestra en el Excel final.
"""

from __future__ import annotations

import json
import re
from textwrap import wrap

import pandas as pd


PHYSICAL_LINES_COLUMN = "_LINEAS_DOCUMENTO"
OUTPUT_MODE_SINGLE = "Un movimiento por fila (concepto completo en una celda)"
OUTPUT_MODE_PHYSICAL = "Un renglón físico del PDF por fila"

FIELD_SPLIT_RE = re.compile(
    r"\s+(?=(?:REF|RASTREO|CUENTA|EMISOR|RECEPTOR|ORDENANTE|BENEFICIARIO|CONCEPTO|HORA|CLAVE|FOLIO|AUT|SUC|CAJA)\s*:?)",
    re.IGNORECASE,
)


def encode_physical_lines(lines) -> str:
    cleaned = [re.sub(r"\s+", " ", str(line or "")).strip() for line in lines]
    return json.dumps([line for line in cleaned if line], ensure_ascii=False)


def _fallback_lines(description: str, bank: str = "GENÉRICO") -> list[str]:
    """Separa campos semanticos y limita renglones largos de datos generados."""
    text = re.sub(r"\s+", " ", str(description or "")).strip()
    if not text:
        return [""]
    parts = [part.strip() for part in FIELD_SPLIT_RE.split(text) if part.strip()]
    if not parts:
        parts = [text]
    widths = {
        "SCOTIABANK": 34,
        "BANAMEX": 48,
        "BANAMEX MICUENTA": 48,
        "CITIBANAMEX": 48,
        "BBVA": 52,
        "BBVA 1": 52,
        "BBVA 2": 52,
        "SANTANDER": 48,
        "AZTECA": 45,
        "HSBC": 52,
        "BANORTE": 58,
    }
    width = widths.get(str(bank).upper(), 56)
    output = []
    for part in parts:
        output.extend(wrap(part, width=width, break_long_words=True, break_on_hyphens=False) or [part])
    return output or [text]


def derive_physical_lines(description: str, bank: str = "GENÉRICO") -> list[str]:
    """Expone la separacion usada cuando el renglon no proviene de un PDF."""
    return _fallback_lines(description, bank)


def decode_physical_lines(row, bank: str = "GENÉRICO") -> list[str]:
    value = row.get(PHYSICAL_LINES_COLUMN) if hasattr(row, "get") else None
    if value is not None and not (isinstance(value, float) and pd.isna(value)):
        try:
            decoded = json.loads(str(value))
            if isinstance(decoded, list):
                cleaned = [re.sub(r"\s+", " ", str(item or "")).strip() for item in decoded]
                cleaned = [item for item in cleaned if item]
                if cleaned:
                    return cleaned
        except (TypeError, ValueError, json.JSONDecodeError):
            pass
    return _fallback_lines(row.get("DESCRIPCION", ""), bank)


def compact_excel_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Elimina columnas tecnicas y conserva un movimiento por fila."""
    return df.drop(columns=[PHYSICAL_LINES_COLUMN, "_LINEAS_ORIGEN"], errors="ignore").copy()


def expand_document_rows(df: pd.DataFrame, bank: str = "GENÉRICO") -> pd.DataFrame:
    """Convierte cada renglón físico conservado en una fila independiente."""
    if df.empty:
        return compact_excel_rows(df)
    source_columns = [column for column in df.columns if column != PHYSICAL_LINES_COLUMN and not (str(bank).upper() == "SANTANDER SELECT" and column == "_LINEAS_ORIGEN")]
    rows = []
    for movement_number, (_, source) in enumerate(df.iterrows(), start=1):
        lines = decode_physical_lines(source, bank)
        for line_number, line in enumerate(lines, start=1):
            new_row = {column: source.get(column, pd.NA) for column in source_columns}
            new_row["DESCRIPCION"] = line
            new_row["MOVIMIENTO"] = movement_number
            new_row["RENGLON_MOVIMIENTO"] = line_number
            new_row["ES_CONTINUACION"] = "NO" if line_number == 1 else "SI"
            if str(bank).upper() == "SANTANDER SELECT":
                try:
                    origins = json.loads(str(source.get("_LINEAS_ORIGEN", "[]")))
                    if line_number <= len(origins):
                        new_row["HOJA_ORIGINAL"] = origins[line_number-1]["page"]
                        new_row["Y_ORIGINAL"] = origins[line_number-1]["y"]
                except (ValueError, TypeError, KeyError):
                    pass
                new_row.pop("_LINEAS_ORIGEN", None)
            if line_number > 1:
                if str(bank).upper() == "SANTANDER SELECT": new_row["FOLIO"] = pd.NA
                new_row["FECHA"] = pd.NaT
                new_row["DEPOSITO"] = pd.NA
                new_row["RETIRO"] = pd.NA
                new_row["SALDO"] = pd.NA
                if "SALDO_ORIGINAL" in new_row:
                    new_row["SALDO_ORIGINAL"] = pd.NA
            rows.append(new_row)
    output_columns = source_columns + ["MOVIMIENTO", "RENGLON_MOVIMIENTO", "ES_CONTINUACION"]
    return pd.DataFrame(rows, columns=output_columns)


def prepare_excel_layout(df: pd.DataFrame, mode: str, bank: str = "GENÉRICO") -> pd.DataFrame:
    if mode == OUTPUT_MODE_PHYSICAL:
        return expand_document_rows(df, bank)
    return compact_excel_rows(df)
