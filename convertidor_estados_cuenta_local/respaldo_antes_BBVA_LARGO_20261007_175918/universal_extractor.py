"""Extractor bancario reutilizable PDF -> movimientos.

El resultado semantico conserva un movimiento por fila y una columna tecnica
con los renglones fisicos detectados. La interfaz decide despues cual de las
dos representaciones exportar a Excel.
"""

from __future__ import annotations

import re
from datetime import datetime

import fitz
import pandas as pd

from excel_layouts import PHYSICAL_LINES_COLUMN, encode_physical_lines
from banamex_micuenta_parser import (
    BANK_ID as BANAMEX_MICUENTA,
    JUN_JUL_BANK_ID as BANAMEX_MICUENTA_JUN_JUL,
    extract_banamex_micuenta_movements,
    is_banamex_jun_jul_layout,
    is_banamex_micuenta,
)
from scotiabank_parser import extract_scotiabank_movements


MONTH_TO_NUM = {
    "ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DIC": 12,
    "JAN": 1, "APR": 4, "AUG": 8, "DEC": 12,
}
FULL_MONTH_TO_NUM = {
    "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
    "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12,
}
DATE_RE = re.compile(r"^\s*(\d{1,2})[-/ ]([A-ZÁÉÍÓÚ]{3,}|\d{1,2})[-/ ](\d{2,4})(?=\s|[A-Z0-9]|$)", re.I)
SHORT_DATE_RE = re.compile(r"^\s*(\d{1,2})[-/ ]([A-ZÁÉÍÓÚ]{3})(?=\s|$)", re.I)
MONEY_RE = re.compile(r"(?<![A-Z0-9])(?:\$\s*)?(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2}(?![A-Z0-9])")
WITHDRAWAL_WORDS = [
    "RETIRO", "CARGO", "PAGO", "CHEQUE PAGADO", "COMISION", "COMISIÓN",
    "IVA", "INTERESES", "PENALIZACION", "PENALIZACIÓN", "CAPITAL", "SEGUROS",
    "ADMINISTRACION", "ADMINISTRACIÓN", "SPEI ENVIADO", "TRASPASO A",
    "TRANSFERENCIA ENVIADA", "DOMICILIACION", "DOMICILIACIÓN", "DISPOSICION", "DISPOSICIÓN",
]


def decode_pdf_font_text(value) -> str:
    """Recupera texto HSBC embebido como bytes EBCDIC CP037."""
    text = str(value or "")
    compact = [character for character in text if not character.isspace()]
    unusual = sum(
        ord(character) >= 128 or character in "ðñòóôõö÷øù"
        for character in compact
    )
    suspicious_ratio = unusual / len(compact) if compact else 0.0
    has_ebcdic_spaces = "@" in text and any(marker in text for marker in ("Ã", "ä", "Å", "Õ", "Ù", "×", "[@"))
    if suspicious_ratio < 0.18 and not has_ebcdic_spaces:
        return text
    try:
        decoded = text.encode("latin-1").decode("cp037")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return text
    printable = sum(character.isprintable() or character in "\r\n\t" for character in decoded)
    return decoded if decoded and printable / len(decoded) > 0.90 else text


def clean_text(value) -> str:
    if value is None:
        return ""
    text = decode_pdf_font_text(value).replace("\r", " ").replace("\n", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return "" if text.lower() in {"nan", "none", "null"} else text


def amount_to_float(value) -> float:
    text = str(value or "").replace("$", "").replace(",", "").replace(" ", "").strip()
    try:
        return float(text)
    except (TypeError, ValueError):
        return 0.0


def first_pdf_text(pdf_bytes, pages=2) -> str:
    with fitz.open(stream=pdf_bytes, filetype="pdf") as document:
        chunks = []
        for index in range(min(pages, len(document))):
            chunks.extend(decode_pdf_font_text(line) for line in (document[index].get_text("text") or "").splitlines())
        return "\n".join(chunks)


def detect_bank(pdf_bytes) -> str:
    if is_banamex_micuenta(pdf_bytes):
        if is_banamex_jun_jul_layout(pdf_bytes):
            return BANAMEX_MICUENTA_JUN_JUL
        return BANAMEX_MICUENTA
    text = first_pdf_text(pdf_bytes, pages=2).upper()
    top = text[:6000]
    issuer = text[:2000]
    if "CUENTA SCOTIA" in issuer or "SCOTIABANK INVERLAT" in issuer:
        return "SCOTIABANK"
    if "EMITIDO POR: HSBC" in issuer or "HSBC MÉXICO" in issuer or "HSBC MEXICO" in issuer:
        return "HSBC"
    if "SANTANDER" in issuer:
        return "SANTANDER"
    if "BANCO AZTECA" in issuer or "GUARDADITO" in issuer:
        return "AZTECA"
    if "BBVA MEXICO, S.A." in text or "GRUPO FINANCIERO BBVA MEXICO" in text:
        return "BBVA"
    if "CITIBANAMEX" in issuer or "MICUENTA" in issuer:
        return "CITIBANAMEX"
    if "ENLACE NEGOCIOS" in issuer or "BANORTE" in issuer:
        return "BANORTE"
    if "BANAMEX" in issuer:
        return "CITIBANAMEX"
    if "BANCOMER" in issuer or "BBVA" in issuer:
        return "BBVA"
    if "SCOTIABANK" in top:
        return "SCOTIABANK"
    if "HSBC" in top:
        return "HSBC"
    return "GENÉRICO"


def _period_date(value: str):
    text = clean_text(value).upper().replace(".", "")
    numeric = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{2,4})", text)
    if numeric:
        day, month, year = map(int, numeric.groups())
        year = year + 2000 if year < 100 else year
        return pd.Timestamp(year=year, month=month, day=day)
    named = re.fullmatch(r"(\d{1,2})[-/ ]([A-Z]{3})[-/ ](\d{2,4})", text)
    if named and named.group(2) in MONTH_TO_NUM:
        year = int(named.group(3))
        year = year + 2000 if year < 100 else year
        return pd.Timestamp(year=year, month=MONTH_TO_NUM[named.group(2)], day=int(named.group(1)))
    return None


def extract_statement_period(pdf_bytes):
    """Obtiene las fechas completas impresas en la caratula, cuando existen."""
    text = first_pdf_text(pdf_bytes, pages=4).upper()
    token = r"\d{1,2}(?:/\d{1,2}/|[-/ ][A-Z]{3}[-/ ])\d{2,4}"
    patterns = [
        rf"(?:PERIODO|PERÍODO)\s*({token})\s*(?:AL|A|-|/)\s*({token})",
        rf"DEL\s+({token})\s+AL\s+({token})",
        rf"({token})\s+AL\s+({token})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.I)
        if match:
            start, end = _period_date(match.group(1)), _period_date(match.group(2))
            if start is not None and end is not None and start <= end:
                return start.normalize(), end.normalize()
    month_words = "|".join(FULL_MONTH_TO_NUM)
    long_match = re.search(
        rf"DEL\s+(\d{{1,2}})\s+DE\s+({month_words})(?:\s+(?:(?:DEL|DE)\s+)?(20\d{{2}}))?"
        rf"\s+AL\s+(\d{{1,2}})\s+DE\s+({month_words})(?:\s+(?:(?:DEL|DE)\s+)?(20\d{{2}}))?",
        text,
        re.I,
    )
    if long_match:
        first_year = long_match.group(3)
        second_year = long_match.group(6)
        year_end = int(second_year or first_year or datetime.now().year)
        year_start = int(first_year or second_year or year_end)
        start = pd.Timestamp(year=year_start, month=FULL_MONTH_TO_NUM[long_match.group(2).upper()], day=int(long_match.group(1)))
        end = pd.Timestamp(year=year_end, month=FULL_MONTH_TO_NUM[long_match.group(5).upper()], day=int(long_match.group(4)))
        if start <= end:
            return start.normalize(), end.normalize()
    return None, None


def extract_lines_from_pdf(pdf_bytes):
    lines = []
    with fitz.open(stream=pdf_bytes, filetype="pdf") as document:
        for page_number, page in enumerate(document, start=1):
            for raw in (page.get_text("text") or "").splitlines():
                line = clean_text(raw)
                if line:
                    lines.append((page_number, line))
    return lines


def normalize_short_date_lines(lines, pdf_bytes):
    text = first_pdf_text(pdf_bytes, pages=3)
    years = re.findall(r"\b20\d{2}\b", text)
    year = max(set(years), key=years.count) if years else str(datetime.now().year)
    output = []
    for page_number, line in lines:
        if not DATE_RE.search(line):
            match = SHORT_DATE_RE.search(line)
            if match:
                replacement = f"{int(match.group(1)):02d}-{match.group(2).upper()}-{year[-2:]}"
                line = replacement + line[match.end():]
        output.append((page_number, line))
    return output


def _skip_noise(line):
    upper = line.upper()
    noise = [
        "DETALLE DE MOVIMIENTOS", "DETALLE DE OPERACIONES", "FECHA DESCRIP", "DESCRIPCIÓN / ESTABLECIMIENTO",
        "DESCRIPCION / ESTABLECIMIENTO", "MONTO DEL DEPOSITO", "MONTO DEL DEPÓSITO", "MONTO DEL RETIRO",
        "CONTINUACIÓN", "CONTINUACION", "ENLACE NEGOCIOS", "PÁGINA", "PAGINA", "TOTAL DEPÓSITOS",
        "TOTAL DEPOSITOS", "TOTAL RETIROS",
    ]
    if upper.strip() in {"FECHA", "DESCRIPCION", "DESCRIPCIÓN", "SALDO", "PESOS", "RETIROS", "DEPOSITOS", "DEPÓSITOS"}:
        return True
    return any(marker in upper for marker in noise) and not DATE_RE.search(upper)


def _parse_movement_block(block):
    if not block:
        return None
    match = DATE_RE.search(block[0])
    if not match:
        return None
    month_text = match.group(2)[:3].upper()
    if month_text.isdigit():
        month_text = next((key for key, number in MONTH_TO_NUM.items() if number == int(month_text) and key not in {"JAN", "APR", "AUG", "DEC"}), month_text)
    date_text = f"{int(match.group(1)):02d}-{month_text}-{match.group(3)[-2:]}"
    first_description = block[0][match.end():].strip()
    physical_lines = [first_description] + list(block[1:])
    physical_lines = [clean_text(line) for line in physical_lines if clean_text(line)]
    text = " ".join(physical_lines)
    amounts = MONEY_RE.findall(text)
    deposit = withdrawal = balance = 0.0
    upper = text.upper()
    if amounts:
        numbers = [amount_to_float(value) for value in amounts]
        if "SALDO ANTERIOR" in upper:
            balance = numbers[-1]
            structured_count = 1
        else:
            balance = numbers[-1] if len(numbers) >= 2 else 0.0
            movement = numbers[-2] if len(numbers) >= 2 else numbers[-1]
            is_withdrawal = any(word in upper for word in WITHDRAWAL_WORDS) and not any(word in upper for word in ("SPEI RECIBIDO", "DEPOSITO", "DEPÓSITO", "ABONO"))
            if is_withdrawal:
                withdrawal = movement
            else:
                deposit = movement
            structured_count = min(3, len(amounts))
        for amount in reversed(amounts[-structured_count:]):
            position = text.rfind(amount)
            if position >= 0:
                text = (text[:position] + text[position + len(amount):]).strip()
        text = re.sub(r"\s+", " ", text).strip(" ,")

    concept_lines = []
    for physical in physical_lines:
        if re.fullmatch(r"(?:\$?\s*[\d,]+\.\d{2}\s*)+", physical):
            continue
        concept_lines.append(physical)
    return {
        "FECHA": date_text,
        "DESCRIPCION": text,
        "DEPOSITO": deposit,
        "RETIRO": withdrawal,
        "SALDO": balance,
        PHYSICAL_LINES_COLUMN: encode_physical_lines(concept_lines or [text]),
    }


def _lines_to_movements(lines):
    blocks = []
    current = []
    current_page = None
    for page_number, line in lines:
        if current_page is not None and page_number != current_page:
            if current:
                blocks.append(current)
            current = []
        current_page = page_number
        if _skip_noise(line):
            if current:
                blocks.append(current)
                current = []
            continue
        if DATE_RE.search(line):
            if current:
                blocks.append(current)
            current = [line]
        elif current:
            current.append(line)
    if current:
        blocks.append(current)
    rows = [row for row in (_parse_movement_block(block) for block in blocks) if row]
    return pd.DataFrame(rows, columns=["FECHA", "DESCRIPCION", "DEPOSITO", "RETIRO", "SALDO", PHYSICAL_LINES_COLUMN])


def extract_banorte_account_lines(pdf_bytes):
    output = []
    in_movements = False
    in_account = False
    after_header = False
    stop_markers = [
        "INVERSIÓN VISTA PREFERENTE", "INVERSION VISTA PREFERENTE", "OTROS▼", "CARGOS OBJETADOS",
        "INFORME DE DEPÓSITOS", "INFORME DE DEPOSITOS", "GANANCIA ANUAL TOTAL", "COMPROBANTE FISCAL",
    ]
    with fitz.open(stream=pdf_bytes, filetype="pdf") as document:
        for page_number, page in enumerate(document, start=1):
            for raw in (page.get_text("text") or "").splitlines():
                line = clean_text(raw)
                if not line:
                    continue
                upper = line.upper()
                if "DETALLE DE MOVIMIENTOS" in upper:
                    in_movements, in_account, after_header = True, False, False
                    continue
                if not in_movements:
                    continue
                if any(marker in upper for marker in stop_markers):
                    if in_account:
                        return output
                    continue
                if upper in {"CUENTA PREFERENTE", "ENLACE NEGOCIOS BASICA", "ENLACE NEGOCIOS BÁSICA"}:
                    in_account = True
                    continue
                if in_account and ("FECHA" in upper or "DESCRIP" in upper or "MONTO DEL" in upper or upper == "SALDO"):
                    continue
                if in_account and DATE_RE.search(line):
                    after_header = True
                if in_account and after_header and not upper.startswith("SIN MOVIMIENTOS"):
                    output.append((page_number, line))
    return output


def extract_hsbc_movements(pdf_bytes):
    """Extrae HSBC aun cuando el PDF mezcla texto normal y fuente EBCDIC."""
    lines = extract_lines_from_pdf(pdf_bytes)
    joined = "\n".join(line for _, line in lines)
    period_dates = re.findall(r"\b(\d{2})/(\d{2})/(20\d{2})\b", joined)
    if period_dates:
        _, month_text, year_text = period_dates[0]
        month, year = int(month_text), int(year_text)
    else:
        month, year = datetime.now().month, datetime.now().year

    first_detail = next((index for index, (_, line) in enumerate(lines) if "DETALLE MOVIMIENTOS" in line.upper()), len(lines))
    opening = 0.0
    for _, line in lines[:first_detail]:
        match = MONEY_RE.search(line)
        if match:
            opening = amount_to_float(match.group(0))
            break

    blocks = []
    in_detail = False
    current = None
    money_count = 0
    for _, line in lines:
        upper = line.upper()
        if "DETALLE MOVIMIENTOS" in upper:
            in_detail = True
            current = None
            money_count = 0
            continue
        if not in_detail:
            continue
        if re.fullmatch(r"(?:0[1-9]|[12]\d|3[01])", line):
            current = [line]
            money_count = 0
            continue
        if current is None:
            continue
        current.append(line)
        if MONEY_RE.fullmatch(line.strip()):
            money_count += 1
            if money_count >= 2:
                blocks.append(current)
                current = None
                money_count = 0

    rows = []
    running = round(float(opening), 2)
    for block_index, block in enumerate(blocks, start=1):
        day = int(block[0])
        amounts = [amount_to_float(match.group(0)) for line in block[1:] for match in MONEY_RE.finditer(line)]
        if len(amounts) < 2:
            continue
        movement_amount, balance = amounts[-2], amounts[-1]
        physical = [
            clean_text(line) for line in block[1:]
            if clean_text(line) and not MONEY_RE.fullmatch(clean_text(line))
        ]
        description = " ".join(physical).strip() or f"MOVIMIENTO HSBC {block_index:04d}"
        delta = round(balance - running, 2)
        tolerance = max(0.02, movement_amount * 0.0001)
        if abs(delta - movement_amount) <= tolerance:
            deposit, withdrawal = movement_amount, 0.0
        elif abs(delta + movement_amount) <= tolerance:
            deposit, withdrawal = 0.0, movement_amount
        else:
            upper = description.upper()
            is_withdrawal = any(word in upper for word in WITHDRAWAL_WORDS) and not any(word in upper for word in ("ABONO", "DEPOSITO", "SPEI RECIBIDO"))
            deposit, withdrawal = (0.0, movement_amount) if is_withdrawal else (movement_amount, 0.0)
        rows.append({
            "FECHA": f"{day:02d}-{next((name for name, number in MONTH_TO_NUM.items() if number == month and name not in {'JAN', 'APR', 'AUG', 'DEC'}), month):s}-{str(year)[-2:]}",
            "DESCRIPCION": description,
            "DEPOSITO": deposit,
            "RETIRO": withdrawal,
            "SALDO": balance,
            PHYSICAL_LINES_COLUMN: encode_physical_lines(physical or [description]),
        })
        running = round(balance, 2)

    if not rows:
        return pd.DataFrame(columns=["FECHA", "DESCRIPCION", "DEPOSITO", "RETIRO", "SALDO", PHYSICAL_LINES_COLUMN])
    opening_row = {
        "FECHA": rows[0]["FECHA"],
        "DESCRIPCION": "SALDO ANTERIOR",
        "DEPOSITO": 0.0,
        "RETIRO": 0.0,
        "SALDO": opening,
        PHYSICAL_LINES_COLUMN: encode_physical_lines(["SALDO ANTERIOR"]),
    }
    return pd.DataFrame([opening_row, *rows])


def parse_pdf_movements(pdf_bytes, bank="AUTO"):
    selected_bank = detect_bank(pdf_bytes) if bank == "AUTO" else bank
    period_start, period_end = extract_statement_period(pdf_bytes)
    if selected_bank == "SCOTIABANK":
        dataframe = extract_scotiabank_movements(pdf_bytes)
    elif selected_bank == "HSBC":
        dataframe = extract_hsbc_movements(pdf_bytes)
    elif selected_bank in {BANAMEX_MICUENTA, BANAMEX_MICUENTA_JUN_JUL} or (selected_bank == "CITIBANAMEX" and is_banamex_micuenta(pdf_bytes)):
        if selected_bank == "CITIBANAMEX":
            selected_bank = BANAMEX_MICUENTA_JUN_JUL if is_banamex_jun_jul_layout(pdf_bytes) else BANAMEX_MICUENTA
        try:
            dataframe = extract_banamex_micuenta_movements(pdf_bytes, period_start, period_end)
        except ValueError as coordinate_error:
            # MiCuenta tiene generaciones históricas con tablas auxiliares
            # diferentes. El formato nuevo usa conciliación estricta; las
            # variantes anteriores conservan el lector existente sin impedir
            # que el usuario las abra.
            lines = normalize_short_date_lines(extract_lines_from_pdf(pdf_bytes), pdf_bytes)
            dataframe = _lines_to_movements(lines)
            dataframe.attrs["coordinate_parser_warning"] = str(coordinate_error)
    else:
        lines = extract_banorte_account_lines(pdf_bytes) if selected_bank == "BANORTE" else normalize_short_date_lines(extract_lines_from_pdf(pdf_bytes), pdf_bytes)
        dataframe = _lines_to_movements(lines)
    if dataframe.empty:
        return dataframe, selected_bank
    description = dataframe["DESCRIPCION"].astype(str).str.strip()
    deposit = pd.to_numeric(dataframe["DEPOSITO"], errors="coerce").fillna(0.0)
    withdrawal = pd.to_numeric(dataframe["RETIRO"], errors="coerce").fillna(0.0)
    dataframe = dataframe[~((description == "") & (deposit.abs() < 0.005) & (withdrawal.abs() < 0.005))].copy()
    suspicious = dataframe["DESCRIPCION"].astype(str).str.upper().apply(
        lambda value: (
            ("ESTADO DE CUENTA" in value and "NO. DE CUENTA" in value)
            or value.startswith("RESUMEN INFORMATIVO")
        )
    )
    dataframe = dataframe[~suspicious].copy().reset_index(drop=True)
    if period_start is not None:
        dataframe.attrs["period_start"] = period_start
    if period_end is not None:
        dataframe.attrs["period_end"] = period_end
    dataframe.attrs["bank"] = selected_bank
    return dataframe, selected_bank
