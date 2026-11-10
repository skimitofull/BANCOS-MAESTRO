import io
import json
import importlib
import os
import re
from datetime import datetime

import fitz  # PyMuPDF
import numpy as np
import pandas as pd
import streamlit as st

import synthetic_engine as _synthetic_engine

# La aplicación puede permanecer abierta mientras se actualiza el motor.
# Recargarlo garantiza que la regla universal de variedad se aplique sin
# depender de cerrar manualmente Streamlit.
_synthetic_engine = importlib.reload(_synthetic_engine)
EXPENSE_CONCEPTS = _synthetic_engine.EXPENSE_CONCEPTS
INCOME_CONCEPTS = _synthetic_engine.INCOME_CONCEPTS
GenerationConfig = _synthetic_engine.GenerationConfig
expand_concepts = _synthetic_engine.expand_concepts
generate_identity = _synthetic_engine.generate_identity
generate_movements = _synthetic_engine.generate_movements
merge_base_and_generated = _synthetic_engine.merge_base_and_generated
movement_checks = _synthetic_engine.movement_checks
from multi_bank_templates import generate_themed_pdf_bytes
import movement_templates as _movement_templates
_movement_templates = importlib.reload(_movement_templates)
import santander_select as _santander_select
_santander_select = importlib.reload(_santander_select)
generate_movement_pdf_bytes = _movement_templates.generate_movement_pdf_bytes
from scotiabank_chart import generate_scotiabank_chart_png
import linked_periods as _linked_periods

# Streamlit conserva modulos importados cuando actualiza una aplicacion que ya
# estaba abierta. La recarga evita mezclar una interfaz nueva con un motor viejo.
_linked_periods = importlib.reload(_linked_periods)
LinkedPeriodConfig = _linked_periods.LinkedPeriodConfig
TARGET_AVERAGE = _linked_periods.TARGET_AVERAGE
TARGET_EXPENSES = _linked_periods.TARGET_EXPENSES
TARGET_FINAL = _linked_periods.TARGET_FINAL
AVERAGE_RELATION_AUTO = _linked_periods.AVERAGE_RELATION_AUTO
AVERAGE_RELATION_ABOVE = _linked_periods.AVERAGE_RELATION_ABOVE
AVERAGE_RELATION_BELOW = _linked_periods.AVERAGE_RELATION_BELOW
AVERAGE_RELATION_NEAR = _linked_periods.AVERAGE_RELATION_NEAR
AVERAGE_RELATIONS = _linked_periods.AVERAGE_RELATIONS
TIMING_EARLY = _linked_periods.TIMING_EARLY
TIMING_BALANCED = _linked_periods.TIMING_BALANCED
TIMING_LATE = _linked_periods.TIMING_LATE
TIMING_OPTIONS = _linked_periods.TIMING_OPTIONS
BANAMEX_TRANSFER_BANKS = _linked_periods.BANAMEX_TRANSFER_BANKS
link_statement_periods = _linked_periods.link_statement_periods
validate_linked_periods = _linked_periods.validate_linked_periods
from scotiabank_parser import extract_scotiabank_movements
from excel_layouts import (
    OUTPUT_MODE_PHYSICAL,
    OUTPUT_MODE_SINGLE,
    PHYSICAL_LINES_COLUMN,
    encode_physical_lines,
    prepare_excel_layout,
)
import banamex_micuenta_parser as _banamex_micuenta_parser
import universal_extractor as _universal_extractor

# La extracción Banamex es sensible a los límites de las tablas auxiliares.
# Recargar el módulo evita que una sesión de Streamlit ya abierta conserve el
# lector anterior después de actualizar el programa.
_banamex_micuenta_parser = importlib.reload(_banamex_micuenta_parser)
_universal_extractor = importlib.reload(_universal_extractor)
parse_pdf_movements = _universal_extractor.parse_pdf_movements


@st.cache_data(show_spinner=False)
def parse_pdf_movements_cached(pdf_bytes, bank_option):
    # BBVA LARGO: lector por columnas, versión 2.
    """Evita volver a extraer el mismo PDF en cada cambio de la simulacion guiada."""
    cache_revision = "banamex-micuenta-jun-jul-segmentos-fecha-20260924"
    _ = cache_revision
    return parse_pdf_movements(pdf_bytes, bank_option)


def build_variable_income_plan(monthly_average, variation_percent, period_count, seed):
    """Totales distintos, con centavos, dentro del rango y con promedio exacto."""
    monthly_average = round(float(monthly_average), 2)
    variation = max(0.0, float(variation_percent)) / 100.0
    period_count = max(1, int(period_count))
    if period_count == 1 or variation <= 0.0:
        return [monthly_average] * period_count
    rng = np.random.default_rng(int(seed) + 7919)
    deviations = rng.uniform(-1.0, 1.0, period_count)
    deviations = deviations - deviations.mean()
    largest = float(np.max(np.abs(deviations)))
    if largest > 0:
        deviations = deviations / largest * variation
    values = [round(monthly_average * (1.0 + float(value)), 2) for value in deviations]
    # Corrige centavos para que el promedio del conjunto sea exactamente el indicado.
    values[-1] = round(values[-1] + monthly_average * period_count - sum(values), 2)
    target_total_cents = int(round(monthly_average * 100)) * period_count
    lower_cents = int(round(monthly_average * (1.0 - variation) * 100))
    upper_cents = int(round(monthly_average * (1.0 + variation) * 100))
    base_cents = [int(round(value * 100)) for value in values]

    # Introduce centavos realistas sin cambiar la suma. El ultimo periodo absorbe
    # la compensacion, por lo que el promedio del plan permanece exacto.
    cents_rng = np.random.default_rng(int(seed) + 104729)
    for _ in range(500):
        revised = []
        for base in base_cents[:-1]:
            candidate = base + int(cents_rng.integers(-97, 98))
            candidate = min(max(candidate, lower_cents), upper_cents)
            if candidate % 100 == 0:
                candidate += 37 if candidate + 37 <= upper_cents else -37
            revised.append(candidate)
        final_candidate = target_total_cents - sum(revised)
        if lower_cents <= final_candidate <= upper_cents and final_candidate % 100 != 0:
            revised.append(final_candidate)
            if len(set(revised)) == len(revised):
                result = [value / 100.0 for value in revised]
                if period_count >= 3:
                    # Evita secuencias artificiales totalmente ascendentes o
                    # descendentes. El orden sigue dependiendo de la semilla.
                    order_rng = np.random.default_rng(int(seed) + 130363)
                    for _ in range(30):
                        candidate = [result[position] for position in order_rng.permutation(period_count)]
                        ascending = all(candidate[i] < candidate[i + 1] for i in range(period_count - 1))
                        descending = all(candidate[i] > candidate[i + 1] for i in range(period_count - 1))
                        if not ascending and not descending:
                            return candidate
                    return [max(result), min(result), *sorted(result)[1:-1]]
                return result
    # Respaldo extremadamente improbable: conserva rango y promedio aunque no
    # haya sido posible satisfacer la regla adicional de centavos para todos.
    return [value / 100.0 for value in base_cents]


def parse_semicolon_catalog(text, fallback):
    """Convierte una lista editable separada por punto y coma en conceptos unicos."""
    if text is None:
        return tuple(fallback)
    values = [re.sub(r"\s+", " ", item).strip().upper() for item in str(text or "").split(";")]
    values = [item for item in values if item]
    return tuple(dict.fromkeys(values))

BANK_CODES = {
    "BANAMEX": "002",
    "BANAMEX MICUENTA": "002",
    "BANAMEX MICUENTA (JUN-JUL)": "002",
    "BANAMEX (JUN-JUL)": "002",
    "BBVA LARGO": "012",
    "BBVA 2": "012",
    "SANTANDER": "014",
    "SANTANDER SELECT": "014",
    "HSBC": "021",
    "SCOTIABANK": "044",
    "BANORTE": "072",
    "AZTECA": "127",
    "DIDÁCTICO GENÉRICO": "999",
}

# =========================================================
# APP UNIFICADA
# FASE 1: PDF estado de cuenta -> Excel + resumen
# FASE 2: Excel -> PDF estilo Banorte
# =========================================================

PAGE_W_PT = 612.0
PAGE_H_PT = 792.0

FONT_REGULAR_FILE = "MYRIADPRO-COND.OTF"
FONT_BOLD_FILE = "MYRIADPRO-BOLDCOND.OTF"
FONT_REGULAR_NAME = "MyriadProCond"
FONT_BOLD_NAME = "MyriadProBoldCond"

# Calibración Banorte detectada sobre el PDF enviado.
# V3: se abrió la separación entre FECHA y DESCRIPCIÓN para evitar que se empalmen.
FONT_SIZE_DEFAULT = 9.0
HEADER_FONT_SIZE_DEFAULT = 9.0
LINE_H_DEFAULT = 10.80
ROW_EXTRA_DEFAULT = 1.88

X_LINE_LEFT = 50.40
X_LINE_RIGHT = 561.82
X_FECHA = 54.31
X_DESC_DEFAULT = 93.50
DESC_W_DEFAULT = 254.00
X_DEP_R = 416.91
X_RET_R = 489.84
X_SALDO_R = 559.07
Y_TITLE = 59.68
Y_SUBTITLE = 75.70
Y_HEADER = 91.53
Y_HEADER_LINE = 95.955
Y_BOTTOM = 742.0
BASELINE_OFFSET = 9.58
SEPARATOR_GRAY = (0.862745, 0.862745, 0.862745)

MONTHS_ES = {
    1: "ENE", 2: "FEB", 3: "MAR", 4: "ABR", 5: "MAY", 6: "JUN",
    7: "JUL", 8: "AGO", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DIC"
}
MONTH_TO_NUM = {
    "ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DIC": 12,
    "JAN": 1, "APR": 4, "AUG": 8, "DEC": 12,
}
DATE_RE = re.compile(r"^\s*(\d{1,2})[-/ ]([A-ZÁÉÍÓÚ]{3,}|\d{1,2})[-/ ](\d{2,4})(?=\s|[A-Z0-9]|$)", re.I)
SHORT_DATE_RE = re.compile(r"^\s*(\d{1,2})[-/ ]([A-ZÁÉÍÓÚ]{3})(?=\s|$)", re.I)
MONEY_RE = re.compile(r"(?<![A-Z0-9])(?:\$\s*)?(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2}(?![A-Z0-9])")
WITHDRAWAL_WORDS = [
    "RETIRO", "CARGO", "PAGO", "CHEQUE PAGADO", "COMISION", "COMISIÓN",
    "IVA", "INTERESES", "PENALIZACION", "PENALIZACIÓN", "CAPITAL", "SEGUROS", "ADMINISTRACION", "ADMINISTRACIÓN", "SPEI ENVIADO", "TRASPASO A", "TRANSFERENCIA ENVIADA", "DOMICILIACION",
    "DOMICILIACIÓN", "DISPOSICION", "DISPOSICIÓN"
]
DEPOSIT_WORDS = [
    "DEPOSITO", "DEPÓSITO", "ABONO", "SPEI RECIBIDO", "TRASPASO", "TRANSFERENCIA RECIBIDA"
]


def clean_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and np.isnan(value):
        return ""
    txt = str(value).replace("\r", " ").replace("\n", " ").strip()
    txt = re.sub(r"\s+", " ", txt)
    if txt.lower() in {"nan", "none", "null"}:
        return ""
    return txt


def amount_to_float(value):
    if value is None or value == "":
        return 0.0
    txt = str(value).replace("$", "").replace(",", "").replace(" ", "").strip()
    if not txt or txt.lower() in {"nan", "none", "null"}:
        return 0.0
    try:
        return float(txt)
    except Exception:
        return 0.0


def format_money_input(value):
    """Representa importes editables con separadores de miles y dos decimales."""
    return f"{amount_to_float(value):,.2f}"


def _normalize_money_input(text_key, numeric_key, min_value, max_value):
    """Valida y normaliza un campo monetario cuando el usuario termina de editarlo."""
    raw = st.session_state.get(text_key, "")
    cleaned = str(raw).replace("$", "").replace(",", "").replace(" ", "").strip()
    try:
        value = float(cleaned)
        if min_value is not None and value < min_value:
            raise ValueError
        if max_value is not None and value > max_value:
            raise ValueError
    except (TypeError, ValueError):
        value = float(st.session_state.get(numeric_key, min_value or 0.0))
        st.session_state[f"{numeric_key}__money_error"] = True
    else:
        st.session_state.pop(f"{numeric_key}__money_error", None)
    st.session_state[numeric_key] = value
    st.session_state[text_key] = format_money_input(value)


def money_input(label, *, value=0.0, key, min_value=0.0, max_value=None, disabled=False, help=None):
    """Campo monetario que muestra 1,234,567.89 y devuelve un float."""
    text_key = f"{key}__money_text"
    if key not in st.session_state:
        st.session_state[key] = float(value)
    if text_key not in st.session_state:
        st.session_state[text_key] = format_money_input(st.session_state[key])
    st.text_input(
        label,
        key=text_key,
        disabled=disabled,
        help=help,
        on_change=_normalize_money_input,
        args=(text_key, key, min_value, max_value),
    )
    if st.session_state.get(f"{key}__money_error"):
        st.caption("⚠️ Escribe un importe válido, por ejemplo: 1,250,000.00")
    # También interpreta el texto durante la ejecución actual; al perder el foco
    # el callback lo normaliza visualmente con comas y dos decimales.
    cleaned = str(st.session_state[text_key]).replace("$", "").replace(",", "").replace(" ", "").strip()
    try:
        parsed = float(cleaned)
        if min_value is not None and parsed < min_value:
            raise ValueError
        if max_value is not None and parsed > max_value:
            raise ValueError
        st.session_state[key] = parsed
    except (TypeError, ValueError):
        pass
    return float(st.session_state[key])


def format_monto(value):
    num = amount_to_float(value)
    if abs(num) < 0.005:
        return ""
    return f"{num:,.2f}"


def format_fecha(value):
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    if isinstance(value, (datetime, pd.Timestamp)):
        return f"{value.day:02d}-{MONTHS_ES[value.month]}-{str(value.year)[-2:]}"
    txt = clean_text(value).upper()
    m = DATE_RE.search(txt)
    if m:
        day = int(m.group(1))
        mon = m.group(2)[:3].upper()
        year = m.group(3)[-2:]
        return f"{day:02d}-{mon}-{year}"
    try:
        dt = pd.to_datetime(txt, dayfirst=True, errors="raise")
        return f"{dt.day:02d}-{MONTHS_ES[dt.month]}-{str(dt.year)[-2:]}"
    except Exception:
        return txt


def first_pdf_text(pdf_bytes, pages=2):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    chunks = []
    for i in range(min(pages, len(doc))):
        chunks.append(doc[i].get_text("text") or "")
    return "\n".join(chunks)


def detect_bank(pdf_bytes):
    txt = first_pdf_text(pdf_bytes, pages=2).upper()
    top = txt[:6000]
    if "BBVA MEXICO, S.A." in txt or "GRUPO FINANCIERO BBVA MEXICO" in txt:
        return "BBVA"
    if "BANCO AZTECA" in top or "GUARDADITO" in top:
        return "AZTECA"
    if "SCOTIABANK" in top or "CUENTA SCOTIA" in top:
        return "SCOTIABANK"
    if "SANTANDER" in top:
        return "SANTANDER"
    if "BANAMEX" in top or "CITIBANAMEX" in top or "MICUENTA" in top:
        return "CITIBANAMEX"
    if "BBVA" in top or "BANCOMER" in top:
        return "BBVA"
    if "BANORTE" in top or "ENLACE NEGOCIOS" in top or "MONTO DEL DEPOSITO" in top:
        return "BANORTE"
    if "HSBC" in top:
        return "HSBC"
    return "GENÉRICO"


def normalize_short_date_lines(lines, pdf_bytes):
    """Agrega el año del periodo a fechas bancarias como 07/ABR o 13 FEB."""
    text = first_pdf_text(pdf_bytes, pages=3)
    years = re.findall(r"\b20\d{2}\b", text)
    year = max(set(years), key=years.count) if years else str(datetime.now().year)
    normalized = []
    for pno, line in lines:
        if not DATE_RE.search(line):
            match = SHORT_DATE_RE.search(line)
            if match:
                replacement = f"{int(match.group(1)):02d}-{match.group(2).upper()}-{year[-2:]}"
                line = replacement + line[match.end():]
        normalized.append((pno, line))
    return normalized


def extract_lines_from_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    lines = []
    for pno, page in enumerate(doc, start=1):
        text = page.get_text("text") or ""
        for line in text.splitlines():
            line = clean_text(line)
            if line:
                lines.append((pno, line))
    return lines


def skip_noise(line):
    u = line.upper()
    noise = [
        "DETALLE DE MOVIMIENTOS", "FECHA DESCRIP", "FECHA", "DESCRIPCIÓN / ESTABLECIMIENTO",
        "DESCRIPCION / ESTABLECIMIENTO", "MONTO DEL DEPOSITO", "MONTO DEL DEPÓSITO",
        "MONTO DEL RETIRO", "SALDO", "PESOS", "CONTINUACIÓN", "CONTINUACION",
        "ENLACE NEGOCIOS", "PÁGINA", "PAGINA", "TOTAL"
    ]
    return any(n in u for n in noise) and not DATE_RE.search(u)


def parse_movement_block(block):
    if not block:
        return None
    first = block[0]
    m = DATE_RE.search(first)
    if not m:
        return None
    fecha = f"{int(m.group(1)):02d}-{m.group(2)[:3].upper()}-{m.group(3)[-2:]}"
    first_desc = first[m.end():].strip()
    all_lines = [first_desc] + block[1:]
    text = " ".join([clean_text(x) for x in all_lines if clean_text(x)])

    amounts = MONEY_RE.findall(text)
    deposito = retiro = saldo = 0.0
    upper = text.upper()

    if amounts:
        nums = [amount_to_float(x) for x in amounts]
        if "SALDO ANTERIOR" in upper:
            saldo = nums[-1]
            deposito = 0.0
            retiro = 0.0
            for amt in reversed(amounts[-1:]):
                idx = text.rfind(amt)
                if idx >= 0:
                    text = (text[:idx] + text[idx + len(amt):]).strip()
            text = re.sub(r"\s+", " ", text).strip(" ,")
            return {
                "FECHA": fecha,
                "DESCRIPCION": text,
                "DEPOSITO": deposito,
                "RETIRO": retiro,
                "SALDO": saldo,
            }
        saldo = nums[-1] if len(nums) >= 2 else 0.0
        mov = nums[-2] if len(nums) >= 2 else nums[-1]
        if len(nums) >= 3:
            # En layouts bien extraídos suele venir depósito/retiro/saldo.
            saldo = nums[-1]
            if any(w in upper for w in WITHDRAWAL_WORDS) and not any(w in upper for w in ["SPEI RECIBIDO", "DEPOSITO", "DEPÓSITO"]):
                retiro = nums[-2]
            else:
                deposito = nums[-2]
        elif len(nums) == 2:
            if any(w in upper for w in WITHDRAWAL_WORDS) and not any(w in upper for w in ["SPEI RECIBIDO", "DEPOSITO", "DEPÓSITO", "ABONO"]):
                retiro = mov
            else:
                deposito = mov
        else:
            if any(w in upper for w in WITHDRAWAL_WORDS):
                retiro = mov
            else:
                deposito = mov
        # Quitar importes al final de la descripción para que no se dupliquen.
        for amt in reversed(amounts[-3:]):
            idx = text.rfind(amt)
            if idx >= 0:
                text = (text[:idx] + text[idx + len(amt):]).strip()
        text = re.sub(r"\s+", " ", text).strip(" ,")

    return {
        "FECHA": fecha,
        "DESCRIPCION": text,
        "DEPOSITO": deposito,
        "RETIRO": retiro,
        "SALDO": saldo,
    }


def _lines_to_movements(lines):
    blocks = []
    current = []
    for _, line in lines:
        if skip_noise(line):
            continue
        if DATE_RE.search(line):
            if current:
                blocks.append(current)
            current = [line]
        elif current:
            current.append(line)
    if current:
        blocks.append(current)

    rows = []
    for block in blocks:
        row = parse_movement_block(block)
        if row:
            physical_lines = []
            for line_index, raw_line in enumerate(block):
                physical = str(raw_line)
                if line_index == 0:
                    match = DATE_RE.search(physical)
                    if match:
                        physical = physical[match.end():].strip()
                physical = clean_text(physical)
                if not physical:
                    continue
                # Los importes que PyMuPDF entrega como renglones aislados ya
                # estan en sus columnas numericas; no son parte del concepto.
                if re.fullmatch(r"(?:\$?\s*[\d,]+\.\d{2}\s*)+", physical):
                    continue
                physical_lines.append(physical)
            row[PHYSICAL_LINES_COLUMN] = encode_physical_lines(
                physical_lines or [row.get("DESCRIPCION", "")]
            )
            rows.append(row)
    return pd.DataFrame(rows, columns=["FECHA", "DESCRIPCION", "DEPOSITO", "RETIRO", "SALDO", PHYSICAL_LINES_COLUMN])


def extract_banorte_account_lines(pdf_bytes):
    """
    Extrae únicamente el bloque real de movimientos Banorte.
    Corrige el caso típico donde Banorte pega la fecha con la descripción:
    13-MAR-26BNET... sin espacio después del año.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    out = []
    in_movements_title = False
    in_account = False
    after_header = False

    stop_markers = [
        "INVERSIÓN VISTA PREFERENTE", "INVERSION VISTA PREFERENTE",
        "OTROS▼", "OTROS", "CARGOS OBJETADOS", "INFORME DE DEPÓSITOS",
        "INFORME DE DEPOSITOS", "GANANCIA ANUAL TOTAL", "COMPROBANTE FISCAL"
    ]

    for pno, page in enumerate(doc, start=1):
        for raw in (page.get_text("text") or "").splitlines():
            line = clean_text(raw)
            if not line:
                continue
            u = line.upper()

            if "DETALLE DE MOVIMIENTOS" in u:
                in_movements_title = True
                in_account = False
                after_header = False
                continue

            if not in_movements_title:
                continue

            if any(marker in u for marker in stop_markers):
                if in_account:
                    return out
                continue

            if u in {"CUENTA PREFERENTE", "ENLACE NEGOCIOS BASICA", "ENLACE NEGOCIOS BÁSICA"}:
                in_account = True
                continue

            if in_account and ("FECHA" in u or "DESCRIP" in u or "MONTO DEL" in u or u == "SALDO"):
                continue

            if in_account and DATE_RE.search(line):
                after_header = True

            if in_account and after_header:
                # Evitar textos legales y notas posteriores.
                if u.startswith("SIN MOVIMIENTOS"):
                    continue
                out.append((pno, line))

    return out


def _legacy_extract_scotiabank_movements(pdf_bytes):
    """Extrae la tabla Scotia por coordenadas para no mezclarla con la carátula."""
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
                    text = clean_text(span.get("text", ""))
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
                candidates = sorted(
                    [(bbox[1], bbox[0], text) for text, bbox in in_row if left <= bbox[0] < right]
                )
                selected = []
                last_y = None
                for item_y, item_x, item_text in candidates:
                    if last_y is not None and item_y - last_y > 11.0:
                        break
                    selected.append((item_y, item_x, item_text))
                    last_y = item_y
                return " ".join(text for _, _, text in selected)

            concept = contiguous_column(89, 250)
            reference = contiguous_column(250, 376)
            deposit_text = " ".join(text for text, bbox in in_row if 376 <= bbox[0] < 440)
            withdrawal_text = " ".join(text for text, bbox in in_row if 440 <= bbox[0] < 512)
            balance_text = " ".join(text for text, bbox in in_row if 512 <= bbox[0] < 590)
            day_text, month_text = date_text.split()
            month = MONTH_TO_NUM.get(month_text[:3], 1)
            fecha = f"{int(day_text):02d}-{MONTHS_ES[month]}-{str(year)[-2:]}"
            description = clean_text(concept + (f" REF {reference}" if reference else ""))
            def column_amount(value):
                match = MONEY_RE.search(value)
                return amount_to_float(match.group(0)) if match else 0.0

            rows.append({
                "FECHA": fecha,
                "DESCRIPCION": description,
                "DEPOSITO": column_amount(deposit_text),
                "RETIRO": column_amount(withdrawal_text),
                "SALDO": column_amount(balance_text),
            })

    df = pd.DataFrame(rows, columns=["FECHA", "DESCRIPCION", "DEPOSITO", "RETIRO", "SALDO"])
    if not df.empty:
        first = df.iloc[0]
        opening = round(float(first["SALDO"]) - float(first["DEPOSITO"]) + float(first["RETIRO"]), 2)
        opening_row = pd.DataFrame([{
            "FECHA": first["FECHA"], "DESCRIPCION": "SALDO ANTERIOR",
            "DEPOSITO": 0.0, "RETIRO": 0.0, "SALDO": opening,
        }])
        df = pd.concat([opening_row, df], ignore_index=True)
    return df


def _legacy_parse_pdf_movements(pdf_bytes, bank="AUTO"):
    if bank == "AUTO":
        bank = detect_bank(pdf_bytes)

    if bank == "BANORTE":
        lines = extract_banorte_account_lines(pdf_bytes)
    elif bank == "SCOTIABANK":
        df = extract_scotiabank_movements(pdf_bytes)
        return df.reset_index(drop=True), bank
    else:
        lines = normalize_short_date_lines(extract_lines_from_pdf(pdf_bytes), pdf_bytes)

    df = _lines_to_movements(lines)
    if df.empty:
        return df, bank

    # Limpieza final: eliminar movimientos vacíos, pero conservar SALDO ANTERIOR aunque sea 0.00.
    df = df[~((df["DESCRIPCION"].astype(str).str.strip() == "") & (df["DEPOSITO"].fillna(0).astype(float).abs() < 0.005) & (df["RETIRO"].fillna(0).astype(float).abs() < 0.005))].copy()
    return df.reset_index(drop=True), bank


def make_summary(df):
    if df.empty:
        return {
            "saldo_inicial": 0.0, "ingresos": 0.0, "egresos": 0.0, "flujo_neto": 0.0,
            "saldo_final": 0.0, "saldo_final_calculado": 0.0, "diferencia": 0.0,
            "saldo_promedio_mensual": 0.0, "saldo_minimo": 0.0,
            "egresos_cubiertos": True, "saldos_no_negativos": True,
            "mov_ingresos": 0, "mov_egresos": 0, "mov_total": 0,
        }
    dep = pd.to_numeric(df["DEPOSITO"], errors="coerce").fillna(0)
    ret = pd.to_numeric(df["RETIRO"], errors="coerce").fillna(0)
    saldo = pd.to_numeric(df["SALDO"], errors="coerce")
    saldo_valid = saldo.dropna()
    saldo_inicial = float(saldo_valid.iloc[0]) if len(saldo_valid) else 0.0
    ingresos = float(dep.sum())
    egresos = float(ret.sum())
    saldo_final = float(saldo_valid.iloc[-1]) if len(saldo_valid) else 0.0
    saldo_final_calculado = round(saldo_inicial + ingresos - egresos, 2)
    reported_summary = df.attrs.get("statement_summary", {})
    reported_average = reported_summary.get("SALDO_PROMEDIO_MENSUAL_REPORTADO")
    if reported_average is not None:
        saldo_promedio = round(float(reported_average), 2)
    else:
        try:
            saldo_promedio = _linked_periods.daily_average_balance(
                df,
                df.attrs.get("period_start"),
                df.attrs.get("period_end"),
            )
        except Exception:
            saldo_promedio = round(float(saldo_valid.mean()), 2) if len(saldo_valid) else 0.0
    saldo_minimo = float(saldo_valid.min()) if len(saldo_valid) else 0.0
    return {
        "saldo_inicial": saldo_inicial,
        "ingresos": ingresos,
        "egresos": egresos,
        "flujo_neto": round(ingresos - egresos, 2),
        "saldo_final": saldo_final,
        "saldo_final_calculado": saldo_final_calculado,
        "diferencia": round(saldo_final - saldo_final_calculado, 2),
        "saldo_promedio_mensual": saldo_promedio,
        "saldo_minimo": saldo_minimo,
        "egresos_cubiertos": saldo_inicial + ingresos + 0.005 >= egresos,
        "saldos_no_negativos": saldo_minimo >= -0.005 and saldo_final >= -0.005,
        "mov_ingresos": int((dep.abs() > 0.005).sum()),
        "mov_egresos": int((ret.abs() > 0.005).sum()),
        "mov_total": int(len(df)),
    }


def build_excel_bytes(df, bank, resumen):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_out = df.copy()
        for c in ["DEPOSITO", "RETIRO", "SALDO"]:
            if c in df_out.columns:
                df_out[c] = pd.to_numeric(df_out[c], errors="coerce")
        df_out.to_excel(writer, sheet_name="Movimientos", index=False)
        wb = writer.book
        ws = writer.sheets["Movimientos"]
        fmt_header = wb.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "align": "center"})
        fmt_money = wb.add_format({"num_format": "#,##0.00", "border": 1})
        fmt_text = wb.add_format({"border": 1, "text_wrap": True, "valign": "top"})
        fmt_date = wb.add_format({"border": 1, "align": "center"})
        ws.set_row(0, 20)
        for col, name in enumerate(df_out.columns):
            ws.write(0, col, name, fmt_header)
            width = 18
            cell_format = fmt_text
            if name == "FECHA":
                width, cell_format = 14, fmt_date
            elif name == "DESCRIPCION":
                width, cell_format = 80, fmt_text
            elif name in {"DEPOSITO", "RETIRO", "SALDO", "SALDO_ORIGINAL"}:
                width, cell_format = 16, fmt_money
            elif name in {"DOCUMENTO", "ID_MOVIMIENTO", "TIPO_TRANSFERENCIA"}:
                width = 28
            elif name in {"MOVIMIENTO", "RENGLON_MOVIMIENTO", "ES_CONTINUACION"}:
                width = 20
            ws.set_column(col, col, width, cell_format)
        n = len(df_out)
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, max(n, 1), max(len(df_out.columns) - 1, 0))

        res = pd.DataFrame([
            ["Banco detectado/seleccionado", bank],
            ["Saldo inicial", resumen["saldo_inicial"]],
            ["Total ingresos", resumen["ingresos"]],
            ["Total egresos", resumen["egresos"]],
            ["Flujo neto del periodo", resumen["flujo_neto"]],
            ["Saldo utilizado del periodo anterior", max(0.0, resumen["egresos"] - resumen["ingresos"])],
            ["Saldo final calculado", resumen["saldo_final_calculado"]],
            ["Saldo final reportado", resumen["saldo_final"]],
            ["Diferencia contable", resumen["diferencia"]],
            ["Saldo promedio mensual", resumen["saldo_promedio_mensual"]],
            ["Saldo mínimo", resumen["saldo_minimo"]],
            ["Regla saldo inicial + ingresos >= egresos", "CUMPLE" if resumen["egresos_cubiertos"] else "NO CUMPLE"],
            ["Regla saldos no negativos", "CUMPLE" if resumen["saldos_no_negativos"] else "NO CUMPLE"],
            ["Movimientos de ingresos", resumen["mov_ingresos"]],
            ["Movimientos de egresos", resumen["mov_egresos"]],
            ["Movimientos totales", resumen["mov_total"]],
            ["Generado", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ], columns=["Concepto", "Valor"])
        res.to_excel(writer, sheet_name="Resumen", index=False)
        ws2 = writer.sheets["Resumen"]
        ws2.set_column(0, 0, 32)
        ws2.set_column(1, 1, 24)
        money_fmt = wb.add_format({"num_format": "#,##0.00"})
        ws2.set_row(0, 20, fmt_header)
        ws2.set_column(1, 1, 24, money_fmt)
        last_excel_row = len(df_out) + 1
        # Formulas auditables y valores en cache para que el resumen se vea
        # correctamente incluso antes de que Excel recalcule el libro.
        ws2.write_formula(2, 1, "='Movimientos'!E2", money_fmt, resumen["saldo_inicial"])
        ws2.write_formula(3, 1, f"=SUM('Movimientos'!C2:C{last_excel_row})", money_fmt, resumen["ingresos"])
        ws2.write_formula(4, 1, f"=SUM('Movimientos'!D2:D{last_excel_row})", money_fmt, resumen["egresos"])
        ws2.write_formula(5, 1, "=B4-B5", money_fmt, resumen["flujo_neto"])
        ws2.write_formula(6, 1, "=MAX(0,B5-B4)", money_fmt, max(0.0, resumen["egresos"] - resumen["ingresos"]))
        ws2.write_formula(7, 1, "=B3+B4-B5", money_fmt, resumen["saldo_final_calculado"])
        ws2.write_formula(8, 1, f"=LOOKUP(2,1/('Movimientos'!E2:E{last_excel_row}<>\"\"),'Movimientos'!E2:E{last_excel_row})", money_fmt, resumen["saldo_final"])
        ws2.write_formula(9, 1, "=B9-B8", money_fmt, resumen["diferencia"])
        ws2.write_formula(11, 1, f"=MIN('Movimientos'!E2:E{last_excel_row})", money_fmt, resumen["saldo_minimo"])
        ws2.write_formula(12, 1, '=IF(B3+B4>=B5,"CUMPLE","NO CUMPLE")', None, "CUMPLE" if resumen["egresos_cubiertos"] else "NO CUMPLE")
        ws2.write_formula(13, 1, '=IF(AND(B12>=0,B9>=0),"CUMPLE","NO CUMPLE")', None, "CUMPLE" if resumen["saldos_no_negativos"] else "NO CUMPLE")
        if bank == "SANTANDER SELECT":
            from xlsxwriter.utility import xl_col_to_name
            select_saldo_col = xl_col_to_name(df_out.columns.get_loc("SALDO"))
            select_dep_col = xl_col_to_name(df_out.columns.get_loc("DEPOSITO"))
            select_ret_col = xl_col_to_name(df_out.columns.get_loc("RETIRO"))
            ws2.write_formula(2, 1, f"='Movimientos'!{select_saldo_col}2", money_fmt, resumen["saldo_inicial"])
            ws2.write_formula(3, 1, f"=SUM('Movimientos'!{select_dep_col}2:{select_dep_col}{last_excel_row})", money_fmt, resumen["ingresos"])
            ws2.write_formula(4, 1, f"=SUM('Movimientos'!{select_ret_col}2:{select_ret_col}{last_excel_row})", money_fmt, resumen["egresos"])
            ws2.write_formula(8, 1, f'=LOOKUP(2,1/(\'Movimientos\'!{select_saldo_col}2:{select_saldo_col}{last_excel_row}<>""),\'Movimientos\'!{select_saldo_col}2:{select_saldo_col}{last_excel_row})', money_fmt, resumen["saldo_final"])
            ws2.write_formula(11, 1, f"=MIN('Movimientos'!{select_saldo_col}2:{select_saldo_col}{last_excel_row})", money_fmt, resumen["saldo_minimo"])
    output.seek(0)
    return output.getvalue()


def build_linked_excel_bytes(periods, summary):
    """Crea un libro de auditoria con el origen de cada movimiento."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="dd-mmm-yyyy") as writer:
        summary.to_excel(writer, sheet_name="Resumen", index=False)
        consolidated = []
        for index, period in enumerate(periods, start=1):
            export = period.copy()
            export["FECHA"] = pd.to_datetime(export["FECHA"], errors="coerce")
            export.to_excel(writer, sheet_name=f"Periodo {index}", index=False)
            consolidated.append(export.assign(PERIODO=index))
        consolidated_df = pd.concat(consolidated, ignore_index=True)
        consolidated_df.to_excel(writer, sheet_name="Consolidado", index=False)

        workbook = writer.book
        header = workbook.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#7A0019", "align": "center", "valign": "vcenter"})
        money = workbook.add_format({"num_format": "$#,##0.00;[Red]($#,##0.00);-", "align": "right"})
        date_fmt = workbook.add_format({"num_format": "dd-mmm-yyyy", "align": "center"})
        original_fmt = workbook.add_format({"bg_color": "#E2F0D9", "font_color": "#006100"})
        fictitious_fmt = workbook.add_format({"bg_color": "#FFF2CC", "font_color": "#9C6500"})
        system_fmt = workbook.add_format({"bg_color": "#D9EAF7", "font_color": "#1F4E78"})
        sheet_frames = {"Resumen": summary, "Consolidado": consolidated_df}
        sheet_frames.update({f"Periodo {i}": period for i, period in enumerate(periods, start=1)})

        for sheet_name, dataframe in sheet_frames.items():
            worksheet = writer.sheets[sheet_name]
            worksheet.hide_gridlines(2)
            worksheet.freeze_panes(1, 0)
            for col_index, column in enumerate(dataframe.columns):
                worksheet.write(0, col_index, column, header)
                width = 18
                if column == "DESCRIPCION":
                    width = 72
                elif column in {"DOCUMENTO", "ID_MOVIMIENTO", "TIPO_TRANSFERENCIA"}:
                    width = 28
                worksheet.set_column(col_index, col_index, width)
                if column in {"DEPOSITO", "RETIRO", "SALDO", "SALDO_ORIGINAL"} or column.startswith("SALDO_") or column.endswith("_TOTALES"):
                    worksheet.set_column(col_index, col_index, width, money)
                elif column in {"FECHA", "FECHA_INICIAL", "FECHA_FINAL"}:
                    worksheet.set_column(col_index, col_index, width, date_fmt)
            worksheet.set_row(0, 24)
            if len(dataframe):
                worksheet.autofilter(0, 0, len(dataframe), len(dataframe.columns) - 1)
            if "ORIGEN" in dataframe.columns:
                origin_col = dataframe.columns.get_loc("ORIGEN")
                for value, fmt in (("ORIGINAL", original_fmt), ("FICTICIO", fictitious_fmt), ("SISTEMA", system_fmt)):
                    worksheet.conditional_format(1, origin_col, max(len(dataframe), 1), origin_col, {
                        "type": "text", "criteria": "containing", "value": value, "format": fmt,
                    })

        instructions = workbook.add_worksheet("Instrucciones")
        instructions.hide_gridlines(2)
        instructions.set_column("A:A", 110)
        instructions.write("A1", "AUDITORIA DE PERIODOS ENLAZADOS", header)
        notes = [
            "MATERIAL DIDACTICO - DATOS FICTICIOS. No es un documento bancario autentico.",
            "ORIGINAL conserva fecha, concepto e importe extraidos del PDF fuente.",
            "FICTICIO identifica cada movimiento agregado por el generador local.",
            "SALDO es el saldo recalculado de la simulacion; SALDO_ORIGINAL conserva el valor extraido para cotejo.",
            "El saldo final de cada periodo se usa como saldo inicial del siguiente.",
            "Los PDF fuente no se modifican ni se sobrescriben.",
        ]
        for row_index, note in enumerate(notes, start=2):
            instructions.write(row_index - 1, 0, note)
    output.seek(0)
    return output.getvalue()


def apply_linked_proposal_to_manual(summary):
    """Carga los montos sugeridos en los controles editables de Streamlit."""
    for _, row in summary.iterrows():
        period = int(row["PERIODO"])
        if bool(row.get("CALCULO_AUTOMATICO_MONTOS", False)) or bool(row.get("CALCULO_GUIADO_INGRESOS", False)):
            st.session_state[f"linked_amount_mode_{period}"] = "Montos manuales"
            if "INGRESOS_FICTICIOS_CANTIDAD" in row:
                st.session_state[f"linked_inc_count_{period}"] = int(row["INGRESOS_FICTICIOS_CANTIDAD"])
            if "EGRESOS_FICTICIOS_CANTIDAD" in row:
                st.session_state[f"linked_exp_count_{period}"] = int(row["EGRESOS_FICTICIOS_CANTIDAD"])
            st.session_state[f"linked_inc_total_{period}"] = float(row["INGRESOS_FICTICIOS"])
            st.session_state[f"linked_manual_exp_total_{period}"] = float(row["EGRESOS_FICTICIOS"])


def apply_numeric_suggestion(widget_key, value, enable_key=None):
    """Aplica una alternativa del asistente antes de la siguiente simulacion."""
    if enable_key:
        st.session_state[enable_key] = True
    st.session_state[widget_key] = float(value)
    st.session_state[f"{widget_key}__money_text"] = format_money_input(value)


def apply_synthetic_average_option(average_value, monthly_income_value):
    """Fija promedio e ingresos para que una propuesta aceptada sea estable."""
    apply_numeric_suggestion("gen_target_average", average_value)
    apply_numeric_suggestion(
        "gen_monthly_income_reference",
        monthly_income_value,
        "gen_use_income_reference",
    )


def show_feasibility_options(
    exc, key_prefix, average_widget_key=None, income_widget_key=None, income_enable_key=None
):
    """Presenta alternativas numericas aplicables para una combinacion inviable."""
    options = getattr(exc, "options", None) or {}
    if not options:
        return
    period = int(options.get("period", 1))
    rows = []
    maximum_average = options.get("maximum_viable_average")
    minimum_average = options.get("minimum_viable_average")
    closest_average = options.get("closest_average")
    maximum_final = options.get("maximum_viable_final")
    maximum_final_plan_average = options.get("maximum_viable_final_plan_average")
    required_income = options.get("required_monthly_income")
    required_plan_average = options.get("required_plan_average")
    if maximum_average is not None:
        rows.append({"OPCION": "Promedio máximo alcanzable con estos datos", "MONTO_SUGERIDO": float(maximum_average)})
    if minimum_average is not None:
        rows.append({"OPCION": "Promedio mínimo alcanzable con estos datos", "MONTO_SUGERIDO": float(minimum_average)})
    if closest_average is not None:
        rows.append({"OPCION": "Promedio más cercano encontrado", "MONTO_SUGERIDO": float(closest_average)})
    if maximum_final is not None:
        rows.append({"OPCION": "Mantener ingresos y usar el saldo final máximo", "MONTO_SUGERIDO": float(maximum_final)})
    if maximum_final_plan_average is not None:
        rows.append({"OPCION": "Nuevo saldo final promedio del plan", "MONTO_SUGERIDO": float(maximum_final_plan_average)})
    if required_income is not None:
        rows.append({"OPCION": f"Ingreso total requerido en el periodo {period}", "MONTO_SUGERIDO": float(required_income)})
    if required_plan_average is not None:
        rows.append({"OPCION": "Nuevo ingreso mensual promedio del plan", "MONTO_SUGERIDO": float(required_plan_average)})
    if rows:
        st.markdown("##### Opciones numéricas sugeridas")
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    button_columns = st.columns(2)
    if closest_average is not None or maximum_average is not None or minimum_average is not None:
        suggested_average = closest_average if closest_average is not None else (maximum_average if maximum_average is not None else minimum_average)
        average_label = "Aplicar promedio más cercano" if closest_average is not None else ("Aplicar promedio máximo viable" if maximum_average is not None else "Aplicar promedio mínimo viable")
        with button_columns[0]:
            if average_widget_key == "gen_target_average" and required_income is not None:
                st.button(
                    average_label + " con ingreso compatible",
                    on_click=apply_synthetic_average_option,
                    args=(float(suggested_average), float(required_income)),
                    key=f"{key_prefix}_apply_average_{period}",
                    use_container_width=True,
                )
            else:
                st.button(
                    average_label,
                    on_click=apply_numeric_suggestion,
                    args=(average_widget_key or f"linked_target_value_{period}", float(suggested_average)),
                    key=f"{key_prefix}_apply_average_{period}",
                    use_container_width=True,
                )
    elif maximum_final_plan_average is not None:
        with button_columns[0]:
            st.button(
                "Aplicar nuevo promedio de saldo final",
                on_click=apply_numeric_suggestion,
                args=("linked_final_plan_average", float(maximum_final_plan_average)),
                key=f"{key_prefix}_apply_final_plan_{period}",
                use_container_width=True,
            )
    elif maximum_final is not None:
        with button_columns[0]:
            st.button(
                "Aplicar saldo final máximo viable",
                on_click=apply_numeric_suggestion,
                args=(f"linked_target_value_{period}", float(maximum_final)),
                key=f"{key_prefix}_apply_final_{period}",
                use_container_width=True,
            )
    if required_plan_average is not None:
        with button_columns[1]:
            st.button(
                "Aplicar nuevo promedio de ingresos",
                on_click=apply_numeric_suggestion,
                args=("linked_income_plan_average", float(required_plan_average)),
                key=f"{key_prefix}_apply_plan_income_{period}",
                use_container_width=True,
            )
    elif required_income is not None:
        with button_columns[1]:
            st.button(
                "Aplicar ingreso requerido",
                on_click=apply_numeric_suggestion,
                args=(
                    income_widget_key or f"linked_monthly_income_target_{period}",
                    float(required_income),
                    income_enable_key,
                ),
                key=f"{key_prefix}_apply_income_{period}",
                use_container_width=True,
            )


def find_font(filename):
    base = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(base, filename),
        os.path.join(base, "fuentes", filename),
        os.path.join(os.getcwd(), filename),
        os.path.join(os.getcwd(), "fuentes", filename),
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    raise FileNotFoundError(f"No encontré {filename}. Colócalo junto al app.py o dentro de la carpeta fuentes.")


def normalize_excel_columns(df):
    cols_original = list(df.columns)
    norm = {str(c).strip().upper(): c for c in cols_original}
    aliases = {
        "FECHA": ["FECHA", "DIA", "DÍA", "DATE"],
        "DESCRIPCION": ["DESCRIPCION", "DESCRIPCIÓN", "DESCRIPCION / ESTABLECIMIENTO", "DESCRIPCIÓN / ESTABLECIMIENTO", "CONCEPTO"],
        "DEPOSITO": ["DEPOSITO", "DEPÓSITO", "MONTO DEL DEPOSITO", "MONTO DEL DEPÓSITO", "ABONO", "INGRESO"],
        "RETIRO": ["RETIRO", "MONTO DEL RETIRO", "CARGO", "EGRESO"],
        "SALDO": ["SALDO", "SALDO EN PDF"],
    }
    selected = {}
    for target, options in aliases.items():
        for op in options:
            if op in norm:
                selected[target] = norm[op]
                break
    if len(selected) < 5:
        if df.shape[1] < 5:
            raise ValueError("El Excel debe tener 5 columnas: FECHA, DESCRIPCION, DEPOSITO, RETIRO y SALDO.")
        df = df.iloc[:, :5].copy()
        df.columns = ["FECHA", "DESCRIPCION", "DEPOSITO", "RETIRO", "SALDO"]
    else:
        df = df[[selected["FECHA"], selected["DESCRIPCION"], selected["DEPOSITO"], selected["RETIRO"], selected["SALDO"]]].copy()
        df.columns = ["FECHA", "DESCRIPCION", "DEPOSITO", "RETIRO", "SALDO"]
    df = df.fillna("")
    return df.reset_index(drop=True)


def consolidate_continuation_rows(df):
    out = []
    for _, row in df.iterrows():
        fecha = clean_text(row.get("FECHA", ""))
        desc = clean_text(row.get("DESCRIPCION", ""))
        dep = amount_to_float(row.get("DEPOSITO", ""))
        ret = amount_to_float(row.get("RETIRO", ""))
        saldo = amount_to_float(row.get("SALDO", ""))
        has_amount = abs(dep) > 0.005 or abs(ret) > 0.005 or abs(saldo) > 0.005
        if not fecha and not has_amount and out:
            if desc:
                out[-1]["DESCRIPCION"] = clean_text(out[-1]["DESCRIPCION"] + " " + desc)
        else:
            out.append({"FECHA": row.get("FECHA", ""), "DESCRIPCION": desc, "DEPOSITO": dep, "RETIRO": ret, "SALDO": saldo})
    return pd.DataFrame(out, columns=["FECHA", "DESCRIPCION", "DEPOSITO", "RETIRO", "SALDO"])


def read_excel_file(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".xls"):
        raw = pd.read_excel(uploaded_file, engine="xlrd")
    else:
        raw = pd.read_excel(uploaded_file, engine="openpyxl")
    return consolidate_continuation_rows(normalize_excel_columns(raw))


def text_width(font, text, size):
    if not text:
        return 0.0
    return font.text_length(str(text), fontsize=size)


def wrap_text(text, max_width, font, size):
    text = clean_text(text)
    if not text:
        return [""]
    words = text.split(" ")
    lines = []
    current = ""
    for word in words:
        trial = word if not current else current + " " + word
        if text_width(font, trial, size) <= max_width:
            current = trial
        else:
            if current:
                lines.append(current)
            if text_width(font, word, size) > max_width:
                piece = ""
                for ch in word:
                    trial_piece = piece + ch
                    if text_width(font, trial_piece, size) <= max_width:
                        piece = trial_piece
                    else:
                        if piece:
                            lines.append(piece)
                        piece = ch
                current = piece
            else:
                current = word
    if current:
        lines.append(current)
    return lines


def insert_text(page, x, y, text, fontfile, fontname, size):
    if text:
        page.insert_text(fitz.Point(x, y), text, fontname=fontname, fontfile=fontfile, fontsize=size, color=(0, 0, 0))


def insert_right(page, right_x, y, text, font, fontfile, fontname, size):
    if text:
        x = right_x - text_width(font, text, size)
        insert_text(page, x, y, text, fontfile, fontname, size)


def draw_header(page, regular_path, bold_path, header_size, x_desc=X_DESC_DEFAULT):
    page.draw_rect(fitz.Rect(50, 760, 562, 782), color=(0.65, 0.0, 0.08), fill=(0.65, 0.0, 0.08))
    page.insert_text(fitz.Point(203, 775), "DOCUMENTO FICTICIO - USO DIDACTICO", fontname=FONT_BOLD_NAME, fontfile=bold_path, fontsize=9, color=(1, 1, 1))
    insert_text(page, 223.91, Y_TITLE, "DETALLE DE MOVIMIENTOS (PESOS) (CONTINUACIÓN)", bold_path, FONT_BOLD_NAME, header_size)
    insert_text(page, 270.65, Y_SUBTITLE, "Enlace Negocios Basica", regular_path, FONT_REGULAR_NAME, header_size)
    insert_text(page, 59.04, Y_HEADER, "FECHA", bold_path, FONT_BOLD_NAME, header_size)
    insert_text(page, x_desc, Y_HEADER, "DESCRIPCIÓN / ESTABLECIMIENTO", bold_path, FONT_BOLD_NAME, header_size)
    insert_text(page, 353.55, Y_HEADER, "MONTO DEL DEPOSITO", bold_path, FONT_BOLD_NAME, header_size)
    insert_text(page, 431.04, Y_HEADER, "MONTO DEL RETIRO", bold_path, FONT_BOLD_NAME, header_size)
    insert_text(page, 533.28, Y_HEADER, "SALDO", bold_path, FONT_BOLD_NAME, header_size)
    page.draw_line(fitz.Point(X_LINE_LEFT, Y_HEADER_LINE), fitz.Point(X_LINE_RIGHT, Y_HEADER_LINE), color=(0, 0, 0), width=1.05)


def generate_banorte_pdf_bytes(df, font_size=FONT_SIZE_DEFAULT, header_size=HEADER_FONT_SIZE_DEFAULT, line_h=LINE_H_DEFAULT, row_extra=ROW_EXTRA_DEFAULT, x_desc=X_DESC_DEFAULT, desc_w=DESC_W_DEFAULT):
    regular_path = find_font(FONT_REGULAR_FILE)
    bold_path = find_font(FONT_BOLD_FILE)
    font_reg = fitz.Font(fontfile=regular_path)
    doc = fitz.open()
    page = doc.new_page(width=PAGE_W_PT, height=PAGE_H_PT)
    draw_header(page, regular_path, bold_path, header_size, x_desc=x_desc)
    current_sep_y = Y_HEADER_LINE
    for _, row in df.iterrows():
        fecha = format_fecha(row["FECHA"])
        descripcion = clean_text(row["DESCRIPCION"])
        deposito = format_monto(row["DEPOSITO"])
        retiro = format_monto(row["RETIRO"])
        saldo = format_monto(row["SALDO"])
        lines = wrap_text(descripcion, desc_w, font_reg, font_size)
        row_h = max(1, len(lines)) * line_h + row_extra
        if current_sep_y + row_h > Y_BOTTOM:
            page = doc.new_page(width=PAGE_W_PT, height=PAGE_H_PT)
            draw_header(page, regular_path, bold_path, header_size, x_desc=x_desc)
            current_sep_y = Y_HEADER_LINE
        y = current_sep_y + BASELINE_OFFSET
        insert_text(page, X_FECHA, y, fecha, regular_path, FONT_REGULAR_NAME, font_size)
        for i, line in enumerate(lines):
            insert_text(page, x_desc, y + i * line_h, line, regular_path, FONT_REGULAR_NAME, font_size)
        insert_right(page, X_DEP_R, y, deposito, font_reg, regular_path, FONT_REGULAR_NAME, font_size)
        insert_right(page, X_RET_R, y, retiro, font_reg, regular_path, FONT_REGULAR_NAME, font_size)
        insert_right(page, X_SALDO_R, y, saldo, font_reg, regular_path, FONT_REGULAR_NAME, font_size)
        current_sep_y = current_sep_y + row_h
        page.draw_line(fitz.Point(X_LINE_LEFT, current_sep_y), fitz.Point(X_LINE_RIGHT, current_sep_y), color=SEPARATOR_GRAY, width=0.75)
    return doc.tobytes(deflate=True, garbage=4)


# =========================================================
# STREAMLIT UI
# =========================================================
st.set_page_config(page_title="Estados de cuenta - Excel y PDF", layout="wide", page_icon="🏦")
st.title("🏦 Estados de cuenta: PDF → Excel → PDF")
st.caption("Extrae movimientos, conserva archivos base y genera ejercicios ficticios en varios estilos bancarios.")

# Acabado independiente del formato y de los datos.
select_finish = st.sidebar.selectbox(
    "Acabado Santander Select", ["Escaneo editable (texto)", "Nítido"],
    help="Se aplica sólo a SANTANDER SELECT. Ambas opciones conservan texto real. Escaneo editable usa contornos de fuente con apariencia escaneada.",
    key="select_finish_choice",
)
select_scope = st.sidebar.selectbox("Contenido Santander Select", ["Solo movimientos", "Estado completo editable"], key="select_scope")
select_cover_enabled = select_scope == "Estado completo editable"
select_cover_fields = {}
with st.sidebar.expander("Datos de la carátula y cabeceras", expanded=select_cover_enabled):
    st.caption("Datos ficticios. Los totales y el saldo promedio se calculan desde el Excel.")
    for field,label,default in [
        ("holder","Nombre ficticio","PERSONA FICTICIA PARA CLASE"),
        ("address","Domicilio ficticio","CALLE DIDACTICA 123, CIUDAD DE EJEMPLO"),
        ("rfc","RFC de ejemplo","XAXX010101000"),("client","Cliente de ejemplo","00000000"),
        ("account","Cuenta de ejemplo","0000000000"),("secondary_account","Segunda cuenta sin movimientos","0000000001"),
        ("clabe","CLABE de ejemplo","000000000000000000"),("branch","Sucursal de ejemplo","SUCURSAL DIDACTICA"),
        ("phone","Teléfono de ejemplo","0000000000"),("postal","Código postal fiscal ficticio","00000"),
        ("regime","Régimen fiscal de ejemplo","EJEMPLO")]:
        select_cover_fields[field] = st.text_input(label, default, max_chars=160 if field=="address" else 80, key="select_cover_"+field)
    for field,label in [("interest","Intereses brutos"),("tax","ISR retenido"),("commissions","Comisiones"),("gat","GAT nominal"),("gat_real","GAT real")]:
        select_cover_fields[field] = st.number_input(label,value=0.0,key="select_field_"+field)
    select_cover_fields["net_interest"] = select_cover_fields["interest"] - select_cover_fields["tax"]
select_recalculate = st.sidebar.checkbox("Recalcular saldos desde depósitos y retiros", value=True)
select_opening_override = st.sidebar.checkbox("Indicar saldo inicial manualmente",value=False)
select_opening = st.sidebar.number_input("Saldo inicial del ejercicio",value=0.0) if select_opening_override else None
select_full_month = st.sidebar.checkbox("Usar el periodo del mes completo",value=True)
select_reference = None
select_annex_rows = None
with st.sidebar.expander("Elementos fijos del PDF de referencia",expanded=False):
    st.caption("Opcional: conserva solo el logo superior y el pie fijo. Las carátulas, cabeceras, movimientos y datos fiscales se reconstruyen como texto didáctico.")
    select_reference_upload = st.file_uploader("PDF Santander de referencia",type=["pdf"],key="select_fixed_reference")
    if select_reference_upload:
        select_reference = select_reference_upload.getvalue()
        import fitz as select_fitz, hashlib as select_hashlib
        select_ref_doc = select_fitz.open(stream=select_reference,filetype="pdf")
        select_ref_key = select_hashlib.sha256(select_reference).hexdigest()[:16]
        first_annex = st.number_input("Primera hoja de abreviaturas / avisos",min_value=1,max_value=len(select_ref_doc),value=min(8,len(select_ref_doc)))
        annex_top = st.number_input("Inicio del texto fijo en esa hoja (puntos)",min_value=132,max_value=700,value=202,help="Excluye cadenas fiscales y datos variables que aparezcan antes del catálogo.")
        if st.button("Reconocer texto de los anexos seleccionados"):
            try:
                from select_editable_document import extract_fixed_annexes
                with st.spinner("Reconociendo anexos; después revisa el texto y sus posiciones..."):
                    st.session_state["select_annex_"+select_ref_key] = extract_fixed_annexes(select_reference,int(first_annex),int(annex_top))
            except Exception as error:
                st.error("No se pudieron reconocer los anexos: "+str(error))
        if "select_annex_"+select_ref_key in st.session_state:
            st.caption("Revisa errores OCR, espacios y datos específicos. El texto de estos anexos también será vectorial y editable.")
            edited_annex = st.data_editor(pd.DataFrame(st.session_state["select_annex_"+select_ref_key]),num_rows="dynamic",key="select_annex_editor_"+select_ref_key,hide_index=True)
            select_annex_rows = edited_annex.to_dict("records")
            st.download_button("Guardar plantilla de anexos revisados",json.dumps(select_annex_rows,ensure_ascii=False,indent=2).encode("utf-8"),file_name="select_anexos_revisados.json",mime="application/json")
    saved_annex = st.file_uploader("Plantilla de anexos revisados (JSON)",type=["json"],key="select_saved_annex")
    if saved_annex:
        try:select_annex_rows = json.loads(saved_annex.getvalue())
        except Exception as error:st.error("Plantilla no válida: "+str(error))
    st.caption("Sin anexos revisados se usa un anexo educativo de ejemplo, no una copia completa de los avisos originales.")
with st.sidebar.expander("Fuentes instalables Santander",expanded=False):
    from pathlib import Path as SelectPath
    import zipfile as select_zip
    select_font_zip = io.BytesIO()
    with select_zip.ZipFile(select_font_zip,"w",select_zip.ZIP_DEFLATED) as archive:
        for font in SelectPath(__file__).parent.glob("CLC_Select_*.ttf"):archive.write(font,font.name)
        regular = SelectPath(__file__).with_name("CLC_Movimientos_Select_Regular.ttf")
        archive.write(regular,regular.name)
        for license_path in SelectPath(__file__).parent.glob("LICENSE-Select-*.txt"):archive.write(license_path,license_path.name)
    st.download_button("Descargar todas las fuentes",select_font_zip.getvalue(),file_name="FUENTES_SELECT_EDITABLES.zip",mime="application/zip")
    st.caption("Instálalas para editar en un editor PDF compatible. Son aproximaciones didácticas, no las fuentes oficiales del banco.")
    select_example = SelectPath(__file__).with_name("PLANTILLA_SELECT_EDITABLE.xlsx")
    if select_example.exists():
        st.download_button("Descargar Excel de movimientos y carátula",select_example.read_bytes(),file_name=select_example.name,mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
_generate_movement_pdf_original = generate_movement_pdf_bytes

def generate_movement_pdf_bytes(df, identity, start, end, style):
    if style == "SANTANDER SELECT":
        from select_editable_document import generate_editable_document
        identity = dict(identity)
        excel_meta = df.attrs.get("select_cover", {})
        if excel_meta.get("period_start") and excel_meta.get("period_end"):
            start, end = pd.Timestamp(excel_meta["period_start"]), pd.Timestamp(excel_meta["period_end"])
        elif select_full_month:
            start = pd.Timestamp(start).replace(day=1)
            end = pd.Timestamp(end) + pd.offsets.MonthEnd(0)
        # The new export is always vector text, including its scan-looking option.
        identity["select_finish"] = "NITIDO" if select_finish=="Nítido" else "ESCANEO_EDITABLE"
        identity["select_cover"] = {**select_cover_fields, **excel_meta}
        identity["select_document"] = {"scope":"COMPLETO" if select_cover_enabled else "MOVIMIENTOS", "opening":select_opening, "recalculate":select_recalculate,"reference":select_reference,"annex_rows":select_annex_rows}
        return generate_editable_document(df,identity,start,end)
    return _generate_movement_pdf_original(df,identity,start,end,style)

tab1, tab2, tab3, tab4, tab5 = st.tabs(["Fase 1 · PDF a Excel", "Fase 2 · Excel a PDF", "Generador didáctico", "Periodos enlazados", "Revisión tipográfica"])

with tab1:
    st.subheader("Fase 1 · Convertir estado de cuenta PDF a Excel")
    col_a, col_b = st.columns([2, 1])
    with col_a:
        pdf_file = st.file_uploader("Sube el PDF del estado de cuenta", type=["pdf"], key="pdf_in")
    with col_b:
        bank_option = st.selectbox("Banco", ["AUTO", "BANORTE", "SCOTIABANK", "SANTANDER SELECT", "SANTANDER", "AZTECA", "HSBC", "BBVA", "BBVA LARGO", "BANAMEX MICUENTA", "BANAMEX MICUENTA (JUN-JUL)", "CITIBANAMEX", "GENÉRICO"], index=0)
        extraction_layout = st.radio(
            "Formato del Excel",
            [OUTPUT_MODE_SINGLE, OUTPUT_MODE_PHYSICAL],
            key="pdf_excel_layout",
        )
    if pdf_file:
        pdf_bytes = pdf_file.getvalue()
        try:
            select_scan_route = bank_option == "SANTANDER SELECT"
            if bank_option == "AUTO":
                scan_probe = fitz.open(stream=pdf_bytes, filetype="pdf")
                if not scan_probe[0].get_text().strip():
                    from santander_select_ocr import is_select_scan
                    select_scan_route = is_select_scan(pdf_bytes)
            if select_scan_route:
                @st.cache_data(show_spinner="Leyendo movimientos de la imagen; puede tardar varios minutos…")
                def read_select_scan(content, reader_version="20261009-columns-physical-v4"):
                    from santander_select_ocr import extract_select
                    return extract_select(content), "SANTANDER SELECT"
                df_mov, bank_detected = read_select_scan(pdf_bytes)
            else:
                df_mov, bank_detected = parse_pdf_movements(pdf_bytes, bank_option)
            if select_scan_route:
                reconciliation = df_mov.attrs.get("reconciliation", {})
                if reconciliation.get("financial_reconciled"):
                    st.success("Importes conciliados contra la carátula y los saldos de cada operación. Revisa todavía nombres, referencias y espacios del texto OCR.")
                else:
                    st.error("LECTURA SIN CONCILIAR: no está confirmada la integridad de los movimientos.")
                    for problem in reconciliation.get("failures", []): st.write("• " + problem)
                st.caption("El saldo final impreso por sí solo no valida una extracción. La fila SALDO ANTERIOR no es una operación.")
                observations = pd.DataFrame(df_mov.attrs.get("ocr_review", []))
                st.dataframe(observations, use_container_width=True)
                st.download_button("Descargar observaciones de revisión", observations.to_csv(index=False).encode("utf-8-sig"), file_name="SANTANDER_SELECT_REVISION.csv")
            resumen = make_summary(df_mov)
            if select_scan_route:
                resumen["mov_total"] = df_mov.attrs.get("reconciliation", {}).get("movement_count", len(df_mov)-1)
            df_export = prepare_excel_layout(df_mov, extraction_layout, bank_detected)
            if select_scan_route:
                detected_count = df_mov.attrs.get("reconciliation", {}).get("movement_count", len(df_mov))
                st.info(f"Banco usado: {bank_detected}. Operaciones: {detected_count}. Filas del Excel: {len(df_export)} (incluye saldo anterior y continuaciones).")
            else:
                st.success(f"Banco usado: {bank_detected}. Movimientos detectados: {len(df_mov)}. Filas del Excel: {len(df_export)}")
            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("Ingresos", f"${resumen['ingresos']:,.2f}")
            k2.metric("Egresos", f"${resumen['egresos']:,.2f}")
            k3.metric("Saldo final", f"${resumen['saldo_final']:,.2f}")
            k4.metric("Mov. ingresos", resumen["mov_ingresos"])
            k5.metric("Mov. egresos", resumen["mov_egresos"])
            st.dataframe(df_export, use_container_width=True, height=420)
            excel_bytes = build_excel_bytes(df_export, bank_detected, resumen)
            st.download_button(
                "📥 Descargar Excel de movimientos",
                data=excel_bytes,
                file_name=f"MOVIMIENTOS_{bank_detected}_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            with st.expander("Notas de extracción"):
                st.write("El formato de una fila conserva todo el concepto en una sola celda. El formato por renglón usa los saltos físicos conservados por el extractor; las filas de continuación dejan fecha e importes vacíos para no duplicar montos.")
        except Exception as e:
            st.error(f"No se pudo procesar el PDF: {e}")
    else:
        st.info("Sube un PDF para comenzar.")

with tab2:
    st.subheader("Fase 2 · Convertir Excel a PDF multibanco")
    excel_pdf_style = st.selectbox(
        "Estilo bancario de salida",
        ["BANORTE", "SCOTIABANK", "AZTECA", "BANAMEX", "BANAMEX (JUN-JUL)", "BBVA LARGO", "BBVA 2", "SANTANDER", "SANTANDER SELECT", "HSBC", "DIDÁCTICO GENÉRICO"],
        key="excel_pdf_style",
    )
    with st.expander("Fuentes necesarias", expanded=False):
        st.write("Coloca estas fuentes junto al app.py o dentro de la carpeta fuentes:")
        st.code("MYRIADPRO-COND.OTF\nMYRIADPRO-BOLDCOND.OTF")
    with st.sidebar:
        st.subheader("Ajuste fino PDF Banorte")
        st.caption("Estos controles solo se aplican cuando eliges BANORTE en la fase 2.")
        font_size = st.number_input("Tamaño fuente movimientos", min_value=7.0, max_value=12.0, value=FONT_SIZE_DEFAULT, step=0.1)
        header_size = st.number_input("Tamaño fuente encabezado", min_value=7.0, max_value=12.0, value=HEADER_FONT_SIZE_DEFAULT, step=0.1)
        line_h = st.number_input("Interlineado", min_value=8.0, max_value=14.0, value=LINE_H_DEFAULT, step=0.05)
        row_extra = st.number_input("Espacio extra por movimiento", min_value=0.0, max_value=5.0, value=ROW_EXTRA_DEFAULT, step=0.05)
        x_desc = st.number_input("Inicio descripción X", min_value=84.0, max_value=105.0, value=X_DESC_DEFAULT, step=0.25)
        desc_w = st.number_input("Ancho descripción", min_value=230.0, max_value=270.0, value=DESC_W_DEFAULT, step=0.5)
    excel_file = st.file_uploader("Sube el Excel de movimientos", type=["xlsx", "xls", "xlsm"], key="excel_in")
    if excel_file:
        try:
            if excel_pdf_style == "SANTANDER SELECT":
                from select_excel_reader import read_select_excel
                df_pdf = read_select_excel(excel_file.getvalue())
            else:
                df_pdf = read_excel_file(excel_file)
            st.success(f"Excel leído correctamente: {len(df_pdf)} movimientos.")
            if excel_pdf_style == "SANTANDER SELECT":
                from select_editable_document import parse_select_date
                df_pdf["FECHA"] = df_pdf["FECHA"].map(parse_select_date)
                excel_select_meta = {}
                select_book = pd.ExcelFile(io.BytesIO(excel_file.getvalue()))
                if "CARATULA" in select_book.sheet_names:
                    select_meta_rows = pd.read_excel(select_book,sheet_name="CARATULA",dtype=str).fillna("")
                    if {"CAMPO","VALOR"}.issubset(select_meta_rows.columns):
                        excel_select_meta = {str(row.CAMPO).strip():row.VALOR for row in select_meta_rows.itertuples() if str(row.CAMPO).strip() and str(row.VALOR).strip()}
                        st.caption("Se cargarán los datos de la hoja CARATULA del Excel; prevalecen sobre los campos laterales.")
                st.caption("Edita fechas, conceptos, folios e importes en la tabla antes de generar el PDF. Los saldos se recalculan al exportar si está activada la opción lateral.")
                import hashlib
                editor_key = "select_editor_" + hashlib.sha256(excel_file.getvalue()).hexdigest()[:16]
                # read_select_excel preserves edited continuation lines and zero-prefixed folios.
                original_description = df_pdf["DESCRIPCION"].copy()
                df_pdf = st.data_editor(
                    df_pdf, key=editor_key, use_container_width=True, height=420,
                    column_order=[c for c in ["FECHA", "FOLIO", "DESCRIPCION", "DEPOSITO", "RETIRO", "SALDO"] if c in df_pdf.columns],
                    disabled=[c for c in df_pdf.columns if c not in {"FECHA", "FOLIO", "DESCRIPCION", "DEPOSITO", "RETIRO", "SALDO"}],
                    num_rows="fixed",
                )
                changed = df_pdf["DESCRIPCION"].astype(str) != original_description.astype(str)
                df_pdf.attrs["select_cover"] = excel_select_meta
                for row_index in df_pdf.index[changed]:
                    df_pdf.at[row_index, "_LINEAS_DOCUMENTO"] = encode_physical_lines(str(df_pdf.at[row_index, "DESCRIPCION"]).splitlines())
                from pathlib import Path
                font_file = Path(__file__).with_name("CLC_Select_Escaneo_Editable.ttf")
                if font_file.exists():
                    st.download_button("Descargar fuente Escaneo Editable para instalar", font_file.read_bytes(), file_name=font_file.name, mime="font/ttf", key="select_font_download")
                st.caption("Escaneo editable conserva texto real. Para editar el PDF fuera del convertidor, instala la fuente y utiliza un editor de PDF compatible.")
            else:
                st.dataframe(df_pdf.head(50), use_container_width=True, height=380)
            if st.button("Generar PDF didáctico", type="primary", use_container_width=True):
                parsed_dates = pd.to_datetime(df_pdf["FECHA"], errors="coerce", dayfirst=True)
                valid_dates = parsed_dates.dropna()
                pdf_start = valid_dates.min().date() if not valid_dates.empty else datetime.now().date().replace(day=1)
                pdf_end = valid_dates.max().date() if not valid_dates.empty else datetime.now().date()
                identity = generate_identity(2026, bank_code=BANK_CODES.get(excel_pdf_style, "999"))
                if excel_pdf_style == "BANORTE":
                    pdf_out = generate_banorte_pdf_bytes(df_pdf, font_size=font_size, header_size=header_size, line_h=line_h, row_extra=row_extra, x_desc=x_desc, desc_w=desc_w)
                elif excel_pdf_style == "DIDÁCTICO GENÉRICO":
                    pdf_out = generate_themed_pdf_bytes(df_pdf, identity, pdf_start, pdf_end, "GENERICO")
                else:
                    pdf_out = generate_movement_pdf_bytes(df_pdf, identity, pdf_start, pdf_end, excel_pdf_style)
                st.success("PDF generado correctamente.")
                st.download_button(
                    "📥 Descargar PDF",
                    data=pdf_out,
                    file_name=f"{excel_pdf_style.replace(' ', '_')}_GENERADO_{datetime.now():%Y%m%d_%H%M%S}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
                if excel_pdf_style == "SCOTIABANK":
                    chart_png = generate_scotiabank_chart_png(df_pdf)
                    st.download_button(
                        "Descargar gráfica Scotiabank · PNG transparente",
                        data=chart_png,
                        file_name=f"SCOTIABANK_GRAFICA_TRANSPARENTE_{datetime.now():%Y%m%d_%H%M%S}.png",
                        mime="image/png",
                        use_container_width=True,
                    )
        except Exception as e:
            st.error(f"No se pudo procesar el Excel: {e}")
    else:
        st.info("Sube un Excel de movimientos para generar el PDF.")

with tab3:
    st.subheader("Generador de movimientos ficticios para clase")
    st.warning("Los archivos de esta sección contienen datos ficticios y deben conservar la leyenda de uso didáctico.")
    synthetic_source_mode = st.radio(
        "Origen del ejercicio",
        ["Sin archivo base", "Excel base", "PDF base"],
        horizontal=True,
        key="synthetic_source_mode",
        help="PDF base conserva sus movimientos originales y agrega movimientos nuevos en un solo periodo.",
    )
    has_synthetic_base = synthetic_source_mode != "Sin archivo base"
    target_mode = st.radio(
        "Objetivo del cálculo",
        [TARGET_FINAL, TARGET_AVERAGE],
        horizontal=True,
        key="gen_target_mode",
    )
    c1, c2, c3 = st.columns(3)
    with c1:
        start_date = st.date_input("Fecha inicial", value=datetime.now().date().replace(day=1), key="gen_start")
        end_date = st.date_input("Fecha final", value=datetime.now().date(), key="gen_end")
        opening_balance = money_input("Saldo inicial", min_value=0.0, value=100000.0, key="gen_opening_balance")
    with c2:
        minimum_movements = 1 if target_mode == TARGET_AVERAGE else 0
        income_count = st.number_input(
            "Número de ingresos nuevos" if has_synthetic_base else "Número de ingresos",
            min_value=minimum_movements, value=max(8, minimum_movements), step=1,
        )
        expense_count = st.number_input(
            "Número de egresos nuevos" if has_synthetic_base else "Número de egresos",
            min_value=minimum_movements, value=max(18, minimum_movements), step=1,
        )
    with c3:
        if target_mode == TARGET_AVERAGE:
            target_balance = money_input(
                "Saldo promedio mensual deseado", min_value=0.0,
                value=75000.0, key="gen_target_average",
            )
            use_income_reference = st.checkbox(
                "Usar ingreso mensual aproximado",
                value=False,
                key="gen_use_income_reference",
                help="Si se desactiva, el programa calculará los ingresos y egresos necesarios. El saldo final siempre será un resultado.",
            )
            if use_income_reference:
                monthly_income_reference = money_input(
                    "Ingreso mensual aproximado",
                    min_value=0.0,
                    value=250000.0,
                    key="gen_monthly_income_reference",
                )
            else:
                monthly_income_reference = None
                st.caption("Ingresos, egresos y saldo final: calculados automáticamente.")
            total_income = 0.0
        else:
            total_income = money_input(
                "Monto total de ingresos nuevos" if has_synthetic_base else "Monto total de ingresos",
                min_value=0.0, value=250000.0, key="gen_total_income",
            )
            target_balance = money_input(
                "Saldo final deseado", min_value=0.0,
                value=75000.0, key="gen_target_final",
            )
            monthly_income_reference = None
        seed = st.number_input("Semilla reproducible", min_value=1, value=2026, step=1)
        layout_mode = st.radio(
            "Presentación del concepto",
            ["Concepto completo en una fila", "Concepto separado en varias filas"],
        )

    base_file = None
    base_pdf = None
    synthetic_pdf_bank_option = "AUTO"
    synthetic_banamex_banks = list(BANAMEX_TRANSFER_BANKS)
    if synthetic_source_mode == "Excel base":
        base_file = st.file_uploader(
            "Sube el Excel base (se conservarán sus movimientos)",
            type=["xlsx", "xls", "xlsm"],
            key="base_synthetic_excel",
        )
    elif synthetic_source_mode == "PDF base":
        pdf_source_a, pdf_source_b = st.columns(2)
        with pdf_source_a:
            base_pdf = st.file_uploader(
                "Sube un estado de cuenta PDF",
                type=["pdf"],
                key="base_synthetic_pdf",
            )
        with pdf_source_b:
            synthetic_pdf_bank_option = st.selectbox(
                "Banco del PDF",
                ["AUTO", "BANORTE", "SCOTIABANK", "SANTANDER", "AZTECA", "HSBC", "BBVA", "BBVA LARGO", "BANAMEX MICUENTA", "BANAMEX MICUENTA (JUN-JUL)", "CITIBANAMEX", "GENÉRICO"],
                key="synthetic_pdf_bank_option",
                help="AUTO detecta el formato y aplica su diseño de salida.",
            )

    with st.expander("Bancos y conceptos permitidos en esta operación", expanded=False):
        synthetic_banamex_banks = st.multiselect(
            "Bancos permitidos para movimientos ficticios",
            options=list(BANAMEX_TRANSFER_BANKS),
            default=list(BANAMEX_TRANSFER_BANKS),
            key="synthetic_banamex_transfer_banks",
            help="Solo los bancos que permanezcan seleccionados podrán aparecer en movimientos nuevos.",
        )
        excluded = [bank for bank in BANAMEX_TRANSFER_BANKS if bank not in synthetic_banamex_banks]
        st.caption("Bancos excluidos: " + (", ".join(excluded) if excluded else "NINGUNO"))
        synthetic_income_concepts_text = st.text_area(
            "Conceptos permitidos para ingresos",
            value="; ".join(INCOME_CONCEPTS),
            key="synthetic_income_concepts",
            help="Separa los conceptos con punto y coma. Lo que elimines no se utilizará.",
        )
        synthetic_expense_concepts_text = st.text_area(
            "Conceptos permitidos para egresos",
            value="; ".join(EXPENSE_CONCEPTS),
            key="synthetic_expense_concepts",
            help="Puedes eliminar conceptos o agregar otros; se respetará esta lista exacta.",
        )
        synthetic_income_concepts = parse_semicolon_catalog(synthetic_income_concepts_text, INCOME_CONCEPTS)
        synthetic_expense_concepts = parse_semicolon_catalog(synthetic_expense_concepts_text, EXPENSE_CONCEPTS)

    if synthetic_source_mode == "PDF base":
        bank_style = "DIDÁCTICO GENÉRICO"
        st.caption("El estilo de salida se elegirá automáticamente según el banco detectado en el PDF.")
    else:
        bank_style = st.selectbox(
            "Estilo de salida",
            ["DIDÁCTICO GENÉRICO", "BANORTE", "SCOTIABANK", "AZTECA", "BANAMEX", "BANAMEX (JUN-JUL)", "BBVA LARGO", "BBVA 2", "SANTANDER", "SANTANDER SELECT", "HSBC"],
            key="synthetic_style",
        )

    if st.button("Crear ejercicio ficticio", type="primary", use_container_width=True):
        try:
            output_start, output_end = start_date, end_date
            detected_bank = bank_style
            base_for_average = None
            if synthetic_source_mode == "PDF base":
                if base_pdf is None:
                    raise ValueError("Sube un PDF base para conservar sus movimientos.")
                extracted, detected_bank = parse_pdf_movements_cached(
                    base_pdf.getvalue(), synthetic_pdf_bank_option
                )
                if extracted.empty:
                    raise ValueError("No se encontraron movimientos en el PDF base.")
                base_for_average = extracted
            elif target_mode == TARGET_AVERAGE and synthetic_source_mode == "Excel base":
                if base_file is None:
                    raise ValueError("Sube un Excel base para conservar sus movimientos.")
                base_for_average = read_excel_file(base_file)
                if base_for_average.empty:
                    raise ValueError("No se encontraron movimientos en el Excel base.")
            elif target_mode == TARGET_AVERAGE:
                base_for_average = pd.DataFrame(
                    columns=["FECHA", "DESCRIPCION", "DEPOSITO", "RETIRO", "SALDO"]
                )
                base_for_average.attrs.update({
                    "period_start": pd.Timestamp(start_date),
                    "period_end": pd.Timestamp(end_date),
                })

            if target_mode == TARGET_AVERAGE:
                if synthetic_source_mode != "PDF base":
                    linked_bank_by_style = {
                        "SCOTIABANK": "SCOTIABANK", "SANTANDER": "SANTANDER", "SANTANDER SELECT": "SANTANDER",
                        "AZTECA": "AZTECA", "HSBC": "HSBC", "BBVA LARGO": "BBVA",
                        "BBVA 2": "BBVA", "BANAMEX": "BANAMEX MICUENTA", "BANAMEX (JUN-JUL)": "BANAMEX MICUENTA (JUN-JUL)",
                        "BANORTE": "BANORTE", "DIDÁCTICO GENÉRICO": "GENERICO",
                    }
                    detected_bank = linked_bank_by_style.get(bank_style, "GENERICO")
                pdf_config = LinkedPeriodConfig(
                    income_count=int(income_count),
                    expense_count=int(expense_count),
                    total_income=0.0,
                    target_mode=target_mode,
                    target_value=float(target_balance),
                    transfer_banks=tuple(synthetic_banamex_banks),
                    income_concepts=synthetic_income_concepts,
                    expense_concepts=synthetic_expense_concepts,
                    auto_calculate_amounts=True,
                    average_target_final=None,
                    target_total_monthly_income=(
                        float(monthly_income_reference)
                        if monthly_income_reference is not None else None
                    ),
                    approximate_monthly_income=monthly_income_reference is not None,
                )
                periods, period_summary = link_statement_periods(
                    [(
                        base_pdf.name if synthetic_source_mode == "PDF base"
                        else (base_file.name if synthetic_source_mode == "Excel base" else "EJERCICIO_SIN_BASE"),
                        detected_bank,
                        base_for_average,
                    )],
                    [pdf_config],
                    first_opening_balance=float(opening_balance),
                    seed=int(seed),
                )
                generated = periods[0]
                st.session_state["synthetic_variety_summary"] = period_summary.iloc[0].to_dict()
                output_start = pd.Timestamp(period_summary.iloc[0]["FECHA_INICIAL"]).date()
                output_end = pd.Timestamp(period_summary.iloc[0]["FECHA_FINAL"]).date()
                st.session_state["synthetic_calculation_summary"] = period_summary.iloc[0].to_dict()
                style_by_bank = {
                    "SCOTIABANK": "SCOTIABANK", "SANTANDER": "SANTANDER", "SANTANDER SELECT": "SANTANDER", "AZTECA": "AZTECA",
                    "HSBC": "HSBC", "BBVA": "BBVA LARGO", "BANAMEX MICUENTA": "BANAMEX", "BANAMEX MICUENTA (JUN-JUL)": "BANAMEX (JUN-JUL)",
                    "CITIBANAMEX": "BANAMEX", "BANORTE": "BANORTE",
                }
                if bank_style != "SANTANDER SELECT":
                    bank_style = style_by_bank.get(str(detected_bank).upper(), "DIDÁCTICO GENÉRICO")
            elif synthetic_source_mode == "PDF base":
                pdf_config = LinkedPeriodConfig(
                    income_count=int(income_count),
                    expense_count=int(expense_count),
                    total_income=float(total_income),
                    target_mode=target_mode,
                    target_value=float(target_balance),
                    transfer_banks=tuple(synthetic_banamex_banks),
                    income_concepts=synthetic_income_concepts,
                    expense_concepts=synthetic_expense_concepts,
                )
                periods, period_summary = link_statement_periods(
                    [(base_pdf.name, detected_bank, extracted)],
                    [pdf_config],
                    first_opening_balance=float(opening_balance),
                    seed=int(seed),
                )
                generated = periods[0]
                st.session_state["synthetic_variety_summary"] = period_summary.iloc[0].to_dict()
                output_start = pd.Timestamp(period_summary.iloc[0]["FECHA_INICIAL"]).date()
                output_end = pd.Timestamp(period_summary.iloc[0]["FECHA_FINAL"]).date()
                style_by_bank = {
                    "SCOTIABANK": "SCOTIABANK", "SANTANDER": "SANTANDER", "SANTANDER SELECT": "SANTANDER", "AZTECA": "AZTECA",
                    "HSBC": "HSBC", "BBVA": "BBVA LARGO", "BANAMEX MICUENTA": "BANAMEX", "BANAMEX MICUENTA (JUN-JUL)": "BANAMEX (JUN-JUL)",
                    "CITIBANAMEX": "BANAMEX", "BANORTE": "BANORTE",
                }
                if bank_style != "SANTANDER SELECT":
                    bank_style = style_by_bank.get(str(detected_bank).upper(), "DIDÁCTICO GENÉRICO")
            else:
                config = GenerationConfig(
                    start_date=start_date,
                    end_date=end_date,
                    income_count=int(income_count),
                    expense_count=int(expense_count),
                    total_income=float(total_income),
                    opening_balance=float(opening_balance),
                    target_mode=target_mode,
                    target_balance=float(target_balance),
                    seed=int(seed),
                    income_concepts=synthetic_income_concepts,
                    expense_concepts=synthetic_expense_concepts,
                    transfer_banks=tuple(synthetic_banamex_banks),
                )
                generated = generate_movements(config)
                if synthetic_source_mode == "Excel base":
                    if base_file is None:
                        raise ValueError("Sube un Excel base para conservar sus movimientos.")
                    generated = merge_base_and_generated(read_excel_file(base_file), generated)
                st.session_state.pop("synthetic_calculation_summary", None)
            checks = movement_checks(generated)
            if synthetic_source_mode == "PDF base":
                excel_mode = OUTPUT_MODE_PHYSICAL if layout_mode.startswith("Concepto separado") else OUTPUT_MODE_SINGLE
                export_df = prepare_excel_layout(generated, excel_mode, detected_bank)
            else:
                export_df = expand_concepts(generated) if layout_mode.startswith("Concepto separado") else generated
            code_by_bank = {"SCOTIABANK": "044", "SANTANDER": "014", "AZTECA": "127", "HSBC": "021", "BBVA": "012", "BANAMEX MICUENTA": "002", "BANAMEX MICUENTA (JUN-JUL)": "002", "CITIBANAMEX": "002", "BANORTE": "072"}
            identity = generate_identity(int(seed), bank_code=code_by_bank.get(str(detected_bank).upper(), BANK_CODES.get(bank_style, "999")))
            st.session_state["synthetic_df"] = export_df
            # El Excel puede ir en una o varias filas. El PDF conserva siempre
            # un movimiento semantico y deja al perfil del banco componer sus
            # renglones fisicos; mezclar ambos formatos causaba filas duplicadas.
            st.session_state["synthetic_pdf_df"] = generated
            st.session_state["synthetic_checks"] = checks
            st.session_state["synthetic_identity"] = identity
            st.session_state["synthetic_bank_style"] = bank_style
            st.session_state["synthetic_detected_bank"] = detected_bank
            st.session_state["synthetic_start_date"] = output_start
            st.session_state["synthetic_end_date"] = output_end
            st.session_state["synthetic_source_used"] = synthetic_source_mode
        except Exception as exc:
            st.error(f"No se pudo generar el ejercicio: {exc}")
            show_feasibility_options(
                exc,
                "synthetic_feasibility",
                average_widget_key="gen_target_average",
                income_widget_key="gen_monthly_income_reference",
                income_enable_key="gen_use_income_reference",
            )

    if "synthetic_df" in st.session_state:
        export_df = st.session_state["synthetic_df"]
        pdf_df = st.session_state.get("synthetic_pdf_df", export_df)
        checks = st.session_state["synthetic_checks"]
        identity = st.session_state["synthetic_identity"]
        output_start = st.session_state.get("synthetic_start_date", start_date)
        output_end = st.session_state.get("synthetic_end_date", end_date)
        source_used = st.session_state.get("synthetic_source_used", "Sin archivo base")
        a, b, c, d = st.columns(4)
        a.metric("Ingresos", f"${checks['INGRESOS']:,.2f}")
        b.metric("Egresos", f"${checks['EGRESOS']:,.2f}")
        c.metric("Saldo final", f"${checks['SALDO_FINAL']:,.2f}")
        d.metric("Saldo promedio", f"${checks['SALDO_PROMEDIO']:,.2f}")
        calculation = st.session_state.get("synthetic_calculation_summary")
        if calculation:
            st.markdown("#### Resultado del cálculo automático mensual")
            p1, p2, p3, p4, p5 = st.columns(5)
            p1.metric("Saldo inicial", f"${float(calculation['SALDO_INICIAL_SIMULADO']):,.2f}")
            p2.metric("Ingresos originales", f"${float(calculation['INGRESOS_ORIGINALES']):,.2f}")
            p3.metric("Ingresos nuevos", f"${float(calculation['INGRESOS_FICTICIOS']):,.2f}")
            p4.metric("Egresos nuevos", f"${float(calculation['EGRESOS_FICTICIOS']):,.2f}")
            p5.metric("Saldo final resultante", f"${float(calculation['SALDO_FINAL']):,.2f}")
            difference = float(calculation["DIFERENCIA_OBJETIVO"])
            reference = calculation.get("INGRESO_MENSUAL_OBJETIVO")
            reference_text = ""
            if reference is not None and not pd.isna(reference):
                actual_monthly_income = float(calculation["INGRESOS_TOTALES"])
                reference_text = (
                    f" Referencia inicial de ingresos: ${float(reference):,.2f}; "
                    f"ingreso mensual viable calculado: ${actual_monthly_income:,.2f}."
                )
            st.info(
                f"Promedio solicitado: ${float(calculation['VALOR_OBJETIVO']):,.2f} · "
                f"promedio obtenido: ${float(calculation['SALDO_PROMEDIO_MENSUAL']):,.2f} · "
                f"diferencia: ${difference:,.2f}.{reference_text} "
                "El saldo inicial participa en los saldos diarios, pero no se contabiliza como ingreso mensual."
            )
            if bool(calculation.get("OBJETIVO_DENTRO_DE_TOLERANCIA", False)):
                st.success("El saldo promedio mensual quedó dentro de la tolerancia matemática permitida.")
            else:
                st.warning(
                    "Con los movimientos originales, las cantidades seleccionadas y la regla de no generar saldos negativos, "
                    "no existe una coincidencia exacta. Se muestra la propuesta segura más cercana para que puedas modificar "
                    "el promedio, la referencia de ingresos o las cantidades de movimientos."
                )
                closest_average = float(calculation["SALDO_PROMEDIO_MENSUAL"])
                proposal_monthly_income = float(calculation["INGRESOS_TOTALES"])
                numeric_options = [{
                    "OPCIÓN": "Propuesta segura más cercana",
                    "PROMEDIO": f"${closest_average:,.2f}",
                    "INGRESOS NUEVOS": f"${float(calculation['INGRESOS_FICTICIOS']):,.2f}",
                    "EGRESOS NUEVOS": f"${float(calculation['EGRESOS_FICTICIOS']):,.2f}",
                    "SALDO FINAL": f"${float(calculation['SALDO_FINAL']):,.2f}",
                }]
                applicable_averages = [("Aplicar el promedio más cercano", closest_average)]
                for label, field, button_label in (
                    ("Límite inferior estimado", "PROMEDIO_MINIMO_ESTIMADO", "Aplicar límite inferior"),
                    ("Límite superior estimado", "PROMEDIO_MAXIMO_ESTIMADO", "Aplicar límite superior"),
                ):
                    value = calculation.get(field)
                    if value is not None and not pd.isna(value):
                        value = float(value)
                        numeric_options.append({
                            "OPCIÓN": label,
                            "PROMEDIO": f"${value:,.2f}",
                            "INGRESOS NUEVOS": "Se recalculan",
                            "EGRESOS NUEVOS": "Se recalculan",
                            "SALDO FINAL": "Será resultado",
                        })
                        if all(abs(value - previous) > 0.005 for _, previous in applicable_averages):
                            applicable_averages.append((button_label, value))
                st.markdown("##### Opciones numéricas más cercanas")
                st.dataframe(pd.DataFrame(numeric_options), use_container_width=True, hide_index=True)
                st.caption(
                    f"Al aplicar una opción también se fijarán ${proposal_monthly_income:,.2f} como ingresos "
                    "totales del mes. Esto evita que la recomendación cambie de escala en la siguiente ejecución."
                )
                option_columns = st.columns(len(applicable_averages))
                for option_index, (button_label, value) in enumerate(applicable_averages):
                    with option_columns[option_index]:
                        st.button(
                            button_label,
                            on_click=apply_synthetic_average_option,
                            args=(value, proposal_monthly_income),
                            key=f"synthetic_apply_average_option_{option_index}",
                            use_container_width=True,
                        )
        if checks["CUADRA"] and checks.get("SALDOS_NO_NEGATIVOS", False):
            st.success("Validación contable: los movimientos cuadran y ningún saldo es negativo.")
        else:
            st.error(f"Validación contable pendiente. Diferencia: ${checks['DIFERENCIA']:,.2f}")
        st.caption(f"Titular ficticio: {identity['NOMBRE']} · Cuenta ficticia terminación {identity['CUENTA'][-4:]}")
        if source_used == "PDF base":
            st.info(
                f"PDF base conservado · Banco detectado: {st.session_state.get('synthetic_detected_bank', 'GENÉRICO')} · "
                "los conteos y montos capturados corresponden solamente a movimientos nuevos."
            )
            variety = st.session_state.get("synthetic_variety_summary", {})
            st.caption(
                f"Variedad: {int(variety.get('NOMBRES_NUEVOS_UNICOS', 0))} nombres nuevos únicos · "
                f"{int(variety.get('BANCOS_NUEVOS_UTILIZADOS', 0))} bancos utilizados."
            )
            if variety.get("DISTRIBUCION_BANCOS_INGRESOS"):
                st.caption("Ingresos nuevos · " + str(variety["DISTRIBUCION_BANCOS_INGRESOS"]))
            if variety.get("DISTRIBUCION_BANCOS_EGRESOS"):
                st.caption("Egresos nuevos · " + str(variety["DISTRIBUCION_BANCOS_EGRESOS"]))
            if int(variety.get("REPETICIONES_CONSECUTIVAS_BANCOS_NUEVOS", 0)) == 0:
                st.success("Variedad validada: ningún par de movimientos nuevos consecutivos usa el mismo banco y no hay nombres repetidos.")
            else:
                st.error("La validación de variedad detectó bancos consecutivos repetidos; la descarga fue bloqueada.")
            if variety.get("BANCOS_AUTORIZADOS_YA_USADOS", False):
                st.warning(
                    "Todos los bancos autorizados ya aparecían en el PDF base; se reutilizaron de forma balanceada, sin repetir el mismo banco consecutivamente."
                )
            if variety.get("CATALOGO_NOMBRES_AMPLIADO", False):
                st.warning(
                    "Se agotaron las combinaciones simples de nombres no usadas; el programa creó nombres ficticios nuevos con nombre compuesto."
                )
        st.dataframe(export_df, use_container_width=True, height=420)
        summary = make_summary(pdf_df)
        excel_out = build_excel_bytes(export_df, st.session_state["synthetic_bank_style"], summary)
        st.download_button(
            "Descargar Excel didáctico",
            data=excel_out,
            file_name=f"EJERCICIO_FICTICIO_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        if st.session_state["synthetic_bank_style"] == "BANORTE":
            pdf_out = generate_banorte_pdf_bytes(pdf_df)
            st.download_button(
                "Descargar PDF didáctico Banorte",
                data=pdf_out,
                file_name=f"EJERCICIO_BANORTE_FICTICIO_{datetime.now():%Y%m%d_%H%M%S}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        elif st.session_state["synthetic_bank_style"] == "SCOTIABANK":
            pdf_out = generate_movement_pdf_bytes(pdf_df, identity, output_start, output_end, "SCOTIABANK")
            st.download_button(
                "Descargar PDF didáctico Scotiabank",
                data=pdf_out,
                file_name=f"EJERCICIO_SCOTIABANK_FICTICIO_{datetime.now():%Y%m%d_%H%M%S}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
            chart_png = generate_scotiabank_chart_png(pdf_df)
            st.download_button(
                "Descargar gráfica Scotiabank · PNG transparente",
                data=chart_png,
                file_name=f"SCOTIABANK_GRAFICA_TRANSPARENTE_{datetime.now():%Y%m%d_%H%M%S}.png",
                mime="image/png",
                use_container_width=True,
                key="synthetic_scotia_chart_download",
            )
        else:
            style = st.session_state["synthetic_bank_style"]
            if style in {"AZTECA", "BANAMEX", "BANAMEX (JUN-JUL)", "BBVA LARGO", "BBVA 2", "SANTANDER", "SANTANDER SELECT", "HSBC"}:
                pdf_out = generate_movement_pdf_bytes(pdf_df, identity, output_start, output_end, style)
            else:
                pdf_out = generate_themed_pdf_bytes(pdf_df, identity, output_start, output_end, "GENERICO")
            st.download_button(
                f"Descargar PDF didáctico {style}",
                data=pdf_out,
                file_name=f"EJERCICIO_{style.replace(' ', '_')}_FICTICIO_{datetime.now():%Y%m%d_%H%M%S}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )


with tab4:
    st.subheader("Periodos enlazados multibanco")
    st.warning("La salida es una copia de simulación con la leyenda MATERIAL DIDÁCTICO. Los PDF fuente no se modifican.")
    st.caption("El procesamiento es local. Puedes combinar documentos de bancos distintos y conservar todos sus movimientos originales.")

    header_a, header_b, header_c = st.columns(3)
    with header_a:
        linked_document_count = st.number_input(
            "Número de documentos a enlazar", min_value=2, value=3, step=1, key="linked_document_count"
        )
    with header_b:
        linked_bank_option = st.selectbox(
            "Banco de los documentos",
            ["AUTO", "BANORTE", "SCOTIABANK", "SANTANDER", "AZTECA", "HSBC", "BBVA", "BBVA LARGO", "BANAMEX MICUENTA", "BANAMEX MICUENTA (JUN-JUL)", "CITIBANAMEX", "GENÉRICO"],
            key="linked_bank_option",
            help="AUTO detecta el banco de cada PDF y permite mezclar bancos.",
        )
    with header_c:
        linked_excel_layout = st.radio(
            "Formato de los Excel", [OUTPUT_MODE_SINGLE, OUTPUT_MODE_PHYSICAL], key="linked_excel_layout_mode"
        )

    linked_files = st.file_uploader(
        f"Sube {int(linked_document_count)} estados de cuenta",
        type=["pdf"], accept_multiple_files=True, key="linked_statement_files",
    )
    top_a, top_b = st.columns(2)
    with top_a:
        linked_opening = money_input(
            "Nuevo saldo inicial del primer periodo", min_value=0.0, value=25000.0, key="linked_opening"
        )
        st.caption(
            "El saldo inicial no se suma a los ingresos mensuales; sí abre la cadena: "
            "saldo siguiente = saldo anterior + ingreso − egreso."
        )
    with top_b:
        linked_seed = st.number_input("Semilla reproducible", min_value=1, value=2026, step=1, key="linked_seed")

    with st.expander("Plan de ingresos mensuales variables", expanded=True):
        plan_enabled = st.checkbox(
            "Usar un ingreso promedio con variación entre meses",
            value=True,
            key="linked_variable_income_plan",
            help="Se aplica a saldo promedio mensual y saldo final; puede desactivarse para capturar ingresos manuales.",
        )
        plan_a, plan_b = st.columns(2)
        with plan_a:
            plan_income_average = money_input(
                "Ingreso mensual promedio",
                min_value=0.0,
                value=1250000.0,
                key="linked_income_plan_average",
            )
        with plan_b:
            plan_income_variation = st.number_input(
                "Variación máxima hacia arriba o abajo (%)",
                min_value=0.0,
                max_value=100.0,
                value=7.0,
                step=0.5,
                key="linked_income_plan_variation",
            )
        planned_monthly_incomes = build_variable_income_plan(
            plan_income_average, plan_income_variation, int(linked_document_count), int(linked_seed)
        )
        if plan_enabled:
            st.caption(
                " · ".join(
                    f"Periodo {position}: ${amount:,.2f}" for position, amount in enumerate(planned_monthly_incomes, start=1)
                )
            )

    with st.expander("Plan de saldos finales", expanded=False):
        final_balance_mode = st.radio(
            "Forma de definir el saldo final",
            ["Manual por periodo", "Saldo final promedio con variación"],
            horizontal=True,
            key="linked_final_balance_mode",
        )
        final_plan_enabled = final_balance_mode == "Saldo final promedio con variación"
        final_plan_a, final_plan_b = st.columns(2)
        with final_plan_a:
            final_plan_average = money_input(
                "Saldo final promedio del conjunto",
                min_value=0.0,
                value=600000.0,
                key="linked_final_plan_average",
                disabled=not final_plan_enabled,
            )
        with final_plan_b:
            final_plan_variation = st.number_input(
                "Variación máxima del saldo final (%)",
                min_value=0.0,
                max_value=100.0,
                value=8.0,
                step=0.5,
                key="linked_final_plan_variation",
                disabled=not final_plan_enabled,
            )
        planned_final_balances = build_variable_income_plan(
            final_plan_average, final_plan_variation, int(linked_document_count), int(linked_seed) + 4001
        )
        if final_plan_enabled:
            st.caption(
                " · ".join(
                    f"Periodo {position}: ${amount:,.2f}" for position, amount in enumerate(planned_final_balances, start=1)
                )
            )

    with st.expander("Catálogos de conceptos para movimientos nuevos", expanded=False):
        st.caption("Separa cada concepto con punto y coma. Los ingresos y egresos usan listas independientes.")
        concept_a, concept_b = st.columns(2)
        with concept_a:
            income_concepts_text = st.text_area(
                "Conceptos de ingresos",
                value="TRANSFER SPEI RECIBIDA; TRANSFERENCIA RECIBIDA; TRANSFERENCIA INT; COBRO A CLIENTE; PAGO RECIBIDO; DEPOSITO EN EFECTIVO; ABONO POR SERVICIOS",
                height=130,
                key="linked_income_concepts",
            )
        with concept_b:
            expense_concepts_text = st.text_area(
                "Conceptos de egresos",
                value="TRANSFER SPEI ENVIADA; TRANSFERENCIA ENVIADA; TRANSFERENCIA INT; PAGO A PROVEEDOR; PAGO DE SERVICIOS; COBRO DE COMISION; RETIRO EN EFECTIVO; PAGO DE IMPUESTOS",
                height=130,
                key="linked_expense_concepts",
            )
        linked_income_concepts = parse_semicolon_catalog(income_concepts_text, INCOME_CONCEPTS)
        linked_expense_concepts = parse_semicolon_catalog(expense_concepts_text, EXPENSE_CONCEPTS)
        st.caption(f"Se usarán {len(linked_income_concepts)} conceptos de ingreso y {len(linked_expense_concepts)} de egreso.")

    with st.expander("Bancos para transferencias ficticias de todos los formatos", expanded=False):
        selected_banamex_banks = st.multiselect(
            "Bancos permitidos",
            options=list(BANAMEX_TRANSFER_BANKS),
            default=list(BANAMEX_TRANSFER_BANKS),
            key="linked_banamex_transfer_banks",
            help="Los bancos que no selecciones quedarán excluidos de los movimientos ficticios.",
        )
        excluded_banamex_banks = [bank for bank in BANAMEX_TRANSFER_BANKS if bank not in selected_banamex_banks]
        st.caption(
            "Incluidos: " + (", ".join(selected_banamex_banks) if selected_banamex_banks else "NINGUNO")
            + "\n\nExcluidos: " + (", ".join(excluded_banamex_banks) if excluded_banamex_banks else "NINGUNO")
        )

    st.markdown("#### Configuración por periodo cronológico")
    st.caption("Los documentos se ordenan por fecha. En cada periodo puedes fijar el total de egresos, el saldo final o el saldo promedio mensual.")
    linked_configs = []
    for index in range(1, int(linked_document_count) + 1):
        with st.expander(f"Periodo cronológico {index}", expanded=index <= 3):
            objective_mode = st.selectbox(
                "Dato que deseas fijar", [TARGET_EXPENSES, TARGET_FINAL, TARGET_AVERAGE], key=f"linked_target_mode_{index}"
            )
            amount_mode = "Sugerencia automática"
            average_relation = AVERAGE_RELATION_AUTO
            average_tolerance_percent = 0.25
            income_timing = TIMING_BALANCED
            expense_timing = TIMING_BALANCED
            income_concentration = 80.0
            expense_concentration = 80.0
            income_window_start, income_window_end = 1, 31
            expense_window_start, expense_window_end = 1, 31
            max_movements_per_day = 8
            business_days_only = False
            if objective_mode == TARGET_AVERAGE:
                amount_mode = st.radio(
                    "Forma de definir los montos",
                    ["Sugerencia automática", "Fijar ingreso mensual y calcular egresos", "Montos manuales"],
                    horizontal=True,
                    key=f"linked_amount_mode_{index}",
                )
                relation_a, relation_b = st.columns(2)
                with relation_a:
                    average_relation = st.selectbox(
                        "Relación deseada entre promedio y saldo final",
                        AVERAGE_RELATIONS,
                        key=f"linked_average_relation_{index}",
                        help="Determina si se adelantan ingresos, se retrasan egresos o se busca una distribución equilibrada.",
                    )
                with relation_b:
                    average_tolerance_percent = st.number_input(
                        "Tolerancia del promedio (%)",
                        min_value=0.0,
                        max_value=10.0,
                        value=0.25,
                        step=0.05,
                        key=f"linked_average_tolerance_{index}",
                    )
                with st.expander("Calendario y concentración de movimientos nuevos", expanded=False):
                    custom_calendar = st.checkbox(
                        "Definir manualmente las ventanas de fechas",
                        value=False,
                        key=f"linked_custom_calendar_{index}",
                    )
                    calendar_a, calendar_b, calendar_c = st.columns(3)
                    with calendar_a:
                        income_concentration = st.number_input(
                            "Ingresos dentro de su ventana (%)", min_value=0.0, max_value=100.0,
                            value=80.0, step=5.0, key=f"linked_income_concentration_{index}",
                        )
                    with calendar_b:
                        expense_concentration = st.number_input(
                            "Egresos dentro de su ventana (%)", min_value=0.0, max_value=100.0,
                            value=80.0, step=5.0, key=f"linked_expense_concentration_{index}",
                        )
                    with calendar_c:
                        max_movements_per_day = st.number_input(
                            "Máximo de movimientos nuevos por día", min_value=1, max_value=100,
                            value=8, step=1, key=f"linked_max_daily_movements_{index}",
                        )
                    if custom_calendar:
                        custom_a, custom_b = st.columns(2)
                        with custom_a:
                            income_timing = st.selectbox(
                                "Momento de ingresos", TIMING_OPTIONS, key=f"linked_income_timing_{index}"
                            )
                            income_days_a, income_days_b = st.columns(2)
                            with income_days_a:
                                income_window_start = st.number_input(
                                    "Día inicial ingresos", min_value=1, max_value=31, value=1,
                                    step=1, key=f"linked_income_window_start_{index}",
                                )
                            with income_days_b:
                                income_window_end = st.number_input(
                                    "Día final ingresos", min_value=1, max_value=31, value=5,
                                    step=1, key=f"linked_income_window_end_{index}",
                                )
                        with custom_b:
                            expense_timing = st.selectbox(
                                "Momento de egresos", TIMING_OPTIONS, key=f"linked_expense_timing_{index}"
                            )
                            expense_days_a, expense_days_b = st.columns(2)
                            with expense_days_a:
                                expense_window_start = st.number_input(
                                    "Día inicial egresos", min_value=1, max_value=31, value=25,
                                    step=1, key=f"linked_expense_window_start_{index}",
                                )
                            with expense_days_b:
                                expense_window_end = st.number_input(
                                    "Día final egresos", min_value=1, max_value=31, value=31,
                                    step=1, key=f"linked_expense_window_end_{index}",
                                )
                    else:
                        if average_relation == AVERAGE_RELATION_ABOVE:
                            income_timing, expense_timing = TIMING_EARLY, TIMING_LATE
                            income_window_start, income_window_end = 1, 5
                            expense_window_start, expense_window_end = 25, 31
                            st.info("Política: ingresos tempranos y egresos tardíos para elevar el promedio.")
                        elif average_relation == AVERAGE_RELATION_BELOW:
                            income_timing, expense_timing = TIMING_LATE, TIMING_EARLY
                            income_window_start, income_window_end = 25, 31
                            expense_window_start, expense_window_end = 1, 5
                            st.info("Política: egresos tempranos e ingresos tardíos para reducir el promedio.")
                        elif average_relation == AVERAGE_RELATION_NEAR:
                            income_timing = expense_timing = TIMING_BALANCED
                            st.info("Política: ingresos y egresos distribuidos durante todo el periodo.")
                        else:
                            st.info("El programa elegirá la política al comparar el promedio solicitado con el saldo final.")
                    business_days_only = st.checkbox(
                        "Colocar movimientos nuevos solamente en días hábiles",
                        value=False,
                        key=f"linked_business_days_only_{index}",
                    )
            suggest_counts = False
            if objective_mode == TARGET_AVERAGE and amount_mode == "Sugerencia automática":
                suggest_counts = st.checkbox(
                    "Sugerir también cuántos movimientos nuevos agregar",
                    value=True,
                    key=f"linked_auto_counts_{index}",
                    help="La cantidad se basa en la densidad y proporción de ingresos/egresos del documento original.",
                )
            row_a, row_b, row_c, row_d = st.columns(4)
            with row_a:
                minimum_count = 1 if objective_mode == TARGET_AVERAGE else 0
                income_number = st.number_input("Ingresos nuevos", min_value=minimum_count, value=max(3, minimum_count), step=1, key=f"linked_inc_count_{index}", disabled=suggest_counts)
                expense_number = st.number_input("Egresos nuevos", min_value=minimum_count, value=max(4, minimum_count), step=1, key=f"linked_exp_count_{index}", disabled=suggest_counts)
            with row_b:
                if objective_mode == TARGET_AVERAGE:
                    objective_value = money_input("Saldo promedio mensual deseado", min_value=0.0, value=float(linked_opening), key=f"linked_target_value_{index}")
                elif objective_mode == TARGET_FINAL and plan_enabled:
                    target_monthly_income = float(planned_monthly_incomes[index - 1])
                    income_total = 0.0
                    st.metric("Ingreso total asignado al periodo", f"${target_monthly_income:,.2f}")
                else:
                    income_total = money_input("Monto de ingresos nuevos", min_value=0.0, value=45000.0, key=f"linked_inc_total_{index}")
            if objective_mode == TARGET_AVERAGE:
                if amount_mode == "Sugerencia automática":
                    with row_c:
                        final_policy = st.radio(
                            "Saldo final para este promedio",
                            ["Sugerido por el programa", "Fijar saldo final"],
                            key=f"linked_average_final_policy_{index}",
                        )
                    with row_d:
                        minimum_allowed = money_input(
                            "Saldo diario mínimo permitido", min_value=0.0, value=0.0,
                            key=f"linked_minimum_balance_{index}",
                        )
                    if final_policy == "Fijar saldo final":
                        advanced_a, advanced_b = st.columns(2)
                        with advanced_a:
                            average_final = money_input("Saldo final deseado", min_value=0.0, value=float(linked_opening), key=f"linked_average_final_{index}")
                        with advanced_b:
                            st.metric("Cálculo combinado", "Promedio + saldo final")
                    else:
                        average_final = None
                    income_total = 0.0
                    manual_expense_total = None
                    target_monthly_income = float(planned_monthly_incomes[index - 1]) if plan_enabled else None
                    if plan_enabled:
                        st.info(
                            f"Plan activo: se respetarán ${target_monthly_income:,.2f} de ingresos totales en este periodo. "
                            "El saldo inicial no cuenta como ingreso; el programa calculará egresos, promedio y saldo final."
                        )
                    else:
                        st.info("El programa usará estas cantidades para sugerir montos y saldo final. Después podrás copiar toda la propuesta al modo manual y modificar cantidades o montos.")
                elif amount_mode == "Fijar ingreso mensual y calcular egresos":
                    average_final = None
                    minimum_allowed = 0.0
                    manual_expense_total = None
                    income_total = 0.0
                    with row_c:
                        if plan_enabled:
                            target_monthly_income = float(planned_monthly_incomes[index - 1])
                            st.metric("Ingreso total asignado al periodo", f"${target_monthly_income:,.2f}")
                        else:
                            target_monthly_income = money_input(
                                "Ingresos totales del mes",
                                min_value=0.0,
                                value=1250000.0,
                                key=f"linked_monthly_income_target_{index}",
                                help="Incluye los ingresos originales del PDF y los nuevos que agregará el programa.",
                            )
                    with row_d:
                        st.metric("Cálculo guiado", "Egreso y saldo final")
                    st.caption("La simulación calculará cuánto egreso es compatible con ese ingreso y con el saldo promedio mensual solicitado.")
                    guided_final_a, guided_final_b = st.columns(2)
                    with guided_final_a:
                        guided_final_enabled = st.checkbox(
                            "También fijar el saldo final",
                            value=False,
                            key=f"linked_guided_final_enabled_{index}",
                        )
                        if guided_final_enabled:
                            average_final = money_input(
                                "Saldo final deseado", min_value=0.0, value=float(linked_opening),
                                key=f"linked_guided_average_final_{index}",
                            )
                    with guided_final_b:
                        minimum_allowed = money_input(
                            "Saldo diario mínimo permitido", min_value=0.0, value=0.0,
                            key=f"linked_guided_minimum_balance_{index}",
                        )
                else:
                    average_final = None
                    minimum_allowed = 0.0
                    with row_c:
                        income_total = money_input("Monto total de ingresos nuevos", min_value=0.0, value=45000.0, key=f"linked_inc_total_{index}")
                    with row_d:
                        manual_expense_total = money_input("Monto total de egresos nuevos", min_value=0.0, value=38000.0, key=f"linked_manual_exp_total_{index}")
                    target_monthly_income = None
                    st.caption("El promedio mensual funcionará como referencia; el resumen mostrará el promedio realmente obtenido y su diferencia.")
            else:
                with row_d:
                    target_label = "Total de egresos nuevos" if objective_mode == TARGET_EXPENSES else "Saldo final deseado"
                    default_target = 38000.0 if objective_mode == TARGET_EXPENSES else 25000.0
                    if objective_mode == TARGET_FINAL and final_plan_enabled:
                        objective_value = float(planned_final_balances[index - 1])
                        st.metric("Saldo final asignado al periodo", f"${objective_value:,.2f}")
                    else:
                        objective_value = money_input(target_label, min_value=0.0, value=default_target, key=f"linked_target_value_{index}")
                average_final = None
                minimum_allowed = 0.0
                manual_expense_total = None
                if objective_mode != TARGET_FINAL or not plan_enabled:
                    target_monthly_income = None
                elif objective_mode == TARGET_FINAL:
                    st.caption(
                        "El plan variable fija los ingresos totales de este mes. El motor calculará los egresos "
                        "necesarios para alcanzar el saldo final deseado."
                    )
            linked_configs.append(LinkedPeriodConfig(
                income_count=int(income_number), expense_count=int(expense_number),
                total_income=float(income_total), target_mode=objective_mode, target_value=float(objective_value),
                require_income_at_least_expense=False,
                auto_calculate_amounts=(
                    objective_mode == TARGET_AVERAGE
                    and amount_mode == "Sugerencia automática"
                    and target_monthly_income is None
                ),
                average_target_final=float(average_final) if average_final is not None else None,
                minimum_balance=float(minimum_allowed),
                manual_total_expense=float(manual_expense_total) if manual_expense_total is not None else None,
                auto_calculate_counts=bool(suggest_counts),
                target_total_monthly_income=float(target_monthly_income) if target_monthly_income is not None else None,
                planned_average_income=float(plan_income_average) if plan_enabled and target_monthly_income is not None else None,
                planned_variation_percent=float(plan_income_variation) if plan_enabled and target_monthly_income is not None else None,
                planned_average_final=float(final_plan_average) if objective_mode == TARGET_FINAL and final_plan_enabled else None,
                planned_final_variation_percent=float(final_plan_variation) if objective_mode == TARGET_FINAL and final_plan_enabled else None,
                income_concepts=linked_income_concepts,
                expense_concepts=linked_expense_concepts,
                transfer_banks=tuple(selected_banamex_banks),
                average_relation=average_relation,
                average_tolerance_percent=float(average_tolerance_percent),
                income_timing=income_timing,
                expense_timing=expense_timing,
                income_concentration_percent=float(income_concentration),
                expense_concentration_percent=float(expense_concentration),
                income_window_start=int(income_window_start),
                income_window_end=int(income_window_end),
                expense_window_start=int(expense_window_start),
                expense_window_end=int(expense_window_end),
                max_movements_per_day=int(max_movements_per_day),
                business_days_only=bool(business_days_only),
            ))

    guided_income_mode = any(config.target_total_monthly_income is not None for config in linked_configs)
    if guided_income_mode:
        st.markdown("#### Simulación guiada en tiempo real")
        if len(linked_files or []) == int(linked_document_count):
            try:
                preview_bases = []
                for uploaded in linked_files:
                    extracted, detected = parse_pdf_movements_cached(uploaded.getvalue(), linked_bank_option)
                    if extracted.empty:
                        raise ValueError(f"No se encontraron movimientos en {uploaded.name}.")
                    preview_bases.append((uploaded.name, detected, extracted))
                preview_periods, preview_summary = link_statement_periods(
                    preview_bases, linked_configs, first_opening_balance=float(linked_opening), seed=int(linked_seed)
                )
                preview_validation = validate_linked_periods(preview_periods, preview_summary)
                preview_columns = [
                    "PERIODO", "INGRESO_PROMEDIO_DEL_PLAN", "VARIACION_MAXIMA_INGRESOS_PCT",
                    "INGRESO_MENSUAL_OBJETIVO", "INGRESOS_TOTALES", "EGRESOS_TOTALES",
                    "SALDO_FINAL_PROMEDIO_DEL_PLAN", "VARIACION_MAXIMA_SALDO_FINAL_PCT",
                    "SALDO_PROMEDIO_MENSUAL", "DIFERENCIA_OBJETIVO", "TOLERANCIA_OBJETIVO",
                    "SALDO_FINAL", "DIFERENCIA_PROMEDIO_MENOS_FINAL",
                    "RELACION_PROMEDIO_FINAL_OBTENIDA", "RELACION_PROMEDIO_FINAL_CUMPLIDA",
                    "PROMEDIO_MINIMO_ESTIMADO", "PROMEDIO_MAXIMO_ESTIMADO", "SALDO_MINIMO",
                ]
                st.dataframe(preview_summary[preview_columns], use_container_width=True, hide_index=True)
                if (
                    preview_validation["SALDOS_NO_NEGATIVOS"]
                    and preview_validation["OBJETIVOS_CUMPLIDOS"]
                    and preview_validation["RELACIONES_PROMEDIO_FINAL_CUMPLIDAS"]
                ):
                    st.success("La combinación es viable: alcanza el promedio, respeta la relación con el saldo final y no produce saldos negativos.")
                else:
                    st.warning("La combinación requiere ajustes; revisa el rango alcanzable, la relación promedio/final, el ingreso o las ventanas de fechas.")
            except Exception as exc:
                st.warning(f"La combinación actual no es viable: {exc}")
                show_feasibility_options(exc, "live_feasibility")
        else:
            st.info("Sube todos los PDF para activar la comprobación. Después, cada cambio de ingreso o promedio actualizará esta simulación.")

    if st.button("Analizar, mezclar y enlazar periodos", type="primary", use_container_width=True, key="build_linked_periods"):
        if len(linked_files or []) != int(linked_document_count):
            st.error(f"Debes subir exactamente {int(linked_document_count)} PDF, o cambiar el número de documentos.")
        else:
            try:
                bases = []
                for uploaded in linked_files:
                    extracted, detected = parse_pdf_movements_cached(uploaded.getvalue(), linked_bank_option)
                    if extracted.empty:
                        raise ValueError(f"No se encontraron movimientos en {uploaded.name}.")
                    bases.append((uploaded.name, detected, extracted))
                periods, linked_summary = link_statement_periods(
                    bases, linked_configs, first_opening_balance=float(linked_opening), seed=int(linked_seed)
                )
                validation = validate_linked_periods(periods, linked_summary)
                pdf_outputs, output_styles = [], []
                style_by_bank = {
                    "SCOTIABANK": "SCOTIABANK", "SANTANDER": "SANTANDER", "SANTANDER SELECT": "SANTANDER", "AZTECA": "AZTECA",
                    "HSBC": "HSBC", "BBVA": "BBVA LARGO", "BANAMEX MICUENTA": "BANAMEX", "BANAMEX MICUENTA (JUN-JUL)": "BANAMEX (JUN-JUL)", "CITIBANAMEX": "BANAMEX", "BANORTE": "BANORTE",
                }
                code_by_bank = {"SCOTIABANK": "044", "SANTANDER": "014", "AZTECA": "127", "HSBC": "021", "BBVA": "012", "BANAMEX MICUENTA": "002", "BANAMEX MICUENTA (JUN-JUL)": "002", "CITIBANAMEX": "002", "BANORTE": "072"}
                for period_index, period in enumerate(periods, start=1):
                    bank = str(linked_summary.iloc[period_index - 1]["BANCO"])
                    style = style_by_bank.get(bank, "DIDÁCTICO GENÉRICO")
                    identity = generate_identity(int(linked_seed) + period_index, bank_code=code_by_bank.get(bank, "999"))
                    pdf_columns = ["FECHA", "DESCRIPCION", "DEPOSITO", "RETIRO", "SALDO"]
                    pdf_metadata = (
                        "ORIGEN", PHYSICAL_LINES_COLUMN, "BANCO_TRANSFERENCIA",
                        "NOMBRE_CONTRAPARTE", "FORMATO_BANAMEX_MICUENTA",
                        "CUENTA_BANAMEX", "REFERENCIA_BANAMEX", "RASTREO_BANAMEX",
                    )
                    pdf_columns.extend(column for column in pdf_metadata if column in period.columns)
                    pdf_table = period[pdf_columns].copy()
                    start = pd.Timestamp(linked_summary.iloc[period_index - 1]["FECHA_INICIAL"]).date()
                    end = pd.Timestamp(linked_summary.iloc[period_index - 1]["FECHA_FINAL"]).date()
                    if style == "BANORTE":
                        pdf_output = generate_banorte_pdf_bytes(pdf_table)
                    elif style == "DIDÁCTICO GENÉRICO":
                        pdf_output = generate_themed_pdf_bytes(pdf_table, identity, start, end, "GENERICO")
                    else:
                        pdf_output = generate_movement_pdf_bytes(pdf_table, identity, start, end, style)
                    pdf_outputs.append(pdf_output)
                    output_styles.append(style)
                st.session_state["linked_periods"] = periods
                st.session_state["linked_summary"] = linked_summary
                st.session_state["linked_validation"] = validation
                st.session_state["linked_pdfs"] = pdf_outputs
                st.session_state["linked_styles"] = output_styles
                st.success(f"Los {len(periods)} periodos quedaron enlazados, calculados y auditados.")
            except Exception as exc:
                st.error(f"No se pudieron enlazar los periodos: {exc}")
                show_feasibility_options(exc, "build_feasibility")

    if "linked_periods" in st.session_state:
        periods = st.session_state["linked_periods"]
        linked_summary = st.session_state["linked_summary"]
        validation = st.session_state["linked_validation"]
        banks = linked_summary["BANCO"].astype(str).tolist()
        excel_periods = [prepare_excel_layout(period, linked_excel_layout, bank) for period, bank in zip(periods, banks)]
        linked_excel_bytes = build_linked_excel_bytes(excel_periods, linked_summary)
        st.dataframe(linked_summary, use_container_width=True, height=260)
        if (
            linked_summary["REPETICIONES_CONSECUTIVAS_BANCOS_NUEVOS"].fillna(0).eq(0).all()
            and linked_summary["NOMBRES_NUEVOS_UNICOS"].fillna(0).eq(
                linked_summary["MOVIMIENTOS_FICTICIOS"].fillna(0)
            ).all()
        ):
            st.success(
                "Variedad validada en todos los periodos: no hay bancos ficticios consecutivos iguales ni nombres ficticios repetidos."
            )
        if linked_summary["CATALOGO_NOMBRES_AMPLIADO"].fillna(False).any():
            st.warning(
                "Se agotaron nombres simples no utilizados y el programa amplió automáticamente el catálogo con nombres compuestos únicos."
            )
        proposal_mask = linked_summary.get("CALCULO_AUTOMATICO_MONTOS", pd.Series(False, index=linked_summary.index)).fillna(False)
        if "CALCULO_GUIADO_INGRESOS" in linked_summary:
            proposal_mask = proposal_mask | linked_summary["CALCULO_GUIADO_INGRESOS"].fillna(False)
        if proposal_mask.any():
            st.markdown("#### Propuesta automática para saldo promedio mensual")
            st.dataframe(
                linked_summary.loc[
                    proposal_mask,
                    ["PERIODO", "INGRESO_MENSUAL_OBJETIVO", "SALDO_INICIAL_SIMULADO", "INGRESOS_ORIGINALES", "INGRESOS_FICTICIOS_CANTIDAD", "INGRESOS_FICTICIOS", "INGRESOS_TOTALES", "EGRESOS_FICTICIOS_CANTIDAD", "EGRESOS_FICTICIOS", "SALDO_PROMEDIO_MENSUAL", "DIFERENCIA_OBJETIVO", "TOLERANCIA_OBJETIVO", "SALDO_FINAL", "DIFERENCIA_PROMEDIO_MENOS_FINAL", "RELACION_PROMEDIO_FINAL_OBTENIDA", "RELACION_PROMEDIO_FINAL_CUMPLIDA", "PROMEDIO_MINIMO_ESTIMADO", "PROMEDIO_MAXIMO_ESTIMADO", "SALDO_MINIMO"],
                ],
                use_container_width=True,
                hide_index=True,
            )
            st.button(
                "Copiar propuesta a montos manuales para modificarla",
                on_click=apply_linked_proposal_to_manual,
                args=(linked_summary.copy(),),
                use_container_width=True,
                key="apply_linked_proposal_manual",
            )
            automatic_rows = linked_summary[proposal_mask]
            if (automatic_rows["DIFERENCIA_OBJETIVO"].abs() > 0.05).any():
                st.warning(
                    "Con las condiciones actuales no existe una coincidencia exacta. Se muestra la propuesta segura más cercana; "
                    "puedes copiarla al modo manual y cambiar cantidades o montos."
                )
        if "CALCULO_MANUAL_MONTOS" in linked_summary and linked_summary["CALCULO_MANUAL_MONTOS"].fillna(False).any():
            manual_rows = linked_summary[linked_summary["CALCULO_MANUAL_MONTOS"].fillna(False)]
            if (manual_rows["DIFERENCIA_OBJETIVO"].abs() > 0.05).any():
                st.warning("Modo manual: el saldo promedio mensual es una referencia; revisa en el resumen la diferencia contra el promedio deseado.")
        if (
            validation["TODOS_CUADRAN"]
            and validation["CADENA_CUADRA"]
            and validation["OBJETIVOS_CUMPLIDOS"]
            and validation["SALDOS_NO_NEGATIVOS"]
            and validation["SALDOS_FINALES_NO_NEGATIVOS"]
            and validation["EGRESOS_CUBIERTOS"]
            and validation["RELACIONES_PROMEDIO_FINAL_CUMPLIDAS"]
        ):
            st.success("Validación completa: los periodos cuadran, no existen saldos negativos, los objetivos y relaciones promedio/final se cumplieron y cada saldo final pasa al periodo siguiente.")
        else:
            st.error("La validación detectó una diferencia; revisa el resumen antes de usar la salida.")
        if validation["SALDO_MINIMO_GLOBAL"] < 0:
            st.warning(f"La simulación contiene un saldo mínimo negativo de ${validation['SALDO_MINIMO_GLOBAL']:,.2f}. Puedes aumentar el saldo inicial o los ingresos.")
        st.download_button(
            "Descargar Excel maestro de auditoría", data=linked_excel_bytes,
            file_name=f"PERIODOS_ENLAZADOS_AUDITORIA_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
        )
        download_columns = st.columns(min(3, len(periods)))
        for index, (period, excel_period, pdf_output, bank) in enumerate(
            zip(periods, excel_periods, st.session_state["linked_pdfs"], banks), start=1
        ):
            safe_bank = re.sub(r"[^A-Z0-9]+", "_", bank.upper()).strip("_") or "BANCO"
            with download_columns[(index - 1) % len(download_columns)]:
                st.download_button(
                    f"PDF didáctico · periodo {index} · {bank}", data=pdf_output,
                    file_name=f"{safe_bank}_PERIODO_{index}_DIDACTICO.pdf", mime="application/pdf",
                    use_container_width=True, key=f"linked_pdf_download_{index}",
                )
                st.download_button(
                    f"Excel · periodo {index}", data=build_excel_bytes(excel_period, bank, make_summary(period)),
                    file_name=f"{safe_bank}_PERIODO_{index}_AUDITORIA.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True, key=f"linked_excel_download_{index}",
                )
                if "SCOTIA" in str(bank).upper():
                    chart_png = generate_scotiabank_chart_png(period)
                    st.download_button(
                        f"Gráfica PNG transparente · periodo {index}",
                        data=chart_png,
                        file_name=f"SCOTIABANK_PERIODO_{index}_GRAFICA_TRANSPARENTE.png",
                        mime="image/png",
                        use_container_width=True,
                        key=f"linked_scotia_chart_download_{index}",
                    )
        for index, (period, bank) in enumerate(zip(excel_periods, banks), start=1):
            with st.expander(f"Ver movimientos del periodo {index} · {bank}"):
                st.dataframe(period, use_container_width=True, height=380)


with tab5:
    st.subheader("Herramientas de revisión tipográfica")
    st.write("FontForge está instalado para trabajar los contornos. FontTools mantiene las fuentes; FontBakery comprueba su calidad técnica. Estas revisiones no certifican que una fuente sea idéntica al original.")
    forge_font = st.selectbox("Fuente para abrir en FontForge", ["Regular", "Escaneo editable"], key="select_forge_font")
    if st.button("Abrir fuente en FontForge", key="select_open_forge"):
        try:
            from pathlib import Path
            import subprocess
            executable = Path(r"C:\Program Files\FontForgeBuilds\bin\fontforge.exe")
            font_path = Path(__file__).with_name("CLC_Select_Escaneo_Editable.ttf" if forge_font == "Escaneo editable" else "CLC_Movimientos_Select_Regular.ttf")
            if not executable.exists():
                raise RuntimeError("No se encontró FontForge en su carpeta de instalación.")
            subprocess.Popen([str(executable), str(font_path)])
            st.info("Se abrió FontForge para trabajar la fuente. Guarda las pruebas con otro nombre para conservar la fuente del convertidor.")
        except Exception as error:
            st.error(str(error))
    if st.button("Revisar fuentes con FontBakery", key="select_check_fonts"):
        try:
            from select_tools import font_quality
            with st.spinner("Revisando fuentes localmente…"):
                quality = font_quality()
            if quality["exit_code"] == 0:
                st.success("Revisión terminada sin fallos según el perfil ejecutado.")
            else:
                st.warning("La revisión encontró incidencias. Descarga el informe; no se presentan estas fuentes como certificadas.")
            st.download_button("Descargar informe técnico", quality["report"], file_name="FONTBAKERY_SELECT.json", mime="application/json")
            with st.expander("Detalles técnicos"):
                st.text(quality["log"][-15000:])
        except Exception as error:
            st.error(str(error))
    st.markdown("### Comparar dos PDFs")
    st.caption("Los archivos deben tener hojas tamaño carta. El resultado marca diferencias de posición, forma y contenido; si los datos cambian, también aparecen como diferencias. No es un porcentaje de coincidencia tipográfica.")
    c1,c2 = st.columns(2)
    reference_pdf = c1.file_uploader("PDF original de referencia", type=["pdf"], key="select_compare_original")
    candidate_pdf = c2.file_uploader("PDF del ejercicio", type=["pdf"], key="select_compare_candidate")
    reference_page = c1.number_input("Hoja del original", min_value=1, value=1, key="select_compare_page_original")
    candidate_page = c2.number_input("Hoja del ejercicio", min_value=1, value=1, key="select_compare_page_candidate")
    compare_area = st.selectbox("Zona que se compara", ["Resumen", "Movimientos", "Pagina sin logo ni pie"], key="select_compare_area")
    if st.button("Comparar visualmente", disabled=reference_pdf is None or candidate_pdf is None, key="select_compare_run"):
        try:
            from select_tools import compare_visual
            with st.spinner("Comparando las imágenes de los PDFs…"):
                comparison = compare_visual(reference_pdf.getvalue(), candidate_pdf.getvalue(), int(reference_page), int(candidate_page), compare_area)
            st.write(f"Índice estructural de la zona: {comparison['score']:.3f}. Incluye diferencias de datos y distribución.")
            cols = st.columns(3)
            for col,key,label in zip(cols,["reference","candidate","difference"],["Original","Ejercicio","Diferencias en rojo"]):
                col.image(comparison[key], caption=label)
            st.download_button("Descargar mapa de diferencias", comparison["difference"], file_name="DIFERENCIAS_SELECT.png", mime="image/png")
        except Exception as error:
            st.error(str(error))
