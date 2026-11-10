"""Motor multibanco para enlazar periodos de ejercicios didacticos."""

from __future__ import annotations

import random
import re
from dataclasses import dataclass, replace
from datetime import date, datetime

import pandas as pd

from synthetic_engine import (
    EXPENSE_CONCEPTS,
    FIRST_NAMES,
    INCOME_CONCEPTS,
    LAST_NAMES,
    GenerationConfig,
    generate_movements,
    movement_checks,
)
from excel_layouts import PHYSICAL_LINES_COLUMN, derive_physical_lines, encode_physical_lines


TARGET_EXPENSES = "MONTO TOTAL DE EGRESOS"
TARGET_FINAL = "SALDO FINAL"
TARGET_AVERAGE = "SALDO PROMEDIO MENSUAL"
LEGACY_TARGET_AVERAGE = "SALDO PROMEDIO DIARIO"
TARGET_MODES = (TARGET_EXPENSES, TARGET_FINAL, TARGET_AVERAGE, LEGACY_TARGET_AVERAGE)

AVERAGE_RELATION_AUTO = "AUTOMÁTICO SEGÚN OBJETIVO"
AVERAGE_RELATION_ABOVE = "PROMEDIO MAYOR QUE SALDO FINAL"
AVERAGE_RELATION_BELOW = "PROMEDIO MENOR QUE SALDO FINAL"
AVERAGE_RELATION_NEAR = "PROMEDIO CERCANO AL SALDO FINAL"
AVERAGE_RELATIONS = (
    AVERAGE_RELATION_AUTO,
    AVERAGE_RELATION_ABOVE,
    AVERAGE_RELATION_BELOW,
    AVERAGE_RELATION_NEAR,
)

TIMING_EARLY = "TEMPRANO"
TIMING_BALANCED = "EQUILIBRADO"
TIMING_LATE = "TARDÍO"
TIMING_OPTIONS = (TIMING_EARLY, TIMING_BALANCED, TIMING_LATE)

BANAMEX_TRANSFER_BANKS = (
    "BBVA MEXICO", "SANTANDER", "BANORTE", "HSBC", "SCOTIABANK",
    "BANCO AZTECA", "BANBAJIO", "INBURSA", "AFIRME", "BANCOPPEL",
)

BANAMEX_BANK_CLABE_CODES = {
    "BBVA MEXICO": "012", "SANTANDER": "014", "BANORTE": "072",
    "HSBC": "021", "SCOTIABANK": "044", "BANCO AZTECA": "127",
    "BANBAJIO": "030", "INBURSA": "036", "AFIRME": "062",
    "BANCOPPEL": "137",
}


def _is_banamex_micuenta_bank(value):
    return str(value or "").upper().startswith("BANAMEX MICUENTA")

MONTHS = {
    "ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DIC": 12,
    "JAN": 1, "APR": 4, "AUG": 8, "DEC": 12,
}


class LinkedFeasibilityError(ValueError):
    """Error recuperable que incluye alternativas numericas para la interfaz."""

    def __init__(self, message, options=None):
        super().__init__(message)
        self.options = dict(options or {})


@dataclass(frozen=True)
class LinkedPeriodConfig:
    income_count: int
    expense_count: int
    total_income: float
    target_mode: str = TARGET_EXPENSES
    target_value: float = 0.0
    require_income_at_least_expense: bool = False
    auto_calculate_amounts: bool = False
    average_target_final: float | None = None
    minimum_balance: float = 0.0
    manual_total_expense: float | None = None
    auto_calculate_counts: bool = False
    target_total_monthly_income: float | None = None
    approximate_monthly_income: bool = False
    planned_average_income: float | None = None
    planned_variation_percent: float | None = None
    planned_average_final: float | None = None
    planned_final_variation_percent: float | None = None
    income_concepts: tuple[str, ...] | None = None
    expense_concepts: tuple[str, ...] | None = None
    transfer_banks: tuple[str, ...] | None = None
    excluded_counterparty_names: tuple[str, ...] = ()
    previously_used_transfer_banks: tuple[str, ...] = ()
    average_relation: str = AVERAGE_RELATION_AUTO
    average_tolerance_percent: float = 0.25
    income_timing: str = TIMING_BALANCED
    expense_timing: str = TIMING_BALANCED
    income_concentration_percent: float = 80.0
    expense_concentration_percent: float = 80.0
    income_window_start: int = 1
    income_window_end: int = 31
    expense_window_start: int = 1
    expense_window_end: int = 31
    max_movements_per_day: int = 8
    business_days_only: bool = False

    def __post_init__(self):
        # Compatibilidad con la firma antigua de cuatro argumentos, donde el
        # cuarto valor era directamente el total de egresos.
        if isinstance(self.target_mode, (int, float)):
            object.__setattr__(self, "target_value", float(self.target_mode))
            object.__setattr__(self, "target_mode", TARGET_EXPENSES)
        elif self.target_mode == LEGACY_TARGET_AVERAGE:
            object.__setattr__(self, "target_mode", TARGET_AVERAGE)

    @property
    def total_expense(self) -> float:
        """Compatibilidad de lectura con versiones anteriores."""
        return self.target_value if self.target_mode == TARGET_EXPENSES else 0.0


def parse_statement_date(value, fallback_year: int | None = None) -> pd.Timestamp:
    if isinstance(value, pd.Timestamp):
        return value.normalize()
    if isinstance(value, (datetime, date)):
        return pd.Timestamp(value).normalize()
    text = str(value or "").strip().upper()
    if re.match(r"^\d{4}-\d{2}-\d{2}(?:[ T].*)?$", text):
        parsed_iso = pd.to_datetime(value, errors="coerce", yearfirst=True)
        if not pd.isna(parsed_iso):
            return pd.Timestamp(parsed_iso).normalize()
    match = re.match(r"^(\d{1,2})[-/ ]([A-Z]{3})[-/ ]?(\d{2,4})?$", text)
    if match:
        day = int(match.group(1))
        month = MONTHS.get(match.group(2))
        raw_year = match.group(3)
        year = int(raw_year) if raw_year else int(fallback_year or datetime.now().year)
        if year < 100:
            year += 2000
        if month:
            return pd.Timestamp(year=year, month=month, day=day)
    parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
    if pd.isna(parsed):
        raise ValueError(f"No se pudo interpretar la fecha: {value}")
    return pd.Timestamp(parsed).normalize()


def _transfer_type(description: str, deposit: float, withdrawal: float) -> str:
    text = str(description).upper()
    if not any(word in text for word in ("SPEI", "TRANSFER", "TRASPASO")):
        return "OTRO"
    if float(deposit) > 0.005:
        return "RECIBIDA"
    if float(withdrawal) > 0.005:
        return "ENVIADA"
    return "TRANSFERENCIA"


def _normalize_base(base: pd.DataFrame, document_name: str):
    required = ["FECHA", "DESCRIPCION", "DEPOSITO", "RETIRO", "SALDO"]
    missing = [column for column in required if column not in base.columns]
    if missing:
        raise ValueError(f"Faltan columnas en {document_name}: {', '.join(missing)}")
    period_start_attr = base.attrs.get("period_start")
    period_end_attr = base.attrs.get("period_end")
    selected = required + ([PHYSICAL_LINES_COLUMN] if PHYSICAL_LINES_COLUMN in base.columns else [])
    selected += [c for c in ("FECHA_LIQ_BBVA", "AJUSTE_LIQ_BBVA") if c in base.columns]
    work = base[selected].copy()
    work["DEPOSITO"] = pd.to_numeric(work["DEPOSITO"], errors="coerce").fillna(0.0)
    work["RETIRO"] = pd.to_numeric(work["RETIRO"], errors="coerce").fillna(0.0)
    work["SALDO"] = pd.to_numeric(work["SALDO"], errors="coerce")
    opening_mask = work["DESCRIPCION"].astype(str).str.upper().str.strip().eq("SALDO ANTERIOR")
    opening_rows = work[opening_mask]
    opening = float(opening_rows["SALDO"].dropna().iloc[0]) if not opening_rows.empty else 0.0
    work = work[~opening_mask].copy()
    if work.empty:
        # El Generador didactico tambien usa este motor sin archivo base. En
        # ese caso el periodo viene definido explicitamente en los atributos y
        # no existen movimientos originales que conservar.
        if period_start_attr is None or period_end_attr is None:
            raise ValueError(f"{document_name} no contiene movimientos para enlazar.")
        period_start = parse_statement_date(period_start_attr).normalize()
        period_end = parse_statement_date(period_end_attr).normalize()
        if period_end < period_start:
            raise ValueError("La fecha final no puede ser anterior a la fecha inicial.")
        work["SALDO_ORIGINAL"] = pd.Series(dtype=float)
        work["ORIGEN"] = pd.Series(dtype=str)
        work["DOCUMENTO"] = pd.Series(dtype=str)
        work["ID_MOVIMIENTO"] = pd.Series(dtype=str)
        work["ORDEN_ORIGINAL"] = pd.Series(dtype=int)
        work["TIPO_TRANSFERENCIA"] = pd.Series(dtype=str)
        return work, opening, period_start, period_end
    fallback_year = datetime.now().year
    work["FECHA"] = [parse_statement_date(value, fallback_year) for value in work["FECHA"]]
    period_start = parse_statement_date(period_start_attr) if period_start_attr is not None else work["FECHA"].min()
    period_end = parse_statement_date(period_end_attr) if period_end_attr is not None else work["FECHA"].max()
    period_start = min(period_start, work["FECHA"].min())
    period_end = max(period_end, work["FECHA"].max())
    work["SALDO_ORIGINAL"] = work["SALDO"]
    work["ORIGEN"] = "ORIGINAL"
    work["DOCUMENTO"] = document_name
    work["ID_MOVIMIENTO"] = [f"ORI-{index + 1:04d}" for index in range(len(work))]
    work["ORDEN_ORIGINAL"] = range(len(work))
    work["TIPO_TRANSFERENCIA"] = [
        _transfer_type(row.DESCRIPCION, row.DEPOSITO, row.RETIRO) for row in work.itertuples()
    ]
    return work, opening, period_start.normalize(), period_end.normalize()


def _validate_config(config: LinkedPeriodConfig):
    if config.target_mode not in TARGET_MODES:
        raise ValueError(f"Objetivo no reconocido: {config.target_mode}")
    if config.income_count < 0 or config.expense_count < 0:
        raise ValueError("Las cantidades de movimientos no pueden ser negativas.")
    if config.total_income < 0 or config.target_value < 0:
        raise ValueError("Los montos y saldos objetivo no pueden ser negativos.")
    if config.income_count == 0 and config.total_income > 0.005:
        raise ValueError("Un monto de ingresos positivo requiere al menos un ingreso.")


def _normalized_name(value):
    return re.sub(r"\s+", " ", re.sub(r"[^A-ZÁÉÍÓÚÜÑ ]", " ", str(value or "").upper())).strip()


def _counterparties_in(dataframe):
    names = set()
    patterns = (
        r"POR ORDEN DE\s+(.+?)(?=\s+(?:CTA\.?\s*ORDENANTE|CTA\.ORDENANTE|REF\.|RASTREO|$))",
        r"AL BENEF\.?\s+(.+?)(?=\s+(?:\(DATO|NO VERIFICADO|CTA\.?\s*BENEFICIARIO|CUENTA:|$))",
        r"(?:ORDENANTE|BENEFICIARIO):\s*(.+?)(?=\s+CUENTA:|$)",
    )
    for description in dataframe.get("DESCRIPCION", pd.Series(dtype=str)).astype(str):
        for pattern in patterns:
            match = re.search(pattern, description, re.I)
            if match:
                name = _normalized_name(match.group(1))
                if len(name) >= 5:
                    names.add(name)
    return names


def _transfer_banks_in(dataframe):
    found = set()
    descriptions = " ".join(dataframe.get("DESCRIPCION", pd.Series(dtype=str)).astype(str)).upper()
    for bank in BANAMEX_TRANSFER_BANKS:
        if re.search(rf"PAGO (?:RECIBIDO DE|INTERBANCARIO A)\s+{re.escape(bank)}\b", descriptions):
            found.add(bank)
    return found


def _assign_unique_counterparties(generated, excluded_names, seed):
    output = generated.copy()
    excluded = {_normalized_name(name) for name in excluded_names if _normalized_name(name)}
    candidates = [
        f"{first} {paternal} {maternal}"
        for first in FIRST_NAMES
        for paternal in LAST_NAMES
        for maternal in LAST_NAMES
        if paternal != maternal
    ]
    rng = random.Random(int(seed) + 65537)
    rng.shuffle(candidates)
    available = [name for name in candidates if _normalized_name(name) not in excluded]
    requested = len(output)
    exhausted = requested > len(available)
    if exhausted:
        expanded = [
            f"{first} {second} {paternal} {maternal}"
            for first in FIRST_NAMES for second in FIRST_NAMES
            for paternal in LAST_NAMES for maternal in LAST_NAMES
            if first != second and paternal != maternal
        ]
        rng.shuffle(expanded)
        available.extend(name for name in expanded if _normalized_name(name) not in excluded)
    if requested > len(available):
        raise ValueError("Se agotaron todas las combinaciones de nombres ficticios disponibles.")
    selected = available[:requested]
    descriptions = []
    for description, name in zip(output["DESCRIPCION"].astype(str), selected):
        revised = re.sub(
            r"((?:ORDENANTE|BENEFICIARIO):\s*)(.+?)(\s+CUENTA:)",
            lambda match: match.group(1) + name + match.group(3),
            description,
            count=1,
            flags=re.I,
        )
        descriptions.append(revised)
    output["DESCRIPCION"] = descriptions
    output["NOMBRE_CONTRAPARTE"] = selected
    output["CATALOGO_NOMBRES_AMPLIADO"] = exhausted
    return output


def _generated_skeleton(start: date, end: date, config: LinkedPeriodConfig, seed: int, document_name: str, bank: str):
    _validate_config(config)
    bank = str(bank or "").upper()
    income_concepts = config.income_concepts
    expense_concepts = config.expense_concepts
    allowed_banks = tuple(
        dict.fromkeys(BANAMEX_TRANSFER_BANKS if config.transfer_banks is None else config.transfer_banks)
    )
    if (config.income_count + config.expense_count) and not allowed_banks:
        raise ValueError("Selecciona al menos un banco permitido para los movimientos ficticios.")
    if _is_banamex_micuenta_bank(bank):
        # La estructura visual sigue siendo MiCuenta, pero el concepto elegido
        # por el usuario no se sustituye silenciosamente.
        if income_concepts is None:
            income_concepts = ("TRANSFERENCIA RECIBIDA",)
        if expense_concepts is None:
            expense_concepts = ("TRANSFERENCIA ENVIADA",)
    provisional_expense = float(max(config.expense_count, 0) * 1000)
    safe_opening = max(1_000_000_000.0, provisional_expense + config.total_income + 1_000_000.0)
    generated = generate_movements(GenerationConfig(
        start_date=start,
        end_date=end,
        income_count=int(config.income_count),
        expense_count=int(config.expense_count),
        total_income=float(config.total_income),
        opening_balance=safe_opening,
        target_mode="SALDO FINAL",
        target_balance=round(safe_opening + config.total_income - provisional_expense, 2),
        seed=int(seed),
        income_concepts=income_concepts,
        expense_concepts=expense_concepts,
        transfer_banks=allowed_banks,
    ))
    generated = generated[~generated["DESCRIPCION"].astype(str).str.upper().eq("SALDO ANTERIOR")].copy()
    generated["SALDO_ORIGINAL"] = pd.NA
    generated["ORIGEN"] = "FICTICIO"
    generated["DOCUMENTO"] = document_name
    generated["ID_MOVIMIENTO"] = [f"FIC-{index + 1:04d}" for index in range(len(generated))]
    generated["ORDEN_ORIGINAL"] = pd.NA
    generated["TIPO_TRANSFERENCIA"] = [
        _transfer_type(row.DESCRIPCION, row.DEPOSITO, row.RETIRO) for row in generated.itertuples()
    ]
    if _is_banamex_micuenta_bank(bank):
        bank_rng = random.Random(int(seed) + 31337)
        unused_banks = [bank for bank in allowed_banks if bank not in set(config.previously_used_transfer_banks)]
        reused_banks = [bank for bank in allowed_banks if bank in set(config.previously_used_transfer_banks)]
        bank_rng.shuffle(unused_banks)
        bank_rng.shuffle(reused_banks)
        bank_cycle = unused_banks + reused_banks
        generated["BANCO_TRANSFERENCIA"] = [bank_cycle[index % len(bank_cycle)] for index in range(len(generated))]
        generated["BANCOS_AUTORIZADOS_YA_USADOS"] = not bool(unused_banks)
        identifier_rng = random.Random(int(seed) * 104729 + 8191)
        used_accounts, used_references, used_traces = set(), set(), set()

        def unique_digits(length, used, prefix=""):
            while True:
                value = prefix + "".join(str(identifier_rng.randrange(10)) for _ in range(length - len(prefix)))
                if value not in used:
                    used.add(value)
                    return value

        def clabe_for(transfer_bank):
            code = BANAMEX_BANK_CLABE_CODES.get(transfer_bank, "002")
            first_seventeen = code + unique_digits(14, set())
            weights = (3, 7, 1)
            check_digit = (10 - sum(int(digit) * weights[index % 3] for index, digit in enumerate(first_seventeen)) % 10) % 10
            clabe = first_seventeen + str(check_digit)
            if clabe in used_accounts:
                return clabe_for(transfer_bank)
            used_accounts.add(clabe)
            return clabe

        accounts, references, traces = [], [], []
        for row_index, row in generated.reset_index(drop=True).iterrows():
            transfer_bank = str(row["BANCO_TRANSFERENCIA"])
            movement_date = pd.Timestamp(row["FECHA"])
            incoming = float(row.get("DEPOSITO", 0.0) or 0.0) > 0.005
            accounts.append(clabe_for(transfer_bank))
            references.append(unique_digits(7, used_references))
            if not incoming:
                # Los pagos interbancarios del MiCuenta de referencia usan 18
                # digitos, normalmente con el prefijo 0859.
                traces.append(unique_digits(18, used_traces, "0859"))
            elif transfer_bank == "BBVA MEXICO":
                # Ejemplo real: 2026060740014BMOVP000450275950 (30 caracteres).
                prefix = movement_date.strftime("%Y%m%d") + "40014BMOVP"
                traces.append(unique_digits(30, used_traces, prefix))
            elif transfer_bank == "HSBC":
                # Ejemplo real: HSB5810693 (10 caracteres).
                traces.append(unique_digits(10, used_traces, "HSB"))
            else:
                # Para otros bancos se conserva una clave numerica compacta
                # de 18 posiciones, sin series consecutivas entre movimientos.
                prefix = BANAMEX_BANK_CLABE_CODES.get(transfer_bank, "002")
                traces.append(unique_digits(18, used_traces, prefix))
        generated["CUENTA_BANAMEX"] = accounts
        generated["REFERENCIA_BANAMEX"] = references
        generated["RASTREO_BANAMEX"] = traces
        generated["FORMATO_BANAMEX_MICUENTA"] = True
    # La unicidad de contrapartes es una regla general, no exclusiva de Banamex.
    generated = _assign_unique_counterparties(generated, config.excluded_counterparty_names, seed)
    generated[PHYSICAL_LINES_COLUMN] = [
        encode_physical_lines(derive_physical_lines(str(description), bank))
        for description in generated["DESCRIPCION"]
    ]
    return generated


def _rebalance_final_transfer_banks(period, config, seed, bank_style):
    """Asigna bancos a todo movimiento ficticio, sin vecinos iguales."""
    output = period.copy()
    allowed = list(dict.fromkeys(BANAMEX_TRANSFER_BANKS if config.transfer_banks is None else config.transfer_banks))
    if not allowed:
        raise ValueError("Selecciona al menos un banco permitido para los movimientos ficticios.")
    fictitious_indexes = output.index[output["ORIGEN"].astype(str).eq("FICTICIO")].tolist()
    if not fictitious_indexes:
        return output
    rng = random.Random(int(seed) + 424242)
    prior = set(config.previously_used_transfer_banks)
    preferred = [bank for bank in allowed if bank not in prior]
    repeated = [bank for bank in allowed if bank in prior]
    rng.shuffle(preferred)
    rng.shuffle(repeated)
    bank_order = preferred + repeated

    direction_for = {
        index: "INGRESO" if float(output.at[index, "DEPOSITO"] or 0.0) > 0.005 else "EGRESO"
        for index in fictitious_indexes
    }
    remaining = {}
    income_extra = 0
    for direction in ("INGRESO", "EGRESO"):
        count = sum(value == direction for value in direction_for.values())
        base, extra = divmod(count, len(bank_order))
        if direction == "INGRESO":
            direction_order = bank_order[:]
            income_extra = extra
        else:
            offset = income_extra % len(bank_order)
            direction_order = bank_order[offset:] + bank_order[:offset]
        remaining[direction] = {
            bank: base + (1 if position < extra else 0)
            for position, bank in enumerate(direction_order)
        }

    def original_bank_at(index):
        description = str(output.at[index, "DESCRIPCION"]).upper()
        for candidate in allowed:
            if re.search(rf"PAGO (?:RECIBIDO DE|INTERBANCARIO A)\s+{re.escape(candidate)}\b", description):
                return candidate
        return None

    events = []
    for index in output.index:
        if index in direction_for:
            events.append(("FICTICIO", index, direction_for[index]))
        else:
            # Un movimiento no bancario rompe la consecutividad visual.
            events.append(("ORIGINAL", index, original_bank_at(index)))

    income_initial = tuple(remaining["INGRESO"][bank] for bank in bank_order)
    expense_initial = tuple(remaining["EGRESO"][bank] for bank in bank_order)
    memo = {}

    def solve(position, previous_visual_bank, previous_new_bank, income_left, expense_left):
        key = (position, previous_visual_bank, previous_new_bank, income_left, expense_left)
        if key in memo:
            return memo[key]
        if position >= len(events):
            result = () if not any(income_left) and not any(expense_left) else None
            memo[key] = result
            return result
        kind, index, value = events[position]
        if kind == "ORIGINAL":
            # Un original bancario participa en la vecindad visual, pero nunca
            # borra la memoria del banco nuevo anterior. Asi tampoco se repite
            # un banco entre dos movimientos nuevos aunque haya originales o
            # movimientos no bancarios intercalados.
            result = solve(position + 1, value, previous_new_bank, income_left, expense_left)
            memo[key] = result
            return result
        direction = value
        counts = income_left if direction == "INGRESO" else expense_left
        next_bank = None
        if position + 1 < len(events) and events[position + 1][0] == "ORIGINAL":
            next_bank = events[position + 1][2]
        candidates = [
            i for i, count in enumerate(counts)
            if count > 0
            and bank_order[i] != previous_visual_bank
            and bank_order[i] != previous_new_bank
            and bank_order[i] != next_bank
        ]
        candidates.sort(key=lambda i: (-counts[i], i))
        for bank_position in candidates:
            revised = list(counts)
            revised[bank_position] -= 1
            next_income = tuple(revised) if direction == "INGRESO" else income_left
            next_expense = tuple(revised) if direction == "EGRESO" else expense_left
            tail = solve(
                position + 1,
                bank_order[bank_position],
                bank_order[bank_position],
                next_income,
                next_expense,
            )
            if tail is not None:
                result = ((index, bank_order[bank_position]),) + tail
                memo[key] = result
                return result
        memo[key] = None
        return None

    solution = solve(0, None, None, income_initial, expense_initial)
    if solution is None:
        raise ValueError(
            "No fue posible distribuir los bancos sin repeticiones consecutivas. "
            "Autoriza al menos dos bancos cuando exista más de un movimiento ficticio."
        )
    assignments = dict(solution)

    identifier_rng = random.Random(int(seed) * 130363 + 97)
    used_accounts, used_traces = set(), set()

    def unique_value(length, used, prefix=""):
        while True:
            value = prefix + "".join(str(identifier_rng.randrange(10)) for _ in range(length - len(prefix)))
            if value not in used:
                used.add(value)
                return value

    def clabe_for(transfer_bank):
        code = BANAMEX_BANK_CLABE_CODES.get(transfer_bank, "002")
        while True:
            first_seventeen = code + "".join(str(identifier_rng.randrange(10)) for _ in range(14))
            weights = (3, 7, 1)
            check = (10 - sum(int(digit) * weights[i % 3] for i, digit in enumerate(first_seventeen)) % 10) % 10
            clabe = first_seventeen + str(check)
            if clabe not in used_accounts:
                used_accounts.add(clabe)
                return clabe

    for index, transfer_bank in assignments.items():
        output.at[index, "BANCO_TRANSFERENCIA"] = transfer_bank
        if not _is_banamex_micuenta_bank(bank_style):
            continue
        output.at[index, "CUENTA_BANAMEX"] = clabe_for(transfer_bank)
        incoming = direction_for[index] == "INGRESO"
        movement_date = pd.Timestamp(output.at[index, "FECHA"])
        if not incoming:
            trace = unique_value(18, used_traces, "0859")
        elif transfer_bank == "BBVA MEXICO":
            trace = unique_value(30, used_traces, movement_date.strftime("%Y%m%d") + "40014BMOVP")
        elif transfer_bank == "HSBC":
            trace = unique_value(10, used_traces, "HSB")
        else:
            trace = unique_value(18, used_traces, BANAMEX_BANK_CLABE_CODES.get(transfer_bank, "002"))
        output.at[index, "RASTREO_BANAMEX"] = trace
    assigned_sequence = [str(output.at[index, "BANCO_TRANSFERENCIA"]).strip().upper() for index in fictitious_indexes]
    repetitions = sum(left == right for left, right in zip(assigned_sequence, assigned_sequence[1:]))
    if repetitions:
        raise ValueError("La validación detectó bancos repetidos en movimientos ficticios consecutivos.")
    return output


def _rescale_expenses(generated: pd.DataFrame, total_expense: float) -> pd.DataFrame:
    output = generated.copy()
    mask = pd.to_numeric(output["RETIRO"], errors="coerce").fillna(0.0) > 0.005
    count = int(mask.sum())
    total_expense = round(float(total_expense), 2)
    if count == 0:
        if total_expense > 0.005:
            raise ValueError("El objetivo requiere egresos, pero se seleccionaron cero egresos nuevos.")
        return output
    if total_expense < count:
        raise ValueError(f"El egreso calculado (${total_expense:,.2f}) es menor al minimo de $1.00 por cada uno de los {count} egresos.")
    current = pd.to_numeric(output.loc[mask, "RETIRO"], errors="coerce").fillna(0.0)
    shares = current / current.sum() if current.sum() else pd.Series([1 / count] * count, index=current.index)
    cents = int(round(total_expense * 100))
    raw = shares * cents
    allocated = raw.apply(int)
    remainder = cents - int(allocated.sum())
    order = (raw - allocated).sort_values(ascending=False).index.tolist()
    for index in order[:remainder]:
        allocated.loc[index] += 1
    output.loc[mask, "RETIRO"] = allocated / 100.0
    return output


def _mix_and_recalculate(original, generated, opening, seed, document_name, period_start):
    rng = random.Random(seed)
    original, generated = original.copy(), generated.copy()
    original["ORDEN_MEZCLA"] = [rng.random() for _ in range(len(original))]
    generated["ORDEN_MEZCLA"] = [rng.random() for _ in range(len(generated))]

    all_rows = pd.concat([original, generated], ignore_index=True, sort=False)
    all_rows["FECHA"] = pd.to_datetime(all_rows["FECHA"]).dt.normalize()
    final_balance = round(
        float(opening)
        + float(pd.to_numeric(all_rows["DEPOSITO"], errors="coerce").fillna(0.0).sum())
        - float(pd.to_numeric(all_rows["RETIRO"], errors="coerce").fillna(0.0).sum()),
        2,
    )
    if final_balance < -0.005:
        shortage = abs(final_balance)
        raise ValueError(
            f"El objetivo produciria un saldo final negativo de ${final_balance:,.2f}. "
            f"Faltan ${shortage:,.2f}: aumenta el saldo inicial o los ingresos, "
            "reduce los egresos o selecciona un saldo final/promedio mayor."
        )

    # Primero se calendarizan los movimientos originales y los ingresos
    # ficticios. Los egresos ficticios se insertan despues en la primera
    # posicion viable, para no consumir fondos que necesita un cargo original.
    synthetic_expense_mask = (
        all_rows["ORIGEN"].astype(str).eq("FICTICIO")
        & (pd.to_numeric(all_rows["RETIRO"], errors="coerce").fillna(0.0) > 0.005)
    )
    expense_rows = all_rows[synthetic_expense_mask].sort_values(["FECHA", "ORDEN_MEZCLA"], kind="stable").to_dict("records")
    base_rows = all_rows[~synthetic_expense_mask].sort_values(["FECHA", "ORDEN_MEZCLA"], kind="stable")
    pending, scheduled = base_rows.to_dict("records"), []
    available = round(float(opening), 2)
    while pending:
        row = pending.pop(0)
        deposit = float(row.get("DEPOSITO", 0.0) or 0.0)
        withdrawal = float(row.get("RETIRO", 0.0) or 0.0)
        while withdrawal > available + deposit + 0.005:
            future = next((i for i, candidate in enumerate(pending)
                if candidate.get("ORIGEN") == "FICTICIO" and float(candidate.get("DEPOSITO", 0.0) or 0.0) > 0.005), None)
            if future is None:
                break
            income = pending.pop(future)
            income["FECHA"] = row["FECHA"]
            income["ORDEN_MEZCLA"] = -1.0
            scheduled.append(income)
            available = round(available + float(income.get("DEPOSITO", 0.0) or 0.0), 2)
        if withdrawal > available + deposit + 0.005:
            shortage = round(withdrawal - available - deposit, 2)
            required_opening = round(float(opening) + shortage, 2)
            movement = str(row.get("DESCRIPCION", "MOVIMIENTO"))[:80]
            movement_date = pd.Timestamp(row["FECHA"]).strftime("%d/%m/%Y")
            raise ValueError(
                f"No hay liquidez suficiente para '{movement}' del {movement_date}. "
                f"Faltan ${shortage:,.2f}. Con los ingresos actuales, el saldo inicial "
                f"debe ser al menos ${required_opening:,.2f}."
            )
        scheduled.append(row)
        available = round(available + deposit - withdrawal, 2)

    def sequence_balances(rows):
        running = round(float(opening), 2)
        balances = []
        for candidate in rows:
            running = round(
                running
                + float(candidate.get("DEPOSITO", 0.0) or 0.0)
                - float(candidate.get("RETIRO", 0.0) or 0.0),
                2,
            )
            balances.append(running)
        return balances

    # Cada egreso ficticio conserva su fecha cuando es viable; si causaria un
    # sobregiro se desplaza al primer punto posterior con fondos suficientes.
    for expense in expense_rows:
        desired_date = pd.Timestamp(expense["FECHA"]).normalize()
        inserted = False
        for position in range(len(scheduled) + 1):
            previous_date = pd.Timestamp(scheduled[position - 1]["FECHA"]).normalize() if position else pd.Timestamp(period_start)
            next_date = pd.Timestamp(scheduled[position]["FECHA"]).normalize() if position < len(scheduled) else None
            candidate_date = max(desired_date, previous_date)
            if next_date is not None and candidate_date > next_date:
                continue
            candidate = expense.copy()
            candidate["FECHA"] = candidate_date
            trial = scheduled[:position] + [candidate] + scheduled[position:]
            if min(sequence_balances(trial), default=float(opening)) >= -0.005:
                scheduled = trial
                inserted = True
                break
        if not inserted:
            raise ValueError(
                "No existe una fecha viable para colocar todos los egresos sin generar "
                "un saldo negativo. Aumenta ingresos o saldo inicial, o reduce egresos."
            )

    calculated_balances = sequence_balances(scheduled)
    if calculated_balances and min(calculated_balances) < -0.005:
        raise ValueError("La validacion de liquidez detecto un saldo negativo; no se genero la salida.")
    output = [{
        "FECHA": pd.Timestamp(period_start), "DESCRIPCION": "SALDO ANTERIOR",
        "DEPOSITO": 0.0, "RETIRO": 0.0, "SALDO": round(float(opening), 2),
        "SALDO_ORIGINAL": pd.NA, "ORIGEN": "SISTEMA", "DOCUMENTO": document_name,
        "ID_MOVIMIENTO": "APERTURA", "TIPO_TRANSFERENCIA": "APERTURA", "ORDEN_ORIGINAL": pd.NA,
        "BANCO_TRANSFERENCIA": "", "FORMATO_BANAMEX_MICUENTA": False,
        "CUENTA_BANAMEX": "", "REFERENCIA_BANAMEX": "", "RASTREO_BANAMEX": "",
        "NOMBRE_CONTRAPARTE": "", "CONCEPTO_SOLICITADO": "", "CATALOGO_NOMBRES_AMPLIADO": False,
        "BANCOS_AUTORIZADOS_YA_USADOS": False,
        PHYSICAL_LINES_COLUMN: encode_physical_lines(["SALDO ANTERIOR"]),
    }]
    running = round(float(opening), 2)
    for row in scheduled:
        deposit = round(float(row.get("DEPOSITO", 0.0) or 0.0), 2)
        withdrawal = round(float(row.get("RETIRO", 0.0) or 0.0), 2)
        running = round(running + deposit - withdrawal, 2)
        output.append({
            "FECHA": pd.Timestamp(row["FECHA"]), "DESCRIPCION": str(row.get("DESCRIPCION", "")),
            "DEPOSITO": deposit, "RETIRO": withdrawal, "SALDO": running,
            "SALDO_ORIGINAL": row.get("SALDO_ORIGINAL", pd.NA), "ORIGEN": row.get("ORIGEN", ""),
            "DOCUMENTO": document_name, "ID_MOVIMIENTO": row.get("ID_MOVIMIENTO", ""),
            "TIPO_TRANSFERENCIA": row.get("TIPO_TRANSFERENCIA", "OTRO"),
            "ORDEN_ORIGINAL": row.get("ORDEN_ORIGINAL", pd.NA),
            "BANCO_TRANSFERENCIA": row.get("BANCO_TRANSFERENCIA", ""),
            "FORMATO_BANAMEX_MICUENTA": str(row.get("FORMATO_BANAMEX_MICUENTA", "")).upper() == "TRUE",
            "CUENTA_BANAMEX": row.get("CUENTA_BANAMEX", ""),
            "REFERENCIA_BANAMEX": row.get("REFERENCIA_BANAMEX", ""),
            "RASTREO_BANAMEX": row.get("RASTREO_BANAMEX", ""),
            "NOMBRE_CONTRAPARTE": row.get("NOMBRE_CONTRAPARTE", ""),
            "CONCEPTO_SOLICITADO": row.get("CONCEPTO_SOLICITADO", ""),
            "CATALOGO_NOMBRES_AMPLIADO": str(row.get("CATALOGO_NOMBRES_AMPLIADO", "")).upper() == "TRUE",
            "BANCOS_AUTORIZADOS_YA_USADOS": str(row.get("BANCOS_AUTORIZADOS_YA_USADOS", "")).upper() == "TRUE",
            PHYSICAL_LINES_COLUMN: row.get(PHYSICAL_LINES_COLUMN, encode_physical_lines([str(row.get("DESCRIPCION", ""))])),
        })
    result = pd.DataFrame(output)
    if float(result["SALDO"].min()) < -0.005 or float(result["SALDO"].iloc[-1]) < -0.005:
        raise ValueError("No se genero el documento porque uno de sus saldos resulto negativo.")
    return result


def daily_average_balance(period: pd.DataFrame, period_start=None, period_end=None) -> float:
    """Promedio de saldos al cierre de cada dia natural del periodo."""
    if period.empty:
        return 0.0
    dates = pd.to_datetime(period["FECHA"]).dt.normalize()
    start = parse_statement_date(period_start) if period_start is not None else dates.min()
    end = parse_statement_date(period_end) if period_end is not None else dates.max()
    if end < start:
        raise ValueError("La fecha final del periodo es anterior a la inicial.")
    opening = float(pd.to_numeric(period["SALDO"], errors="coerce").iloc[0])
    movements = period[~period["DESCRIPCION"].astype(str).str.upper().str.strip().eq("SALDO ANTERIOR")].copy()
    movements["FECHA"] = pd.to_datetime(movements["FECHA"]).dt.normalize()
    movements["NETO"] = pd.to_numeric(movements["DEPOSITO"], errors="coerce").fillna(0.0) - pd.to_numeric(movements["RETIRO"], errors="coerce").fillna(0.0)
    net_by_day = movements.groupby("FECHA")["NETO"].sum().to_dict()
    running, balances = opening, []
    for day in pd.date_range(start, end, freq="D"):
        running = round(running + float(net_by_day.get(day, 0.0)), 2)
        balances.append(running)
    return round(sum(balances) / len(balances), 2)


def _objective_tolerance(config: LinkedPeriodConfig) -> float:
    """Tolerancia monetaria del objetivo, expresada como porcentaje configurable."""
    percent = max(0.0, float(config.average_tolerance_percent or 0.0))
    return round(max(0.05, abs(float(config.target_value)) * percent / 100.0), 2)


def _relation_from_values(average: float, final: float, tolerance: float) -> str:
    difference = float(average) - float(final)
    if abs(difference) <= max(0.05, float(tolerance)):
        return AVERAGE_RELATION_NEAR
    return AVERAGE_RELATION_ABOVE if difference > 0 else AVERAGE_RELATION_BELOW


def _resolved_average_relation(config: LinkedPeriodConfig, opening: float) -> str:
    if config.average_relation in {AVERAGE_RELATION_ABOVE, AVERAGE_RELATION_BELOW, AVERAGE_RELATION_NEAR}:
        return config.average_relation
    comparison = config.average_target_final
    if comparison is None:
        comparison = opening
    return _relation_from_values(config.target_value, comparison, _objective_tolerance(config))


def _relation_matches(requested: str, achieved: str) -> bool:
    if requested == AVERAGE_RELATION_AUTO:
        return True
    return requested == achieved


def _automatic_timing(relation: str) -> tuple[str, str]:
    if relation == AVERAGE_RELATION_ABOVE:
        return TIMING_EARLY, TIMING_LATE
    if relation == AVERAGE_RELATION_BELOW:
        return TIMING_LATE, TIMING_EARLY
    return TIMING_BALANCED, TIMING_BALANCED


def _schedule_generated_dates(
    generated: pd.DataFrame,
    config: LinkedPeriodConfig,
    start: pd.Timestamp,
    end: pd.Timestamp,
    seed: int,
    opening: float,
    forced_relation: str | None = None,
) -> pd.DataFrame:
    """Distribuye únicamente fechas ficticias con concentración monetaria y cupo diario."""
    if generated.empty:
        return generated.copy()
    scheduled = generated.copy()
    calendar = list(pd.date_range(start, end, freq="D"))
    if config.business_days_only:
        business_calendar = [day for day in calendar if day.weekday() < 5]
        if business_calendar:
            calendar = business_calendar
    if not calendar:
        raise ValueError("El periodo no contiene fechas disponibles para movimientos ficticios.")

    relation = forced_relation or _resolved_average_relation(config, opening)
    automatic_income, automatic_expense = _automatic_timing(relation)
    use_automatic_windows = forced_relation is not None or (
        config.average_relation == AVERAGE_RELATION_AUTO
        and config.income_timing == TIMING_BALANCED
        and config.expense_timing == TIMING_BALANCED
        and int(config.income_window_start) == 1
        and int(config.income_window_end) >= 31
        and int(config.expense_window_start) == 1
        and int(config.expense_window_end) >= 31
    )
    income_timing = automatic_income if use_automatic_windows else config.income_timing
    expense_timing = automatic_expense if use_automatic_windows else config.expense_timing
    rng = random.Random(int(seed) + 1709)
    daily_counts: dict[pd.Timestamp, int] = {day: 0 for day in calendar}
    daily_limit = max(1, int(config.max_movements_per_day or 1))

    def default_window(timing: str) -> tuple[int, int]:
        count = len(calendar)
        if timing == TIMING_EARLY:
            return 1, min(5, count)
        if timing == TIMING_LATE:
            return max(1, count - 5), count
        return 1, count

    def available_window(timing: str, requested_start: int, requested_end: int, automatic: bool) -> list[pd.Timestamp]:
        if automatic:
            first, last = default_window(timing)
        else:
            first = max(1, min(int(requested_start), len(calendar)))
            last = max(first, min(int(requested_end), len(calendar)))
        return calendar[first - 1:last] or calendar

    def choose_day(candidates: list[pd.Timestamp], order: int, total: int) -> pd.Timestamp:
        usable = [day for day in candidates if daily_counts.get(day, 0) < daily_limit]
        if not usable:
            usable = [day for day in calendar if daily_counts.get(day, 0) < daily_limit]
        if not usable:
            raise ValueError(
                f"El máximo de {daily_limit} movimientos nuevos por día no alcanza para distribuir "
                f"{len(scheduled)} movimientos en el periodo."
            )
        target_position = round(order * (len(usable) - 1) / max(total - 1, 1))
        jitter = rng.choice([-1, 0, 0, 0, 1]) if len(usable) > 2 else 0
        target_position = max(0, min(len(usable) - 1, target_position + jitter))
        day = usable[target_position]
        daily_counts[day] = daily_counts.get(day, 0) + 1
        return day

    def schedule_flow(mask: pd.Series, amount_column: str, timing: str, concentration: float, first: int, last: int):
        indexes = scheduled.index[mask].tolist()
        if not indexes:
            return
        automatic = use_automatic_windows
        window = available_window(timing, first, last, automatic)
        outside = [day for day in calendar if day not in set(window)] or calendar
        amounts = pd.to_numeric(scheduled.loc[indexes, amount_column], errors="coerce").fillna(0.0)
        ordered = amounts.sort_values(ascending=False).index.tolist()
        target_amount = float(amounts.sum()) * max(0.0, min(100.0, float(concentration))) / 100.0
        concentrated, remainder, accumulated = [], [], 0.0
        for row_index in ordered:
            if accumulated + 0.005 < target_amount:
                concentrated.append(row_index)
                accumulated += float(amounts.loc[row_index])
            else:
                remainder.append(row_index)
        for order, row_index in enumerate(concentrated):
            scheduled.loc[row_index, "FECHA"] = choose_day(window, order, len(concentrated))
        for order, row_index in enumerate(remainder):
            scheduled.loc[row_index, "FECHA"] = choose_day(outside, order, len(remainder))

    income_mask = pd.to_numeric(scheduled["DEPOSITO"], errors="coerce").fillna(0.0) > 0.005
    expense_mask = pd.to_numeric(scheduled["RETIRO"], errors="coerce").fillna(0.0) > 0.005
    schedule_flow(
        income_mask, "DEPOSITO", income_timing, config.income_concentration_percent,
        config.income_window_start, config.income_window_end,
    )
    schedule_flow(
        expense_mask, "RETIRO", expense_timing, config.expense_concentration_percent,
        config.expense_window_start, config.expense_window_end,
    )
    return scheduled


def _estimated_average_range(original, generated, opening, config, start, end, seed, document_name):
    """Extremos seguros aproximados conservando cantidades y montos."""
    estimates = []
    for offset, relation in enumerate((AVERAGE_RELATION_BELOW, AVERAGE_RELATION_ABOVE), start=1):
        try:
            trial_generated = _schedule_generated_dates(
                generated, config, start, end, seed + offset * 37, opening, forced_relation=relation
            )
            trial = _mix_and_recalculate(original, trial_generated, opening, seed + offset * 101, document_name, start)
            if float(trial["SALDO"].min()) + 0.005 >= float(config.minimum_balance):
                estimates.append(daily_average_balance(trial, start, end))
        except ValueError:
            continue
    if not estimates:
        return pd.NA, pd.NA
    return round(min(estimates), 2), round(max(estimates), 2)


def _objective_expense(original, skeleton, opening, config, start, end, seed, document_name):
    original_deposits = float(pd.to_numeric(original["DEPOSITO"], errors="coerce").fillna(0.0).sum())
    original_withdrawals = float(pd.to_numeric(original["RETIRO"], errors="coerce").fillna(0.0).sum())
    if config.target_mode == TARGET_EXPENSES:
        return round(config.target_value, 2)
    if config.target_mode == TARGET_FINAL:
        return round(opening + original_deposits + config.total_income - original_withdrawals - config.target_value, 2)
    incomes_only = skeleton[pd.to_numeric(skeleton["DEPOSITO"], errors="coerce").fillna(0.0) > 0.005].copy()
    baseline = _mix_and_recalculate(original, incomes_only, opening, seed, document_name, start)
    baseline_average = daily_average_balance(baseline, start, end)
    expense_rows = skeleton[pd.to_numeric(skeleton["RETIRO"], errors="coerce").fillna(0.0) > 0.005]
    if expense_rows.empty:
        if abs(baseline_average - config.target_value) > 0.01:
            raise ValueError(f"Con cero egresos, el saldo promedio posible es ${baseline_average:,.2f}.")
        return 0.0
    if config.target_value > baseline_average + 0.01:
        maximum_viable_average = baseline_average
        if config.expense_count > 0:
            try:
                minimum_expenses = _rescale_expenses(skeleton, float(config.expense_count))
                minimum_expense_period = _mix_and_recalculate(
                    original, minimum_expenses, opening, seed, document_name, start
                )
                maximum_viable_average = daily_average_balance(minimum_expense_period, start, end)
            except ValueError:
                maximum_viable_average = baseline_average
        income_rows = skeleton[pd.to_numeric(skeleton["DEPOSITO"], errors="coerce").fillna(0.0) > 0.005]
        required_monthly_income = None
        if not income_rows.empty:
            days = (end - start).days + 1
            shares = pd.to_numeric(income_rows["DEPOSITO"], errors="coerce")
            shares = shares / shares.sum()
            persistence = income_rows["FECHA"].apply(
                lambda value: ((end - pd.Timestamp(value).normalize()).days + 1) / days
            )
            income_coefficient = float((shares * persistence).sum())
            if income_coefficient > 0:
                current_monthly_income = float(config.target_total_monthly_income or (original_deposits + config.total_income))
                required_monthly_income = round(
                    current_monthly_income + (config.target_value - maximum_viable_average) / income_coefficient, 2
                )
        options = {"maximum_viable_average": round(maximum_viable_average, 2)}
        if required_monthly_income is not None:
            options["required_monthly_income"] = required_monthly_income
            if config.planned_average_income and config.target_total_monthly_income:
                period_factor = float(config.target_total_monthly_income) / float(config.planned_average_income)
                if period_factor > 0:
                    options["required_plan_average"] = round(required_monthly_income / period_factor, 2)
        raise LinkedFeasibilityError(
            f"Con los ingresos seleccionados, el promedio maximo viable es aproximadamente "
            f"${maximum_viable_average:,.2f}; el objetivo solicitado de ${config.target_value:,.2f} es mayor.",
            options,
        )
    days = (end - start).days + 1
    shares = pd.to_numeric(expense_rows["RETIRO"], errors="coerce")
    shares = shares / shares.sum()
    persistence = expense_rows["FECHA"].apply(lambda value: ((end - pd.Timestamp(value).normalize()).days + 1) / days)
    coefficient = float((shares * persistence).sum())
    return round((baseline_average - config.target_value) / coefficient, 2)


def _adjust_average(original, generated, opening, seed, document_name, start, end, target, tolerance=0.05):
    linked = _mix_and_recalculate(original, generated, opening, seed, document_name, start)
    for _ in range(5):
        achieved = daily_average_balance(linked, start, end)
        difference = achieved - target
        if abs(difference) <= max(0.02, float(tolerance)):
            return linked, achieved, generated
        expense_mask = pd.to_numeric(generated["RETIRO"], errors="coerce").fillna(0.0) > 0.005
        if not expense_mask.any():
            break
        first_index = generated.loc[expense_mask, "FECHA"].astype("datetime64[ns]").idxmin()
        expense_date = pd.Timestamp(generated.loc[first_index, "FECHA"]).normalize()
        weight = ((end - expense_date).days + 1) / ((end - start).days + 1)
        revised = round(float(generated.loc[first_index, "RETIRO"]) + difference / weight, 2)
        if revised < 1.0:
            break
        generated = generated.copy()
        generated.loc[first_index, "RETIRO"] = revised
        linked = _mix_and_recalculate(original, generated, opening, seed, document_name, start)
    achieved = daily_average_balance(linked, start, end)
    if abs(achieved - target) > max(0.05, float(tolerance)):
        raise ValueError(f"No fue posible alcanzar el saldo promedio solicitado. Resultado mas cercano: ${achieved:,.2f}.")
    return linked, achieved, generated


def _solve_average_suggest_final(original, opening, config, start, end, seed_skeleton, seed_mix, document_name, bank):
    """Sugiere ingresos, egresos y saldo final minimizando montos agregados."""
    if config.income_count <= 0 or config.expense_count <= 0:
        raise ValueError("El calculo automatico requiere al menos un ingreso y un egreso nuevos.")
    def evaluate(total_income, total_expense):
        total_income, total_expense = round(float(total_income), 2), round(float(total_expense), 2)
        if total_income < config.income_count or total_expense < config.expense_count:
            return None
        trial_config = replace(config, total_income=total_income, auto_calculate_amounts=False)
        skeleton = _generated_skeleton(start.date(), end.date(), trial_config, seed_skeleton, document_name, bank)
        skeleton = _schedule_generated_dates(skeleton, config, start, end, seed_skeleton, opening)
        generated = _rescale_expenses(skeleton, total_expense)
        linked = _mix_and_recalculate(original, generated, opening, seed_mix, document_name, start)
        average = daily_average_balance(linked, start, end)
        minimum = float(linked["SALDO"].min())
        if minimum < config.minimum_balance - 0.005:
            return None
        return average - config.target_value, linked, generated, average, total_income, total_expense

    scale = max(abs(opening), abs(config.target_value), float(original["DEPOSITO"].sum()), float(original["RETIRO"].sum()), 10000.0)

    def search(variable):
        # Estos montos base deben depender de la cantidad de movimientos, no
        # del promedio solicitado. Antes se usaba 5% del objetivo: al aceptar
        # una sugerencia mayor aumentaba tambien el piso de la siguiente
        # busqueda y la recomendacion crecia indefinidamente.
        fixed_income = max(float(config.income_count) * 100.0, 100.0)
        fixed_expense = max(float(config.expense_count) * 100.0, 100.0)
        low = fixed_income if variable == "income" else fixed_expense
        upper = max(low + 1000.0, scale * 2.0)
        best, bracket = None, None
        for _ in range(6):
            valid = []
            for step in range(61):
                amount = low + (upper - low) * step / 60
                try:
                    result = evaluate(amount, fixed_expense) if variable == "income" else evaluate(fixed_income, amount)
                except ValueError:
                    result = None
                if result is None:
                    continue
                valid.append(result)
                if best is None or abs(result[0]) < abs(best[0]):
                    best = result
            for left, right in zip(valid, valid[1:]):
                if left[0] == 0 or right[0] == 0 or left[0] * right[0] < 0:
                    bracket = (left, right)
                    break
            if bracket or (best is not None and abs(best[0]) <= _objective_tolerance(config)):
                break
            upper *= 5.0
        if bracket:
            left, right = bracket
            variable_index = 4 if variable == "income" else 5
            for _ in range(45):
                amount = (left[variable_index] + right[variable_index]) / 2
                try:
                    middle = evaluate(amount, fixed_expense) if variable == "income" else evaluate(fixed_income, amount)
                except ValueError:
                    break
                if middle is None:
                    break
                if best is None or abs(middle[0]) < abs(best[0]):
                    best = middle
                if abs(middle[0]) <= 0.02:
                    break
                if left[0] * middle[0] <= 0:
                    right = middle
                else:
                    left = middle
        exact = best if best is not None and abs(best[0]) <= _objective_tolerance(config) else None
        return exact, best

    search_results = (search("income"), search("expense"))
    exact_candidates = [exact for exact, _ in search_results if exact is not None]
    approximate_candidates = [closest for _, closest in search_results if closest is not None]
    candidates = exact_candidates or approximate_candidates
    if not candidates:
        raise ValueError(
            "No existe siquiera una propuesta segura con esos datos sin producir saldos negativos. "
            "Aumenta el saldo inicial, reduce el saldo minimo avanzado o modifica las cantidades."
        )
    if exact_candidates:
        best = min(candidates, key=lambda item: item[4] + item[5])
    else:
        # Conservar la alternativa segura mas cercana permite editarla manualmente.
        best = min(candidates, key=lambda item: (abs(item[0]), item[4] + item[5]))
    return best[1], best[3], best[2], best[4], best[5]


def _solve_average_approximate_income(
    original, opening, config, start, end, seed_skeleton, seed_mix, document_name, bank
):
    """Busca una solución cerrada cerca del ingreso mensual indicado.

    La referencia es flexible: se exploran calendarios e ingresos alternativos y
    solo se devuelve una propuesta cuyos saldos ya fueron recalculados. Esto
    evita sugerir un ingreso local que vuelva a resultar insuficiente al
    aplicarlo en una segunda ejecución.
    """
    if config.income_count <= 0 or config.expense_count <= 0:
        raise ValueError("El cálculo guiado requiere al menos un ingreso y un egreso nuevos.")
    original_income = float(pd.to_numeric(original["DEPOSITO"], errors="coerce").fillna(0.0).sum())
    original_expense = float(pd.to_numeric(original["RETIRO"], errors="coerce").fillna(0.0).sum())
    requested_monthly = max(float(config.target_total_monthly_income or 0.0), original_income + config.income_count)
    tolerance = _objective_tolerance(config)
    relations = list(dict.fromkeys((
        _resolved_average_relation(config, opening),
        AVERAGE_RELATION_ABOVE,
        AVERAGE_RELATION_BELOW,
        AVERAGE_RELATION_NEAR,
    )))
    factors = (1.0, 0.95, 1.05, 0.85, 1.15, 0.70, 1.30, 0.50, 1.50, 2.0, 2.5, 3.0, 4.0, 5.0)
    queue = [round(max(original_income + config.income_count, requested_monthly * factor), 2) for factor in factors]
    seen, candidates = set(), []
    maximum_monthly = max(requested_monthly * 20.0, original_income + abs(opening) + abs(config.target_value) * 20.0, 100000.0)

    def add_candidate(linked, generated, average, synthetic_income, synthetic_expense):
        minimum = float(pd.to_numeric(linked["SALDO"], errors="coerce").min())
        if minimum < float(config.minimum_balance) - 0.005:
            return
        monthly_total = original_income + float(synthetic_income)
        candidates.append((
            abs(float(average) - float(config.target_value)),
            abs(monthly_total - requested_monthly),
            float(synthetic_income) + float(synthetic_expense),
            linked, generated, float(average), float(synthetic_income), float(synthetic_expense),
        ))

    attempts = 0
    while queue and attempts < 80:
        monthly_total = round(float(queue.pop(0)), 2)
        monthly_key = int(round(monthly_total * 100))
        if monthly_key in seen or monthly_total > maximum_monthly:
            continue
        seen.add(monthly_key)
        attempts += 1
        synthetic_income = round(monthly_total - original_income, 2)
        if synthetic_income < float(config.income_count) - 0.005:
            continue
        trial_config = replace(
            config,
            total_income=synthetic_income,
            auto_calculate_amounts=False,
            approximate_monthly_income=False,
        )
        try:
            base = _generated_skeleton(
                start.date(), end.date(), trial_config, seed_skeleton, document_name, bank
            )
        except ValueError:
            continue
        for relation_index, relation in enumerate(relations):
            try:
                scheduled = _schedule_generated_dates(
                    base, trial_config, start, end,
                    seed_skeleton + relation_index * 211, opening,
                    forced_relation=relation,
                )
                required_expense = _objective_expense(
                    original, scheduled, opening, trial_config, start, end,
                    seed_mix + relation_index * 223, document_name,
                )
            except LinkedFeasibilityError as exc:
                required_income = (exc.options or {}).get("required_monthly_income")
                if required_income is not None and float(required_income) <= maximum_monthly:
                    queue.append(round(float(required_income) * 1.002, 2))
                continue
            except ValueError:
                continue
            maximum_expense = round(
                float(opening) + original_income + synthetic_income - original_expense - float(config.minimum_balance), 2
            )
            if maximum_expense < float(config.expense_count) - 0.005:
                continue
            if required_expense > maximum_expense + 0.005:
                shortage = required_expense - maximum_expense
                next_income = monthly_total + shortage
                if next_income <= maximum_monthly:
                    queue.append(round(next_income * 1.002, 2))
            expense_options = [min(required_expense, maximum_expense)]
            if required_expense > maximum_expense + 0.005:
                expense_options.append(maximum_expense)
            for expense_total in dict.fromkeys(round(float(value), 2) for value in expense_options):
                if expense_total < float(config.expense_count) - 0.005:
                    continue
                try:
                    generated = _rescale_expenses(scheduled, expense_total)
                    linked = _mix_and_recalculate(
                        original, generated, opening,
                        seed_mix + relation_index * 227, document_name, start,
                    )
                    average = daily_average_balance(linked, start, end)
                    if abs(average - config.target_value) > tolerance:
                        try:
                            linked, average, generated = _adjust_average(
                                original, generated, opening,
                                seed_mix + relation_index * 227, document_name,
                                start, end, config.target_value, tolerance,
                            )
                            expense_total = float(pd.to_numeric(generated["RETIRO"], errors="coerce").fillna(0.0).sum())
                        except ValueError:
                            pass
                    add_candidate(linked, generated, average, synthetic_income, expense_total)
                except ValueError:
                    continue
        exact = [candidate for candidate in candidates if candidate[0] <= tolerance]
        if exact:
            best = min(exact, key=lambda item: (item[1], item[2]))
            return best[3], best[5], best[4], best[6], best[7]

    if candidates:
        best = min(candidates, key=lambda item: (item[0], item[1], item[2]))
        return best[3], best[5], best[4], best[6], best[7]
    raise LinkedFeasibilityError(
        "No se encontró una combinación segura ni siquiera variando el ingreso mensual aproximado y el calendario. "
        "Aumenta el saldo inicial, reduce el saldo mínimo o modifica las cantidades de movimientos.",
        {"requested_monthly_income": round(requested_monthly, 2)},
    )


def _solve_average_auto(original, opening, config, start, end, seed_skeleton, seed_mix, document_name, bank):
    """Calcula simultaneamente ingresos y egresos para promedio y saldo final."""
    if config.income_count <= 0 or config.expense_count <= 0:
        raise ValueError("El calculo automatico del saldo promedio mensual requiere al menos un ingreso y un egreso nuevos.")
    if config.average_target_final is None:
        return _solve_average_suggest_final(
            original, opening, config, start, end, seed_skeleton, seed_mix, document_name, bank
        )
    target_final = float(config.average_target_final)
    if target_final < config.minimum_balance - 0.005:
        raise ValueError("El saldo final deseado no puede ser menor que el saldo minimo permitido.")
    original_income = float(pd.to_numeric(original["DEPOSITO"], errors="coerce").fillna(0.0).sum())
    original_expense = float(pd.to_numeric(original["RETIRO"], errors="coerce").fillna(0.0).sum())
    # I_ficticio - E_ficticio queda fijado por el saldo final solicitado.
    income_minus_expense = round(target_final - opening - original_income + original_expense, 2)
    lower = max(float(config.income_count), income_minus_expense + float(config.expense_count), 0.0)
    lower = round(lower, 2)
    def evaluate(total_income):
        total_income = round(float(total_income), 2)
        total_expense = round(total_income - income_minus_expense, 2)
        if total_income < config.income_count or total_expense < config.expense_count:
            return None
        trial_config = replace(
            config,
            total_income=total_income,
            auto_calculate_amounts=False,
            require_income_at_least_expense=False,
        )
        skeleton = _generated_skeleton(start.date(), end.date(), trial_config, seed_skeleton, document_name, bank)
        skeleton = _schedule_generated_dates(skeleton, config, start, end, seed_skeleton, opening)
        generated = _rescale_expenses(skeleton, total_expense)
        linked = _mix_and_recalculate(original, generated, opening, seed_mix, document_name, start)
        average = daily_average_balance(linked, start, end)
        return average - config.target_value, linked, generated, average, total_income, total_expense

    scale = max(abs(opening), abs(config.target_value), abs(target_final), original_income, original_expense, 10000.0)
    best = None
    bracket = None
    upper = max(lower + 1000.0, scale * 2.0)
    for _ in range(6):
        valid = []
        for step in range(61):
            candidate = lower + (upper - lower) * step / 60
            try:
                result = evaluate(candidate)
            except ValueError:
                result = None
            if result is None:
                continue
            valid.append(result)
            if best is None or abs(result[0]) < abs(best[0]):
                best = result
        for left, right in zip(valid, valid[1:]):
            if left[0] == 0 or right[0] == 0 or left[0] * right[0] < 0:
                bracket = (left, right)
                break
        if bracket or (best is not None and abs(best[0]) <= _objective_tolerance(config)):
            break
        upper *= 5.0
    if bracket:
        left, right = bracket
        for _ in range(45):
            candidate_income = (left[4] + right[4]) / 2
            try:
                middle = evaluate(candidate_income)
            except ValueError:
                middle = None
            if middle is None:
                break
            if abs(middle[0]) < abs(best[0]):
                best = middle
            if abs(middle[0]) <= 0.02:
                best = middle
                break
            if left[0] * middle[0] <= 0:
                right = middle
            else:
                left = middle
    if best is None or abs(best[0]) > _objective_tolerance(config):
        closest = f" El resultado mas cercano fue ${best[3]:,.2f}." if best is not None else ""
        raise ValueError(
            "No existe una combinacion viable con esas cantidades de movimientos, saldo promedio mensual y saldo final. "
            "Cambia el saldo final, el promedio o el numero de movimientos." + closest
        )
    if float(best[1]["SALDO"].min()) < float(config.minimum_balance) - 0.005:
        raise ValueError(
            f"La propuesta alcanza el promedio y saldo final, pero su saldo minimo seria ${float(best[1]['SALDO'].min()):,.2f}, "
            f"menor al minimo solicitado de ${config.minimum_balance:,.2f}."
        )
    return best[1], best[3], best[2], best[4], best[5]


def _solve_average_fixed_income_final(
    original, opening, config, start, end, seed_skeleton, seed_mix, document_name, bank
):
    """Con ingreso y saldo final fijos, busca fechas que alcancen el promedio."""
    if config.average_target_final is None:
        raise ValueError("Falta el saldo final para el cálculo combinado.")
    original_income = float(pd.to_numeric(original["DEPOSITO"], errors="coerce").fillna(0.0).sum())
    original_expense = float(pd.to_numeric(original["RETIRO"], errors="coerce").fillna(0.0).sum())
    total_income = round(float(config.total_income), 2)
    target_final = round(float(config.average_target_final), 2)
    total_expense = round(opening + original_income + total_income - original_expense - target_final, 2)
    if total_expense < float(config.expense_count) - 0.005:
        raise LinkedFeasibilityError(
            f"Con el ingreso mensual fijado, el saldo final de ${target_final:,.2f} requeriría "
            f"${total_expense:,.2f} de egresos nuevos, menor que lo necesario para crear "
            f"{config.expense_count} movimientos.",
            {"maximum_viable_final": round(opening + original_income + total_income - original_expense, 2)},
        )
    base = _generated_skeleton(start.date(), end.date(), config, seed_skeleton, document_name, bank)
    base = _rescale_expenses(base, total_expense)
    tolerance = _objective_tolerance(config)
    candidates = []
    concentration_values = sorted({
        0.0, 20.0, 40.0, 60.0, 80.0, 90.0, 100.0,
        float(config.income_concentration_percent), float(config.expense_concentration_percent),
    })
    attempt = 0
    found_within_tolerance = False
    for income_concentration in concentration_values:
        for expense_concentration in concentration_values:
            attempt += 1
            trial_config = replace(
                config,
                income_concentration_percent=income_concentration,
                expense_concentration_percent=expense_concentration,
            )
            for variation in range(3):
                try:
                    generated = _schedule_generated_dates(
                        base, trial_config, start, end, seed_skeleton + attempt * 101 + variation * 17, opening
                    )
                    linked = _mix_and_recalculate(
                        original, generated, opening, seed_mix + attempt * 103 + variation * 19,
                        document_name, start,
                    )
                    minimum = float(linked["SALDO"].min())
                    if minimum < float(config.minimum_balance) - 0.005:
                        continue
                    average = daily_average_balance(linked, start, end)
                    candidates.append((abs(average - config.target_value), linked, generated, average))
                    if abs(average - config.target_value) <= tolerance:
                        found_within_tolerance = True
                        break
                except ValueError:
                    continue
            if found_within_tolerance:
                break
        if found_within_tolerance:
            break
    if not candidates:
        raise LinkedFeasibilityError(
            "No existe una distribución segura para combinar ese ingreso, saldo final y saldo mínimo.",
            {"required_opening_balance": round(float(opening + config.minimum_balance), 2)},
        )
    best = min(candidates, key=lambda item: item[0])
    minimum_average = min(item[3] for item in candidates)
    maximum_average = max(item[3] for item in candidates)
    if best[0] > tolerance:
        raise LinkedFeasibilityError(
            f"Con ingreso mensual de ${(original_income + total_income):,.2f} y saldo final de "
            f"${target_final:,.2f}, no se alcanzó el promedio solicitado de ${config.target_value:,.2f} "
            f"dentro de la tolerancia configurada. Las distribuciones seguras exploradas quedaron entre "
            f"${minimum_average:,.2f} y ${maximum_average:,.2f}. "
            f"La propuesta más cercana es ${best[3]:,.2f}.",
            {
                "minimum_viable_average": round(minimum_average, 2),
                "maximum_viable_average": round(maximum_average, 2),
                "closest_average": round(best[3], 2),
            },
        )
    return best[1], best[3], best[2], total_income, total_expense


def link_statement_periods(base_periods, configs, first_opening_balance, seed=2026):
    """Enlaza de 2 a N documentos, incluso de bancos distintos."""
    if len(base_periods) != len(configs):
        raise ValueError("Debe existir una configuracion por cada documento.")
    if not base_periods:
        raise ValueError("No se recibieron documentos.")
    normalized = []
    for item in base_periods:
        if len(item) == 2:
            name, base = item
            bank = "GENERICO"
        else:
            name, bank, base = item
        original, source_opening, start, end = _normalize_base(base, name)
        normalized.append((start, name, str(bank).upper(), original, source_opening, end))
    normalized.sort(key=lambda item: item[0])
    results, summaries = [], []
    counterparties_used_across_periods = set()
    transfer_banks_used_across_periods = set()
    opening = round(float(first_opening_balance), 2)
    for period_index, ((start, name, bank, original, source_opening, end), config) in enumerate(zip(normalized, configs), start=1):
        names_already_used = _counterparties_in(original) | counterparties_used_across_periods
        banks_already_used = _transfer_banks_in(original) | transfer_banks_used_across_periods
        config = replace(
            config,
            excluded_counterparty_names=tuple(sorted(names_already_used)),
            previously_used_transfer_banks=tuple(sorted(banks_already_used)),
        )
        original_income_amount = round(float(pd.to_numeric(original["DEPOSITO"], errors="coerce").fillna(0.0).sum()), 2)
        if config.target_total_monthly_income is not None:
            target_monthly_income = round(float(config.target_total_monthly_income), 2)
            if target_monthly_income + 0.005 < original_income_amount:
                raise ValueError(
                    f"En el periodo {period_index}, los movimientos originales ya contienen ${original_income_amount:,.2f} "
                    f"de ingresos, mayor que el total mensual solicitado de ${target_monthly_income:,.2f}."
                )
            config = replace(config, total_income=round(target_monthly_income - original_income_amount, 2))
        counts_suggested = bool(
            config.target_mode == TARGET_AVERAGE
            and config.auto_calculate_counts
        )
        if counts_suggested:
            original_income_count = int((pd.to_numeric(original["DEPOSITO"], errors="coerce").fillna(0.0) > 0.005).sum())
            original_expense_count = int((pd.to_numeric(original["RETIRO"], errors="coerce").fillna(0.0) > 0.005).sum())
            original_count = original_income_count + original_expense_count
            # Agregar cerca de 30% conserva una densidad didactica realista sin
            # dominar los movimientos originales. Se respeta la proporcion del PDF.
            suggested_total = max(4, min(16, round(max(original_count, 4) * 0.30)))
            income_ratio = original_income_count / original_count if original_count else 0.45
            suggested_income_count = max(1, min(suggested_total - 1, round(suggested_total * income_ratio)))
            suggested_expense_count = max(1, suggested_total - suggested_income_count)
            config = replace(
                config,
                income_count=suggested_income_count,
                expense_count=suggested_expense_count,
            )
        if (
            config.target_mode == TARGET_AVERAGE
            and config.average_target_final is not None
            and config.average_relation != AVERAGE_RELATION_AUTO
        ):
            target_relation = _relation_from_values(
                config.target_value, config.average_target_final, _objective_tolerance(config)
            )
            if not _relation_matches(config.average_relation, target_relation):
                raise LinkedFeasibilityError(
                    f"El promedio solicitado de ${config.target_value:,.2f} y el saldo final de "
                    f"${config.average_target_final:,.2f} producen la relación '{target_relation}', no "
                    f"'{config.average_relation}'. Cambia uno de los dos objetivos o selecciona relación automática.",
                    {
                        "period": period_index,
                        "average_target": round(float(config.target_value), 2),
                        "final_target": round(float(config.average_target_final), 2),
                    },
                )
        fixed_income_final_average = bool(
            config.target_mode == TARGET_AVERAGE
            and config.target_total_monthly_income is not None
            and config.average_target_final is not None
        )
        approximate_income_average = bool(
            config.target_mode == TARGET_AVERAGE
            and config.target_total_monthly_income is not None
            and config.approximate_monthly_income
            and config.average_target_final is None
        )
        auto_average = config.target_mode == TARGET_AVERAGE and config.auto_calculate_amounts
        manual_average = config.target_mode == TARGET_AVERAGE and config.manual_total_expense is not None
        if fixed_income_final_average:
            linked, average, generated, calculated_income, expense_total = _solve_average_fixed_income_final(
                original, opening, config, start, end,
                seed + period_index * 101, seed + period_index * 1009, name, bank,
            )
            skeleton = generated
        elif approximate_income_average:
            linked, average, generated, calculated_income, expense_total = _solve_average_approximate_income(
                original, opening, config, start, end,
                seed + period_index * 101, seed + period_index * 1009, name, bank,
            )
            skeleton = generated
        elif auto_average:
            linked, average, generated, calculated_income, expense_total = _solve_average_auto(
                original, opening, config, start, end,
                seed + period_index * 101, seed + period_index * 1009, name, bank,
            )
            skeleton = generated
        elif manual_average:
            calculated_income = float(config.total_income)
            expense_total = float(config.manual_total_expense)
            skeleton = _generated_skeleton(start.date(), end.date(), config, seed + period_index * 101, name, bank)
            skeleton = _schedule_generated_dates(
                skeleton, config, start, end, seed + period_index * 101, opening
            )
            generated = _rescale_expenses(skeleton, expense_total)
            linked = _mix_and_recalculate(original, generated, opening, seed + period_index * 1009, name, start)
            average = daily_average_balance(linked, start, end)
        else:
            calculated_income = float(config.total_income)
            skeleton = _generated_skeleton(start.date(), end.date(), config, seed + period_index * 101, name, bank)
            if config.target_mode == TARGET_AVERAGE:
                skeleton = _schedule_generated_dates(
                    skeleton, config, start, end, seed + period_index * 101, opening
                )
            try:
                expense_total = _objective_expense(
                    original, skeleton, opening, config, start, end, seed + period_index * 1009, name
                )
            except LinkedFeasibilityError as exc:
                exc.options.setdefault("period", period_index)
                raise
        if expense_total < -0.005:
            if config.target_mode == TARGET_FINAL and config.target_total_monthly_income is not None:
                additional_income = round(abs(float(expense_total)), 2)
                maximum_viable_final = round(float(config.target_value) + float(expense_total), 2)
                required_monthly_income = round(float(config.target_total_monthly_income) + additional_income, 2)
                options = {
                    "period": period_index,
                    "maximum_viable_final": maximum_viable_final,
                    "required_monthly_income": required_monthly_income,
                }
                if config.planned_average_income and config.target_total_monthly_income:
                    period_factor = float(config.target_total_monthly_income) / float(config.planned_average_income)
                    if period_factor > 0:
                        options["required_plan_average"] = round(required_monthly_income / period_factor, 2)
                if config.planned_average_final and config.target_value:
                    final_factor = float(config.target_value) / float(config.planned_average_final)
                    if final_factor > 0:
                        options["maximum_viable_final_plan_average"] = round(maximum_viable_final / final_factor, 2)
                raise LinkedFeasibilityError(
                    f"El saldo final de ${config.target_value:,.2f} supera los fondos posibles con el ingreso "
                    f"asignado. El saldo final maximo sin inventar egresos negativos es ${maximum_viable_final:,.2f}; "
                    f"para conservar el objetivo se requiere un ingreso mensual cercano a ${required_monthly_income:,.2f}.",
                    options,
                )
            raise ValueError(f"En el periodo {period_index}, el objetivo exige egresos negativos (${expense_total:,.2f}). Aumenta el saldo final/promedio o reduce ingresos.")
        original_income = round(float(pd.to_numeric(original["DEPOSITO"], errors="coerce").fillna(0.0).sum()), 2)
        original_expense = round(float(pd.to_numeric(original["RETIRO"], errors="coerce").fillna(0.0).sum()), 2)
        total_income_planned = round(original_income + float(calculated_income), 2)
        total_expense_planned = round(original_expense + max(0.0, expense_total), 2)
        if config.target_mode == TARGET_AVERAGE and config.target_total_monthly_income is not None:
            max_synthetic_expense = round(
                opening + total_income_planned - original_expense - float(config.minimum_balance), 2
            )
            if expense_total > max_synthetic_expense + 0.005:
                additional_funds = round(expense_total - max_synthetic_expense, 2)
                minimum_average_hint = ""
                if max_synthetic_expense >= config.expense_count:
                    try:
                        capped_generated = _rescale_expenses(skeleton, max_synthetic_expense)
                        capped_period = _mix_and_recalculate(
                            original, capped_generated, opening, seed + period_index * 1009, name, start
                        )
                        minimum_viable_average = daily_average_balance(capped_period, start, end)
                        minimum_average_hint = (
                            f" Con ingresos mensuales totales de ${total_income_planned:,.2f}, el promedio minimo "
                            f"viable con esta distribucion es aproximadamente ${minimum_viable_average:,.2f}."
                        )
                    except ValueError:
                        minimum_average_hint = ""
                required_monthly_income = round(total_income_planned + additional_funds, 2)
                options = {
                    "period": period_index,
                    "minimum_viable_average": round(float(minimum_viable_average), 2)
                    if minimum_average_hint else None,
                    "required_monthly_income": required_monthly_income,
                }
                if config.planned_average_income and config.target_total_monthly_income:
                    period_factor = float(config.target_total_monthly_income) / float(config.planned_average_income)
                    if period_factor > 0:
                        options["required_plan_average"] = round(required_monthly_income / period_factor, 2)
                raise LinkedFeasibilityError(
                    f"El promedio de ${config.target_value:,.2f} requiere aproximadamente ${expense_total:,.2f} "
                    f"de egresos nuevos, pero solo pueden cubrirse ${max_synthetic_expense:,.2f} sin sobregiro. "
                    f"Faltan ${additional_funds:,.2f}; el ingreso mensual total tendria que subir a cerca de "
                    f"${required_monthly_income:,.2f}, o debe elevarse el promedio objetivo.{minimum_average_hint}",
                    options,
                )
        if config.require_income_at_least_expense and total_expense_planned > total_income_planned + 0.005:
            additional_income = round(total_expense_planned - total_income_planned, 2)
            feasibility_hint = ""
            max_synthetic_expense = round(total_income_planned - original_expense, 2)
            if config.target_mode == TARGET_AVERAGE and max_synthetic_expense >= config.expense_count:
                try:
                    capped_generated = _rescale_expenses(skeleton, max_synthetic_expense)
                    capped_period = _mix_and_recalculate(
                        original, capped_generated, opening, seed + period_index * 1009, name, start
                    )
                    strict_average = daily_average_balance(capped_period, start, end)
                    feasibility_hint = (
                        f" Con las fechas y proporciones actuales, el saldo promedio compatible con esa regla "
                        f"seria aproximadamente ${strict_average:,.2f}."
                    )
                except ValueError:
                    feasibility_hint = ""
            raise ValueError(
                f"En el periodo {period_index}, el objetivo requiere egresos totales de ${total_expense_planned:,.2f}, "
                f"pero los ingresos totales son ${total_income_planned:,.2f}. Para cumplir la regla ingresos >= egresos, "
                f"existe un faltante neto actual de ${additional_income:,.2f}. Reduce egresos, eleva el saldo promedio/final "
                f"objetivo o permite usar parte del saldo inicial.{feasibility_hint}"
            )
        if not fixed_income_final_average and not approximate_income_average and not auto_average and not manual_average:
            generated = _rescale_expenses(skeleton, max(0.0, expense_total))
            if config.target_mode == TARGET_AVERAGE:
                linked, average, generated = _adjust_average(
                    original, generated, opening, seed + period_index * 1009, name, start, end,
                    config.target_value, _objective_tolerance(config),
                )
            else:
                linked = _mix_and_recalculate(original, generated, opening, seed + period_index * 1009, name, start)
                average = daily_average_balance(linked, start, end)
        linked = _rebalance_final_transfer_banks(
            linked, config, seed + period_index * 4099, bank
        )
        checks = movement_checks(linked)
        original_only = linked[linked["ORIGEN"] == "ORIGINAL"]
        fictitious_only = linked[linked["ORIGEN"] == "FICTICIO"]
        synthetic_income = round(float(fictitious_only["DEPOSITO"].sum()), 2)
        synthetic_expense = round(float(fictitious_only["RETIRO"].sum()), 2)
        synthetic_income_count = int((pd.to_numeric(fictitious_only["DEPOSITO"], errors="coerce").fillna(0.0) > 0.005).sum())
        synthetic_expense_count = int((pd.to_numeric(fictitious_only["RETIRO"], errors="coerce").fillna(0.0) > 0.005).sum())
        new_counterparties = {
            _normalized_name(value) for value in fictitious_only.get("NOMBRE_CONTRAPARTE", pd.Series(dtype=str)).astype(str)
            if _normalized_name(value)
        }
        normalized_counterparty_sequence = [
            _normalized_name(value)
            for value in fictitious_only.get("NOMBRE_CONTRAPARTE", pd.Series(dtype=str)).astype(str)
            if _normalized_name(value)
        ]
        if len(normalized_counterparty_sequence) != len(set(normalized_counterparty_sequence)):
            raise ValueError("La validación detectó nombres ficticios repetidos; no se generó la salida.")
        new_transfer_banks = {
            str(value).upper() for value in fictitious_only.get("BANCO_TRANSFERENCIA", pd.Series(dtype=str)).astype(str)
            if str(value).strip()
        }
        configured_banks = {
            str(value).strip().upper()
            for value in (BANAMEX_TRANSFER_BANKS if config.transfer_banks is None else config.transfer_banks)
            if str(value).strip()
        }
        unauthorized_banks = sorted(new_transfer_banks - configured_banks)
        if unauthorized_banks:
            raise ValueError(
                "Se bloqueó la salida porque contiene bancos ficticios que fueron excluidos: "
                + ", ".join(unauthorized_banks)
            )
        default_income_concepts = ("TRANSFERENCIA RECIBIDA",) if _is_banamex_micuenta_bank(bank) else INCOME_CONCEPTS
        default_expense_concepts = ("TRANSFERENCIA ENVIADA",) if _is_banamex_micuenta_bank(bank) else EXPENSE_CONCEPTS
        configured_income_concepts = {
            str(value).strip().upper()
            for value in (default_income_concepts if config.income_concepts is None else config.income_concepts)
            if str(value).strip()
        }
        configured_expense_concepts = {
            str(value).strip().upper()
            for value in (default_expense_concepts if config.expense_concepts is None else config.expense_concepts)
            if str(value).strip()
        }
        requested_concepts = fictitious_only.get("CONCEPTO_SOLICITADO", pd.Series("", index=fictitious_only.index)).astype(str).str.strip().str.upper()
        income_mask = pd.to_numeric(fictitious_only["DEPOSITO"], errors="coerce").fillna(0) > 0.005
        invalid_income_concepts = sorted(set(requested_concepts[income_mask]) - configured_income_concepts - {""})
        invalid_expense_concepts = sorted(set(requested_concepts[~income_mask]) - configured_expense_concepts - {""})
        if invalid_income_concepts or invalid_expense_concepts:
            invalid = sorted(set(invalid_income_concepts + invalid_expense_concepts))
            raise ValueError(
                "Se bloqueó la salida porque contiene conceptos ficticios que fueron eliminados: "
                + ", ".join(invalid)
            )
        new_bank_sequence = [
            str(value).strip().upper()
            for value in fictitious_only.get("BANCO_TRANSFERENCIA", pd.Series(dtype=str)).astype(str)
            if str(value).strip()
        ]
        consecutive_bank_repetitions = sum(
            left == right for left, right in zip(new_bank_sequence, new_bank_sequence[1:])
        )
        if consecutive_bank_repetitions:
            raise ValueError("La validación detectó bancos repetidos en movimientos ficticios consecutivos.")
        income_bank_distribution = fictitious_only[
            pd.to_numeric(fictitious_only["DEPOSITO"], errors="coerce").fillna(0) > 0.005
        ].get("BANCO_TRANSFERENCIA", pd.Series(dtype=str)).value_counts().to_dict()
        expense_bank_distribution = fictitious_only[
            pd.to_numeric(fictitious_only["RETIRO"], errors="coerce").fillna(0) > 0.005
        ].get("BANCO_TRANSFERENCIA", pd.Series(dtype=str)).value_counts().to_dict()
        expanded_name_catalog = bool(
            fictitious_only.get("CATALOGO_NOMBRES_AMPLIADO", pd.Series(False, index=fictitious_only.index)).fillna(False).any()
        )
        all_authorized_banks_previously_used = bool(
            fictitious_only.get("BANCOS_AUTORIZADOS_YA_USADOS", pd.Series(False, index=fictitious_only.index)).fillna(False).any()
        )
        achieved = synthetic_expense if config.target_mode == TARGET_EXPENSES else (checks["SALDO_FINAL"] if config.target_mode == TARGET_FINAL else average)
        target_tolerance = _objective_tolerance(config) if config.target_mode == TARGET_AVERAGE else 0.05
        achieved_relation = _relation_from_values(average, checks["SALDO_FINAL"], target_tolerance)
        relation_ok = _relation_matches(config.average_relation, achieved_relation)
        minimum_estimated_average, maximum_estimated_average = (pd.NA, pd.NA)
        if config.target_mode == TARGET_AVERAGE:
            minimum_estimated_average, maximum_estimated_average = _estimated_average_range(
                original, generated, opening, config, start, end,
                seed + period_index * 2017, name,
            )
        range_ok = True
        if config.target_mode == TARGET_AVERAGE and not pd.isna(minimum_estimated_average) and not pd.isna(maximum_estimated_average):
            range_ok = bool(
                float(minimum_estimated_average) - target_tolerance
                <= float(config.target_value)
                <= float(maximum_estimated_average) + target_tolerance
            )
        summaries.append({
            "PERIODO": period_index, "BANCO": bank, "DOCUMENTO": name,
            "FECHA_INICIAL": start, "FECHA_FINAL": end,
            "SALDO_INICIAL_FUENTE": round(float(source_opening), 2),
            "SALDO_INICIAL_SIMULADO": round(float(opening), 2),
            "OBJETIVO": config.target_mode, "VALOR_OBJETIVO": round(config.target_value, 2),
            "VALOR_OBTENIDO": round(float(achieved), 2), "DIFERENCIA_OBJETIVO": round(float(achieved) - config.target_value, 2),
            "TOLERANCIA_OBJETIVO": target_tolerance,
            "OBJETIVO_DENTRO_DE_TOLERANCIA": bool(abs(float(achieved) - config.target_value) <= target_tolerance),
            "RELACION_PROMEDIO_FINAL_SOLICITADA": config.average_relation if config.target_mode == TARGET_AVERAGE else "NO APLICA",
            "RELACION_PROMEDIO_FINAL_OBTENIDA": achieved_relation if config.target_mode == TARGET_AVERAGE else "NO APLICA",
            "RELACION_PROMEDIO_FINAL_CUMPLIDA": bool(relation_ok) if config.target_mode == TARGET_AVERAGE else True,
            "DIFERENCIA_PROMEDIO_MENOS_FINAL": round(float(average) - float(checks["SALDO_FINAL"]), 2),
            "PROMEDIO_MINIMO_ESTIMADO": minimum_estimated_average,
            "PROMEDIO_MAXIMO_ESTIMADO": maximum_estimated_average,
            "OBJETIVO_DENTRO_RANGO_ESTIMADO": bool(range_ok),
            "CALENDARIO_INGRESOS": config.income_timing if config.target_mode == TARGET_AVERAGE else "NO APLICA",
            "CALENDARIO_EGRESOS": config.expense_timing if config.target_mode == TARGET_AVERAGE else "NO APLICA",
            "CONCENTRACION_INGRESOS_PCT": round(float(config.income_concentration_percent), 2),
            "CONCENTRACION_EGRESOS_PCT": round(float(config.expense_concentration_percent), 2),
            "MAX_MOVIMIENTOS_NUEVOS_POR_DIA": int(config.max_movements_per_day),
            "CALCULO_AUTOMATICO_MONTOS": bool(auto_average or fixed_income_final_average or approximate_income_average),
            "CALCULO_GUIADO_INGRESOS": bool(config.target_total_monthly_income is not None),
            "INGRESO_MENSUAL_APROXIMADO": bool(approximate_income_average),
            "CANTIDADES_SUGERIDAS_AUTOMATICAMENTE": counts_suggested,
            "CALCULO_MANUAL_MONTOS": bool(manual_average),
            "INGRESO_MENSUAL_OBJETIVO": round(float(config.target_total_monthly_income), 2) if config.target_total_monthly_income is not None else pd.NA,
            "INGRESO_PROMEDIO_DEL_PLAN": round(float(config.planned_average_income), 2) if config.planned_average_income is not None else pd.NA,
            "VARIACION_MAXIMA_INGRESOS_PCT": round(float(config.planned_variation_percent), 2) if config.planned_variation_percent is not None else pd.NA,
            "SALDO_FINAL_PROMEDIO_DEL_PLAN": round(float(config.planned_average_final), 2) if config.planned_average_final is not None else pd.NA,
            "VARIACION_MAXIMA_SALDO_FINAL_PCT": round(float(config.planned_final_variation_percent), 2) if config.planned_final_variation_percent is not None else pd.NA,
            "SALDO_FINAL_OBJETIVO_PROMEDIO": round(
                float(config.average_target_final if config.average_target_final is not None else checks["SALDO_FINAL"]), 2
            ) if (auto_average or fixed_income_final_average) else pd.NA,
            "MOVIMIENTOS_ORIGINALES": int(len(original_only)), "MOVIMIENTOS_FICTICIOS": int(len(fictitious_only)),
            "INGRESOS_FICTICIOS_CANTIDAD": synthetic_income_count,
            "EGRESOS_FICTICIOS_CANTIDAD": synthetic_expense_count,
            "INGRESOS_FICTICIOS": synthetic_income, "EGRESOS_FICTICIOS": synthetic_expense,
            "NOMBRES_BASE_DETECTADOS": len(names_already_used),
            "NOMBRES_NUEVOS_UNICOS": len(new_counterparties),
            "CATALOGO_NOMBRES_AMPLIADO": expanded_name_catalog,
            "BANCOS_BASE_DETECTADOS": len(banks_already_used),
            "BANCOS_NUEVOS_UTILIZADOS": len(new_transfer_banks),
            "BANCOS_PERMITIDOS_CONFIGURADOS": ", ".join(sorted(configured_banks)),
            "CONCEPTOS_INGRESO_CONFIGURADOS": "; ".join(sorted(configured_income_concepts)),
            "CONCEPTOS_EGRESO_CONFIGURADOS": "; ".join(sorted(configured_expense_concepts)),
            "DISTRIBUCION_BANCOS_INGRESOS": ", ".join(f"{bank}: {count}" for bank, count in sorted(income_bank_distribution.items())),
            "DISTRIBUCION_BANCOS_EGRESOS": ", ".join(f"{bank}: {count}" for bank, count in sorted(expense_bank_distribution.items())),
            "REPETICIONES_CONSECUTIVAS_BANCOS_NUEVOS": consecutive_bank_repetitions,
            "BANCOS_AUTORIZADOS_YA_USADOS": all_authorized_banks_previously_used,
            "INGRESOS_ORIGINALES": original_income,
            "INGRESOS_TOTALES": checks["INGRESOS"], "EGRESOS_TOTALES": checks["EGRESOS"],
            "FLUJO_NETO": round(checks["INGRESOS"] - checks["EGRESOS"], 2),
            "FONDOS_DISPONIBLES": round(opening + checks["INGRESOS"], 2),
            "SALDO_INICIAL_CONTADO_COMO_INGRESO": False,
            "EGRESOS_CUBIERTOS": bool(opening + checks["INGRESOS"] + 0.005 >= checks["EGRESOS"]),
            "REGLA_INGRESOS_EGRESOS": "OPCIONAL" if config.require_income_at_least_expense else "NO APLICA",
            "INGRESOS_MAYORES_IGUALES_EGRESOS": bool(checks["INGRESOS"] + 0.005 >= checks["EGRESOS"]),
            "SALDO_FINAL": checks["SALDO_FINAL"], "SALDO_PROMEDIO_MENSUAL": average,
            "SALDO_MINIMO": round(float(linked["SALDO"].min()), 2), "CUADRA": bool(checks["CUADRA"]),
        })
        linked.attrs.update({"period_start": start, "period_end": end, "bank": bank})
        results.append(linked)
        counterparties_used_across_periods.update(new_counterparties)
        transfer_banks_used_across_periods.update(new_transfer_banks)
        opening = checks["SALDO_FINAL"]
    return results, pd.DataFrame(summaries)


def link_scotiabank_periods(base_periods, configs, first_opening_balance, seed=2026):
    """Alias compatible; el motor actual es multibanco."""
    return link_statement_periods(base_periods, configs, first_opening_balance, seed)


def validate_linked_periods(periods, summary):
    chain_differences = []
    for index in range(1, len(periods)):
        chain_differences.append(round(float(periods[index]["SALDO"].iloc[0]) - float(periods[index - 1]["SALDO"].iloc[-1]), 2))
    tolerances = pd.to_numeric(summary.get("TOLERANCIA_OBJETIVO", 0.05), errors="coerce").fillna(0.05)
    differences_ok = pd.to_numeric(summary.get("DIFERENCIA_OBJETIVO", 0), errors="coerce").fillna(0).abs() <= tolerances
    manual_reference = summary.get("CALCULO_MANUAL_MONTOS", pd.Series(False, index=summary.index)).fillna(False).astype(bool)
    objective_ok = bool((differences_ok | manual_reference).all()) if len(summary) else False
    no_negative_balances = bool(periods) and all(float(period["SALDO"].min()) >= -0.005 for period in periods)
    no_negative_finals = bool(periods) and all(float(period["SALDO"].iloc[-1]) >= -0.005 for period in periods)
    flow_rule_ok = bool(summary["INGRESOS_MAYORES_IGUALES_EGRESOS"].all()) if "INGRESOS_MAYORES_IGUALES_EGRESOS" in summary else True
    coverage_ok = bool(summary["EGRESOS_CUBIERTOS"].all()) if "EGRESOS_CUBIERTOS" in summary else no_negative_finals
    relation_ok = bool(summary["RELACION_PROMEDIO_FINAL_CUMPLIDA"].all()) if "RELACION_PROMEDIO_FINAL_CUMPLIDA" in summary else True
    return {
        "PERIODOS": len(periods), "TODOS_CUADRAN": bool(summary["CUADRA"].all()) if len(summary) else False,
        "OBJETIVOS_CUMPLIDOS": objective_ok,
        "SALDOS_NO_NEGATIVOS": no_negative_balances,
        "SALDOS_FINALES_NO_NEGATIVOS": no_negative_finals,
        "INGRESOS_MAYORES_IGUALES_EGRESOS": flow_rule_ok,
        "EGRESOS_CUBIERTOS": coverage_ok,
        "RELACIONES_PROMEDIO_FINAL_CUMPLIDAS": relation_ok,
        "CADENA_CUADRA": all(abs(value) < 0.01 for value in chain_differences),
        "DIFERENCIAS_CADENA": chain_differences,
        "SALDO_MINIMO_GLOBAL": min(float(period["SALDO"].min()) for period in periods),
    }
