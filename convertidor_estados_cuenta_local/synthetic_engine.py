"""Motor de datos ficticios para ejercicios didacticos.

No contiene identidades ni cuentas reales. Los movimientos se generan de forma
determinista cuando se proporciona la misma semilla, lo que permite repetir una
clase y auditar los calculos.
"""

from __future__ import annotations

import random
import re
from dataclasses import dataclass
from datetime import date, datetime, timedelta

import numpy as np
import pandas as pd


FIRST_NAMES = [
    "ALEJANDRA", "ANDREA", "CAROLINA", "CLAUDIA", "DANIEL", "DIEGO",
    "EDUARDO", "FERNANDA", "GABRIELA", "JORGE", "JOSE", "LAURA",
    "LUIS", "MARIANA", "MIGUEL", "PAOLA", "ROBERTO", "SOFIA",
]
LAST_NAMES = [
    "AGUILAR", "CASTILLO", "CHAVEZ", "CONTRERAS", "FLORES", "GARCIA",
    "GUTIERREZ", "HERNANDEZ", "LOPEZ", "MARTINEZ", "MENDOZA", "MORALES",
    "NAVARRO", "RAMIREZ", "REYES", "RODRIGUEZ", "SANCHEZ", "TORRES",
]
INCOME_CONCEPTS = [
    "TRANSFERENCIA SPEI RECIBIDA", "DEPOSITO EN EFECTIVO", "COBRO A CLIENTE",
    "ABONO POR SERVICIOS", "PAGO DE FACTURA", "TRASPASO ENTRE CUENTAS",
]
EXPENSE_CONCEPTS = [
    "TRANSFERENCIA SPEI ENVIADA", "PAGO A PROVEEDOR", "RENTA DEL MES",
    "SERVICIO DE ENERGIA", "SERVICIO DE TELECOMUNICACIONES", "COMPRA DE INSUMOS",
    "PAGO DE IMPUESTOS", "COMISION BANCARIA", "PAGO DE HONORARIOS",
    "MANTENIMIENTO Y SERVICIOS", "COMPRA CON TARJETA", "RETIRO EN EFECTIVO",
]
TRANSFER_BANKS = (
    "BBVA MEXICO", "SANTANDER", "BANORTE", "HSBC", "SCOTIABANK",
    "BANCO AZTECA", "BAJIO", "INBURSA", "AFIRME", "BANCOPPEL",
)


@dataclass(frozen=True)
class GenerationConfig:
    start_date: date
    end_date: date
    income_count: int
    expense_count: int
    total_income: float
    opening_balance: float
    target_mode: str = "SALDO FINAL"
    target_balance: float = 0.0
    seed: int = 2026
    include_weekends: bool = False
    income_concepts: tuple[str, ...] | None = None
    expense_concepts: tuple[str, ...] | None = None
    transfer_banks: tuple[str, ...] | None = None


def _digits(rng: random.Random, count: int) -> str:
    return "".join(str(rng.randrange(10)) for _ in range(count))


def _clabe_check_digit(first_seventeen: str) -> str:
    """Calcula el digito de control CLABE con ponderadores 3, 7 y 1."""
    if len(first_seventeen) != 17 or not first_seventeen.isdigit():
        raise ValueError("La base de una CLABE debe contener exactamente 17 digitos.")
    weights = (3, 7, 1)
    weighted = sum(
        (int(digit) * weights[index % 3]) % 10
        for index, digit in enumerate(first_seventeen)
    )
    return str((10 - weighted % 10) % 10)


def generate_identity(seed: int = 2026, bank_code: str = "044") -> dict:
    """Genera una identidad obviamente sintetica con formatos plausibles."""
    rng = random.Random(seed)
    first = rng.choice(FIRST_NAMES)
    paternal, maternal = rng.sample(LAST_NAMES, 2)
    institution = re.sub(r"\D", "", str(bank_code)).zfill(3)[-3:]
    plaza = _digits(rng, 3)
    account_component = _digits(rng, 11)
    clabe_base = institution + plaza + account_component
    clabe = clabe_base + _clabe_check_digit(clabe_base)
    rfc = f"{paternal[:2]}{maternal[:1]}{first[:1]}900101" + _digits(rng, 3)
    return {
        "NOMBRE": f"{first} {paternal} {maternal}",
        "RFC": rfc,
        "CUENTA": _digits(rng, 10),
        "CLABE": clabe,
        "PLAZA": plaza,
        "CLIENTE": _digits(rng, 8),
        "DOMICILIO": "AV. EJEMPLO 100, COL. CENTRO, C.P. 00000",
    }


def _valid_dates(config: GenerationConfig) -> list[date]:
    if config.end_date < config.start_date:
        raise ValueError("La fecha final no puede ser anterior a la fecha inicial.")
    days = []
    current = config.start_date
    while current <= config.end_date:
        if config.include_weekends or current.weekday() < 5:
            days.append(current)
        current += timedelta(days=1)
    if not days:
        raise ValueError("El periodo seleccionado no contiene dias disponibles.")
    return days


def _unequal_split(total: float, count: int, rng: np.random.Generator, minimum: float = 1.0) -> list[float]:
    if count < 0:
        raise ValueError("El numero de movimientos no puede ser negativo.")
    if count == 0:
        if abs(total) > 0.005:
            raise ValueError("No puede distribuirse un monto positivo entre cero movimientos.")
        return []
    total = round(float(total), 2)
    if total < minimum * count:
        raise ValueError("El monto total es demasiado pequeno para el numero de movimientos.")
    weights = rng.lognormal(mean=0.0, sigma=0.75, size=count)
    raw = minimum + (total - minimum * count) * weights / weights.sum()
    cents = np.floor(raw * 100).astype(int)
    remaining = int(round(total * 100)) - int(cents.sum())
    order = np.argsort(-(raw * 100 - cents))
    for i in range(remaining):
        cents[order[i % count]] += 1
    values = [round(x / 100, 2) for x in cents]
    rng.shuffle(values)
    return values


def _expense_total_for_average(config: GenerationConfig, incomes: list[float], dates: list[date], rng: np.random.Generator) -> float:
    """Aproxima el total de retiros que conduce al saldo promedio solicitado."""
    n = config.income_count + config.expense_count
    if config.expense_count == 0:
        return 0.0
    # Una primera aproximacion estable; el ajuste fino ocurre sobre el ultimo retiro.
    expected_income_effect = sum(incomes) * 0.5
    desired_reduction = config.opening_balance + expected_income_effect - config.target_balance
    return max(float(config.expense_count), round(desired_reduction * 2, 2))


def _unique_counterparty_names(count: int, rng: random.Random) -> list[str]:
    candidates = [
        f"{first} {paternal} {maternal}"
        for first in FIRST_NAMES
        for paternal in LAST_NAMES
        for maternal in LAST_NAMES
        if paternal != maternal
    ]
    rng.shuffle(candidates)
    if count > len(candidates):
        expanded = [
            f"{first} {second} {paternal} {maternal}"
            for first in FIRST_NAMES
            for second in FIRST_NAMES
            for paternal in LAST_NAMES
            for maternal in LAST_NAMES
            if first != second and paternal != maternal
        ]
        rng.shuffle(expanded)
        candidates.extend(expanded)
    if count > len(candidates):
        raise ValueError(
            "Se agotaron las combinaciones de nombres ficticios y no fue posible crear más datos únicos."
        )
    return candidates[:count]


def _varied_transfer_banks(count: int, rng: random.Random, allowed_banks=None) -> list[str]:
    """Distribuye bancos de forma aleatoria, balanceada y sin vecinos iguales."""
    if count <= 0:
        return []
    banks = list(dict.fromkeys(TRANSFER_BANKS if allowed_banks is None else allowed_banks))
    if not banks:
        raise ValueError("Selecciona al menos un banco permitido para los movimientos ficticios.")
    if count > 1 and len(banks) < 2:
        raise ValueError(
            "Para generar más de un movimiento ficticio sin repetir bancos consecutivos debes seleccionar al menos dos bancos."
        )
    result = []
    while len(result) < count:
        cycle = banks[:]
        rng.shuffle(cycle)
        if result and cycle[0] == result[-1]:
            cycle[0], cycle[1] = cycle[1], cycle[0]
        result.extend(cycle)
    return result[:count]


def _rich_description(
    concept: str,
    rng: random.Random,
    movement_date: date,
    incoming: bool,
    person: str | None = None,
) -> str:
    reference = _digits(rng, 12)
    if not any(token in concept.upper() for token in ("SPEI", "TRANSF", "TRASPASO")):
        return f"{concept} REF {reference}"
    person = person or f"{rng.choice(FIRST_NAMES)} {rng.choice(LAST_NAMES)} {rng.choice(LAST_NAMES)}"
    account = _digits(rng, 18)
    letters = "ABCDEFGHJKLMNPQRSTUVWXYZ"
    if incoming:
        trace = _digits(rng, 12) + "".join(rng.choice(letters + "0123456789") for _ in range(8))
    else:
        trace = movement_date.strftime("%Y%m%d") + "40044B36L" + _digits(rng, 13)
    role = "ORDENANTE" if incoming else "BENEFICIARIO"
    return f"{concept} {role}: {person} CUENTA: {account} RASTREO: {trace} REF: {reference} CONCEPTO: TRANSFERENCIA INTERBANCARIA"


def _daily_average(df: pd.DataFrame, start_date: date, end_date: date) -> float:
    """Promedio de los saldos al cierre de cada dia natural del periodo."""
    opening = float(pd.to_numeric(df["SALDO"], errors="coerce").iloc[0])
    movements = df[~df["DESCRIPCION"].astype(str).str.upper().str.strip().eq("SALDO ANTERIOR")].copy()
    movements["FECHA"] = pd.to_datetime(movements["FECHA"]).dt.normalize()
    movements["NETO"] = (
        pd.to_numeric(movements["DEPOSITO"], errors="coerce").fillna(0.0)
        - pd.to_numeric(movements["RETIRO"], errors="coerce").fillna(0.0)
    )
    net_by_day = movements.groupby("FECHA")["NETO"].sum().to_dict()
    running, balances = opening, []
    for day in pd.date_range(start_date, end_date, freq="D"):
        running = round(running + float(net_by_day.get(day.normalize(), 0.0)), 2)
        balances.append(running)
    return round(sum(balances) / len(balances), 2)


def _safe_schedule_generated(rows: list[dict], opening_balance: float) -> list[dict]:
    """Adelanta ingresos ficticios y rechaza cualquier resultado negativo."""
    final_balance = round(
        float(opening_balance)
        + sum(float(row.get("DEPOSITO", 0.0) or 0.0) for row in rows)
        - sum(float(row.get("RETIRO", 0.0) or 0.0) for row in rows),
        2,
    )
    if final_balance < -0.005:
        raise ValueError(
            f"La configuracion produciria un saldo final negativo de ${final_balance:,.2f}. "
            f"Faltan ${abs(final_balance):,.2f}; aumenta ingresos o saldo inicial, o reduce egresos."
        )
    pending, scheduled = [dict(row) for row in rows], []
    available = round(float(opening_balance), 2)
    while pending:
        row = pending.pop(0)
        deposit = float(row.get("DEPOSITO", 0.0) or 0.0)
        withdrawal = float(row.get("RETIRO", 0.0) or 0.0)
        while withdrawal > available + deposit + 0.005:
            income_pos = next((index for index, candidate in enumerate(pending) if float(candidate.get("DEPOSITO", 0.0) or 0.0) > 0.005), None)
            if income_pos is None:
                raise ValueError(
                    f"No hay fondos suficientes para un egreso de ${withdrawal:,.2f}. "
                    "Aumenta el saldo inicial o los ingresos, o reduce los egresos."
                )
            income = pending.pop(income_pos)
            income["FECHA"] = row["FECHA"]
            scheduled.append(income)
            available = round(available + float(income.get("DEPOSITO", 0.0) or 0.0), 2)
        scheduled.append(row)
        available = round(available + deposit - withdrawal, 2)
        if available < -0.005:
            raise ValueError("La configuracion produciria un saldo negativo y fue rechazada.")
    return scheduled


def generate_movements(config: GenerationConfig) -> pd.DataFrame:
    if config.income_count < 0 or config.expense_count < 0:
        raise ValueError("Las cantidades de movimientos deben ser mayores o iguales a cero.")
    dates = _valid_dates(config)
    rng = np.random.default_rng(config.seed)
    py_rng = random.Random(config.seed)
    incomes = _unequal_split(config.total_income, config.income_count, rng)
    if config.target_mode.upper().startswith("SALDO PROMEDIO"):
        expense_total = _expense_total_for_average(config, incomes, dates, rng)
    else:
        expense_total = config.opening_balance + config.total_income - config.target_balance
        if expense_total < -0.005:
            raise ValueError("El saldo final solicitado requiere mas ingresos que los configurados.")
        expense_total = max(0.0, expense_total)
    expenses = _unequal_split(expense_total, config.expense_count, rng) if config.expense_count else []

    income_source = INCOME_CONCEPTS if config.income_concepts is None else config.income_concepts
    expense_source = EXPENSE_CONCEPTS if config.expense_concepts is None else config.expense_concepts
    income_catalog = tuple(item.strip().upper() for item in income_source if str(item).strip())
    expense_catalog = tuple(item.strip().upper() for item in expense_source if str(item).strip())
    if config.income_count and not income_catalog:
        raise ValueError("El catalogo de conceptos de ingreso no puede estar vacio.")
    if config.expense_count and not expense_catalog:
        raise ValueError("El catalogo de conceptos de egreso no puede estar vacio.")
    rows = []
    unique_names = iter(_unique_counterparty_names(config.income_count + config.expense_count, py_rng))
    for i, amount in enumerate(incomes):
        concept = py_rng.choice(income_catalog)
        movement_date = py_rng.choice(dates)
        person = next(unique_names)
        rows.append({"FECHA": movement_date, "DESCRIPCION": _rich_description(concept, py_rng, movement_date, True, person), "DEPOSITO": amount, "RETIRO": 0.0, "NOMBRE_CONTRAPARTE": person, "CONCEPTO_SOLICITADO": concept})
    for i, amount in enumerate(expenses):
        concept = py_rng.choice(expense_catalog)
        movement_date = py_rng.choice(dates)
        person = next(unique_names)
        rows.append({"FECHA": movement_date, "DESCRIPCION": _rich_description(concept, py_rng, movement_date, False, person), "DEPOSITO": 0.0, "RETIRO": amount, "NOMBRE_CONTRAPARTE": person, "CONCEPTO_SOLICITADO": concept})
    py_rng.shuffle(rows)
    rows.sort(key=lambda row: row["FECHA"])

    rows = _safe_schedule_generated(rows, config.opening_balance)
    assigned_banks = _varied_transfer_banks(len(rows), py_rng, config.transfer_banks)
    for row, transfer_bank in zip(rows, assigned_banks):
        row["BANCO_TRANSFERENCIA"] = transfer_bank
        row["CATALOGO_NOMBRES_AMPLIADO"] = config.income_count + config.expense_count > (
            len(FIRST_NAMES) * len(LAST_NAMES) * (len(LAST_NAMES) - 1)
        )

    balance = round(float(config.opening_balance), 2)
    output = [{"FECHA": config.start_date, "DESCRIPCION": "SALDO ANTERIOR", "DEPOSITO": 0.0, "RETIRO": 0.0, "SALDO": balance}]
    for row in rows:
        balance = round(balance + row["DEPOSITO"] - row["RETIRO"], 2)
        output.append({**row, "SALDO": balance})

    df = pd.DataFrame(output)
    if config.target_mode.upper().startswith("SALDO PROMEDIO") and config.expense_count:
        expense_indexes = df.index[df["RETIRO"] > 0].tolist()
        days = (config.end_date - config.start_date).days + 1
        weighted_effect = sum(
            float(df.loc[idx, "RETIRO"])
            * ((config.end_date - pd.Timestamp(df.loc[idx, "FECHA"]).date()).days + 1) / days
            for idx in expense_indexes
        )
        current_avg = _daily_average(df, config.start_date, config.end_date)
        average_without_expenses = current_avg + weighted_effect
        required_effect = average_without_expenses - config.target_balance
        if required_effect < 0:
            raise ValueError("El saldo promedio solicitado requiere ingresos adicionales.")
        scale = required_effect / weighted_effect if weighted_effect else 0.0
        for idx in expense_indexes:
            df.loc[idx, "RETIRO"] = max(0.01, round(float(df.loc[idx, "RETIRO"]) * scale, 2))

        def recalculate() -> None:
            running = float(df.loc[0, "SALDO"])
            for row_idx in df.index[1:]:
                running = round(running + float(df.loc[row_idx, "DEPOSITO"]) - float(df.loc[row_idx, "RETIRO"]), 2)
                df.loc[row_idx, "SALDO"] = running

        recalculate()
        # Ajuste de centavos usando el retiro con mayor permanencia en el periodo.
        for _ in range(5):
            if float(df["SALDO"].min()) < -0.005:
                raise ValueError(
                    "El saldo promedio solicitado obligaria a generar un saldo negativo. "
                    "Aumenta el saldo promedio objetivo, el saldo inicial o los ingresos."
                )
            average_delta = _daily_average(df, config.start_date, config.end_date) - config.target_balance
            if abs(average_delta) <= 0.02:
                break
            adjustment_idx = min(expense_indexes, key=lambda idx: pd.Timestamp(df.loc[idx, "FECHA"]))
            weight = ((config.end_date - pd.Timestamp(df.loc[adjustment_idx, "FECHA"]).date()).days + 1) / days
            adjusted = round(float(df.loc[adjustment_idx, "RETIRO"]) + average_delta / weight, 2)
            if adjusted <= 0:
                raise ValueError("No es posible alcanzar ese saldo promedio con la configuracion actual.")
            df.loc[adjustment_idx, "RETIRO"] = adjusted
            recalculate()
        if abs(_daily_average(df, config.start_date, config.end_date) - config.target_balance) > 0.05:
            raise ValueError("No fue posible alcanzar el saldo promedio mensual solicitado con esta configuracion.")
    if float(pd.to_numeric(df["SALDO"], errors="coerce").min()) < -0.005:
        raise ValueError("No se genero el ejercicio porque produciria uno o mas saldos negativos.")
    df["FECHA"] = pd.to_datetime(df["FECHA"])
    df.attrs.update({"period_start": pd.Timestamp(config.start_date), "period_end": pd.Timestamp(config.end_date)})
    return df


def merge_base_and_generated(base: pd.DataFrame, generated: pd.DataFrame) -> pd.DataFrame:
    """Conserva el archivo base y recalcula todos los saldos cronologicamente."""
    columns = ["FECHA", "DESCRIPCION", "DEPOSITO", "RETIRO", "SALDO"]
    metadata = [c for c in ("_LINEAS_DOCUMENTO", "FECHA_LIQ_BBVA", "AJUSTE_LIQ_BBVA", "BANCO_TRANSFERENCIA", "NOMBRE_CONTRAPARTE", "CONCEPTO_SOLICITADO", "ORIGEN", "REFERENCIA_CONCEPTO_BBVA") if c in base.columns or c in generated.columns]
    columns += metadata
    base = base.reindex(columns=columns).copy()
    opening = float(pd.to_numeric(base["SALDO"], errors="coerce").dropna().iloc[0]) if len(base) else 0.0
    additions = generated[generated["DESCRIPCION"] != "SALDO ANTERIOR"].copy()
    combined = pd.concat([base, additions], ignore_index=True)
    combined["FECHA"] = pd.to_datetime(combined["FECHA"], errors="coerce")
    combined["DEPOSITO"] = pd.to_numeric(combined["DEPOSITO"], errors="coerce").fillna(0.0)
    combined["RETIRO"] = pd.to_numeric(combined["RETIRO"], errors="coerce").fillna(0.0)
    combined = combined.sort_values("FECHA", kind="stable").reset_index(drop=True)
    balance = opening
    for idx in combined.index:
        if str(combined.loc[idx, "DESCRIPCION"]).upper().strip() == "SALDO ANTERIOR":
            combined.loc[idx, "SALDO"] = balance
            continue
        balance = round(balance + float(combined.loc[idx, "DEPOSITO"]) - float(combined.loc[idx, "RETIRO"]), 2)
        combined.loc[idx, "SALDO"] = balance
    if float(pd.to_numeric(combined["SALDO"], errors="coerce").min()) < -0.005:
        raise ValueError(
            "Al combinar el archivo base con los movimientos nuevos se produciria un saldo negativo. "
            "Aumenta el saldo inicial o los ingresos, o reduce/reprograma los egresos."
        )
    result = combined[columns]
    result.attrs.update(generated.attrs)
    return result


def expand_concepts(df: pd.DataFrame) -> pd.DataFrame:
    """Convierte cada renglon visual del concepto en una fila independiente."""
    rows = []
    for _, row in df.iterrows():
        text = str(row.get("DESCRIPCION", "")).strip()
        parts = [p.strip() for p in re.split(r"\s+(?=(?:REF|RASTREO|CUENTA|EMISOR|RECEPTOR|ORDENANTE|BENEFICIARIO|CONCEPTO|HORA|CLAVE)\b)", text) if p.strip()]
        if not parts:
            parts = [""]
        for index, part in enumerate(parts):
            new_row = row.copy()
            new_row["DESCRIPCION"] = part
            if index:
                new_row["FECHA"] = pd.NaT
                new_row["DEPOSITO"] = 0.0
                new_row["RETIRO"] = 0.0
                new_row["SALDO"] = np.nan
            rows.append(new_row)
    return pd.DataFrame(rows, columns=df.columns)


def movement_checks(df: pd.DataFrame) -> dict:
    dep = pd.to_numeric(df["DEPOSITO"], errors="coerce").fillna(0.0)
    ret = pd.to_numeric(df["RETIRO"], errors="coerce").fillna(0.0)
    balances = pd.to_numeric(df["SALDO"], errors="coerce").dropna()
    opening = float(balances.iloc[0]) if len(balances) else 0.0
    expected = round(opening + float(dep.sum()) - float(ret.sum()), 2)
    actual = round(float(balances.iloc[-1]), 2) if len(balances) else 0.0
    dates = pd.to_datetime(df["FECHA"], errors="coerce").dropna()
    period_start = pd.Timestamp(df.attrs.get("period_start", dates.min())).date() if len(dates) else datetime.now().date()
    period_end = pd.Timestamp(df.attrs.get("period_end", dates.max())).date() if len(dates) else period_start
    daily_average = _daily_average(df, period_start, period_end) if len(df) else actual
    minimum = round(float(balances.min()), 2) if len(balances) else actual
    return {
        "INGRESOS": round(float(dep.sum()), 2),
        "EGRESOS": round(float(ret.sum()), 2),
        "SALDO_FINAL": actual,
        "SALDO_PROMEDIO": daily_average,
        "SALDO_MINIMO": minimum,
        "SALDOS_NO_NEGATIVOS": minimum >= -0.005 and actual >= -0.005,
        "CUADRA": abs(expected - actual) < 0.01,
        "DIFERENCIA": round(actual - expected, 2),
    }
