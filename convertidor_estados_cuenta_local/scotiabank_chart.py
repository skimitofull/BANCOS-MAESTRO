from __future__ import annotations

from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
from PIL import Image, ImageDraw, ImageFont


BASE = Path(__file__).resolve().parent
FONT_PATH = BASE / "fuentes" / "ArialUnicodeMS.ttf"


def _column(frame: pd.DataFrame, *names: str) -> pd.Series:
    normalized = {str(col).strip().upper(): col for col in frame.columns}
    for name in names:
        key = name.strip().upper()
        if key in normalized:
            return frame[normalized[key]]
    return pd.Series([0.0] * len(frame), index=frame.index)


def _font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    if FONT_PATH.exists():
        return ImageFont.truetype(str(FONT_PATH), size=size)
    return ImageFont.load_default()


def _money(value: float) -> str:
    sign = "-" if value < 0 else ""
    return f"{sign}${abs(value):,.2f}"


def _statement_values(frame: pd.DataFrame) -> tuple[list[float], float, float]:
    data = frame.copy()
    descriptions = _column(data, "DESCRIPCION", "CONCEPTO").fillna("").astype(str).str.upper()
    deposits = pd.to_numeric(_column(data, "DEPOSITO", "DEPÓSITO", "ABONO"), errors="coerce").fillna(0.0)
    withdrawals = pd.to_numeric(_column(data, "RETIRO", "CARGO"), errors="coerce").fillna(0.0)
    balances = pd.to_numeric(_column(data, "SALDO"), errors="coerce")

    interest_mask = descriptions.str.contains(r"\bINTERE(?:S|SES)\b", regex=True)
    commission_mask = descriptions.str.contains("COMISION", regex=False)
    cash_mask = descriptions.str.contains(r"EFECTIVO|CAJERO|\bATM\b", regex=True)

    interest = float(deposits[interest_mask].sum())
    deposit = float(deposits[~interest_mask].sum())
    cash = -float(withdrawals[cash_mask & ~commission_mask].sum())
    commission = -float(withdrawals[commission_mask].sum())
    other = -float(withdrawals[~cash_mask & ~commission_mask].sum())

    valid = balances.dropna()
    if valid.empty:
        opening = 0.0
        closing = deposit + interest + cash + other + commission
    else:
        first_index = valid.index[0]
        first_balance = float(valid.iloc[0])
        if "SALDO ANTERIOR" in descriptions.loc[first_index]:
            opening = first_balance
        else:
            opening = first_balance - float(deposits.loc[first_index]) + float(withdrawals.loc[first_index])
        closing = float(valid.iloc[-1])
    return [deposit, interest, cash, other, commission], opening, closing


def generate_scotiabank_chart_png(frame: pd.DataFrame) -> bytes:
    """Genera la gráfica de carátula Scotiabank como PNG RGBA transparente."""
    values, opening, closing = _statement_values(frame)
    labels = ["Depósitos", "Intereses", "Retiros en efectivo", "Otros cargos*", "Comisiones"]

    width, height = 1400, 700
    image = Image.new("RGBA", (width, height), (255, 255, 255, 0))
    draw = ImageDraw.Draw(image)
    regular = _font(30)
    small = _font(26)
    title_font = _font(31)

    draw.text((20, 14), "Comportamiento de transacciones en tu cuenta", font=title_font, fill=(0, 0, 0, 255))
    balance_text = f"Saldo inicial = {_money(opening)}    Saldo final = {_money(closing)}"
    balance_box = draw.textbbox((0, 0), balance_text, font=regular)
    draw.text((width - 24 - (balance_box[2] - balance_box[0]), 58), balance_text, font=regular, fill=(0, 0, 0, 255))

    left, right, top, bottom = 170, width - 25, 125, 545
    max_abs = max([abs(value) for value in values] + [1.0])
    scale_limit = max_abs * 1.22
    baseline = int(top + (bottom - top) / 2)
    pixels_per_unit = ((bottom - top) / 2) / scale_limit

    draw.rectangle((left, top, right, bottom), outline=(0, 0, 0, 255), width=2)
    draw.line((left, baseline, right, baseline), fill=(105, 105, 105, 255), width=2)

    ticks = np.linspace(-scale_limit, scale_limit, 7)
    for tick in ticks:
        y = int(baseline - tick * pixels_per_unit)
        label = f"{tick:,.2f}"
        box = draw.textbbox((0, 0), label, font=small)
        draw.text((left - 14 - (box[2] - box[0]), y - 14), label, font=small, fill=(0, 0, 0, 255))

    plot_width = right - left
    step = plot_width / len(values)
    bar_width = int(step * 0.58)
    for index, (label, value) in enumerate(zip(labels, values)):
        center = int(left + step * (index + 0.5))
        value_y = int(baseline - value * pixels_per_unit)
        if abs(value) > 0.004:
            y1, y2 = sorted((baseline, value_y))
            draw.rectangle((center - bar_width // 2, y1, center + bar_width // 2, y2), fill=(135, 135, 135, 255))
        value_text = _money(value)
        value_box = draw.textbbox((0, 0), value_text, font=small)
        value_text_y = value_y - 38 if value >= 0 else value_y + 8
        draw.text((center - (value_box[2] - value_box[0]) / 2, value_text_y), value_text, font=small, fill=(0, 0, 0, 255))

        label_box = draw.textbbox((0, 0), label, font=small)
        draw.text((center - (label_box[2] - label_box[0]) / 2, bottom + 18), label, font=small, fill=(0, 0, 0, 255))

    draw.text((25, height - 58), "* Incluye impuestos", font=small, fill=(0, 0, 0, 255))
    output = BytesIO()
    image.save(output, format="PNG", optimize=True)
    return output.getvalue()
