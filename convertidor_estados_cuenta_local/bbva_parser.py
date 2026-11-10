"""Lectura por columnas de BBVA Libretón, con conciliación del resumen."""
import re
import fitz
import pandas as pd
from excel_layouts import PHYSICAL_LINES_COLUMN, encode_physical_lines

MONTHS = {m: i + 1 for i, m in enumerate('ENE FEB MAR ABR MAY JUN JUL AGO SEP OCT NOV DIC'.split())}
DATE = re.compile(r'^\d{2}/[A-Z]{3}$')
MONEY = re.compile(r'^\d[\d,]*\.\d{2}$')

def extract_bbva_movements(blob):
    doc = fitz.open(stream=blob, filetype='pdf')
    text = '\n'.join(p.get_text() for p in doc)
    period = re.search(r'DEL\s+(\d{2}/\d{2}/\d{4})\s+AL', text)
    if not period:
        raise ValueError('BBVA: no se encontró el periodo del estado de cuenta.')
    start = pd.to_datetime(period[1], dayfirst=True)
    opening_match = re.search(r'Saldo Anterior\s+([\d,]+\.\d{2})', text)
    if not opening_match:
        raise ValueError('BBVA: no se encontró el saldo anterior.')
    opening = float(opening_match[1].replace(',', ''))
    rows = []
    current = None
    for page in doc:
        words = page.get_text('words')
        dates = [w for w in words if w[0] < 50 and DATE.fullmatch(w[4])]
        if not dates:
            continue
        header = [w[1] for w in words if w[4] == 'OPER' and w[0] < 50]
        if not header:
            continue
        low = max(header) + 6
        stop = min([w[1] for w in words if w[4] == 'Total' and w[1] > low] + [764.0])
        lines = []
        for w in sorted(words, key=lambda w: (w[1], w[0])):
            if not low < w[1] < stop:
                continue
            if lines and abs(lines[-1][0] - w[1]) < 2:
                lines[-1][1].append(w)
            else:
                lines.append([w[1], [w]])
        for _, group in lines:
            group.sort(key=lambda w: w[0])
            op = next((w[4] for w in group if w[0] < 50 and DATE.fullmatch(w[4])), None)
            desc = ' '.join(w[4] for w in group if 80 <= w[0] < 360)
            if op:
                if current:
                    rows.append(current)
                day, month = op.split('/')
                date = pd.Timestamp(start.year - (MONTHS[month] > start.month and start.month == 1), MONTHS[month], int(day))
                liq = next((w[4] for w in group if 50 <= w[0] < 82 and DATE.fullmatch(w[4])), op)
                money = [(w[0], float(w[4].replace(',', ''))) for w in group if MONEY.fullmatch(w[4])]
                current = {'FECHA': date, 'FECHA_LIQ_BBVA': liq, 'DEPOSITO': sum(v for x,v in money if 420 <= x < 470), 'RETIRO': sum(v for x,v in money if 360 <= x < 420), 'SALDO': None, '_lines': [desc]}
                balances = [v for x,v in money if 470 <= x < 535]
                if balances:
                    current['_printed_balance'] = balances[0]
                    liquid_balances = [v for x,v in money if x >= 535]
                    current['AJUSTE_LIQ_BBVA'] = (liquid_balances[0] - balances[0]) if liquid_balances else 0.
            elif current and desc and not any(s in desc for s in ('No. de Cuenta', 'No. de Cliente', 'Libretón Premium', 'Estado de Cuenta')):
                current['_lines'].append(desc)
    if current:
        rows.append(current)
    running = opening
    for n, row in enumerate(rows, 1):
        if row['DEPOSITO'] == row['RETIRO'] == 0:
            raise ValueError(f'BBVA: movimiento {n} sin importe en cargos o abonos.')
        running = round(running + row['DEPOSITO'] - row['RETIRO'], 2)
        printed = row.pop('_printed_balance', None)
        if printed is not None and abs(printed-running) > .02:
            raise ValueError(f'BBVA: saldo incongruente en movimiento {n}: {printed:,.2f} frente a {running:,.2f}.')
        row['SALDO'] = running
        physical = row.pop('_lines')
        row['DESCRIPCION'] = ' '.join(physical)
        row[PHYSICAL_LINES_COLUMN] = encode_physical_lines(physical)
    for label, col in [('CARGOS', 'RETIRO'), ('ABONOS', 'DEPOSITO')]:
        match = re.search(r'TOTAL IMPORTE '+label+r'\s+([\d,]+\.\d{2})', text)
        count = re.search(r'TOTAL MOVIMIENTOS '+label+r'\s+(\d+)', text)
        if match and abs(sum(r[col] for r in rows)-float(match[1].replace(',', ''))) > .02:
            raise ValueError('BBVA: los importes no concilian con el resumen de '+label+'.')
        if count and sum(r[col] > 0 for r in rows) != int(count[1]):
            raise ValueError('BBVA: el número de movimientos no concilia con '+label+'.')
    return pd.DataFrame([{'FECHA': start, 'DESCRIPCION': 'SALDO ANTERIOR', 'DEPOSITO': 0., 'RETIRO': 0., 'SALDO': opening, PHYSICAL_LINES_COLUMN: encode_physical_lines(['SALDO ANTERIOR'])}, *rows])
