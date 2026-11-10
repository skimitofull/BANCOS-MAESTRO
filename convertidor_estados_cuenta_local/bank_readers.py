"""Read-only Afirme and Monex readers. Extracted from Alpha converter; no PDF generation."""
import re,json,io,unicodedata
from decimal import Decimal,InvalidOperation
from datetime import datetime
import fitz,pandas as pd
MONTHS={'ENERO': 1, 'FEBRERO': 2, 'MARZO': 3, 'ABRIL': 4, 'MAYO': 5, 'JUNIO': 6, 'JULIO': 7, 'AGOSTO': 8, 'SEPTIEMBRE': 9, 'OCTUBRE': 10, 'NOVIEMBRE': 11, 'DICIEMBRE': 12}
SHORT={k[:3]:v for k,v in MONTHS.items()}
MONEY=re.compile(r'^\(?-?\$?\s*\d[\d,]*\.\d{2}\)?(?:[+-])?$')
def plain(text):
    return ''.join(c for c in unicodedata.normalize('NFKD',text.upper()) if not unicodedata.combining(c))

def lines(words,tolerance=3):
    result=[]
    for word in sorted(words,key=lambda w:(w['top'],w['x0'])):
        match=next((row for row in reversed(result[-4:]) if abs(row['top']-word['top'])<=tolerance),None)
        if match is None:
            match={'top':word['top'],'words':[]}; result.append(match)
        match['words'].append(word)
    for row in result:
        row['words'].sort(key=lambda w:w['x0'])
        row['text']=' '.join(w['text'] for w in row['words'])
    return sorted(result,key=lambda row:row['top'])

def amount(value):
    try:
        text=str(value).replace('$','').replace(',','').replace(' ','')
        if text.startswith('('):text='-'+text.strip('()')
        return float(Decimal(text))
    except InvalidOperation: return value
def clean(text):
    return str(text).replace('\xa0',' ').replace('‐','-').replace('–','-').replace('ﬁ','fi').replace('ﬂ','fl')

def date(value):
    value=plain(clean(value))
    m=re.search(r'\b(\d{1,2})/(\d{1,2})/(\d{2,4})\b',value)
    if m:
        d,mo,y=map(int,m.groups());y=y+2000 if y<100 else y
    else:
        m=re.search(r'\b(20\d{2})-(\d{2})-(\d{2})\b',value)
        if m:y,mo,d=map(int,m.groups())
        else:
            m=re.search(r'\b(\d{1,2})\s+(?:DE\s+)?([A-Z]+)\.?\s+(?:(?:DE|DEL)\s+)?(20\d{2})\b',value)
            if not m:return ''
            d,month,y=m.groups();d=int(d);y=int(y);mo=MONTHS.get(month,SHORT.get(month,0))
    try:return datetime(y,mo,d).strftime('%Y-%m-%d')
    except ValueError:return ''

def afirme(pages):
    cover='\n'.join(plain(row['text']) for row in pages[0]['lines'])
    period=re.search(r'PERIODO DE\s+(\d{1,2}\s+[A-Z]{3}\s+20\d{2})\s+AL\s+(\d{1,2}\s+[A-Z]{3}\s+20\d{2})',cover)
    if not period:raise ValueError('Afirme: no se identificó el periodo impreso. Incluye la carátula del estado de cuenta.')
    start,end=date(period[1]),date(period[2])
    if not start or not end or start[:7]!=end[:7]:raise ValueError('Afirme: revisa el periodo; este lector necesita mes y año inequívocos para los días de los movimientos.')
    def cover_value(label):
        m=re.search(label+r'\s*\$\s*([\d,]+\.\d{2})',cover)
        if not m:raise ValueError('Afirme: falta el total de '+label.lower()+' en la carátula.')
        return amount(m[1])
    totals=dict(bank='AFIRME',opening=cover_value('SALDO INICIAL'),deposits=cover_value('DEPOSITOS'),withdrawals=cover_value('RETIROS'),final=cover_value('SALDO AL CORTE'),period_start=start,period_end=end)
    movements=[]
    for page in pages:
        active=False;current=None;anchors={};reference_x=0;date_x=0
        for row in page['lines']:
            words=row['words'];upper=plain(row['text'])
            if 'DESCRIPCION' in upper and 'DEPOSITOS' in upper and 'RETIROS' in upper and 'SALDO' in upper:
                for w in words:
                    token=plain(w['text'])
                    if token in ('DEPOSITOS','RETIROS','SALDO'):anchors[token]=w['x1']
                    elif token=='REFERENCIA':reference_x=w['x0']
                    elif token=='DIA':date_x=w['x1']
                active=len(anchors)==3
                continue
            if not active:continue
            if upper.startswith(('DEPOSITOS RETIROS EN EFECTIVO','METODO DE PAGO','CADENA ORIGINAL','IMPORTANTE','RESUMEN','SUS AHORROS','SALDO INICIAL','SI DESEA RECIBIR')):
                active=False;current=None;continue
            first=words[0] if words else {}
            amounts=[w for w in words if MONEY.fullmatch(w['text']) and w['x1']>=min(anchors.values())-60]
            new=bool(first and re.fullmatch(r'\d{1,2}',first['text']) and first['x1']<=date_x+10 and amounts)
            if new:
                day=int(first['text'])
                full=date(f'{day}/{start[5:7]}/{start[:4]}')
                if not full or not start<=full<=end:raise ValueError('Afirme: día fuera del periodo en la página '+str(page['number']))
                current=dict(FECHA=full,DESCRIPCION='',DEPOSITO='',RETIRO='',SALDO='',REFERENCIA='',PAGINA=page['number'],_FILAS=[])
                movements.append(current)
            if current is None:continue
            visual=dict(FECHA=current['FECHA'] if new else '',DESCRIPCION='',DEPOSITO='',RETIRO='',SALDO='',REFERENCIA='')
            description=[]
            for i,w in enumerate(words):
                if i==0 and new:continue
                if w in amounts:
                    column=min(anchors,key=lambda k:abs(w['x1']-anchors[k]))
                    if abs(w['x1']-anchors[column])>page['width']*.06:raise ValueError('Afirme: importe fuera de las columnas esperadas.')
                    visual[{'DEPOSITOS':'DEPOSITO','RETIROS':'RETIRO','SALDO':'SALDO'}[column]]=amount(w['text']);continue
                if w['text']=='$' and w['x0']>min(anchors.values())-80:continue
                if new and re.fullmatch(r'\d+',w['text']) and reference_x-5<=w['x0']<min(anchors.values())-55:visual['REFERENCIA']=w['text'];continue
                description.append(w['text'])
            visual['DESCRIPCION']=' '.join(description)
            current['_FILAS'].append(visual)
            for key in ('DEPOSITO','RETIRO','SALDO','REFERENCIA'):
                if visual[key]!='':current[key]=visual[key]
    if not movements:raise ValueError('Afirme: no se identificaron operaciones en las columnas del PDF.')
    running=Decimal(str(totals['opening']))
    for i,mov in enumerate(movements,1):
        running+=Decimal(str(mov['DEPOSITO'] or 0))-Decimal(str(mov['RETIRO'] or 0))
        if mov['SALDO']=='' or abs(running-Decimal(str(mov['SALDO'])))>Decimal('.02'):raise ValueError(f'Afirme: no cuadra el saldo del movimiento {i}; revisar la página {mov["PAGINA"]}. Calculado {running}, leído {mov["SALDO"]}.')
        mov['DESCRIPCION']=' '.join(r['DESCRIPCION'] for r in mov['_FILAS'])
    deposits=sum(Decimal(str(m['DEPOSITO'] or 0)) for m in movements)
    withdrawals=sum(Decimal(str(m['RETIRO'] or 0)) for m in movements)
    for name,value in [('deposits',deposits),('withdrawals',withdrawals),('final',running)]:
        if abs(value-Decimal(str(totals[name])))>Decimal('.02'):raise ValueError('Afirme: la extracción no coincide con el total impreso de '+name+'.')
    totals['count']=len(movements)
    return movements,totals

def monex(content):
    import fitz
    doc=fitz.open(stream=content,filetype='pdf');pages=[]
    with doc:
        first=plain(doc[0].get_text())
        period=re.search(r'DEL\s+(\d{1,2}\s+[A-Z]+\s+20\d{2})\s+AL\s+(\d{1,2}\s+[A-Z]+\s+20\d{2})',first)
        if not period:raise ValueError('Monex: falta la carátula con el periodo del estado de cuenta.')
        start,end=date(period[1]),date(period[2])
        if not start or not end:raise ValueError('Monex: periodo no reconocido.')
        contract=re.search(r'CONTRATO:\s*(\d+)',first);contract=contract[1] if contract else ''
        for i,p in enumerate(doc):
            words=[dict(text=w[4],x0=w[0],x1=w[2],top=w[1],bottom=w[3]) for w in p.get_text('words')]
            bands=[(item[1].y0,item[1].y1) for draw in p.get_drawings() if draw['fill'] for item in draw['items'] if item[0]=='re' and item[1].width>p.rect.width*.7 and item[1].height>2]
            pages.append(dict(number=i+1,text=plain(p.get_text()),lines=lines(words,2),bands=bands,width=p.rect.width,height=p.rect.height))
    controls={};currency='MXN';movements=[];account_usd=''
    for page in pages:
        found=re.search(r'(?m)^CUENTA\s+(\d+)\s*$',page['text'])
        if found and 'MOVIMIENTOS DE' in page['text']:account_usd=found[1]
    for page in pages:
        text=page['text'];summary=re.search(r'RESUMEN CUENTA\s+(PESO MEXICANO|DOLAR AMERICANO)',text)
        if summary:
            currency='USD' if summary[1]=='DOLAR AMERICANO' else 'MXN'
            def total(pattern):
                match=re.search(pattern+r'\s*([\d,]+\.\d{2})',text)
                if not match:raise ValueError('Monex: falta un total impreso de '+currency+'.')
                return amount(match[1])
            controls[currency]=dict(bank='MONEX',currency=currency,opening=total('SALDO INICIAL:'),deposits=total(r'\+ TOTAL ABONOS:'),withdrawals=total('- TOTAL CARGOS:'),final=total('SALDO TOTAL:'),period_start=start,period_end=end)
        rows=page['lines'];header=next((r for r in rows if 'ABONOS' in plain(r['text']) and 'CARGOS' in plain(r['text']) and 'REFERENCIA' in plain(r['text'])),None)
        if not header:
            if 'MOVIMIENTOS DE' in text:raise ValueError('Monex: encabezados de movimientos no reconocidos en página '+str(page['number']))
            continue
        if re.search(r'MOVIMIENTOS DE[^\n]*DOLAR AMERICANO',text):currency='USD'
        if currency not in controls:raise ValueError('Monex: incluye el resumen de la cuenta '+currency+' para validar los importes.')
        ref=next(w['x0'] for w in header['words'] if plain(w['text'])=='REFERENCIA')
        desc=next(w['x0'] for w in header['words'] if plain(w['text'])=='DESCRIPCION')
        abono=next(w['x1'] for w in header['words'] if plain(w['text'])=='ABONOS')
        heads=[]
        for r in rows:
            dates=''.join(w['text'] for w in r['words'] if w['x0']<desc-5)
            match=re.fullmatch(r'(\d{1,2}/[A-Za-z]{3})(?:\((\d{1,2}/[A-Za-z]{3})\))?',dates)
            money_words=[w for w in r['words'] if MONEY.fullmatch(w['text']) and w['x1']>=abono-5]
            if match and len(money_words)==6:heads.append((r,match,sorted(money_words,key=lambda w:w['x1'])))
        if not heads:continue
        if not page['bands']:raise ValueError('Monex: no se identificaron los límites de las filas; usa Tabla PDF para este diseño.')
        for r,match,money_words in heads:
            center=(r['words'][0]['top']+r['words'][0]['bottom'])/2
            band=next(((a,b) for a,b in page['bands'] if a<=center<=b),None)
            if band:low,high=band
            else:
                low=max((b for a,b in page['bands'] if b<center),default=header['top']+18)
                high=min((a for a,b in page['bands'] if a>center),default=page['height']-35)
            def full_date(value):
                day,month=value.split('/')
                return date(f'{day}/{SHORT.get(month.upper(),0)}/{start[:4]}')
            liquid=full_date(match[1]);pact=full_date(match[2]) if match[2] else ''
            if not liquid or not start<=liquid<=end:raise ValueError('Monex: fecha de liquidación fuera del periodo.')
            numbers=[amount(w['text']) for w in money_words]
            reference=' '.join(w['text'] for w in r['words'] if ref-3<=w['x0']<abono-65)
            visual=[]
            for line in rows:
                parts=[w['text'] for w in line['words'] if desc-2<=w['x0'] and w['x1']<ref-1 and low<=w['top']<high]
                if parts:visual.append(' '.join(parts))
            mov=dict(FECHA=liquid,DESCRIPCION=' '.join(visual),DEPOSITO=numbers[0],RETIRO=numbers[1],SALDO=numbers[5],MONEDA=currency,REFERENCIA=reference,CONTRATO=contract,CUENTA=account_usd if currency=='USD' else '',**{'FECHA PACTADA':pact,'MOVIMIENTO GARANTÍA':numbers[2],'SALDO GARANTÍA':numbers[3],'SALDO DISPONIBLE':numbers[4],'PAGINA':page['number'],'_FILAS':[]})
            for i,line in enumerate(visual or ['']):
                row={k:(v if i==0 else '') for k,v in mov.items() if k not in ('_FILAS','PAGINA')};row['DESCRIPCION']=line;mov['_FILAS'].append(row)
            movements.append(mov)
        # A row can continue at the top of the following page before its first band.
        first_low=min(page['bands'])[0]
        continuation=[]
        for r in rows:
            if header['top']+18<r['top']<first_low:
                parts=[w['text'] for w in r['words'] if desc-2<=w['x0'] and w['x1']<ref-1]
                value=' '.join(parts)
                if value and not value.upper().startswith(('CUENTA ','DESCRIP')):continuation.append(value)
        previous=[m for m in movements if m['PAGINA']<page['number'] and m['MONEDA']==currency]
        if previous and continuation:
            for line in continuation:
                previous[-1]['_FILAS'].append(dict(DESCRIPCION=line))
            previous[-1]['DESCRIPCION']+=' '+' '.join(continuation)
    for currency,control in controls.items():
        rows=[m for m in movements if m['MONEDA']==currency]
        if not rows:raise ValueError('Monex: no se identificaron movimientos de '+currency+'.')
        running=Decimal(str(control['opening']))
        for mov in rows:
            running+=Decimal(str(mov['DEPOSITO']))-Decimal(str(mov['RETIRO']))
            if abs(running-Decimal(str(mov['SALDO'])))>Decimal('.02'):raise ValueError('Monex: no cuadra el saldo '+currency+' en referencia '+mov['REFERENCIA']+'.')
        for key,value in [('deposits',sum(m['DEPOSITO'] for m in rows)),('withdrawals',sum(m['RETIRO'] for m in rows)),('final',float(running))]:
            if abs(Decimal(str(value))-Decimal(str(control[key])))>Decimal('.02'):raise ValueError(f'Monex: no coincide {key} en {currency}: extraído {value}, impreso {control[key]}, {len(rows)} movimientos.')
        control['count']=len(rows)
    return movements,dict(bank='MONEX',count=len(movements),currencies=list(controls.values()))

def extract_bank_movements(content,bank):
    if bank=='AFIRME':
        with fitz.open(stream=content,filetype='pdf') as doc:
            pages=[]
            for i,p in enumerate(doc):
                words=[dict(text=w[4],x0=w[0],x1=w[2],top=w[1],bottom=w[3]) for w in p.get_text('words')]
                pages.append(dict(number=i+1,width=p.rect.width,lines=lines(words,2)))
        rows,control=afirme(pages)
        rows.insert(0,dict(FECHA=control['period_start'],DESCRIPCION='SALDO ANTERIOR',DEPOSITO=0.,RETIRO=0.,SALDO=control['opening'],_FILAS=[dict(DESCRIPCION='SALDO ANTERIOR')]))
    else:rows,control=monex(content)
    records=[]
    for row in rows:
        record={k:v for k,v in row.items() if k!='_FILAS'}
        record['FECHA']=pd.Timestamp(record['FECHA'])
        record['_LINEAS_DOCUMENTO']=json.dumps([r.get('DESCRIPCION','') for r in row['_FILAS']],ensure_ascii=False)
        records.append(record)
    frame=pd.DataFrame(records);frame.attrs['bank_control']=control
    if 'period_start' in control:frame.attrs.update(period_start=control['period_start'],period_end=control['period_end'])
    return frame

def export_monex(frame):
    control=frame.attrs['bank_control'];output=io.BytesIO()
    with pd.ExcelWriter(output,engine='xlsxwriter') as writer:
        summary=[]
        for c in control['currencies']:
            frame[frame.MONEDA==c['currency']].drop(columns=['_LINEAS_DOCUMENTO'],errors='ignore').to_excel(writer,sheet_name='Movimientos '+c['currency'],index=False)
            for label,key in [('Movimientos','count'),('Saldo inicial','opening'),('Abonos','deposits'),('Cargos','withdrawals'),('Saldo final','final')]:summary.append(dict(MONEDA=c['currency'],CONCEPTO=label,VALOR=c[key]))
        pd.DataFrame(summary).to_excel(writer,sheet_name='Resumen por moneda',index=False)
        for name,ws in writer.sheets.items():
            ws.freeze_panes(1,0);ws.set_column(0,0,14);ws.set_column(1,1,90);ws.set_column(2,18,20)
    return output.getvalue()
