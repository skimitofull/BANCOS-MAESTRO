"""Select reader: independent columns, physical ink lines and financial reconciliation."""
from pathlib import Path
from decimal import Decimal
import sys,re,json,hashlib
import fitz,pandas as pd
LIBS=Path(r'C:\Users\clopez\Documents\ChatGPT\LOCAL CLC\tmp\santander_v2\libs')
MONTHS={m:i+1 for i,m in enumerate(['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'])}
def compact(s):return re.sub(r'\s+','',str(s)).upper()
def money(s):
    s=re.sub(r'[\s$]','',str(s));s=s.replace('O','0').replace('o','0')
    s=re.sub(r'^[^\d\-]+','',s)
    if not re.fullmatch(r'-?\d[\d,.]*[.,]\d{2}',s):return None
    return float(re.sub(r'[.,]','',s[:-3])+'.'+s[-2:])
def restore_known_spaces(text):
    # These are fixed labels visible in the reference, not inferred counterparty data.
    aliases=[('VAPORCOMISION','IVA POR COMISION '),('VAPORINTERESES','IVA POR INTERESES '),('(1)DATONOVERIFICADOPORESTAINSTITUCION','(1) DATO NO VERIFICADO POR ESTA INSTITUCION ')]
    for key,label in aliases:
        if compact(text).startswith(key):
            m=re.match(''.join(re.escape(c)+r'\s*' for c in key),text,re.I)
            return label+text[m.end():] if m else label+compact(text)[len(key):]
    prefixes=[('ADMINISTRACIONRENTA','ADMINISTRACION RENTA '),('IVAPORCOMISION','IVA POR COMISION '),('ABONOTRANSFERENCIASPEIHORA','ABONO TRANSFERENCIA SPEI HORA '),('PAGOTRANSFERENCIASPEIHORA','PAGO TRANSFERENCIA SPEI HORA '),('RECIBIDODE','RECIBIDO DE '),('ENVIADOA','ENVIADO A '),('DELACUENTA','DE LA CUENTA '),('ALACUENTA','A LA CUENTA '),('DELCLIENTE','DEL CLIENTE '),('ALCLIENTE','AL CLIENTE '),('CLAVEDERASTREO','CLAVE DE RASTREO '),('CONCEPTO','CONCEPTO '),('REF','REF '),('RFC','RFC '),('CARGOPAGOTARJETACREDITO','CARGO PAGO TARJETA CREDITO '),('CARGOPORINTERESESDECREDITO','CARGO POR INTERESES DE CREDITO '),('CARGOCAPITALDECREDITO','CARGO CAPITAL DE CREDITO '),('IVAPORINTERESES','IVA POR INTERESES '),('COBROPRIMADESEGUROS','COBRO PRIMA DE SEGUROS ')]
    norm=compact(text)
    for key,label in prefixes:
        if norm.startswith(key):
            # Keep remaining OCR characters; do not guess references or names.
            m=re.match(''.join(re.escape(c)+r'\s*' for c in key),text,re.I)
            return label+text[m.end():] if m else label+norm[len(key):]
    return text
def bands(mask):
    groups=[];a=None
    for i,ink in enumerate(mask):
        if ink and a is None:a=i
        if not ink and a is not None:groups.append((a,i));a=None
    if a is not None:groups.append((a,len(mask)))
    return [(a,b) for a,b in groups if b-a>=3]
def recognize(engine,image):
    result,_=engine(image,use_det=False,use_cls=False)
    if not result:return '',0.
    if len(result[0])==2:
        return ' '.join(str(r[0]) for r in result),min(float(r[1]) for r in result)
    return ' '.join(str(r[1]) for r in result),min(float(r[2]) for r in result)
def read_cells(engine,image,x0,x1,start,end,scale=3,field=None):
    import numpy as np
    col=image[round(start*scale):round(end*scale),round(x0*scale):round(x1*scale),:3]
    ink=(col.min(axis=2)<115).sum(axis=1)>2
    detected,_=engine(col,use_cls=False)
    alternatives=[]
    for box,text,conf in detected or []:
        yy=start+sum(v[1] for v in box)/len(box)/scale
        alternatives.append((yy,text,float(conf)))
    rows=[]
    for a,b in bands(ink):
        y=start+(a+b)/2/scale
        pad=9
        crop=col[max(0,a-pad):min(len(col),b+pad)]
        candidates=[r for r in alternatives if abs(r[0]-y)<3]
        if candidates:
            _,text,conf=min(candidates,key=lambda r:abs(r[0]-y))
        else:
            dark=crop.min(axis=2)<115
            xs=np.flatnonzero(dark.any(axis=0))
            if len(xs):crop=crop[:,max(0,xs.min()-3):min(crop.shape[1],xs.max()+4)]
            text,conf=recognize(engine,crop)
        original_text=text
        if field=='fecha' and parse_date(text) is None:
            retry=image[round((y-6)*scale):round((y+6)*scale),round(x0*scale):round(x1*scale),:3]
            candidate,candidate_conf=recognize(engine,retry)
            if parse_date(candidate) is not None:text,conf=candidate,candidate_conf
        rows.append({'y':round(y,3),'top':start+a/scale,'bottom':start+b/scale,'text':text,'confidence':conf})
        if text!=original_text:rows[-1]['retry_original_text']=original_text
    return rows
def parse_date(text):
    norm=compact(text).replace('O','0')
    m=re.search(r'(\d{2})[-./]?([A-Z0]{3})[-./]?(\d{4})',norm)
    if not m:return None
    mon=m[2].replace('0','O')
    if mon not in MONTHS:return None
    try:return pd.Timestamp(int(m[3]),MONTHS[mon],int(m[1]))
    except ValueError:return None
def table_bounds(image,page_number):
    import numpy as np
    # Header background is gray across all six columns, unlike the short opening strip.
    s=3;x=round(560*s);gray=image[:,:,0]
    lo,hi=(240,355) if page_number==2 else (125,165)
    runs=bands((np.abs(gray[lo*s:hi*s,x].astype('int16')-196)<8))
    found=[(lo+a/s,lo+b/s) for a,b in runs if 9<=(b-a)/s<=25]
    if not found:return None
    start=found[0][1]+.5
    # A yellow closing balance ends this table. Exclude the TOTAL immediately above it.
    stripe=image[round(start*s):735*s,round(550*s),:3].astype('int16')
    yellow=(stripe[:,0]>230)&(stripe[:,1]>170)&(stripe[:,1]<230)&(stripe[:,2]<140)
    ys=np.flatnonzero(yellow)
    end=start+ys[0]/s-12 if len(ys) else 735
    return start,end
def scan_document(pdf_bytes,cache=None,progress=None):
    if str(LIBS) not in sys.path:sys.path.insert(0,str(LIBS))
    import numpy as np
    from rapidocr_onnxruntime import RapidOCR
    engine=RapidOCR(det_limit_side_len=2400,det_limit_type='max',intra_op_num_threads=4,inter_op_num_threads=1)
    doc=fitz.open(stream=pdf_bytes,filetype='pdf');pages=[]
    for i,page in enumerate(doc):
        if i==0:continue
        target=Path(cache)/f'page_{i+1}.json' if cache else None
        if target and target.exists():
            pages.append(json.loads(target.read_text(encoding='utf-8')));continue
        if progress:progress(f'Hoja {i+1}/{len(doc)}: separando columnas y renglones')
        pix=page.get_pixmap(matrix=fitz.Matrix(3,3))
        image=np.frombuffer(pix.samples,np.uint8).reshape(pix.height,pix.width,pix.n)
        bounds=table_bounds(image,i+1)
        if bounds is None:
            data={'page':i+1,'columns':{},'table':None}
        else:
            start,end=bounds
            cols={}
            for label,x0,x1 in [('fecha',30.5,79.5),('folio',80.3,113.05),('descripcion',113.65,361.7),('deposito',363,429.9),('retiro',430.7,497.05),('saldo',498,581.6)]:
                cols[label]=read_cells(engine,image,x0,x1,start,end,field=label)
            data={'page':i+1,'table':[start,end],'columns':cols}
        if target:target.parent.mkdir(parents=True,exist_ok=True);target.write_text(json.dumps(data,ensure_ascii=False),encoding='utf-8')
        pages.append(data)
    return pages,read_summary(doc,engine)
def read_summary(doc,engine):
    import numpy as np
    # First-page right summary has opening, deposits, withdrawals and final in that order.
    p=doc[0];pix=p.get_pixmap(matrix=fitz.Matrix(3,3));a=np.frombuffer(pix.samples,np.uint8).reshape(pix.height,pix.width,pix.n)
    cells=read_cells(engine,a,490,581,435,490)
    values=[money(r['text']) for r in cells]
    if len(values)<4 or any(v is None for v in values[:4]):return {'readable':False,'raw':cells}
    return dict(readable=True,opening=values[0],deposits=values[1],withdrawals=values[2],final=values[3],raw=cells)
def nearest(rows,y,tolerance=3):
    group=[r for r in rows if abs(r['y']-y)<tolerance]
    return min(group,key=lambda r:abs(r['y']-y)) if group else None
def parse_pages(pages,summary):
    movements=[];issues=[];physical_missing=[]
    for data in pages:
        cols=data.get('columns',{})
        if not cols:continue
        anchors=[]
        for r in cols.get('fecha',[]):
            dt=parse_date(r['text'])
            if dt is None:
                physical_missing.append({'Hoja':data['page'],'Y':r['y'],'Texto':r['text'],'Problema':'Fecha ilegible: ancla detectada en imagen, no interpretada'})
                continue
            anchors.append((r,dt))
        leading=[r for r in cols.get('descripcion',[]) if not anchors or r['y']<anchors[0][0]['y']-3]
        if leading and movements:
            prior=movements[-1]
            detail=json.loads(prior['_LINEAS_DOCUMENTO'])
            origins=json.loads(prior['_LINEAS_ORIGEN'])
            detail.extend(restore_known_spaces(r['text']) for r in leading)
            origins.extend({'page':data['page'],'y':r['y'],'raw':r['text']} for r in leading)
            prior.update(DESCRIPCION='\n'.join(detail),_LINEAS_DOCUMENTO=json.dumps(detail,ensure_ascii=False),_LINEAS_ORIGEN=json.dumps(origins,ensure_ascii=False),LINEAS_FISICAS=len(detail))
        for j,(a,date) in enumerate(anchors):
            stop=anchors[j+1][0]['y']-3 if j+1<len(anchors) else data['table'][1]
            desc=[r for r in cols['descripcion'] if a['y']-3<=r['y']<stop]
            detail=[restore_known_spaces(r['text']) for r in desc]
            problems=[]
            if not desc or abs(desc[0]['y']-a['y'])>3:problems.append('Falta concepto en primera línea')
            folio=nearest(cols['folio'],a['y'])
            folio_text=compact(folio['text']) if folio else ''
            if not re.fullmatch(r'\d{7}',folio_text):problems.append('Folio no leído como siete dígitos')
            amounts={};confidence=[a['confidence']]+[r['confidence'] for r in desc]
            for label in ['deposito','retiro','saldo']:
                cell=nearest(cols[label],a['y'])
                amounts[label]=money(cell['text']) if cell else (None if label=='saldo' else 0.)
                if cell:confidence.append(cell['confidence'])
                if amounts[label] is None:problems.append('Importe ilegible en '+label)
            if min(confidence or [0])<.95:problems.append('Texto o importe de confianza baja')
            if any(re.search(r'[A-Z]{14}',s) for s in detail):problems.append('Revisar espacios y referencias del concepto')
            movements.append({'FECHA':date,'FOLIO':folio_text,'DESCRIPCION':'\n'.join(detail),'DEPOSITO':amounts['deposito'],'RETIRO':amounts['retiro'],'SALDO':amounts['saldo'],'_LINEAS_DOCUMENTO':json.dumps(detail,ensure_ascii=False),'_LINEAS_ORIGEN':json.dumps([{'page':data['page'],'y':r['y'],'raw':r['text']} for r in desc],ensure_ascii=False),'ORIGEN':'ORIGINAL','HOJA_ORIGINAL':data['page'],'Y_ORIGINAL':a['y'],'LINEAS_FISICAS':len(detail),'CONFIANZA_OCR':min(confidence or [0]),'REVISION_OCR':'; '.join(problems),'ID_MOVIMIENTO':len(movements)+1})
    df=pd.DataFrame(movements)
    if df.empty:raise ValueError('No se detectaron movimientos con fechas válidas.')
    previous=summary.get('opening')
    for i,row in df.iterrows():
        if previous is not None and pd.notna(row.SALDO) and pd.notna(row.DEPOSITO) and pd.notna(row.RETIRO):
            delta=round(previous+row.DEPOSITO-row.RETIRO-row.SALDO,2)
            if abs(delta)>.015:
                df.at[i,'REVISION_OCR']+=(('; ' if row.REVISION_OCR else '')+f'Saldo no concilia: diferencia {delta:.2f}')
                issues.append({'Movimiento':int(row.ID_MOVIMIENTO),'Hoja':int(row.HOJA_ORIGINAL),'Revisión':f'Saldo no concilia: {delta:.2f}'})
        previous=row.SALDO if pd.notna(row.SALDO) else None
    deposits=round(float(df.DEPOSITO.sum()),2);withdrawals=round(float(df.RETIRO.sum()),2)
    failures=[]
    if not summary.get('readable'):failures.append('No se pudo leer el resumen de la carátula')
    else:
        for label,value in [('deposits',deposits),('withdrawals',withdrawals),('final',float(df.SALDO.iloc[-1]))]:
            if abs(value-summary[label])>.015:failures.append(f'{label}: movimientos {value:.2f}, carátula {summary[label]:.2f}')
    if df[['DEPOSITO','RETIRO','SALDO']].isna().any().any():failures.append('Hay importes sin interpretar: no se deben tratar como cero')
    if physical_missing:failures.append(f'{len(physical_missing)} fechas físicas sin interpretar')
    if issues:failures.append(f'{len(issues)} saldos intermedios sin conciliar')
    report={'movement_count':len(df),'physical_lines':int(df.LINEAS_FISICAS.sum()),'deposits':deposits,'withdrawals':withdrawals,'opening':summary.get('opening'),'final':float(df.SALDO.iloc[-1]),'summary':summary,'financial_reconciled':not failures,'failures':failures,'unparsed_dates':physical_missing,'version':'select-columns-physical-v4'}
    for _,row in df.iterrows():
        if row.REVISION_OCR:issues.append({'Movimiento':int(row.ID_MOVIMIENTO),'Hoja':int(row.HOJA_ORIGINAL),'Revisión':row.REVISION_OCR})
    df.attrs.update(ocr_review=issues,ocr_verified=False,financial_reconciled=not failures,reconciliation=report,bank='SANTANDER SELECT')
    # Opening is an explicit non-operation; it is not counted as income or withdrawal.
    if summary.get('readable'):
        first={c:None for c in df.columns};first.update(FECHA=df.FECHA.min().replace(day=1),FOLIO='',DESCRIPCION='SALDO ANTERIOR',DEPOSITO=0.,RETIRO=0.,SALDO=summary['opening'],_LINEAS_DOCUMENTO='["SALDO ANTERIOR"]',ORIGEN='ORIGINAL',HOJA_ORIGINAL=2,LINEAS_FISICAS=1,ID_MOVIMIENTO=0,REVISION_OCR='',CONFIANZA_OCR=1.)
        attrs=df.attrs.copy();df=pd.concat([pd.DataFrame([first]),df],ignore_index=True);df.attrs=attrs
    return df
def extract_select(pdf_bytes):
    pages,summary=scan_document(pdf_bytes)
    return parse_pages(pages,summary)

def is_select_scan(pdf_bytes):
    """Identify an image-only Select reference by its printed logo, not by filename."""
    if str(LIBS) not in sys.path:sys.path.insert(0,str(LIBS))
    import numpy as np
    from rapidocr_onnxruntime import RapidOCR
    doc=fitz.open(stream=pdf_bytes,filetype='pdf')
    if not len(doc):return False
    pix=doc[0].get_pixmap(matrix=fitz.Matrix(3,3),clip=fitz.Rect(25,5,300,62))
    image=np.frombuffer(pix.samples,np.uint8).reshape(pix.height,pix.width,pix.n)
    engine=RapidOCR(det_limit_side_len=1200,det_limit_type='max',intra_op_num_threads=4,inter_op_num_threads=1)
    result,_=engine(image)
    return 'SANTANDERSELECT' in compact(' '.join(r[1] for r in result or []))
