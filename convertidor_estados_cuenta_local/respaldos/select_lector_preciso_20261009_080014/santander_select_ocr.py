"""Lector local de tablas Select escaneadas. Resultado sujeto a revisión."""
from pathlib import Path
import sys,re,json
import fitz
import pandas as pd

LIBS=Path(r'C:\Users\clopez\Documents\ChatGPT\LOCAL CLC\tmp\santander_v2\libs')
MONTHS={name:i+1 for i,name in enumerate(['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'])}
DATE=re.compile(r'^(\d{2})-([A-Z]{3})-(\d{4})(\d{7})(.*)$')

def restore_known_spaces(text):
    # Sólo separaciones de etiquetas conocidas; conservar intactos los datos.
    for compact,spaced in [('ADMINISTRACIONRENTA','ADMINISTRACION RENTA '),('IVAPORCOMISION','IVA POR COMISION '),('ABONOTRANSFERENCIASPEIHORA','ABONO TRANSFERENCIA SPEI HORA '),('PAGOTRANSFERENCIASPEIHORA','PAGO TRANSFERENCIA SPEI HORA '),('RECIBIDODE','RECIBIDO DE '),('ENVIADOA','ENVIADO A '),('DELACUENTA','DE LA CUENTA '),('ALACUENTA','A LA CUENTA '),('DELCLIENTE','DEL CLIENTE '),('ALCLIENTE','AL CLIENTE '),('CLAVEDERASTREO','CLAVE DE RASTREO '),('CONCEPTO','CONCEPTO '),('REF','REF '),('RFC','RFC ')]:
        if text.startswith(compact):return spaced+text[len(compact):]
    return text

def money(text):
    value=text.replace(' ','')
    if not re.fullmatch(r'-?\d[\d,.]*[.,]\d{2}',value):return None
    return float(re.sub(r'[.,]','',value[:-3])+'.'+value[-2:])

def parse_cached(pages):
    rows=[];issues=[];current=None;previous=None
    for data in pages:
        origin=data['origin'];scale=data['scale'];lines=[]
        for box,text,conf in data['lines'] or []:
            x=sum(p[0] for p in box)/4/scale+origin[0];y=sum(p[1] for p in box)/4/scale+origin[1]
            lines.append((y,x,text,conf))
        lines.sort()
        anchors=[]
        for y,x,text,conf in lines:
            match=DATE.match(re.sub(r'\s','',text.upper()))
            if match and x<362 and match[2] in MONTHS:anchors.append((y,match,conf))
        for number,(y,match,confidence) in enumerate(anchors):
            next_y=anchors[number+1][0] if number+1<len(anchors) else 735
            group=[l for l in lines if y-3<=l[0]<next_y-3]
            numeric=[l for l in group if l[1]>362 and abs(l[0]-y)<4]
            amounts=[0.,0.,None];problems=[]
            for yy,x,text,conf in numeric:
                idx=0 if x<430.2 else 1 if x<497.4 else 2
                value=money(text)
                if value is None:problems.append('Importe ilegible: '+text)
                else:amounts[idx]=value
                if conf<.98:problems.append('Importe con confianza baja: '+text)
            detail=[match[5]]+[l[2] for l in group if l[1]<362 and l[0]>y+3]
            detail=[restore_known_spaces(s) for s in detail if s.strip()]
            if any(l[3]<.97 for l in group):problems.append('Texto o referencia con confianza baja')
            if any(re.search(r'[A-Z]{12}',s) for s in detail):problems.append('Revisar espacios del concepto contra la imagen')
            if amounts[2] is None:problems.append('Falta saldo impreso')
            if previous is not None and amounts[2] is not None and abs(previous+amounts[0]-amounts[1]-amounts[2])>.015:problems.append('El saldo no concilia con el movimiento anterior')
            if previous is None:problems.append('Saldo inicial pendiente de verificar en el original')
            previous=amounts[2]
            row={'FECHA':pd.Timestamp(int(match[3]),MONTHS[match[2]],int(match[1])),'FOLIO':match[4],'DESCRIPCION':'\n'.join(detail),'DEPOSITO':amounts[0],'RETIRO':amounts[1],'SALDO':amounts[2],'_LINEAS_DOCUMENTO':json.dumps(detail,ensure_ascii=False),'ORIGEN':'ORIGINAL','HOJA_ORIGINAL':data['page'],'REVISION_OCR':'; '.join(problems) or 'Revisar referencias y concepto','CONFIANZA_OCR':min(l[3] for l in group)}
            rows.append(row)
            issues.append({'Movimiento':len(rows),'Hoja':data['page'],'Revisión':row['REVISION_OCR']})
    df=pd.DataFrame(rows)
    df.attrs['ocr_review']=issues;df.attrs['ocr_verified']=False;df.attrs['bank']='SANTANDER SELECT'
    return df

def extract_select(pdf_bytes):
    if str(LIBS) not in sys.path:sys.path.insert(0,str(LIBS))
    try:
        import numpy as np
        from rapidocr_onnxruntime import RapidOCR
    except ImportError as error:raise RuntimeError('No está disponible el lector de imágenes Santander Select') from error
    engine=RapidOCR(det_limit_side_len=2400,intra_op_num_threads=4,inter_op_num_threads=1)
    doc=fitz.open(stream=pdf_bytes,filetype='pdf');pages=[]
    for index,page in enumerate(doc):
        if index==0:continue
        origin=[30,28];clip=fitz.Rect(30,28,581,735)
        pix=page.get_pixmap(matrix=fitz.Matrix(3,3),clip=clip)
        pixels=np.frombuffer(pix.samples,np.uint8).reshape(pix.height,pix.width,pix.n)
        result,_=engine(pixels)
        pages.append({'page':index+1,'origin':origin,'scale':3,'lines':result})
    result=parse_cached(pages)
    result.attrs['ocr_review'].insert(0,{'Movimiento':'General','Hoja':'Todas','Revisión':'Extracción preliminar: comprobar cantidad de operaciones, saldo inicial y final contra el PDF. Puede omitir filas o confundir cifras; no usar como datos confirmados.'})
    return result
