"""Plantilla de movimientos Select calibrada a estados raster de 2026."""
from pathlib import Path
import re,hashlib
import fitz
import pandas as pd

MONTHS=['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC']
X=[31.,80.,113.4,362.4,430.2,497.4,580.4]
SIZE=8.5
PITCH=9.2

def clean(v,default=''):
    if v is None:return default
    try:
        if pd.isna(v):return default
    except (ValueError,TypeError):pass
    return re.sub(r'\s+',' ',str(v)).strip() or default

def fields(text):
    matches=list(re.finditer(r'\b(ORDENANTE|BENEFICIARIO|CUENTA|RASTREO|REF|CONCEPTO|HORA|RFC)\s*:',text,re.I))
    return {m[1].upper():text[m.end():(matches[i+1].start() if i+1<len(matches) else len(text))].strip() for i,m in enumerate(matches)}

def movement_lines(row):
    from excel_layouts import decode_physical_lines
    text=clean(row.get('DESCRIPCION'))
    if clean(row.get('ORIGEN')).upper()=='ORIGINAL' and row.get('_LINEAS_DOCUMENTO') is not None:
        return decode_physical_lines(row,'SANTANDER')
    f=fields(text)
    if not any(k in text.upper() for k in ('SPEI','TRANSFERENCIA')):
        return decode_physical_lines(row,'SANTANDER')
    incoming=float(row.get('DEPOSITO',0) or 0)>0
    bank=clean(row.get('BANCO_TRANSFERENCIA'),'BBVA MEXICO')
    name=clean(row.get('NOMBRE_CONTRAPARTE')) or f.get('ORDENANTE' if incoming else 'BENEFICIARIO','PERSONA FICTICIA')
    account=f.get('CUENTA','000000000000000000')
    trace=f.get('RASTREO','REFERENCIA NO DISPONIBLE')
    ref=f.get('REF','0000000')[-7:]
    concept=clean(row.get('CONCEPTO_SOLICITADO')) or f.get('CONCEPTO','TRANSFERENCIA')
    hour=f.get('HORA','12:00:00')
    if incoming:
        return [f'ABONO TRANSFERENCIA SPEI HORA {hour}',f'RECIBIDO DE {bank}',f'DE LA CUENTA {account}',f'DEL CLIENTE {name}',f'CLAVE DE RASTREO {trace}',f'REF {ref}',f'CONCEPTO {concept}',f'RFC {f.get("RFC","XAXX010101000")}']
    return [f'PAGO TRANSFERENCIA SPEI HORA {hour}',f'ENVIADO A {bank}',f'A LA CUENTA {account}',f'AL CLIENTE {name} ( 1 )','(1 ) DATO NO VERIFICADO POR ESTA INSTITUCION',f'CLAVE DE RASTREO {trace}',f'REF {ref}',f'CONCEPTO {concept}']

def wrap(text,font,width):
    rows=[];line=''
    for word in clean(text).split():
        candidate=(line+' '+word).strip()
        if line and font.text_length(candidate,fontsize=SIZE)>width:
            rows.append(line);line=''
        if font.text_length(word,fontsize=SIZE)>width:
            for char in word:
                if line and font.text_length(line+char,fontsize=SIZE)>width:rows.append(line);line=''
                line+=char
        else:line=(line+' '+word).strip()
    return rows+[line] if line else rows or ['']

def generate_select_pdf(df,identity,start,end,font_path=None):
    font_path=Path(font_path or Path(__file__).with_name('CLC_Movimientos_Select_Regular.ttf'))
    font=fitz.Font(fontfile=str(font_path));doc=fitz.open();page=None;y=0
    def text(x,base,value,size=SIZE):
        page.insert_text((x,base),str(value),fontfile=str(font_path),fontname='CLCSelect',fontsize=size)
    def amount(right,base,value):
        value=f'{float(value):,.2f}'
        text(right-font.text_length(value,fontsize=SIZE),base,value)
    def new_page():
        nonlocal page,y
        page=doc.new_page(width=612,height=792)
        text(31,20,'MOVIMIENTOS - SANTANDER SELECT',10)
        text(31,31,'MATERIAL DIDACTICO - DATOS FICTICIOS',7)
        top=40
        if len(doc)==1:
            opening=df[df.DESCRIPCION.astype(str).str.upper().str.strip()=='SALDO ANTERIOR']
            value=float(opening.SALDO.iloc[0]) if len(opening) else 0.
            page.draw_rect((X[0],top,X[4],top+14),color=None,fill=(.65,.65,.65))
            text(35,top+10,'SALDO FINAL DEL PERIODO ANTERIOR:',9)
            amount(X[4]-3,top+10,value);top+=14
        page.draw_rect((X[0],top,X[-1],top+13),color=None,fill=(.76,.76,.76))
        for left,right,label in zip(X,X[1:],['FECHA','FOLIO','DESCRIPCION','DEPOSITO','RETIRO','SALDO']):
            xx=left+5 if label=='DESCRIPCION' else (left+right-font.text_length(label,fontsize=SIZE))/2
            text(xx,top+9.5,label)
        y=top+13
    work=df[~df.DESCRIPCION.astype(str).str.upper().str.strip().eq('SALDO ANTERIOR')]
    for index,(_,row) in enumerate(work.iterrows()):
        raw=movement_lines(row);lines=[piece for line in raw for piece in wrap(line,font,X[3]-X[2]-3)]
        if page is None:new_page()
        offset=0
        while offset<len(lines):
            capacity=int((752-y)/PITCH)
            if capacity<1:new_page();continue
            count=min(capacity,len(lines)-offset);height=count*PITCH
            if index%2==0:page.draw_rect((X[0],y,X[-1],y+height),color=None,fill=(.92,.92,.92))
            for xx in X[1:-1]:page.draw_line((xx,y),(xx,y+height),color=(.68,.68,.68),width=.3)
            for n,line in enumerate(lines[offset:offset+count]):text(X[2],y+7.1+n*PITCH,line)
            if offset==0:
                date=pd.Timestamp(row.FECHA);text(X[0]+1,y+7.1,f'{date.day:02d}-{MONTHS[date.month-1]}-{date.year}')
                f=fields(clean(row.DESCRIPCION));folio=clean(row.get('FOLIO')) or f.get('REF','0000000')[-7:]
                text(X[1]+1,y+7.1,folio)
                if float(row.get('DEPOSITO',0) or 0):amount(X[4]-2,y+7.1,row.DEPOSITO)
                if float(row.get('RETIRO',0) or 0):amount(X[5]-2,y+7.1,row.RETIRO)
                if not pd.isna(row.SALDO):amount(X[6]-2,y+7.1,row.SALDO)
            y+=height;offset+=count
    if page is None:new_page()
    for i,p in enumerate(doc,1):p.insert_text((31,775),f'MATERIAL DIDACTICO - PAGINA {i} DE {len(doc)}',fontname='helv',fontsize=7)
    doc.subset_fonts()
    return doc.tobytes(garbage=4,deflate=True)
