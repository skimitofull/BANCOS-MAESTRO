"""Editable, visibly didactic Select documents. Original assets are optional references."""
from pathlib import Path
from io import BytesIO
from decimal import Decimal, ROUND_HALF_UP
import json, re, math
import fitz
import pandas as pd
from santander_select import movement_lines, MONTHS

ROOT=Path(__file__).parent
RED=(234/255,0,0); GRAY=(196/255,)*3; LIGHT=(235/255,)*3
YELLOW=(1,213/255,91/255)
COLS=[30,80,113.4,362.4,430.2,497.4,582]
def cash(v):
    return Decimal(str(v)).quantize(Decimal('.01'),rounding=ROUND_HALF_UP)
def parse_select_date(value):
    if isinstance(value,str):
        value=value.strip()
        if re.match(r'^\d{4}-\d{2}-\d{2}',value):return pd.Timestamp(value)
        for es,en in [('ENE','JAN'),('ABR','APR'),('AGO','AUG'),('DIC','DEC')]:
            value=re.sub(r'\b'+es+r'\b',en,value,flags=re.I)
    return pd.to_datetime(value,errors='raise',dayfirst=True)
def normalized(df,start,end,opening=None,recalculate=True):
    out=df.copy(); out.columns=[str(c).strip().upper() for c in out.columns]
    for c in ['FECHA','DESCRIPCION','DEPOSITO','RETIRO']:
        if c not in out: raise ValueError('Falta la columna '+c)
    if 'SALDO' not in out:out['SALDO']=float('nan')
    for c in ['DEPOSITO','RETIRO']:
        out[c]=pd.to_numeric(out[c],errors='raise').fillna(0)
        if (out[c]<0).any():raise ValueError(c+': usa importes positivos en su columna correspondiente.')
    if ((out.DEPOSITO>0)&(out.RETIRO>0)).any():raise ValueError('Un movimiento no puede tener depósito y retiro simultáneos.')
    mask=out.DESCRIPCION.astype(str).str.strip().str.upper().eq('SALDO ANTERIOR')
    prior=out[mask];work=out[~mask].copy()
    if work.empty:raise ValueError('El Excel no contiene movimientos.')
    work['FECHA']=work.FECHA.map(parse_select_date)
    if not work.FECHA.is_monotonic_increasing:raise ValueError('Ordena las fechas de menor a mayor; se conserva el orden de operaciones del Excel.')
    start,end=pd.Timestamp(start).normalize(),pd.Timestamp(end).normalize()
    if start>end or (work.FECHA.dt.normalize()<start).any() or (work.FECHA.dt.normalize()>end).any():raise ValueError('Las fechas deben quedar dentro del periodo elegido.')
    if opening is None:
        if len(prior):opening=pd.to_numeric(prior.SALDO,errors='raise').iloc[0]
        elif pd.notna(work.SALDO.iloc[0]):opening=float(work.SALDO.iloc[0])-work.DEPOSITO.iloc[0]+work.RETIRO.iloc[0]
        else:raise ValueError('Indica saldo inicial o agrega una fila SALDO ANTERIOR con su saldo.')
    balance=cash(opening); initial=balance
    for idx,row in work.iterrows():
        balance+=cash(row.DEPOSITO)-cash(row.RETIRO)
        if not recalculate and pd.notna(row.SALDO) and abs(cash(row.SALDO)-balance)>Decimal('.01'):
            raise ValueError(f'El saldo no concilia en la fila {idx}. Activa recalcular saldos o corrige el Excel.')
        work.at[idx,'SALDO']=float(balance)
    daily={};running=initial
    for day in pd.date_range(start,end):
        rows=work[work.FECHA.dt.normalize()==day]
        if len(rows):running=cash(rows.SALDO.iloc[-1])
        daily[day]=running
    return work,{'initial':float(initial),'deposits':float(sum(map(cash,work.DEPOSITO))),
        'withdrawals':float(sum(map(cash,work.RETIRO))),'final':float(balance),
        'average':float(sum(daily.values())/len(daily)),'days':len(daily),'count':len(work)}

class Layout:
    def __init__(self,fields,start,end,finish,reference=None):
        self.doc=fitz.open();self.fields=fields;self.start=pd.Timestamp(start);self.end=pd.Timestamp(end)
        scan=finish=='ESCANEO_EDITABLE'
        names={'r':'CLC_Select_Escaneo_Editable.ttf' if scan else 'CLC_Movimientos_Select_Regular.ttf'}
        for k,s in [('b','Bold'),('n','NumerosBold'),('i','BoldItalic')]:
            names[k]=f'CLC_Select_{"Escaneo_" if scan else ""}{s}.ttf'
        self.paths={k:ROOT/v for k,v in names.items()};self.fonts={k:fitz.Font(fontfile=str(v)) for k,v in self.paths.items()}
        self.reference=fitz.open(stream=reference,filetype='pdf') if reference else None
    def text(self,x,y,s,size=8.5,k='r',color=(0,0,0),width=None):
        s=str(s)
        if width is not None:size=min(size,size*width/max(self.fonts[k].text_length(s,fontsize=size),.01))
        if size<5.5:raise ValueError('Texto demasiado largo para el campo: '+s[:70])
        self.p.insert_text((x,y),s,fontfile=str(self.paths[k]),fontname='CLC'+k,fontsize=size,color=color)
    def amount(self,right,y,value,size=8.5,k='r',prefix=''):
        s=prefix+f'{value:,.2f}';self.text(right-self.fonts[k].text_length(s,fontsize=size),y,s,size,k)
    def rect(self,x,y,r,b,color):self.p.draw_rect((x,y,r,b),color=None,fill=color)
    def date(self,dt):return f'{dt.day:02d}-{MONTHS[dt.month-1]}-{dt.year}'
    def page(self,cover=False):
        self.p=self.doc.new_page(width=612,height=792)
        if self.reference:
            src=self.reference[0]
            # Only fixed branding area and footer; no account identifiers or barcode.
            for rect in [fitz.Rect(30,12,255,60),fitz.Rect(30,752,582,792)]:
                pix=src.get_pixmap(matrix=fitz.Matrix(2,2),clip=rect)
                self.p.insert_image(rect,stream=pix.tobytes('png'))
            self.rect(30,766,582,792,(1,1,1))
        else:self.text(30,36,'PLANTILLA SELECT PARA CLASE',12,'b',RED)
        self.text(30,12,'EJEMPLO DIDACTICO - DATOS MODIFICADOS - SIN VALIDEZ BANCARIA',7,'b',RED)
        self.text(428,60,'ESTADO DE CUENTA',18,'i',(54/255,58/255,61/255),width=154)
        if not cover:
            self.text(30,111,self.fields.get('holder','PERSONA FICTICIA'),8.5,'b',width=320)
            self.text(385,111,'CLIENTE: '+self.fields.get('client','00000000'),8.5,'b',width=197)
            self.text(365,123,'PERIODO DEL '+self.date(self.start)+' AL '+self.date(self.end),8.5,'b',width=217)
            self.p.draw_line((30,128),(582,128),color=(.6,.6,.6),width=.4)
    def title(self,y,s):self.text(55,y,s,13,'b',RED,width=527)
    def product(self,y,secondary=False):
        self.rect(30,y,582,y+14,GRAY)
        account=self.fields.get('secondary_account','0000000001') if secondary else self.fields.get('account','0000000000')
        self.text(34,y+10,('DINERO CRECIENTE' if secondary else 'CUENTA DE CHEQUES')+' - '+account,8.5,'b',width=265)
        self.text(301,y+10,'CLABE '+self.fields.get('clabe','000000000000000000'),8.5,'b',RED,width=275)
    def wrap(self,s,width,size=8.5,k='r'):
        result=[]
        for raw in str(s).splitlines() or ['']:
            line=''
            for token in re.findall(r'\S+\s*',raw):
                for ch in token:
                    if line and self.fonts[k].text_length(line+ch,fontsize=size)>width:
                        result.append(line.rstrip());line=''
                    line+=ch
            result.append(line.rstrip())
        return result
    def indicators(self,y,t,secondary=False):
        vals=[('Saldo promedio',0 if secondary else t['average']),('Tasa bruta de interes anual',0),('Dias del periodo',t['days'])]
        for j,(label,val) in enumerate(vals):
            self.text(30,y+j*11,label);self.amount(267,y+j*11,val)
        for j,(label,key) in enumerate([('Saldo inicial','initial'),('+ Depositos','deposits'),('- Retiros','withdrawals'),('= Saldo final','final')]):
            self.text(365,y+j*11,label,8.5,'b' if j==3 else 'r');self.amount(580,y+j*11,0 if secondary else t[key],8.5,'n' if j==3 else 'r')
    def finish(self):
        for i,p in enumerate(self.doc,1):
            p.insert_text((30,778),f'EJEMPLO DIDACTICO - SIN VALIDEZ BANCARIA - {i}/{len(self.doc)}',fontsize=8,color=RED)
        self.doc.set_metadata({'title':'Ejercicio Select editable con datos didácticos','subject':'Texto vectorial; fuentes completas incorporadas. Las imágenes fijas conservadas son referencias.'})
        return self.doc.tobytes(garbage=4,deflate=True)

def cover(l,t):
    l.page(True);f=l.fields
    l.text(51,78,f.get('holder','PERSONA FICTICIA'),10,'b',width=300)
    for j,line in enumerate(l.wrap(f.get('address','DOMICILIO FICTICIO'),290)[:4]):l.text(51,91+j*10,line)
    for j,(label,key,default) in enumerate([('CLIENTE','client','00000000'),('RFC','rfc','XAXX010101000'),('MONEDA','currency','MXN'),('SUCURSAL','branch','DIDACTICA'),('TELEFONO','phone','0000000000')]):
        l.text(382,78+j*11,label+': '+str(f.get(key,default)),8.5,'b',width=200)
    l.text(382,133,'PERIODO: '+l.date(l.start)+' AL '+l.date(l.end),8.5,'b',width=200)
    l.text(382,145,'CORTE: '+l.date(l.end),8.5,'b')
    l.title(195,'Resumen informativo.');l.text(30,214,'Resumen intereses y comisiones')
    xs=[30,159,225,278,330,384,441,509,582];l.rect(30,218,582,242,GRAY)
    labels=['PRODUCTO','NUMERO DE\nCUENTA','INTERESES\nBRUTOS','ISR\nRETENIDO','INTERESES\nNETOS','COMISIONES\nCOBRADAS','GAT\nNOMINAL','GAT\nREAL']
    for j,label in enumerate(labels):
        for k,line in enumerate(label.splitlines()):l.text(xs[j]+3,228+k*8,line,7.5,'b',width=xs[j+1]-xs[j]-6)
    for j in range(2):
        y=252+j*12;l.text(32,y,'CHEQUES' if not j else 'DINERO CRECIENTE',8,width=124)
        l.text(162,y,f.get('account' if not j else 'secondary_account','0000000000'),8,width=61)
        for col,key in zip(range(2,8),['interest','tax','net_interest','commissions','gat','gat_real']):
            l.amount(xs[col+1]-4,y,float(f.get(key,0)) if j==0 else 0,8)
    l.text(30,294,'Resumen saldos.');l.rect(30,299,582,326,GRAY)
    for x,s in [(32,'PRODUCTO'),(162,'CUENTA'),(247,'MES ANTERIOR'),(418,'MES ACTUAL')]:l.text(x,310,s,8,'b')
    for x,s in [(247,'MONTO'),(324,'% DISTRIBUCION'),(418,'MONTO'),(497,'% DISTRIBUCION')]:l.text(x,322,s,7.5,'b')
    for j in range(2):
        y=337+j*12;l.text(32,y,'CHEQUES' if not j else 'DINERO CRECIENTE',8,width=125)
        l.text(162,y,f.get('account' if not j else 'secondary_account','0000000000'),8,width=62)
        l.amount(319,y,t['initial'] if not j else 0,8);l.text(337,y,'100.00%' if not j else '0.00%',8)
        l.amount(491,y,t['final'] if not j else 0,8);l.text(524,y,'100.00%' if not j else '0.00%',8)
    l.rect(242,353,582,369,GRAY);l.text(32,365,'TOTAL',10,'b');l.amount(319,365,t['initial'],9,'n');l.text(337,365,'100.00%',8);l.amount(491,365,t['final'],9,'n');l.text(524,365,'100.00%',8)
    l.title(395,'Cuenta de cheques.');l.product(402);l.text(30,427,'SUCURSAL: '+f.get('branch','DIDACTICA'),8.5,'b',width=552)
    l.indicators(445,t);l.text(30,484,'Importes calculados desde el Excel del ejercicio.',7)
    l.title(522,'Grafica cuenta de cheques.');l.text(30,545,'CUENTA '+f.get('account','0000000000'),8.5,'b')
    l.text(30,559,'Saldo inicial');l.amount(212,559,t['initial'])
    values=[t['deposits'],t['withdrawals'],float(f.get('commissions',0)),t['final']]
    total=sum(abs(v) for v in values);angle=-math.pi/2
    colors=[(235/255,70/255,66/255),(251/255,223/255,220/255),(.8,.8,.8),(.65,.65,.65)]
    for val,col in zip(values,colors):
        delta=2*math.pi*abs(val)/max(total,1)
        if delta:
            n=max(3,int(delta*30));points=[fitz.Point(385,578)]+[fitz.Point(385+50*math.cos(angle+delta*i/n),578+50*math.sin(angle+delta*i/n)) for i in range(n+1)]
            l.p.draw_polyline(points,color=None,fill=col,closePath=True)
        angle+=delta
    for j,(label,val,col) in enumerate(zip(['Depositos','Retiros','Comisiones','Saldo final'],values,colors)):
        y=544+j*23;l.rect(447,y-7,454,y,col);l.text(458,y,label,8);l.amount(580,y+10,val,8)
    if any(v<0 for v in values):l.text(240,650,'Grafica: proporciones de valores absolutos.',7)

def movement_pages(l,work,t,complete):
    l.page();
    if complete:
        l.title(149,'Dinero creciente.');l.product(165,True);l.indicators(197,t,True)
        l.title(259,'Detalle de movimientos cuenta de cheques.');l.text(30,277,'CUENTA '+l.fields.get('account','0000000000'),8.5,'b');top=280
    else:
        l.title(149,'Detalle de movimientos cuenta de cheques.');top=161
    def table(top,opening=False):
        if opening:
            l.rect(30,top,344,top+14,(163/255,)*3);l.text(32,top+10,'SALDO FINAL DEL PERIODO ANTERIOR:',8.5);l.amount(342,top+10,t['initial']);top+=14
        l.rect(30,top,582,top+14,GRAY)
        for x,s in zip(COLS,['FECHA','FOLIO','DESCRIPCION','DEPOSITO','RETIRO','SALDO']):l.text(x+2,top+10,s,8.5,'b')
        return top+14
    y=table(top,True)
    for j,(_,row) in enumerate(work.iterrows()):
        if str(row.get('ORIGEN','')).upper()=='ORIGINAL':
            try:
                old=json.loads(str(row.get('_LINEAS_DOCUMENTO','[]')))
            except (ValueError,TypeError):old=[]
            collapse=lambda v:re.sub(r'\s+',' ',str(v)).strip()
            if collapse(' '.join(old))!=collapse(row.DESCRIPCION):
                row=row.copy();row['_LINEAS_DOCUMENTO']=json.dumps(str(row.DESCRIPCION).splitlines(),ensure_ascii=False)
        lines=[line for raw in movement_lines(row) for line in l.wrap(raw,245)]
        offset=0
        if y+len(lines)*9.2>731 and len(lines)*9.2<586:
            l.page();y=table(131)
        while offset<len(lines):
            cap=int((731-y)/9.2)
            if cap<1:l.page();y=table(131);continue
            n=min(cap,len(lines)-offset);bottom=y+n*9.2
            if j%2==0:l.rect(30,y,582,bottom,LIGHT)
            for x in COLS[1:-1]:l.p.draw_line((x,y),(x,bottom),color=(.72,)*3,width=.3)
            for k,line in enumerate(lines[offset:offset+n]):l.text(113.4,y+7.1+k*9.2,line)
            if offset==0:
                l.text(31,y+7.1,l.date(row.FECHA),width=48)
                folio=str(row.get('FOLIO','')).replace('nan','')
                l.text(81,y+7.1,folio,width=31)
                if row.DEPOSITO:l.amount(428,y+7.1,row.DEPOSITO)
                if row.RETIRO:l.amount(495,y+7.1,row.RETIRO)
                l.amount(580,y+7.1,row.SALDO)
            y=bottom;offset+=n
    if y+30>731:l.page();y=table(131)
    l.text(115,y+10,'TOTAL',9,'b');l.amount(428,y+10,t['deposits'],9,'n');l.amount(495,y+10,t['withdrawals'],9,'n');y+=14
    l.rect(362,y,582,y+14,YELLOW);l.text(365,y+10,'SALDO FINAL DEL PERIODO:',8.5,'b',width=146);l.amount(580,y+10,t['final'],9,'n',prefix='$');y+=28
    if complete:
        if y+115>731:l.page();y=143
        l.title(y,'Detalle de movimientos Dinero Creciente.');y+=10;l.product(y,True);y+=20
        l.rect(30,y,344,y+14,(163/255,)*3);l.text(32,y+10,'SALDO FINAL DEL PERIODO ANTERIOR:');l.amount(342,y+10,0);y+=14
        y=table(y);l.text(115,y+10,'TOTAL',9,'b');l.amount(428,y+10,0);l.amount(495,y+10,0);y+=14
        l.rect(362,y,582,y+14,YELLOW);l.text(365,y+10,'SALDO FINAL DEL PERIODO:',8.5,'b',width=146);l.amount(580,y+10,0,9,'n',prefix='$');y+=32
    return y

def fiscal(l,y):
    if y+275>731:l.page();y=150
    l.title(y,'Informacion fiscal - ejemplo sin validez.');y+=22
    for x,s in [(30,'UUID DEL EJERCICIO'),(215,'CERTIFICADO EMISOR'),(400,'CERTIFICADO SAT')]:l.text(x,y,s,8,'b')
    for x in [30,215,400]:l.text(x,y+12,'FICTICIO - NO VERIFICABLE',8)
    y+=35
    items=[('FOLIO INTERNO','EJEMPLO-0001'),('FECHA DE EXPEDICION',l.date(l.end)),('FECHA DE CERTIFICACION','NO APLICA'),('REGIMEN EMISOR','DIDACTICO'),('LUGAR DE EXPEDICION','CIUDAD DE EJEMPLO'),('UNIDAD DE MEDIDA','SERVICIO'),('METODO DE PAGO','EJEMPLO'),('REGIMEN RECEPTOR',l.fields.get('regime','EJEMPLO')),('CP FISCAL RECEPTOR',l.fields.get('postal','00000')),('USO CFDI','SIN EFECTO FISCAL'),('TIPO DE COMPROBANTE','EJERCICIO')]
    for label,val in items:l.text(30,y,label,8,'b');l.text(220,y,val,8,width=290);y+=10
    for label in ['SELLO DEL EMISOR','SELLO DEL SAT','CADENA DE CERTIFICACION']:
        y+=13;l.text(30,y,label+' - NO VALIDO',8,'b');y+=11;l.text(30,y,'EJEMPLO DIDACTICO / SIN FIRMA NI CERTIFICACION / DATOS FICTICIOS',7.5)

def annexes(l,rows,preserve=False):
    for number in sorted({int(r['PAGINA']) for r in rows}):
        l.page()
        page_rows=[r for r in rows if int(r['PAGINA'])==number]
        if preserve and l.reference and 1<=number<=len(l.reference):
            top=max(132,min(float(r['Y'])-float(r['TAMANO'])*1.1 for r in page_rows))
            clip=fitz.Rect(30,top,582,731)
            pix=l.reference[number-1].get_pixmap(matrix=fitz.Matrix(2,2),clip=clip)
            l.p.insert_image(clip,stream=pix.tobytes('png'))
            # White overlays replace recognized text, leaving the fixed artwork visible.
            # OCR misses remain part of the reference image until reviewed/added.
            for r in page_rows:
                x,y,width,size=map(float,[r['X'],r['Y'],r['ANCHO'],r['TAMANO']])
                l.rect(max(30,x-.7),max(top,y-size*1.1),min(582,x+width),y+1.5,(1,1,1))
        for r in page_rows:
            txt=str(r['TEXTO']);x=float(r['X']);y=float(r['Y']);width=float(r['ANCHO']);size=float(r['TAMANO'])
            if not 30<=x<582 or not 132<=y<731:continue
            l.text(x,y,txt,max(5.5,min(size,16)),str(r.get('ESTILO','r')),RED if r.get('ROJO') else (0,0,0),width=width)

def generate_editable_document(df,identity,start,end):
    settings=identity.get('select_document',{});fields=identity.get('select_cover',{})
    work,t=normalized(df,start,end,settings.get('opening'),settings.get('recalculate',True))
    full=settings.get('scope','MOVIMIENTOS')=='COMPLETO'
    l=Layout(fields,start,end,identity.get('select_finish','ESCANEO_EDITABLE'),settings.get('reference'))
    if full:cover(l,t)
    y=movement_pages(l,work,t,full)
    if full:
        fiscal(l,y)
        rows=settings.get('annex_rows') or json.loads((ROOT/'select_fixed_text.json').read_text(encoding='utf-8'))
        annexes(l,rows,preserve=bool(settings.get('annex_rows')))
    return l.finish()

def extract_fixed_annexes(pdf_bytes,first_page,top=132):
    """OCR only explicitly selected fixed appendix pages; caller reviews resulting text."""
    from santander_select_ocr import LIBS
    import sys
    if str(LIBS) not in sys.path:sys.path.insert(0,str(LIBS))
    from rapidocr_onnxruntime import RapidOCR
    import numpy as np
    engine=RapidOCR(det_limit_side_len=2400,intra_op_num_threads=4,inter_op_num_threads=4)
    doc=fitz.open(stream=pdf_bytes,filetype='pdf');rows=[]
    for i in range(first_page-1,len(doc)):
        clip=fitz.Rect(30,top if i==first_page-1 else 132,582,731)
        pix=doc[i].get_pixmap(matrix=fitz.Matrix(3,3));img=np.frombuffer(pix.samples,np.uint8).reshape(pix.height,pix.width,pix.n)
        lines,_=engine(img)
        for box,text,conf in lines or []:
            x=min(v[0] for v in box)/3;y0=min(v[1] for v in box)/3;y=max(v[1] for v in box)/3
            right=max(v[0] for v in box)/3
            if not (clip.x0<=x<clip.x1 and clip.y0<=y0 and y<clip.y1):continue
            patch=img[max(0,int(y0*3)):int(y*3)+1,int(x*3):int(right*3)+1,:3]
            patch=patch.astype('int16')
            red=bool(((patch[:,:,0]>patch[:,:,1]+40)&(patch[:,:,0]>160)).any()) if patch.size else False
            rows.append({'PAGINA':i+1,'X':round(x,2),'Y':round(y,2),'ANCHO':round(min(582-x,right-x+2),2),'TAMANO':round(max(5.5,min(16,(y-y0)*1.08)),2),'ESTILO':'b' if red else 'r','ROJO':red,'TEXTO':text,'CONFIANZA':round(float(conf),3)})
    if not rows:raise ValueError('No se detectó texto en las páginas elegidas.')
    return rows
