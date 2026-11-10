"""Carátula didáctica reconstruida; sin imagen ni datos del original."""
from pathlib import Path
import fitz,pandas as pd

def totals(df):
    descriptions=df['DESCRIPCION'].astype(str).str.strip().str.upper()
    opening=df[descriptions.eq('SALDO ANTERIOR')]
    work=df[~descriptions.eq('SALDO ANTERIOR')]
    if work.empty:raise ValueError('La carátula necesita al menos un movimiento.')
    dep=pd.to_numeric(work['DEPOSITO'],errors='raise').fillna(0)
    ret=pd.to_numeric(work['RETIRO'],errors='raise').fillna(0)
    saldo=pd.to_numeric(work['SALDO'],errors='coerce')
    if len(opening):initial=float(pd.to_numeric(opening['SALDO'],errors='raise').iloc[0])
    elif pd.notna(saldo.iloc[0]):initial=float(saldo.iloc[0]-dep.iloc[0]+ret.iloc[0])
    else:raise ValueError('Falta el saldo anterior o el saldo del primer movimiento para calcular la carátula.')
    final=round(initial+float(dep.sum())-float(ret.sum()),2)
    printed=float(saldo.dropna().iloc[-1]) if saldo.notna().any() else None
    return {'initial':initial,'deposits':float(dep.sum()),'withdrawals':float(ret.sum()),'final':final,'count':len(work),'difference':None if printed is None else round(printed-final,2)}

def cover_pdf(df,fields,start,end,finish='NITIDO'):
    data=totals(df);root=Path(__file__).parent
    regular=root/('CLC_Select_Escaneo_Editable.ttf' if finish=='ESCANEO_EDITABLE' else 'CLC_Movimientos_Select_Regular.ttf')
    paths={'r':regular,'b':root/'Select_Referencia_Bold.ttf','n':root/'Select_Referencia_Numeros_Bold.ttf','i':root/'Select_Referencia_BoldItalic.ttf'}
    fonts={k:fitz.Font(fontfile=str(v)) for k,v in paths.items()}
    doc=fitz.open();p=doc.new_page(width=612,height=792)
    def text(x,y,value,size=8.5,kind='r',color=(0,0,0)):
        p.insert_text((x,y),str(value),fontsize=size,fontfile=str(paths[kind]),fontname=kind,color=color)
    def fitted(x,y,value,width,size=8.5,kind='r'):
        value=str(value);length=fonts[kind].text_length(value,fontsize=size)
        text(x,y,value,min(size,size*width/max(length,.01)),kind)
    def amount(right,y,value,size=8.5,kind='r'):
        value=f'{value:,.2f}';text(right-fonts[kind].text_length(value,fontsize=size),y,value,size,kind)
    def bar(y,height=14,x0=30,x1=582,gray=.77):p.draw_rect((x0,y,x1,y+height),color=None,fill=(gray,gray,gray))
    text(30,28,'MATERIAL DIDACTICO - DATOS FICTICIOS',10,'b',(0.65,0,0))
    text(30,49,'FORMATO DE PRACTICA SANTANDER SELECT',10,'b')
    fitted(350,72,'ESTADO DE CUENTA',232,18,'i')
    text(30,86,'CARATULA RECONSTRUIDA PARA EJERCICIO',8)
    fitted(30,116,fields.get('holder','PERSONA FICTICIA'),300,11,'b')
    address=fields.get('address','CALLE DIDACTICA 123, CIUDAD DE EJEMPLO')
    words=str(address).split();line='';y=132
    for word in words:
        candidate=(line+' '+word).strip()
        if fonts['r'].text_length(candidate,fontsize=8.5)>285 and line:text(30,y,line);line=word;y+=11
        else:line=candidate
    if line:fitted(30,y,line,285)
    for y,label,key in [(110,'CLIENTE','client'),(125,'RFC','rfc'),(140,'CUENTA','account'),(155,'CLABE','clabe'),(170,'SUCURSAL','branch')]:
        text(348,y,label,8.5,'b');fitted(410,y,fields.get(key,''),170)
    text(348,185,'PERIODO',8.5,'b');fitted(410,185,f'{pd.Timestamp(start):%d/%m/%Y} - {pd.Timestamp(end):%d/%m/%Y}',170)
    red=(1,0,0)
    text(56,220,'Resumen informativo.',13,'b',red)
    text(30,240,'Resumen de saldos del ejercicio',10)
    bar(249,24)
    for x,value in [(40,'PRODUCTO'),(210,'CUENTA'),(358,'SALDO INICIAL'),(480,'SALDO FINAL')]:text(x,265,value,8.5,'b')
    text(40,292,'CUENTA DIDACTICA');fitted(210,292,fields.get('account','0000000000'),126)
    amount(428,292,data['initial']);amount(574,292,data['final'])
    bar(305,20,x0=344)
    text(40,320,'TOTAL',10,'b');amount(428,320,data['initial'],10.5,'n');amount(574,320,data['final'],10.5,'n')
    text(56,365,'Cuenta de cheques.',13,'b',red)
    bar(376)
    fitted(40,386,'CUENTA DIDACTICA - '+fields.get('account','0000000000'),255,8.5,'b')
    fitted(301,386,'CLABE: '+fields.get('clabe','000000000000000000'),274,8.5,'b')
    for y,label,value in [(416,'Saldo inicial',data['initial']),(436,'+ Depositos',data['deposits']),(456,'- Retiros',data['withdrawals'])]:text(348,y,label);amount(574,y,value)
    text(40,416,'Numero de movimientos');text(242,416,data['count'])
    text(40,436,'Dias del periodo');text(242,436,(pd.Timestamp(end)-pd.Timestamp(start)).days+1)
    text(40,456,'Moneda');text(242,456,'MXN')
    p.draw_rect((344,469,582,492),color=None,fill=(1,.84,.35))
    text(350,485,'SALDO FINAL:',10,'b');amount(574,485,data['final'],10.5,'n')
    text(30,535,'DATOS DE EJEMPLO PARA CLASE',10,'b')
    text(30,552,'Los totales se calculan a partir de los movimientos del ejercicio.',9)
    if data['difference'] is not None and abs(data['difference'])>.015:
        text(30,575,'REVISAR: el ultimo saldo de la tabla difiere del saldo calculado.',9,'b',(0.65,0,0))
        text(30,591,f'Diferencia: {data["difference"]:,.2f} MXN. No se modificaron los saldos de la tabla.',8.5)
    text(30,745,'MATERIAL DIDACTICO - DATOS FICTICIOS - CARATULA',9,'b',(0.65,0,0))
    doc.set_metadata({'title':'Carátula didáctica de movimientos ficticios','subject':'Plantilla reconstruida; sin logos, QR ni sellos bancarios'})
    return doc.tobytes(garbage=4,deflate=True)

def prepend_cover(movements,df,identity,start,end):
    finish=identity.get('select_finish','NITIDO');cover=cover_pdf(df,identity['select_cover'],start,end,finish)
    if finish=='ESCANEO':
        from santander_select import scan_finish
        cover=scan_finish(cover)
    result=fitz.open(stream=cover,filetype='pdf');body=fitz.open(stream=movements,filetype='pdf');result.insert_pdf(body)
    return result.tobytes(garbage=4,deflate=True)
