"""Bankaool reader, formula workbook and visibly didactic editable layout."""
from pathlib import Path
from io import BytesIO
from decimal import Decimal
import json,re,calendar,math
import fitz,pandas as pd
ROOT=Path(__file__).parent
EXTRA=['FECHA_APLICACION','TIPO_MOVIMIENTO','ES_IVA_COMISION']
def spans(page):return [{**s,'text':s['text'].replace('\xa0',' ')} for b in page.get_text('dict')['blocks'] for line in b.get('lines',[]) for s in line['spans']]
def money(value):return float(re.sub(r'[^0-9.\-]','',str(value)) or 0)
def movement_type(desc,dep,ret):
 t=str(desc).upper();compact=re.sub(r'[^A-Z0-9]','',t)
 if compact.startswith('IVA'):return 'IVA_COMISION'
 if 'COMISION' in compact:return 'COMISION'
 if 'SPEI' in t:return 'SPEI_RECIBIDO' if dep else 'SPEI_ENVIADO'
 if 'TRANSFERENCIA' in t:return 'TRANSFERENCIA_RECIBIDA' if dep else 'TRANSFERENCIA_ENVIADA'
 return 'ABONO' if dep else 'CARGO'
def extract_bankaool(content):
 doc=fitz.open(stream=content,filetype='pdf');first=spans(doc[0]);period=re.search(r'Del (\d{2}/\d{2}/\d{4}) al (\d{2}/\d{2}/\d{4})',doc[0].get_text().replace('\xa0',' '))
 if not period:raise ValueError('No se encontró el periodo Bankaool. Este lector requiere un PDF con texto seleccionable.')
 def value(y,left=520):
  matches=[s['text'] for s in first if abs(s['origin'][1]-y)<.2 and s['origin'][0]>left]
  if not matches and y in [163,175,187] and left==75:return ''
  if not matches:raise ValueError(f'No se encontró el campo Bankaool en la altura {y}; revisar la carátula.')
  return matches[0]
 fields={'period_start':str(pd.to_datetime(period[1],dayfirst=True).date()),'period_end':str(pd.to_datetime(period[2],dayfirst=True).date()),'opening':money(value(322)), 'original_deposits':money(value(336)), 'original_other_charges':money(value(350)), 'original_commissions':money(value(378)), 'original_commission_iva':money(value(392)), 'original_final':money(value(420)), 'original_average':money(value(468)), 'holder':value(129,75),'client':value(112,75),'rfc':value(146,75),'address':'\n'.join(value(y,75) for y in [163,175,187]),'account':value(322,140) if False else next(s['text'] for s in first if abs(s['origin'][1]-322)<.2 and 140<s['origin'][0]<295),'clabe':next(s['text'] for s in first if abs(s['origin'][1]-364)<.2 and 100<s['origin'][0]<295),'product':'MI CUENTA MAESTRA BANKAOOL','branch':'POLANCO','currency':'MXN','interest':0.,'tax':0.,'annual_rate':0.,'gat':'N/A','gat_real':'N/A','bank_model_version':'BANKAOOL_V1'}
 for y,key in [(336,'product'),(350,'currency'),(378,'branch')]:fields[key]=next(s['text'] for s in first if abs(s['origin'][1]-y)<.2 and 100<s['origin'][0]<295)
 fields.update(interest=money(value(170)),tax=money(value(406)),annual_rate=money(value(150)),gat=value(110),gat_real=value(130))
 fields['address']='\n'.join(s['text'] for s in first if 75<s['origin'][0]<301 and 162<s['origin'][1]<190)
 rows=[]
 for number,page in enumerate(doc,1):
  if 'Movimientos Realizados' not in page.get_text().replace('\xa0',' '):continue
  ss=spans(page);heads=sorted([s for s in ss if s['origin'][0]<69 and 260<s['origin'][1]<736 and re.fullmatch(r'\d{2}/\d{2}/\d{4}',s['text'])],key=lambda s:s['origin'][1])
  for j,head in enumerate(heads):
   y=head['origin'][1];limit=heads[j+1]['origin'][1]-2 if j+1<len(heads) else 735
   same=[s for s in ss if abs(s['origin'][1]-y)<.25]
   desc=sorted([s for s in ss if 197<=s['origin'][0]<349 and y-.1<=s['origin'][1]<limit],key=lambda s:s['origin'][1])
   # Stop before a page total; monetary columns are read only on the operation head.
   lines=[s['text'].replace('R\ufffdNICA','RÓNICA') for s in desc if s['text']!='Totales']
   dep=sum(money(s['text']) for s in same if 431<=s['origin'][0]<513 and '$' in s['text']);ret=sum(money(s['text']) for s in same if 349<=s['origin'][0]<431 and '$' in s['text'])
   saldo=next(money(s['text']) for s in same if 513<=s['origin'][0]<595 and '$' in s['text'])
   applied=next(s['text'] for s in same if 69<=s['origin'][0]<128);folio=next(s['text'] for s in same if 128<=s['origin'][0]<198)
   text='\n'.join(lines);kind=movement_type(text,dep,ret)
   rows.append({'FECHA':pd.to_datetime(head['text'],dayfirst=True),'FECHA_APLICACION':pd.to_datetime(applied,dayfirst=True),'FOLIO':folio,'DESCRIPCION':text,'DEPOSITO':dep,'RETIRO':ret,'SALDO':saldo,'TIPO_MOVIMIENTO':kind,'ES_COMISION':'SI' if kind=='COMISION' else 'NO','ES_IVA_COMISION':'SI' if kind=='IVA_COMISION' else 'NO','HOJA_ORIGINAL':number,'Y_ORIGINAL':y,'ORIGEN':'ORIGINAL','_LINEAS_DOCUMENTO':json.dumps(lines,ensure_ascii=False)})
 df=pd.DataFrame(rows);balance=Decimal(str(fields['opening']))
 for row in df.itertuples():
  balance+=Decimal(str(row.DEPOSITO))-Decimal(str(row.RETIRO))
  if abs(balance-Decimal(str(row.SALDO)))>Decimal('.01'):raise ValueError(f'El saldo no concilia en referencia {row.FOLIO}; revisar la extracción.')
 commissions=float(df.loc[df.ES_COMISION=='SI','RETIRO'].sum());iva=float(df.loc[df.ES_IVA_COMISION=='SI','RETIRO'].sum())
 assert abs(df.DEPOSITO.sum()-fields['original_deposits'])<.01
 assert abs(df.RETIRO.sum()-commissions-iva-fields['original_other_charges'])<.01
 assert abs(commissions-fields['original_commissions'])<.01 and abs(iva-fields['original_commission_iva'])<.01
 assert abs(float(balance)-fields['original_final'])<.01
 opening={'FECHA':pd.Timestamp(fields['period_start']),'FECHA_APLICACION':pd.Timestamp(fields['period_start']),'FOLIO':'','DESCRIPCION':'SALDO ANTERIOR','DEPOSITO':0.,'RETIRO':0.,'SALDO':fields['opening'],'ES_COMISION':'NO','ES_IVA_COMISION':'NO','TIPO_MOVIMIENTO':'SALDO_ANTERIOR','ORIGEN':'ORIGINAL'}
 df=pd.concat([pd.DataFrame([opening]),df],ignore_index=True);df.attrs['bankaool_cover']=fields;df.attrs['reconciliation']={'financial_reconciled':True,'movement_count':len(rows),'deposits':float(df.DEPOSITO.sum()),'withdrawals':float(df.RETIRO.sum()),'opening':fields['opening'],'final':float(balance)}
 return df
def normalized(df,start,end):
 from select_editable_document import normalized as normalize
 work,t=normalize(df,start,end)
 work['FECHA_APLICACION']=pd.to_datetime(work.get('FECHA_APLICACION',work.FECHA),errors='coerce').fillna(work.FECHA)
 # Source accounting balances follow transaction order. Compare both day bases explicitly.
 for basis in ['FECHA','FECHA_APLICACION']:
  balance=t['initial'];daily=[]
  for day in pd.date_range(start,end):
   current=work[work[basis].dt.normalize()==day]
   if len(current):balance=float(current.SALDO.iloc[-1])
   daily.append(balance)
  t['average_'+basis]=sum(daily)/len(daily)
 work['TIPO_MOVIMIENTO']=[movement_type(r.DESCRIPCION,r.DEPOSITO,r.RETIRO) for r in work.itertuples()]
 t['commissions']=float(work.loc[work.TIPO_MOVIMIENTO=='COMISION','RETIRO'].sum());t['commission_iva']=float(work.loc[work.TIPO_MOVIMIENTO=='IVA_COMISION','RETIRO'].sum());t['other_charges']=t['withdrawals']-t['commissions']-t['commission_iva']
 return work,t
def workbook_model(df,cover_fields=None):
 from select_formula_workbook import prepare_model,HEADERS,serial
 attrs=df.attrs.copy()
 if df.FECHA.isna().any():
  merged=[]
  for _,r in df.iterrows():
   if pd.notna(r.FECHA):merged.append(r.to_dict())
   elif merged:merged[-1]['DESCRIPCION']+='\n'+str(r.DESCRIPCION)
  df=pd.DataFrame(merged);df.attrs=attrs
 fields=df.attrs.get('bankaool_cover',{}).copy();generated='ORIGEN' in df and df.ORIGEN.astype(str).str.upper().eq('GENERADO').any()
 if generated:fields.update(cover_fields or {})
 old_start=pd.Timestamp(fields['period_start']) if fields.get('period_start') else None
 dates=pd.to_datetime(df.FECHA);start=dates.min().replace(day=1);end=dates.max()+pd.offsets.MonthEnd(0)
 if generated and df.attrs.get('period_start') is not None and df.attrs.get('period_end') is not None:
  candidate_start=pd.Timestamp(df.attrs['period_start']);candidate_end=pd.Timestamp(df.attrs['period_end'])
  if candidate_start<=dates.min() and dates.max()<=candidate_end:start,end=candidate_start,candidate_end
 if generated or not fields.get('period_start') or dates.min()<pd.Timestamp(fields['period_start']) or dates.max()>pd.Timestamp(fields.get('period_end',end)):fields.update(period_start=str(start.date()),period_end=str(end.date()))
 fields.update(product='MI CUENTA MAESTRA BANKAOOL',bank_model_version='BANKAOOL_V1');rows=[];extra=[]
 if not df.DESCRIPCION.astype(str).str.upper().eq('SALDO ANTERIOR').any():
  first=df.iloc[0];opening=float(first.SALDO)-float(first.DEPOSITO or 0)+float(first.RETIRO or 0)
  prior={'FECHA':pd.Timestamp(fields['period_start']),'FECHA_APLICACION':pd.Timestamp(fields['period_start']),'FOLIO':'','DESCRIPCION':'SALDO ANTERIOR','DEPOSITO':0.,'RETIRO':0.,'SALDO':opening,'ES_COMISION':'NO','ORIGEN':'ORIGINAL'}
  df=pd.concat([pd.DataFrame([prior]),df],ignore_index=True);df.attrs=attrs
 for i,r in df.iterrows():
  if str(r.get('ORIGEN','')).upper()=='GENERADO':
   r=r.copy()
   if movement_type(r.DESCRIPCION,r.DEPOSITO,r.RETIRO) not in ['COMISION','IVA_COMISION'] and 'SPEI' not in str(r.DESCRIPCION).upper():r['DESCRIPCION']=('ABONO POR SPEI' if r.DEPOSITO else 'CARGO POR SPEI')+'\n'+str(r.DESCRIPCION)
   ref=r.get('FOLIO');r['FOLIO']=f'{pd.Timestamp(r.FECHA):%y%m}{i:04d}' if ref is None or pd.isna(ref) or str(ref)=='' else str(ref)
  line=[None if pd.isna(r.get(c)) else r.get(c) for c in HEADERS]
  line[0]=str(pd.Timestamp(r.FECHA).date());line[6]=i;line[7]=1;line[8]='NO';rows.append(line)
  kind='SALDO_ANTERIOR' if r.DESCRIPCION=='SALDO ANTERIOR' else movement_type(r.DESCRIPCION,r.DEPOSITO,r.RETIRO)
  application=pd.to_datetime(r.get('FECHA_APLICACION'),errors='coerce')
  if pd.isna(application):application=pd.Timestamp(r.FECHA)
  elif old_start is not None and generated and not pd.Timestamp(fields['period_start'])<=application<=pd.Timestamp(fields['period_end']):application+=pd.DateOffset(months=(start.year-old_start.year)*12+start.month-old_start.month)
  extra.append([serial(application),kind,'SI' if kind=='IVA_COMISION' else 'NO'])
 model=prepare_model(rows,fields);n=len(model['rows'])+1
 keyrow=model['keyrow'];met=model['metadata'];b=lambda k:f'CARATULA!$B${keyrow[k]}'
 iva=sum(float(r[4] or 0) for r,e in zip(model['rows'],extra) if e[2]=='SI')
 keyrow['commission_iva']=len(met)+2;met.append(['commission_iva',None,'IVA de comisiones: automático'])
 model['formulas'].append(['CARATULA',f"B{keyrow['commission_iva']}",f'=SUMIFS(MOVIMIENTOS!E3:E{n},MOVIMIENTOS!R3:R{n},"SI")',iva]);model['expected']['commission_iva']=iva
 for entry in model['formulas']:
  if entry[0]=='CARATULA' and entry[1]==f"B{keyrow['other_charges']}":entry[2]=f'={b("withdrawals")}-{b("commissions")}-{b("commission_iva")}';entry[3]-=iva
  if entry[0]=='CARATULA' and entry[1] in ['F5','G5']:entry[3]-=iva
 model['expected']['other_charges']-=iva;model['extra_rows']=extra
 assert len(model['rows'])==len(extra),'Las fechas de aplicación deben corresponder a cada fila.'
 model['formulas']=[r for r in model['formulas'] if not(r[0]=='CARATULA' and re.fullmatch(r'[FG][4-7]',r[1]))]
 model['chart_labels']=['COMISIONES','DEPÓSITOS','RETIROS EN EFECTIVO','INTERÉS PAGADO','IVA','RETENCIÓN ISR','OTROS CARGOS','SALDO FINAL']
 for row,key in enumerate(['commissions','deposits',None,'interest','commission_iva','tax','other_charges','final_balance'],4):
  formula=f'={b(key)}' if key else '=0';value=model['expected'].get(key,fields.get(key,0)) if key else 0
  model['formulas'] += [['CARATULA',f'F{row}',formula,value],['CARATULA',f'G{row}',f'=ABS(F{row})',abs(value)]]
 return model
def read_excel(content):
 from select_excel_reader import read_select_excel
 # Preserve the physical concept parser and recover Bankaool-specific dates by operation order.
 df=read_select_excel(content);book=pd.ExcelFile(BytesIO(content));raw=pd.read_excel(book,sheet_name='MOVIMIENTOS',dtype={'FOLIO':str});heads=raw[raw.FECHA.notna()]
 for col in EXTRA:
  if col in heads:df[col]=heads[col].values
 fields=df.attrs.pop('select_cover',{});work,t=normalized(df,fields['period_start'],fields['period_end']);fields.update(commissions=t['commissions'],commission_iva=t['commission_iva'],other_charges=t['other_charges'],final_balance=t['final'],average_balance=t['average']);df.attrs['bankaool_cover']=fields
 return df
def build_excel(df,cover_fields=None):
 import xlsxwriter
 from select_formula_workbook import HEADERS
 model=workbook_model(df,cover_fields);buf=BytesIO();book=xlsxwriter.Workbook(buf,{'in_memory':True});book.set_calc_mode('auto');sheets={name:book.add_worksheet(name) for name in ['MOVIMIENTOS','CARATULA','SALDOS_DIARIOS','CONTROL']}
 header=book.add_format({'bold':True,'bg_color':'#000000','font_color':'white','text_wrap':True});date=book.add_format({'num_format':'yyyy-mm-dd'});amount=book.add_format({'num_format':'#,##0.00'});inp=book.add_format({'num_format':'#,##0.00','bg_color':'#FFF2CC'});auto=book.add_format({'num_format':'#,##0.00','bg_color':'#EAF1F8'});txt=book.add_format({'num_format':'@','text_wrap':True})
 m=sheets['MOVIMIENTOS'];m.write_row(0,0,HEADERS+EXTRA,header);m.set_row(0,36);m.freeze_panes(1,0);m.set_column(0,1,15);m.set_column(2,2,70);m.set_column(3,5,18);m.set_column(6,17,19)
 for i,(row,extra) in enumerate(zip(model['rows'],model['extra_rows']),1):
  for j,v in enumerate(row+extra):
   if v is not None:m.write(i,j,v,date if j in [0,15] else inp if j in [3,4] or(j==5 and i==1) else amount if j==5 else txt if j in [1,2] else None)
  m.set_row(i,max(20,14*len(str(row[2]).splitlines())))
 c=sheets['CARATULA'];c.write_row(0,0,['CAMPO','VALOR','USO'],header);c.set_column(0,0,26);c.set_column(1,1,55);c.set_column(2,2,45)
 for i,row in enumerate(model['metadata'],1):
  for j,v in enumerate(row):
   if v is not None:c.write(i,j,v,date if j==1 and str(row[0]).startswith('period_') else txt if isinstance(v,str) else inp)
 d=sheets['SALDOS_DIARIOS'];d.write_row(0,0,['FECHA','DEPOSITOS','RETIROS','FLUJO NETO','SALDO AL CIERRE'],header);d.set_column(0,4,19)
 c.write_row(2,4,['CONCEPTO','IMPORTE','MAGNITUD'],header)
 for i,label in enumerate(model['chart_labels'],3):c.write(i,4,label)
 chart=book.add_chart({'type':'doughnut'});chart.add_series({'categories':'=CARATULA!$E$4:$E$11','values':'=CARATULA!$G$4:$G$11'});c.insert_chart('E14',chart)
 t=sheets['CONTROL'];t.write_row(0,0,['COMPARACION','ORIGINAL','EDITADO','CAMBIO'],header);t.set_column(0,3,20)
 for i,key in enumerate(['opening','deposits','withdrawals','final','movement_count'],1):t.write(i,0,key);t.write(i,1,model['controls'][key])
 for sheet,cell,formula,value in model['formulas']:sheets[sheet].write_formula(cell,formula,date if sheet=='SALDOS_DIARIOS' and cell.startswith('A') else auto,value)
 book.close();return buf.getvalue()
def generate_pdf(df,identity,start,end,scope='CARATULA_MOVIMIENTOS'):
 fields={**df.attrs.get('bankaool_cover',{}),**identity.get('bankaool_cover',{})}
 if identity.get('bankaool_generated'):
  donor=df.attrs.get('bankaool_cover',{}).get('period_start')
  if donor and 'FECHA_APLICACION' in df:
   donor=pd.Timestamp(donor);offset=(pd.Timestamp(start).year-donor.year)*12+pd.Timestamp(start).month-donor.month
   df=df.copy();df['FECHA_APLICACION']=pd.to_datetime(df.FECHA_APLICACION,errors='coerce')+pd.DateOffset(months=offset)
  fields['period_start']=str(pd.Timestamp(start).date());fields['period_end']=str(pd.Timestamp(end).date())
 start=pd.Timestamp(fields.get('period_start',start)).normalize();end=pd.Timestamp(fields.get('period_end',end)).normalize()
 work,t=normalized(df,start,end);template=fitz.open(ROOT/'bankaool_template.pdf');doc=fitz.open();arial=ROOT/'Bankaool_Arial.ttf';bold=ROOT/'Bankaool_Arial_Bold.ttf';times=ROOT/'Bankaool_Times.ttf'
 fonts={k:fitz.Font(fontfile=str(p)) for k,p in [('a',arial),('b',bold),('t',times)]};paths={'a':arial,'b':bold,'t':times}
 def text(page,x,y,value,size=9,k='a',right=None,width=None,color=(0,0,0)):
  value=str(value);length=fonts[k].text_length(value,fontsize=size);scale=min(1,width/max(length,.01)) if width else 1
  if right is not None:x=right-length*scale
  page.insert_text((x,y),value,fontfile=str(paths[k]),fontname='BK'+k,fontsize=size,color=color,morph=(fitz.Point(x,y),fitz.Matrix(scale,1)))
 def moneytext(page,right,y,value,size=9,k='a'):text(page,0,y,f'$ {value:,.2f}',size,k,right=right)
 def page(kind):
  doc.insert_pdf(template,from_page=kind,to_page=kind);p=doc[-1]
  if kind:
   # Restore the rounded date frame removed together with the original date values.
   shape=p.new_shape();r=5;k=.55228475*r;x0,y0,x1,y1=305,50,595,90
   shape.draw_line((x0+r,y0),(x1-r,y0));shape.draw_bezier((x1-r,y0),(x1-r+k,y0),(x1,y0+r-k),(x1,y0+r))
   shape.draw_line((x1,y0+r),(x1,y1-r));shape.draw_bezier((x1,y1-r),(x1,y1-r+k),(x1-r+k,y1),(x1-r,y1))
   shape.draw_line((x1-r,y1),(x0+r,y1));shape.draw_bezier((x0+r,y1),(x0+r-k,y1),(x0,y1-r+k),(x0,y1-r))
   shape.draw_line((x0,y1-r),(x0,y0+r));shape.draw_bezier((x0,y0+r),(x0,y0+r-k),(x0+r-k,y0),(x0+r,y0));shape.finish(color=(0,0,0),width=2);shape.commit()
  for y,key,default in [(112,'client','00000000000'),(129,'holder','PERSONA FICTICIA PARA CLASE'),(146,'rfc','XAXX010101000')]:text(p,78,y,fields.get(key,default),width=215)
  for j,line in enumerate(str(fields.get('address','DOMICILIO DIDACTICO')).splitlines()[:3]):text(p,78,163+j*12,line,width=217)
  text(p,0,65,f'Del {start:%d/%m/%Y} al {end:%d/%m/%Y}',10,right=590);text(p,0,84,f'{end+pd.Timedelta(days=1):%d/%m/%Y}',10,right=590)
  text(p,15,12,'EJEMPLO DIDACTICO - DATOS FICTICIOS - SIN VALIDEZ BANCARIA',8,'b',color=(.8,0,0));return p
 if scope!='MOVIMIENTOS':
  p=page(0)
  for y,key,default in [(322,'account','00000000'),(336,'product','MI CUENTA MAESTRA BANKAOOL'),(350,'currency','PESOS MXN'),(364,'clabe','000000000000000000'),(378,'branch','SUCURSAL DIDACTICA')]:text(p,0,y,fields.get(key,default),right=285,width=245)
  for y,val in [(110,fields.get('gat','N/A')),(130,fields.get('gat_real','N/A')),(150,f"{float(fields.get('annual_rate',0)):.2f} %")]:text(p,0,y,val,10,right=585)
  for y,val,size in [(170,float(fields.get('interest',0)),10),(195,t['commissions'],10),(215,t['commission_iva'],10),(235,t['commissions']+t['commission_iva'],10),(322,t['initial'],9),(336,t['deposits'],9),(350,t['other_charges'],9),(364,float(fields.get('interest',0)),9),(378,t['commissions'],9),(392,t['commission_iva'],9),(406,float(fields.get('tax',0)),9),(420,t['final'],9),(468,t['final'],9),(482,t['final'],9)]:moneytext(p,585,y,val,size)
  for y,val in [(392,0),(468,t['average']),(510,float(fields.get('interest',0))),(524,float(fields.get('tax',0))),(571,0),(585,t['commission_iva']),(599,t['commissions']+t['commission_iva'])]:moneytext(p,285,y,val)
  text(p,0,482,t['days'],right=285);text(p,0,496,f"{float(fields.get('annual_rate',0)):.2f} %",right=285)
 positions=all(pd.notna(r.get('HOJA_ORIGINAL')) and pd.notna(r.get('Y_ORIGINAL')) for _,r in work.iterrows())
 p=None;current=None;y=273.618
 def table(p):
  text(p,15,225,'Movimientos Realizados (PESOS MXN)',12,'b');p.draw_rect((15,230,595,261.598),fill=(0,0,0),color=None)
  for x in [69,128,198,349,431,513]:p.draw_line((x,230),(x,261.598),width=.25,color=(.4,)*3)
  for x,value in [(23.49,'Fecha Movi-\nmiento'),(77.54,'Fecha Aplica-\nción'),(143.66,'Referencia'),(229.71,'Transacción Descripción'),(379.1,'Cargo'),(460.43,'Abono'),(543.77,'Saldo')]:
   for j,line in enumerate(value.splitlines()):
    origin=33.494 if j and x<69 else 94.662 if j and x<128 else x
    text(p,origin,242.242+j*11.199 if '\n' in value else 249.242,line,8,color=(1,1,1))
 for index,(_,row) in enumerate(work.iterrows()):
  number=int(row.HOJA_ORIGINAL) if positions else 0;lines=str(row.DESCRIPCION).splitlines()
  if str(row.get('ORIGEN','')).upper()!='ORIGINAL':
   if row.TIPO_MOVIMIENTO not in ['COMISION','IVA_COMISION']:lines=['ABONO POR SPEI' if row.DEPOSITO else 'CARGO POR SPEI']+lines
  if not positions:
   wrapped=[]
   for line in lines:
    current=''
    for word in line.split():
     if current and fonts['t'].text_length(current+' '+word,fontsize=9)>149:wrapped.append(current);current=word
     else:current=(current+' '+word).strip()
    wrapped.append(current)
   lines=wrapped
  height=len(lines)*12.349+16.698
  if p is None or (positions and current!=number) or (not positions and y+height>735):p=page(1);table(p);current=number;y=273.618
  if positions:y=float(row.Y_ORIGINAL)
  text(p,25,y,f'{row.FECHA:%d/%m/%Y}',9,'t');text(p,81.5,y,f'{row.FECHA_APLICACION:%d/%m/%Y}',9,'t');ref=str(row.get('FOLIO',''))
  if ref in ['','nan','None']:ref=f'{row.FECHA:%y%m}{index+1:04d}'
  text(p,163-fonts['t'].text_length(ref,fontsize=9)/2,y,ref,9,'t',width=66)
  for j,line in enumerate(lines):text(p,198,y+j*12.349,line,9,'t',width=151)
  if row.RETIRO:moneytext(p,431,y,row.RETIRO,9,'t')
  if row.DEPOSITO:moneytext(p,513,y,row.DEPOSITO,9,'t')
  moneytext(p,595,y,row.SALDO,9,'t');bottom=y+height-12.02;p.draw_line((15,bottom),(595,bottom),width=.5,color=(.3,)*3);y+=height
 if y+35>735:p=page(1);table(p);y=273.618
 text(p,198,y,'Totales',9,'t');moneytext(p,431,y,t['withdrawals'],9,'t');moneytext(p,513,y,t['deposits'],9,'t')
 if scope!='MOVIMIENTOS':
  p=page(2);categories=[('COMISIONES',t['commissions']),('DEPÓSITOS',t['deposits']),('RETIROS EN EFECTIVO',0),('INTERÉS PAGADO',float(fields.get('interest',0))),('IVA',t['commission_iva']),('RETENCIÓN ISR',float(fields.get('tax',0))),('OTROS CARGOS',t['other_charges']),('SALDO FINAL',t['final'])];colors=[(.78,)*3,(.44,)*3,(1,.48,0),(1,0,.2),(0,.55,.65),(0,.65,.28),(.49,.14,.46),(.97,.94,0)];total=sum(abs(v) for _,v in categories)
  text(p,80,242,'SALDO INICIAL',9);moneytext(p,310,242,t['initial'])
  for j,((name,value),color) in enumerate(zip(categories,colors)):
   y=262+j*20;p.draw_rect((50,y-12,70,y+3),fill=color,color=None);text(p,80,y,name,9);text(p,0,y,f'{abs(value)/max(total,1):.0%}',9,right=230);moneytext(p,310,y,value)
  cx,cy,r=470,310,80;angle=0
  for (_,v),color in zip(categories,colors):
   delta=2*math.pi*abs(v)/max(total,1)
   if delta:
    n=max(3,int(delta*50));points=[fitz.Point(cx,cy)]+[fitz.Point(cx+r*math.cos(angle+delta*i/n),cy+r*math.sin(angle+delta*i/n)) for i in range(n+1)];p.draw_polyline(points,fill=color,color=(1,1,1),width=.3,closePath=True)
   angle+=delta
  p.draw_circle((cx,cy),40,fill=(1,1,1),color=None)
 if scope=='COMPLETO':
  page(3);p=page(4);text(p,30,280,'EJEMPLO DIDACTICO: SIN FOLIO, SELLOS NI CERTIFICACION FISCAL',10,'b',width=550)
 for i,p in enumerate(doc,1):text(p,15,784,f'EJEMPLO DIDACTICO / SIN VALIDEZ BANCARIA / {i} de {len(doc)}',8,'b',color=(.8,0,0))
 doc.set_metadata({'title':'Bankaool - ejercicio didáctico editable','subject':'Formato didáctico; importes recalculados desde Excel.'});return doc.tobytes(garbage=4,deflate=True)
