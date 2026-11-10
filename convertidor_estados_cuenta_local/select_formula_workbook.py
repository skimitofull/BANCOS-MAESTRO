"""Automatic Select workbook model shared by the app and its Excel templates."""
from pathlib import Path
from io import BytesIO
from datetime import datetime,timedelta
from decimal import Decimal
import calendar,json,re
import pandas as pd
MODEL_VERSION='SELECT_FORMULAS_V1'
HEADERS=['FECHA','FOLIO','DESCRIPCION','DEPOSITO','RETIRO','SALDO','ID_MOVIMIENTO','RENGLON_MOVIMIENTO','ES_CONTINUACION','HOJA_ORIGINAL','Y_ORIGINAL','TEXTO_OCR_CRUDO','REVISION_OCR','CONFIANZA_OCR','ES_COMISION']
MODEL_FIELDS=['opening','deposits','withdrawals','final_balance','average_balance','days','operations','income_count','expense_count','other_charges','net_interest','commissions']
DEFAULT_FIELDS={'holder':'PERSONA FICTICIA PARA CLASE','address':'DOMICILIO FICTICIO','rfc':'XAXX010101000','client':'00000000','account':'0000000000','secondary_account':'0000000001','clabe':'000000000000000000','branch':'SUCURSAL DIDACTICA','phone':'0000000000','postal':'00000','regime':'EJEMPLO','interest':0,'tax':0,'gat':0,'gat_real':0,'annual_rate':0,'currency':'MXN'}
def is_commission(description):
 text=re.sub(r'\s+',' ',str(description)).strip().upper()
 compact=re.sub(r'[^A-Z0-9]','',text)
 return ('ADMINISTRACIONRENTA' in compact or 'COMISION' in compact) and not compact.startswith('IVA')
def serial(value):return (pd.Timestamp(value).normalize()-pd.Timestamp('1899-12-30')).days
def prepare_model(rows,metadata,source_controls=None,source_name=''):
 rows=[list(r[:14])+[r[14] if len(r)>14 else None] for r in rows]
 fields={**DEFAULT_FIELDS,**dict(metadata)}
 dated=[pd.Timestamp(r[0]) for r in rows if r[0] is not None and str(r[0]).strip()]
 if not dated:raise ValueError('El Excel necesita fechas de movimientos.')
 start=pd.Timestamp(fields.get('period_start') or min(dated).replace(day=1)).normalize()
 end=pd.Timestamp(fields.get('period_end') or max(dated).replace(day=calendar.monthrange(max(dated).year,max(dated).month)[1])).normalize()
 if start>end or (end-start).days>365:raise ValueError('Periodo inválido o mayor de un año.')
 fields['period_start']=serial(start);fields['period_end']=serial(end)
 opening=next((r[5] for r in rows if str(r[2]).upper()=='SALDO ANTERIOR'),None)
 if opening is None:
  first=next(r for r in rows if r[0] is not None)
  if first[5] is None:raise ValueError('Falta SALDO ANTERIOR o un saldo que permita obtener la apertura.')
  opening=float(first[5])-float(first[3] or 0)+float(first[4] or 0)
  rows.insert(0,[str(start.date()),'', 'SALDO ANTERIOR',0.,0.,opening,0,1,'NO',None,None,'SALDO ANTERIOR','',1.,None])
 if str(rows[0][2]).upper()!='SALDO ANTERIOR':raise ValueError('SALDO ANTERIOR debe ser la primera fila.')
 formulas=[];heads=[];balance=Decimal(str(opening));by_day={};commissions=0.
 for idx,row in enumerate(rows):
  excel=idx+2
  if row[0] is not None and str(row[0]).strip():
   dt=pd.Timestamp(row[0]).normalize();row[0]=serial(dt)
   if idx==0:row[5]=float(opening);row[14]='NO';continue
   dep=float(row[3] or 0);ret=float(row[4] or 0)
   balance+=Decimal(str(dep))-Decimal(str(ret));row[5]=float(balance);heads.append(excel);by_day[dt]=float(balance)
   row[14]=str(row[14]).upper() if row[14] is not None else 'SI' if is_commission(row[2]) else 'NO'
   if row[14]=='SI':commissions+=ret
   formulas.append(['MOVIMIENTOS',f'F{excel}',f'=ROUND($F$2+SUM($D$3:D{excel})-SUM($E$3:E{excel}),2)',float(balance)])
  else:row[0]=None;row[14]=None
 n=len(rows)+1
 for name in ['commissions','net_interest',*MODEL_FIELDS]:fields.pop(name,None)
 fields['select_model_version']=MODEL_VERSION
 fields.update(previous_distribution=None,current_distribution=None)
 fields.update({name:None for name in ['opening','deposits','withdrawals','final_balance','average_balance','days','operations','income_count','expense_count','commissions','other_charges','net_interest']})
 labels={'previous_distribution':'Distribución mes anterior: automática','current_distribution':'Distribución mes actual: automática','holder':'Nombre: editable','address':'Domicilio: editable','rfc':'RFC: editable','client':'Cliente: editable','account':'Cuenta: editable','secondary_account':'Segunda cuenta: editable','clabe':'CLABE: texto editable','branch':'Sucursal: editable','phone':'Teléfono: texto editable','postal':'CP: texto editable','regime':'Régimen: editable','interest':'Intereses brutos: editable','tax':'ISR retenido: editable','gat':'GAT nominal: editable','gat_real':'GAT real: editable','annual_rate':'Tasa anual: editable','currency':'Moneda: editable','period_start':'Inicio del periodo: editable','period_end':'Fin del periodo: editable','select_model_version':'Modelo automático: conservar','opening':'Saldo inicial: editar MOVIMIENTOS!F2','deposits':'Depósitos: automático','withdrawals':'Retiros: automático','final_balance':'Saldo final: automático','average_balance':'Saldo promedio diario: automático','days':'Días del periodo: automático','operations':'Número de operaciones: automático','income_count':'Operaciones con depósitos: automático','expense_count':'Operaciones con retiros: automático','commissions':'Comisiones: retiros marcados SI','other_charges':'Otros cargos: retiros menos comisiones','net_interest':'Intereses netos: brutos menos ISR'}
 meta=[[k,v,labels.get(k,'Dato editable')] for k,v in fields.items()]
 keyrow={k:i+2 for i,(k,_,_) in enumerate(meta)}
 b=lambda key:f"CARATULA!$B${keyrow[key]}"
 dep=sum(float(r[3] or 0) for r in rows[1:]);ret=sum(float(r[4] or 0) for r in rows[1:]);days=(end-start).days+1
 running=float(opening);daily=[]
 for offset in range(366):
  day=start+timedelta(days=offset)
  if day in by_day:running=by_day[day]
  active=offset<days
  daily.append([serial(day),0.,0.,0.,running] if active else ['', '', '', '', ''])
  row=offset+2
  formulas.extend([
   ['SALDOS_DIARIOS',f'A{row}',f'=IF({b("period_start")}+{offset}<={b("period_end")},{b("period_start")}+{offset},"")',serial(day) if active else ''],
   ['SALDOS_DIARIOS',f'B{row}',f'=IF(A{row}="","",SUMIFS(MOVIMIENTOS!$D$3:$D${n},MOVIMIENTOS!$A$3:$A${n},A{row}))',None],
   ['SALDOS_DIARIOS',f'C{row}',f'=IF(A{row}="","",SUMIFS(MOVIMIENTOS!$E$3:$E${n},MOVIMIENTOS!$A$3:$A${n},A{row}))',None],
   ['SALDOS_DIARIOS',f'D{row}',f'=IF(A{row}="","",B{row}-C{row})',None],
   ['SALDOS_DIARIOS',f'E{row}',f'=IF(A{row}="","",ROUND({b("opening")}+D{row},2))' if offset==0 else f'=IF(A{row}="","",ROUND(E{row-1}+D{row},2))',running if active else '']])
  daily_dep=sum(float(r[3] or 0) for r in rows[1:] if r[0]==serial(day));daily_ret=sum(float(r[4] or 0) for r in rows[1:] if r[0]==serial(day))
  for f in formulas[-5:]:
   if f[1].startswith('B'):f[3]=daily_dep if active else ''
   elif f[1].startswith('C'):f[3]=daily_ret if active else ''
   elif f[1].startswith('D'):f[3]=daily_dep-daily_ret if active else ''
 calc={'opening':('=MOVIMIENTOS!F2',float(opening)),'deposits':(f'=SUM(MOVIMIENTOS!D3:D{n})',dep),'withdrawals':(f'=SUM(MOVIMIENTOS!E3:E{n})',ret),'final_balance':(f'=MOVIMIENTOS!F{heads[-1]}',float(balance)),'average_balance':('=AVERAGE(SALDOS_DIARIOS!E2:E367)',sum(r[4] for r in daily[:days])/days),'days':(f'={b("period_end")}-{b("period_start")}+1',days),'operations':(f'=COUNT(MOVIMIENTOS!A3:A{n})',len(heads)),'income_count':(f'=COUNTIFS(MOVIMIENTOS!D3:D{n},">0")',sum(float(r[3] or 0)>0 for r in rows[1:])),'expense_count':(f'=COUNTIFS(MOVIMIENTOS!E3:E{n},">0")',sum(float(r[4] or 0)>0 for r in rows[1:])),'commissions':(f'=SUMIFS(MOVIMIENTOS!E3:E{n},MOVIMIENTOS!O3:O{n},"SI")',commissions),'other_charges':(f'={b("withdrawals")}-{b("commissions")}',ret-commissions),'net_interest':(f'={b("interest")}-{b("tax")}',float(fields['interest'])-float(fields['tax']))}
 calc.update(previous_distribution=(f'=IF({b("opening")}=0,0,1)',1. if float(opening) else 0.),current_distribution=(f'=IF({b("final_balance")}=0,0,1)',1. if float(balance) else 0.))
 for key,(formula,value) in calc.items():formulas.append(['CARATULA',f'B{keyrow[key]}',formula,value])
 for row,(label,key) in enumerate([('Abonos','deposits'),('Otros cargos','other_charges'),('Comisiones cobradas','commissions'),('Saldo final','final_balance')],4):
  formulas.append(['CARATULA',f'F{row}',f'={b(key)}',calc[key][1]]);formulas.append(['CARATULA',f'G{row}',f'=ABS(F{row})',abs(calc[key][1])])
 controls=source_controls or {'opening':float(opening),'deposits':dep,'withdrawals':ret,'final':float(balance),'movement_count':len(heads),'physical_lines':len(rows)-1}
 for row,key in enumerate(['opening','deposits','withdrawals','final_balance','operations'],2):
  formulas.append(['CONTROL',f'C{row}',f'={b(key)}',calc[key][1]]);origkey='final' if key=='final_balance' else 'movement_count' if key=='operations' else key
  formulas.append(['CONTROL',f'D{row}',f'=ROUND(C{row}-B{row},2)',calc[key][1]-controls[origkey]])
 return {'rows':rows,'metadata':meta,'keyrow':keyrow,'formulas':formulas,'daily':daily,'days':days,'controls':controls,'source_name':source_name,'expected':{k:v[1] for k,v in calc.items()}}

def update_calculated_fields(work,totals,fields):
 fields=dict(fields)
 flags=work.get('ES_COMISION',pd.Series(None,index=work.index,dtype=object))
 inferred=work.DESCRIPCION.map(is_commission)
 tokens=flags.fillna('').astype(str).str.strip().str.upper()
 marked=tokens.eq('SI') | (~tokens.isin(['SI','NO']) & inferred)
 commission=float(work.loc[marked,'RETIRO'].sum())
 fields.update(opening=totals['initial'],deposits=totals['deposits'],withdrawals=totals['withdrawals'],final_balance=totals['final'],average_balance=totals['average'],days=totals['days'],operations=totals['count'],income_count=int((work.DEPOSITO>0).sum()),expense_count=int((work.RETIRO>0).sum()),commissions=commission,other_charges=totals['withdrawals']-commission,net_interest=float(fields.get('interest',0))-float(fields.get('tax',0)))
 fields.update(previous_distribution=1. if totals['initial'] else 0.,current_distribution=1. if totals['final'] else 0.)
 return fields

def build_select_formula_excel(df,cover_fields=None):
 import xlsxwriter
 if '_LINEAS_ORIGEN' in df.columns or '_LINEAS_DOCUMENTO' in df.columns:
  from excel_layouts import expand_document_rows
  original_attrs=df.attrs.copy();df=expand_document_rows(df,'SANTANDER SELECT');df.attrs=original_attrs
 rows=[]
 def clean(v):return None if v is None or (not isinstance(v,(list,dict)) and pd.isna(v)) else v
 for idx,(_,row) in enumerate(df.iterrows()):
  data=[clean(row.get(c)) for c in HEADERS]
  if data[0] is not None:data[0]=str(pd.Timestamp(data[0]).date())
  data[1]='' if data[1] is None else str(data[1]);rows.append(data)
 metadata={**(cover_fields or {}),**df.attrs.get('select_cover',{})}
 model=prepare_model(rows,metadata,df.attrs.get('reconciliation'))
 buffer=BytesIO();book=xlsxwriter.Workbook(buffer,{'in_memory':True});book.set_calc_mode('auto')
 sheets={name:book.add_worksheet(name) for name in ['MOVIMIENTOS','CARATULA','SALDOS_DIARIOS','CONTROL']}
 header=book.add_format({'bold':True,'bg_color':'#17324D','font_color':'white','text_wrap':True});money=book.add_format({'num_format':'#,##0.00'});date=book.add_format({'num_format':'yyyy-mm-dd'});text=book.add_format({'num_format':'@'});pct=book.add_format({'num_format':'0.00%','bg_color':'#EAF1F8'});inputfmt=book.add_format({'num_format':'#,##0.00','bg_color':'#FFF2CC'});autofmt=book.add_format({'num_format':'#,##0.00','bg_color':'#EAF1F8'})
 m=sheets['MOVIMIENTOS'];m.write_row(0,0,HEADERS,header);m.freeze_panes(1,0);m.set_row(0,36);m.set_column(0,1,14);m.set_column(2,2,80);m.set_column(3,5,16,money);m.set_column(6,10,18);m.set_column(11,12,65);m.set_column(13,14,17)
 for i,row in enumerate(model['rows'],1):
  for j,value in enumerate(row):
   if value is not None:m.write(i,j,value,date if j==0 else text if j==1 else inputfmt if j in [3,4] or (j==5 and i==1) else money if j==5 else None)
 c=sheets['CARATULA'];c.write_row(0,0,['CAMPO','VALOR','USO'],header);c.set_column(0,0,25);c.set_column(1,1,58);c.set_column(2,2,46);c.set_column(4,4,25);c.set_column(5,6,19,money);c.set_row(2,64)
 cover_text=book.add_format({'num_format':'@','bg_color':'#FFF2CC','text_wrap':True})
 cover_date=book.add_format({'num_format':'yyyy-mm-dd','bg_color':'#FFF2CC'})
 auto_money=book.add_format({'num_format':'#,##0.00','bg_color':'#EAF1F8'})
 auto_count=book.add_format({'num_format':'0','bg_color':'#EAF1F8'})
 for i,(key,value,label) in enumerate(model['metadata'],1):c.write(i,0,key);c.write(i,2,label);c.write(i,1,value,cover_date if key.startswith('period_') else cover_text if isinstance(value,str) else inputfmt) if value is not None else None
 c.write_row(2,4,['CONCEPTO','IMPORTE','MAGNITUD GRAFICA'],header)
 for i,label in enumerate(['Abonos','Otros cargos','Comisiones cobradas','Saldo final'],3):c.write(i,4,label)
 chart=book.add_chart({'type':'pie'});chart.add_series({'categories':'=CARATULA!$E$4:$E$7','values':'=CARATULA!$G$4:$G$7'});chart.set_title({'name':'Gráfica cuenta de cheques'});c.insert_chart('E10',chart,{'x_scale':1.1,'y_scale':1.1})
 d=sheets['SALDOS_DIARIOS'];d.write_row(0,0,['FECHA','DEPOSITOS','RETIROS','FLUJO NETO','SALDO AL CIERRE'],header);d.set_column(0,0,15,date);d.set_column(1,4,20,money)
 t=sheets['CONTROL'];t.write_row(0,0,['COMPARACION','ORIGINAL','EDITADO','CAMBIO'],header);t.set_column(0,0,28);t.set_column(1,3,20,money)
 for i,(label,key) in enumerate([('Saldo anterior','opening'),('Depósitos','deposits'),('Retiros','withdrawals'),('Saldo final','final'),('Operaciones','movement_count')],1):t.write(i,0,label);t.write(i,1,model['controls'][key],money)
 for name,cell,formula,value in model['formulas']:
  fmt=date if name=='SALDOS_DIARIOS' and cell.startswith('A') else autofmt if name=='MOVIMIENTOS' else money
  if name=='CARATULA':
   fmt=pct if cell in {f"B{model['keyrow'][k]}" for k in ['previous_distribution','current_distribution']} else auto_count if cell in {f"B{model['keyrow'][k]}" for k in ['days','operations','income_count','expense_count']} else auto_money
  sheets[name].write_formula(cell,formula,fmt,value)
 book.close();return buffer.getvalue()
