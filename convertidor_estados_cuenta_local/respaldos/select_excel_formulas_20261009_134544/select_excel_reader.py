"""Import physical Select Excel lines without flattening their edited descriptions."""
import io,json,re
import pandas as pd
from select_editable_document import parse_select_date
def read_select_excel(content):
    raw=pd.read_excel(io.BytesIO(content),dtype={'FOLIO':str})
    raw.columns=[str(c).strip().upper() for c in raw.columns]
    aliases={'DESCRIPCIÃ“N':'DESCRIPCION','CONCEPTO':'DESCRIPCION','DEPÃ“SITO':'DEPOSITO','CARGO':'RETIRO','ABONO':'DEPOSITO'}
    raw=raw.rename(columns={k:v for k,v in aliases.items() if k in raw and v not in raw})
    for col in ['FECHA','DESCRIPCION','DEPOSITO','RETIRO']:
        if col not in raw:raise ValueError('Falta la columna '+col)
    rows=[];detail=[];origins=[]
    def finish():
        if not rows:return
        rows[-1]['DESCRIPCION']='\n'.join(detail)
        rows[-1]['_LINEAS_DOCUMENTO']=json.dumps(detail,ensure_ascii=False)
        rows[-1]['_LINEAS_ORIGEN']=json.dumps(origins,ensure_ascii=False)
    for idx,row in raw.iterrows():
        desc='' if pd.isna(row.DESCRIPCION) else str(row.DESCRIPCION).strip()
        date=row.FECHA
        hasdate=pd.notna(date) and str(date).strip()!=''
        values={c:(float('nan') if c=='SALDO' else 0.) if pd.isna(row.get(c)) or str(row.get(c,'')).strip()=='' else pd.to_numeric(row.get(c),errors='raise') for c in ['DEPOSITO','RETIRO','SALDO']}
        if hasdate or desc.upper()=='SALDO ANTERIOR':
            finish();detail=desc.splitlines() or [''];origins=[]
            rows.append({'FECHA':parse_select_date(date) if hasdate else parse_select_date(raw.FECHA.dropna().iloc[0]),'FOLIO':'' if pd.isna(row.get('FOLIO')) else str(row.get('FOLIO','')).strip(),'DESCRIPCION':desc,**values,'ORIGEN':'ORIGINAL','HOJA_ORIGINAL':row.get('HOJA_ORIGINAL')})
        elif desc:
            if not rows:raise ValueError(f'LÃ­nea {idx+2}: continuaciÃ³n sin movimiento anterior.')
            if values['DEPOSITO'] or values['RETIRO'] or pd.notna(values['SALDO']):raise ValueError(f'LÃ­nea {idx+2}: contiene importes pero no tiene fecha.')
            detail.extend(desc.splitlines())
        if desc:
            for line in desc.splitlines():
                origins.append({'page':row.get('HOJA_ORIGINAL'),'y':row.get('Y_ORIGINAL'),'raw':line})
    finish()
    result=pd.DataFrame(rows)
    if result.empty:raise ValueError('El Excel estÃ¡ vacÃ­o.')
    return result
