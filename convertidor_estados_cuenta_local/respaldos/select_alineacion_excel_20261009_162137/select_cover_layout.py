"""Measured cover layout for the visibly didactic Select template."""
import math,hashlib,re
import fitz
RED=(234/255,0,0);GRAY=(196/255,)*3

def draw_cover(l,t):
    l.page(True);f=l.fields
    def text(x,y,s,size=8.5,bold=False,width=None,color=(0,0,0),center=False,key_override=None):
        key=key_override or ('cs' if bold and size==8.5 and 'cs' in l.fonts else 'cb' if bold else 'cr');s=str(s)
        # Calibration ratios measured against the source at 300 dpi.
        sx,sy=1.,1.
        if size==13 and bold:sy=.96;y+=.12
        elif size==10 and not bold:sx,sy=1.0587,1.078
        elif size==8.5 and not bold and key!='cp':
            numeric=bool(re.fullmatch(r'[\d.,%$ +\-]+',s))
            sx,sy=(1.025,.96) if numeric else (1.014,.93)
            if not numeric:y+=.12
        if size==10 and bold and s=='TOTAL':sx=1.025
        length=l.fonts[key].text_length(s,fontsize=size)*sx
        scale=(min(1.,width/max(length,.01)) if width else 1.)*sx
        if center and width:x+=(width-l.fonts[key].text_length(s,fontsize=size)*scale)/2
        l.p.insert_text((x,y),s,fontfile=str(l.paths[key]),fontname='CLC'+key,fontsize=size,color=color,morph=(fitz.Point(x,y),fitz.Matrix(scale,sy)))
    def amt(right,y,v,size=8.5,bold=False,prefix=''):
        s=prefix+f'{float(v):,.2f}';key='cb' if bold else 'cr'
        sx=1.025 if size==8.5 and not bold else 1.055 if size==10 and not bold else 1.
        text(right-l.fonts[key].text_length(s,fontsize=size)*sx,y,s,size,bold)
    def rule(x,y,r,b):l.p.draw_line((x,y),(r,b),color=(.65,)*3,width=.25)
    def title(y,label,icon):
        text(57,y,label,13,True,color=RED)
        if l.reference:
            rect=fitz.Rect(30,y-17,55,y+3)
            l.p.insert_image(rect,stream=l.reference[0].get_pixmap(matrix=fitz.Matrix(3,3),clip=rect).tobytes('png'))
        else:
            l.p.draw_rect((32,y-14,48,y+1),color=RED,width=.7)
            text(35,y-2,icon,10,True,color=RED)
    text(51,76,f.get('holder','PERSONA FICTICIA'),9,True,width=285)
    for j,line in enumerate(l.wrap(f.get('address','DOMICILIO FICTICIO'),275,8.5,'cb')[:4]):text(51,[84.36,91.2,99.6,108][j],line,8.5,True,width=275)
    text(382,77,'CODIGO DE CLIENTE NO.'+f.get('client','00000000'),10,True,width=200)
    text(382,92,'R.F.C.',10,True);text(435,92,f.get('rfc','XAXX010101000'),10,True,width=147)
    labels=[('MONEDA','MONEDA NACIONAL' if f.get('currency','MXN')=='MXN' else f.get('currency','MXN')),('SUCURSAL',f.get('branch','DIDACTICA')),('TELEFONO',f.get('phone','0000000000')),('PERIODO','DEL '+l.date(l.start)+' AL '+l.date(l.end)),('CORTE AL',l.date(l.end))]
    for j,(label,value) in enumerate(labels):text(382.2,104.16+j*9.6,label,8.5,True);text(435,104.16+j*9.6,value,8.5,True,width=147)
    # Educational decorative bars occupy the source barcode box; no banking payload.
    digest=hashlib.sha256(('DIDACTICO '+str(f.get('client',''))).encode()).digest()
    x=51
    for n in range(310):
        w=.45+(digest[n%32]%3)*.25
        if n%2==0:l.rect(x,110,x+w,142,(0,0,0))
        x+=w+.3
        if x>313:break
    text(51,149,'EJEMPLO DIDACTICO - CODIGO SIN VALIDEZ',6.5,width=262)
    title(196,'Resumen informativo.','R')
    text(30,215.5,'Resumen intereses y comisiones.',10)
    xs=[30,168.5,241,295.5,349,404,475.5,530,581.5]
    l.rect(30,218.3,581.5,241.7,GRAY)
    labels=['PRODUCTO','NUMERO DE\nCUENTA','INTERESES\nBRUTOS','ISR RETENIDO\n(0.90%)','INTERESES\nNETOS','COMISIONES\nCOBRADAS','GAT\nNOMINAL*','GAT\nREAL**']
    for j,label in enumerate(labels):
        parts=label.splitlines()
        for k,part in enumerate(parts):text(xs[j]+2,233 if len(parts)==1 else 227.5+k*11,part,8.5,True,width=xs[j+1]-xs[j]-4,center=True)
    products=[f.get('product','SUPERCUENTA CHEQUES-PAGO FIJO'),f.get('secondary_product','DINERO CRECIENTE SANTANDER')]
    accounts=[f.get('account','0000000000'),f.get('secondary_account','0000000001')]
    for j in range(2):
        y=253.68+j*11.28;text(30.36,y,products[j],8.5,width=110 if j==0 else 98,key_override='cp' if 'cp' in l.fonts else None);text(168.5,y,accounts[j],8.5,width=72.5,center=True)
        for col,key in zip(range(2,6),['interest','tax','net_interest','commissions']):
            value=float(f.get(key,0)) if not j else 0
            txt=f'{value:,.2f}';text(xs[col]+2,y,txt,8.5,width=xs[col+1]-xs[col]-4,center=True)
        if j:
            for col,key in [(6,'gat'),(7,'gat_real')]:text(xs[col]+2,y,f'{float(f.get(key,0)):.2f}%',8.5,True,width=xs[col+1]-xs[col]-4,center=True)
    for x in xs[1:-1]:rule(x,218.3,x,268.4)
    rule(30,268.4,581.5,268.4)
    text(30,295,'Resumen saldos.',10)
    sx=[30,157,242,411.5,580]
    l.rect(30,299.5,580,313.8,GRAY)
    for left,right,label in zip(sx,sx[1:],['PRODUCTO','NUMERO DE CUENTA','MES ANTERIOR','MES ACTUAL']):text(left+2,309,label,8.5,True,width=right-left-4,center=True)
    for x,w,label in [(242,84.5,'Monto'),(326.5,85,'% de distribucion'),(411.5,84.5,'Monto'),(496,84,'% de distribucion')]:text(x+2,324,label,8.5,True,width=w-4,center=True)
    for j in range(2):
        y=335.4+j*10.08;text(30.36,y,products[j],8.5,width=110 if j==0 else 98,key_override='cp' if 'cp' in l.fonts else None);text(157,y,accounts[j],8.5,width=85,center=True)
        amt(331,y,t['initial'] if not j else 0);text(371,y,'100.00%' if not j else '0.00%',8.5)
        amt(500,y,t['final'] if not j else 0);text(540,y+.72,'100.00%' if not j else '0.00%',8.5,width=39)
    l.rect(242,353,580,369,GRAY)
    text(30,365.04,'TOTAL',10,True);amt(334,365.04,t['initial'],10,True);text(367,365.04,'100.00%',10,True)
    amt(502,365.04,t['final'],10,True);text(538,365.04,'100.00%',10,True,width=40)
    for x in sx[1:-1]:rule(x,299.5,x,369)
    rule(30,369,580,369)
    title(397,'Cuenta de cheques.','$')
    l.rect(30,401.7,581.5,415.5,GRAY)
    text(49,412.5,products[0],8.5,True,width=159);text(215,412.5,accounts[0],8.5,True,width=76)
    text(296.5,412.5,'CUENTA CLABE: '+f.get('clabe','000000000000000000'),9,True,color=RED,width=280)
    text(48,425,'SUCURSAL  '+f.get('branch','DIDACTICA'),8.5,True,width=532)
    for y,label,value,bold in [(445,'Saldo promedio',f"{t['average']:,.2f}",False),(456,'Tasa bruta de interes anual',f"{float(f.get('annual_rate',0)):.4f}%",True),(467,'Dias del periodo',str(t['days']),False),(478,'No requiere saldo promedio minimo','',False)]:
        text(49.5,y,label,8.5,bold)
        # Right-align values in the whole numeric column; retain their glyph height.
        key='cb' if bold else 'cr';sx=1. if bold else 1.025
        text(282-l.fonts[key].text_length(value,fontsize=8.5)*sx,y,value,8.5,bold)
    for j,(label,key) in enumerate([('Saldo inicial','initial'),('+ Depositos','deposits'),('- Retiros','withdrawals'),('= Saldo final','final')]):
        y=445+j*11.5;text(339,y,label,10 if j==3 else 8.5,j==3);amt(534,y,t[key],10.5 if j==3 else 8.5,j==3)
    title(530,'Grafico cuenta de cheques.','G')
    text(51,559,products[0],11,True,width=258);text(51,574,'No. de cuenta '+accounts[0],11,True,width=258)
    text(51,589,'Saldo inicial de '+f"${t['initial']:,.2f}",10)
    rule(52,599,210,599)
    # Source graph separates commissions from other withdrawals to avoid double counting.
    commissions=float(f.get('commissions',0));other=t['withdrawals']-commissions
    values=[other,commissions,t['deposits'],t['final']]
    colors=[(251/255,223/255,220/255),(.85,.3,.3),(235/255,70/255,66/255),(234/255,0,0)]
    cx,cy,r=412,571.5,49;angle=-math.pi/2;total=sum(abs(v) for v in values)
    for value,color in zip(values,colors):
        delta=2*math.pi*abs(value)/max(total,1)
        if delta:
            n=max(3,int(delta*50));points=[fitz.Point(cx,cy)]+[fitz.Point(cx+r*math.cos(angle+delta*i/n),cy+r*math.sin(angle+delta*i/n)) for i in range(n+1)]
            l.p.draw_polyline(points,color=(1,1,1),fill=color,width=.25,closePath=True)
        angle+=delta
    text(320,518,'Saldo final  '+f"${t['final']:,.2f}",8.5,width=141)
    text(468,567,'Otros cargos  '+f'${other:,.2f}',8.5,width=112)
    text(287,603,'Abonos  '+f"${t['deposits']:,.2f}",8.5,width=90)
    text(431,634,'Comisiones cobradas  '+f'${commissions:,.2f}',8.5,width=149)
    for a,b in [((459,564),(467,563)),((374,596),(366,600)),((429,617),(432,625))]:rule(*a,*b)
    rule(30,680,581.5,680)
    rule(30,735,577,735)
    text(32,744,'*GAT NOMINAL Y REAL: ANTES DE IMPUESTOS.',8.5)
    text(30,755,'**GAT REAL: LA GAT REAL ES EL RENDIMIENTO QUE SE OBTENDRIA DESPUES DE DESCONTAR LA INFLACION.',8,width=475)
