"""Hojas de movimientos calibradas con las referencias proporcionadas."""

from __future__ import annotations

import os
import random
import re
from functools import lru_cache

import fitz
import pandas as pd

from excel_layouts import PHYSICAL_LINES_COLUMN, decode_physical_lines

W, H = 612.0, 792.0
GREEN = (0.00, 0.39, 0.24)
NAVY = (0.00, 0.22, 0.28)
BLUE = (0.00, 0.25, 0.52)
RED = (0.86, 0.00, 0.08)
GRAY = (0.88, 0.88, 0.88)
LIGHT_BLUE = (0.88, 0.93, 0.95)


def _font(name):
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fuentes", name)
    if not os.path.exists(path):
        raise FileNotFoundError(f"Falta {name} en la carpeta fuentes.")
    return path


def _first_existing(*paths):
    for path in paths:
        if path and os.path.exists(path):
            return path
    raise FileNotFoundError("No se encontro ninguna tipografia compatible.")


def _fonts_for_style(style):
    base = os.path.dirname(os.path.abspath(__file__))
    free = os.path.join(base, "fuentes", "libres")
    windows = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    if style in {"BANAMEX", "BANAMEX (JUN-JUL)"}:
        # Subconjuntos tipograficos incrustados en el estado MiCuenta de
        # referencia proporcionado por el usuario.
        regular = _font("BanamexDisplay-Regular-subset.otf")
        bold = _font("BanamexDisplay-Bold-subset.otf")
    elif style in {"AZTECA", "BBVA 2", "HSBC"}:
        regular = _first_existing(os.path.join(windows, "arial.ttf"), os.path.join(free, "Arimo-Variable.ttf"), _font("MYRIADPRO-COND.OTF"))
        bold = _first_existing(os.path.join(windows, "arialbd.ttf"), os.path.join(free, "SourceSans3-Bold.ttf"), _font("MYRIADPRO-BOLDCOND.OTF"))
    elif style in {"BBVA 1", "BBVA LARGO", "SANTANDER"}:
        regular = _first_existing(os.path.join(windows, "ARIALN.TTF"), os.path.join(free, "RobotoCondensed-Variable.ttf"), _font("MYRIADPRO-COND.OTF"))
        bold = _first_existing(os.path.join(windows, "ARIALNB.TTF"), os.path.join(free, "SourceSans3-Bold.ttf"), _font("MYRIADPRO-BOLDCOND.OTF"))
    elif style == "SCOTIABANK":
        # Perfil validado en EDO_CTA_MASTER_APP_V1_15_SCOTIA_FRANJA_MOVIMIENTO.
        # El archivo usa Arial Unicode MS Regular tanto en cuerpo como en titulos.
        regular = _first_existing(_font("ArialUnicodeMS.ttf"), os.path.join(windows, "arial.ttf"))
        bold = regular
    else:
        regular, bold = _font("MYRIADPRO-COND.OTF"), _font("MYRIADPRO-BOLDCOND.OTF")
    return regular, bold


def _money(value, symbol=True):
    try:
        number = float(value)
        if pd.isna(number):
            return ""
        return ("$" if symbol else "") + f"{number:,.2f}"
    except Exception:
        return ""


def _txt(page, x, y, text, fontfile, fontname, size=7.5, color=(0, 0, 0)):
    if text is not None and str(text) != "":
        page.insert_text((x, y), str(text), fontfile=fontfile, fontname=fontname, fontsize=size, color=color)


def _right(page, right, y, text, font, fontfile, fontname, size=7.5, color=(0, 0, 0)):
    text = str(text)
    _txt(page, right - font.text_length(text, fontsize=size), y, text, fontfile, fontname, size, color)


def _native_txt(page, x, y, text, fontname="helv", size=9, color=(0,0,0)):
    if text is not None and str(text)!="":
        page.insert_text((x,y),str(text),fontname=fontname,fontsize=size,color=color)


def _native_right(page, right, y, text, font, fontname="helv", size=9, color=(0,0,0)):
    text=str(text)
    _native_txt(page,right-font.text_length(text,fontsize=size),y,text,fontname,size,color)


def _wrap(text, width, font, size):
    result, current = [], ""
    for word in str(text or "").split():
        trial = word if not current else current + " " + word
        if font.text_length(trial, fontsize=size) <= width:
            current = trial
        else:
            if current:
                result.append(current)
            if font.text_length(word, fontsize=size) <= width:
                current = word
            else:
                piece = ""
                for char in word:
                    candidate = piece + char
                    if font.text_length(candidate, fontsize=size) <= width:
                        piece = candidate
                    else:
                        if piece:
                            result.append(piece)
                        piece = char
                current = piece
    if current:
        result.append(current)
    return result or [""]


def _physical_lines(text, width, font, size):
    """Build bank rows from semantic fields, not from one undifferentiated paragraph.

    The source converters that match the references treat continuation rows as
    first-class rows.  Splitting at the labels first preserves that behaviour;
    width wrapping is only used inside each physical field.
    """
    value = str(text or "").replace("\r", " ").replace("\n", " ").strip()
    parts = [
        part.strip()
        for part in re.split(
            r"\s+(?=(?:REF|RASTREO|CUENTA|EMISOR|RECEPTOR|ORDENANTE|BENEFICIARIO|CONCEPTO|HORA|CLAVE)\s*:?)",
            value,
            flags=re.IGNORECASE,
        )
        if part.strip()
    ]
    if not parts:
        parts = [value]

    # REF and CONCEPTO are one terminal detail row in the calibrated examples.
    merged = []
    for part in parts:
        if merged and part.upper().startswith("CONCEPTO") and merged[-1].upper().startswith("REF"):
            merged[-1] += " " + part
        else:
            merged.append(part)

    lines = []
    for part in merged:
        lines.extend(_wrap(part, width, font, size))
    return lines or [""]


_FIELD_RE = re.compile(
    r"\b(REF|RASTREO|CUENTA|EMISOR|RECEPTOR|ORDENANTE|BENEFICIARIO|CONCEPTO|HORA|CLAVE)\s*:?\s*",
    re.IGNORECASE,
)


def _semantic_fields(text):
    value = " ".join(str(text or "").replace("\r", " ").replace("\n", " ").split())
    matches = list(_FIELD_RE.finditer(value))
    if not matches:
        return value, {}
    concept = value[:matches[0].start()].strip()
    fields = {}
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(value)
        fields[match.group(1).upper()] = value[match.end():end].strip()
    return concept, fields


def _amount(row, key):
    try:
        value = float(row.get(key) or 0)
        return 0.0 if pd.isna(value) else value
    except Exception:
        return 0.0


def _movement_direction(row):
    """Clasifica por la columna contable: cargo=enviado, abono=recibido."""
    deposit = _amount(row, "DEPOSITO")
    withdrawal = _amount(row, "RETIRO")
    if deposit > 0 and withdrawal <= 0:
        return "RECIBIDO"
    if withdrawal > 0 and deposit <= 0:
        return "ENVIADO"
    return ""


def _is_transfer(row):
    concept, _ = _semantic_fields(row.get("DESCRIPCION", ""))
    return any(word in concept.upper() for word in ("SPEI", "TRANSFER", "TRASPASO"))


def _transfer_bank(index, row=None):
    banks = ["BBVA MEXICO", "SANTANDER", "BANAMEX", "BANORTE", "SCOTIABANK", "BANCO AZTECA", "BAJIO", "INBURSA"]
    if row is not None:
        assigned = row.get("BANCO_TRANSFERENCIA", "")
        if assigned is not None and not pd.isna(assigned) and str(assigned).strip():
            return str(assigned).strip().upper()
    return banks[int(index) % len(banks)]


def _digits_only(value, length=None):
    digits = re.sub(r"\D", "", str(value or ""))
    if length:
        digits = digits[-length:].zfill(length)
    return digits


def _alnum_only(value, length=None):
    cleaned = re.sub(r"[^A-Z0-9]", "", str(value or "").upper())
    if length:
        cleaned = cleaned[-length:].rjust(length, "0")
    return cleaned


def _date(value, fmt="%d/%m/%Y"):
    return "" if pd.isna(value) else pd.to_datetime(value).strftime(fmt).upper()


MONTHS_ES_SHORT = {
    1: "ENE", 2: "FEB", 3: "MAR", 4: "ABR", 5: "MAY", 6: "JUN",
    7: "JUL", 8: "AGO", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DIC",
}
MONTHS_ES_LONG = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio",
    7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre",
}


def _banamex_short_date(value):
    stamp = pd.to_datetime(value)
    return f"{stamp.day:02d}", MONTHS_ES_SHORT[stamp.month]


def _banamex_period_text(start, end):
    return (
        f"del {start.day:02d} de {MONTHS_ES_LONG[start.month]} "
        f"al {end.day:02d} de {MONTHS_ES_LONG[end.month]} del {end.year}"
    )


def _watermark(page, bold):
    h=page.rect.height
    page.draw_rect((42,h-16,570,h-4),color=(1,1,1),fill=(1,1,1))
    text="MATERIAL DIDACTICO - DATOS FICTICIOS"
    native=fitz.Font("hebo")
    x=(page.rect.width-native.text_length(text,fontsize=6.2))/2
    _native_txt(page,x,h-7,text,"hebo",6.2,(.35,.35,.35))


def _folio(seed, index):
    return f"{random.Random(seed + index).randrange(1000, 9999999):07d}"


def _azteca_lines(row,font):
    concept,fields=_semantic_fields(row.get("DESCRIPCION","")); upper=concept.upper()
    if upper=="SALDO ANTERIOR": return [concept]
    name=fields.get("ORDENANTE") or fields.get("BENEFICIARIO") or "PERSONA FICTICIA"
    account=_digits_only(fields.get("CUENTA"),18); trace=_digits_only(fields.get("RASTREO"),20); ref=_digits_only(fields.get("REF"),7)
    if "SPEI" in upper:
        first="TRANSFERENCIA SPEI A SU FAVOR" if float(row.get("DEPOSITO") or 0)>0 else "TRANSFERENCIA SPEI ENVIADA"
        role="NOM ORIGI" if float(row.get("DEPOSITO") or 0)>0 else "BENEFICIARIO"
        return [first,f"EMISOR: {_transfer_bank(0, row)}",f"CUENTA: {account}",f"{role}: {name}",f"RASTREO: {trace}",f"REF: {ref}","CONCEPTO: Transferencia interbancaria"]
    if "TRASPASO" in upper or "TRANSFERENCIA" in upper:
        return ["ABONO TRASPASO DE CUENTA",f"CUENTA: {account}","CLIENTE:",name.replace(" ","*"),f"FOLIO: {ref.zfill(9)}",f"TRASPASO {account[-14:]}"]
    if "DEPOSITO EN EFECTIVO" in upper:
        return [concept,"MEGA SAN JERONI","MO CALERAS"]
    return _wrap(row.get("DESCRIPCION",""),210,font,11.5)


def _azteca_header(page,identity,start,end,title,first_group,regular,bold,font):
    _txt(page,345,60,"Banco Azteca",bold,"Bold",24,GREEN)
    _txt(page,32,78,identity["NOMBRE"],bold,"Bold",11.5,GREEN)
    _txt(page,32,112,f"No. Cliente: {identity['CLIENTE']}",regular,"Regular",10)
    _txt(page,32,130,f"No. Cuenta: {identity['CUENTA']}",regular,"Regular",10)
    _txt(page,32,174,f"Periodo: del {start:%d/%m/%Y} al {end:%d/%m/%Y}",bold,"Bold",11)
    if first_group:
        page.draw_rect((32,210,580,236),color=GREEN,fill=GREEN); _txt(page,45,228,"Banco Azteca S. A. Institución de Banca Múltiple",bold,"Bold",12,(1,1,1))
        _txt(page,40,255,"Resumen Mensual de Movimientos de la Cuenta en Moneda Nacional",bold,"Bold",11,GREEN)
        sal=pd.to_numeric(identity.get("_balances",[]),errors="coerce")
        page.draw_rect((32,405,580,420),color=GREEN,fill=GREEN); _txt(page,45,417,title,bold,"Bold",11.5,(1,1,1))
        header_top=434
    else:
        page.draw_rect((32,195,580,210),color=GREEN,fill=GREEN); _txt(page,45,207,title,bold,"Bold",11.5,(1,1,1)); header_top=225
    page.draw_rect((32,header_top,580,header_top+28),color=GREEN,fill=GREEN)
    for x,label in [(55.57,"Fecha"),(196.67,"Concepto"),(336.53,"Monto de la Operación"),(475.27,"Lugar o Canal de\nOperación")]:
        for n,line in enumerate(label.split("\n")): _txt(page,x,header_top+18+n*12,line,bold,"Bold",11.5,(1,1,1))
    return header_top


def _azteca(df, identity, start, end, regular, bold, font):
    doc=fitz.open(); page_h=871.65; movement_no=0
    groups=[("Total Depósitos del mes",df[pd.to_numeric(df["DEPOSITO"],errors="coerce").fillna(0)>0],True),("Total Retiros del mes",df[pd.to_numeric(df["RETIRO"],errors="coerce").fillna(0)>0],False)]
    for group_no,(title,data,is_deposit) in enumerate(groups):
        page=None; top=0; first_group_page=True
        for idx,row in data.iterrows():
            lines=_azteca_lines(row,font)
            row_h=max(22.68,len(lines)*12.86+(0.71 if len(lines)>1 else 0))
            if page is None or top+row_h>805:
                page=doc.new_page(width=W,height=page_h)
                is_first_document_page=(group_no==0 and first_group_page)
                header_top=_azteca_header(page,identity,start,end,title,is_first_document_page,regular,bold,font)
                _watermark(page,bold)
                if is_first_document_page: top=477.8
                elif first_group_page: top=267.0
                else: top=109.3
                first_group_page=False
            y=top+11.13
            fill=(.95,.95,.95) if movement_no%2 else (1,1,1)
            page.draw_rect((32,top,580,top+row_h),color=fill,fill=fill,overlay=False)
            _txt(page,43.42,y,_date(row["FECHA"]),regular,"Regular",11.5)
            for n,line in enumerate(lines): _txt(page,112.49,y+n*12.86,line,regular,"Regular",11.5)
            amount=row["DEPOSITO"] if is_deposit else row["RETIRO"]
            _txt(page,333.88,y,("(+) " if is_deposit else "(-) ")+_money(amount),regular,"Regular",11.5)
            channel="SPEI" if "SPEI" in str(row["DESCRIPCION"]).upper() else ("VENTANILLA" if "EFECTIVO" in str(row["DESCRIPCION"]).upper() else "BANCA ELECTRONICA")
            _right(page,566.65,y,channel,font,regular,"Regular",11.5)
            top+=row_h; movement_no+=1
    return doc


def _banamex_lines(row, movement_no):
    concept,fields=_semantic_fields(row.get("DESCRIPCION","")); upper=concept.upper()
    if upper=="SALDO ANTERIOR": return [concept]
    if str(row.get("ORIGEN", "")).upper() == "ORIGINAL" and row.get(PHYSICAL_LINES_COLUMN) is not None:
        preserved = decode_physical_lines(row, "BANAMEX MICUENTA")
        if preserved:
            # En MiCuenta el deposito y el saldo se imprimen en el renglon
            # inmediatamente posterior al ultimo texto del ingreso. La
            # extraccion del PDF no conserva ese renglon porque esta vacio en
            # la columna CONCEPTO; se repone aqui sin alterar el texto original.
            deposit = pd.to_numeric(pd.Series([row.get("DEPOSITO", 0)]), errors="coerce").fillna(0).iloc[0]
            if float(deposit) > 0.005 and preserved[-1] != "":
                return preserved + [""]
            return preserved
    ref=_digits_only(fields.get("REF"),7) or _digits_only(row.get("DESCRIPCION"),7)
    name=fields.get("ORDENANTE") or fields.get("BENEFICIARIO") or "CARLOS MARTINEZ LOPEZ"
    account=_digits_only(fields.get("CUENTA"),18)
    trace=_alnum_only(fields.get("RASTREO"),32)
    transfer_bank=_transfer_bank(movement_no, row)
    detail=fields.get("CONCEPTO") or "TRANSFERENCIA"
    if str(row.get("FORMATO_BANAMEX_MICUENTA", "")).upper() == "TRUE":
        account = _digits_only(row.get("CUENTA_BANAMEX"), 18)
        ref = _digits_only(row.get("REFERENCIA_BANAMEX"), 7)
        trace = _alnum_only(row.get("RASTREO_BANAMEX"))
        if float(row.get("DEPOSITO") or 0)>0:
            return [
                f"PAGO RECIBIDO DE {transfer_bank}",
                f"POR ORDEN DE {name}",
                "CTA.ORDENANTE",
                f"{account} REF.{ref}",
                "RASTREO:",
                trace,
                "",
            ]
        return [
            f"PAGO INTERBANCARIO A {transfer_bank}",
            f"AL BENEF. {name}",
            "(DATO NO VERIFICADO POR ESTA",
            "INSTITUCION) CTA.BENEFICIARIO",
            f"{account} CLAVE",
            f"RASTREO {trace}",
            f"REF. {ref}",
            "MISMO DIA",
        ]
    if any(key in upper for key in ("SPEI","TRANSFERENCIA","TRASPASO")):
        if float(row.get("DEPOSITO") or 0)>0:
            lines = [
                f"PAGO RECIBIDO DE {transfer_bank}", f"POR ORDEN DE {name}",
                "CTA.ORDENANTE", f"{account} REF.{ref}", detail, "RASTREO:", trace,
            ]
            # Los depositos nuevos deben reservar, igual que los originales,
            # un renglon fisico vacio antes de imprimir DEPOSITO y SALDO.
            if str(row.get("ORIGEN", "")).upper() == "FICTICIO":
                lines.append("")
            return lines
        bank_lines = transfer_bank.split(" ", 1)
        first_bank_line = f"PAGO INTERBANCARIO A {bank_lines[0]}"
        second_bank_line = f"{bank_lines[1]} AL BENEF." if len(bank_lines) > 1 else "AL BENEF."
        return [first_bank_line,second_bank_line,name,"NO VERIFICADO POR ESTA","INSTITUCION) CTA.BENEFICIARIO",f"{account} CLAVE","SUC 0","CAJA 0 AUT 0 HORA 0: 0"]
    lines = [concept,f"CAJA 0071 AUT {ref}","HORA 03:34 SUC 0511"]
    if (
        str(row.get("ORIGEN", "")).upper() == "FICTICIO"
        and float(row.get("DEPOSITO") or 0) > 0
    ):
        lines.append("")
    return lines


def _banamex_page(page,df,identity,start,end,page_no,first,regular,bold,font,bold_font,jun_jul=False):
    _txt(page,24.9,42,"Banamex",bold,"BanamexBold",20,(.8,0,.18)); _right(page,579,43,f"Página {page_no}",font,regular,"BanamexRegular",9,NAVY)
    if first:
        page.draw_rect((18,84,579,740),color=NAVY,width=.7)
        page.draw_rect((24.9,99,285,119),color=NAVY,fill=NAVY); _txt(page,31,114,"MiCuenta",bold,"BanamexBold",13,(1,1,1))
        _txt(page,24.9,136,"Resumen",regular,"BanamexRegular",9,NAVY); _txt(page,138,136,_banamex_period_text(start,end),regular,"BanamexRegular",9)
        dep=float(pd.to_numeric(df["DEPOSITO"],errors="coerce").fillna(0).sum()); ret=float(pd.to_numeric(df["RETIRO"],errors="coerce").fillna(0).sum()); sal=pd.to_numeric(df["SALDO"],errors="coerce").dropna(); opening=float(sal.iloc[0]); closing=float(sal.iloc[-1])
        for n,(label,value) in enumerate([("Saldo anterior",opening),("(+) Depósitos",dep),("(-) Retiros/Otros cargos",ret),("Saldo al Corte",closing)]):
            yy=173+n*22; page.draw_rect((24.9,yy-14,285,yy+3),color=LIGHT_BLUE,fill=LIGHT_BLUE); _txt(page,31,yy,label,regular,"BanamexRegular",10,NAVY); _right(page,282,yy,_money(value,False),font,regular,"BanamexRegular",10)
        if jun_jul:
            page.draw_rect((18.8,341.0,566.5,358.02),color=LIGHT_BLUE,fill=LIGHT_BLUE)
            _txt(page,24.9,353.4,"Detalle de Operaciones",bold,"BanamexBold",11)
            _txt(page,142.0,353.4,"En pesos Moneda Nacional",regular,"BanamexRegular",9)
        header_y=379.4 if jun_jul else 369.6
    else:
        _txt(page,24.9,81,identity["NOMBRE"],regular,"BanamexRegular",9)
        if jun_jul:
            page.draw_rect((18.8,108.0,566.5,124.32),color=LIGHT_BLUE,fill=LIGHT_BLUE)
            _txt(page,24.9,119.8,"Detalle de Operaciones",bold,"BanamexBold",11)
            _txt(page,142.0,119.8,"En pesos Moneda Nacional",regular,"BanamexRegular",9)
        else:
            _txt(page,24.9,106,"Detalle de Operaciones",bold,"BanamexBold",11,NAVY)
        header_y=145.7 if jun_jul else 135.9
    for x,label in [(24.9,"FECHA"),(81.6,"CONCEPTO"),(356.1,"RETIROS"),(429.3,"DEPÓSITOS"),(532.8,"SALDO")]: _txt(page,x,header_y,label,bold,"BanamexBold",9,NAVY)


def _banamex(df, identity, start, end, regular, bold, font, jun_jul=False):
    month_tokens={"ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"}
    for movement_no,(_,row) in enumerate(df.iterrows()):
        concept_lines=[str(value).strip().upper() for value in _banamex_lines(row,movement_no)]
        embedded_date=any(
            re.fullmatch(r"\d{1,2}", concept_lines[index] or "")
            and concept_lines[index+1] in month_tokens
            for index in range(max(0,len(concept_lines)-1))
        )
        if embedded_date:
            raise ValueError(
                "El PDF Banamex no se genero porque el lector encontro varios movimientos "
                "fusionados: una fecha quedo dentro de la columna CONCEPTO. Vuelve a cargar "
                "el PDF base para separarlos antes de exportar."
            )
    doc=fitz.open(); page=None; y=0; page_no=1; first=True; page_is_first=False; rows_page=0; bold_font=fitz.Font(fontfile=bold)
    row_height=11.65
    for movement_no,(idx,row) in enumerate(df.iterrows()):
        lines=_banamex_lines(row,movement_no)
        for line_index,line in enumerate(lines):
            if jun_jul:
                page_capacity=29 if page_is_first else 49
            else:
                page_capacity=30 if page_is_first else 50
            if page is None or rows_page>=page_capacity:
                page_is_first=first
                page=doc.new_page(width=W,height=H); _banamex_page(page,df,identity,start,end,page_no,page_is_first,regular,bold,font,bold_font,jun_jul)
                _watermark(page,bold)
                if jun_jul:
                    grid_top=384.95 if page_is_first else 151.25
                    y=grid_top+9.45
                    page.draw_line((18.8,grid_top),(566.5,grid_top),color=(.35,.35,.35),width=.35)
                else:
                    y=384.60 if page_is_first else 150.49
                rows_page=0; page_no+=1; first=False
            is_first=line_index==0; is_last=line_index==len(lines)-1
            if is_first:
                day_text, month_text = _banamex_short_date(row["FECHA"])
                _txt(page,24.94,y,day_text,regular,"BanamexRegular",9); _txt(page,41.95,y,month_text,regular,"BanamexRegular",9)
            _txt(page,81.63,y,line,regular,"BanamexRegular",9)
            if is_last:
                _right(page,390.57,y,_money(row["RETIRO"],False) if float(row["RETIRO"] or 0) else "",font,regular,"BanamexRegular",9)
                _right(page,477.95,y,_money(row["DEPOSITO"],False) if float(row["DEPOSITO"] or 0) else "",font,regular,"BanamexRegular",9)
                _right(page,560.74,y,"" if pd.isna(row["SALDO"]) else _money(row["SALDO"],False),font,regular,"BanamexRegular",9)
            # La variante JUN-JUL imprime la retícula horizontal de 29/49
            # posiciones; el perfil Banamex anterior conserva su salida.
            y+=row_height; rows_page+=1
            if jun_jul:
                page.draw_line((18.8,grid_top+rows_page*row_height),(566.5,grid_top+rows_page*row_height),color=(.35,.35,.35),width=.35)
    return doc


def _assert_banamex_render_integrity(doc, df):
    """Comprueba que el PDF conserve cada fila calculada sin desplazarla.

    La auditoria/Excel es la fuente de verdad del generador enlazado. Esta
    validacion se ejecuta sobre el PDF ya dibujado y verifica, en el mismo
    orden, la fecha, el primer renglon del concepto y la terna
    retiro/deposito/saldo. Si una columna, una fecha o un importe se mueve de
    renglon, el PDF no se entrega.
    """
    rendered_headers = []
    rendered_concept_lines = []
    rendered_amounts = []
    for page_index, page in enumerate(doc):
        table_top = 360.0 if page_index == 0 else 125.0
        words = [
            word for word in page.get_text("words", sort=True)
            if table_top <= float(word[1]) < 740.0
        ]
        groups = []
        for word in words:
            group = next(
                (item for item in reversed(groups) if abs(item[0] - float(word[1])) <= 1.5),
                None,
            )
            if group is None:
                group = [float(word[1]), []]
                groups.append(group)
            group[1].append(word)
        for _, line_words in groups:
            concept = " ".join(
                word[4] for word in sorted(line_words, key=lambda item: item[0])
                if 75 <= word[0] < 335
            ).strip()
            if concept and concept.upper() != "CONCEPTO":
                rendered_concept_lines.append(concept)
            day = next(
                (word[4] for word in line_words if 20 <= word[0] < 41 and re.fullmatch(r"\d{1,2}", word[4])),
                None,
            )
            month = next(
                (word[4] for word in line_words if 41 <= word[0] < 75 and re.fullmatch(r"[A-Z]{3}", word[4])),
                None,
            )
            if day and month:
                rendered_headers.append((day.zfill(2), month, concept))

            columns = []
            for left, right in ((335, 415), (415, 505), (505, 580)):
                values = [
                    word[4] for word in line_words
                    if left <= word[0] < right and re.fullmatch(r"[\d,]+\.\d{2}", word[4])
                ]
                columns.append(values[-1] if values else "")
            if any(columns):
                rendered_amounts.append(tuple(columns))

    expected_headers = []
    expected_concept_lines = []
    expected_amounts = []
    for movement_no, (_, row) in enumerate(df.iterrows()):
        day, month = _banamex_short_date(row["FECHA"])
        lines = _banamex_lines(row, movement_no)
        expected_headers.append((day, month, lines[0] if lines else ""))
        expected_concept_lines.extend(line for line in lines if str(line).strip())
        expected_amounts.append((
            _money(row.get("RETIRO", 0), False) if float(row.get("RETIRO", 0) or 0) else "",
            _money(row.get("DEPOSITO", 0), False) if float(row.get("DEPOSITO", 0) or 0) else "",
            "" if pd.isna(row.get("SALDO")) else _money(row.get("SALDO"), False),
        ))

    if len(rendered_headers) != len(expected_headers):
        raise ValueError(
            "El PDF Banamex no se genero: la cantidad de fechas dibujadas "
            f"({len(rendered_headers)}) no coincide con las filas calculadas ({len(expected_headers)})."
        )
    if len(rendered_amounts) != len(expected_amounts):
        raise ValueError(
            "El PDF Banamex no se genero: la cantidad de renglones monetarios "
            f"({len(rendered_amounts)}) no coincide con las filas calculadas ({len(expected_amounts)})."
        )
    if rendered_concept_lines != expected_concept_lines:
        mismatch = next(
            (
                index for index, (expected, rendered) in enumerate(
                    zip(expected_concept_lines, rendered_concept_lines), start=1
                )
                if expected != rendered
            ),
            min(len(expected_concept_lines), len(rendered_concept_lines)) + 1,
        )
        expected = expected_concept_lines[mismatch - 1] if mismatch <= len(expected_concept_lines) else "<fin>"
        rendered = rendered_concept_lines[mismatch - 1] if mismatch <= len(rendered_concept_lines) else "<fin>"
        raise ValueError(
            f"El PDF Banamex no se genero: el concepto se desplazo o quedo incompleto en el renglon fisico {mismatch}. "
            f"Esperado '{expected}'; dibujado '{rendered}'."
        )
    for index, (expected, rendered) in enumerate(zip(expected_headers, rendered_headers), start=1):
        if expected != rendered:
            raise ValueError(
                f"El PDF Banamex no se genero: fecha o concepto desplazado en el movimiento {index}. "
                f"Esperado {expected}; dibujado {rendered}."
            )
    for index, (expected, rendered) in enumerate(zip(expected_amounts, rendered_amounts), start=1):
        if expected != rendered:
            raise ValueError(
                f"El PDF Banamex no se genero: retiro/deposito/saldo desplazado en el movimiento {index}. "
                f"Esperado {expected}; dibujado {rendered}."
            )


def _bbva_financial_area(page, df, identity, start, end, regular, bold, font, profile):
    dep=pd.to_numeric(df["DEPOSITO"],errors="coerce").fillna(0)
    ret=pd.to_numeric(df["RETIRO"],errors="coerce").fillna(0)
    sal=pd.to_numeric(df["SALDO"],errors="coerce").dropna()
    opening=float(sal.iloc[0]) if len(sal) else 0.0; closing=float(sal.iloc[-1]) if len(sal) else 0.0
    _txt(page,28,60,"BBVA",bold,"Bold",27,BLUE)
    _right(page,570,30,"Estado de Cuenta",font,bold,"Bold",11)
    _right(page,570,45,"Libretón Premium" if profile==1 else "TARJETA BASICA BBVA",font,regular,"Regular",9)
    _right(page,570,59,"PAGINA 1",font,bold,"Bold",8)
    _txt(page,42,108,identity["NOMBRE"],regular,"Regular",9.5)
    _txt(page,42,124,identity["DOMICILIO"],regular,"Regular",8.5)
    page.draw_rect((320,70,570,160),color=(.2,.2,.2),width=.55)
    page.draw_rect((320,70,435,160),color=(.75,.75,.75),fill=(.75,.75,.75))
    info=[("Periodo",f"DEL {start:%d/%m/%Y} AL {end:%d/%m/%Y}"),("Fecha de Corte",f"{end:%d/%m/%Y}"),("No. de Cuenta",identity["CUENTA"]),("No. de Cliente",identity["CLIENTE"]),("R.F.C",identity["RFC"]),("No. Cuenta CLABE",identity["CLABE"])]
    for n,(label,value) in enumerate(info):
        yy=84+n*14; _txt(page,325,yy,label,regular,"Regular",8.5); _right(page,565,yy,value,font,regular,"Regular",8.5)
        if n<len(info)-1: page.draw_line((320,yy+4),(570,yy+4),color=(.25,.25,.25),width=.35)
    _txt(page,20,222,"Información Financiera",regular,"Regular",13)
    _right(page,570,222,"MONEDA NACIONAL",font,regular,"Regular",12)
    page.draw_line((18,226),(570,226),color=(0,0,0),width=1.2)
    for left,right,title in [(18,300,"Rendimiento"),(310,570,"Comportamiento")]:
        page.draw_rect((left,232,right,250),color=(.7,.7,.7),fill=(.7,.7,.7)); _txt(page,left+4,245,title,bold,"Bold",8.5)
    daily = df.copy()
    daily["_date"] = pd.to_datetime(daily["FECHA"], errors="coerce").dt.normalize()
    end_balances = daily.groupby("_date", sort=True)["SALDO"].last()
    day_index = pd.date_range(start, end, freq="D")
    average = end_balances.reindex(end_balances.index.union(day_index)).sort_index().ffill().reindex(day_index).fillna(opening).mean()
    left_rows=[("Saldo Promedio",float(average)),("Días del Periodo",(end-start).days+1),("Tasa Bruta Anual",0),("Intereses a Favor (+)",0),("Total Comisiones",0)]
    right_rows=[("Saldo Anterior",opening),(f"Depósitos / Abonos (+)  {int((dep>0).sum())}",float(dep.sum())),(f"Retiros / Cargos (-)  {int((ret>0).sum())}",float(ret.sum())),("Saldo Final (+)",closing),("Saldo Promedio Mínimo Mensual",0)]
    for n in range(5):
        yy=265+n*20
        for left,right,(label,value) in [(18,300,left_rows[n]),(310,570,right_rows[n])]:
            page.draw_rect((left,yy-13,right,yy+5),color=(.25,.25,.25),width=.35)
            _txt(page,left+5,yy,label,regular,"Regular",8.3)
            val=f"{int(value)}" if isinstance(value,int) else _money(value,False)
            _right(page,right-4,yy,val,font,regular,"Regular",8.3)
    _txt(page,18,505,"Detalle de Movimientos Realizados",bold,"Bold",10)
    page.draw_line((18,511),(590,511),color=(0,0,0),width=1.2)


def _bbva1_lines(row):
    concept,fields=_semantic_fields(row.get("DESCRIPCION",""))
    # "CONCEPTO" dentro del detalle no inicia un segundo campo.
    detail_match = re.search(r'CONCEPTO\s*:\s*(.*?)(?=\s+(?:REF|RASTREO|CUENTA|ORDENANTE|BENEFICIARIO)\s*:|$)', str(row.get("DESCRIPCION", "")), re.I)
    if detail_match:
        fields["CONCEPTO"] = detail_match.group(1)
    if concept.upper()=="SALDO ANTERIOR": return [concept]
    physical = row.get(PHYSICAL_LINES_COLUMN)
    if physical is not None and not pd.isna(physical):
        return decode_physical_lines(row, "BBVA")
    # Excel sin metadatos: recuperar los renglones de un concepto BBVA completo.
    match = re.search(r"\s+(?=(?:\d{10}\s+\d{3}\s+|\w*BNET\s+|REFBNTC))", concept)
    if match:
        title, details = concept[:match.start()], concept[match.end():]
        return [title] + re.split(r"\s+(?=(?:\d{18,20}\b|MBAN\w+\b|\d{8}\w{12,}\b))", details)
    ref=_digits_only(fields.get("REF"),10) or _digits_only(row.get("DESCRIPCION"),10)
    if "SPEI" in concept.upper() or "TRANSFERENCIA" in concept.upper() or "TRASPASO" in concept.upper():
        name=fields.get("ORDENANTE") or fields.get("BENEFICIARIO") or "PERSONA FICTICIA"
        account=_digits_only(fields.get("CUENTA"),20)
        trace=fields.get("RASTREO") or "00000000000000000000000000"
        detail=fields.get("CONCEPTO") or "Sin informacion"
        bank=str(row.get("BANCO_TRANSFERENCIA", "BANORTE")).upper()
        code={"SANTANDER":"014", "BBVA MEXICO":"012", "BANAMEX":"002", "SCOTIABANK":"044", "HSBC":"021", "BANORTE":"072"}.get(bank,"072")
        return [concept, f"{ref} {code}  {pd.to_datetime(row['FECHA']):%y%m%d} {detail}", account, trace, name]
    return [concept, f"{ref}  CUENTA: BMOV"]


def _bbva2_lines(row):
    concept,fields=_semantic_fields(row.get("DESCRIPCION",""))
    if concept.upper()=="SALDO ANTERIOR": return [concept]
    ref=_digits_only(fields.get("REF"),10) or _digits_only(row.get("DESCRIPCION"),10)
    if "SPEI" in concept.upper() or "TRANSFERENCIA" in concept.upper() or "TRASPASO" in concept.upper():
        name=fields.get("ORDENANTE") or fields.get("BENEFICIARIO") or "PERSONA FICTICIA"
        account=_digits_only(fields.get("CUENTA"),18)
        trace=_digits_only(fields.get("RASTREO"),20)
        detail=(fields.get("CONCEPTO") or "TRANSFERENCIA").lower()
        return [concept, detail, account, trace, name]
    return [concept, f"BNET {ref}"]


def _bbva_page(page, df, identity, start, end, regular, bold, font, profile, first):
    if first:
        _bbva_financial_area(page,df,identity,start,end,regular,bold,font,profile)
        header_y=526
    else:
        _txt(page,28,60,"BBVA",bold,"Bold",27,BLUE)
        _right(page,570,30,"Estado de Cuenta",font,bold,"Bold",11)
        _right(page,570,45,"Libretón Premium" if profile==1 else "TARJETA BASICA BBVA",font,regular,"Regular",9)
        page.draw_rect((330,75,570,111),color=(.25,.25,.25),width=.55); page.draw_rect((330,75,465,111),color=GRAY,fill=GRAY)
        _txt(page,335,90,"No. de Cuenta",regular,"Regular",7); _right(page,565,90,identity["CUENTA"],font,regular,"Regular",8)
        _txt(page,335,105,"No. de Cliente",regular,"Regular",7); _right(page,565,105,identity["CLIENTE"],font,regular,"Regular",8)
        header_y=137
    if profile==1:
        labels=[(16,"FECHA\nOPER"),(51,"LIQ"),(81.78,"DESCRIPCION"),(216.16,"REFERENCIA"),(370.81,"CARGOS"),(428.29,"ABONOS"),(476.99,"OPERACION"),(544.94,"LIQUIDACION")]
    else:
        labels=[(14.53,"FECHA\nOPER"),(59.15,"LIQ"),(106.18,"DESCRIPCION"),(318.53,"REFERENCIA"),(382,"CARGOS"),(425,"ABONOS"),(475,"OPERACION"),(550,"LIQUIDACION")]
    for x,label in labels:
        for n,line in enumerate(label.split("\n")): _txt(page,x,header_y+n*10,line,regular,"Regular",8 if profile==2 else 10)


def _bbva(df, identity, start, end, regular, bold, italic, font, profile):
    doc=fitz.open(); page=None; y=0; first_page=True; page_bottom=0
    step=12.0 if profile==1 else 11.18
    normal_start=167.0 if profile==1 else 165.4
    first_start=555.0 if profile==1 else 553.0
    # La variante 2 reserva mucho mas espacio inferior en la caratula. En las
    # referencias suministradas caben como maximo unas seis filas fisicas
    # (normalmente 2 o 3 movimientos) antes de continuar en la pagina siguiente.
    first_bottom=750.0 if profile==1 else 625.0
    normal_bottom=750.0
    for idx,row in df.iterrows():
        lines=_bbva1_lines(row) if profile==1 else _bbva2_lines(row)
        wrapped = []
        for line in lines:
            piece = ""
            for char in str(line):
                if piece and font.text_length(piece + char, fontsize=9.5) > 272:
                    split = piece.rfind(" ")
                    if split > 0:
                        wrapped.append(piece[:split]); piece = piece[split+1:] + char
                    else:
                        wrapped.append(piece); piece = char
                else:
                    piece += char
            wrapped.append(piece)
        lines = wrapped
        needed=len(lines)*step
        if page is None or y+needed>page_bottom:
            creating_first=first_page
            page=doc.new_page(width=W,height=H)
            _bbva_page(page,df,identity,start,end,regular,bold,font,profile,creating_first)
            _watermark(page,bold)
            y=first_start if creating_first else normal_start
            page_bottom=first_bottom if creating_first else normal_bottom
            first_page=False
        for line_index,line in enumerate(lines):
            first=line_index==0
            date=_date(row["FECHA"],"%d/%b") if first else ""
            if profile==1:
                liquid = row.get("FECHA_LIQ_BBVA", date)
                if pd.isna(liquid): liquid = date
                _txt(page,16,y,date,regular,"Regular",9.5); _txt(page,51,y,str(liquid) if first else "",regular,"Regular",9.5)
                _txt(page,85.5,y,line,italic if line_index==1 else regular,"Italic" if line_index==1 else "Regular",9.5)
                if first:
                    _right(page,416.05,y,_money(row["RETIRO"],False) if float(row["RETIRO"] or 0) else "",font,regular,"Regular",9.5)
                    _right(page,464.59,y,_money(row["DEPOSITO"],False) if float(row["DEPOSITO"] or 0) else "",font,regular,"Regular",9.5)
                    if not pd.isna(row["SALDO"]):
                        _right(page,526.95,y,_money(row["SALDO"],False),font,regular,"Regular",9.5); _right(page,598.54,y,_money(float(row["SALDO"])+(0. if pd.isna(row.get("AJUSTE_LIQ_BBVA",0.)) else float(row.get("AJUSTE_LIQ_BBVA",0.))),False),font,regular,"Regular",9.5)
            else:
                _txt(page,14.53,y,date,regular,"Regular",10); _txt(page,59.15,y,date,regular,"Regular",10)
                _txt(page,106.18,y,line,regular,"Regular",9.5 if first else 9)
                if line_index==1:
                    ref=_digits_only(_semantic_fields(row.get("DESCRIPCION",""))[1].get("REF"),10)
                    _txt(page,318.53,y,f"Referencia {ref}",regular,"Regular",9.5)
                if first:
                    _right(page,417.37,y,_money(row["RETIRO"],False) if float(row["RETIRO"] or 0) else "",font,regular,"Regular",8.5)
                    _right(page,457.84,y,_money(row["DEPOSITO"],False) if float(row["DEPOSITO"] or 0) else "",font,regular,"Regular",8.5)
                    if not pd.isna(row["SALDO"]):
                        _right(page,512.92,y,_money(row["SALDO"],False),font,regular,"Regular",8.5); _right(page,594.5,y,_money(row["SALDO"],False),font,regular,"Regular",8.5)
            y+=step
    for number, sheet in enumerate(doc, 1):
        sheet.draw_rect((490, 48, 605, 65), color=None, fill=(1,1,1), overlay=True)
        _right(sheet,598,59,f"PAGINA {number} / {len(doc)}",font,bold,"Bold",8)
    return doc


def _santander_lines(row, idx, font):
    concept, fields = _semantic_fields(row.get("DESCRIPCION", ""))
    if concept.upper() == "SALDO ANTERIOR":
        return [concept]
    if _is_transfer(row):
        direction = _movement_direction(row)
        name = fields.get("ORDENANTE") or fields.get("BENEFICIARIO") or "PERSONA FICTICIA"
        account = _digits_only(fields.get("CUENTA"), 18)
        trace = _digits_only(fields.get("RASTREO"), 22)
        reference = _digits_only(fields.get("REF"), 7)
        seed = int((reference or "1234567")[-7:])
        stamp = f"{seed % 24:02d}:{(seed // 24) % 60:02d}:{(seed // 1440) % 60:02d}"
        if direction == "RECIBIDO":
            return _wrap(f"ABONO TRANSF RAPIDA SANTANDER Transferencia a {name.title()}",250,font,7.1)
        detail = fields.get("CONCEPTO") or "TRANSFERENCIA INTERBANCARIA"
        return [
            f"PAGO TRANSFERENCIA SPEI HORA {stamp}",
            f"ENVIADO {_transfer_bank(idx, row)}",
            f"A LA CUENTA {account}",
            f"AL CLIENTE {name} (*)",
            "(1) DATO NO VERIFICADO POR ESTA INSTITUCION",
            f"CLAVE DE RASTREO {trace}",
            f"REF {reference}",
            f"CONCEPTO {detail}",
        ]
    return _physical_lines(row.get("DESCRIPCION", ""), 250, font, 7.1)


def _santander_first_header(page, df, identity, start, end, regular, bold, font):
    dep=float(pd.to_numeric(df["DEPOSITO"],errors="coerce").fillna(0).sum())
    ret=float(pd.to_numeric(df["RETIRO"],errors="coerce").fillna(0).sum())
    sal=pd.to_numeric(df["SALDO"],errors="coerce").dropna()
    opening=float(sal.iloc[0]) if len(sal) else 0.0
    closing=float(sal.iloc[-1]) if len(sal) else 0.0
    _txt(page,32,47,"S",bold,"Bold",28,RED)
    _txt(page,55,47,"Santander",bold,"Bold",25,RED)
    _right(page,570,55,"ESTADO DE CUENTA",font,bold,"Bold",14,(.18,.18,.18))
    page.draw_line((32,100),(580,100),color=(.4,.4,.4),width=.7)
    _txt(page,32,116,identity["NOMBRE"],bold,"Bold",7.5)
    _right(page,570,116,f"CODIGO DE CLIENTE NO.{identity['CLIENTE']}",font,bold,"Bold",7.5)
    _right(page,570,129,f"PERIODO DEL {start:%d-%b-%Y} AL {end:%d-%b-%Y}",font,bold,"Bold",7.5)
    _txt(page,32,153,"DineroCreciente",regular,"Regular",14,RED)
    _txt(page,72,169,"S A N T A N D E R",bold,"Bold",8.5,RED)
    page.draw_rect((32,174,580,191),color=(.75,.75,.75),fill=(.75,.75,.75))
    _txt(page,72,186,f"DINERO CRECIENTE SANTANDER {identity['CUENTA'][:2]}-{identity['CUENTA'][2:10]}-2",bold,"Bold",7.5)
    for n,(left_label,left_value,right_label,right_value) in enumerate([
        ("Saldo promedio",float(sal.mean()) if len(sal) else 0,"Saldo inicial",opening),
        ("Dias del periodo",(end-start).days+1,"+ Depositos",dep),
        ("","","- Retiros",ret),
        ("","","= saldo final",closing),
    ]):
        yy=201+n*12
        if left_label:
            _txt(page,110,yy,left_label,regular,"Regular",7.1)
            _right(page,302,yy,str(int(left_value)) if left_label.startswith("Dias") else _money(left_value,False),font,regular,"Regular",7.1)
        is_total=right_label.startswith("=")
        _txt(page,352,yy,right_label,bold if is_total else regular,"Bold" if is_total else "Regular",9 if is_total else 7.1)
        _right(page,510,yy,_money(right_value,False),font,bold if is_total else regular,"Bold" if is_total else "Regular",9 if is_total else 7.1)
    page.draw_line((32,240),(580,240),color=(.8,.8,.8),width=.5)
    _txt(page,63,266,"Detalle de movimientos cuenta de cheques.",bold,"Bold",11,RED)
    _txt(page,35,277,f"SUPERCUENTA CHEQUES-SALDO PROM {identity['CUENTA'][:2]}-{identity['CUENTA'][2:10]}-2",regular,"Regular",7.5)


def _santander_table_header(page, table_y, regular, bold):
    page.draw_rect((32,table_y,580,table_y+14),color=(.75,.75,.75),fill=(.75,.75,.75))
    _txt(page,38,table_y+10,"SALDO FINAL DEL PERIODO ANTERIOR:",bold,"Bold",8.5)
    page.draw_rect((32,table_y+14,580,table_y+30),color=(.75,.75,.75),fill=(.75,.75,.75))
    for x,label in [(44.24,"FECHA"),(86.34,"FOLIO"),(121.09,"DESCRIPCION"),(378.59,"DEPOSITO"),(451.24,"RETIRO"),(528.09,"SALDO")]:
        _txt(page,x,table_y+26,label,regular,"Regular",7.1)


def _santander(df,identity,start,end,regular,bold,font):
    doc=fitz.open(); page=None; y=0; first_page=True
    for idx,row in df.iterrows():
        lines=_santander_lines(row,idx,font); h=max(9.9,len(lines)*9.25+0.25)
        if page is None or y+h>744:
            page=doc.new_page(width=W,height=H)
            if first_page:
                _santander_first_header(page,df,identity,start,end,regular,bold,font)
                table_y=280; y=319.2
            else:
                _txt(page,32,47,"S",bold,"Bold",28,RED); _txt(page,55,47,"Santander",bold,"Bold",25,RED)
                _right(page,570,55,"ESTADO DE CUENTA",font,bold,"Bold",14,(.18,.18,.18))
                _txt(page,32,102,identity["NOMBRE"],bold,"Bold",7.5)
                table_y=118; y=147.5
            _santander_table_header(page,table_y,regular,bold)
            _watermark(page,bold); first_page=False
        if idx%2==0: page.draw_rect((32,y-10,580,y+h-10),color=(.92,.92,.92),fill=(.92,.92,.92),overlay=False)
        _txt(page,33.25,y,_date(row["FECHA"],"%d-%b-%Y"),regular,"Regular",7.1); _txt(page,81,y,_folio(4020,idx),regular,"Regular",7.1)
        for n,line in enumerate(lines): _txt(page,114,y+n*9.25,line,regular,"Regular",7.1)
        _right(page,426.94,y,_money(row["DEPOSITO"],False) if _amount(row,"DEPOSITO") else "",font,regular,"Regular",7.1)
        _right(page,494.93,y,_money(row["RETIRO"],False) if _amount(row,"RETIRO") else "",font,regular,"Regular",7.1)
        _right(page,577.72,y,"" if pd.isna(row["SALDO"]) else _money(row["SALDO"],False),font,regular,"Regular",7.1)
        y+=h
    return doc


@lru_cache(maxsize=1)
def _scotia_concept_blocks():
    """Carga los bloques reales entregados en DATOS SCOTIA.xlsx."""
    base = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(base, "datos", "DATOS SCOTIA.xlsx"),
        r"C:\Users\clopez\Music\nuevo iv\DATOS SCOTIA.xlsx",
    ]
    source = next((path for path in candidates if os.path.exists(path)), None)
    if source is None:
        return {"INGRESOS": [], "EGRESOS": []}
    result = {}
    for sheet in ("INGRESOS", "EGRESOS"):
        raw = pd.read_excel(source, sheet_name=sheet, header=None, dtype=str).fillna("")
        blocks = []
        current = None
        for _, values in raw.iterrows():
            date_value = str(values.iloc[0]).strip() if len(values) else ""
            concept_value = str(values.iloc[1]).strip() if len(values) > 1 else ""
            reference_value = str(values.iloc[2]).strip() if len(values) > 2 else ""
            if date_value:
                if current and current["lines"]:
                    blocks.append(current)
                current = {"lines": [], "reference": reference_value}
            if current is not None and concept_value:
                current["lines"].append(re.sub(r"\s+", " ", concept_value).strip())
        if current and current["lines"]:
            blocks.append(current)
        result[sheet] = blocks
    return result


def _adapt_scotia_dates(lines, movement_date):
    """Conserva el formato del catálogo, actualizando únicamente su fecha."""
    stamp = pd.to_datetime(movement_date)
    month = stamp.strftime("%b").upper()
    output = []
    for line in lines:
        line = re.sub(
            r"\b\d{1,2}\s+[A-Z]{3}\s+20\d{2}(?=\s+\d{2}:\d{2}:\d{2})",
            f"{stamp.day} {month} {stamp.year}",
            str(line).upper(),
        )
        line = re.sub(r"(?<=FECHA OPERACION: )\d{1,2}\s+[A-Z]{3}\b", f"{stamp.day:02d} {month}", line)
        output.append(line)
    return output


def _adapt_scotia_fictitious_identity(lines, row):
    """Conserva el bloque Scotia, pero aplica banco y nombre únicos del motor."""
    output = list(lines)
    _, fields = _semantic_fields(row.get("DESCRIPCION", ""))
    name = fields.get("ORDENANTE") or fields.get("BENEFICIARIO") or row.get("NOMBRE_CONTRAPARTE", "")
    bank = _transfer_bank(0, row)
    display_bank = {
        "BBVA MEXICO": "BBVA BANCOMER",
        "BANCO AZTECA": "AZTECA",
        "BANBAJIO": "BAJIO",
    }.get(bank, bank)
    bank_code = {
        "BBVA MEXICO": "012", "SANTANDER": "014", "BANORTE": "072",
        "HSBC": "021", "SCOTIABANK": "044", "BANCO AZTECA": "127",
        "BANBAJIO": "030", "INBURSA": "036", "AFIRME": "062", "BANCOPPEL": "137",
    }.get(bank, "002")
    account = _digits_only(fields.get("CUENTA"), 18)
    account = bank_code + account[3:]
    incoming = float(row.get("DEPOSITO") or 0) > 0
    if incoming and len(output) >= 8:
        output[3], output[6], output[7] = display_bank, str(name).upper(), f"/{account}"
    elif not incoming and len(output) >= 8:
        output[1] = display_bank
        output[2] = f"TRANSFERENCIA A {str(name).split()[0].upper()}"
        output[6], output[7] = str(name).upper(), f"/{account}"
    return output


def _scotia_lines(row, idx, font):
    concept,fields=_semantic_fields(row.get("DESCRIPCION",""))
    upper=concept.upper()
    if str(row.get("ORIGEN", "")).upper() == "ORIGINAL" and row.get(PHYSICAL_LINES_COLUMN) is not None:
        preserved = decode_physical_lines(row, "SCOTIABANK")
        if preserved:
            # La extraccion por coordenadas puede unir al final de la primera
            # linea la referencia que visualmente pertenece a otra columna.
            preserved[0] = re.sub(r"\s+\d{16,20}$", "", preserved[0]).strip()
            if any(key in upper for key in ("SPEI", "TRANSF", "TRASPASO")) and float(row.get("DEPOSITO") or 0) > 0 and len(preserved) < 9:
                preserved.append(str(10000 + int((_digits_only(fields.get("REF"), 5) or "40002")) % 90000))
            fitted_lines = []
            for physical_line in preserved:
                fitted_lines.extend(_wrap(physical_line, 155, font, 8.35) or [physical_line])
            return fitted_lines, None
    catalog = _scotia_concept_blocks()
    catalog_key = "INGRESOS" if float(row.get("DEPOSITO") or 0) > 0 else "EGRESOS"
    blocks = catalog.get(catalog_key, [])
    # Todos los movimientos ficticios Scotia deben usar exclusivamente el
    # catálogo entregado. También se usa para transferencias originales cuyo
    # extractor no haya conservado renglones físicos completos.
    is_fictitious = str(row.get("ORIGEN", "")).upper() == "FICTICIO"
    is_incomplete_original_transfer = any(key in upper for key in ("SPEI", "TRANSF", "TRASPASO"))
    if blocks and (is_fictitious or is_incomplete_original_transfer):
        block = blocks[int(idx) % len(blocks)]
        lines = _adapt_scotia_dates(block["lines"], row["FECHA"])
        if is_fictitious:
            lines = _adapt_scotia_fictitious_identity(lines, row)
        return lines, block["reference"]
    if not any(key in upper for key in ("SPEI","TRANSF","TRASPASO")):
        return _wrap(row.get("DESCRIPCION",""),155,font,8.35), None
    bank_codes={"BBVA MEXICO":"012","BBVA BANCOMER":"012","BANCOPPEL":"137","BANCO AZTECA":"127","AZTECA":"127","SANTANDER":"014","BANORTE":"072","BANBAJIO":"030","BAJIO":"030","INBURSA":"036","AFIRME":"062","HSBC":"021","SCOTIABANK":"044","BANAMEX":"002"}
    bank=_transfer_bank(idx, row)
    bank_code=bank_codes.get(bank,"002")
    name=fields.get("ORDENANTE") or fields.get("BENEFICIARIO") or "PERSONA FICTICIA"
    raw_account=_digits_only(fields.get("CUENTA"),18)
    account=bank_code + raw_account[3:]
    trace=_alnum_only(fields.get("RASTREO"))
    seed=int((_digits_only(fields.get("REF"),6) or "123456")[-6:])
    hh,mm,ss=seed%24,(seed//24)%60,(seed//1440)%60
    stamp=f"{pd.to_datetime(row['FECHA']):%d %b %Y} {hh:02d}:{mm:02d}:{ss:02d}".upper()
    incoming=float(row.get("DEPOSITO") or 0)>0
    if incoming:
        if "TRANSFERENCIA INT" in upper:
            variant="TRANSFERENCIA INT"
        elif re.search(r"\bTRANSFER\b",upper):
            variant="TRANSFER"
        elif "PAGO" in upper:
            variant="PAGO ENV"
        elif "COBRO" in upper:
            variant="COBRO"
        elif "DEPOSITO" in upper:
            variant="DEPOSITO"
        else:
            variant="TRANSFERENCIA"
        operation_code=str(10000 + seed % 90000)
        return ["TRANSF INTERBANCARIA SPEI","TRANSF INTERBANCARIA SPEI",variant,bank,stamp,trace,name,f"/{account}",operation_code], None
    short_name=" ".join(name.split()[:2])
    if not trace:
        trace=f"{pd.to_datetime(row['FECHA']):%Y%m%d}40044B36L{seed:013d}"
    return ["SWEB TRANSF.INTERB SPEI",bank,f"TRANSFERENCIA A {short_name}",f"/{account[-8:]} {hh:02d}:{mm:02d}:{ss:02d}",trace,f"FECHA OPERACION: {pd.to_datetime(row['FECHA']):%d %b}".upper(),name,f"/{account}"], None


def _scotia(df,identity,start,end,regular,bold,font):
    doc=fitz.open(); page=None; page_no=1; rows_page=0
    body_size=8.35; line_h=9.12; first_top=80.64; rows_per_page=73
    x_grid=[45.84,88.80,250.80,376.80,439.92,511.92,587.04]

    def new_page():
        nonlocal page, page_no, rows_page
        page=doc.new_page(width=W,height=H)
        _txt(page,53.5,26,identity["CLIENTE"],regular,"Regular",8.0)
        _txt(page,245,26,f"PAGINA {page_no}",regular,"Regular",8.0)
        _right(page,583.4,26,f"Cuenta {identity['CUENTA']}",font,regular,"Regular",8.0)
        page.draw_line((45.75,50.83),(586.88,50.83),color=(0,0,0),width=1.2)
        _txt(page,266,61,"Detalle de tus movimientos",bold,"Bold",7.8)
        page.draw_line((45.75,62),(586.88,62),color=(0,0,0),width=1.2)
        page.draw_rect((45.84,66.72,587.04,80.64),color=(0,0,0),fill=(0,0,0))
        for x,label in [(53.5,"Fecha"),(151,"Concepto"),(277,"Origen / Referencia"),(391,"Depósito"),(464,"Retiro"),(538,"Saldo")]:
            _txt(page,x,76.7,label,bold,"Bold",8.0,(1,1,1))
        for x in x_grid:
            page.draw_line((x,66.72),(x,80.64),color=(.58,.58,.58),width=.72,overlay=True)
        _watermark(page,bold)
        rows_page=0
        page_no+=1

    for movement_no,(idx,row) in enumerate(df.iterrows()):
        lines,catalog_ref=_scotia_lines(row,movement_no,font)
        wrapped_lines=[]
        for concept_line in lines:
            wrapped_lines.extend(_wrap(concept_line,155,font,body_size) or [concept_line])
        lines=wrapped_lines or [""]
        fill=(.88,.88,.88) if movement_no%2 else (1,1,1)
        _,fields=_semantic_fields(row.get("DESCRIPCION",""))
        is_opening = str(row.get("DESCRIPCION", "")).upper().strip() == "SALDO ANTERIOR"
        ref="" if is_opening else (_digits_only(catalog_ref,20) or _digits_only(fields.get("REF"),20) or _digits_only(row.get("DESCRIPCION"),20))
        line_offset=0
        while line_offset < len(lines):
            if page is None or rows_page >= rows_per_page:
                new_page()
            take=min(rows_per_page-rows_page,len(lines)-line_offset)
            segment_top=first_top+rows_page*line_h
            segment_h=take*line_h
            page.draw_rect((x_grid[0],segment_top,x_grid[-1],segment_top+segment_h),color=fill,fill=fill,overlay=False)
            # Las 73 posiciones se cuentan por renglon fisico. Si un movimiento
            # cruza de pagina, su franja conserva el color hasta que termine.
            for x in x_grid:
                page.draw_line((x,segment_top),(x,segment_top+segment_h),color=(.35,.39,.39),width=.72,overlay=True)
            y=segment_top+7.56
            if line_offset==0:
                _txt(page,53.52,y,_date(row["FECHA"],"%d %b"),regular,"Regular",body_size)
                _txt(page,252.24,y,ref,regular,"Regular",body_size)
                _right(page,436.08,y,_money(row["DEPOSITO"]) if float(row["DEPOSITO"] or 0) else "",font,regular,"Regular",body_size)
                _right(page,508.07,y,_money(row["RETIRO"]) if float(row["RETIRO"] or 0) else "",font,regular,"Regular",body_size)
                _right(page,583.37,y,"" if pd.isna(row["SALDO"]) else _money(row["SALDO"]),font,regular,"Regular",body_size)
            for local_index,line in enumerate(lines[line_offset:line_offset+take]):
                _txt(page,91.68,y+local_index*line_h,line,regular,"Regular",body_size)
            rows_page+=take
            line_offset+=take
    return doc


def _hsbc_brand(page, identity, page_no, regular, bold, font):
    page.draw_rect((0,5,W,11),color=RED,fill=RED)
    _txt(page,38,44,"HSBC",regular,"Regular",24,(.12,.12,.12))
    page.draw_rect((126,24,150,48),color=RED,fill=RED)
    page.draw_line((126,24),(150,48),color=(1,1,1),width=1.2)
    page.draw_line((150,24),(126,48),color=(1,1,1),width=1.2)
    _txt(page,258,32,"CUENTA PREMIER",bold,"Bold",9)
    _txt(page,276,47,"Estado de Cuenta",bold,"Bold",9)
    _right(page,570,35,"HSBC Premier",font,bold,"Bold",13,(.25,.25,.25))
    _txt(page,28,88,identity["NOMBRE"],bold,"Bold",8)
    _right(page,570,772,f"PAG. {page_no}",font,regular,"Regular",7)


def _hsbc_movement_header(page, y, identity, regular, bold):
    page.draw_rect((38,y,570,y+15),color=(0,0,0),fill=(0,0,0))
    _txt(page,42,y+11,f"DETALLE MOVIMIENTOS CUENTA PREMIER No. {identity['CUENTA']}",bold,"Bold",7,(1,1,1))
    page.draw_rect((38,y+15,570,y+34),color=(.76,.76,.76),fill=(.76,.76,.76))
    labels=[(42,"Dia"),(72,"Descripcion"),(284,"Referencia/\nSerial"),(355,"Retiro/Cargo"),(432,"Deposito/Abono"),(526,"Saldo")]
    for x,label in labels:
        for n,line in enumerate(label.split("\n")):
            _txt(page,x,y+26+n*7,line,bold,"Bold",6.5)
    return y+44


def _hsbc_first_page(page, df, identity, start, end, regular, bold, font):
    _hsbc_brand(page,identity,1,regular,bold,font)
    dep=float(pd.to_numeric(df["DEPOSITO"],errors="coerce").fillna(0).sum())
    ret=float(pd.to_numeric(df["RETIRO"],errors="coerce").fillna(0).sum())
    sal=pd.to_numeric(df["SALDO"],errors="coerce").dropna()
    opening=float(sal.iloc[0]) if len(sal) else 0.0; closing=float(sal.iloc[-1]) if len(sal) else 0.0
    _txt(page,42,115,identity["DOMICILIO"],regular,"Regular",7)
    _txt(page,42,205,"NUMERO DE CUENTA",bold,"Bold",7); _txt(page,42,218,identity["CUENTA"],regular,"Regular",7)
    _txt(page,180,205,"CLABE INTERBANCARIA",bold,"Bold",7); _txt(page,180,218,identity["CLABE"],regular,"Regular",7)
    _txt(page,360,118,"RESUMEN DE CUENTAS",bold,"Bold",7)
    page.draw_line((355,111),(560,111),color=(0,0,0),width=2)
    for n,(label,value) in enumerate([("Saldo Inicial del Periodo",opening),("Depositos/Abonos",dep),("Retiros/Cargos",ret),("Saldo Final del Periodo",closing)]):
        yy=135+n*24; total=n==3
        _txt(page,365,yy,label,bold if total else regular,"Bold" if total else "Regular",7)
        _right(page,558,yy,_money(value),font,bold if total else regular,"Bold" if total else "Regular",7)
    _txt(page,365,245,f"Periodo del {start:%d/%m/%Y} al {end:%d/%m/%Y}",regular,"Regular",7)
    page.draw_rect((38,315,570,330),color=(0,0,0),fill=(0,0,0))
    _txt(page,42,326,"RESUMEN DE TU CUENTA PREMIER",bold,"Bold",7,(1,1,1))
    for n,(label,value) in enumerate([("Comisiones Cobradas en el Mes",0),("Saldo Promedio en el Mes",float(sal.mean()) if len(sal) else 0),("Tasa Promedio Nominal",0),("Intereses Netos",0)]):
        yy=342+n*18
        if n%2: page.draw_rect((38,yy-10,570,yy+4),color=(.82,.82,.82),fill=(.82,.82,.82))
        _txt(page,42,yy,label,regular,"Regular",7); _right(page,566,yy,_money(value),font,regular,"Regular",7)
    return _hsbc_movement_header(page,420,identity,regular,bold)


def _hsbc_chart(page, y, df, regular, bold, font):
    ret_series=pd.to_numeric(df["RETIRO"],errors="coerce").fillna(0)
    ret=float(ret_series.sum())
    cash=float(ret_series[df["DESCRIPCION"].astype(str).str.upper().str.contains("EFECTIVO",na=False)].sum())
    other=max(0.0,ret-cash)
    sal=pd.to_numeric(df["SALDO"],errors="coerce").dropna()
    opening=float(sal.iloc[0]) if len(sal) else 0.0; closing=float(sal.iloc[-1]) if len(sal) else 0.0
    _txt(page,40,y,"ColDi: Operacion procesada por CoDi",bold,"Bold",7)
    page.draw_oval((105,y+36,205,y+82),color=(.25,.25,.25),fill=(.5,.5,.5),width=.6)
    page.draw_oval((105,y+28,205,y+72),color=(.25,.25,.25),fill=(.65,.65,.65),width=.6)
    _txt(page,92,y+27,f"Saldo final {_money(closing)}",regular,"Regular",6.5)
    _txt(page,152,y+94,f"Otros cargos {_money(other)}",regular,"Regular",6.5)
    page.draw_rect((365,y+30,545,y+100),color=(.2,.2,.2),width=.7)
    rows=[("Saldo Inicial",opening),("Saldo Final",closing),("Comisiones",0),("Retiro en Efectivo",cash),("Otros Cargos",other)]
    for n,(label,value) in enumerate(rows):
        yy=y+43+n*13
        if n: page.draw_line((365,yy-9),(545,yy-9),color=(.65,.65,.65),width=.35)
        _txt(page,395,yy,label,bold if n<2 else regular,"Bold" if n<2 else "Regular",6.5)
        _right(page,540,yy,_money(value),font,regular,"Regular",6.5)


def _hsbc_transfer_record(row, idx):
    concept,fields=_semantic_fields(row.get("DESCRIPCION",""))
    direction=_movement_direction(row)
    name=fields.get("ORDENANTE") if direction=="RECIBIDO" else fields.get("BENEFICIARIO")
    name=name or "PERSONA FICTICIA"
    reference=_digits_only(fields.get("REF"),12)
    seed=int((reference or "123456")[-6:])
    return {
        "date":pd.to_datetime(row["FECHA"]).strftime("%d/%m/%Y"),
        "time":f"{seed%24:02d}:{(seed//24)%60:02d}:{(seed//1440)%60:02d}",
        "bank":_transfer_bank(idx, row), "name":name,
        "account":_digits_only(fields.get("CUENTA"),18),
        "concept":fields.get("CONCEPTO") or concept,
        "amount":_amount(row,"DEPOSITO") or _amount(row,"RETIRO"),
        "trace":_digits_only(fields.get("RASTREO"),20),
        "reference":reference, "direction":direction,
    }


def _hsbc_transfer_section(doc, records, title, identity, regular, bold, font, page_no):
    page=None; y=0
    x=[38,82,125,181,251,343,410,469,521,570]
    headers=["Fecha de\nOperacion","Hora de\nOperacion","Participante\nBancario","Nombre del\nCliente","Cuenta","Concepto del\npago","Monto del pago","Clave de\nRastreo","Numero de\nReferencia"]

    def new_page(number):
        p=doc.new_page(width=W,height=H)
        _hsbc_brand(p,identity,number,regular,bold,font)
        p.draw_rect((38,104,570,118),color=(0,0,0),fill=(0,0,0)); _txt(p,120,114,title,bold,"Bold",6.4,(1,1,1))
        p.draw_rect((38,118,570,146),color=(.76,.76,.76),fill=(.76,.76,.76))
        for i,label in enumerate(headers):
            for n,line in enumerate(label.split("\n")): _txt(p,x[i]+2,128+n*7,line,bold,"Bold",5.6)
        _watermark(p,bold)
        return p

    for record in records:
        if page is None:
            page_no+=1; page=new_page(page_no); y=157
        values=[record["date"],record["time"],record["bank"],record["name"],record["account"],record["concept"],_money(record["amount"]),record["trace"],record["reference"]]
        wrapped=[_wrap(value,x[i+1]-x[i]-4,font,5.7)[:4] for i,value in enumerate(values)]
        row_h=max(24,max(len(lines) for lines in wrapped)*7+4)
        if y+row_h>724:
            page_no+=1; page=new_page(page_no); y=157
        for i,lines in enumerate(wrapped):
            for n,line in enumerate(lines): _txt(page,x[i]+2,y+n*7,line,regular,"Regular",5.7)
        page.draw_line((38,y+row_h-4),(570,y+row_h-4),color=(.65,.65,.65),width=.35)
        for xx in x: page.draw_line((xx,y-7),(xx,y+row_h-4),color=(.72,.72,.72),width=.25)
        y+=row_h
    return page_no


def _hsbc(df, identity, start, end, regular, bold, font):
    """HSBC calibrado con los tres estados de cuenta originales suministrados."""
    doc=fitz.open(); page=doc.new_page(width=W,height=H); page_no=1
    y=_hsbc_first_page(page,df,identity,start,end,regular,bold,font); _watermark(page,bold)
    for idx,row in df.iterrows():
        if str(row.get("DESCRIPCION","")).upper().strip()=="SALDO ANTERIOR": continue
        h=19.5
        if y+h>650:
            page=doc.new_page(width=W,height=H); page_no+=1
            _hsbc_brand(page,identity,page_no,regular,bold,font)
            y=_hsbc_movement_header(page,104,identity,regular,bold); _watermark(page,bold)
        concept,_=_semantic_fields(row.get("DESCRIPCION",""))
        digits=re.sub(r"\D","",str(row.get("DESCRIPCION","")))
        serial=(digits[-16:] if digits else ("08045219"+_folio(81,idx)))
        _txt(page,43,y,_date(row["FECHA"],"%d"),regular,"Regular",7)
        for n,line in enumerate(_wrap(concept,202,font,7)[:2]): _txt(page,72,y+n*8,line,regular,"Regular",7)
        _txt(page,300,y,serial[:8],regular,"Regular",6.5); _txt(page,300,y+8,serial[8:],regular,"Regular",6.5)
        _right(page,410,y,_money(row["RETIRO"]) if _amount(row,"RETIRO") else "",font,regular,"Regular",7)
        _right(page,485,y,_money(row["DEPOSITO"]) if _amount(row,"DEPOSITO") else "",font,regular,"Regular",7)
        _right(page,566,y,"" if pd.isna(row["SALDO"]) else _money(row["SALDO"]),font,regular,"Regular",7)
        page.draw_line((38,y+h-7),(570,y+h-7),color=(.7,.7,.7),width=.35); y+=h
    if y+130>720:
        page=doc.new_page(width=W,height=H); page_no+=1; _hsbc_brand(page,identity,page_no,regular,bold,font); y=120; _watermark(page,bold)
    _hsbc_chart(page,y,df,regular,bold,font)
    transfers=[_hsbc_transfer_record(row,idx) for idx,row in df.iterrows() if _is_transfer(row)]
    sent=[r for r in transfers if r["direction"]=="ENVIADO"]
    received=[r for r in transfers if r["direction"]=="RECIBIDO"]
    period=f"durante el periodo del {start:%d/%m/%Y} al {end:%d/%m/%Y}"
    if sent: page_no=_hsbc_transfer_section(doc,sent,"Informacion SPEI's Enviados "+period,identity,regular,bold,font,page_no)
    if received: page_no=_hsbc_transfer_section(doc,received,"Informacion SPEI's Recibidos "+period,identity,regular,bold,font,page_no)
    return doc


def generate_movement_pdf_bytes(df, identity, start, end, style):
    regular,bold=_fonts_for_style(style)
    base=os.path.dirname(os.path.abspath(__file__)); free=os.path.join(base,"fuentes","libres"); windows=os.path.join(os.environ.get("WINDIR",r"C:\Windows"),"Fonts")
    if style in {"BBVA 1","BBVA LARGO","SANTANDER"}:
        italic=_first_existing(os.path.join(windows,"ARIALNI.TTF"),os.path.join(free,"RobotoCondensed-Italic-Variable.ttf"),regular)
    else:
        italic=_first_existing(os.path.join(windows,"ariali.ttf"),os.path.join(free,"Arimo-Italic-Variable.ttf"),regular)
    font=fitz.Font(fontfile=regular)
    if style=="AZTECA": doc=_azteca(df,identity,start,end,regular,bold,font)
    elif style=="BANAMEX": doc=_banamex(df,identity,start,end,regular,bold,font)
    elif style=="BANAMEX (JUN-JUL)": doc=_banamex(df,identity,start,end,regular,bold,font,True)
    elif style in {"BBVA 1", "BBVA LARGO"}: doc=_bbva(df,identity,start,end,regular,bold,italic,font,1)
    elif style=="BBVA 2": doc=_bbva(df,identity,start,end,regular,bold,italic,font,2)
    elif style=="SANTANDER": doc=_santander(df,identity,start,end,regular,bold,font)
    elif style=="SCOTIABANK": doc=_scotia(df,identity,start,end,regular,bold,font)
    elif style=="HSBC": doc=_hsbc(df,identity,start,end,regular,bold,font)
    else: raise ValueError(f"Estilo de movimientos no soportado: {style}")
    if style in {"BANAMEX", "BANAMEX (JUN-JUL)"}:
        _assert_banamex_render_integrity(doc, df)
    # Incrusta solamente los glifos utilizados. Mantiene texto vectorial y
    # editable, pero evita repetir archivos completos de tipografia.
    # BanamexDisplay ya proviene como subconjunto del PDF de referencia;
    # volver a subconjuntarlo puede dañar sus indices CFF.
    if style not in {"BANAMEX", "BANAMEX (JUN-JUL)"}:
        doc.subset_fonts()
    return doc.tobytes(deflate=True,garbage=4,clean=True)
