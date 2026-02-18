import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import re
import json
import io
import os
import requests
import zipfile
import urllib3
from datetime import datetime
import xlsxwriter
import time
import urllib.parse

# --- 1. CONFIGURACIÓN Y SEGURIDAD ---
st.set_page_config(page_title="RAPIDITO AI - Portal Contable", layout="wide", page_icon="📊")
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- ENLACES DE CONEXIÓN ---
URL_WS = "https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline?wsdl"
HEADERS_WS = {"Content-Type": "text/xml;charset=UTF-8","User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36"}

# El CSV de lectura (Respaldo)
URL_SHEET = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSkIGy-ovamvkCQjnjuT9kV7BndRqOeZlrEUEy9BZUH-oGISXG2a_own9BMbzTV21giZXgBqGxlTjkp/pub?output=csv"

# ⚠️ PEGA AQUÍ LA URL DE TU GOOGLE APPS SCRIPT (Termina en /exec)
URL_API_VIRAL = "https://script.google.com/macros/s/AKfycbz3vRq203m7vcdor30hJiXuAGNGr8n_kM2dCpf63LW4KhaeY9wqAijBC473AwywYes/exec" 

# --- CONEXIÓN AL CEREBRO VIRAL (Google Apps Script) ---
def conectar_api(payload):
    """Función para hablar con Google Sheets (Login e Invitaciones)"""
    try:
        # Si la URL no está configurada, retornamos error simulado
        if "TU_URL" in URL_API_VIRAL: 
            return {"exito": False, "mensaje": "Falta configurar URL_API_VIRAL en el código."}
        
        r = requests.post(URL_API_VIRAL, json=payload, timeout=10)
        return r.json()
    except Exception as e:
        return {"exito": False, "mensaje": f"Error de conexión: {str(e)}"}

# --- LOGGING Y SUGERENCIAS (Legacy) ---
def registrar_actividad(usuario, accion, cantidad=None, sugerencia=None):
    URL_LOGGING = "https://script.google.com/macros/s/AKfycbyk0CWehcUec47HTGMjqsCs0sTKa_9J3ZU_Su7aRxfwmNa76-dremthTuTPf-FswZY/exec"
    detalle_accion = f"{accion} ({cantidad} XMLs)" if cantidad is not None else accion
    payload = {"usuario": str(usuario), "accion": str(detalle_accion)}
    if sugerencia: payload["sugerencia"] = str(sugerencia)
    try: 
        requests.post(URL_LOGGING, json=payload, timeout=5)
        return True
    except: return False

# --- 2. SISTEMA DE LOGIN VIRAL Y ESTADO ---
if "autenticado" not in st.session_state: st.session_state.autenticado = False
if "id_proceso" not in st.session_state: st.session_state.id_proceso = 0
if "data_compras_cache" not in st.session_state: st.session_state.data_compras_cache = []
if "data_ventas_cache" not in st.session_state: st.session_state.data_ventas_cache = []
if "invitaciones_disponibles" not in st.session_state: st.session_state.invitaciones_disponibles = 0

if not st.session_state.autenticado:
    st.sidebar.title("🔐 Acceso RAPIDITO")
    st.sidebar.markdown("---")
    user = st.sidebar.text_input("Usuario (Email)")
    password = st.sidebar.text_input("Contraseña", type="password")
    
    if st.sidebar.button("Iniciar Sesión"):
        with st.spinner("Verificando credenciales..."):
            resp = conectar_api({"accion": "LOGIN", "usuario": user.strip(), "clave": password.strip()})
            
            if resp.get("exito"):
                st.session_state.autenticado = True
                st.session_state.usuario_actual = user.strip()
                st.session_state.invitaciones_disponibles = resp.get("invitaciones", 0)
                registrar_actividad(user, "ENTRÓ AL PORTAL")
                st.rerun()
            else:
                st.sidebar.error(f"Error: {resp.get('mensaje', 'Credenciales incorrectas')}")
    
    st.sidebar.markdown("---")
    st.sidebar.info("ℹ️ Este sistema es exclusivo por invitación. Pídele acceso a un colega contador.")
    st.stop()

# --- 3. MEMORIA DE APRENDIZAJE ---
if 'memoria' not in st.session_state:
    archivo_memoria = "conocimiento_contable.json"
    if os.path.exists(archivo_memoria):
        with open(archivo_memoria, "r", encoding="utf-8") as f: st.session_state.memoria = json.load(f)
    else: st.session_state.memoria = {"empresas": {}}

def guardar_memoria():
    with open("conocimiento_contable.json", "w", encoding="utf-8") as f: json.dump(st.session_state.memoria, f, indent=4, ensure_ascii=False)

# --- HELPER: DESCOMPRIMIR ZIP Y XMLs ---
def procesar_archivos_entrada(lista_archivos):
    xmls_procesables = []
    for file in lista_archivos:
        if file.name.lower().endswith('.xml'):
            xmls_procesables.append(io.BytesIO(file.getvalue()))
        elif file.name.lower().endswith('.zip'):
            try:
                with zipfile.ZipFile(file) as z:
                    for filename in z.namelist():
                        if filename.lower().endswith('.xml') and not filename.startswith('__MACOSX'):
                            xmls_procesables.append(io.BytesIO(z.read(filename)))
            except: pass
    return xmls_procesables

# --- 4. MOTOR DE EXTRACCIÓN XML (VERSION BLINDADA V2) ---
def extraer_datos_robusto(xml_file):
    try:
        if isinstance(xml_file, (io.BytesIO, io.StringIO)): xml_file.seek(0)
        tree = ET.parse(xml_file)
        root = tree.getroot()
        xml_data = None
        
        for elem in root.iter():
            if 'comprobante' in elem.tag.lower() and elem.text and ("<" in elem.text or "&lt;" in elem.text):
                try:
                    clean_text = re.sub(r'<\?xml.*?\?>', '', elem.text).strip()
                    xml_data = ET.fromstring(clean_text)
                    break
                except: continue
        
        if xml_data is None: xml_data = root

        root_tag = xml_data.tag.lower()
        if 'notacredito' in root_tag: tipo_doc = "NC"
        elif 'comprobanteretencion' in root_tag: tipo_doc = "RET"
        elif 'liquidacioncompra' in root_tag: tipo_doc = "LC"
        else: tipo_doc = "FC" 

        def buscar(tags):
            for t in tags:
                f = xml_data.find(f".//{t}")
                if f is not None and f.text: return f.text.strip()
            return ""
            
        def buscar_float(tags):
            val_str = buscar(tags)
            try: return float(val_str) if val_str else 0.0
            except: return 0.0

        razon_social = buscar(["razonSocial"]).upper()
        ruc_emisor = buscar(["ruc"])
        
        estab = buscar(["estab"]) or "000"
        pto = buscar(["ptoEmi"]) or "000"
        sec = buscar(["secuencial"]) or "000000000"
        num_fact_completo = f"{estab}-{pto}-{sec}"
        
        fecha_emision = buscar(["fechaEmision"])
        num_autori = buscar(["numeroAutorizacion"]) or buscar(["claveAcceso"])
        
        mes_nombre = "DESCONOCIDO"
        if "/" in fecha_emision:
            try:
                meses_dict = {"01":"ENERO","02":"FEBRERO","03":"MARZO","04":"ABRIL","05":"MAYO","06":"JUNIO","07":"JULIO","08":"AGOSTO","09":"SEPTIEMBRE","10":"OCTUBRE","11":"NOVIEMBRE","12":"DICIEMBRE"}
                mes_nombre = meses_dict.get(fecha_emision.split('/')[1], "DESCONOCIDO")
            except: pass

        ruc_cliente = buscar(["identificacionComprador", "identificacionSujetoRetenido"])
        nombre_cliente = buscar(["razonSocialComprador", "razonSocialSujetoRetenido"]).upper()

        # --- LÓGICA DE CLASIFICACIÓN 10 VS 13 DÍGITOS ---
        len_id = len(ruc_cliente)
        info_json = st.session_state.memoria["empresas"].get(razon_social)

        if tipo_doc == "NC":
            detalle_final, memo_final = "", ""
        elif len_id == 10:
            memo_final = "PERSONAL"
            # Si está en el JSON, usa la categoría guardada; si no, No Deducible.
            detalle_final = info_json.get("DETALLE", "No Deducible") if info_json else "No Deducible"
        else:
            # Lógica estándar para 13 dígitos o más
            if info_json:
                detalle_final = info_json.get("DETALLE", "OTROS")
                memo_final = info_json.get("MEMO", "PROFESIONAL")
            else:
                detalle_final, memo_final = "OTROS", "PROFESIONAL"

        base_data = {
            "TIPO": tipo_doc, "TIPO DE DOCUMENTO": tipo_doc,
            "MES": mes_nombre, "FECHA": fecha_emision, "N. FACTURA": num_fact_completo, 
            "RUC": ruc_emisor, "NOMBRE": razon_social, "N AUTORIZACION": num_autori,
            "CONTRIBUYENTE": ruc_cliente, "RUC CLIENTE": ruc_cliente, "CLIENTE": nombre_cliente,
            "DETALLE": detalle_final, "MEMO": memo_final
        }

        if tipo_doc == "RET":
            rt_renta, rt_iva, base_renta, base_iva = 0.0, 0.0, 0.0, 0.0
            sustento_formateado = ""
            
            doc_sus_node = xml_data.find(".//numDocSustento")
            doc_sus_raw = doc_sus_node.text.strip() if (doc_sus_node is not None and doc_sus_node.text) else ""
            
            if doc_sus_raw:
                parts = doc_sus_raw.replace('-','').strip()
                if len(parts) >= 15: 
                    sustento_formateado = f"{parts[0:3]}-{parts[3:6]}-{parts[6:]}"
                elif len(doc_sus_raw.split('-')) == 3:
                    sustento_formateado = doc_sus_raw

            lista_retenciones = xml_data.findall(".//impuesto") + xml_data.findall(".//retencion")

            for item in lista_retenciones:
                cod_node = item.find("codigo")
                cod = cod_node.text.strip() if (cod_node is not None and cod_node.text) else ""
                try:
                    val_node = item.find("valorRetenido")
                    val = float(val_node.text.strip() if (val_node is not None and val_node.text) else "0")
                except: val = 0.0
                try:
                    base_node = item.find("baseImponible")
                    base = float(base_node.text.strip() if (base_node is not None and base_node.text) else "0")
                except: base = 0.0
                
                if cod == "1": rt_renta += val; base_renta += base
                elif cod == "2": rt_iva += val; base_iva += base

            base_data.update({
                "ruc_recep": ruc_cliente, "nomrecep": nombre_cliente, "fechaemi": fecha_emision,
                "razonsocial": razon_social, "ruc_emisor": ruc_emisor, "numfact": sustento_formateado, 
                "numreten": num_fact_completo, "baserenta": base_renta, "rt_renta": rt_renta,
                "baseiva": base_iva, "rt_iva": rt_iva, "numautori": num_autori,
                "fecautori": buscar(["fechaAutorizacion"]) or fecha_emision,
                "SUSTENTO": sustento_formateado, "TOTAL RET": rt_renta + rt_iva
            })
            return base_data

        else: 
            m = -1 if tipo_doc == "NC" else 1
            total = buscar_float(["importeTotal", "total", "valorModificado"]) * m
            propina = buscar_float(["propina"]) * m
            
            base_0, base_12_15, iva_12_15 = 0.0, 0.0, 0.0
            no_obj_iva, exento_iva = 0.0, 0.0
            otra_base, otro_monto_iva, ice_val = 0.0, 0.0, 0.0
            
            for imp in xml_data.findall(".//totalImpuesto"):
                try:
                    cod = imp.find("codigo").text
                    cod_por = imp.find("codigoPorcentaje").text
                    base = float(imp.find("baseImponible").text or 0) * m
                    valor = float(imp.find("valor").text or 0) * m
                    
                    if cod == "2":
                        if cod_por == "0": base_0 += base
                        elif cod_por in ["2", "3", "4", "8", "10"]:
                            base_12_15 += base; iva_12_15 += valor
                        elif cod_por == "6": no_obj_iva += base
                        elif cod_por == "7": exento_iva += base
                        else: otra_base += base; otro_monto_iva += valor
                    elif cod == "3": ice_val += valor
                    else: otra_base += base; otro_monto_iva += valor
                except: continue 

            items = [d.find("descripcion").text for d in xml_data.findall(".//detalle") if d.find("descripcion") is not None]
            subdetalle = " | ".join(items[:5]) if items else ""

            base_data.update({
                "SUBDETALLE": subdetalle,
                "OTRA BASE IVA": otra_base, "OTRO IVA": otro_monto_iva, 
                "MONTO ICE": ice_val, "PROPINAS": propina,
                "EXENTO DE IVA": exento_iva, "NO OBJ IVA": no_obj_iva, 
                "BASE. 0": base_0, "BASE. 12 / 15": base_12_15,
                "IVA.": iva_12_15, "TOTAL": total
            })
            return base_data
    except Exception as e:
        print(f"Error procesando XML: {e}")
        return None

# --- 5. LÓGICA DE INTEGRACIÓN (CRUCE VENTAS) ---
def procesar_ventas_con_retenciones(lista_datos_crudos):
    ventas, retenciones_map = [], {}
    for dato in lista_datos_crudos:
        if dato["TIPO"] == "FC": ventas.append(dato)
        elif dato["TIPO"] == "RET" and dato.get("SUSTENTO"): retenciones_map[dato["SUSTENTO"]] = dato

    ventas_integradas = []
    for venta in ventas:
        num_fact = venta["N. FACTURA"] 
        ret_asociada = retenciones_map.get(num_fact, {}) 
        fila = {
            "MES": venta.get("MES"), "FECHA": venta.get("FECHA"), "N. FACTURA": num_fact,
            "RUC": venta.get("RUC CLIENTE"), "CLIENTE": venta.get("CLIENTE"),
            "DETALLE": "SERVICIOS", "MEMO": "PROFESIONAL", "MONTO REEMBOLS": 0.0,
            "BASE. 0": venta.get("BASE. 0", 0), "BASE. 12 / 15": venta.get("BASE. 12 / 15", 0),
            "IVA": venta.get("IVA.", 0), "TOTAL": venta.get("TOTAL", 0),
            "FECHA RET": ret_asociada.get("fechaemi", ""),
            "N° RET": ret_asociada.get("numreten", ""),
            "N° AUTORIZACIÓN": ret_asociada.get("numautori", ""),
            "RET RENTA": ret_asociada.get("rt_renta", 0), 
            "RET IVA": ret_asociada.get("rt_iva", 0),
            "ISD": 0.0, "TOTAL RET": ret_asociada.get("TOTAL RET", 0)
        }
        ventas_integradas.append(fila)
    return ventas_integradas

# --- 6. GENERADOR MULTI-EXCEL MAESTRO (CON MARCA VIRAL) ---
def generar_excel_multiexcel(data_compras=None, data_ventas_ret=None, data_sri_lista=None, sri_mode=None):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book
        f_azul = wb.add_format({'bold':True,'align':'center','border':1,'bg_color':'#002060','font_color':'white'})
        f_amar = wb.add_format({'bold':True,'align':'center','border':1,'bg_color':'#FFD966'})
        f_verd = wb.add_format({'bold':True,'align':'center','border':1,'bg_color':'#92D050'})
        f_gris = wb.add_format({'bold':True,'align':'center','border':1,'bg_color':'#F2F2F2'})
        f_num = wb.add_format({'num_format':'_-$ * #,##0.00_-','border':1})
        f_tot = wb.add_format({'bold':True,'num_format':'_-$ * #,##0.00_-','border':1,'bg_color':'#EFEFEF'})
        
        texto_pie = "&LGenerado por RAPIDITO AI&RConsigue tu cuenta gratis en: rapidito.ec"

        if sri_mode:
            df = pd.DataFrame(data_sri_lista)
            if sri_mode == "NC":
                cols = ["NOMBRE","RUC","N AUTORIZACION","FECHA","TIPO DE DOCUMENTO","N. FACTURA","MES","RUC CLIENTE","CLIENTE","PROPINAS","BASE. 0","NO OBJ IVA","BASE. 12 / 15","IVA.","TOTAL"]
                header_fmt = f_amar; sheet_name = "NOTAS DE CREDITO"
            elif sri_mode == "RET":
                cols = ["ruc_recep", "nomrecep", "fechaemi", "razonsocial", "ruc_emisor", "numfact", "numreten", "baserenta", "rt_renta", "baseiva", "rt_iva", "numautori", "fecautori"]
                header_fmt = f_verd; sheet_name = "RETENCIONES"
            else: 
                cols = ["MES","FECHA","N. FACTURA","TIPO DE DOCUMENTO","RUC","CONTRIBUYENTE","NOMBRE","DETALLE","MEMO","OTRA BASE IVA","OTRO IVA","MONTO ICE","PROPINAS","EXENTO DE IVA","NO OBJ IVA","BASE. 0","BASE. 12 / 15","IVA.","TOTAL","SUBDETALLE"]
                header_fmt = f_azul; sheet_name = "FACTURAS"

            for c in cols: 
                if c not in df.columns: df[c] = ""
            df = df[cols]
            ws = wb.add_worksheet(sheet_name)
            ws.set_footer(texto_pie) 
            for i, c in enumerate(cols): ws.write(0, i, c, header_fmt)
            for r, row in enumerate(df.values, 1):
                for c, val in enumerate(row): ws.write(r, c, val, f_num if isinstance(val, (int,float)) else wb.add_format({'border':1}))
            ws.set_column(0, len(cols)-1, 15)

        else:
            meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
            if data_compras:
                df_c = pd.DataFrame(data_compras)
                orden_c = ["MES","FECHA","N. FACTURA","TIPO DE DOCUMENTO","RUC","CONTRIBUYENTE","NOMBRE","DETALLE","MEMO","OTRA BASE IVA","OTRO IVA","MONTO ICE","PROPINAS","EXENTO DE IVA","NO OBJ IVA","BASE. 0","BASE. 12 / 15","IVA.","TOTAL","SUBDETALLE"]
                for c in orden_c: 
                    if c not in df_c.columns: df_c[c] = ""
                df_c = df_c[orden_c]
                ws_c = wb.add_worksheet('COMPRAS')
                ws_c.set_footer(texto_pie) 
                for i, c in enumerate(orden_c):
                    fmt = f_amar if i in range(9, 15) else f_azul
                    ws_c.write(0, i, c, fmt)
                for r, row in enumerate(df_c.values, 1):
                    for c, val in enumerate(row): ws_c.write(r, c, val, f_num if isinstance(val, (int,float)) else wb.add_format({'border':1}))
                ft = len(df_c) + 1; ws_c.write(ft, 0, "TOTAL", f_tot)
                for cidx in range(9, 19): 
                    l = xlsxwriter.utility.xl_col_to_name(cidx); ws_c.write_formula(ft, cidx, f"=SUM({l}2:{l}{ft})", f_tot)

                ws_ra = wb.add_worksheet('REPORTE ANUAL')
                ws_ra.set_footer(texto_pie) 
                ws_ra.set_column('A:K', 14); ws_ra.merge_range('B1:B2', "Negocios y\nServicios", f_azul)
                cats=["VIVIENDA","SALUD","EDUCACION","ALIMENTACION","VESTIMENTA","TURISMO","NO DEDUCIBLE","SERVICIOS BASICOS"]
                icos=["🏠","❤️","🎓","🛒","🧢","✈️","🚫","💡"]
                for i,(ct,ic) in enumerate(zip(cats,icos)): ws_ra.write(0,i+2,ic,f_azul); ws_ra.write(1,i+2,ct.title(),f_azul)
                ws_ra.merge_range('K1:K2',"Total Mes",f_azul); ws_ra.write('B3',"PROFESIONALES",f_gris); ws_ra.merge_range('C3:J3',"GASTOS PERSONALES",f_gris)
                
                cols_gasto = ["P","Q","O","N","J","M"] 
                for r, mes in enumerate(meses):
                    fila = r+4; ws_ra.write(r+3,0,mes.title(),f_num)
                    f_pr = "+".join([f"SUMIFS('COMPRAS'!${l}:${l},'COMPRAS'!$A:$A,\"{mes}\",'COMPRAS'!$I:$I,\"PROFESIONAL\")" for l in ["P","Q","O","N","J","M"]])
                    ws_ra.write_formula(r+3,1,"="+f_pr,f_num)
                    for cidx, cat in enumerate(cats):
                        f_pe = "+".join([f"SUMIFS('COMPRAS'!${l}:${l},'COMPRAS'!$A:$A,\"{mes}\",'COMPRAS'!$H:$H,\"{cat}\")" for l in cols_gasto])
                        ws_ra.write_formula(r+3,cidx+2,"="+f_pe,f_num)
                    ws_ra.write_formula(r+3,10,f"=SUM(B{fila}:J{fila})",f_num)
                ws_ra.write(15,0,"TOTAL",f_tot)
                for c in range(1,11): l=xlsxwriter.utility.xl_col_to_name(c); ws_ra.write_formula(15,c,f"=SUM({l}4:{l}15)",f_tot)

            if data_ventas_ret:
                df_v = pd.DataFrame(data_ventas_ret)
                orden_v = ["MES","FECHA","N. FACTURA","RUC","CLIENTE","DETALLE","MEMO","MONTO REEMBOLS","BASE. 0","BASE. 12 / 15","IVA","TOTAL","FECHA RET","N° RET","N° AUTORIZACIÓN","RET RENTA","RET IVA","ISD","TOTAL RET"]
                for c in orden_v: 
                    if c not in df_v.columns: df_v[c] = ""
                df_v = df_v[orden_v]
                ws_v = wb.add_worksheet('VENTAS')
                ws_v.set_footer(texto_pie) 
                for i, c in enumerate(orden_v): ws_v.write(0, i, c, f_verd if i >= 12 else f_azul)
                for r, row in enumerate(df_v.values, 1):
                    for c, val in enumerate(row): ws_v.write(r, c, val, f_num if isinstance(val, (int,float)) else wb.add_format({'border':1}))
                ft_v = len(df_v) + 1; ws_v.write(ft_v, 0, "TOTAL", f_tot)
                for cidx in range(7, 19): l = xlsxwriter.utility.xl_col_to_name(cidx); ws_v.write_formula(ft_v, cidx, f"=SUM({l}2:{l}{ft_v})", f_tot)

                ws_p = wb.add_worksheet('PROYECCION')
                ws_p.set_footer(texto_pie) 
                ws_p.set_column('A:A', 12); ws_p.set_column('B:M', 15)
                ws_p.merge_range('A1:D1', f"PERIODO: {datetime.now().year}", f_azul)
                for i, h in enumerate(["VENTAS", "COMPRAS", "TOTAL"]): ws_p.write(i+2, 0, h, f_azul)
                for c, mes in enumerate(meses):
                    col = c + 1; l = xlsxwriter.utility.xl_col_to_name(col)
                    ws_p.write(1, col, mes, f_azul)
                    ws_p.write_formula(2, col, f"=SUMIFS(VENTAS!$I:$I,VENTAS!$A:$A,\"{mes}\") + SUMIFS(VENTAS!$J:$J,VENTAS!$A:$A,\"{mes}\")", f_num)
                    if data_compras: ws_p.write_formula(3, col, 
                            f"=SUMIFS('COMPRAS'!$P:$P,'COMPRAS'!$A:$A,{l}$2,'COMPRAS'!$I:$I,\"PROFESIONAL\") + "
                            f"SUMIFS('COMPRAS'!$Q:$Q,'COMPRAS'!$A:$A,{l}$2,'COMPRAS'!$I:$I,\"PROFESIONAL\") + "
                            f"SUMIFS('COMPRAS'!$O:$O,'COMPRAS'!$A:$A,{l}$2,'COMPRAS'!$I:$I,\"PROFESIONAL\") + "
                            f"SUMIFS('COMPRAS'!$N:$N,'COMPRAS'!$A:$A,{l}$2,'COMPRAS'!$I:$I,\"PROFESIONAL\") + "
                            f"SUMIFS('COMPRAS'!$J:$J,'COMPRAS'!$A:$A,{l}$2,'COMPRAS'!$I:$I,\"PROFESIONAL\")", 
                            f_num)
                    else: ws_p.write(3, col, 0, f_num)
                    ws_p.write_formula(4, col, f"={l}3-{l}4", f_tot)
                lt = xlsxwriter.utility.xl_col_to_name(len(meses)+1)
                ws_p.write(1, len(meses)+1, "TOTAL", f_azul)
                for r in range(2,5): ws_p.write_formula(r, len(meses)+1, f"=SUM(B{r+1}:{xlsxwriter.utility.xl_col_to_name(len(meses))}{r+1})", f_tot)

    return output.getvalue()

# --- 7. INTERFAZ ---
st.title(f"🚀 RAPIDITO AI - Portal Contable")

with st.sidebar:
    st.header("⚙️ Panel de Control")
    
    # 1. NUEVO INFORME
    if st.button("🧹 NUEVO INFORME", type="primary", use_container_width=True):
        st.session_state.id_proceso += 1
        st.session_state.data_compras_cache = []
        st.session_state.data_ventas_cache = []
        st.rerun()
    
    st.markdown("---")

    # 2. CONFIGURACIÓN MAESTRO
    if st.session_state.usuario_actual == "GABRIEL":
        st.subheader("🔑 Master Config")
        up_xls = st.file_uploader("Actualizar JSON (Excel)", type=["xlsx"], key=f"mst_{st.session_state.id_proceso}")
        if up_xls:
            df = pd.read_excel(up_xls)
            for _, r in df.iterrows():
                nm = str(r.get("NOMBRE","")).upper().strip()
                if nm: st.session_state.memoria["empresas"][nm] = {"DETALLE":str(r.get("DETALLE","OTROS")).upper(),"MEMO":str(r.get("MEMO","PROFESIONAL")).upper()}
            guardar_memoria(); st.success("Memoria guardada.")
        st.markdown("---")

    # 3. BUZÓN DE SUGERENCIAS
    st.subheader("📬 Sugerencias")
    sug_text = st.text_area("¿Cómo podemos mejorar?", key="txt_sugerencia")
    if st.button("Enviar Sugerencia", use_container_width=True):
        if sug_text:
            registrar_actividad(st.session_state.usuario_actual, "SUGERENCIA", sugerencia=sug_text)
            st.success("¡Gracias!")
        else: st.warning("Escribe algo primero.")

    st.markdown("---")

    # 4. CERRAR SESIÓN
    if st.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state.autenticado = False; st.rerun()

# --- CUERPO PRINCIPAL ---
st.subheader("💎 Gana Meses PRO")
inv = st.session_state.invitaciones_disponibles
if inv > 0:
    with st.expander(f"🎁 Regalar Invitación ({inv} pases disponibles)"):
        email = st.text_input("Email de tu colega")
        if st.button("Generar Pase"):
            resp = conectar_api({"accion": "INVITAR", "usuario": st.session_state.usuario_actual, "invitado": email})
            if resp.get("exito"):
                st.success("¡Pase generado!")
                msg = urllib.parse.quote(f"🎁 Te regalo un pase para *RAPIDITO AI*.\n👤 Usuario: {email}\n🔑 Clave: Rapidito2026\n👉 https://pruebas1998.streamlit.app")
                st.markdown(f'<a href="https://wa.me/?text={msg}" target="_blank"><button style="background-color:#25D366;color:white;width:100%;font-weight:bold;padding:12px;border-radius:8px;border:none;cursor:pointer;">📲 Enviar por WhatsApp</button></a>', unsafe_allow_html=True)

tab_xml, tab_sri = st.tabs(["📂 Subir XMLs (Manual/ZIP)", "📡 Descarga SRI (TXT)"])

with tab_xml:
    c1, c2 = st.columns(2)
    with c1:
        up_c = st.file_uploader("Compras (XML o ZIP)", type=["xml", "zip"], accept_multiple_files=True, key=f"c_{st.session_state.id_proceso}")
        if up_c and st.button("Procesar Compras"):
            data = [extraer_datos_robusto(x) for x in procesar_archivos_entrada(up_c)]
            data = [d for d in data if d and d["TIPO"] in ["FC","NC"]]
            st.session_state.data_compras_cache = data
            st.download_button("📥 Excel Compras", generar_excel_multiexcel(data_compras=data), "Compras.xlsx")
    with c2:
        up_v = st.file_uploader("Ventas (XML o ZIP)", type=["xml", "zip"], accept_multiple_files=True, key=f"v_{st.session_state.id_proceso}")
        if up_v and st.button("Procesar Ventas"):
            raw = [extraer_datos_robusto(x) for x in procesar_archivos_entrada(up_v)]
            data = procesar_ventas_con_retenciones([d for d in raw if d])
            st.session_state.data_ventas_cache = data
            st.download_button("📥 Excel Ventas", generar_excel_multiexcel(data_ventas_ret=data), "Ventas.xlsx")

with tab_sri:
    up_txt = st.file_uploader("TXT del SRI", type=["txt"])
    if up_txt and st.button("Descargar Masivo"):
        claves = list(dict.fromkeys(re.findall(r'\d{49}', up_txt.read().decode("latin-1"))))
        bar, lst = st.progress(0), []
        for i, cl in enumerate(claves):
            try:
                r = requests.post(URL_WS, data=f'<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ec="http://ec.gob.sri.ws.autorizacion"><soapenv:Body><ec:autorizacionComprobante><claveAccesoComprobante>{cl}</claveAccesoComprobante></ec:autorizacionComprobante></soapenv:Body></soapenv:Envelope>', headers=HEADERS_WS, verify=False, timeout=5)
                if "<autorizaciones>" in r.text:
                    d = extraer_datos_robusto(io.BytesIO(r.content))
                    if d: lst.append(d)
            except: pass
            bar.progress((i+1)/len(claves))
        if lst: st.download_button("📊 Excel SRI", generar_excel_multiexcel(data_sri_lista=lst), "SRI_Masivo.xlsx")
                if lst: 
                    st.success(f"✅ Completado. {len(lst)} documentos.")
                    registrar_actividad(st.session_state.usuario_actual, f"EXCEL SRI {titulo}", len(lst))
                    c1, c2 = st.columns(2)
                    with c1: st.download_button(f"📦 ZIP {titulo}", zip_buffer.getvalue(), f"{titulo}.zip")
                    with c2: st.download_button(f"📊 Excel {titulo}", generar_excel_multiexcel(data_sri_lista=lst, sri_mode=tipo_filtro), f"{titulo}.xlsx")
                else: st.warning("No se encontraron documentos.")

    s1, s2, s3 = st.tabs(["Facturas", "Notas Crédito", "Retenciones"])
    with s1: bloque_sri("Facturas Recibidas", "FC", "sri_fc")
    with s2: bloque_sri("Notas de Crédito", "NC", "sri_nc")
    with s3: bloque_sri("Retenciones", "RET", "sri_ret")




