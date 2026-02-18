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

# ENLACES DE CONEXIÓN
URL_WS = "https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline?wsdl"
HEADERS_WS = {"Content-Type": "text/xml;charset=UTF-8","User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36"}
URL_API_VIRAL = "https://script.google.com/macros/s/AKfycbz3vRq203m7vcdor30hJiXuAGNGr8n_kM2dCpf63LW4KhaeY9wqAijBC473AwywYes/exec" 

# --- CONEXIÓN AL CEREBRO VIRAL ---
def conectar_api(payload):
    try:
        if "TU_URL" in URL_API_VIRAL: return {"exito": False, "mensaje": "Falta configurar URL_API_VIRAL."}
        r = requests.post(URL_API_VIRAL, json=payload, timeout=10)
        return r.json()
    except: return {"exito": False, "mensaje": "Error de conexión."}

def registrar_actividad(usuario, accion, cantidad=None, sugerencia=None):
    URL_LOGGING = "https://script.google.com/macros/s/AKfycbyk0CWehcUec47HTGMjqsCs0sTKa_9J3ZU_Su7aRxfwmNa76-dremthTuTPf-FswZY/exec"
    detalle = f"{accion} ({cantidad} XMLs)" if cantidad is not None else accion
    payload = {"usuario": str(usuario), "accion": str(detalle)}
    if sugerencia: payload["sugerencia"] = str(sugerencia)
    try: requests.post(URL_LOGGING, json=payload, timeout=5); return True
    except: return False

# --- 2. SISTEMA DE LOGIN Y ESTADO ---
if "autenticado" not in st.session_state: st.session_state.autenticado = False
if "id_proceso" not in st.session_state: st.session_state.id_proceso = 0
if "data_compras_cache" not in st.session_state: st.session_state.data_compras_cache = []
if "data_ventas_cache" not in st.session_state: st.session_state.data_ventas_cache = []
if "invitaciones_disponibles" not in st.session_state: st.session_state.invitaciones_disponibles = 0

if not st.session_state.autenticado:
    st.sidebar.title("🔐 Acceso RAPIDITO")
    u, p = st.sidebar.text_input("Usuario (Email)"), st.sidebar.text_input("Contraseña", type="password")
    if st.sidebar.button("Iniciar Sesión"):
        with st.spinner("Verificando..."):
            resp = conectar_api({"accion": "LOGIN", "usuario": u.strip(), "clave": p.strip()})
            if resp.get("exito"):
                st.session_state.autenticado, st.session_state.usuario_actual = True, u.strip()
                st.session_state.invitaciones_disponibles = resp.get("invitaciones", 0)
                registrar_actividad(u, "ENTRÓ AL PORTAL"); st.rerun()
            else: st.sidebar.error("Credenciales incorrectas")
    st.sidebar.info("ℹ️ Exclusivo por invitación.")
    st.stop()

# --- 3. MEMORIA DE APRENDIZAJE ---
if 'memoria' not in st.session_state:
    if os.path.exists("conocimiento_contable.json"):
        with open("conocimiento_contable.json", "r", encoding="utf-8") as f: st.session_state.memoria = json.load(f)
    else: st.session_state.memoria = {"empresas": {}}

def guardar_memoria():
    with open("conocimiento_contable.json", "w", encoding="utf-8") as f: json.dump(st.session_state.memoria, f, indent=4, ensure_ascii=False)

def procesar_archivos_entrada(lista):
    xmls = []
    for f in lista:
        if f.name.lower().endswith('.xml'): xmls.append(io.BytesIO(f.getvalue()))
        elif f.name.lower().endswith('.zip'):
            with zipfile.ZipFile(f) as z:
                for n in z.namelist():
                    if n.lower().endswith('.xml') and not n.startswith('__MACOSX'): xmls.append(io.BytesIO(z.read(n)))
    return xmls

# --- 4. MOTOR DE EXTRACCIÓN (LÓGICA 10/13 DÍGITOS INTEGRADA) ---
def extraer_datos_robusto(xml_file):
    try:
        xml_file.seek(0); tree = ET.parse(xml_file); root = tree.getroot(); xml_data = None
        for elem in root.iter():
            if 'comprobante' in elem.tag.lower() and elem.text and "<" in elem.text:
                xml_data = ET.fromstring(re.sub(r'<\?xml.*?\?>', '', elem.text).strip()); break
        if xml_data is None: xml_data = root

        def buscar(tags):
            for t in tags:
                f = xml_data.find(f".//{t}")
                if f is not None and f.text: return f.text.strip()
            return ""

        tipo_doc = "NC" if "notacredito" in xml_data.tag.lower() else "RET" if "retencion" in xml_data.tag.lower() else "LC" if "liquidacion" in xml_data.tag.lower() else "FC"
        razon_social = buscar(["razonSocial"]).upper()
        ruc_emisor = buscar(["ruc"])
        num_fact = f"{buscar(['estab']) or '000'}-{buscar(['ptoEmi']) or '000'}-{buscar(['secuencial']) or '000'}"
        fecha = buscar(["fechaEmision"])
        ruc_cli = buscar(["identificacionComprador", "identificacionSujetoRetenido"])
        nom_cli = buscar(["razonSocialComprador", "razonSocialSujetoRetenido"]).upper()

        # --- LÓGICA DE CLASIFICACIÓN 10/13 DÍGITOS ---
        len_id = len(ruc_cli)
        tipo_id = "RUC" if len_id == 13 else "CEDULA" if len_id == 10 else "OTRO"
        info_json = st.session_state.memoria["empresas"].get(razon_social)
        
        if len_id == 10:
            memo_final = "PERSONAL"
            detalle_final = info_json["DETALLE"] if info_json else "NO DEDUCIBLE"
        else:
            detalle_final = info_json["DETALLE"] if info_json else "OTROS"
            memo_final = info_json["MEMO"] if info_json else "PROFESIONAL"

        data = {
            "TIPO": tipo_doc, "TIPO DE DOCUMENTO": tipo_doc, "FECHA": fecha, "N. FACTURA": num_fact,
            "RUC": ruc_emisor, "NOMBRE": razon_social, "N AUTORIZACION": buscar(["numeroAutorizacion", "claveAcceso"]),
            "TIPO ID": tipo_id, "CONTRIBUYENTE": ruc_cli, "CLIENTE": nom_cli, "DETALLE": detalle_final, "MEMO": memo_final
        }
        
        # Mapeo de Mes
        if "/" in fecha:
            ms = {"01":"ENERO","02":"FEBRERO","03":"MARZO","04":"ABRIL","05":"MAYO","06":"JUNIO","07":"JULIO","08":"AGOSTO","09":"SEPTIEMBRE","10":"OCTUBRE","11":"NOVIEMBRE","12":"DICIEMBRE"}
            data["MES"] = ms.get(fecha.split('/')[1], "DESCONOCIDO")

        # Retenciones
        if tipo_doc == "RET":
            rt_renta, rt_iva, base_renta, base_iva = 0.0, 0.0, 0.0, 0.0
            sus = ""
            node = xml_data.find(".//numDocSustento")
            if node is not None and node.text:
                p = node.text.replace('-','')
                if len(p) >= 15: sus = f"{p[0:3]}-{p[3:6]}-{p[6:]}"
            for item in (xml_data.findall(".//impuesto") + xml_data.findall(".//retencion")):
                try:
                    c = item.find("codigo").text
                    v = float(item.find("valorRetenido").text or 0)
                    b = float(item.find("baseImponible").text or 0)
                    if c == "1": rt_renta += v; base_renta += b
                    elif c == "2": rt_iva += v; base_iva += b
                except: continue
            data.update({"ruc_recep": ruc_cli, "nomrecep": nom_cli, "fechaemi": fecha, "razonsocial": razon_social, "ruc_emisor": ruc_emisor, "numfact": sus, "numreten": num_fact, "baserenta": base_renta, "rt_renta": rt_renta, "baseiva": base_iva, "rt_iva": rt_iva, "TOTAL RET": rt_renta + rt_iva, "SUSTENTO": sus})
        
        # Facturas / NC
        else:
            m = -1 if tipo_doc == "NC" else 1
            b0, b12, i12, ice, prop, no_obj, exento, otra_b, otro_i = 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0
            for imp in xml_data.findall(".//totalImpuesto"):
                try:
                    c, cp = imp.find("codigo").text, imp.find("codigoPorcentaje").text
                    b, v = float(imp.find("baseImponible").text or 0)*m, float(imp.find("valor").text or 0)*m
                    if c == "2":
                        if cp == "0": b0 += b
                        elif cp in ["2","3","4","8","10"]: b12 += b; i12 += v
                        elif cp == "6": no_obj += b
                        elif cp == "7": exento += b
                        else: otra_b += b; otro_i += v
                    elif c == "3": ice += v
                except: continue
            
            total_val = 0.0
            for t_tag in ["importeTotal", "total", "valorModificado"]:
                f = xml_data.find(f".//{t_tag}")
                if f is not None: total_val = float(f.text) * m; break

            items = [d.find("descripcion").text for d in xml_data.findall(".//detalle") if d.find("descripcion") is not None]
            data.update({"OTRA BASE IVA": otra_b, "OTRO IVA": otro_i, "MONTO ICE": ice, "PROPINAS": prop, "EXENTO DE IVA": exento, "NO OBJ IVA": no_obj, "BASE. 0": b0, "BASE. 12 / 15": b12, "IVA.": i12, "TOTAL": total_val, "SUBDETALLE": " | ".join(items[:5])})
        return data
    except: return None

# --- 5. LÓGICA DE INTEGRACIÓN ---
def procesar_ventas_con_retenciones(lista):
    vts, rets = [], {}
    for d in lista:
        if d["TIPO"] == "FC": vts.append(d)
        elif d["TIPO"] == "RET" and d.get("SUSTENTO"): rets[d["SUSTENTO"]] = d
    res = []
    for v in vts:
        r = rets.get(v["N. FACTURA"], {})
        res.append({
            "MES": v.get("MES"), "FECHA": v["FECHA"], "N. FACTURA": v["N. FACTURA"], "TIPO ID": v["TIPO ID"],
            "RUC": v["RUC CLIENTE"], "CLIENTE": v["CLIENTE"], "DETALLE": "SERVICIOS", "MEMO": "PROFESIONAL",
            "BASE. 0": v.get("BASE. 0", 0), "BASE. 12 / 15": v.get("BASE. 12 / 15", 0), "IVA": v.get("IVA.", 0), "TOTAL": v.get("TOTAL", 0),
            "FECHA RET": r.get("fechaemi", ""), "N° RET": r.get("numreten", ""), "RET RENTA": r.get("rt_renta", 0), "RET IVA": r.get("rt_iva", 0), "TOTAL RET": r.get("TOTAL RET", 0)
        })
    return res

# --- 6. GENERADOR EXCEL MAESTRO (CON TODAS LAS PESTAÑAS) ---
def generar_excel_multiexcel(data_compras=None, data_ventas_ret=None, data_sri_lista=None, sri_mode=None):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book
        f_azul = wb.add_format({'bold':True,'align':'center','border':1,'bg_color':'#002060','font_color':'white'})
        f_amar = wb.add_format({'bold':True,'align':'center','border':1,'bg_color':'#FFD966'})
        f_verd = wb.add_format({'bold':True,'align':'center','border':1,'bg_color':'#92D050'})
        f_num = wb.add_format({'num_format':'_-$ * #,##0.00_-','border':1})
        f_tot = wb.add_format({'bold':True,'num_format':'_-$ * #,##0.00_-','border':1,'bg_color':'#EFEFEF'})
        texto_pie = "&LGenerado por RAPIDITO AI&Rrapidito.ec"

        if sri_mode:
            df = pd.DataFrame(data_sri_lista)
            if sri_mode == "NC": cols = ["NOMBRE","RUC","N AUTORIZACION","FECHA","TIPO DE DOCUMENTO","N. FACTURA","MES","RUC CLIENTE","CLIENTE","PROPINAS","BASE. 0","NO OBJ IVA","BASE. 12 / 15","IVA.","TOTAL"]
            elif sri_mode == "RET": cols = ["ruc_recep", "nomrecep", "fechaemi", "razonsocial", "ruc_emisor", "numfact", "numreten", "baserenta", "rt_renta", "baseiva", "rt_iva", "numautori"]
            else: cols = ["MES","FECHA","N. FACTURA","TIPO DE DOCUMENTO","RUC","TIPO ID","CONTRIBUYENTE","NOMBRE","DETALLE","MEMO","BASE. 0","BASE. 12 / 15","IVA.","TOTAL"]
            for c in cols: 
                if c not in df.columns: df[c] = ""
            ws = wb.add_worksheet("SRI")
            for i, c in enumerate(cols): ws.write(0, i, c, f_azul)
            for r, row in enumerate(df[cols].values, 1):
                for c, v in enumerate(row): ws.write(r, c, v, f_num if isinstance(v, (float,int)) else wb.add_format({'border':1}))
        else:
            if data_compras:
                df_c = pd.DataFrame(data_compras)
                # SE AGREGA TIPO ID
                orden_c = ["MES","FECHA","N. FACTURA","TIPO DE DOCUMENTO","RUC","TIPO ID","CONTRIBUYENTE","NOMBRE","DETALLE","MEMO","OTRA BASE IVA","OTRO IVA","MONTO ICE","PROPINAS","EXENTO DE IVA","NO OBJ IVA","BASE. 0","BASE. 12 / 15","IVA.","TOTAL","SUBDETALLE"]
                for c in orden_c: 
                    if c not in df_c.columns: df_c[c] = ""
                ws_c = wb.add_worksheet('COMPRAS')
                for i, c in enumerate(orden_c): ws_c.write(0, i, c, f_amar if i in range(10, 16) else f_azul)
                for r, row in enumerate(df_c[orden_c].values, 1):
                    for c, v in enumerate(row): ws_c.write(r, c, v, f_num if isinstance(v, (float,int)) else wb.add_format({'border':1}))
                
                # REPORTE ANUAL
                meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
                ws_ra = wb.add_worksheet('REPORTE ANUAL')
                cats=["VIVIENDA","SALUD","EDUCACION","ALIMENTACION","VESTIMENTA","TURISMO","NO DEDUCIBLE","SERVICIOS BASICOS"]
                for i, ct in enumerate(cats): ws_ra.write(1, i+2, ct.title(), f_azul)
                # Columnas de sumatoria (P, Q, O, N... ajustadas por la nueva columna TIPO ID)
                # BASE 0(Q), BASE 12/15(R), IVA(S), TOTAL(T)
                for r, mes in enumerate(meses):
                    ws_ra.write(r+3, 0, mes, f_num)
                    # Fórmula simplificada: Suma de base 0 + base 12 + IVA + Ice + Otros (Columnas Q a U en el nuevo orden)
                    f_pr = f"SUMIFS('COMPRAS'!$T:$T,'COMPRAS'!$A:$A,\"{mes}\",'COMPRAS'!$J:$J,\"PROFESIONAL\")"
                    ws_ra.write_formula(r+3, 1, "="+f_pr, f_num)
                    for cidx, cat in enumerate(cats):
                        f_pe = f"SUMIFS('COMPRAS'!$T:$T,'COMPRAS'!$A:$A,\"{mes}\",'COMPRAS'!$I:$I,\"{cat}\")"
                        ws_ra.write_formula(r+3, cidx+2, "="+f_pe, f_num)

            if data_ventas_ret:
                df_v = pd.DataFrame(data_ventas_ret)
                orden_v = ["MES","FECHA","N. FACTURA","TIPO ID","RUC","CLIENTE","DETALLE","MEMO","BASE. 0","BASE. 12 / 15","IVA","TOTAL","N° RET","RET RENTA","RET IVA","TOTAL RET"]
                for c in orden_v: 
                    if c not in df_v.columns: df_v[c] = ""
                ws_v = wb.add_worksheet('VENTAS')
                for i, c in enumerate(orden_v): ws_v.write(0, i, c, f_verd if i >= 12 else f_azul)
                for r, row in enumerate(df_v[orden_v].values, 1):
                    for c, v in enumerate(row): ws_v.write(r, c, v, f_num if isinstance(v, (float,int)) else wb.add_format({'border':1}))

    return output.getvalue()

# --- 7. INTERFAZ ---
st.title(f"🚀 RAPIDITO AI - {st.session_state.usuario_actual}")

with st.sidebar:
    if st.button("🧹 NUEVO INFORME", type="primary"):
        st.session_state.id_proceso += 1; st.session_state.data_compras_cache = []; st.session_state.data_ventas_cache = []; st.rerun()
    st.markdown("---")
    if st.session_state.usuario_actual == "GABRIEL":
        st.header("Master Config")
        up_xls = st.file_uploader("Cargar JSON (Excel)", type=["xlsx"])
        if up_xls:
            df = pd.read_excel(up_xls)
            for _, r in df.iterrows():
                nm = str(r.get("NOMBRE","")).upper().strip()
                if nm: st.session_state.memoria["empresas"][nm] = {"DETALLE":str(r.get("DETALLE","OTROS")).upper(),"MEMO":str(r.get("MEMO","PROFESIONAL")).upper()}
            guardar_memoria(); st.success("Memoria guardada.")

    st.header("📬 Sugerencias")
    sug = st.text_area("Mejoras:")
    if st.button("Enviar"): registrar_actividad(st.session_state.usuario_actual, "SUGERENCIA", sugerencia=sug); st.success("Enviado")

    if st.button("Cerrar Sesión"): st.session_state.autenticado = False; st.rerun()

tab_xml, tab_sri = st.tabs(["📂 Procesar XMLs", "📡 SRI Masivo"])

with tab_xml:
    st1, st2, st3 = st.tabs(["🛒 Compras", "💰 Ventas", "📑 Integral"])
    with st1:
        up = st.file_uploader("XML/ZIP Compras", type=["xml","zip"], accept_multiple_files=True, key=f"c_{st.session_state.id_proceso}")
        if up and st.button("Generar Compras"):
            data = [extraer_datos_robusto(x) for x in procesar_archivos_entrada(up)]
            data = [d for d in data if d and d["TIPO"] in ["FC","NC"]]
            st.session_state.data_compras_cache = data
            st.download_button("📥 Excel Compras", generar_excel_multiexcel(data_compras=data), "Compras.xlsx")
    with st2:
        up = st.file_uploader("XML/ZIP Ventas", type=["xml","zip"], accept_multiple_files=True, key=f"v_{st.session_state.id_proceso}")
        if up and st.button("Generar Ventas"):
            raw = [extraer_datos_robusto(x) for x in procesar_archivos_entrada(up)]
            data = procesar_ventas_con_retenciones([d for d in raw if d])
            st.session_state.data_ventas_cache = data
            st.download_button("📥 Excel Ventas", generar_excel_multiexcel(data_ventas_ret=data), "Ventas.xlsx")
    with st3:
        if st.button("Generar Informe Integral"):
            if st.session_state.data_compras_cache and st.session_state.data_ventas_cache:
                st.download_button("📥 INFORME INTEGRAL", generar_excel_multiexcel(st.session_state.data_compras_cache, st.session_state.data_ventas_cache), "Integral.xlsx")
            else: st.warning("Procese Compras y Ventas primero.")

with tab_sri:
    up_txt = st.file_uploader("TXT SRI", type=["txt"])
    if up_txt and st.button("Descarga SRI"):
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
        if lst: st.download_button("📊 Excel SRI", generar_excel_multiexcel(data_sri_lista=lst, sri_mode="FC"), "SRI.xlsx")
