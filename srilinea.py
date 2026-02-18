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
        if "TU_URL" in URL_API_VIRAL: return {"exito": False, "mensaje": "Falta URL API"}
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

# Cache para descargas SRI para que no desaparezcan los botones
if "sri_results" not in st.session_state: st.session_state.sri_results = {}

if not st.session_state.autenticado:
    st.sidebar.title("🔐 Acceso RAPIDITO")
    u, p = st.sidebar.text_input("Usuario"), st.sidebar.text_input("Clave", type="password")
    if st.sidebar.button("Entrar"):
        resp = conectar_api({"accion": "LOGIN", "usuario": u.strip(), "clave": p.strip()})
        if resp.get("exito"):
            st.session_state.autenticado, st.session_state.usuario_actual = True, u.strip()
            st.session_state.invitaciones_disponibles = resp.get("invitaciones", 0)
            registrar_actividad(u, "LOGIN"); st.rerun()
        else: st.sidebar.error("Error de acceso")
    st.stop()

# --- 3. MEMORIA JSON ---
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

# --- 4. MOTOR DE EXTRACCIÓN (LÓGICA 10 DÍGITOS) ---
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

        tipo = "NC" if "notacredito" in xml_data.tag.lower() else "RET" if "retencion" in xml_data.tag.lower() else "FC"
        razon_social = buscar(["razonSocial"]).upper()
        ruc_emisor = buscar(["ruc"])
        num_fact = f"{buscar(['estab']) or '000'}-{buscar(['ptoEmi']) or '000'}-{buscar(['secuencial']) or '000'}"
        fecha = buscar(["fechaEmision"])
        ruc_cli = buscar(["identificacionComprador", "identificacionSujetoRetenido"])
        nom_cli = buscar(["razonSocialComprador", "razonSocialSujetoRetenido"]).upper()

        # CLASIFICACIÓN 10 VS 13 DÍGITOS
        len_id = len(ruc_cli)
        info_json = st.session_state.memoria["empresas"].get(razon_social)
        
        if len_id == 10:
            memo_final = "PERSONAL"
            detalle_final = info_json["DETALLE"] if info_json else "No Deducible"
        else:
            detalle_final = info_json["DETALLE"] if info_json else "OTROS"
            memo_final = info_json["MEMO"] if info_json else "PROFESIONAL"

        data = {
            "TIPO": tipo, "TIPO DE DOCUMENTO": tipo, "FECHA": fecha, "N. FACTURA": num_fact,
            "RUC": ruc_emisor, "NOMBRE": razon_social, "N AUTORIZACION": buscar(["numeroAutorizacion", "claveAcceso"]),
            "CONTRIBUYENTE": ruc_cli, "RUC CLIENTE": ruc_cli, "CLIENTE": nom_cli,
            "DETALLE": detalle_final, "MEMO": memo_final
        }
        
        if "/" in fecha:
            ms = {"01":"ENERO","02":"FEBRERO","03":"MARZO","04":"ABRIL","05":"MAYO","06":"JUNIO","07":"JULIO","08":"AGOSTO","09":"SEPTIEMBRE","10":"OCTUBRE","11":"NOVIEMBRE","12":"DICIEMBRE"}
            data["MES"] = ms.get(fecha.split('/')[1], "DESCONOCIDO")

        if tipo == "RET":
            r_renta, r_iva, b_renta, b_iva = 0.0, 0.0, 0.0, 0.0
            node = xml_data.find(".//numDocSustento")
            sus = node.text.replace('-','') if (node is not None and node.text) else ""
            if len(sus) >= 15: sus = f"{sus[0:3]}-{sus[3:6]}-{sus[6:]}"
            for item in (xml_data.findall(".//impuesto") + xml_data.findall(".//retencion")):
                try:
                    c = item.find("codigo").text
                    v = float(item.find("valorRetenido").text or 0)
                    b = float(item.find("baseImponible").text or 0)
                    if c == "1": r_renta += v; b_renta += b
                    elif c == "2": r_iva += v; b_iva += b
                except: continue
            data.update({"numfact": sus, "numreten": num_fact, "baserenta": b_renta, "rt_renta": r_renta, "baseiva": b_iva, "rt_iva": r_iva, "TOTAL RET": r_renta+r_iva, "SUSTENTO": sus, "fechaemi": fecha, "razonsocial": razon_social, "ruc_emisor": ruc_emisor})
            return data
        else:
            m = -1 if tipo == "NC" else 1
            b0, b12, i12, ice, no_obj, exento = 0.0, 0.0, 0.0, 0.0, 0.0, 0.0
            for imp in xml_data.findall(".//totalImpuesto"):
                try:
                    c, cp = imp.find("codigo").text, imp.find("codigoPorcentaje").text
                    b, v = float(imp.find("baseImponible").text or 0)*m, float(imp.find("valor").text or 0)*m
                    if c == "2":
                        if cp == "0": b0 += b
                        elif cp in ["2","3","4","8","10"]: b12 += b; i12 += v
                        elif cp == "6": no_obj += b
                        elif cp == "7": exento += b
                except: continue
            
            total_val = 0.0
            for t_tag in ["importeTotal", "total", "valorModificado"]:
                f = xml_data.find(f".//{t_tag}")
                if f is not None: total_val = float(f.text) * m; break
            
            items = [d.find("descripcion").text for d in xml_data.findall(".//detalle") if d.find("descripcion") is not None]
            data.update({"BASE. 0": b0, "BASE. 12 / 15": b12, "IVA.": i12, "TOTAL": total_val, "NO OBJ IVA": no_obj, "EXENTO DE IVA": exento, "SUBDETALLE": " | ".join(items[:5])})
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
            "MES": v.get("MES"), "FECHA": v["FECHA"], "N. FACTURA": v["N. FACTURA"],
            "RUC": v["RUC CLIENTE"], "CLIENTE": v["CLIENTE"], "DETALLE": "SERVICIOS", "MEMO": "PROFESIONAL",
            "BASE. 0": v.get("BASE. 0", 0), "BASE. 12 / 15": v.get("BASE. 12 / 15", 0), "IVA": v.get("IVA.", 0), "TOTAL": v.get("TOTAL", 0),
            "FECHA RET": r.get("fechaemi", ""), "N° RET": r.get("numreten", ""), "RET RENTA": r.get("rt_renta", 0), "RET IVA": r.get("rt_iva", 0), "TOTAL RET": r.get("TOTAL RET", 0)
        })
    return res

# --- 6. GENERADOR EXCEL ---
def generar_excel_multiexcel(data_compras=None, data_ventas_ret=None, data_sri_lista=None, sri_mode=None):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book
        f_azul = wb.add_format({'bold':True,'align':'center','border':1,'bg_color':'#002060','font_color':'white'})
        f_amar = wb.add_format({'bold':True,'align':'center','border':1,'bg_color':'#FFD966'})
        f_num = wb.add_format({'num_format':'_-$ * #,##0.00_-','border':1})
        texto_pie = "&LGenerado por RAPIDITO AI&Rrapidito.ec"

        if data_sri_lista:
            df = pd.DataFrame(data_sri_lista)
            cols = ["MES","FECHA","N. FACTURA","TIPO DE DOCUMENTO","RUC","CONTRIBUYENTE","NOMBRE","DETALLE","MEMO","BASE. 0","BASE. 12 / 15","IVA.","TOTAL"]
            if sri_mode == "RET": cols = ["fechaemi", "razonsocial", "ruc_emisor", "numfact", "numreten", "baserenta", "rt_renta", "baseiva", "rt_iva"]
            for c in cols: 
                if c not in df.columns: df[c] = ""
            ws = wb.add_worksheet('REPORTE')
            for i, c in enumerate(cols): ws.write(0, i, c, f_azul)
            for r, row in enumerate(df[cols].values, 1):
                for c, v in enumerate(row): ws.write(r, c, v, f_num if isinstance(v, (float,int)) else wb.add_format({'border':1}))
        else:
            if data_compras:
                df_c = pd.DataFrame(data_compras)
                orden_c = ["MES","FECHA","N. FACTURA","TIPO DE DOCUMENTO","RUC","CONTRIBUYENTE","NOMBRE","DETALLE","MEMO","BASE. 0","BASE. 12 / 15","IVA.","TOTAL"]
                ws_c = wb.add_worksheet('COMPRAS')
                for i, c in enumerate(orden_c): ws_c.write(0, i, c, f_azul)
                for r, row in enumerate(df_c[orden_c].values, 1):
                    for c, v in enumerate(row): ws_c.write(r, c, v, f_num if isinstance(v, (float,int)) else wb.add_format({'border':1}))
            if data_ventas_ret:
                df_v = pd.DataFrame(data_ventas_ret)
                ws_v = wb.add_worksheet('VENTAS')
                for i, c in enumerate(df_v.columns): ws_v.write(0, i, c, f_azul)
                for r, row in enumerate(df_v.values, 1):
                    for c, v in enumerate(row): ws_v.write(r, c, v, f_num if isinstance(v, (float,int)) else wb.add_format({'border':1}))
    return output.getvalue()

# --- 7. INTERFAZ ORGANIZADA ---
st.title(f"🚀 RAPIDITO AI - {st.session_state.get('usuario_actual', 'Portal Contable')}")

with st.sidebar:
    st.header("⚙️ Panel de Control")
    if st.button("🧹 NUEVO INFORME", type="primary", use_container_width=True):
        st.session_state.id_proceso += 1
        st.session_state.data_compras_cache, st.session_state.data_ventas_cache = [], []
        st.session_state.sri_results = {}
        st.rerun()
    st.markdown("---")
    if st.session_state.usuario_actual == "GABRIEL":
        st.subheader("🔑 Master Config")
        up_xls = st.file_uploader("Actualizar JSON", type=["xlsx"], key=f"mst_{st.session_state.id_proceso}")
        if up_xls:
            df = pd.read_excel(up_xls)
            for _, r in df.iterrows():
                nm = str(r.get("NOMBRE","")).upper().strip()
                if nm: st.session_state.memoria["empresas"][nm] = {"DETALLE":str(r.get("DETALLE","OTROS")).upper(),"MEMO":str(r.get("MEMO","PROFESIONAL")).upper()}
            guardar_memoria(); st.success("Guardado.")
    st.subheader("📬 Sugerencias")
    sug = st.text_area("Ideas:")
    if st.button("Enviar Sugerencia", use_container_width=True):
        if sug: registrar_actividad(st.session_state.usuario_actual, "SUGERENCIA", sugerencia=sug); st.success("Enviado")
    st.markdown("---")
    if st.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state.autenticado = False; st.rerun()

# CUERPO PRINCIPAL
st.subheader("💎 Gana Meses PRO")
inv = st.session_state.invitaciones_disponibles
if inv > 0:
    with st.expander(f"🎁 Regalar Invitación ({inv} disponibles)"):
        email = st.text_input("Email colega")
        if st.button("Generar Pase"):
            resp = conectar_api({"accion": "INVITAR", "usuario": st.session_state.usuario_actual, "invitado": email})
            if resp.get("exito"):
                msg = urllib.parse.quote(f"🎁 Pase exclusivo *RAPIDITO AI*.\n👤 Usuario: {email}\n🔑 Clave: Rapidito2026\n👉 https://pruebas1998.streamlit.app")
                st.markdown(f'<a href="https://wa.me/?text={msg}" target="_blank"><button style="background-color:#25D366;color:white;width:100%;font-weight:bold;padding:10px;border-radius:8px;border:none;">📲 WhatsApp</button></a>', unsafe_allow_html=True)

tab_manual, tab_sri = st.tabs(["📂 Subir XMLs (Manual/ZIP)", "📡 Descarga SRI (TXT)"])

# --- TAB MANUAL ---
with tab_manual:
    m1, m2, m3 = st.tabs(["🛒 Compras y NC", "💰 Ventas y Retenciones", "📑 Informe Integral"])
    with m1:
        up_c = st.file_uploader("Compras (XML/ZIP)", type=["xml","zip"], accept_multiple_files=True, key=f"c_{st.session_state.id_proceso}")
        if up_c and st.button("Procesar Compras"):
            data = [extraer_datos_robusto(x) for x in procesar_archivos_entrada(up_c)]
            data = [d for d in data if d and d["TIPO"] in ["FC","NC"]]
            st.session_state.data_compras_cache = data
            st.download_button("📥 Excel Compras", generar_excel_multiexcel(data_compras=data), "Compras.xlsx")
    with m2:
        up_v = st.file_uploader("Ventas (XML/ZIP)", type=["xml","zip"], accept_multiple_files=True, key=f"v_{st.session_state.id_proceso}")
        if up_v and st.button("Procesar Ventas"):
            raw = [extraer_datos_robusto(x) for x in procesar_archivos_entrada(up_v)]
            data = procesar_ventas_con_retenciones([d for d in raw if d])
            st.session_state.data_ventas_cache = data
            st.download_button("📥 Excel Ventas", generar_excel_multiexcel(data_ventas_ret=data), "Ventas.xlsx")
    with m3:
        if st.button("🚀 Generar Informe Integral"):
            if st.session_state.data_compras_cache and st.session_state.data_ventas_cache:
                st.download_button("📥 DESCARGAR INTEGRAL", generar_excel_multiexcel(st.session_state.data_compras_cache, st.session_state.data_ventas_cache), "Integral.xlsx")
            else: st.error("Primero procesa Compras y Ventas.")

# --- TAB SRI (CON BOTONES PERSISTENTES) ---
with tab_sri:
    def bloque_sri_persistente(titulo, tipo_filtro, key):
        st.subheader(titulo)
        up = st.file_uploader(f"TXT {titulo}", type=["txt"], key=f"up_{key}")
        
        # Botón para disparar proceso
        if up and st.button(f"🚀 Descargar {titulo}", key=f"btn_{key}"):
            claves = list(dict.fromkeys(re.findall(r'\d{49}', up.read().decode("latin-1"))))
            if claves:
                bar, status = st.progress(0), st.empty()
                lst, zip_buffer = [], io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zf:
                    for i, cl in enumerate(claves):
                        try:
                            r = requests.post(URL_WS, data=f'<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ec="http://ec.gob.sri.ws.autorizacion"><soapenv:Body><ec:autorizacionComprobante><claveAccesoComprobante>{cl}</claveAccesoComprobante></ec:autorizacionComprobante></soapenv:Body></soapenv:Envelope>', headers=HEADERS_WS, verify=False, timeout=5)
                            if "<autorizaciones>" in r.text:
                                zf.writestr(f"{cl}.xml", r.text)
                                d = extraer_datos_robusto(io.BytesIO(r.content))
                                if d and (d["TIPO"] == tipo_filtro or (tipo_filtro=="FC" and d["TIPO"]=="LC")): lst.append(d)
                        except: pass
                        bar.progress((i+1)/len(claves))
                        status.text(f"Procesando {i+1} de {len(claves)}")
                
                # Guardamos en session_state para que no se borren los botones
                st.session_state.sri_results[key] = {"data": lst, "zip": zip_buffer.getvalue()}

        # Mostrar botones si existen resultados en el estado
        if key in st.session_state.sri_results:
            res = st.session_state.sri_results[key]
            if res["data"]:
                st.success(f"✅ {len(res['data'])} documentos listos.")
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(f"📊 Excel {titulo}", generar_excel_multiexcel(data_sri_lista=res["data"], sri_mode=tipo_filtro), f"{titulo}.xlsx", key=f"dl_ex_{key}")
                with col2:
                    st.download_button(f"📦 ZIP XMLs {titulo}", res["zip"], f"{titulo}.zip", key=f"dl_zip_{key}")
            else:
                st.warning("No se encontraron documentos válidos en el archivo.")

    s1, s2, s3 = st.tabs(["Facturas", "Notas Crédito", "Retenciones"])
    with s1: bloque_sri_persistente("Facturas Recibidas", "FC", "sri_fc")
    with s2: bloque_sri_persistente("Notas de Crédito", "NC", "sri_nc")
    with s3: bloque_sri_persistente("Retenciones", "RET", "sri_ret")
