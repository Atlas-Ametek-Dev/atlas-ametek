import streamlit as st
from supabase import create_client, Client
import io
import time
from pypdf import PdfReader
from deep_translator import GoogleTranslator
import streamlit.components.v1 as components
import pandas as pd
import base64
from streamlit_pdf_viewer import pdf_viewer
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from fpdf import FPDF
import fitz  # PyMuPDF
import requests
from PIL import Image

# --- 1. CONFIGURACIÓN INICIAL ---
st.set_page_config(page_title="Gestión de Manufactura", layout="wide")

# --- CONEXIÓN A SUPABASE ---
SUPABASE_URL = st.secrets["SUPABASE_URL"].strip()
SUPABASE_KEY = st.secrets["SUPABASE_KEY"].strip()

@st.cache_resource
def init_connection():
    return create_client(SUPABASE_URL, SUPABASE_KEY)

supabase: Client = init_connection()

@st.cache_data(ttl=1800, show_spinner=False)
def descargar_bytes_pdf(url_pdf):
    """Descarga los bytes del archivo PDF con caché para carga rápida."""
    if not url_pdf or not str(url_pdf).startswith("http"):
        return None
    try:
        resp = requests.get(url_pdf, timeout=15)
        if resp.status_code == 200:
            return resp.content
    except Exception as e:
        st.error(f"Error al descargar PDF: {e}")
    return None

def generar_pdf_reporte_po(po_numero, piezas_lista, comentario_po=""):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Título principal
    pdf.set_font("Arial", 'B', 16)
    pdf.set_text_color(0, 74, 153)
    pdf.cell(0, 10, f"PO: {po_numero}", ln=True)
    
    # Nota de la PO
    if comentario_po:
        pdf.set_font("Arial", 'I', 10)
        pdf.set_text_color(146, 64, 14)
        pdf.multi_cell(0, 6, f"Nota de la PO: {comentario_po}")
        pdf.ln(3)
    else:
        pdf.ln(4)
        
    # Línea separadora
    pdf.set_draw_color(0, 74, 153)
    pdf.set_line_width(0.5)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(6)
    
    # Detalle de piezas
    for idx, pz in enumerate(piezas_lista):
        pn = str(pz.get("part_num", "-"))
        cant = pz.get("cantidad", 1)
        fecha = str(pz.get("fecha", "-"))
        
        estados_raw = pz.get("estados", [])
        if isinstance(estados_raw, list):
            estados_txt = ", ".join(estados_raw) if estados_raw else "En proceso"
        else:
            estados_txt = str(estados_raw or "En proceso")
            
        nota_pz = str(pz.get("comentario_pz", "") or "").strip()
        
        pdf.set_font("Arial", 'B', 12)
        pdf.set_text_color(30, 58, 138)
        pdf.cell(0, 7, f"#{pn}", ln=True)
        
        pdf.set_font("Arial", '', 10)
        pdf.set_text_color(55, 65, 81)
        pdf.cell(0, 6, f"   Cantidad: {cant} pzas", ln=True)
        pdf.cell(0, 6, f"   Entrega: {fecha}", ln=True)
        pdf.cell(0, 6, f"   Estados: {estados_txt}", ln=True)
        
        if nota_pz:
            pdf.set_font("Arial", 'I', 10)
            pdf.set_text_color(180, 83, 9)
            pdf.multi_cell(0, 6, f"   Nota Pieza: {nota_pz}")
            
        pdf.ln(4)
        pdf.set_draw_color(229, 231, 235)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.ln(4)
        
    # Retorno compatible con fpdf y fpdf2
    salida = pdf.output(dest='S')
    if isinstance(salida, str):
        return salida.encode('latin-1', errors='replace')
    return bytes(salida)

def generar_pdf_requisicion(po_number, partidas):
    """Genera la requisición formal en PDF horizontal con hipervínculos funcionales y celdas autoajustables."""
    partidas_herramientas = [
        it for it in partidas 
        if any(term in str(it.get("tipo", "")).lower() for term in ["herramienta", "tool", "inserto", "cortador", "broca", "machuelo", "fresa"])
    ]
    partidas_a_procesar = partidas_herramientas if partidas_herramientas else partidas

    pdf = FPDF(orientation="L", unit="mm", format="Letter")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_margins(left=14, top=12, right=14)
    pdf.add_page()

    # --- ENCABEZADO SUPERIOR ---
    pdf.set_font("Helvetica", "B", 16)
    pdf.set_text_color(0, 74, 153)
    pdf.cell(0, 8, "REQUISICION DE COMPRA (HERRAMIENTAS) - SHIRSA MAQUINADOS", ln=True, align="C")

    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(50, 50, 50)
    pdf.cell(0, 6, f"Orden de Produccion / PO de Referencia: {po_number}", ln=True, align="C")

    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(110, 110, 110)
    pdf.cell(0, 5, f"Fecha de Aprobacion: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True, align="C")
    pdf.ln(5)

    w_cols = [15, 15, 22, 60, 48, 25, 65]
    encabezados = ["Cant.", "Und.", "Tipo", "Descripcion Tecnica / Modelo", "Medida / Dimension", "Enlace", "Observaciones"]

    def dibujar_encabezado():
        pdf.set_fill_color(0, 74, 153)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 9)
        for i, enc in enumerate(encabezados):
            pdf.cell(w_cols[i], 8, enc, border=1, ln=(1 if i == len(encabezados) - 1 else 0), align="C", fill=True)
        pdf.set_font("Helvetica", "", 8.5)
        pdf.set_text_color(30, 30, 30)

    dibujar_encabezado()

    alto_linea = 4.5
    for it in partidas_a_procesar:
        raw_link = str(it.get("link", "") or "").strip()
        
        if raw_link.startswith("http"):
            if "mcmaster" in raw_link.lower():
                texto_link = "McMaster"
            elif "mscdirect" in raw_link.lower():
                texto_link = "MSC Direct"
            else:
                texto_link = "Ver Link"
        else:
            texto_link = "-"

        valores = [
            str(it.get("cantidad", 1)),
            str(it.get("unidad", "PZA")),
            str(it.get("tipo", "")),
            str(it.get("descripcion", "")),
            str(it.get("medida", "") or "-"),
            texto_link,
            str(it.get("motivo_uso", "") or "-")
        ]

        pdf.set_font("Helvetica", "", 8.5)
        lineas_celdas = []
        for i, val in enumerate(valores):
            try:
                lineas = pdf.multi_cell(w_cols[i], alto_linea, val, dry_run=True, output="LINES")
                num_lineas = max(len(lineas), 1)
            except TypeError:
                ancho_t = pdf.get_string_width(val)
                num_lineas = max(int(ancho_t / max(w_cols[i] - 4, 1)) + 1, len(val.split("\n")), 1)
            lineas_celdas.append(num_lineas)

        altura_fila = max(max(lineas_celdas) * alto_linea + 3, 7.5)

        if pdf.get_y() + altura_fila > pdf.page_break_trigger:
            pdf.add_page()
            dibujar_encabezado()

        pos_x_inicio = pdf.get_x()
        pos_y_inicio = pdf.get_y()

        for i, val in enumerate(valores):
            pdf.rect(pos_x_inicio, pos_y_inicio, w_cols[i], altura_fila)
            pdf.set_xy(pos_x_inicio, pos_y_inicio + 1.5)

            if i == 5:
                if raw_link.startswith("http"):
                    pdf.set_text_color(0, 102, 204)
                    pdf.set_font("Helvetica", "U", 8.5)
                    pdf.cell(w_cols[i], alto_linea, val, border=0, align="C", link=raw_link)
                    pdf.set_text_color(30, 30, 30)
                    pdf.set_font("Helvetica", "", 8.5)
                else:
                    pdf.cell(w_cols[i], alto_linea, "-", border=0, align="C")
            else:
                alineacion = "C" if i in [0, 1, 2] else "L"
                pdf.multi_cell(w_cols[i], alto_linea, val, border=0, align=alineacion)

            pos_x_inicio += w_cols[i]

        pdf.set_xy(pdf.l_margin, pos_y_inicio + altura_fila)

    pdf.ln(5)
    pdf.set_font("Helvetica", "I", 8.5)
    pdf.set_text_color(90, 90, 90)
    pdf.cell(0, 5, "Documento emitido y aprobado por Supervision de Taller para compra inmediata.", ln=True, align="L")

    return bytes(pdf.output())

def enviar_correo_gerente(po_number, pdf_bytes, destinatarios_to=None, destinatarios_cc=None):
    """Envía la requisición formal a destinatarios principales y en copia (CC)."""
    remitente = st.secrets["SMTP_USER"]
    password = st.secrets["SMTP_PASSWORD"]

    if not destinatarios_to:
        destinatarios_to = st.secrets.get("REQUISICION_TO", "ericslosberg@gmail.com")
    if not destinatarios_cc:
        destinatarios_cc = st.secrets.get("REQUISICION_CC", "shirsa.produccion@nae-supply.com, juan.rodriguez@nae-supply.com, diseno.dos.shirsa@nae-supply.com, soporte@nae-supply.com")

    lista_to = [c.strip() for c in destinatarios_to.split(",") if c.strip()]
    lista_cc = [c.strip() for c in destinatarios_cc.split(",") if c.strip()]
    todos_destinatarios = lista_to + lista_cc

    msg = MIMEMultipart()
    msg["From"] = remitente
    msg["To"] = ", ".join(lista_to)
    if lista_cc:
        msg["Cc"] = ", ".join(lista_cc)
    msg["Subject"] = f"🔔 Requisición Aprobada para Compra - PO {po_number}"

    cuerpo = f"""Buen día Estimada Gerencia y Compras,

Favor de procesar la orden de compra de material / herramienta para la PO {po_number}.

La requisición ha sido validada contra inventario y aprobada por Supervisión de Taller.
En el archivo PDF horizontal adjunto encontrarán la relación detallada de ítems, medidas y justificaciones.

Saludos cordiales,
Sistema de Control y Gestión Atlas - Shirsa Maquinados
"""
    msg.attach(MIMEText(cuerpo, "plain"))

    adjunto = MIMEApplication(pdf_bytes, _subtype="pdf")
    adjunto.add_header('Content-Disposition', 'attachment', filename=f"Requisicion_PO_{po_number}.pdf")
    msg.attach(adjunto)

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(remitente, password)
        server.sendmail(remitente, todos_destinatarios, msg.as_string())

@st.cache_data(ttl=600, show_spinner=False)
def obtener_pos():
    try:
        res = supabase.table("piezas").select("po_number").execute()
        pos_unicas = sorted(list(set(
            str(r["po_number"]).strip() 
            for r in res.data 
            if r.get("po_number")
        )))
        return {po: po for po in pos_unicas}
    except Exception as e:
        st.error(f"Error al obtener POs: {e}")
        return {}

st.title("🏭 Sistema de Gestión de Manufactura")

# --- 5. SISTEMA DE LOGIN CON USUARIOS DE SUPERVISOR ---
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.rol_actual = None
    st.session_state.usuario_actual = None

if not st.session_state.autenticado:
    st.sidebar.header("🔐 Iniciar Sesión")
    rol_elegido = st.sidebar.selectbox("Selecciona tu rol:", ["Desarrollador", "Maquinista", "Inventarios", "Supervisor", "Gerente"])
    
    # Selector de usuario específico si es Supervisor
    if rol_elegido == "Supervisor":
        usuario_elegido = st.sidebar.selectbox("Usuario Supervisor:", ["Jonathan", "Ilagor", "Alessandro"])
    else:
        usuario_elegido = rol_elegido

    password = st.sidebar.text_input("Contraseña:", type="password")

    # Contraseñas individuales de supervisor y generales de roles
    passwords_supervisores = {
        "Jonathan": "S2S3_Jon",
        "Ilagor": "S2S3_Ila",
        "Alessandro": "S2S3_Ale"
    }

    passwords_generales = {
        "Desarrollador": "120501",
        "Maquinista": "Maq",
        "Inventarios": "Inv2026",
        "Gerente": "GJE2026"
    }

    if st.sidebar.button("Entrar"):
        valido = False
        if rol_elegido == "Supervisor":
            # Permite la contraseña individual o la clave maestra 'S2S3'
            if password == passwords_supervisores.get(usuario_elegido) or password == "S2S3":
                valido = True
        else:
            if password == passwords_generales.get(rol_elegido):
                valido = True

        if valido:
            st.session_state.autenticado = True
            st.session_state.rol_actual = rol_elegido
            st.session_state.usuario_actual = usuario_elegido
            st.rerun()
        else:
            st.sidebar.error("Contraseña incorrecta.")
    st.stop()
else:
    rol = st.session_state.rol_actual
    usuario = st.session_state.usuario_actual
    st.sidebar.markdown(f"*Usuario:* {usuario} ({rol})")
    if st.sidebar.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.session_state.usuario_actual = None
        st.rerun()
# ==========================================
# VENTANA 1: DESARROLLADOR (COMPLETO)
# ==========================================
if rol == "Desarrollador":
    st.header("⚙️ Área de Desarrollo: Gestión de POs y Planos")

    tab_nuevos_planos, tab_hoja_recurrentes = st.tabs([
        "📤 Carga de Planos Nuevos (PDFs)", 
        "📋 Órdenes Recurrentes (Heredar Planos)"
    ])

    # ------------------------------------------------------------------
    # 1. CARGA MASIVA DE PLANOS NUEVOS (TRADUCCIÓN Y STORAGE)
    # ------------------------------------------------------------------
    with tab_nuevos_planos:
        st.subheader("Subir planos PDF para piezas nuevas")
        
        col_po_nueva, col_idioma = st.columns([2, 1])
        with col_po_nueva:
            po = st.text_input("Número de PO a trabajar:", placeholder="Ej: 50054061, 50054647, 50054832")
        with col_idioma:
            idioma_origen = st.selectbox("Idioma del plano para traducción:", ["de (Alemán)", "en (Inglés)"])
            cod_idioma = "de" if "Alemán" in idioma_origen else "en"

        archivos_pdf = st.file_uploader(
            "Sube los planos PDF (Originales y _TL)", 
            type="pdf", 
            accept_multiple_files=True
        )

        if st.button("Procesar y Guardar Todos"):
    
            if not po.strip():
                st.error("⚠️ Debes ingresar el número de PO.")
            elif not archivos_pdf:
                st.warning("⚠️ Debes subir al menos un archivo PDF.")
            else:
                procesados = {}
                for archivo in archivos_pdf:
                    nombre_archivo = archivo.name
                    nombre_sin_ext = archivo.name.replace(".pdf", "").replace(".PDF", "")
                    if "_TL" in nombre_archivo.upper():
                        base_name = nombre_sin_ext.upper().replace("_TL", "")
                        if base_name not in procesados:
                            procesados[base_name] = {}
                        procesados[base_name]['tl'] = archivo
                    else:
                        base_name = nombre_sin_ext.upper()
                        if base_name not in procesados:
                            procesados[base_name] = {}
                        procesados[base_name]['org'] = archivo
        
                for nombre_pieza, archivos in procesados.items():
                    with st.spinner(f'Procesando: {nombre_pieza}...'):
                        try:
                            timestamp = int(time.time())
                            
                            # 1. Procesar archivo original
                            url_org = ""
                            if 'org' in archivos:
                                archivo_org = archivos['org']
                                ruta = f"{nombre_pieza}_ORG_{timestamp}.pdf"
                                supabase.storage.from_("planos").upload(
                                    ruta, 
                                    archivo_org.getvalue(), 
                                    {"content-type": "application/pdf"}
                                )
                                url_org = supabase.storage.from_("planos").get_public_url(ruta)
        
                            # 2. Procesar archivo TL y traducción (únicamente si se subió un archivo _TL)
                            url_tl = ""
                            texto_traducido_tl = ""
                            texto = ""
                            
                            if 'tl' in archivos:
                                archivo_tl = archivos['tl']
                                
                                # Lectura del PDF
                                try:
                                    import pypdf
                                    import io
                                    reader = pypdf.PdfReader(io.BytesIO(archivo_tl.getvalue()))
                                    texto = "".join([p.extract_text() for p in reader.pages if p.extract_text()])
                                except Exception as e_pdf:
                                    st.warning(f"No se pudo extraer texto del PDF TL: {e_pdf}")
                                    texto = ""
        
                                # Traducción en bloques de menos de 450 caracteres
                                if texto.strip():
                                    try:
                                        time.sleep(0.4)
                                        from deep_translator import GoogleTranslator
                                        traductor = GoogleTranslator(source=cod_idioma, target='es')
                                        texto_traducido_tl = traductor.translate(texto[:2000].strip())
                                    except Exception:
                                        try:
                                            from deep_translator import MyMemoryTranslator
                                            
                                            mapa_idiomas_mymemory = {
                                                'de': 'german',
                                                'en': 'english',
                                                'fr': 'french',
                                                'it': 'italian'
                                            }
                                            idioma_origen_mm = mapa_idiomas_mymemory.get(cod_idioma, 'german')
                                            traductor_alt = MyMemoryTranslator(source=idioma_origen_mm, target='spanish')
                    
                                            lineas = [l.strip() for l in texto.split("\n") if l.strip()]
                                            partes_traducidas = []
                                            bloque_temp = ""
                    
                                            for linea in lineas:
                                                if len(bloque_temp) + len(linea) + 1 > 450:
                                                    if bloque_temp:
                                                        time.sleep(0.4)
                                                        partes_traducidas.append(traductor_alt.translate(bloque_temp))
                                                    bloque_temp = linea[:450]
                                                else:
                                                    bloque_temp = f"{bloque_temp} {linea}".strip()
                    
                                            if bloque_temp:
                                                time.sleep(0.4)
                                                partes_traducidas.append(traductor_alt.translate(bloque_temp))
                    
                                            texto_traducido_tl = "\n".join(partes_traducidas)
                                        except Exception as e_final:
                                            texto_traducido_tl = texto[:2000]
                                            st.warning(f"⚠️ No se pudo traducir: {e_final}")
                                else:
                                    texto_traducido_tl = ""
        
                                # Subir plano TL a Supabase Storage
                                ruta_tl = f"{nombre_pieza}_TL_{timestamp}.pdf"
                                supabase.storage.from_("planos").upload(
                                    ruta_tl, 
                                    archivo_tl.getvalue(), 
                                    {"content-type": "application/pdf"}
                                )
                                url_tl = supabase.storage.from_("planos").get_public_url(ruta_tl)
        
                            # 3. Guardar / Actualizar en la base de datos Supabase
                            supabase.table("piezas").upsert({
                                "po_number": po,
                                "Part number": nombre_pieza,
                                "url_pdf_org": url_org,
                                "url_pdf_tl": url_tl,
                                "texto_traducido": texto_traducido_tl
                            }, on_conflict="po_number, Part number").execute()
        
                            st.success(f"✅ Pieza {nombre_pieza} procesada y guardada correctamente.")
        
                        except Exception as e:
                            st.error(f"Error procesando {nombre_pieza}: {e}")

                            # 3. Guardar en la tabla unificada 'piezas'
                            supabase.table("piezas").insert({
                                "po_number": po.strip(),
                                "Part number": nombre_pieza,
                                "url_pdf_org": url_org,
                                "url_pdf_tl": url_tl,
                                "texto_traducido": texto_traducido_tl,
                                "cantidad_solicitada": 1,
                                "piezas_terminadas": 0,
                                "porcentaje_avance": 0
                            }).execute()

                            st.success(f"✅ Pieza '{nombre_pieza}' guardada en PO {po.strip()}.")
                        except Exception as e:
                            st.error(f"Error procesando {nombre_pieza}: {e}")

    # ------------------------------------------------------------------
    # 2. PIEZAS RECURRENTES / ÓRDENES DE PRODUCCIÓN (HOJAS AMETEK)
    # ------------------------------------------------------------------
    with tab_hoja_recurrentes:
        st.subheader("Cargar partida de hoja PO usando piezas anteriores")
        st.caption("Usa esta pestaña para órdenes como la PO 50054832 donde las piezas ya se habían maquinado.")

        try:
            # Traer historial de piezas para reutilizar planos y parámetros
            res_historico = supabase.table("piezas").select("*").execute()
            todas_piezas = res_historico.data

            if todas_piezas:
                # Diccionario único por Part number
                catalogo_base = {}
                for p in todas_piezas:
                    p_num = str(p.get("Part number", "")).strip()
                    if p_num and p_num not in catalogo_base:
                        catalogo_base[p_num] = p

                with st.form("form_hoja_ametek", clear_on_submit=False):
                    c_po, c_part = st.columns(2)
                    with c_po:
                        po_hoja = st.text_input("Número de PO de la hoja:", value="50054832")
                    with c_part:
                        part_sel = st.selectbox(
                            "Seleccionar Número de Parte existente:", 
                            sorted(list(catalogo_base.keys()))
                        )

                    c_qty, c_fec = st.columns(2)
                    with c_qty:
                        cantidad_pedida = st.number_input("Cantidad solicitada (Qty):", min_value=1, value=1, step=1)
                    with c_fec:
                        fecha_requerida = st.date_input("Fecha de entrega (Required Date):", value=datetime.today())

                    if st.form_submit_button("Crear Partida en PO"):
                        if not po_hoja.strip():
                            st.error("⚠️ Debes ingresar el número de PO.")
                        else:
                            molde = catalogo_base[part_sel]
                            
                            nueva_partida = {
                                "po_number": po_hoja.strip(),
                                "Part number": part_sel,
                                "url_pdf_org": molde.get("url_pdf_org", ""),
                                "url_pdf_tl": molde.get("url_pdf_tl", ""),
                                "texto_traducido": molde.get("texto_traducido", ""),
                                "herramienta_necesaria": molde.get("herramienta_necesaria", ""),
                                "acabado": molde.get("acabado", ""),
                                "tratamiento": molde.get("tratamiento", ""),
                                "cantidad_solicitada": cantidad_pedida,
                                "fecha_entrega": str(fecha_requerida),
                                "piezas_terminadas": 0,
                                "porcentaje_avance": 0
                            }
                            
                            try:
                                supabase.table("piezas").insert(nueva_partida).execute()
                                st.success(f"✅ Partida '{part_sel}' ligada a la PO {po_hoja.strip()} ({cantidad_pedida} pzas para el {fecha_requerida}).")
                            except Exception as err:
                                st.error(f"Error al ligar partida: {err}")

            else:
                st.info("Aún no hay piezas registradas en el historial para reutilizar.")

        except Exception as e_cat:
            st.error(f"Error al consultar el catálogo de piezas: {e_cat}")
            
# ==========================================
# VENTANA 2: MAQUINISTA (CÓDIGO COMPLETO)
# ==========================================
elif rol == "Maquinista":
    st.header("🛠️ Área de Maquinado")

    maquinistas = {"Abel": "8352", "Alejandro": "1497", "Francisco": "5206", "Isaac": "3984", "Jesus": "5731", "Rafael": "6475", "Edwin" : "5759","Alessandro":"1234"}
    col_l1, col_l2 = st.columns(2)
    usuario = col_l1.selectbox("Nombre:", list(maquinistas.keys()))
    password_ingresado = col_l2.text_input("Contraseña:", type="password")

    if password_ingresado and password_ingresado == maquinistas.get(usuario):
        st.success(f"Bienvenido, {usuario}.")
   
        dicc_pos = obtener_pos()
        po = st.selectbox("Seleccione la PO a trabajar:", list(dicc_pos.keys()))
        
        # Ahora filtras las piezas de esa PO:
        respuesta = supabase.table("piezas").select("*").eq("po_number", po).execute()
        piezas = respuesta.data
        
        nombres_piezas = {p['Part number']: p for p in piezas}
        pieza_seleccionada = st.selectbox("Selecciona la pieza a trabajar:", list(nombres_piezas.keys()))
        datos_pieza = nombres_piezas[pieza_seleccionada]
        
        # --- REGISTRO DE TRABAJO ---
        with st.expander("📝 Registrar piezas terminadas"):
            with st.form(key="form_registro"):
                piezas_realizadas = st.multiselect("¿Qué piezas terminaste?", nombres_piezas)
                notas_adicionales = st.text_area("Notas:")
                
                if st.form_submit_button("Registrar"):
                    lista_segura = [str(p) for p in piezas_realizadas]
                    notas_finales = str(notas_adicionales) if notas_adicionales else "Sin notas"
                    
                    for pieza_id in piezas_realizadas:
                        update_data = {
                            "maquinista_responsable": usuario,
                            "notas_registro": notas_finales
                        }
                        supabase.table("piezas").update(update_data).eq("Part number", pieza_id).execute()
                    
                    st.success("¡Registro actualizado en la base de datos!")
    
        tab1, tab2, tab3 = st.tabs(["📄 Plano Original", "⚙️ Especificaciones (TL)", "📝 Traducción"])

        with tab1: # Plano Original
            url_org = datos_pieza.get('url_pdf_org')
            if url_org and url_org.startswith("http"):
                st.link_button("↗️ Abrir plano en pestaña nueva", url_org)
                
                pdf_org_bytes = descargar_bytes_pdf(url_org)
                if pdf_org_bytes:
                    pdf_viewer(input=pdf_org_bytes, width=950, height=650)
                else:
                    st.warning("No se pudo cargar el archivo del plano.")
            else:
                st.warning("No hay plano original disponible.")

        with tab2: # Especificaciones TL
            url_tl = datos_pieza.get('url_pdf_tl')
            if url_tl and url_tl.startswith("http"):
                st.link_button("↗️ Abrir archivo TL en pestaña nueva", url_tl)
                
                pdf_tl_bytes = descargar_bytes_pdf(url_tl)
                if pdf_tl_bytes:
                    pdf_viewer(input=pdf_tl_bytes, width=950, height=650)
                else:
                    st.warning("No se pudo cargar el archivo TL.")
            else: 
                st.warning("No hay archivo TL disponible.")

        with tab3: # Traducción
            st.info(datos_pieza.get('texto_traducido') or "Sin traducción disponible.")
    
        # --- FORMULARIO DE ACTUALIZACIÓN ---
        st.subheader("Actualizar Parámetros")
        herramientas_bd = sorted(list(set(h.strip() for p in piezas if p.get('herramienta_necesaria') for h in str(p['herramienta_necesaria']).split('\n') if h.strip())))
        acabados_bd = sorted(list(set(str(p['acabado']).strip() for p in piezas if p.get('acabado') and str(p['acabado']).strip())))
        tratamientos_bd = sorted(list(set(str(p['tratamiento']).strip() for p in piezas if p.get('tratamiento') and str(p['tratamiento']).strip())))
    
        with st.form(key=f"form_{datos_pieza['Part number']}"):
            c1, c2, c3 = st.columns(3)
            with c1:
                herr_sel = st.multiselect("Herramientas existentes:", herramientas_bd, default=[h.strip() for h in str(datos_pieza.get('herramienta_necesaria', '')).split('\n') if h.strip() in herramientas_bd])
                herr_nuevas = st.text_area("Herramientas nuevas (Nombre - Medida - Material):", height=100)
            with c2:
                acabado_sel = st.selectbox("Acabado existente:", ["Ninguno"] + acabados_bd)
                acabado_nuevo = st.text_input("Nuevo acabado:")
                mat_recibido = st.checkbox("¿Material Recibido?", value=bool(datos_pieza.get('material_llegado')))
            with c3:
                trat_sel = st.selectbox("Tratamiento existente:", ["Ninguno"] + tratamientos_bd)
                trat_nuevo = st.text_input("Nuevo tratamiento:")
                comentarios_mat = st.text_area("Notas de material:", value=datos_pieza.get('notas_material') or "", height=68)
    
            if st.form_submit_button("Guardar Parámetros"):
                update_data = {
                    "herramienta_necesaria": '\n'.join(herr_sel + [h.strip() for h in herr_nuevas.split('\n') if h.strip()]),
                    "acabado": acabado_nuevo.strip() if acabado_nuevo.strip() != "" else (acabado_sel if acabado_sel != "Ninguno" else ""),
                    "tratamiento": trat_nuevo.strip() if trat_nuevo.strip() != "" else (trat_sel if trat_sel != "Ninguno" else ""),
                    "material_llegado": mat_recibido,
                    "notas_material": comentarios_mat.strip()
                }
                supabase.table("piezas").update(update_data).eq("Part number", datos_pieza['Part number']).execute()
                st.success("¡Datos guardados!")
                            
    elif password_ingresado:
        st.error("Contraseña incorrecta.")

# ==========================================
# VENTANA 3: INVENTARIOS (Inventario_Atlas)
# ==========================================
elif rol == "Inventarios":
    st.header("📦 Control de Almacén e Inventario Atlas")

    tab_alta, tab_consulta, tab_req, tab_bajas = st.tabs([
    "➕ Altas Atlas", 
    "🔍 Catálogo Atlas", 
    "📝 Requisiciones Atlas", 
    "🗑️ Bajas Atlas"
     ])

    # ----------------------------------------------------
    # TAB 1: ALTA DE NUEVOS ARTÍCULOS (CON REGISTRO EN TABLA SEPARADA)
    # ----------------------------------------------------
    with tab_alta:
        st.subheader("Registrar nuevo ítem en Inventario_Atlas")
        with st.form("form_alta_inventario", clear_on_submit=True):
            col_fec, col_sku, col_cat = st.columns(3)
            with col_fec:
                fecha_reg = st.date_input("Fecha de Registro / Entrada:", value=datetime.today())
            with col_sku:
                sku = st.text_input("SKU (Código único / ID):*")
            with col_cat:
                categoria = st.text_input("CATEGORIA (ej. Endmill, Broca, Inserto):")

            col1, col2, col3 = st.columns(3)
            with col1:
                modelo = st.text_input("MODELO:")
                stock = st.number_input("STOCK Inicial:", min_value=0, value=1, step=1)
            with col2:
                marca = st.text_input("MARCA:")
                unidad = st.text_input("UNIDAD:", value="PZA")
            with col3:
                medida = st.text_input("MEDIDA (ej. 1/2\", 5mm):")
                ubicacion = st.text_input("UBICACION (ej. Rack 6, Gaveta 2):")

            descripcion = st.text_area("DESCRIPCION:", placeholder="Descripción técnica...")
            archivo_imagen = st.file_uploader("Imagen de Referencia:", type=["jpg", "jpeg", "png"])
            status = st.checkbox("STATUS (Activo / Disponible)", value=True)

            if st.form_submit_button("Guardar en Inventario"):
                if not sku.strip():
                    st.error("⚠️ El campo 'SKU' es obligatorio.")
                else:
                    url_imagen = ""
                    if archivo_imagen:
                        try:
                            timestamp = int(time.time())
                            ext = archivo_imagen.name.split(".")[-1]
                            ruta_img = f"ref_{sku.strip()}_{timestamp}.{ext}"
                            supabase.storage.from_("planos").upload(
                                ruta_img, 
                                archivo_imagen.getvalue(), 
                                {"content-type": archivo_imagen.type}
                            )
                            url_imagen = supabase.storage.from_("planos").get_public_url(ruta_img)
                        except Exception as e_img:
                            st.warning(f"No se pudo guardar la imagen: {e_img}")

                    # 1. Registro en la tabla principal (Inventario_Atlas)
                    nuevo_registro = {
                        "SKU": sku.strip(),
                        "CATEGORIA": categoria.strip(),
                        "MODELO": modelo.strip(),
                        "DESCRIPCION": descripcion.strip(),
                        "REFERENCIA": url_imagen,
                        "UNIDAD": unidad.strip(),
                        "MARCA": marca.strip(),
                        "MEDIDA": medida.strip(),
                        "STOCK": str(stock),
                        "STATUS": bool(status),
                        "UBICACION": ubicacion.strip()
                    }

                    # 2. Registro en la tabla de bitácora separada (altas_inventario)
                    log_alta = {
                        "sku": sku.strip(),
                        "descripcion": descripcion.strip(),
                        "categoria": categoria.strip(),
                        "stock_inicial": str(stock),
                        "unidad": unidad.strip(),
                        "responsable": rol,
                        "fecha_alta": str(fecha_reg)
                    }

                    try:
                        # Guardar en catálogo principal
                        supabase.table("Inventario_Atlas").insert(nuevo_registro).execute()
                        
                        # Guardar en tabla separada de altas
                        supabase.table("altas_inventario").insert(log_alta).execute()

                        st.success(f"✅ SKU '{sku.strip()}' dado de alta y registrado en el historial de entradas.")
                    except Exception as err:
                        st.error(f"Error al procesar el alta: {err}")

        # Ver historial de altas recientes
        with st.expander("📜 Ver historial de altas recientes"):
            try:
                res_altas = supabase.table("altas_inventario").select("*").order("created_at", desc=True).limit(20).execute()
                if res_altas.data:
                    df_altas = pd.DataFrame(res_altas.data)
                    cols_altas = ["fecha_alta", "sku", "descripcion", "stock_inicial", "unidad", "responsable"]
                    st.dataframe(df_altas[[c for c in cols_altas if c in df_altas.columns]], use_container_width=True, hide_index=True)
                else:
                    st.info("Sin registros de altas todavía.")
            except Exception as e_alt:
                st.caption(f"No se pudo cargar la vista de altas: {e_alt}")

    # ----------------------------------------------------
    # TAB 2: CATÁLOGO, FICHA TÉCNICA Y EDICIÓN DE STOCK
    # ----------------------------------------------------
    with tab_consulta:
        try:
            res_inv = supabase.table("Inventario_Atlas").select("*").execute()
            data_inv = res_inv.data

            if data_inv:
                df_inv = pd.DataFrame(data_inv)

                # --- 1. SECCIÓN DE FICHA TÉCNICA VISUAL MEJORADA ---
                st.subheader("🔍 Ficha Técnica del Artículo")
                dict_items = {f"{row['SKU']} - {row.get('DESCRIPCION', '')}": row for row in data_inv}
                sel_ver = st.selectbox("Seleccionar artículo a inspeccionar o editar:", list(dict_items.keys()))
                item_sel = dict_items[sel_ver]

                with st.container(border=True):
                    col_izq, col_der = st.columns([1.8, 1.2])

                    with col_izq:
                        st.markdown(f"### {item_sel.get('DESCRIPCION') or 'Sin descripción'}")
                        st.caption(f"**SKU:** `{item_sel.get('SKU')}` | **Categoría:** {item_sel.get('CATEGORIA') or 'N/A'}")

                        # Tarjetas de métricas
                        m1, m2, m3 = st.columns(3)
                        m1.metric("Stock Actual", f"{item_sel.get('STOCK')} {item_sel.get('UNIDAD', '')}")
                        m2.metric("Ubicación", item_sel.get('UBICACION') or "Sin asignar")
                        estado_lbl = "🟢 Activo" if item_sel.get('STATUS') else "🔴 Inactivo"
                        m3.metric("Estatus", estado_lbl)

                        st.divider()
                        st.markdown(f"**Marca:** {item_sel.get('MARCA') or 'N/A'} &nbsp;&nbsp;|&nbsp;&nbsp; **Modelo:** {item_sel.get('MODELO') or 'N/A'} &nbsp;&nbsp;|&nbsp;&nbsp; **Medida:** {item_sel.get('MEDIDA') or 'N/A'}")

                    with col_der:
                        img_url = item_sel.get('REFERENCIA') or item_sel.get('IMAGEN')
                        if img_url and str(img_url).startswith("http"):
                            st.image(img_url, caption=f"Foto SKU: {item_sel.get('SKU')}", use_container_width=True)
                        else:
                            st.info("📷 Este artículo no cuenta con imagen de referencia.")

                # --- 2. MODIFICAR STOCK O DATOS DEL ÍTEM SELECCIONADO ---
                with st.expander(f"✏️ Editar información o surtir stock de {item_sel.get('SKU')}"):
                    with st.form(key=f"form_edit_{item_sel.get('SKU')}"):
                        st.markdown("**Gestión Rápida de Stock**")
                        c_stk1, c_stk2 = st.columns(2)
                        with c_stk1:
                            try:
                                stock_num = int(item_sel.get('STOCK', 0))
                            except Exception:
                                stock_num = 0
                            stock_fijo = st.number_input("Establecer Stock Total:", value=stock_num, min_value=0, step=1)
                        with c_stk2:
                            piezas_agregar = st.number_input("O agregar piezas recibidas (+):", min_value=0, value=0, step=1)

                        st.markdown("**Actualizar Parámetros Generales**")
                        e_c1, e_c2 = st.columns(2)
                        with e_c1:
                            e_ub = st.text_input("Ubicación:", value=str(item_sel.get('UBICACION') or ''))
                            e_med = st.text_input("Medida:", value=str(item_sel.get('MEDIDA') or ''))
                        with e_c2:
                            e_desc = st.text_area("Descripción:", value=str(item_sel.get('DESCRIPCION') or ''), height=68)
                            e_stat = st.checkbox("Disponible / Activo", value=bool(item_sel.get('STATUS', True)))
                            archivo_imagen_edit = st.file_uploader("Actualizar Imagen de Referencia:", type=["jpg", "jpeg", "png"])

                        if st.form_submit_button("Guardar Cambios"):
                            stock_final = stock_fijo + piezas_agregar if piezas_agregar > 0 else stock_fijo
                            sku_val = str(item_sel.get('SKU', '')).strip()

                            update_payload = {
                                "STOCK": str(stock_final),
                                "UBICACION": e_ub.strip(),
                                "MEDIDA": e_med.strip(),
                                "DESCRIPCION": e_desc.strip(),
                                "STATUS": e_stat
                            }

                            try:
                                if archivo_imagen_edit is not None:
                                    timestamp = int(time.time())
                                    ext = archivo_imagen_edit.name.split(".")[-1]
                                    ruta_img = f"ref_{sku_val}_{timestamp}.{ext}"

                                    supabase.storage.from_("planos").upload(
                                        ruta_img, 
                                        archivo_imagen_edit.getvalue(), 
                                        {"content-type": archivo_imagen_edit.type}
                                    )
                                    url_imagen_nueva = supabase.storage.from_("planos").get_public_url(ruta_img)
                                    update_payload["REFERENCIA"] = url_imagen_nueva
                                    update_payload["IMAGEN"] = url_imagen_nueva

                                supabase.table("Inventario_Atlas").update(update_payload).eq("SKU", item_sel.get('SKU')).execute()
                                st.success("✅ Datos e imagen actualizados exitosamente.")
                                st.rerun()

                            except Exception as err_upd:
                                st.error(f"Error al actualizar: {err_upd}")

                # --- 3. TABLA GENERAL DEL CATÁLOGO ---
                st.write("---")
                st.subheader("Catálogo Completo")
                cols_ver = ["SKU", "DESCRIPCION", "MEDIDA", "STOCK", "UNIDAD", "UBICACION", "MARCA", "STATUS"]
                st.dataframe(df_inv[[c for c in cols_ver if c in df_inv.columns]], use_container_width=True, hide_index=True)

            else:
                st.info("No hay registros en la tabla.")

        except Exception as e_cat:
            st.error(f"Error general en catálogo: {e_cat}")



    # ----------------------------------------------------
    # TAB 3: REQUISICIONES ATLAS (ESTILO CATÁLOGO ATLAS)
    # ----------------------------------------------------
    with tab_req:
        st.subheader("📝 Gestión de Requisiciones de Compra")

        # 1. Obtener lista de POs que tienen requisiciones activas o crear una nueva
        res_reqs_all = supabase.table("requisiciones_atlas").select("po_number, estatus").execute()
        datos_reqs = res_reqs_all.data or []
        
        pos_existentes = sorted(list(set(str(r["po_number"]).strip() for r in datos_reqs if r.get("po_number"))))

        col_sel_po, col_nueva_po = st.columns([2, 1.5])
        with col_sel_po:
            opciones_po = ["-- Seleccionar PO existente --"] + pos_existentes + ["➕ Escribir otra PO"]
            po_elegida = st.selectbox("Seleccionar PO de requisición:", opciones_po)

        with col_nueva_po:
            if po_elegida == "➕ Escribir otra PO" or not pos_existentes:
                po_trabajo = st.text_input("Número de nueva PO:", placeholder="Ej. 50054697").strip()
            elif po_elegida != "-- Seleccionar PO existente --":
                po_trabajo = po_elegida
                st.text_input("PO activa:", value=po_trabajo, disabled=True)
            else:
                po_trabajo = ""

        if not po_trabajo:
            st.info("👆 Selecciona una PO de la lista o escribe una nueva para gestionar sus partidas.")
        else:
            # Consultar partidas registradas para esta PO
            res_items_po = supabase.table("requisiciones_atlas").select("*")\
                .eq("po_number", po_trabajo)\
                .order("id").execute()
            
            partidas_actuales = res_items_po.data or []

            # Mostrar estatus general de la requisición de esta PO
            if partidas_actuales:
                estados_po = set(p.get("estatus") for p in partidas_actuales)
                if "Rechazada" in estados_po:
                    motivo_sup = next((p.get("comentarios_supervisor") for p in partidas_actuales if p.get("comentarios_supervisor")), "Sin observaciones.")
                    st.error(f"🔴 **Requisición Rechazada por Supervisión:** {motivo_sup}")
                    st.caption("Modifica las partidas señaladas en el cuadro de edición antes de reenviarla.")
                elif "Enviada" in estados_po:
                    st.info("🔵 **Estado: En Revisión por Supervisión**")
                elif "Aprobada" in estados_po:
                    st.success("🟢 **Estado: Aprobada para Compra**")
                else:
                    st.warning("🟡 **Estado: Borrador en Almacén**")

            st.write("---")

            # ====================================================
            # CUADRO 1: AGREGAR NUEVAS PARTIDAS A LA PO
            # ====================================================
            with st.expander("➕ Cargar nueva partida a la requisición", expanded=(len(partidas_actuales) == 0)):
                with st.form("form_alta_partida_req", clear_on_submit=True):
                    c_n1, c_n2, c_n3 = st.columns([1.2, 2.2, 1.2])
                    with c_n1:
                        n_tipo = st.selectbox("Tipo de ítem:", ["Herramienta", "Material"])
                        n_cant = st.number_input("Cantidad requerida:", min_value=1, value=1, step=1)
                    with c_n2:
                        n_desc = st.text_input("Descripción técnica / Modelo:", placeholder="Ej. Endmill 1/2 4FT / Redondo 1018")
                        n_med = st.text_input("Medida / Especificación:", placeholder="Ej. 1/2 x 1/2 x 3 pulg")
                        n_link = st.text_input("Link de McMaster:", placeholder="Ej. https://www.mcmaster.com/products/")
                    with c_n3:
                        n_und = st.selectbox("Unidad:", ["PZA", "MT", "TRAMO", "KG", "JGO"])
                        n_urg = st.selectbox("Prioridad:", ["Media", "Alta", "Urgente", "Baja"])

                    n_motivo = st.text_input("Notas / Justificación de la compra:")
                    n_solic = st.text_input("Solicitante (Inventarios):", value="Almacén Atlas")

                    btn_guardar_nueva = st.form_submit_button("➕ Agregar partida a la lista")

                    if btn_guardar_nueva:
                        if not n_desc.strip():
                            st.error("La descripción no puede quedar vacía.")
                        else:
                            supabase.table("requisiciones_atlas").insert({
                                "po_number": po_trabajo,
                                "tipo": n_tipo,
                                "descripcion": n_desc.strip(),
                                "medida": n_med.strip(),
                                "cantidad": int(n_cant),
                                "unidad": n_und,
                                "urgencia": n_urg,
                                "motivo_uso": n_motivo.strip(),
                                "solicitante": n_solic.strip(),
                                "link" : n_link,
                                "estatus": "Borrador"
                            }).execute()
                            st.success(f"Partida añadida a la PO {po_trabajo}.")
                            st.rerun()

            # ====================================================
            # CUADRO 2: EDITAR INFORMACIÓN DE PARTIDA EXISTENTE
            # ====================================================
            if partidas_actuales:
                with st.expander("✏️ Editar información o corregir partida", expanded=True):
                    opciones_items_edit = {
                        f"ID {p['id']} - {p.get('descripcion', '')} ({p.get('cantidad', 1)} {p.get('unidad', 'PZA')})": p
                        for p in partidas_actuales
                    }
                    partida_seleccionada_str = st.selectbox(
                        "Seleccionar ítem para inspeccionar o editar:",
                        list(opciones_items_edit.keys())
                    )
                    item_edit = opciones_items_edit[partida_seleccionada_str]

                    st.markdown("**Actualizar Parámetros de la Partida:**")
                    with st.form("form_editar_partida_req", clear_on_submit=False):
                        e_c1, e_c2, e_c3 = st.columns([1.2, 2.2, 1.2])
                        with e_c1:
                            e_tipo = st.selectbox("Tipo:", ["Herramienta", "Material"], index=0 if item_edit.get("tipo") == "Herramienta" else 1)
                            e_cant = st.number_input("Cantidad:", min_value=1, value=int(item_edit.get("cantidad", 1)), step=1)
                        with e_c2:
                            e_desc = st.text_input("Descripción técnica:", value=item_edit.get("descripcion", ""))
                            e_med = st.text_input("Medida / Dimensión:", value=item_edit.get("medida", "") or "")
                            e_link = st.text_input("Nuevo link:", value=item_edit.get("link","") or "")
                        with e_c3:
                            unidades = ["PZA", "MT", "TRAMO", "KG", "JGO"]
                            u_idx = unidades.index(item_edit.get("unidad", "PZA")) if item_edit.get("unidad") in unidades else 0
                            e_und = st.selectbox("Unidad:", unidades, index=u_idx)
                            
                            prioridades = ["Media", "Alta", "Urgente", "Baja"]
                            p_idx = prioridades.index(item_edit.get("urgencia", "Media")) if item_edit.get("urgencia") in prioridades else 0
                            e_urg = st.selectbox("Prioridad:", prioridades, index=p_idx)

                        e_mot = st.text_input("Observaciones / Motivo de uso:", value=item_edit.get("motivo_uso", "") or "")

                        col_btn_ed1, col_btn_ed2 = st.columns([1.5, 2.5])
                        with col_btn_ed1:
                            btn_actualizar = st.form_submit_button("💾 Guardar Cambios")

                        if btn_actualizar:
                            supabase.table("requisiciones_atlas").update({
                                "tipo": e_tipo,
                                "descripcion": e_desc.strip(),
                                "medida": e_med.strip(),
                                "cantidad": int(e_cant),
                                "unidad": e_und,
                                "urgencia": e_urg,
                                "link": e_link,
                                "motivo_uso": e_mot.strip()
                            }).eq("id", item_edit["id"]).execute()
                            st.success(f"Partida ID {item_edit['id']} actualizada correctamente.")
                            st.rerun()

                    # Opción rápida para borrar la partida si no se requiere
                    with st.expander("🗑️ Eliminar esta partida de la requisición"):
                        st.warning(f"¿Deseas quitar la partida '{item_edit.get('descripcion')}'?")
                        if st.button(f"Confirmar eliminación de ID {item_edit['id']}", type="secondary"):
                            supabase.table("requisiciones_atlas").delete().eq("id", item_edit["id"]).execute()
                            st.success("Partida eliminada.")
                            st.rerun()

            # ====================================================
            # TABLA GENERAL Y ACCIÓN DE ENVÍO FINAL
            # ====================================================
            st.write("---")
            st.markdown(f"### 📋 Resumen de Requisición — PO: `{po_trabajo}`")
            if partidas_actuales:
                df_resumen_po = pd.DataFrame(partidas_actuales)
                cols_mostrar = ["id", "tipo", "descripcion", "medida", "cantidad", "unidad", "urgencia", "estatus", "comentarios_supervisor"]
                st.dataframe(df_resumen_po[[c for c in cols_mostrar if c in df_resumen_po.columns]], use_container_width=True, hide_index=True)

                st.write("")
                puede_mandar = any(p.get("estatus") in ["Borrador", "Rechazada"] for p in partidas_actuales)
                if puede_mandar:
                    if st.button("🚀 Enviar Requisición a Validación de Supervisión", type="primary"):
                        supabase.table("requisiciones_atlas").update({
                            "estatus": "Enviada",
                            "comentarios_supervisor": None
                        }).eq("po_number", po_trabajo).in_("estatus", ["Borrador", "Rechazada"]).execute()
                        st.success(f"✅ Requisición de la PO {po_trabajo} enviada a Supervisión.")
                        st.rerun()
                else:
                    st.success("Esta requisición ya fue enviada o aprobada por Supervisión.")
            else:
                st.info("No hay partidas registradas para esta PO. Utiliza el recuadro '➕ Cargar nueva partida' arriba.")
    # ----------------------------------------------------
    # TAB 4: BAJAS & HISTORIAL (RESTAURADO COMPLETO)
    # ----------------------------------------------------
    with tab_bajas:
        st.subheader("🗑️ Registro de Bajas y Salidas de Inventario")

        try:
            res_inv_bajas = supabase.table("Inventario_Atlas").select("*").execute()
            items_bajas = res_inv_bajas.data

            if items_bajas:
                dict_bajas = {f"{it['SKU']} - {it.get('DESCRIPCION', '')} (Stock: {it.get('STOCK')})": it for it in items_bajas}
                sel_baja = st.selectbox("Seleccionar ítem para dar de baja / salida:", list(dict_bajas.keys()))
                item_a_baja = dict_bajas[sel_baja]

                with st.form("form_baja_inventario", clear_on_submit=True):
                    col_b1, col_b2 = st.columns(2)
                    with col_b1:
                        try:
                            stk_max = int(item_a_baja.get("STOCK", 0))
                        except:
                            stk_max = 0
                        cant_baja = st.number_input("Cantidad a dar de baja:", min_value=1, max_value=max(stk_max, 1), value=1, step=1)
                        motivo = st.selectbox("Motivo de la baja:", ["Desgaste / Fin de vida", "Ruptura / Daño", "Uso en producción", "Ajuste de inventario", "Merma"])

                    with col_b2:
                        responsable_baja = st.text_input("Responsable / Solicitante:", placeholder="Ej. Operador / Taller")
                        notas_baja = st.text_area("Notas u observaciones de la salida:")

                    if st.form_submit_button("Confirmar Baja"):
                        nuevo_stock = max(0, stk_max - cant_baja)

                        log_baja = {
                            "sku": item_a_baja.get("SKU"),
                            "descripcion": item_a_baja.get("DESCRIPCION", ""),
                            "cantidad_baja": cant_baja,
                            "motivo": motivo,
                            "responsable": responsable_baja.strip() or rol,
                            "notas": notas_baja.strip(),
                            "stock_restante": str(nuevo_stock)
                        }

                        try:
                            # 1. Descontar en tabla principal
                            supabase.table("Inventario_Atlas").update({"STOCK": str(nuevo_stock)}).eq("SKU", item_a_baja.get("SKU")).execute()

                            # 2. Registrar en tabla de bitácora de bajas
                            supabase.table("bajas_inventario").insert(log_baja).execute()

                            st.success(f"✅ Se dieron de baja {cant_baja} pz(s) de {item_a_baja.get('SKU')}. Stock restante: {nuevo_stock}")
                            st.rerun()
                        except Exception as e_baja:
                            st.error(f"Error al procesar la baja: {e_baja}")

                st.write("---")
                st.subheader("📜 Historial de Salidas / Bajas")
                try:
                    res_hist_bajas = supabase.table("bajas_inventario").select("*").order("created_at", desc=True).limit(50).execute()
                    if res_hist_bajas.data:
                        df_hb = pd.DataFrame(res_hist_bajas.data)
                        cols_hb = ["created_at", "sku", "descripcion", "cantidad_baja", "motivo", "responsable", "stock_restante"]
                        st.dataframe(df_hb[[c for c in cols_hb if c in df_hb.columns]], use_container_width=True, hide_index=True)
                    else:
                        st.info("No hay historial de bajas registrado aún.")
                except Exception as e_h:
                    st.caption(f"No se pudo cargar la bitácora de bajas: {e_h}")
            else:
                st.info("No hay artículos disponibles para dar de baja.")

        except Exception as e_tab_b:
            st.error(f"Error al cargar módulo de bajas: {e_tab_b}")

# ==============================================================================
#                      MÓDULO DE SUPERVISIÓN Y AUDITORÍA
# ==============================================================================
if rol in ["Supervisor", "Desarrollador", "Gerente"]:
    st.header(f"📋 Panel de Supervisión — Operador: {usuario}")

    # Tabs para organizar: Edición de Panorama vs. Historial de Auditoría
    tab_panorama, tab_auditoria = st.tabs(["🔍 Panorama y Edición de POs", "📜 Bitácora de Cambios"])

    # ---------------------------------------------------------
    # TAB 1: PANORAMA Y EDICIÓN DE PARTIDAS
    # ---------------------------------------------------------
    with tab_panorama:
        diccionario_pos = obtener_pos()
        
        if not diccionario_pos:
            st.warning("No hay órdenes de compra registradas en el sistema.")
        else:
            po_sel = st.selectbox("Selecciona la Orden de Compra (PO):", list(diccionario_pos.keys()), key="sup_po_sel")
            
            if po_sel:
                # Consultar las partidas de la PO seleccionada
                res_piezas = supabase.table("piezas").select("*").eq("po_number", po_sel).execute()
                piezas = res_piezas.data
                
                if not piezas:
                    st.info(f"No se encontraron registros para la PO {po_sel}.")
                else:
                    st.subheader(f"Partidas de la PO: {po_sel} (Total: {len(piezas)} registros)")
                    
                    # Desplegar cada partida en un contenedor expandible para editar
                    for item in piezas:
                        id_item = item.get("id")
                        # Compatibilidad con 'Part number' o 'part_num' según el esquema
                        p_num = item.get("Part number") or item.get("part_num") or "-"
                        cant = item.get("cantidad_solicitada") or item.get("cantidad") or 0
                        fecha_req = item.get("fecha") or "-"
                        comentario_actual = item.get("comentario_pz") or ""
                        
                        estados_actuales = item.get("estados_proceso") or item.get("estados") or ["Pendiente"]
                        if not isinstance(estados_actuales, list):
                            estados_actuales = [estados_actuales]
                        
                        titulo_expander = f"Part: {p_num} | Cant: {cant} | Entrega: {fecha_req} | Estado actual: {', '.join(estados_actuales)}"
                        
                        with st.expander(titulo_expander):
                            col1, col2 = st.columns([1, 1])
                            
                            with col1:
                                opciones_estados = [
                                    "Pendiente", 
                                    "Material Solicitado", 
                                    "Material en Taller", 
                                    "En Maquinado", 
                                    "Control de Calidad", 
                                    "Terminado / Listo para Envío", 
                                    "Entregado"
                                ]
                                
                                # Multiselect para los estados de manufactura
                                nuevos_estados = st.multiselect(
                                    "Estados de Proceso:",
                                    opciones_estados,
                                    default=[e for e in estados_actuales if e in opciones_estados] or ["Pendiente"],
                                    key=f"est_{id_item}"
                                )
                                
                            with col2:
                                nuevo_comentario = st.text_area(
                                    "Notas / Observaciones del Supervisor:",
                                    value=comentario_actual,
                                    key=f"com_{id_item}"
                                )
                            
                            # Botón de guardado con disparador de auditoría
                            if st.button("💾 Guardar Cambios en Partida", key=f"btn_save_{id_item}"):
                                cambios_detectados = False
                                
                                # 1. Validar si cambiaron los estados
                                if sorted(nuevos_estados) != sorted(estados_actuales):
                                    supabase.table("piezas").update({
                                        "estados_proceso": nuevos_estados
                                    }).eq("id", id_item).execute()
                                    
                                    registrar_log_cambio(
                                        po_number=po_sel,
                                        campo_editado="estados_proceso",
                                        valor_anterior=str(estados_actuales),
                                        valor_nuevo=str(nuevos_estados),
                                        part_num=p_num,
                                        motivo="Actualización de estatus de maquinado"
                                    )
                                    cambios_detectados = True

                                # 2. Validar si cambió el comentario
                                if nuevo_comentario.strip() != comentario_actual.strip():
                                    supabase.table("piezas").update({
                                        "comentario_pz": nuevo_comentario.strip()
                                    }).eq("id", id_item).execute()
                                    
                                    registrar_log_cambio(
                                        po_number=po_sel,
                                        campo_editado="comentario_pz",
                                        valor_anterior=comentario_actual,
                                        valor_nuevo=nuevo_comentario.strip(),
                                        part_num=p_num,
                                        motivo="Actualización de notas técnicas"
                                    )
                                    cambios_detectados = True
                                
                                if cambios_detectados:
                                    st.success(f"Partida {p_num} actualizada correctamente por {usuario}.")
                                    time.sleep(1)
                                    st.rerun()
                                else:
                                    st.info("No se realizaron modificaciones.")

    # ---------------------------------------------------------
    # TAB 2: BITÁCORA E HISTORIAL COMPLETO DE AUDITORÍA
    # ---------------------------------------------------------
    with tab_auditoria:
        st.subheader("📜 Registro Completo de Modificaciones en el Sistema")
        
        # Filtros rápidos para el supervisor
        col_filtro1, col_filtro2 = st.columns(2)
        with col_filtro1:
            usuario_filtro = st.selectbox(
                "Filtrar por Responsable:", 
                ["Todos", "Jonathan", "Ilagor", "Alessandro", "Desarrollador"]
            )
        with col_filtro2:
            limite_filas = st.slider("Número de registros recientes:", 10, 200, 50)
            
        try:
            query = supabase.table("historial_cambios").select("*").order("fecha_hora", desc=True).limit(limite_filas)
            
            if usuario_filtro != "Todos":
                query = query.eq("usuario", usuario_filtro)
                
            res_logs = query.execute()
            
            if res_logs.data:
                df_logs = pd.DataFrame(res_logs.data)
                
                # Formateo estético de la fecha y hora
                if "fecha_hora" in df_logs.columns:
                    df_logs["fecha_hora"] = pd.to_datetime(df_logs["fecha_hora"]).dt.strftime('%d/%m/%Y %H:%M:%S')
                
                cols_ordenadas = [
                    "fecha_hora", 
                    "usuario", 
                    "rol", 
                    "po_number", 
                    "part_num", 
                    "campo_editado", 
                    "valor_anterior", 
                    "valor_nuevo", 
                    "motivo"
                ]
                
                cols_finales = [c for c in cols_ordenadas if c in df_logs.columns]
                st.dataframe(df_logs[cols_finales], use_container_width=True, hide_index=True)
            else:
                st.info("No hay registros de auditoría que coincidan con la búsqueda.")
        except Exception as e:
            st.error(f"Error al leer la tabla de historial: {e}")

# ==========================================
# VENTANA 5: GERENTE (DASHBOARD GENERAL)
# ==========================================   
elif rol == "Gerente":
    st.header("📊 Tablero de Control Gerencial")

    tab_prod_ger, tab_inv_ger = st.tabs(["🏭 Producción y Órdenes", "📦 Almacén e Inventario"])

    # ----------------------------------------------------
    # PESTAÑA 1: PRODUCCIÓN Y POs
    # ----------------------------------------------------
    with tab_prod_ger:
        st.header("📈 Dashboard de Gerencia")
        dicc_pos = obtener_pos()
        po = st.selectbox("Seleccione la PO a trabajar:", list(dicc_pos.keys()))
        
        # Ahora filtras las piezas de esa PO:
        respuesta = supabase.table("piezas").select("*").eq("po_number", po).execute()
           
        if respuesta.data:
                piezas_data = respuesta.data
                df = pd.DataFrame(piezas_data)
        
                # 1. Preparamos el DataFrame "limpio" (sin columna repetida de PO)
                df['Estado Material'] = df['material_llegado'].apply(lambda x: "✅ Recibido" if x else "❌ Pendiente")
                columnas_ejecutivas = ['Part number', 'Estado Material', 'notas_material', 'porcentaje_avance', 'herramienta_necesaria', 'maquinista_responsable', 'acabado', 'tratamiento']
                df_ejecutivo = df[columnas_ejecutivas].rename(columns={
                    'Part number': 'Pieza', 'notas_material': 'Material', 'porcentaje_avance': 'Avance (%)', 'herramienta_necesaria': 'Herramientas',
                    'maquinista_responsable': 'Maquinista responsable', 'acabado': 'Acabado', 'tratamiento': 'Tratamiento'
                })
        
                # 2. Exportación con Gráfica y Encabezado de PO en Excel
                import io
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    # Inicia la tabla en la fila 4 (índice 3) para dejar espacio al título arriba
                    df_ejecutivo.to_excel(writer, index=False, sheet_name='Reporte_Ejecutivo', startrow=3)
        
                    workbook = writer.book
                    worksheet = writer.sheets['Reporte_Ejecutivo']

                    # Formato del encabezado superior con la PO
                    formato_po_banner = workbook.add_format({
                        'bold': True,
                        'font_size': 13,
                        'fg_color': '#004A99',
                        'font_color': 'white',
                        'align': 'left',
                        'valign': 'vcenter',
                        'border': 1
                    })

                    # Bloque superior para la PO (Celdas A1 a D2)
                    worksheet.merge_range('A1:D2', f"  ORDEN DE PRODUCCIÓN / PO: {po}", formato_po_banner)
        
                    # Formato de cabeceras de columnas de la tabla
                    header_format = workbook.add_format({
                        'bold': True, 
                        'fg_color': '#004A99', 
                        'font_color': 'white', 
                        'border': 1,
                        'align': 'center'
                    })
                    for col_num, value in enumerate(df_ejecutivo.columns.values):
                        worksheet.write(3, col_num, value, header_format)
                    
                    # Ajuste automático seguro del ancho de las celdas
                    for col_num, col_name in enumerate(df_ejecutivo.columns):
                        max_len_datos = max([len(str(val)) for val in df_ejecutivo[col_name].dropna().values], default=0)
                        max_len = max(max_len_datos, len(str(col_name))) + 4
                        worksheet.set_column(col_num, col_num, max(max_len, 12))
        
                    # Insertar Gráfica referenciada a la tabla a partir de la fila 4
                    chart = workbook.add_chart({'type': 'column'})
                    chart.add_series({
                        'name': 'Avance de Producción',
                        'categories': ['Reporte_Ejecutivo', 4, 0, 3 + len(df_ejecutivo), 0],
                        'values':     ['Reporte_Ejecutivo', 4, 3, 3 + len(df_ejecutivo), 3],
                        'fill':       {'color': '#004A99'},
                    })
                    chart.set_x_axis({'label_position': 'low'})
                    
                    worksheet.insert_chart('J2', chart) # La gráfica se acomoda a la derecha
        
                # 3. Botón de descarga
                st.download_button(
                    label="📥 Descargar Reporte Completo (con Gráfica)",
                    data=buffer,
                    file_name=f"Reporte_Gerencia_PO_{po}.xlsx",
                    mime="application/vnd.ms-excel"
                )
        
                # 4. Visualización en la app (Métricas y Tabla)
                col1, col2, col3 = st.columns(3)
                col1.metric("Total de Piezas", len(df))
                col2.metric("Avance Global", f"{int(df['porcentaje_avance'].mean())}%")
                col3.metric("Material Recibido", f"{df['material_llegado'].sum()}/{len(df)}")
        
                st.subheader("Reporte Ejecutivo de Fabricación")
                st.dataframe(df_ejecutivo, use_container_width=True, hide_index=True)

    # ----------------------------------------------------
    # TAB 2: CATÁLOGO, FICHA TÉCNICA Y EDICIÓN DE STOCK
    # ----------------------------------------------------
    with tab_inv_ger:
        try:
            res_inv = supabase.table("Inventario_Atlas").select("*").execute()
            data_inv = res_inv.data

            if data_inv:
                df_inv = pd.DataFrame(data_inv)

                # --- 1. SECCIÓN DE FICHA TÉCNICA VISUAL MEJORADA ---
                st.subheader("🔍 Ficha Técnica del Artículo")
                dict_items = {f"{row['SKU']} - {row.get('DESCRIPCION', '')}": row for row in data_inv}
                sel_ver = st.selectbox("Seleccionar artículo a inspeccionar o editar:", list(dict_items.keys()))
                item_sel = dict_items[sel_ver]

                with st.container(border=True):
                    col_izq, col_der = st.columns([1.8, 1.2])

                    with col_izq:
                        st.markdown(f"### {item_sel.get('DESCRIPCION') or 'Sin descripción'}")
                        st.caption(f"**SKU:** `{item_sel.get('SKU')}` | **Categoría:** {item_sel.get('CATEGORIA') or 'N/A'}")

                        # Tarjetas de métricas
                        m1, m2, m3 = st.columns(3)
                        m1.metric("Stock Actual", f"{item_sel.get('STOCK')} {item_sel.get('UNIDAD', '')}")
                        m2.metric("Ubicación", item_sel.get('UBICACION') or "Sin asignar")
                        estado_lbl = "🟢 Activo" if item_sel.get('STATUS') else "🔴 Inactivo"
                        m3.metric("Estatus", estado_lbl)

                        st.divider()
                        st.markdown(f"**Marca:** {item_sel.get('MARCA') or 'N/A'} &nbsp;&nbsp;|&nbsp;&nbsp; **Modelo:** {item_sel.get('MODELO') or 'N/A'} &nbsp;&nbsp;|&nbsp;&nbsp; **Medida:** {item_sel.get('MEDIDA') or 'N/A'}")

                    with col_der:
                        img_url = item_sel.get('REFERENCIA') or item_sel.get('IMAGEN')
                        if img_url and str(img_url).startswith("http"):
                            st.image(img_url, caption=f"Foto SKU: {item_sel.get('SKU')}", use_container_width=True)
                        else:
                            st.info("📷 Este artículo no cuenta con imagen de referencia.")
                
                # Filtros rápidos para Gerencia
                st.subheader("Catálogo de Almacén")
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    buscar_sku = st.text_input("🔍 Buscar por SKU o Descripción:", placeholder="Ej. Endmill, AT-02, 1/2...")
                with col_f2:
                    categorias_disp = ["Todas"] + sorted([c for c in df_inv["CATEGORIA"].dropna().unique() if str(c).strip()])
                    cat_sel = st.selectbox("Filtrar por Categoría:", categorias_disp)

                # Aplicar filtros
                df_filtrado = df_inv.copy()
                if buscar_sku.strip():
                    term = buscar_sku.strip().lower()
                    df_filtrado = df_filtrado[
                        df_filtrado["SKU"].astype(str).str.lower().str.contains(term) |
                        df_filtrado["DESCRIPCION"].astype(str).str.lower().str.contains(term)
                    ]
                if cat_sel != "Todas":
                    df_filtrado = df_filtrado[df_filtrado["CATEGORIA"] == cat_sel]

                # Visualización interactiva
                cols_mostrar = ["SKU", "DESCRIPCION", "MEDIDA", "STOCK", "UNIDAD", "UBICACION", "MARCA", "REFERENCIA"]
                cols_validas = [col for col in cols_mostrar if col in df_filtrado.columns]

                # st.dataframe configurado con column_config para mostrar fotos directamente en la tabla
                st.dataframe(
                    df_filtrado[cols_validas],
                    column_config={
                        "REFERENCIA": st.column_config.ImageColumn(
                            "Foto Referencia",
                            help="Imagen técnica del artículo"
                        )
                    },
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("No hay inventario registrado en la base de datos.")

        except Exception as e_gi:
            st.error(f"Error cargando dashboard de inventario: {e_gi}")
        
