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


# ==============================================================================
#      FUNCIÓN DE AUDITORÍA EXCLUSIVA PARA SUPERVISORES
# ==============================================================================
def registrar_log_cambio(po_number, campo_editado, valor_anterior, valor_nuevo, part_num="-", motivo=""):
    """Inserta una entrada en la tabla historial_cambios de Supabase."""
    try:
        sup = (
            st.session_state.get("supervisor_activo") 
            or st.session_state.get("usuario_actual") 
            or "Supervisor"
        )
        rol = st.session_state.get("rol_actual", "Supervisor")

        data = {
            "usuario": str(sup),
            "rol": str(rol),
            "po_number": str(po_number),
            "part_num": str(part_num) if part_num else "-",
            "campo_editado": str(campo_editado),
            "valor_anterior": str(valor_anterior) if valor_anterior is not None else "",
            "valor_nuevo": str(valor_nuevo) if valor_nuevo is not None else "",
            "motivo": str(motivo) if motivo else ""
        }
        res = supabase.table("historial_cambios").insert(data).execute()
        return res
    except Exception as e:
        print(f"Error al registrar log de supervisor: {e}")
        return None


# ==============================================================================
#      1. BITÁCORA DE SUPERVISORES (PDF HORIZONTAL CON HORA LOCAL)
# ==============================================================================
def generar_pdf_historial_cambios(registros):
    """Genera un reporte PDF formal de la bitácora de supervisión con hora local."""
    pdf = FPDF(orientation="L", unit="mm", format="Letter")
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.set_margins(left=10, top=10, right=10)
    pdf.add_page()

    # Título principal
    pdf.set_font("Helvetica", "B", 15)
    pdf.set_text_color(0, 74, 153)
    pdf.cell(0, 7, "HISTORIAL Y BITACORA DE MODIFICACIONES - ATLAS", ln=True, align="C")

    # Hora de emisión en tiempo local de México (UTC-6)
    try:
        hora_emision = pd.Timestamp.now(tz="America/Monterrey").strftime("%d/%m/%Y %H:%M:%S")
    except Exception:
        hora_emision = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 5, f"Fecha de emision: {hora_emision}", ln=True, align="C")
    pdf.ln(3)

    # Anchos calculados exactamente para Letter horizontal (Suma = 258 mm)
    w_cols = [34, 25, 22, 26, 35, 58, 58]
    encabezados = ["Fecha / Hora", "Supervisor", "Rol", "PO", "Campo", "Valor Anterior", "Valor Nuevo"]

    def dibujar_encabezado_sup():
        pdf.set_fill_color(0, 74, 153)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 8)
        for i, enc in enumerate(encabezados):
            pdf.cell(w_cols[i], 7, enc, border=1, ln=(1 if i == len(encabezados) - 1 else 0), align="C", fill=True)
        pdf.set_font("Helvetica", "", 7.5)
        pdf.set_text_color(30, 30, 30)

    dibujar_encabezado_sup()
    alto_linea = 4.0

    for reg in registros:
        fecha_raw = str(reg.get("fecha_hora", ""))
        try:
            dt = pd.to_datetime(fecha_raw)
            if dt.tz is None:
                dt = dt.tz_localize("UTC")
            fecha_txt = dt.tz_convert("America/Monterrey").strftime("%d/%m/%Y %H:%M")
        except Exception:
            fecha_txt = fecha_raw[:16]

        valores = [
            fecha_txt,
            str(reg.get("usuario", "-")),
            str(reg.get("rol", "-")),
            str(reg.get("po_number", "-")),
            str(reg.get("campo_editado", "-")),
            str(reg.get("valor_anterior", "-")),
            str(reg.get("valor_nuevo", "-"))
        ]

        # Cálculo dinámico de altura para celdas multilinea
        lineas_celdas = []
        for i, val in enumerate(valores):
            ancho_t = pdf.get_string_width(val)
            num_lineas = max(int(ancho_t / max(w_cols[i] - 3, 1)) + 1, len(val.split("\n")), 1)
            lineas_celdas.append(num_lineas)

        altura_fila = max(max(lineas_celdas) * alto_linea + 3, 7.0)

        if pdf.get_y() + altura_fila > pdf.page_break_trigger:
            pdf.add_page()
            dibujar_encabezado_sup()

        pos_x_inicio = pdf.get_x()
        pos_y_inicio = pdf.get_y()

        for i, val in enumerate(valores):
            pdf.rect(pos_x_inicio, pos_y_inicio, w_cols[i], altura_fila)
            pdf.set_xy(pos_x_inicio + 1, pos_y_inicio + 1.2)
            alineacion = "C" if i in [0, 1, 2, 3] else "L"
            pdf.multi_cell(w_cols[i] - 2, alto_linea, val, border=0, align=alineacion)
            pos_x_inicio += w_cols[i]

        pdf.set_xy(pdf.l_margin, pos_y_inicio + altura_fila)

    salida = pdf.output(dest="S")
    if isinstance(salida, str):
        return salida.encode("latin-1", errors="replace")
    return bytes(salida)





# ==============================================================================
#      FUNCIÓN DE AUDITORÍA EXCLUSIVA PARA MAQUINISTAS
# ==============================================================================
def registrar_log_maquinista(po_number, part_num, tipo_accion, agregadas="", removidas="", detalle="", notas=""):
    """Inserta una entrada en la tabla historial_maquinistas de Supabase."""
    try:
        maq = st.session_state.get("maquinista_activo") or st.session_state.get("usuario_actual") or "Maquinista"
        data = {
            "maquinista": str(maq),
            "po_number": str(po_number),
            "part_num": str(part_num) if part_num else "-",
            "tipo_accion": str(tipo_accion),
            "herramientas_agregadas": str(agregadas) if agregadas else "-",
            "herramientas_removidas": str(removidas) if removidas else "-",
            "detalle_general": str(detalle) if detalle else "-",
            "notas": str(notas) if notas else "-"
        }
        supabase.table("historial_maquinistas").insert(data).execute()
    except Exception as e:
        print(f"Error al registrar log de maquinista: {e}")


# ==============================================================================
#      2. BITÁCORA DE MAQUINISTAS (PDF HORIZONTAL CON HORA LOCAL)
# ==============================================================================
def generar_pdf_historial_maquinistas(registros):
    """Genera el reporte PDF horizontal con celdas de altura calculada de forma exacta."""
    pdf = FPDF(orientation="L", unit="mm", format="Letter")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_margins(left=10, top=12, right=10)
    pdf.add_page()

    # Título principal
    pdf.set_font("Helvetica", "B", 15)
    pdf.set_text_color(0, 74, 153)
    pdf.cell(0, 7, "BITACORA DE TRABAJO Y HERRAMENTAL - MAQUINISTAS", ln=True, align="C")

    # Hora actual convertida a hora local
    try:
        hora_emision = pd.Timestamp.now(tz="America/Monterrey").strftime("%d/%m/%Y %H:%M:%S")
    except Exception:
        hora_emision = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    pdf.set_font("Helvetica", "", 8.5)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 5, f"Fecha de emision: {hora_emision}", ln=True, align="C")
    pdf.ln(4)

    # Anchos calibrados (Suma exacta = 259 mm)
    # [Fecha, Maq, PO, Parte, Acción, Agregadas, Removidas, Detalles]
    w_cols = [28, 22, 18, 25, 26, 46, 46, 48]
    encabezados = [
        "Fecha / Hora", "Maquinista", "PO", "Parte", 
        "Acción", "Herr. Agregadas", "Herr. Removidas", "Detalles / Notas"
    ]

    def dibujar_encabezado_maq():
        pdf.set_fill_color(0, 74, 153)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 8)
        for i, enc in enumerate(encabezados):
            pdf.cell(w_cols[i], 7.5, enc, border=1, ln=(1 if i == len(encabezados) - 1 else 0), align="C", fill=True)
        pdf.set_font("Helvetica", "", 7.5)
        pdf.set_text_color(30, 30, 30)

    dibujar_encabezado_maq()
    alto_renglon = 4.0  # Altura por cada línea de texto

    for reg in registros:
        fecha_raw = str(reg.get("fecha_hora", ""))
        try:
            dt = pd.to_datetime(fecha_raw)
            if dt.tz is None:
                dt = dt.tz_localize("UTC")
            fecha_txt = dt.tz_convert("America/Monterrey").strftime("%d/%m/%Y %H:%M")
        except Exception:
            fecha_txt = fecha_raw[:16]

        valores = [
            fecha_txt,
            str(reg.get("maquinista", "-")),
            str(reg.get("po_number", "-")),
            str(reg.get("part_num", "-")),
            str(reg.get("tipo_accion", "-")),
            str(reg.get("herramientas_agregadas", "-")),
            str(reg.get("herramientas_removidas", "-")),
            f"{reg.get('detalle_general', '')} | {reg.get('notas', '')}".strip(" |")
        ]

        # Calcular renglones reales mediante simulación palabra por palabra
        pdf.set_font("Helvetica", "", 7.5)
        conteo_lineas = []
        for i, texto in enumerate(valores):
            ancho_util = w_cols[i] - 3.5  # Margen interior de celda
            renglones = 0
            for parrafo in str(texto).split("\n"):
                palabras = parrafo.split(" ")
                linea_act = ""
                for p in palabras:
                    prueba = f"{linea_act} {p}".strip()
                    if pdf.get_string_width(prueba) <= ancho_util:
                        linea_act = prueba
                    else:
                        renglones += 1
                        linea_act = p
                if linea_act:
                    renglones += 1
            conteo_lineas.append(max(renglones, 1))

        # Altura de la fila con padding superior e inferior de 3 mm
        altura_fila = max(conteo_lineas) * alto_renglon + 5.0

        # Salto de página automático si no cabe la fila completa
        if pdf.get_y() + altura_fila > pdf.page_break_trigger:
            pdf.add_page()
            dibujar_encabezado_maq()

        pos_x_inicio = pdf.get_x()
        pos_y_inicio = pdf.get_y()

        # Dibujar rectángulos y colocar textos con margen interno limpio
        for i, val in enumerate(valores):
            pdf.rect(pos_x_inicio, pos_y_inicio, w_cols[i], altura_fila)
            pdf.set_xy(pos_x_inicio + 1.8, pos_y_inicio + 2.5)
            alineacion = "C" if i in [0, 1, 2, 3, 4] else "L"
            pdf.multi_cell(w_cols[i] - 3.6, alto_renglon, val, border=0, align=alineacion)
            pos_x_inicio += w_cols[i]

        pdf.set_xy(pdf.l_margin, pos_y_inicio + altura_fila)

    salida = pdf.output(dest="S")
    if isinstance(salida, str):
        return salida.encode("latin-1", errors="replace")
    return bytes(salida)

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

# --- SISTEMA DE LOGIN ---
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.rol_actual = None

if not st.session_state.autenticado:
    st.sidebar.header("🔐 Iniciar Sesión")
    rol_elegido = st.sidebar.selectbox("Selecciona tu rol:", ["Desarrollador", "Maquinista", "Inventarios", "Supervisor", "Gerente"])
    password = st.sidebar.text_input("Contraseña:", type="password")

    passwords = {
        "Desarrollador": "120501",
        "Maquinista": "Maq",
        "Inventarios": "Inv2026",
        "Supervisor": "S2S3",
        "Gerente": "GJE2026"
    }

    if st.sidebar.button("Entrar"):
        if password == passwords.get(rol_elegido):
            st.session_state.autenticado = True
            st.session_state.rol_actual = rol_elegido
            st.rerun()
        else:
            st.sidebar.error("Contraseña incorrecta.")
    st.stop()
else:
    rol = st.session_state.rol_actual
    if st.sidebar.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.session_state.rol_actual = None
        st.session_state.usuario_actual = None
        st.session_state.supervisor_activo = None  # <-- Limpia el supervisor activo
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
                        # Inicialización garantizada para prevenir NameError
                        url_org = None
                        url_tl = None
                        texto_traducido_tl = ""
                        texto = ""
                        timestamp = int(time.time())

                        try:
                            # 1. Procesar archivo original
                            if 'org' in archivos:
                                archivo_org = archivos['org']
                                ruta = f"{nombre_pieza}_ORG_{timestamp}.pdf"
                                supabase.storage.from_("planos").upload(
                                    ruta, 
                                    archivo_org.getvalue(), 
                                    {"content-type": "application/pdf"}
                                )
                                url_org = supabase.storage.from_("planos").get_public_url(ruta)
        
                            # 2. Procesar archivo TL y traducción (si existe)
                            if 'tl' in archivos:
                                archivo_tl = archivos['tl']
                                
                                try:
                                    import pypdf
                                    import io
                                    reader = pypdf.PdfReader(io.BytesIO(archivo_tl.getvalue()))
                                    texto = "".join([p.extract_text() for p in reader.pages if p.extract_text()])
                                except Exception as e_pdf:
                                    st.warning(f"No se pudo extraer texto del PDF TL ({nombre_pieza}): {e_pdf}")
                                    texto = ""
        
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
                                            st.warning(f"⚠️ No se pudo traducir {nombre_pieza}: {e_final}")
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
                            registro_pieza = {
                                "po_number": po.strip(),
                                "Part number": nombre_pieza,
                                "url_pdf_org": url_org,
                                "url_pdf_tl": url_tl,
                                "texto_traducido": texto_traducido_tl,
                                "cantidad_solicitada": 1,
                                "piezas_terminadas": 0,
                                "porcentaje_avance": 0,
                                "es_ensamble": False
                            }

                            # Buscar si la pieza ya existe para esta PO
                            res_existente = supabase.table("piezas").select("id").eq("po_number", po.strip()).eq("Part number", nombre_pieza).execute()

                            if res_existente.data:
                                id_p = res_existente.data[0]["id"]
                                datos_actualizacion = {}
                                if url_org:
                                    datos_actualizacion["url_pdf_org"] = url_org
                                if url_tl:
                                    datos_actualizacion["url_pdf_tl"] = url_tl
                                if texto_traducido_tl:
                                    datos_actualizacion["texto_traducido"] = texto_traducido_tl
                                
                                if datos_actualizacion:
                                    supabase.table("piezas").update(datos_actualizacion).eq("id", id_p).execute()
                            else:
                                supabase.table("piezas").insert(registro_pieza).execute()

                            st.success(f"✅ Pieza '{nombre_pieza}' procesada y guardada correctamente en PO {po.strip()}.")

                        except Exception as e:
                            st.error(f"Error procesando {nombre_pieza}: {e}")

    # ------------------------------------------------------------------
    # 2. PIEZAS RECURRENTES / ÓRDENES DE PRODUCCIÓN (HOJAS AMETEK)
    # ------------------------------------------------------------------
    with tab_hoja_recurrentes:
        st.subheader("Cargar partida de hoja PO usando piezas anteriores")
        st.caption("Usa esta pestaña para órdenes como la PO 50054832 donde las piezas ya se habían maquinado.")

        try:
            res_historico = supabase.table("piezas").select("*").execute()
            todas_piezas = res_historico.data

            if todas_piezas:
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
                                "porcentaje_avance": 0,
                                "es_ensamble": bool(molde.get("es_ensamble", False))
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
# VENTANA 2: MAQUINISTA (SOLO OPERACIÓN EN TALLER)
# ==========================================
elif rol == "Maquinista":
    st.header("🛠️ Área de Maquinado")

    maquinistas = {
        "Abel": "8352",
        "Alejandro": "1497",
        "Francisco": "5206",
        "Isaac": "3984",
        "Jesus": "5731",
        "Rafael": "6475",
        "Edwin": "5759",
        "Alessandro": "1234"
    }

    # Control de acceso estricto
    if not st.session_state.get("maquinista_autenticado", False):
        st.subheader("🔐 Identificación de Maquinista")
        with st.form(key="form_auth_maquinista"):
            col_l1, col_l2 = st.columns(2)
            opciones_maq = ["-- Seleccionar operador --"] + list(maquinistas.keys())
            usuario_sel = col_l1.selectbox("Nombre:", opciones_maq)
            password_ingresado = col_l2.text_input("Contraseña / PIN:", type="password")

            if st.form_submit_button("Ingresar a Taller", use_container_width=True):
                if usuario_sel == "-- Seleccionar operador --":
                    st.error("⚠️ Debes seleccionar tu nombre.")
                elif password_ingresado == maquinistas.get(usuario_sel):
                    st.session_state.maquinista_activo = usuario_sel
                    st.session_state.usuario_actual = usuario_sel
                    st.session_state.maquinista_autenticado = True
                    st.rerun()
                else:
                    st.error("❌ Contraseña incorrecta.")
        st.stop()

    usuario = st.session_state.maquinista_activo
    c_user1, c_user2 = st.columns([4, 1])
    c_user1.success(f"Operador activo: **{usuario}**")
    if c_user2.button("Cambiar Operador"):
        st.session_state.maquinista_activo = None
        st.session_state.maquinista_autenticado = False
        st.session_state.usuario_actual = None
        st.rerun()

    dicc_pos = obtener_pos()
    if not dicc_pos:
        st.warning("No hay POs registradas.")
        st.stop()

    po = st.selectbox("Seleccione la PO a trabajar:", list(dicc_pos.keys()))
    
    respuesta = supabase.table("piezas").select("*").eq("po_number", po).execute()
    piezas = respuesta.data

    if not piezas:
        st.info("No hay piezas cargadas para esta PO.")
        st.stop()
    
    nombres_piezas = {
        (p.get("Part number") or p.get("part_num")): p 
        for p in piezas 
        if (p.get("Part number") or p.get("part_num"))
    }
    
    pieza_seleccionada = st.selectbox("Selecciona la pieza a trabajar:", list(nombres_piezas.keys()))
    datos_pieza = nombres_piezas[pieza_seleccionada]
    col_part = "Part number" if "Part number" in datos_pieza else "part_num"
    
    # 1. REGISTRO DE PIEZAS TERMINADAS
    with st.expander("📝 Registrar piezas terminadas"):
        with st.form(key="form_registro"):
            piezas_realizadas = st.multiselect("¿Qué piezas terminaste?", list(nombres_piezas.keys()))
            notas_adicionales = st.text_area("Notas:")
            
            if st.form_submit_button("Registrar"):
                notas_finales = str(notas_adicionales).strip() if notas_adicionales else "Sin notas"
                
                for p_num in piezas_realizadas:
                    supabase.table("piezas").update({
                        "maquinista_responsable": usuario,
                        "notas_registro": notas_finales
                    }).eq(col_part, p_num).execute()
                    
                    registrar_log_maquinista(
                        po_number=po,
                        part_num=p_num,
                        tipo_accion="Pieza Terminada",
                        detalle=f"Terminada por {usuario}",
                        notas=notas_finales
                    )
                
                st.success("¡Registro actualizado en la base de datos y archivado en bitácora!")
                time.sleep(1)
                st.rerun()

    # 2. PLANOS Y ESPECIFICACIONES
    tab1, tab2, tab3 = st.tabs(["📄 Plano Original", "⚙️ Especificaciones (TL)", "📝 Traducción"])

    with tab1:
        url_org = datos_pieza.get('url_pdf_org')
        if url_org and str(url_org).startswith("http"):
            st.link_button("↗️ Abrir plano en pestaña nueva", url_org)
            pdf_org_bytes = descargar_bytes_pdf(url_org)
            if pdf_org_bytes:
                pdf_viewer(input=pdf_org_bytes, width=950, height=650)
            else:
                st.warning("No se pudo cargar el archivo del plano.")
        else:
            st.warning("No hay plano original disponible.")

    with tab2:
        url_tl = datos_pieza.get('url_pdf_tl')
        if url_tl and str(url_tl).startswith("http"):
            st.link_button("↗️ Abrir archivo TL en pestaña nueva", url_tl)
            pdf_tl_bytes = descargar_bytes_pdf(url_tl)
            if pdf_tl_bytes:
                pdf_viewer(input=pdf_tl_bytes, width=950, height=650)
            else:
                st.warning("No se pudo cargar el archivo TL.")
        else: 
            st.warning("No hay archivo TL disponible.")

    with tab3:
        st.info(datos_pieza.get('texto_traducido') or "Sin traducción disponible.")

    # 3. ACTUALIZAR PARÁMETROS TÉCNICOS
    st.subheader("Actualizar Parámetros")
    herramientas_bd = sorted(list(set(
        h.strip() 
        for p in piezas if p.get('herramienta_necesaria') 
        for h in str(p['herramienta_necesaria']).split('\n') if h.strip()
    )))
    acabados_bd = sorted(list(set(str(p['acabado']).strip() for p in piezas if p.get('acabado') and str(p['acabado']).strip())))
    tratamientos_bd = sorted(list(set(str(p['tratamiento']).strip() for p in piezas if p.get('tratamiento') and str(p['tratamiento']).strip())))

    lista_herr_previas = [h.strip() for h in str(datos_pieza.get('herramienta_necesaria', '')).split('\n') if h.strip()]

    with st.form(key=f"form_{datos_pieza[col_part]}"):
        c1, c2, c3 = st.columns(3)
        with c1:
            herr_sel = st.multiselect(
                "Herramientas existentes:", 
                herramientas_bd, 
                default=[h for h in lista_herr_previas if h in herramientas_bd]
            )
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
            nuevas_ingresadas = [h.strip() for h in herr_nuevas.split('\n') if h.strip()]
            herr_final_lista = list(dict.fromkeys(herr_sel + nuevas_ingresadas))
            herr_final_txt = '\n'.join(herr_final_lista)

            acab_final = acabado_nuevo.strip() if acabado_nuevo.strip() != "" else (acabado_sel if acabado_sel != "Ninguno" else "")
            trat_final = trat_nuevo.strip() if trat_nuevo.strip() != "" else (trat_sel if trat_sel != "Ninguno" else "")

            # Detección de herramientas agregadas y removidas
            set_previo = set(lista_herr_previas)
            set_nuevo = set(herr_final_lista)
            herramientas_anadidas = list(set_nuevo - set_previo)
            herramientas_quitadas = list(set_previo - set_nuevo)

            update_data = {
                "herramienta_necesaria": herr_final_txt,
                "acabado": acab_final,
                "tratamiento": trat_final,
                "material_llegado": mat_recibido,
                "notas_material": comentarios_mat.strip()
            }
            supabase.table("piezas").update(update_data).eq(col_part, datos_pieza[col_part]).execute()

            detalles_cambios = []
            if mat_recibido != bool(datos_pieza.get('material_llegado')):
                detalles_cambios.append(f"Material llegó: {mat_recibido}")
            if acab_final != (datos_pieza.get('acabado') or ""):
                detalles_cambios.append(f"Acabado: {acab_final}")
            if trat_final != (datos_pieza.get('tratamiento') or ""):
                detalles_cambios.append(f"Tratamiento: {trat_final}")

            registrar_log_maquinista(
                po_number=po,
                part_num=datos_pieza[col_part],
                tipo_accion="Actualización de Parámetros",
                agregadas=", ".join(herramientas_anadidas) if herramientas_anadidas else "-",
                removidas=", ".join(herramientas_quitadas) if herramientas_quitadas else "-",
                detalle=" | ".join(detalles_cambios) if detalles_cambios else "Parámetros actualizados",
                notas=comentarios_mat.strip()
            )

            st.success("¡Parámetros actualizados y auditados con éxito!")
            time.sleep(1)
            st.rerun()

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

# ==========================================
# VENTANA 4: SUPERVISOR (MODO EDICIÓN CONTINUA)
# ==========================================
elif rol == "Supervisor":

    st.header("📋 Área de Supervisión: Control Interno")

    # ==========================================
    # 0. CONTROL DE ACCESO POR USUARIO SUPERVISOR
    # ==========================================
    supervisores = {
        "Jonathan": "J6182",
        "Ilagor": "I3954",
        "Sayuri": "S7205",
        "Alessandro": "A1498"
    }

    if "supervisor_activo" not in st.session_state:
        st.session_state.supervisor_activo = None

    if not st.session_state.supervisor_activo:
        st.subheader("🔐 Identificación de Supervisor")
        col_s1, col_s2 = st.columns(2)
        sup_nombre = col_s1.selectbox("Selecciona tu nombre:", list(supervisores.keys()))
        sup_pass = col_s2.text_input("Contraseña / PIN de Supervisor:", type="password")

        if col_s2.button("Ingresar al Panel de Supervisión"):
            if sup_pass == supervisores.get(sup_nombre):
                st.session_state.supervisor_activo = sup_nombre
                st.session_state.usuario_actual = sup_nombre
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")
        st.stop()

    usuario_sup = st.session_state.supervisor_activo
    c_sup_h1, c_sup_h2 = st.columns([4, 1])
    c_sup_h1.success(f"Supervisor en turno: **{usuario_sup}**")
    if c_sup_h2.button("Cambiar Supervisor"):
        st.session_state.supervisor_activo = None
        st.session_state.usuario_actual = None
        st.rerun()

    # Pestañas incluyendo la descarga directa del historial
    tab_piez_term, tab_req, tab_historial_pdf = st.tabs(["🏭 Producción", "📦 Requisiciones", "📥 Descargar Historial (PDF)"])

    # ==========================================
    # CONFIGURACIÓN DE ESTADOS Y COLORES
    # ==========================================
    LISTA_ESTADOS_PIEZAS = [
        "En proceso",
        "Detenida por el torno/fresadora",
        "Pendiente de tratamiento",
        "Con proveedor externo",
        "Material pendiente",
        "Herramienta pendiente",
        "Terminada",
        "Material en almacen",
        "Fabricada en China",
        "Comprado",
        "En fila de calidad",
        "Cancelada",
        "Ya llego de china",
        "Enviado a cliente"
    ]
    
    ESTADOS_COLORES = {
        "Produccion": {"bg": "#FFFFFF", "fg": "#C0504D", "border": "#C0504D"},
        "En proceso": {"bg": "#D9D9D9", "fg": "#000000", "border": "#BFBFBF"},
        "Detenida por el torno/fresadora": {"bg": "#B8CCE4", "fg": "#000000", "border": "#95B3D7"},
        "Pendiente de tratamiento": {"bg": "#FCD5B4", "fg": "#000000", "border": "#FABF8F"},
        "Con proveedor externo": {"bg": "#E5B8E8", "fg": "#000000", "border": "#D899DD"},
        "Material pendiente": {"bg": "#FFFF00", "fg": "#000000", "border": "#E6E600"},
        "Herramienta pendiente": {"bg": "#548235", "fg": "#FFFFFF", "border": "#375623"},
        "Terminada": {"bg": "#000000", "fg": "#FFFFFF", "border": "#000000"},
        "Material en almacen": {"bg": "#00FF00", "fg": "#000000", "border": "#00D900"},
        "Fabricada en China": {"bg": "#33CCCC", "fg": "#000000", "border": "#2BA8A8"},
        "Comprado": {"bg": "#FF6600", "fg": "#FFFFFF", "border": "#CC5200"},
        "En fila de calidad": {"bg": "#FFC000", "fg": "#000000", "border": "#D5A000"},
        "Cancelada": {"bg": "#FFFFFF", "fg": "#C00000", "border": "#C00000"},
        "Ya llego de china": {"bg": "#0099FF", "fg": "#FFFFFF", "border": "#007ACC"},
        "Enviado a cliente": {"bg": "#4C2365", "fg": "#FFFFFF", "border": "#172554"}
    }
    
    CATEGORIAS_FILTRO = ["PO en Producción"] + LISTA_ESTADOS_PIEZAS

    # ==========================================
    # PESTAÑA 1: PRODUCCIÓN (PANORAMA GENERAL + EDICIÓN)
    # ==========================================
    with tab_piez_term:
        dicc_pos = obtener_pos()
        lista_pos = list(dicc_pos.keys())
    
        if not lista_pos:
            st.warning("No hay POs registradas.")
        else:
            # Sincronización interna con session_state
            if "po_activa" not in st.session_state:
                st.session_state.po_activa = lista_pos[0]
            if st.session_state.po_activa not in lista_pos:
                st.session_state.po_activa = lista_pos[0]
    
            # ==========================================
            # SECCIÓN 1 (ARRIBA): PANORAMA GENERAL (MATRIZ DE TALLER)
            # ==========================================
            st.subheader("📊 Panorama General de POs (Matriz de Taller)")
    
            with st.spinner("Generando tablero visual..."):
                res_todas = supabase.table("piezas").select(
                'po_number, "Part number", cantidad_solicitada, fecha_entrega, estados_proceso, estado_proceso, comentarios, comentarios_po, po_en_produccion, es_ensamble'
                ).order("fecha_entrega").execute().data or []
                    
            # 1. Detectar meses disponibles dinámicamente
            nombres_meses = {
                1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
                7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
            }
            meses_set = set()
            for r in res_todas:
                f_str = r.get("fecha_entrega")
                if f_str:
                    try:
                        f_obj = pd.to_datetime(f_str)
                        clave_mes = f_obj.strftime("%Y-%m")
                        etiqueta_mes = f"{nombres_meses.get(f_obj.month, '')} {f_obj.year}"
                        meses_set.add((clave_mes, etiqueta_mes))
                    except Exception:
                        pass
    
            lista_meses_ordenados = sorted(list(meses_set), key=lambda x: x[0])
            opciones_meses = ["Todos los meses"] + [m[1] for m in lista_meses_ordenados]
            mapa_etiqueta_clave = {m[1]: m[0] for m in lista_meses_ordenados}
    
            # --- FILTROS DE VISUALIZACIÓN RÁPIDA (FILA 1) ---
            c_filtro1, c_filtro_env, c_filtro2, c_filtro3, c_filtro4 = st.columns([1.6, 1.8, 2.0, 2.6, 2.0])
            with c_filtro1:
                filtro_solo_prod = st.checkbox("🏭 Solo Producción", value=False)
            with c_filtro_env:
                filtro_ocultar_enviadas = st.checkbox("📦 Ocultar enviadas", value=True, help="Oculta las piezas y POs que ya fueron enviadas al cliente")
            with c_filtro2:
                mes_seleccionado = st.selectbox(
                    "📅 Mes entrega:",
                    options=opciones_meses,
                    index=0,
                    help="Filtra las piezas según su mes de entrega"
                )
            with c_filtro3:
                filtro_terminacion = st.radio(
                    "Filtrar por estatus:",
                    options=["Todas las piezas", "Solo pendientes (Ocultar terminadas)", "Solo terminadas"],
                    horizontal=True,
                    label_visibility="collapsed"
                )
            with c_filtro4:
                busqueda_pieza_filtro = st.text_input("Filtrar por N/P:", placeholder="Ej. 1828...").strip().lower()
    
            # --- FILTRO POR CATEGORÍAS / ESTADOS (FILA 2) ---
            filtro_categorias_sel = st.multiselect(
                "🏷️ Filtrar por categorías / estados de pieza:",
                options=CATEGORIAS_FILTRO,
                default=[],
                placeholder="Selecciona una o más categorías (deja vacío para mostrar todas)..."
            )

            # --- FILTRO POR NÚMERO DE PO (FILA 3) ---
            todas_pos_registradas = sorted(list({str(r.get("po_number", "")).strip() for r in res_todas if r.get("po_number")}))
            opciones_filtro_po = ["-- Todas las POs --"] + todas_pos_registradas
    
            col_po_filtro, col_po_btn = st.columns([3.5, 1.5])
            with col_po_filtro:
                po_filtrada_sel = st.selectbox(
                    "🔍 Filtrar panorama por número de PO:",
                    options=opciones_filtro_po,
                    key="filtro_po_matriz",
                    help="Selecciona o escribe el número de PO para ver únicamente su columna en el tablero"
                )
            with col_po_btn:
                st.write("")
                st.write("")
                if po_filtrada_sel != "-- Todas las POs --":
                    if st.button("Ir a editar esta PO ↗️", use_container_width=True):
                        st.session_state.po_activa = po_filtrada_sel
                        st.rerun()
    
            if res_todas:
                pos_matriz = {}
                clave_mes_filtro = mapa_etiqueta_clave.get(mes_seleccionado, None)
    
                for reg in res_todas:
                    po_k = str(reg.get("po_number", "Sin PO")).strip()
                    pn_r = str(reg.get("Part number", "-")).strip()
                    f_ent = str(reg.get("fecha_entrega") or "Sin fecha").strip()
                    cant_pz_num = reg.get("cantidad_solicitada") or reg.get("cantidad") or 1
    
                    if clave_mes_filtro:
                        try:
                            f_obj = pd.to_datetime(reg.get("fecha_entrega"))
                            if f_obj.strftime("%Y-%m") != clave_mes_filtro:
                                continue
                        except Exception:
                            continue
    
                    if po_k not in pos_matriz:
                        pos_matriz[po_k] = {
                            "comentario": reg.get("comentarios_po") or "",
                            "en_produccion": bool(reg.get("po_en_produccion", False)),
                            "piezas": []
                        }
                    if reg.get("po_en_produccion"):
                        pos_matriz[po_k]["en_produccion"] = True
    
                    raw_ests = reg.get("estados_proceso")
                    if isinstance(raw_ests, list) and len(raw_ests) > 0:
                        lista_est_pz = [e for e in raw_ests if e != "Produccion"]
                    else:
                        st_ind = reg.get("estado_proceso") or "En proceso"
                        lista_est_pz = [st_ind] if st_ind != "Produccion" else ["En proceso"]
    
                    if not lista_est_pz:
                        lista_est_pz = ["En proceso"]
    
                    es_terminada = ("Terminada" in lista_est_pz)
                    es_enviada = ("Enviado a cliente" in lista_est_pz)
                    es_concluida = es_terminada or es_enviada

                    if filtro_ocultar_enviadas and es_enviada:
                        continue

                    if filtro_terminacion == "Solo pendientes (Ocultar terminadas)" and es_concluida:
                        continue
                    if filtro_terminacion == "Solo terminadas" and not es_terminada:
                        continue

                    if busqueda_pieza_filtro and (busqueda_pieza_filtro not in pn_r.lower()):
                        continue
    
                    if filtro_categorias_sel:
                        filtro_estados_sin_prod = [c for c in filtro_categorias_sel if c != "PO en Producción"]
                        coincide_estado_pieza = any(c in lista_est_pz for c in filtro_estados_sin_prod)
                        coincide_po_prod = ("PO en Producción" in filtro_categorias_sel and pos_matriz[po_k]["en_produccion"])

                        if filtro_estados_sin_prod and not coincide_estado_pieza:
                            continue
                        if not filtro_estados_sin_prod and not coincide_po_prod:
                            continue

                    pos_matriz[po_k]["piezas"].append({
                        "part_num": pn_r,
                        "fecha": f_ent,
                        "cantidad": cant_pz_num,
                        "estados": lista_est_pz,
                        "comentario_pz": str(reg.get("comentarios") or "").strip(),
                        "es_ensamble": bool(reg.get("es_ensamble", False))  # <-- Nuevo campo
                    })
    
                lista_pos_cols = [
                    k for k in sorted(list(pos_matriz.keys()))
                    if (not filtro_solo_prod or pos_matriz[k]["en_produccion"])
                    and (po_filtrada_sel == "-- Todas las POs --" or k == po_filtrada_sel)
                    and len(pos_matriz[k]["piezas"]) > 0
                ]
    
                cfg_prod = ESTADOS_COLORES.get("Produccion", {"bg": "#FFFFFF", "fg": "#C0504D", "border": "#C0504D"})
                filas_acciones = [
                    f'<div style="background-color:{cfg_prod["bg"]}; color:{cfg_prod["fg"]}; border:2px dashed {cfg_prod["border"]}; padding:5px 6px; margin-bottom:6px; border-radius:4px; font-size:11px; font-weight:bold; text-align:center;">🏷️ PO en Producción</div>'
                ]
    
                for accion, cfg in ESTADOS_COLORES.items():
                    if accion == "Produccion":
                        continue
                    estilo_acc = f"background-color:{cfg['bg']}; color:{cfg['fg']}; border:1px solid {cfg['border']}; padding:4px 6px; margin-bottom:4px; border-radius:4px; font-size:11px; font-weight:600; text-align:center; white-space:nowrap;"
                    filas_acciones.append(f'<div style="{estilo_acc}">{accion}</div>')
                html_acciones = "".join(filas_acciones)
    
                columnas_pos_html = []
                for po_k in lista_pos_cols:
                    coment_po_k = pos_matriz[po_k]["comentario"]
                    esta_en_prod = pos_matriz[po_k]["en_produccion"]
                    piezas_col = pos_matriz[po_k]["piezas"]

                    bloques_piezas = []
                    for pz in piezas_col:
                        ests_pz = pz["estados"]
                        num_parte = pz["part_num"]
                        fecha_pz = pz["fecha"]
                        cant_pz = pz.get("cantidad", 1)

                        ests_relevantes = [e for e in ests_pz if e != "En proceso"] if len(ests_pz) > 1 else ests_pz
                        if not ests_relevantes:
                            ests_relevantes = ["En proceso"]

                        if len(ests_relevantes) == 1:
                            cfg_c = ESTADOS_COLORES.get(ests_relevantes[0], {"bg": "#D9D9D9", "fg": "#000000", "border": "#BFBFBF"})
                            bg_style = f"background-color: {cfg_c['bg']}; color: {cfg_c['fg']}; border: 1.5px solid {cfg_c['border']};"
                        else:
                            c1 = ESTADOS_COLORES.get(ests_relevantes[0], {}).get("bg", "#D9D9D9")
                            c2 = ESTADOS_COLORES.get(ests_relevantes[1], {}).get("bg", "#F1C40F")
                            cfg_main = ESTADOS_COLORES.get(ests_relevantes[0], {"fg": "#000000", "border": "#000000"})
                            bg_style = f"background: repeating-linear-gradient(45deg, {c1}, {c1} 10px, {c2} 10px, {c2} 20px); color: {cfg_main['fg']}; border: 2px solid {cfg_main['border']};"

                        # Comprobar si forma parte de un ensamble
                        es_ensamble_pz = pz.get("es_ensamble", False)
                        icono_ensamble = " ⚙️" if es_ensamble_pz else ""
                        
                        # En la descripción del tooltip flotante:
                        tooltip_html_lines = [
                            f"<b>⚙️ Pieza:</b> {num_parte}",
                            f"<b>📦 Cantidad:</b> {cant_pz} pzs",
                            f"<b>📅 Entrega:</b> {fecha_pz}",
                            f"<b>🏷️ Estados:</b> {', '.join(ests_pz)}"
                        ]
                        # En el tooltip flotante:
                        if es_ensamble_pz:
                            tooltip_html_lines.append("<b>⚙️ Ensamble:</b> Parte de un ensamble")
                        if pz["comentario_pz"]:
                            tooltip_html_lines.append(f"<b>💬 Nota Pieza:</b> {pz['comentario_pz']}")
                        
                        contenido_tooltip = "<br>".join(tooltip_html_lines)
                        
                        # En el texto visible de la celda:
                        celda_html = f'''<div class="custom-tooltip-wrapper" style="{bg_style} padding:7px 6px; margin-bottom:5px; border-radius:4px; font-size:15px; font-weight:bold; text-align:center; box-shadow:0 1px 2px rgba(0,0,0,0.08);">
                        <span style="display:block; overflow:hidden; text-overflow:ellipsis;">{num_parte}{icono_ensamble}</span>
                        <div class="custom-tooltip-text">{contenido_tooltip}</div>
                        </div>'''
                    
                        bloques_piezas.append(celda_html)

                    html_piezas = "".join(bloques_piezas)

                    if esta_en_prod:
                        bg_po_header = cfg_prod["bg"]
                        fg_po_header = cfg_prod["fg"]
                        border_po_header = f"2px dashed {cfg_prod['border']}"
                    else:
                        bg_po_header = "#004A99"
                        fg_po_header = "#FFFFFF"
                        border_po_header = "1px solid #003366"

                    tooltip_po_header = ""
                    indicador_nota = ""
                    if coment_po_k:
                        indicador_nota = " 📌"
                        tooltip_po_header = f'''<div class="custom-tooltip-text" style="font-size:12px; font-weight:normal; text-align:left;">
<b>📌 Nota de la PO:</b><br>{coment_po_k}
</div>'''

                    header_po_html = f'''<div class="custom-tooltip-wrapper" style="background-color:{bg_po_header}; color:{fg_po_header}; border:{border_po_header}; font-weight:bold; text-align:center; padding:8px 6px; border-radius:5px; margin-bottom:6px; font-size:16px; box-shadow:0 1px 3px rgba(0,0,0,0.12);">
{po_k}{indicador_nota}
{tooltip_po_header}
</div>'''

                    col_html = f'''<td style="vertical-align:top; width:175px; min-width:175px;">
{header_po_html}
{html_piezas}
</td>'''
                    columnas_pos_html.append(col_html)

                html_columnas_unidas = "".join(columnas_pos_html)
    
                css_scroll = """<style>
.matriz-scroll-container {
    overflow-x: scroll !important;
    overflow-y: visible !important;
    white-space: nowrap;
    background-color: #FAFAFA;
    padding: 16px 14px 24px 14px;
    border-radius: 8px;
    border: 1px solid #D1D5DB;
    box-shadow: inset 0 1px 3px rgba(0,0,0,0.05);
}
.matriz-scroll-container::-webkit-scrollbar {
    height: 14px !important;
}
.matriz-scroll-container::-webkit-scrollbar-track {
    background: #E5E7EB !important;
    border-radius: 8px;
}
.matriz-scroll-container::-webkit-scrollbar-thumb {
    background: #004A99 !important;
    border-radius: 8px;
    border: 2px solid #E5E7EB;
}
.matriz-scroll-container::-webkit-scrollbar-thumb:hover {
    background: #003366 !important;
}
.columna-fija-acciones {
    position: sticky !important;
    left: 0 !important;
    z-index: 10 !important;
    background-color: #FAFAFA !important;
    box-shadow: 4px 0 6px -2px rgba(0,0,0,0.15);
}
.custom-tooltip-wrapper {
    position: relative;
    cursor: pointer;
}
.custom-tooltip-wrapper .custom-tooltip-text {
    visibility: hidden;
    width: 250px;
    background-color: #111827;
    color: #FFFFFF;
    text-align: left;
    border-radius: 7px;
    padding: 9px 12px;
    position: absolute;
    z-index: 99999;
    top: 50%;
    left: 105%;
    transform: translateY(-50%);
    font-size: 12px !important;
    font-weight: normal !important;
    line-height: 1.45;
    white-space: normal;
    border: 1px solid #374151;
    box-shadow: 0 4px 14px rgba(0,0,0,0.35);
    opacity: 0;
    transition: opacity 0.15s ease-in-out;
    pointer-events: none;
}
.custom-tooltip-wrapper:hover .custom-tooltip-text {
    visibility: visible;
    opacity: 1;
}
.custom-tooltip-wrapper .custom-tooltip-text::after {
    content: "";
    position: absolute;
    top: 50%;
    right: 100%;
    margin-top: -6px;
    border-width: 6px;
    border-style: solid;
    border-color: transparent #111827 transparent transparent;
}
</style>"""

              # Montaje del contenedor HTML sin sangrías para evitar bloques de código planos
                html_tablero = (
                    f"{css_scroll}\n"
                    f'<div class="matriz-scroll-container">\n'
                    f'<table style="border-collapse:separate; border-spacing:6px; font-family:\'Segoe UI\', Arial, sans-serif;">\n'
                    f'<tr>\n'
                    f'<td class="columna-fija-acciones" style="vertical-align:top; width:160px; min-width:160px;">\n'
                    f'<div style="background-color:#374151; color:#FFFFFF; font-weight:bold; text-align:center; padding:8px 6px; border-radius:5px; margin-bottom:6px; font-size:16px;">Acciones</div>\n'
                    f'{html_acciones}\n'
                    f'</td>\n'
                    f'{html_columnas_unidas}\n'
                    f'</tr>\n'
                    f'</table>\n'
                    f'</div>'
                )

                st.markdown(html_tablero, unsafe_allow_html=True)

            # ===========================================
            # SECCIÓN 1.5 : DESCARGAR REPORTE POR PO
            # ===========================================
            st.markdown("---")
            st.subheader("REPORTE DE PO")

            todas_pos_registradas = sorted(list({str(r.get("po_number", "")).strip() for r in res_todas if r.get("po_number")}))
            opciones_filtro_po = ["-- Todas las POs --"] + todas_pos_registradas
    
            col_po_filtro, col_po_btn, col_po_pdf = st.columns([3.0, 1.2, 1.4])
            with col_po_filtro:
                po_filtrada_sel = st.selectbox(
                    "🔍 Filtrar panorama por número de PO:",
                    options=opciones_filtro_po,
                    key="filtro_po_reporte_pdf",
                    help="Selecciona o escribe el número de PO para ver únicamente su columna en el tablero"
                )
            with col_po_btn:
                st.write("")
                st.write("")
                if po_filtrada_sel != "-- Todas las POs --":
                    if st.button("Ir a editar esta PO ↗️", use_container_width=True):
                        st.session_state.po_activa = po_filtrada_sel
                        st.rerun()

            with col_po_pdf:
                st.write("")
                st.write("")
                po_para_reporte = po_filtrada_sel if po_filtrada_sel != "-- Todas las POs --" else st.session_state.get("po_activa")
                
                if po_para_reporte and po_para_reporte in pos_matriz:
                    datos_po_sel = pos_matriz[po_para_reporte]
                    
                    pdf_bytes = generar_pdf_reporte_po(
                        po_numero=po_para_reporte,
                        piezas_lista=datos_po_sel.get("piezas", []),
                        comentario_po=datos_po_sel.get("comentario", "")
                    )
                    
                    st.download_button(
                        label="📄 Descargar PDF",
                        data=pdf_bytes,
                        file_name=f"Reporte_PO_{po_para_reporte}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
    
            # ==========================================
            # SECCIÓN 2 (ABAJO): EDICIÓN Y GESTIÓN POR PO
            # ==========================================
            st.markdown("---")
            st.subheader("🛠️ Modificación y Detalle por PO")
    
            idx_actual = lista_pos.index(st.session_state.po_activa) if st.session_state.po_activa in lista_pos else 0
    
            po_seleccionada = st.selectbox(
                "Seleccione la PO a trabajar:",
                options=lista_pos,
                index=idx_actual,
                key="sel_po_superv"
            )
            st.session_state.po_activa = po_seleccionada
    
            res_piezas_po = (
                supabase.table("piezas")
                .select("*")
                .eq("po_number", po_seleccionada)
                .order("fecha_entrega")
                .order("id")
                .execute()
                .data or []
            )
    
            comentario_po_val = ""
            po_prod_val = False
            for p_chk in res_piezas_po:
                if p_chk.get("comentarios_po"):
                    comentario_po_val = p_chk.get("comentarios_po")
                if p_chk.get("po_en_produccion"):
                    po_prod_val = True
    
            c_po_tit, c_po_prod, c_po_input, c_po_btn = st.columns([1.8, 1.8, 3.4, 1.0])
    
            with c_po_tit:
                st.markdown(f"### 📦 PO: `{po_seleccionada}`")
                st.caption(f"**Total partidas:** {len(res_piezas_po)}")
    
            with c_po_prod:
                nuevo_en_prod = st.checkbox(
                    "🏭 En Producción",
                    value=po_prod_val,
                    key=f"chk_prod_po_{po_seleccionada}",
                    help="Marca toda la orden de compra con la etiqueta de Producción"
                )
    
            with c_po_input:
                nuevo_coment_po = st.text_input(
                    f"Comentario PO {po_seleccionada}:",
                    value=comentario_po_val,
                    placeholder="Ej. Entrega: 25-Nov / Falta material / Soldadura externa...",
                    key=f"txt_coment_po_{po_seleccionada}",
                    label_visibility="collapsed"
                )
    
            with c_po_btn:
                if st.button("💾 Guardar PO", key=f"btn_save_po_{po_seleccionada}", use_container_width=True):
                    supabase.table("piezas").update({
                        "comentarios_po": nuevo_coment_po.strip(),
                        "po_en_produccion": nuevo_en_prod
                    }).eq("po_number", po_seleccionada).execute()

                    # Registro de auditoría
                    if nuevo_en_prod != po_prod_val:
                        registrar_log_cambio(
                            po_number=po_seleccionada,
                            campo_editado="po_en_produccion",
                            valor_anterior=str(po_prod_val),
                            valor_nuevo=str(nuevo_en_prod),
                            part_num="-",
                            motivo=f"Cambio por {usuario_sup}"
                        )
                    if nuevo_coment_po.strip() != comentario_po_val.strip():
                        registrar_log_cambio(
                            po_number=po_seleccionada,
                            campo_editado="comentarios_po",
                            valor_anterior=comentario_po_val,
                            valor_nuevo=nuevo_coment_po.strip(),
                            part_num="-",
                            motivo=f"Nota PO editada por {usuario_sup}"
                        )

                    st.toast(f"Estatus de PO {po_seleccionada} actualizado", icon="📦")
                    st.rerun()
    
            pz_enfocada = st.session_state.get("pieza_enfoque", None)
    
            for p in res_piezas_po:
                id_pieza = p["id"]
                part_num = p.get("Part number", "Sin N/P")
                fecha_ent = p.get("fecha_entrega") or "Sin fecha"
                cant_sol = p.get("cantidad_solicitada") or p.get("cantidad") or 1
                coment_pieza_act = p.get("comentarios") or ""
    
                es_foco = (pz_enfocada == str(part_num).strip())
    
                raw_estados = p.get("estados_proceso")
                if not raw_estados:
                    est_v = p.get("estado_proceso") or "En proceso"
                    estados_actuales = [est_v] if (est_v in LISTA_ESTADOS_PIEZAS) else ["En proceso"]
                else:
                    estados_actuales = [e for e in raw_estados if e in LISTA_ESTADOS_PIEZAS]
                    if not estados_actuales:
                        estados_actuales = ["En proceso"]
    
                with st.container(border=True):
                    badges_html = []
                    for est in estados_actuales:
                        info_col = ESTADOS_COLORES.get(est, {"bg": "#D9D9D9", "fg": "#000000", "border": "#BFBFBF"})
                        badges_html.append(
                            f"""<span style="
                                display: inline-block;
                                background-color: {info_col['bg']};
                                color: {info_col['fg']};
                                border: 2px solid {info_col['border']};
                                padding: 7px 12px;
                                margin: 3px 5px 3px 0px;
                                border-radius: 6px;
                                font-weight: 700;
                                font-size: 13px;
                                box-shadow: 0 1px 3px rgba(0,0,0,0.1);
                            ">{est}</span>"""
                        )
                    html_etiquetas = "".join(badges_html)
    
                    col_cab_info, col_cab_badges = st.columns([2.5, 7.5])
                    with col_cab_info:
                        icono = "🎯 " if es_foco else "⚙️ "
                        st.markdown(f"### {icono}`{part_num}`")
                        st.caption(f"📅 **Entrega:** {fecha_ent} &nbsp;|&nbsp; 📦 **Cantidad:** {cant_sol} pzs")
                    with col_cab_badges:
                        st.markdown(f"<div style='padding-top: 4px;'>{html_etiquetas}</div>", unsafe_allow_html=True)
    
                    st.write("")
    
                    # 1. Leer el valor actual
                    es_ensamble_val = bool(p.get("es_ensamble", False))
                    
                    # 2. En las columnas de edición, agrega el checkbox (por ejemplo, redistribuyendo col_sel, col_coment, col_btn)
                    col_sel, col_coment, col_ensamble, col_btn = st.columns([4.2, 3.8, 1.2, 0.8])
                    
                    with col_sel:
                        nuevos_estados = st.multiselect(
                            "Editar etiquetas de la pieza:",
                            options=LISTA_ESTADOS_PIEZAS,
                            default=estados_actuales,
                            key=f"multi_est_{id_pieza}",
                            placeholder="Selecciona etiquetas..."
                        )
                    
                    with col_coment:
                        nuevo_coment_pz = st.text_input(
                            "Comentario técnico:",
                            value=coment_pieza_act,
                            placeholder="Ej. Soldar con Argón / Dejar sobremedida...",
                            key=f"txt_pz_{id_pieza}"
                        )
                    
                    with col_ensamble:
                        st.write("")
                        st.write("")
                        nuevo_es_ensamble = st.checkbox("⚙️ Ensamble", value=es_ensamble_val, key=f"chk_ensamble_{id_pieza}")
                    
                    with col_btn:
                        st.write("")
                        st.write("")
                        if st.button("💾", key=f"btn_save_pz_{id_pieza}", help=f"Guardar cambios en {part_num}", use_container_width=True):
                            estados_finales = nuevos_estados if nuevos_estados else ["En proceso"]
                            
                            # Guardar en base de datos incluyendo 'es_ensamble'
                            supabase.table("piezas").update({
                                "estados_proceso": estados_finales,
                                "estado_proceso": estados_finales[0],
                                "comentarios": nuevo_coment_pz.strip(),
                                "es_ensamble": nuevo_es_ensamble
                            }).eq("id", id_pieza).execute()
                    
                            # Registrar en la auditoría si cambió el estatus de ensamble
                            if nuevo_es_ensamble != es_ensamble_val:
                                registrar_log_cambio(
                                    po_number=po_seleccionada,
                                    campo_editado="es_ensamble",
                                    valor_anterior=str(es_ensamble_val),
                                    valor_nuevo=str(nuevo_es_ensamble),
                                    part_num=part_num,
                                    motivo=f"Marcado como ensamble por {usuario_sup}"
                                )
                    
                            st.session_state["pieza_enfoque"] = None
                            st.toast(f"Pieza {part_num} actualizada", icon="✅")
                            st.rerun()

    # ==========================================
    # PESTAÑA 2: REQUISICIONES
    # ==========================================
    with tab_req:
        st.subheader("📬 Validación de Requisiciones de Almacén")
    
        try:
            res_req_sup = supabase.table("requisiciones_atlas").select("*").eq("estatus", "Enviada").execute()
            reqs_pendientes = res_req_sup.data or []
    
            if reqs_pendientes:
                pos_con_req = sorted(list(set(str(r["po_number"]).strip() for r in reqs_pendientes if r.get("po_number"))))
                po_req_sel = st.selectbox("Seleccione la Requisición / PO a dictaminar:", pos_con_req, key="sup_po_req_select")
    
                items_a_revisar = [r for r in reqs_pendientes if str(r["po_number"]).strip() == po_req_sel]
                df_rev = pd.DataFrame(items_a_revisar)
    
                st.markdown(f"##### Partidas solicitadas para la PO: {po_req_sel}")
                cols_r_show = ["id", "tipo", "descripcion", "medida", "cantidad", "unidad", "urgencia", "motivo_uso", "solicitante", "link"]
                st.dataframe(df_rev[[c for c in cols_r_show if c in df_rev.columns]], use_container_width=True, hide_index=True)
    
                c_dict1, c_dict2 = st.columns(2)
    
                # ACCIÓN 1: RECHAZAR PARTIDA ESPECÍFICA O TODAS
                with c_dict1:
                    with st.expander("❌ Rechazar Partida Específica o Requisición", expanded=False):
                        st.markdown("Indica qué ítem requiere corrección por Inventarios:")
                        
                        opciones_items_rechazo = {
                            "⚠️ Rechazar TODA la requisición (todas las partidas)": "TODAS"
                        }
                        for it in items_a_revisar:
                            opciones_items_rechazo[f"ID {it['id']} - {it.get('descripcion')} ({it.get('medida', '')})"] = it['id']
    
                        item_a_rechazar = st.selectbox(
                            "Partida a devolver a Inventarios:", 
                            list(opciones_items_rechazo.keys()),
                            key=f"sel_rechazo_{po_req_sel}"
                        )
                        id_seleccionado = opciones_items_rechazo[item_a_rechazar]
    
                        motivo_rechazo = st.text_area(
                            "Motivo de corrección para Almacén:",
                            placeholder="Ej. Medida de cortador errónea, material agotado con proveedor...",
                            key=f"txt_motivo_{po_req_sel}"
                        )
    
                        if st.button("Devolver para Corrección", type="secondary", key=f"btn_dev_{po_req_sel}"):
                            if not motivo_rechazo.strip():
                                st.error("⚠️ Especifica el motivo del rechazo para que Inventarios sepa qué cambiar.")
                            else:
                                if id_seleccionado == "TODAS":
                                    supabase.table("requisiciones_atlas").update({
                                        "estatus": "Rechazada",
                                        "comentarios_supervisor": motivo_rechazo.strip()
                                    }).eq("po_number", po_req_sel).eq("estatus", "Enviada").execute()

                                    registrar_log_cambio(
                                        po_number=po_req_sel,
                                        campo_editado="requisicion",
                                        valor_anterior="Enviada",
                                        valor_nuevo="Rechazada (Todas)",
                                        part_num="-",
                                        motivo=f"Rechazada por {usuario_sup}: {motivo_rechazo.strip()}"
                                    )
                                    st.warning(f"Se rechazaron todas las partidas de la PO {po_req_sel}.")
                                else:
                                    supabase.table("requisiciones_atlas").update({
                                        "estatus": "Rechazada",
                                        "comentarios_supervisor": motivo_rechazo.strip()
                                    }).eq("id", id_seleccionado).execute()

                                    registrar_log_cambio(
                                        po_number=po_req_sel,
                                        campo_editado="requisicion_partida",
                                        valor_anterior="Enviada",
                                        valor_nuevo=f"Rechazada ID {id_seleccionado}",
                                        part_num="-",
                                        motivo=f"Devuelto por {usuario_sup}: {motivo_rechazo.strip()}"
                                    )
                                    st.warning(f"Partida ID {id_seleccionado} devuelta a Inventarios con la observación asignada.")
                                st.rerun()
    
                # ACCIÓN 2: APROBAR Y MANDAR CORREO CON PDF A GERENCIA
                with c_dict2:
                    with st.expander("✅ Aprobar y Notificar por Correo", expanded=True):
                        default_to = st.secrets.get("REQUISICION_TO", "ericslosberg@gmail.com")
                        default_cc = st.secrets.get("REQUISICION_CC", "shirsa.produccion@nae-supply.com, juan.rodriguez@nae-supply.com, diseno.dos.shirsa@nae-supply.com, soporte@nae-supply.com")
    
                        c_mail1, c_mail2 = st.columns(2)
                        with c_mail1:
                            input_to = st.text_input("Para (Destinatarios principales):", value=default_to, key=f"to_{po_req_sel}")
                        with c_mail2:
                            input_cc = st.text_input("CC (En copia):", value=default_cc, key=f"cc_{po_req_sel}")
    
                        pdf_bytes_req = generar_pdf_requisicion(po_req_sel, items_a_revisar)
    
                        st.download_button(
                            label="📄 Descargar PDF de la Requisición",
                            data=pdf_bytes_req,
                            file_name=f"Requisicion_PO_{po_req_sel}.pdf",
                            mime="application/pdf",
                            key=f"dl_pdf_{po_req_sel}"
                        )
    
                        st.write("")
                        if st.button("🚀 Aprobar Requisición y Enviar Email", type="primary", key=f"btn_aprob_{po_req_sel}"):
                            if not input_to.strip():
                                st.error("⚠️ Debe haber al menos un correo en el campo 'Para'.")
                            else:
                                with st.spinner("Enviando correo con PDF adjunto..."):
                                    try:
                                        enviar_correo_gerente(po_req_sel, pdf_bytes_req, input_to.strip(), input_cc.strip())
    
                                        supabase.table("requisiciones_atlas").update({
                                            "estatus": "Aprobada",
                                            "comentarios_supervisor": f"Aprobada por {usuario_sup} el {datetime.now().strftime('%d/%m/%Y %H:%M')}"
                                        }).eq("po_number", po_req_sel).eq("estatus", "Enviada").execute()

                                        registrar_log_cambio(po_number=po_req_sel,
                                            campo_editado="requisicion",
                                            valor_anterior="Enviada",
                                            valor_nuevo="Aprobada",
                                            part_num="-",
                                            motivo=f"Aprobada y enviada a compras por {usuario_sup}"
                                        )
    
                                        st.success(f"✅ Requisición de la PO {po_req_sel} aprobada y enviada a los destinatarios.")
                                        st.rerun()
                                    except Exception as e_envio:
                                        st.error(f"Error al enviar correo: {e_envio}")
            else:
                st.info("No hay requisiciones pendientes de validación en este momento.")
    
        except Exception as e_r_sup:
            st.error(f"Error en validación de requisiciones: {e_r_sup}")

# ==========================================
    # PESTAÑA: DESCARGA DE HISTORIALES (PDF)
    # ==========================================
    with tab_historial_pdf:
        st.subheader("📥 Exportación de Bitácoras de Auditoría")
        
        col_sup_box, col_maq_box = st.columns(2)

        # 1. BITÁCORA DE SUPERVISORES
        with col_sup_box:
            with st.container(border=True):
                st.markdown("#### 📋 Bitácora de Supervisión")
                st.caption("Cambios de estatus de POs, piezas y validación de requisiciones.")

                try:
                    res_sup = supabase.table("historial_cambios").select("*").execute()
                    logs_sup = res_sup.data or []
                    if logs_sup:
                        logs_sup_ord = sorted(logs_sup, key=lambda x: str(x.get("fecha_hora", "")), reverse=True)[:300]
                        pdf_bytes_sup = generar_pdf_historial_cambios(logs_sup_ord)
                        st.download_button(
                            label="📄 Descargar Bitácora de Supervisión (PDF)",
                            data=pdf_bytes_sup,
                            file_name=f"Bitacora_Supervision_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                        st.caption(f"Registros disponibles: {len(logs_sup_ord)} eventos.")
                    else:
                        st.info("Sin registros de supervisión.")
                except Exception as e_sup:
                    st.warning(f"Error al leer bitácora de supervisión: {e_sup}")

        # 2. BITÁCORA DE MAQUINISTAS
        with col_maq_box:
            with st.container(border=True):
                st.markdown("#### 🛠️ Bitácora de Maquinistas")
                st.caption("Herramientas agregadas/removidas, material y piezas terminadas.")

                try:
                    res_maq = supabase.table("historial_maquinistas").select("*").execute()
                    logs_maq = res_maq.data or []
                    if logs_maq:
                        logs_maq_ord = sorted(logs_maq, key=lambda x: str(x.get("fecha_hora", "")), reverse=True)[:300]
                        pdf_bytes_maq = generar_pdf_historial_maquinistas(logs_maq_ord)
                        st.download_button(
                            label="📄 Descargar Bitácora de Maquinistas (PDF)",
                            data=pdf_bytes_maq,
                            file_name=f"Bitacora_Maquinistas_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                        st.caption(f"Registros disponibles: {len(logs_maq_ord)} eventos.")
                    else:
                        st.info("Sin registros de maquinistas.")
                except Exception as e_maq:
                    st.warning(f"Error al leer bitácora de maquinistas: {e_maq}")





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
        
