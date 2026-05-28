import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import io, os, tempfile, re, time
from docx import Document
import google.generativeai as genai
from PIL import Image

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Presentaciones Comerciales Universales", layout="wide")
st.title("🚀 Generador Universal de Proyectos de Inversión")
st.markdown("Carga cualquier proyecto (Lavaderos, Apps, Inmuebles, etc.) y genera una propuesta comercial de alto nivel.")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Configuración")
    api_key = st.text_input("Introduce tu Google API Key", type="password")
    moneda_local = st.text_input("Moneda Local (ej: UYU, ARS, MXN)", value="UYU")
    st.divider()
    st.info("La IA detectará automáticamente el sector y aplicará el diseño y KPIs adecuados.")

# --- FUNCIONES TÉCNICAS ---
def leer_archivo(file):
    if file.name.lower().endswith('.docx'):
        return "\n".join([p.text for p in Document(file).paragraphs])
    return file.read().decode("utf-8")

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))

# --- MOTOR DE DISEÑO DINÁMICO ---
def crear_slide_pro(prs, titulo, contenido, color_hex, imagen=None):
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    # Fondo
    bg_color = hex_to_rgb(color_hex)
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(20, 20, 20) # Fondo oscuro neutro
    
    # Barra decorativa lateral
    shape = slide.shapes.add_shape(1, 0, 0, Inches(0.1), Inches(7.5))
    shape.fill.solid()
    shape.fill.foreground_color.rgb = bg_color
    shape.line.fill.background()

    # Título
    title_shape = slide.shapes.title
    title_shape.text = titulo.upper()
    p = title_shape.text_frame.paragraphs[0]
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = bg_color
    p.alignment = PP_ALIGN.LEFT

    # Cuerpo
    body_shape = slide.placeholders[1]
    body_shape.width = Inches(5.5) if imagen else Inches(9)
    tf = body_shape.text_frame
    tf.word_wrap = True
    
    for linea in contenido.split('\n'):
        if len(linea.strip()) > 3:
            p = tf.add_paragraph()
            p.text = "• " + linea.strip().replace('*', '').replace('-', '')
            p.font.size = Pt(16)
            p.font.color.rgb = RGBColor(240, 240, 240)
            p.space_after = Pt(10)

    # Imagen con marco
    if imagen:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            imagen.save(tmp.name)
            slide.shapes.add_picture(tmp.name, Inches(5.8), Inches(1.5), height=Inches(4.5))

# --- INTERFAZ DE USUARIO ---
col1, col2, col3 = st.columns(3)
with col1:
    st.subheader("📁 Datos")
    f_doc = st.file_uploader("Contrato/Charla/Word", type=["docx", "txt"])
    f_xlsx = st.file_uploader("Planilla Financiera", type=["xlsx"])
with col2:
    st.subheader("🎥 Multimedia")
    f_media = st.file_uploader("Audio o Video", type=["mp4", "mp3", "wav"])
    f_foto = st.file_uploader("Imagen/Logo/Fachada", type=["jpg", "png", "jpeg"])
with col3:
    st.subheader("✍️ Contexto")
    notas = st.text_area("Explica el proyecto brevemente:", height=150, placeholder="Ej: Lavadero automático en zona transitada...")

# --- PROCESAMIENTO ---
if st.button("🏗️ GENERAR PRESENTACIÓN COMERCIAL"):
    if not api_key:
        st.error("Falta API Key")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-1.5-flash')
            
            with st.spinner("Analizando modelo de negocio y preparando diseño..."):
                # 1. Recolectar datos
                contexto = f"MONEDA LOCAL: {moneda_local}\nCONTEXTO: {notas}\n"
                if f_doc: contexto += f"DOC: {leer_archivo(f_doc)}\n"
                if f_xlsx: contexto += f"EXCEL: {pd.read_excel(f_xlsx).to_string()}\n"
                
                inputs = [contexto]
                if f_foto: inputs.append(Image.open(f_foto))
                if f_media:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f_media.name)[1]) as tmp:
                        tmp.write(f_media.read())
                        t_path = tmp.name
                    gf = genai.upload_file(path=t_path)
                    while gf.state.name == "PROCESSING": time.sleep(2); gf = genai.get_file(gf.name)
                    inputs.append(gf)

                # 2. PROMPT DINÁMICO UNIVERSAL
                prompt = f"""
                Eres un Consultor de Inversiones de élite. Analiza este proyecto de cualquier sector.
                TAREAS:
                1. Detecta el sector del negocio y define una paleta de colores HEX (un solo color principal vibrante).
                2. Calcula/Estima KPIs específicos: ROTE, ROI, Cost to Income, Punto de Equilibrio y Ticket Promedio (si aplica).
                3. Proyecta flujos en {moneda_local} y USD.
                4. Explica beneficios fiscales y rentabilidad Pre y Post impuestos.
                
                GENERA 10 DIAPOSITIVAS:
                Usa este formato exacto: 
                COLOR: #HEX
                SLIDE: Título | Contenido en puntos clave (separados por *)
                
                Diapositivas: 1. Portada, 2. El Problema/Oportunidad, 3. Público Objetivo, 4. Características del Negocio, 
                5. Zonas/Ubicación, 6. Costos de Inversión inicial, 7. Ingresos y Gastos mensuales, 
                8. Beneficios Fiscales y Tax Efficiency, 9. KPIs y Proyecciones, 10. Conclusión y Branding.
                """
                
                res = model.generate_content([prompt] + inputs)
                raw_text = res.text

                # 3. EXTRAER COLOR Y CONSTRUIR PPTX
                color_match = re.search(r'COLOR:\s*(#[0-9A-Fa-f]{6})', raw_text)
                color_hex = color_match.group(1) if color_match else "#D4AF37"
                
                prs = Presentation()
                bloques = raw_text.split("SLIDE:")
                
                for i, bloque in enumerate(bloques):
                    if "|" in bloque:
                        partes = bloque.split("|")
                        titulo = partes[0].replace("COLOR:", "").strip()
                        contenido = partes[1].strip().replace("*", "\n")
                        
                        img_para_slide = Image.open(f_foto) if (i <= 2 and f_foto) else None
                        crear_slide_pro(prs, titulo, contenido, color_hex, img_para_slide)

                # Descarga
                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                st.success(f"✅ ¡Propuesta de {tipo_negocio if 'tipo_negocio' in locals() else 'Negocio'} lista!")
                st.download_button("📥 Descargar Presentación Pro", buf, "Propuesta_Comercial_Universal.pptx")
                
                if f_media: os.remove(t_path)
                with st.expander("Auditoría de Datos (Lo que la IA analizó)"):
                    st.write(raw_text)

        except Exception as e:
            st.error(f"Error: {e}")
