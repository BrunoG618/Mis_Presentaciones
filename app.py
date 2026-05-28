import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import io
import os
import tempfile
from docx import Document
import google.generativeai as genai
from PIL import Image
import time

# --- CONFIGURACIÓN ESTÉTICA ---
AZUL_INVERSOR = RGBColor(0, 32, 96)
NARANJA_OBRA = RGBColor(237, 125, 49)
GRIS_FONDO = RGBColor(242, 242, 242)

st.set_page_config(page_title="Presentaciones Comerciales Pro", layout="wide")
st.title("💼 Presentaciones Comerciales Inteligentes")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input("Introduce tu Google API Key", type="password")
    tipo_negocio = st.selectbox("Sector del Proyecto", ["Reciclaje Inmobiliario", "Inversión Comercial", "Retail", "Gastronomía", "Tech"])
    st.info("Esta herramienta genera KPIs: ROTE, ROI, Cost to Income y Punto de Equilibrio.")

# --- FUNCIONES DE SOPORTE ---
def leer_word(file):
    doc = Document(file)
    return "\n".join([p.text for p in doc.paragraphs])

def aplicar_diseno(slide, titulo_texto):
    """Aplica un diseño ejecutivo a cada diapositiva"""
    slide.background.fill.solid()
    slide.background.fill.foreground_color.rgb = GRIS_FONDO
    
    title = slide.shapes.title
    title.text = titulo_texto
    p = title.text_frame.paragraphs[0]
    p.font.bold = True
    p.font.size = Pt(30)
    p.font.color.rgb = AZUL_INVERSOR
    
    # Línea decorativa
    line = slide.shapes.add_shape(1, Inches(0.5), Inches(1.1), Inches(4), Inches(0.04))
    line.fill.solid()
    line.fill.foreground_color.rgb = NARANJA_OBRA
    line.line.fill.background()

# --- INTERFAZ DE CARGA (Organizada por tipo) ---
st.subheader("📁 Carga de Información del Proyecto")
col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("**📄 Texto y Datos**")
    f_doc = st.file_uploader("Conversación (Word o TXT)", type=["docx", "txt"])
    f_xlsx = st.file_uploader("Planilla Financiera (Excel)", type=["xlsx"])

with col2:
    st.markdown("**🎙️ Multimedia**")
    f_media = st.file_uploader("Grabación Audio o Video", type=["mp3", "mp4", "wav", "m4a"])
    f_notas = st.text_area("Notas manuales o detalles extra:", height=100)

with col3:
    st.markdown("**📸 Visuales**")
    f_foto = st.file_uploader("Imagen de Fachada o Planos", type=["jpg", "png", "jpeg"])

# --- PROCESAMIENTO ---
if st.button("🚀 GENERAR PRESENTACIÓN EJECUTIVA"):
    if not api_key:
        st.error("Por favor, introduce tu API Key.")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-1.5-flash')
            
            with st.spinner("Analizando multimodalidad y calculando rentabilidad..."):
                
                # 1. Preparar Inputs para la IA
                prompt_parts = []
                texto_contexto = f"PROYECTO: {tipo_negocio}\n"
                
                if f_doc: texto_contexto += f"DOCUMENTO: {leer_word(f_doc)}\n"
                if f_notas: texto_contexto += f"NOTAS: {f_notas}\n"
                if f_xlsx:
                    df = pd.read_excel(f_xlsx)
                    texto_contexto += f"EXCEL FINANCIERO: {df.to_string()}\n"
                
                prompt_parts.append(texto_contexto)

                # Procesar Foto
                if f_foto:
                    img = Image.open(f_foto)
                    prompt_parts.append(img)

                # Procesar Audio/Video
                if f_media:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f_media.name)[1]) as tmp:
                        tmp.write(f_media.read())
                        tmp_path = tmp.name
                    g_file = genai.upload_file(path=tmp_path)
                    while g_file.state.name == "PROCESSING":
                        time.sleep(2)
                        g_file = genai.get_file(g_file.name)
                    prompt_parts.append(g_file)

                # 2. EL PROMPT MAESTRO (KPIs y Diseño)
                master_prompt = f"""
                Eres un analista financiero y de diseño estratégico. Analiza TODO lo adjunto.
                Genera el contenido de una presentación de 9 diapositivas para inversores.
                
                KPIs OBLIGATORIOS: ROTE, ROI, Cost to Income, Punto de Equilibrio, Flujos en UYU y USD.
                CONTEXTO LEGAL: Beneficios Fiscales Uruguay (Vivienda Promovida / Ley 18.795).
                DISEÑO: Sugiere Paleta de Colores (Códigos HEX) y Tipografías que se adapten al proyecto.
                
                Formato de respuesta: DIAPOSITIVA | TÍTULO | CONTENIDO (puntos con *)
                
                Diapositivas: 1. Portada, 2. Introducción, 3. Público Objetivo, 4. Ubicación/Zonas, 
                5. Características del Negocio, 6. Costos Iniciales y Gastos, 7. Beneficios Fiscales, 
                8. KPIs y Rentabilidad, 9. Conclusión y Diseño Sugerido.
                """
                
                prompt_parts.insert(0, master_prompt)
                res = model.generate_content(prompt_parts)
                texto_ia = res.text

                # 3. CONSTRUCCIÓN DEL PPTX
                prs = Presentation()
                
                for bloque in texto_ia.split("DIAPOSITIVA"):
                    if "|" in bloque:
                        partes = bloque.split("|")
                        if len(partes) >= 3:
                            slide = prs.slides.add_slide(prs.slide_layouts[1])
                            aplicar_diseno(slide, partes[1].strip())
                            
                            tf = slide.placeholders[1].text_frame
                            tf.word_wrap = True
                            for line in partes[2].strip().split("*"):
                                if len(line.strip()) > 3:
                                    p = tf.add_paragraph()
                                    p.text = "• " + line.strip().replace("**", "")
                                    p.font.size = Pt(17)
                                    p.font.color.rgb = RGBColor(60,60,60)

                # Agregar la foto en la portada si existe
                if f_foto:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
                        img.save(tmp_img.name)
                        prs.slides[0].shapes.add_picture(tmp_img.name, Inches(6), Inches(1.5), width=Inches(3.5))

                # Descarga
                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                
                st.success("✅ ¡Presentación Completa Generada!")
                st.download_button("📥 Descargar Presentación para Inversores", buf, "Propuesta_Comercial_Final.pptx")
                
                if f_media: os.remove(tmp_path)

        except Exception as e:
            st.error(f"Error técnico detectado: {e}")
