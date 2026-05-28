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

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Presentaciones Comerciales", layout="wide")
st.title("🏛️ Presentaciones Comerciales: Versión Definitiva")

# --- ESTÉTICA ---
AZUL_DARK = RGBColor(11, 46, 81)
NARANJA_OBRA = RGBColor(230, 126, 34)

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input("Introduce tu Google API Key", type="password")
    tipo_negocio = st.selectbox("Sector", ["Reciclaje Inmobiliario", "Inversión Inmobiliaria", "Retail", "Gastronomía", "Otro"])
    st.info("Esta versión procesa: Audio, Video, Excel, Word, TXT e Imágenes.")

# --- FUNCIONES DE LECTURA BLINDADAS ---
def safe_read_doc(file):
    """Detecta el tipo de archivo y lo lee sin errores de ZIP"""
    try:
        if file.name.lower().endswith(".docx"):
            doc = Document(file)
            return "\n".join([p.text for p in doc.paragraphs])
        else:
            return file.read().decode("utf-8")
    except Exception as e:
        return f"Error leyendo texto: {e}"

def safe_read_excel(file):
    try:
        df = pd.read_excel(file)
        return df.to_string()
    except Exception as e:
        return f"Error leyendo Excel: {e}"

def aplicar_diseno(slide, titulo_texto):
    """Diseño profesional para cada diapositiva"""
    title = slide.shapes.title
    title.text = titulo_texto
    p = title.text_frame.paragraphs[0]
    p.font.bold = True
    p.font.size = Pt(30)
    p.font.color.rgb = AZUL_DARK
    # Línea decorativa
    line = slide.shapes.add_shape(1, Inches(0.5), Inches(1.15), Inches(3.5), Inches(0.04))
    line.fill.solid()
    line.fill.foreground_color.rgb = NARANJA_OBRA
    line.line.fill.background()

# --- INTERFAZ DE CARGA MULTIMODAL ---
st.subheader("📁 Ingesta de Datos del Proyecto")
col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("🔍 **Documentos y Datos**")
    f_doc = st.file_uploader("Conversación (Word o TXT)", type=["docx", "txt"])
    f_xlsx = st.file_uploader("Excel Financiero", type=["xlsx"])

with col2:
    st.markdown("🎥 **Multimedia**")
    f_media = st.file_uploader("Audio o Video de la charla", type=["mp3", "mp4", "wav", "m4a"])
    f_notas = st.text_area("Notas o detalles adicionales:", height=150)

with col3:
    st.markdown("🖼️ **Visuales**")
    f_foto = st.file_uploader("Imagen de Fachada o Planos", type=["jpg", "png", "jpeg"])

# --- PROCESAMIENTO ---
if st.button("🚀 GENERAR PRESENTACIÓN COMPLETA"):
    if not api_key:
        st.error("Por favor, introduce tu API Key.")
    else:
        try:
            genai.configure(api_key=api_key)
            
            # 1. Buscador dinámico de modelo (Evita Error 404)
            modelos_disponibles = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            target_model = next((m for m in modelos_disponibles if '1.5-flash' in m), modelos_disponibles[0])
            model = genai.GenerativeModel(target_model)

            with st.spinner("Procesando información multimodal y diseñando propuesta..."):
                
                # 2. Recolección de toda la información
                contexto_total = f"PROYECTO: {tipo_negocio}\n"
                inputs_ia = []
                
                if f_doc: contexto_total += f"DOCUMENTO: {safe_read_doc(f_doc)}\n"
                if f_xlsx: contexto_total += f"EXCEL: {safe_read_excel(f_xlsx)}\n"
                if f_notas: contexto_total += f"NOTAS: {f_notas}\n"
                
                inputs_ia.append(contexto_total)

                # Procesar Foto
                if f_foto:
                    img = Image.open(f_foto)
                    inputs_ia.append(img)

                # Procesar Multimedia (Audio/Video)
                if f_media:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f_media.name)[1]) as tmp:
                        tmp.write(f_media.read())
                        tmp_path = tmp.name
                    g_file = genai.upload_file(path=tmp_path)
                    while g_file.state.name == "PROCESSING":
                        time.sleep(2)
                        g_file = genai.get_file(g_file.name)
                    inputs_ia.append(g_file)

                # 3. PROMPT DE NEGOCIO (REQUERIMIENTOS)
                prompt = f"""
                Analiza TODA la info adjunta para este proyecto de {tipo_negocio} en Uruguay.
                Crea una presentación de 9 diapositivas para inversores.
                Usa el separador '---' entre cada diapositiva.
                Formato de cada bloque: Título | Contenido (puntos con *)
                
                REQUERIMIENTOS DE CONTENIDO:
                - Introducción, Público objetivo y Zonas de inversión.
                - Principales características del negocio.
                - Costos de inversión y gastos mensuales (UYU y USD).
                - Beneficios Fiscales Uruguay (Vivienda Promovida / Ley 18.795).
                - KPIs: Punto de Equilibrio, ROTE, ROI y Cost to Income.
                - Sugerencia de diseño: Colores HEX y tipografías acordes al proyecto.
                """

                # 4. LLAMADA A IA Y CONSTRUCCIÓN DE PPTX
                res = model.generate_content([prompt] + inputs_ia)
                texto_ia = res.text
                
                prs = Presentation()
                bloques = texto_ia.split("---")
                
                for bloque in bloques:
                    if "|" in bloque:
                        partes = bloque.split("|")
                        if len(partes) >= 2:
                            slide = prs.slides.add_slide(prs.slide_layouts[1])
                            aplicar_diseno(slide, partes[0].strip().replace("**", ""))
                            
                            tf = slide.placeholders[1].text_frame
                            tf.word_wrap = True
                            for p_text in partes[1].strip().split("*"):
                                if len(p_text.strip()) > 3:
                                    p = tf.add_paragraph()
                                    p.text = "• " + p_text.strip().replace("**", "")
                                    p.font.size = Pt(17)

                # Insertar foto en la portada si existe
                if f_foto:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
                        img.save(tmp_img.name)
                        prs.slides[0].shapes.add_picture(tmp_img.name, Inches(6), Inches(1.4), width=Inches(3.4))

                # 5. DESCARGA
                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                
                st.success("✅ ¡Presentación generada con éxito!")
                st.download_button("📥 Descargar Propuesta Comercial", buf, "Propuesta_Comercial.pptx")
                
                if f_media: os.remove(tmp_path)

        except Exception as e:
            st.error(f"Error técnico detectado: {e}")
