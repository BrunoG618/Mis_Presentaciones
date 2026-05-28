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
import re

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Presentaciones Comerciales Pro", layout="wide")
st.title("🏆 Generador de Propuestas de Inversión")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input("Introduce tu Google API Key", type="password")
    tipo_negocio = st.selectbox("Sector", ["Reciclaje Inmobiliario", "Inversión Comercial", "Otro"])
    st.info("Detectando automáticamente el mejor modelo disponible...")

# --- FUNCIONES DE LECTURA ---
def leer_texto_seguro(file):
    if file.name.lower().endswith('.docx'):
        try:
            doc = Document(file)
            return "\n".join([p.text for p in doc.paragraphs])
        except: return "Error leyendo Word."
    else:
        try: return file.read().decode("utf-8")
        except: return "Error leyendo TXT."

def aplicar_formato_slide(slide, titulo_texto):
    """Aplica diseño corporativo con corrección de color"""
    try:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(245, 245, 245)
        title = slide.shapes.title
        title.text = titulo_texto
        p = title.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(28)
        p.font.color.rgb = RGBColor(0, 51, 102)
    except: pass

# --- INTERFAZ ---
st.subheader("📁 Carga de Información")
col1, col2, col3 = st.columns(3)
with col1:
    f_doc = st.file_uploader("Conversación (Word o TXT)", type=["docx", "txt"])
    f_xlsx = st.file_uploader("Planilla Excel", type=["xlsx"])
with col2:
    f_media = st.file_uploader("Audio o Video", type=["mp3", "mp4", "wav", "m4a"])
    f_notas = st.text_area("Notas adicionales:", height=100)
with col3:
    f_foto = st.file_uploader("Imagen de Fachada", type=["jpg", "png", "jpeg"])

# --- PROCESAMIENTO ---
if st.button("🚀 GENERAR PRESENTACIÓN"):
    if not api_key:
        st.error("Por favor, introduce tu API Key.")
    else:
        try:
            genai.configure(api_key=api_key)
            models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            model_name = next((m for m in models if '1.5-flash' in m), models[0])
            model = genai.GenerativeModel(model_name)
            
            with st.spinner("Procesando datos financieros y visuales..."):
                # 1. Recolectar datos
                contexto = f"PROYECTO: {tipo_negocio}\n"
                if f_doc: contexto += f"CONTENIDO: {leer_texto_seguro(f_doc)}\n"
                if f_notas: contexto += f"NOTAS: {f_notas}\n"
                if f_xlsx:
                    df = pd.read_excel(f_xlsx)
                    contexto += f"DATOS EXCEL: {df.to_string()}\n"
                
                inputs_ia = [contexto]
                if f_foto:
                    img = Image.open(f_foto)
                    inputs_ia.append(img)
                
                if f_media:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f_media.name)[1]) as tmp:
                        tmp.write(f_media.read())
                        tmp_path = tmp.name
                    g_file = genai.upload_file(path=tmp_path)
                    while g_file.state.name == "PROCESSING":
                        time.sleep(2)
                        g_file = genai.get_file(g_file.name)
                    inputs_ia.append(g_file)

                # 2. IA genera contenido
                prompt = """
                Eres un analista de inversiones en Uruguay. Analiza todo lo adjunto.
                Genera el contenido de 8 diapositivas.
                Formato obligatorio: DIAPOSITIVA | TÍTULO | CONTENIDO (puntos con *)
                
                Contenido:
                1. Portada. 2. Visión y Público. 3. Análisis de Fachada (si hay foto).
                4. Inversión Inicial y Gastos (UYU y USD). 
                5. KPIs: ROTE, ROI, Cost to Income, Punto de Equilibrio.
                6. Beneficios Fiscales (Vivienda Promovida Uruguay).
                7. Sugerencia Estética (Colores HEX y Tipografías). 
                8. Conclusión.
                
                Limpia el texto de asteriscos dobles.
                """
                res = model.generate_content([prompt] + inputs_ia)
                texto_ia = res.text

                # 3. Crear PowerPoint
                prs = Presentation()
                
                # CREAR PORTADA INMEDIATAMENTE PARA EVITAR ERROR DE INDEX
                slide_portada = prs.slides.add_slide(prs.slide_layouts[0])
                slide_portada.shapes.title.text = f"PROYECTO: {tipo_negocio}"
                slide_portada.placeholders[1].text = "Análisis de Inversión Estratégica\nMontevideo, Uruguay"
                
                # Insertar imagen en la portada si existe
                if f_foto:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
                        img.save(tmp_img.name)
                        slide_portada.shapes.add_picture(tmp_img.name, Inches(5.5), Inches(1.5), width=Inches(4))

                # Procesar el resto de diapositivas de la IA
                # Usamos una búsqueda más flexible (Mayúsculas o minúsculas)
                slides_content = re.split(r'DIAPOSITIVA|Diapositiva', texto_ia)
                
                for bloque in slides_content:
                    if "|" in bloque:
                        partes = bloque.split("|")
                        if len(partes) >= 3:
                            slide = prs.slides.add_slide(prs.slide_layouts[1])
                            aplicar_formato_slide(slide, partes[1].strip())
                            
                            tf = slide.placeholders[1].text_frame
                            tf.word_wrap = True
                            for line in partes[2].strip().split("*"):
                                if len(line.strip()) > 3:
                                    p = tf.add_paragraph()
                                    p.text = "• " + line.strip().replace("**", "")
                                    p.font.size = Pt(18)

                # 4. Descarga
                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                st.success("✅ ¡Presentación generada con éxito!")
                st.download_button("📥 Descargar Propuesta Comercial", buf, "Propuesta_Inversor.pptx")
                
                if f_media: os.remove(tmp_path)

        except Exception as e:
            st.error(f"Error técnico detectado: {e}"
