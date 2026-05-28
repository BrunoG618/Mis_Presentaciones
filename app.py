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
st.set_page_config(page_title="Presentaciones Comerciales Pro", layout="wide")
st.title("🏆 Generador de Propuestas de Inversión")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input("Introduce tu Google API Key", type="password")
    tipo_negocio = st.selectbox("Sector", ["Reciclaje Inmobiliario", "Inversión Comercial", "Otro"])
    st.info("Esta versión detecta automáticamente el modelo de IA disponible en tu cuenta.")

# --- FUNCIONES DE LECTURA SEGURA ---
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
    """Aplica diseño corporativo corrigiendo el error de fore_color"""
    # Fondo Gris
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(245, 245, 245)
    
    # Título Azul
    title = slide.shapes.title
    title.text = titulo_texto
    p = title.text_frame.paragraphs[0]
    p.font.bold = True
    p.font.size = Pt(28)
    p.font.color.rgb = RGBColor(0, 51, 102) # Azul Marino

# --- INTERFAZ DE CARGA ---
st.subheader("📁 Carga de Archivos")
col1, col2, col3 = st.columns(3)

with col1:
    f_doc = st.file_uploader("Documento (Word o TXT)", type=["docx", "txt"])
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
            
            # SOLUCIÓN AL ERROR 404: Detección automática de modelo
            modelos_disponibles = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            modelo_nombre = next((m for m in modelos_disponibles if '1.5-flash' in m), modelos_disponibles[0])
            model = genai.GenerativeModel(modelo_nombre)
            
            with st.spinner(f"Conectado a {modelo_nombre}. Analizando proyecto..."):
                # Recolectar Información
                contexto = f"PROYECTO: {tipo_negocio}\n"
                if f_doc: contexto += f"TEXTO: {leer_texto_seguro(f_doc)}\n"
                if f_notas: contexto += f"NOTAS: {f_notas}\n"
                if f_xlsx:
                    df = pd.read_excel(f_xlsx)
                    contexto += f"DATOS EXCEL: {df.to_string()}\n"
                
                prompt_input = [contexto]
                if f_foto:
                    img = Image.open(f_foto)
                    prompt_input.append(img)
                
                # Manejo de Audio/Video con espera activa
                if f_media:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f_media.name)[1]) as tmp:
                        tmp.write(f_media.read())
                        tmp_path = tmp.name
                    g_file = genai.upload_file(path=tmp_path)
                    while g_file.state.name == "PROCESSING":
                        time.sleep(2)
                        g_file = genai.get_file(g_file.name)
                    prompt_input.append(g_file)

                # PROMPT MAESTRO
                prompt_ia = """
                Analiza todo el material adjunto para un proyecto de inversión en Uruguay.
                Genera el contenido de 8 diapositivas profesionales.
                Formato: DIAPOSITIVA | TÍTULO | CONTENIDO (puntos con *)
                
                REQUERIMIENTO:
                1. Portada con visión general.
                2. Introducción y Público Objetivo.
                3. Análisis de Fachada y Diseño (si hay foto).
                4. Inversión Inicial y Gastos Mensuales (en UYU y USD).
                5. KPIs: ROTE, ROI, Cost to Income y Punto de Equilibrio.
                6. Beneficios Fiscales (Vivienda Promovida Ley 18.795).
                7. Sugerencia de Colores HEX y Tipografías.
                8. Conclusión.
                
                IMPORTANTE: Limpia el texto de asteriscos. No uses negritas.
                """
                
                prompt_input.insert(0, prompt_ia)
                res = model.generate_content(prompt_input)
                texto_ia = res.text

                # --- GENERACIÓN PPTX ---
                prs = Presentation()
                
                for bloque in texto_ia.split("DIAPOSITIVA"):
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

                # Insertar imagen en portada si existe
                if f_foto:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
                        img.save(tmp_img.name)
                        prs.slides[0].shapes.add_picture(tmp_img.name, Inches(5.5), Inches(1.5), width=Inches(4))

                # Descarga
                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                
                st.success("✅ ¡Presentación Completa Generada!")
                st.download_button("📥 Descargar Presentación", buf, "Propuesta_Inversor.pptx")

        except Exception as e:
            st.error(f"Error técnico detectado: {e}")
