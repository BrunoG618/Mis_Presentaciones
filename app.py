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
import time

# 1. CONFIGURACIÓN
st.set_page_config(page_title="Generador Pro Inversiones", layout="wide")
st.title("💼 Presentaciones Comerciales Inteligentes")

# 2. PANEL LATERAL
with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input("Introduce tu Google API Key", type="password")
    tipo_negocio = st.selectbox("Sector", ["Reciclaje Inmobiliario", "Retail", "Gastronomía", "Tecnología", "Otro"])

# 3. FUNCIONES DE LECTURA
def leer_documento(archivo):
    try:
        if archivo.name.endswith(".docx"):
            doc = Document(archivo)
            return "\n".join([p.text for p in doc.paragraphs])
        return archivo.read().decode("utf-8")
    except: return ""

# 4. INTERFAZ
col1, col2 = st.columns(2)
with col1:
    f_doc = st.file_uploader("Documento (Word o TXT)", type=["docx", "txt"])
    f_xlsx = st.file_uploader("Excel Financiero", type=["xlsx"])
    f_media = st.file_uploader("Audio o Video", type=["mp3", "wav", "mp4", "m4a"])
with col2:
    f_notas = st.text_area("Notas o información del proyecto:", height=250)

# 5. PROCESAMIENTO
if st.button("🚀 GENERAR PRESENTACIÓN"):
    if not api_key:
        st.error("Introduce tu API Key.")
    else:
        try:
            genai.configure(api_key=api_key)
            
            # Buscador de modelos
            modelos = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            seleccionado = next((m for m in modelos if '1.5-flash' in m), modelos[0])
            model = genai.GenerativeModel(seleccionado)

            with st.spinner("Analizando y redactando propuesta..."):
                contexto = f"PROYECTO: {tipo_negocio}\n"
                if f_notas: contexto += f"NOTAS: {f_notas}\n"
                if f_doc: contexto += f"DOC: {leer_documento(f_doc)}\n"
                if f_xlsx:
                    df = pd.read_excel(f_xlsx)
                    contexto += f"EXCEL: {df.to_string()}\n"
                
                inputs = [contexto]

                if f_media:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f_media.name)[1]) as tmp:
                        tmp.write(f_media.read())
                        tmp_path = tmp.name
                    g_file = genai.upload_file(path=tmp_path)
                    while g_file.state.name == "PROCESSING":
                        time.sleep(2)
                        g_file = genai.get_file(g_file.name)
                    inputs.append(g_file)

                # PROMPT REFORZADO
                prompt = """
                Analiza los datos y genera una presentación de inversión para Uruguay. 
                Separa cada diapositiva con la palabra 'DIAPOSITIVA:'. 
                Escribe un título corto y luego el contenido en puntos.
                
                Diapositivas: 
                1. Portada. 2. Visión. 3. Ubicación/Mercado. 4. Inversión (USD). 
                5. KPIs (ROTE, Punto Equilibrio). 6. Beneficios Fiscales Uruguay. 7. Conclusión.
                """

                res = model.generate_content([prompt] + inputs)
                texto_ia = res.text

                # --- GENERACIÓN DE POWERPOINT ROBUSTA ---
                prs = Presentation()
                
                # Intentamos separar por "DIAPOSITIVA"
                bloques = texto_ia.split("DIAPOSITIVA")
                
                if len(bloques) <= 1: # Si la IA no usó la palabra clave, usamos el texto completo
                    slide = prs.slides.add_slide(prs.slide_layouts[1])
                    slide.shapes.title.text = "Resumen del Proyecto"
                    tf = slide.placeholders[1].text_frame
                    tf.text = texto_ia
                else:
                    for bloque in bloques:
                        if len(bloque.strip()) > 10:
                            slide = prs.slides.add_slide(prs.slide_layouts[1])
                            
                            # Intentamos sacar la primera línea como título
                            lineas = bloque.strip().split("\n")
                            titulo = lineas[0].replace(":", "").strip()
                            contenido = "\n".join(lineas[1:])
                            
                            slide.shapes.title.text = titulo if len(titulo) < 50 else "Detalle del Proyecto"
                            tf = slide.placeholders[1].text_frame
                            tf.word_wrap = True
                            
                            # Limpiar y agregar párrafos
                            for p_text in contenido.split("\n"):
                                if p_text.strip():
                                    p = tf.add_paragraph()
                                    p.text = p_text.strip().replace("*", "•")
                                    p.font.size = Pt(16)

                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                
                st.success("✅ ¡Presentación generada con éxito!")
                st.download_button("📥 Descargar PowerPoint", buf, "Propuesta_Comercial.pptx")
                
                # CUADRO DE DEPURACIÓN (Para ver qué hizo la IA)
                with st.expander("Ver respuesta de la IA (por si el PPT falla)"):
                    st.write(texto_ia)

        except Exception as e:
            st.error(f"Error técnico: {e}")
