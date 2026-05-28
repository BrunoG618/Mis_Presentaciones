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

# --- ESTÉTICA ---
COLOR_TITULO = RGBColor(0, 32, 96)  # Azul Marino
COLOR_TEXTO = RGBColor(64, 64, 64)  # Gris Grafito

st.set_page_config(page_title="Generador Pro Inversiones", layout="wide")
st.title("🏆 Presentador Comercial Inteligente")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input("Introduce tu Google API Key", type="password")
    tipo_negocio = st.selectbox("Sector", ["Reciclaje Inmobiliario", "Retail", "Gastronomía", "Tecnología", "Otro"])

# --- FUNCIONES ---
def leer_archivo(file):
    if file.name.endswith('.docx'):
        doc = Document(file)
        return "\n".join([p.text for p in doc.paragraphs])
    return file.read().decode("utf-8")

def limpiar(t):
    """Limpia asteriscos, símbolos y basura del texto"""
    return t.replace("**", "").replace("#", "").replace("*", "").strip()

# --- INTERFAZ ---
col1, col2 = st.columns(2)
with col1:
    f_doc = st.file_uploader("Documento de texto (Word o TXT)", type=["docx", "txt"])
    f_xlsx = st.file_uploader("Excel Financiero", type=["xlsx"])
    f_media = st.file_uploader("Audio o Video de la charla", type=["mp3", "wav", "mp4", "m4a"])
with col2:
    f_notas = st.text_area("Escribe aquí o pega info extra:", height=250)

# --- PROCESAMIENTO ---
if st.button("🚀 GENERAR PRESENTACIÓN"):
    if not api_key:
        st.error("Falta la API Key en la barra lateral.")
    else:
        try:
            genai.configure(api_key=api_key)
            modelos = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            seleccionado = next((m for m in modelos if '1.5-flash' in m), modelos[0])
            model = genai.GenerativeModel(seleccionado)

            with st.spinner("Analizando y Diseñando diapositivas..."):
                # Recolectar datos
                contexto = f"PROYECTO: {tipo_negocio}\n"
                if f_notas: contexto += f"NOTAS: {f_notas}\n"
                if f_doc: contexto += f"DOC: {leer_archivo(f_doc)}\n"
                if f_xlsx:
                    df = pd.read_excel(f_xlsx)
                    contexto += f"EXCEL: {df.to_string()}\n"
                
                inputs_ia = [contexto]

                # Manejo de Multimedia
                if f_media:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f_media.name)[1]) as tmp:
                        tmp.write(f_media.read())
                        tmp_path = tmp.name
                    g_file = genai.upload_file(path=tmp_path)
                    while g_file.state.name == "PROCESSING":
                        time.sleep(2)
                        g_file = genai.get_file(g_file.name)
                    inputs_ia.append(g_file)

                # Prompt Simplificado (Más efectivo)
                prompt = """
                Eres un analista de negocios experto. Analiza la información y genera exactamente 7 diapositivas.
                Usa el separador '---' entre cada diapositiva.
                
                Estructura de cada diapositiva:
                Línea 1: El título de la diapositiva
                Líneas siguientes: El contenido en puntos cortos (usa un punto por línea)
                
                IMPORTANTE: No uses asteriscos, no uses negritas, no uses la palabra 'Slide'. 
                Solo el título y los puntos.
                
                Contenido requerido: 1. Portada, 2. Visión, 3. Ubicación y Mercado Montevideo, 
                4. Análisis Financiero USD, 5. KPIs (ROTE, Punto de Equilibrio), 
                6. Beneficios Fiscales Uruguay, 7. Cierre.
                """

                res = model.generate_content([prompt] + inputs_ia)
                texto_ia = res.text

                # --- CONSTRUCCIÓN DEL POWERPOINT ---
                prs = Presentation()
                
                # Separamos por el separador triple
                bloques = texto_ia.split("---")
                
                for bloque in bloques:
                    lineas = [l.strip() for l in bloque.split("\n") if l.strip()]
                    if len(lineas) >= 2:
                        # Crear Slide
                        slide = prs.slides.add_slide(prs.slide_layouts[1])
                        
                        # Título (Primera línea)
                        title_shape = slide.shapes.title
                        title_shape.text = limpiar(lineas[0])
                        title_para = title_shape.text_frame.paragraphs[0]
                        title_para.font.color.rgb = COLOR_TITULO
                        title_para.font.bold = True
                        title_para.font.size = Pt(32)

                        # Cuerpo (Resto de líneas)
                        body_shape = slide.placeholders[1]
                        tf = body_shape.text_frame
                        tf.word_wrap = True
                        
                        for line in lineas[1:]:
                            p = tf.add_paragraph()
                            p.text = "• " + limpiar(line)
                            p.font.size = Pt(18)
                            p.font.color.rgb = COLOR_TEXTO
                            p.space_after = Pt(10)

                # Guardar
                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                
                st.success("✅ ¡Presentación Generada con Éxito!")
                st.download_button("📥 DESCARGAR POWERPOINT PARA INVERSORES", buf, "Propuesta_Comercial.pptx")
                
                with st.expander("Ver Reporte de IA (Auditoría)"):
                    st.write(texto_ia)

        except Exception as e:
            st.error(f"Error técnico: {e}")
