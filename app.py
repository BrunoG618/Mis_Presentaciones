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

# 1. CONFIGURACIÓN DE PÁGINA
st.set_page_config(page_title="Generador Pro Inversiones", layout="wide")
st.title("💼 Presentaciones Comerciales Inteligentes")

# 2. PANEL LATERAL
with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input("Introduce tu Google API Key", type="password")
    st.info("Obtén tu clave gratis en: aistudio.google.com")
    tipo_negocio = st.selectbox("Sector", ["Reciclaje Inmobiliario", "Retail", "Gastronomía", "Tecnología", "Otro"])

# 3. FUNCIONES DE LECTURA
def leer_documento(archivo):
    try:
        if archivo.name.endswith(".docx"):
            doc = Document(archivo)
            return "\n".join([p.text for p in doc.paragraphs])
        return archivo.read().decode("utf-8")
    except: return ""

def configurar_slide(slide, titulo_texto):
    title = slide.shapes.title
    title.text = titulo_texto
    title_para = title.text_frame.paragraphs[0]
    title_para.font.bold = True
    title_para.font.size = Pt(28)
    title_para.font.color.rgb = RGBColor(31, 119, 180)

# 4. INTERFAZ DE CARGA
col1, col2 = st.columns(2)
with col1:
    f_doc = st.file_uploader("Documento (Word o TXT)", type=["docx", "txt"])
    f_xlsx = st.file_uploader("Excel Financiero", type=["xlsx"])
    f_media = st.file_uploader("Audio o Video", type=["mp3", "wav", "mp4", "m4a"])
with col2:
    f_notas = st.text_area("Información adicional o notas:", height=250)

# 5. PROCESAMIENTO
if st.button("🚀 GENERAR PRESENTACIÓN PROFESIONAL"):
    if not api_key:
        st.error("Por favor, introduce tu API Key.")
    else:
        try:
            genai.configure(api_key=api_key)
            
            # --- SOLUCIÓN DEFINITIVA AL 404: DETECCIÓN DINÁMICA ---
            available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            
            # Priorizamos los modelos 1.5 por su capacidad de ver video/audio
            seleccionado = None
            for target in ['models/gemini-1.5-flash', 'models/gemini-1.5-pro', 'models/gemini-pro']:
                if target in available_models:
                    seleccionado = target
                    break
            
            if not seleccionado:
                seleccionado = available_models[0] # Fallback al primero disponible

            model = genai.GenerativeModel(seleccionado)
            st.success(f"Conectado con éxito mediante: {seleccionado}")

            with st.spinner("Analizando datos y diseñando propuesta..."):
                # Recolectar info
                contexto = f"PROYECTO: {tipo_negocio}\n"
                if f_notas: contexto += f"NOTAS: {f_notas}\n"
                if f_doc: contexto += f"DOC: {leer_documento(f_doc)}\n"
                if f_xlsx:
                    df = pd.read_excel(f_xlsx)
                    contexto += f"EXCEL: {df.to_string()}\n"
                
                inputs = [contexto]

                # Manejo de Multimedia
                if f_media:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f_media.name)[1]) as tmp:
                        tmp.write(f_media.read())
                        tmp_path = tmp.name
                    
                    g_file = genai.upload_file(path=tmp_path)
                    while g_file.state.name == "PROCESSING":
                        time.sleep(2)
                        g_file = genai.get_file(g_file.name)
                    inputs.append(g_file)

                # Prompt para la IA
                prompt = """
                Actúa como un experto en inversiones en Uruguay. Analiza los datos y genera 7 diapositivas comerciales excelentes.
                Formato de respuesta estricto:
                SLIDE | TÍTULO | CONTENIDO (en puntos clave con *)
                
                Diapositivas requeridas:
                1. Portada y Visión. 2. Ubicación y Mercado. 3. Análisis de Inversión (USD). 
                4. KPIs: ROTE, ROI y Punto de Equilibrio. 5. Beneficios Fiscales (Vivienda Promovida). 
                6. Rol del Gestor y Socio Administrador. 7. Conclusión.
                """

                res = model.generate_content([prompt] + inputs)
                
                # Crear PowerPoint
                prs = Presentation()
                for bloque in res.text.split("SLIDE"):
                    if "|" in bloque:
                        partes = bloque.split("|")
                        if len(partes) >= 3:
                            slide = prs.slides.add_slide(prs.slide_layouts[1])
                            configurar_slide(slide, partes[1].strip())
                            tf = slide.placeholders[1].text_frame
                            for p in partes[2].strip().split("*"):
                                if len(p.strip()) > 3:
                                    tf.add_paragraph().text = "• " + p.strip()

                # Descarga
                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                st.success("✅ ¡Presentación generada con éxito!")
                st.download_button("📥 Descargar PowerPoint Profesional", buf, "Propuesta_Inversion.pptx")

        except Exception as e:
            st.error(f"Error técnico: {e}")
