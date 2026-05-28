import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
import io
import os
import tempfile
from docx import Document
import google.generativeai as genai

# 1. CONFIGURACIÓN BÁSICA
st.set_page_config(page_title="Presentaciones Uruguay", layout="wide")
st.title("💼 Generador de Proyectos de Inversión")

# 2. PANEL LATERAL
with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input("Introduce tu Google API Key", type="password")
    tipo = st.selectbox("Sector", ["Reciclaje Inmobiliario", "Retail", "Gastronomía", "Otro"])

# 3. INTERFAZ DE CARGA
col1, col2 = st.columns(2)
with col1:
    st.subheader("📁 Archivos")
    f_doc = st.file_uploader("Documento (Word o TXT)", type=["docx", "txt"])
    f_xlsx = st.file_uploader("Excel Financiero", type=["xlsx"])
    f_media = st.file_uploader("Audio o Video", type=["mp3", "wav", "mp4", "m4a"])

with col2:
    st.subheader("📝 Notas")
    f_notas = st.text_area("Info extra:", height=200)

# 4. FUNCIÓN PARA LEER ARCHIVOS
def procesar_informacion():
    texto = f"Proyecto: {tipo}\n"
    if f_notas: texto += f"Notas: {f_notas}\n"
    if f_doc:
        try:
            if f_doc.name.endswith(".docx"):
                doc = Document(f_doc)
                texto += "\n".join([p.text for p in doc.paragraphs])
            else:
                texto += f_doc.read().decode("utf-8")
        except: texto += "Error leyendo documento.\n"
    if f_xlsx:
        try:
            df = pd.read_excel(f_xlsx)
            texto += f"\nDatos Excel: {df.to_string()}"
        except: texto += "Error leyendo Excel.\n"
    return texto

# 5. GENERACIÓN
if st.button("🚀 GENERAR PRESENTACIÓN"):
    if not api_key:
        st.error("Falta la API Key")
    else:
        try:
            genai.configure(api_key=api_key)
            # Usamos el nombre de modelo más estándar
            model = genai.GenerativeModel('gemini-1.5-flash')
            
            with st.spinner("Analizando y diseñando..."):
                input_ia = [procesar_informacion()]
                
                # Manejo de audio/video
                if f_media:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f_media.name)[1]) as tmp:
                        tmp.write(f_media.read())
                        nombre_tmp = tmp.name
                    f_google = genai.upload_file(path=nombre_tmp)
                    input_ia.append(f_google)

                prompt = """
                Analiza la info y genera 7 diapositivas para una presentación comercial en Uruguay.
                Formato obligatorio: SLIDE | TÍTULO | CONTENIDO (en puntos clave)
                Incluye: Introducción, Inversión (USD), KPIs (ROTE, Punto de Equilibrio), Beneficios Fiscales y Conclusión.
                """

                res = model.generate_content([prompt] + input_ia)
                
                # Crear el PowerPoint
                prs = Presentation()
                for bloque in res.text.split("SLIDE"):
                    if "|" in bloque:
                        partes = bloque.split("|")
                        if len(partes) >= 3:
                            slide = prs.slides.add_slide(prs.slide_layouts[1])
                            slide.shapes.title.text = partes[1].strip()
                            tf = slide.placeholders[1].text_frame
                            for p in partes[2].strip().split("*"):
                                if len(p.strip()) > 3:
                                    tf.add_paragraph().text = "• " + p.strip()

                output = io.BytesIO()
                prs.save(output)
                output.seek(0)
                
                st.success("✅ ¡Lista!")
                st.download_button("📥 Descargar PowerPoint", output, "Proyecto.pptx")

        except Exception as e:
            st.error(f"Error técnico: {e}")
