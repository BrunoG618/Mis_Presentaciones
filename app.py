import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
import io
import os
import tempfile
from docx import Document
import google.generativeai as genai
import time

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador Pro Uruguay", layout="wide")
st.title("💼 Generador de Proyectos de Inversión")

# --- SIDEBAR ---
with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input("Introduce tu Google API Key", type="password")
    tipo_negocio = st.selectbox("Sector", ["Reciclaje Inmobiliario", "Retail", "Gastronomía", "Tecnología", "Otro"])
    st.info("Buscando el mejor modelo disponible para tu cuenta...")

# --- FUNCIONES DE LECTURA ---
def leer_documento(archivo):
    try:
        if archivo.name.endswith(".docx"):
            doc = Document(archivo)
            return "\n".join([p.text for p in doc.paragraphs])
        else:
            return archivo.read().decode("utf-8")
    except: return "Error leyendo archivo."

# --- INTERFAZ ---
col1, col2 = st.columns(2)
with col1:
    st.subheader("📁 Archivos de Origen")
    f_doc = st.file_uploader("Conversación (Word o TXT)", type=["docx", "txt"])
    f_xlsx = st.file_uploader("Excel Financiero", type=["xlsx"])
    f_media = st.file_uploader("Audio o Video", type=["mp3", "wav", "mp4", "m4a"])
with col2:
    st.subheader("📝 Notas Extra")
    f_notas = st.text_area("Información adicional:", height=200)

# --- PROCESAMIENTO ---
if st.button("🚀 GENERAR PRESENTACIÓN"):
    if not api_key:
        st.error("Falta la API Key en la barra lateral.")
    else:
        try:
            genai.configure(api_key=api_key)
            
            # --- ESCÁNER DE MODELOS DISPONIBLES ---
            modelos_disponibles = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            
            # Prioridad de búsqueda
            seleccionado = None
            for opcion in ['models/gemini-1.5-flash', 'models/gemini-1.5-flash-latest', 'models/gemini-1.5-pro', 'models/gemini-pro']:
                if opcion in modelos_disponibles:
                    seleccionado = opcion
                    break
            
            if not seleccionado:
                seleccionado = modelos_disponibles[0] # Usar el primero que haya si no encuentra los preferidos

            st.success(f"Conectado con éxito al modelo: {seleccionado}")
            model = genai.GenerativeModel(seleccionado)

            with st.spinner("Analizando información financiera y legal..."):
                texto_total = f"PROYECTO: {tipo_negocio}\n"
                if f_doc: texto_total += f"DOCUMENTO: {leer_documento(f_doc)}\n"
                if f_notas: texto_total += f"NOTAS: {f_notas}\n"
                if f_xlsx:
                    df = pd.read_excel(f_xlsx)
                    texto_total += f"EXCEL: {df.to_string()}\n"
                
                partes_ia = [texto_total]

                # Multimedia
                if f_media:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f_media.name)[1]) as tmp:
                        tmp.write(f_media.read())
                        nombre_tmp = tmp.name
                    
                    archivo_g = genai.upload_file(path=nombre_tmp)
                    while archivo_g.state.name == "PROCESSING":
                        time.sleep(2)
                        archivo_g = genai.get_file(archivo_g.name)
                    partes_ia.append(archivo_g)

                prompt = """
                Eres un experto en inversiones en Uruguay. Analiza los datos y genera 7 diapositivas comerciales.
                Formato: SLIDE | TÍTULO | CONTENIDO
                Incluye: Introducción, Inversión (USD), KPIs (ROTE, Punto Equilibrio), Beneficios Fiscales (Vivienda Promovida) y Conclusión.
                Se muy profesional con los números.
                """

                res = model.generate_content([prompt] + partes_ia)
                
                # CREAR POWERPOINT
                prs = Presentation()
                for bloque in res.text.split("SLIDE"):
                    if "|" in bloque:
                        partes = bloque.split("|")
                        if len(partes) >= 3:
                            slide = prs.slides.add_slide(prs.slide_layouts[1])
                            slide.shapes.title.text = partes[1].strip()
                            tf = slide.placeholders[1].text_frame
                            for p in partes[2].strip().split("*"):
                                if len(p.strip()) > 2:
                                    tf.add_paragraph().text = "• " + p.strip()

                output = io.BytesIO()
                prs.save(output)
                output.seek(0)
                
                st.success("✅ ¡Presentación generada!")
                st.download_button("📥 Descargar PowerPoint", output, "Propuesta_Inversion.pptx")

        except Exception as e:
            st.error(f"Detalle técnico: {e}")
            st.info("Modelos detectados en tu cuenta: " + str(genai.list_models()))
