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

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Inversor Pro", layout="wide")
st.title("💼 Generador de Presentaciones Comerciales")

# --- SIDEBAR ---
with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input("Introduce tu Google API Key", type="password")
    tipo_negocio = st.selectbox("Sector", ["Reciclaje Inmobiliario", "Retail", "Gastronomía", "Tecnología", "Otro"])

# --- FUNCIONES ---
def leer_documento(archivo):
    try:
        if archivo.name.endswith(".docx"):
            doc = Document(archivo)
            return "\n".join([p.text for p in doc.paragraphs])
        else:
            return archivo.read().decode("utf-8")
    except:
        return "Error al leer el documento."

# --- INTERFAZ ---
col1, col2 = st.columns(2)
with col1:
    st.subheader("📁 Carga de Datos")
    f_doc = st.file_uploader("Conversación (Word o TXT)", type=["docx", "txt"])
    f_xlsx = st.file_uploader("Excel Financiero", type=["xlsx"])
    f_media = st.file_uploader("Audio o Video", type=["mp3", "wav", "mp4", "m4a"])

with col2:
    st.subheader("📝 Notas")
    f_notas = st.text_area("Información extra:", height=200)

# --- PROCESAMIENTO ---
if st.button("🚀 GENERAR PRESENTACIÓN"):
    if not api_key:
        st.error("Introduce la clave API en la barra lateral.")
    else:
        try:
            genai.configure(api_key=api_key)
            
            # Intentar conectar con el modelo más estable
            model = genai.GenerativeModel('gemini-1.5-flash')
            
            with st.spinner("Analizando información..."):
                texto_total = f"PROYECTO: {tipo_negocio}\n"
                
                if f_doc: texto_total += f"DOCUMENTO: {leer_documento(f_doc)}\n"
                if f_notas: texto_total += f"NOTAS: {f_notas}\n"
                if f_xlsx:
                    df = pd.read_excel(f_xlsx)
                    texto_total += f"DATOS EXCEL: {df.to_string()}\n"
                
                partes_ia = [texto_total]

                # Manejo de Multimedia (Audio/Video)
                if f_media:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f_media.name)[1]) as tmp:
                        tmp.write(f_media.read())
                        nombre_tmp = tmp.name
                    
                    st.info("Procesando archivo multimedia...")
                    archivo_g = genai.upload_file(path=nombre_tmp)
                    
                    # Esperar a que el archivo esté listo en los servidores de Google
                    while archivo_g.state.name == "PROCESSING":
                        time.sleep(2)
                        archivo_g = genai.get_file(archivo_g.name)
                    
                    partes_ia.append(archivo_g)

                # El Pedido a la IA
                prompt = """
                Analiza la información y genera 7 diapositivas para una presentación comercial en Uruguay.
                Formato: SLIDE | TÍTULO | CONTENIDO
                Incluye: Introducción, Inversión (USD), KPIs (ROTE, Punto de Equilibrio), Beneficios Fiscales (Vivienda Promovida) y Conclusión.
                """

                res = model.generate_content([prompt] + partes_ia)
                
                # Crear PowerPoint
                prs = Presentation()
                for bloque in res.text.split("SLIDE"):
                    if "|" in bloque:
                        partes = bloque.split("|")
                        if len(partes) >= 3:
                            titulo_s = partes[1].strip()
                            contenido_s = partes[2].strip()
                            
                            slide = prs.slides.add_slide(prs.slide_layouts[1])
                            slide.shapes.title.text = titulo_s
                            tf = slide.placeholders[1].text_frame
                            for p in contenido_s.split("*"):
                                if len(p.strip()) > 3:
                                    tf.add_paragraph().text = "• " + p.strip()

                # Descarga
                output = io.BytesIO()
                prs.save(output)
                output.seek(0)
                
                st.success("✅ ¡Presentación lista!")
                st.download_button("📥 Descargar PowerPoint", output, "Proyecto_Inversion.pptx")

        except Exception as e:
            st.error(f"Ocurrió un detalle técnico: {e}")
