import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import io
import tempfile
import os
from docx import Document
import google.generativeai as genai

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Presentaciones Pro Uruguay", layout="wide")
st.title("🚀 Generador de Proyectos de Inversión")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("Configuración de IA")
    api_key = st.text_input("Introduce tu Google API Key", type="password")
    tipo_negocio = st.selectbox("Sector", ["Reciclaje Inmobiliario", "Retail", "Gastronomía", "Tecnología", "Otro"])
    st.info("Sistema de auto-detección de modelo Gemini activado.")

# --- FUNCIONES DE LECTURA ---
def leer_texto(file):
    if file.name.endswith('.docx'):
        try:
            doc = Document(file)
            return "\n".join([p.text for p in doc.paragraphs])
        except: return "Error en Word."
    else:
        try: return file.read().decode("utf-8")
        except: return "Error en TXT."

def configurar_slide(slide, titulo_texto):
    try:
        title = slide.shapes.title
        title.text = titulo_texto
        title_para = title.text_frame.paragraphs[0]
        title_para.font.bold = True
        title_para.font.size = Pt(28)
        title_para.font.color.rgb = RGBColor(31, 119, 180)
    except: pass

# --- INTERFAZ ---
col1, col2 = st.columns(2)
with col1:
    st.subheader("📁 Archivos")
    doc_file = st.file_uploader("Documento (Word o TXT)", type=["docx", "txt"])
    excel_file = st.file_uploader("Excel Financiero", type=["xlsx"])
    media_file = st.file_uploader("Audio o Video", type=["mp3", "wav", "m4a", "mp4"])
with col2:
    st.subheader("📝 Notas")
    notas = st.text_area("Información adicional:", height=250)

# --- PROCESAMIENTO ---
if st.button("✨ GENERAR PRESENTACIÓN"):
    if not api_key:
        st.error("Falta la clave API.")
    else:
        try:
            genai.configure(api_key=api_key)
            
            # SISTEMA DE AUTO-DETECCIÓN DE MODELO
            model_names = ['gemini-1.5-flash', 'gemini-1.5-pro', 'gemini-pro']
            model = None
            
            with st.spinner("Buscando conexión con el cerebro de IA..."):
                for name in model_names:
                    try:
                        test_model = genai.GenerativeModel(name)
                        # Prueba rápida de conexión
                        test_model.generate_content("Hola", generation_config={"max_output_tokens": 10})
                        model = test_model
                        st.success(f"Conectado exitosamente usando el modelo: {name}")
                        break
                    except:
                        continue
                
                if model is None:
                    st.error("No se pudo conectar con ningún modelo de Gemini. Verifica tu API Key o los permisos en Google AI Studio.")
                    st.stop()

                datos_ia = []
                texto_contexto = f"PROYECTO: {tipo_negocio}\n"
                
                if doc_file: texto_contexto += f"DOC: {leer_texto(doc_file)}\n"
                if excel_file:
                    df = pd.read_excel(excel_file)
                    texto_contexto += f"EXCEL: {df.to_string()}\n"
                if notas: texto_contexto += f"NOTAS: {notas}\n"
                
                datos_ia.append(texto_contexto)

                tmp_path = None
                if media_file:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(media_file.name)[1]) as tmp:
                        tmp.write(media_file.read())
                        tmp_path = tmp.name
                    g_file = genai.upload_file(path=tmp_path)
                    datos_ia.append(g_file)

                prompt = """
                Eres un analista de inversiones experto. Crea una presentación de 7 diapositivas.
                Formato estricto: SLIDE | TÍTULO | CONTENIDO (puntos breves con *)
                Diapositivas: Portada, Resumen, Ubicación/Mercado, Inversión USD, KPIs (ROTE, ROI, Punto Equilibrio), Beneficios Fiscales Uruguay, Conclusión.
                """

                response = model.generate_content([prompt] + datos_ia)
                analisis = response.text

                # Crear PPTX
                prs = Presentation()
                # Limpiar el texto para procesar slides
                slides_raw = analisis.split("SLIDE")
                
                for bloque in slides_raw:
                    if "|" in bloque:
                        partes = bloque.split("|")
                        if len(partes) >= 3:
                            titulo = partes[1].strip().replace(":", "")
                            contenido = partes[2].strip()
                            
                            slide = prs.slides.add_slide(prs.slide_layouts[1])
                            configurar_slide(slide, titulo)
                            tf = slide.placeholders[1].text_frame
                            tf.word_wrap = True
                            
                            for p_text in contenido.split("*"):
                                clean_text = p_text.strip()
                                if len(clean_text) > 3:
                                    p = tf.add_paragraph()
                                    p.text = clean_text
                                    p.font.size = Pt(18)

                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                st.success("✅ ¡Presentación generada!")
                st.download_button("📥 Descargar PowerPoint", buf, "Propuesta_Inversion.pptx")
                
                if tmp_path: os.remove(tmp_path)

        except Exception as e:
            st.error(f"Error detallado: {e}")
