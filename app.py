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
st.set_page_config(page_title="Inversiones Inteligentes - Uruguay", layout="wide")
st.title("🚀 Generador de Presentaciones Pro")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("Configuración de IA")
    api_key = st.text_input("Introduce tu Google API Key", type="password")
    st.info("Obtén tu clave gratis en Google AI Studio.")
    tipo_negocio = st.selectbox("Sector", ["Reciclaje Inmobiliario", "Retail", "Gastronomía", "Otro"])

# --- FUNCIONES DE LECTURA SEGURA ---
def leer_archivo_texto(file):
    """Lee archivos .docx o .txt de forma segura"""
    if file.name.endswith('.docx'):
        try:
            doc = Document(file)
            return "\n".join([p.text for p in doc.paragraphs])
        except Exception:
            return "Error leyendo Word."
    else:
        try:
            return file.read().decode("utf-8")
        except Exception:
            return "Error leyendo TXT."

def configurar_slide(slide, titulo_texto):
    title = slide.shapes.title
    title.text = titulo_texto
    title_para = title.text_frame.paragraphs[0]
    title_para.font.bold = True
    title_para.font.size = Pt(30)
    title_para.font.color.rgb = RGBColor(31, 119, 180)

# --- INTERFAZ DE CARGA ---
col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("📄 Documentos")
    doc_file = st.file_uploader("Subir Conversación (Word o TXT)", type=["docx", "txt"])
    excel_file = st.file_uploader("Subir Planilla Excel", type=["xlsx"])

with col2:
    st.subheader("🎬 Multimedia")
    media_file = st.file_uploader("Audio o Video", type=["mp3", "wav", "m4a", "mp4"])

with col3:
    st.subheader("✍️ Notas")
    notas = st.text_area("Notas manuales:", height=200)

# --- PROCESAMIENTO ---
if st.button("🔥 GENERAR PROPUESTA"):
    if not api_key:
        st.error("Introduce tu Google API Key.")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-1.5-flash')
            
            with st.spinner("Analizando información..."):
                contenido_ia = []
                texto_base = f"Proyecto: {tipo_negocio}\n"
                
                # 1. Procesar Documentos (Lógica Corregida)
                if doc_file:
                    texto_base += f"Contenido Documento: {leer_archivo_texto(doc_file)}\n"
                
                if notas:
                    texto_base += f"Notas: {notas}\n"

                # 2. Procesar Excel
                if excel_file:
                    df = pd.read_excel(excel_file)
                    texto_base += f"Datos Financieros Excel:\n{df.to_string()}\n"
                
                contenido_ia.append(texto_base)

                # 3. Procesar Multimedia
                tmp_path = None
                if media_file:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(media_file.name)[1]) as tmp:
                        tmp.write(media_file.read())
                        tmp_path = tmp.name
                    g_file = genai.upload_file(path=tmp_path)
                    contenido_ia.append(g_file)

                # 4. Prompt
                prompt = f"""
                Analiza la información de este proyecto de {tipo_negocio} en Uruguay.
                Crea una presentación de 7 diapositivas profesional.
                Formato: SLIDE | TÍTULO | CONTENIDO
                Incluye: Resumen, Ubicación, Análisis de Inversión USD, KPIs (ROTE, ROI, Punto Equilibrio), 
                Beneficios Fiscales Uruguay (Vivienda Promovida) y Conclusión.
                """

                # 5. Generar
                response = model.generate_content([prompt] + contenido_ia)
                analisis = response.text

                # 6. Crear PPTX
                prs = Presentation()
                for s_line in analisis.split("SLIDE"):
                    if "|" in s_line:
                        partes = s_line.split("|")
                        if len(partes) >= 3:
                            titulo = partes[1].strip()
                            cuerpo = partes[2].strip()
                            
                            slide = prs.slides.add_slide(prs.slide_layouts[1])
                            configurar_slide(slide, titulo)
                            tf = slide.placeholders[1].text_frame
                            tf.word_wrap = True
                            for p in cuerpo.split("* "):
                                if len(p) > 2:
                                    tf.add_paragraph().text = "• " + p.strip()

                # 7. Descarga
                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                st.success("✅ ¡Presentación Lista!")
                st.download_button("📥 Descargar PowerPoint", buf, "Propuesta_Inversor.pptx")
                
                if tmp_path: os.remove(tmp_path)

        except Exception as e:
            st.error(f"Se produjo un error: {e}")
