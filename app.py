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
st.title("🚀 Generador de Presentaciones Multimodal")
st.write("Sube cualquier formato (Video, Audio, Excel, Word) y genera tu propuesta comercial.")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("Configuración de IA")
    api_key = st.text_input("Introduce tu Google API Key (Gratis)", type="password")
    st.markdown("[Obtén tu clave gratis aquí](https://aistudio.google.com/)")
    tipo_negocio = st.selectbox("Sector", ["Reciclaje Inmobiliario", "Retail", "Gastronomía", "Tecnología", "Otro"])
    st.divider()
    st.info("Esta versión usa Gemini 1.5 para analizar videos, audios y datos financieros.")

# --- FUNCIONES DE APOYO ---
def extraer_texto_word(file):
    doc = Document(file)
    return "\n".join([p.text for p in doc.paragraphs])

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
    doc_file = st.file_uploader("Word o TXT", type=["docx", "txt"])
    excel_file = st.file_uploader("Planilla Excel", type=["xlsx"])

with col2:
    st.subheader("🎬 Multimedia")
    media_file = st.file_uploader("Audio o Video de la charla", type=["mp3", "wav", "m4a", "mp4", "mov"])

with col3:
    st.subheader("✍️ Notas")
    notas = st.text_area("Notas manuales:", height=200, placeholder="Pega aquí info extra...")

# --- PROCESAMIENTO ---
if st.button("🔥 GENERAR PROPUESTA INTEGRAL"):
    if not api_key:
        st.error("Introduce tu Google API Key en la barra lateral.")
    else:
        try:
            genai.configure(api_key=api_key)
            # Usamos el modelo Flash que es excelente con archivos pesados
            model = genai.GenerativeModel('gemini-1.5-flash')
            
            with st.spinner("Analizando todos tus archivos (esto puede tardar según el tamaño del video/audio)..."):
                
                contenido_para_ia = []
                
                # 1. Procesar Texto y Notas
                texto_base = f"Proyecto: {tipo_negocio}\nNotas: {notas}\n"
                if doc_file:
                    texto_base += f"Contenido Documento: {extraer_texto_word(doc_file)}\n"
                if excel_file:
                    df = pd.read_excel(excel_file)
                    texto_base += f"Datos Excel: {df.to_string()}\n"
                
                contenido_para_ia.append(texto_base)

                # 2. Procesar Multimedia (Audio/Video)
                if media_file:
                    # Guardar temporalmente para que Gemini lo lea
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(media_file.name)[1]) as tmp:
                        tmp.write(media_file.read())
                        tmp_path = tmp.name
                    
                    # Subir a Google para procesamiento
                    st.info("Subiendo archivo multimedia a la IA...")
                    g_file = genai.upload_file(path=tmp_path)
                    contenido_para_ia.append(g_file)

                # 3. Prompt Maestro
                prompt = f"""
                Analiza absolutamente toda la información adjunta (videos, audios, excel y texto).
                Se trata de un proyecto de {tipo_negocio} en Uruguay.
                
                Genera el contenido detallado para una presentación comercial de 8 diapositivas.
                Formato de respuesta: SLIDE | TÍTULO | CONTENIDO (en puntos)
                
                Debes incluir:
                - Resumen de la conversación/video.
                - Análisis financiero detallado (Inversión en USD).
                - KPIs calculados: ROTE, ROI, Punto de Equilibrio y Cost to Income.
                - Beneficios fiscales en Uruguay (Vivienda Promovida / COMAP).
                - Escenarios de rentabilidad y conclusión.
                """

                # 4. Generar Respuesta
                response = model.generate_content([prompt] + contenido_para_ia)
                analisis = response.text

                # 5. Crear PowerPoint
                prs = Presentation()
                for s_line in analisis.split("SLIDE"):
                    if "|" in s_line:
                        partes = s_line.split("|")
                        titulo = partes[1].strip()
                        cuerpo = partes[2].strip()
                        
                        slide = prs.slides.add_slide(prs.slide_layouts[1])
                        configurar_slide(slide, titulo)
                        tf = slide.placeholders[1].text_frame
                        tf.word_wrap = True
                        for p in cuerpo.split("* "):
                            if len(p) > 2:
                                tf.add_paragraph().text = "• " + p.strip()

                # 6. Descarga
                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                st.success("✅ ¡Análisis Multimodal Completo!")
                st.download_button("📥 Descargar Presentación Pro", buf, "Propuesta_Inversor_Final.pptx")
                
                with st.expander("Ver Reporte Técnico"):
                    st.write(analisis)
                
                # Limpiar archivo temporal
                if media_file: os.remove(tmp_path)

        except Exception as e:
            st.error(f"Se produjo un error: {e}")
