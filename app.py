import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import io
from docx import Document
import google.generativeai as genai

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Fábrica de Inversiones (Gemini)", layout="wide")
st.title("🚀 Presentaciones Comerciales (Powered by Gemini)")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input("Introduce tu Google API Key", type="password")
    tipo_negocio = st.selectbox("Sector del Proyecto", ["Reciclaje Inmobiliario", "Retail", "Gastronomía", "Tecnología", "Otro"])
    st.info("Esta versión usa Google Gemini (Gratuito).")

# --- FUNCIONES ---
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

# --- INTERFAZ ---
col1, col2 = st.columns(2)
with col1:
    st.subheader("📄 Documentos")
    doc_file = st.file_uploader("Subir Word o Texto", type=["docx", "txt"])
    excel_file = st.file_uploader("Subir Planilla Excel", type=["xlsx"])
with col2:
    st.subheader("✍️ Notas")
    notas = st.text_area("Notas manuales o transcripción:", height=200)

# --- BOTÓN DE ACCIÓN ---
if st.button("🔥 GENERAR PRESENTACIÓN COMPLETA"):
    if not api_key:
        st.error("Por favor, introduce tu Google API Key.")
    else:
        try:
            # Configurar Gemini
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-1.5-flash')

            with st.spinner("Gemini está analizando tu proyecto..."):
                
                # Unificar texto
                contexto = f"Proyecto: {tipo_negocio}\n"
                if notas: contexto += notas + "\n"
                if doc_file: contexto += extraer_texto_word(doc_file)
                if excel_file:
                    df = pd.read_excel(excel_file)
                    contexto += f"\nDatos Excel: {df.to_string()}"

                # Prompt
                prompt = f"""
                Actúa como experto en inversiones en Uruguay. Analiza este proyecto: {contexto}
                Genera 7 diapositivas para inversores. Formato:
                SLIDE | TÍTULO | CONTENIDO (en puntos clave)
                Incluye: Visión, Ubicación, Costos USD, KPIs (ROTE, Punto de Equilibrio), 
                Beneficios Fiscales Uruguay y Rol del Gestor.
                """

                response = model.generate_content(prompt)
                analisis = response.text

                # Crear PPTX
                prs = Presentation()
                for s_line in analisis.split("SLIDE"):
                    if "|" in s_line:
                        partes = s_line.split("|")
                        titulo = partes[1].strip()
                        cuerpo = partes[2].strip()
                        
                        slide = prs.slides.add_slide(prs.slide_layouts[1])
                        configurar_slide(slide, titulo)
                        tf = slide.placeholders[1].text_frame
                        for p in cuerpo.split("* "):
                            if len(p) > 2:
                                tf.add_paragraph().text = "• " + p.strip()

                # Descarga
                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                st.success("✅ ¡Presentación lista!")
                st.download_button("📥 Descargar PowerPoint", buf, "Propuesta_Gemini.pptx")
                
        except Exception as e:
            st.error(f"Error: {e}")
