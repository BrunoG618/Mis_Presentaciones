import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import io
import os
import tempfile
import time
from docx import Document
import google.generativeai as genai
from PIL import Image

# --- CONFIGURACIÓN DE ESTILO ---
AZUL_URUGUAY = RGBColor(0, 51, 102)
GRIS_TEXTO = RGBColor(64, 64, 64)
NARANJA_OBRA = RGBColor(255, 102, 0)

st.set_page_config(page_title="Presentaciones Pro Uruguay", layout="wide")
st.title("🏆 Generador de Proyectos de Inversión")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("Configuración de IA")
    api_key = st.text_input("Introduce tu Google API Key", type="password")
    tipo_negocio = st.selectbox("Sector", ["Reciclaje Inmobiliario", "Inversión Comercial", "Retail", "Gastronomía", "Otro"])
    st.info("Esta versión procesa: Texto, Excel, Fotos y Video/Audio.")

# --- FUNCIONES DE LECTURA BLINDADAS (ELIMINA EL ERROR 'NOT A ZIP') ---
def procesar_archivo_texto(file):
    if file.name.lower().endswith('.docx'):
        try:
            doc = Document(file)
            return "\n".join([p.text for p in doc.paragraphs])
        except Exception as e:
            return f"\n[Error leyendo Word: {e}]\n"
    elif file.name.lower().endswith('.txt'):
        try:
            return file.read().decode("utf-8")
        except Exception as e:
            return f"\n[Error leyendo TXT: {e}]\n"
    return ""

def procesar_excel_seguro(file):
    try:
        df = pd.read_excel(file)
        return f"\nDATOS FINANCIEROS (EXCEL):\n{df.to_string()}\n"
    except Exception as e:
        return f"\n[Error leyendo Excel: {e}]\n"

def aplicar_diseno(slide, titulo_texto):
    """Diseño profesional para inversores"""
    title = slide.shapes.title
    title.text = titulo_texto
    p = title.text_frame.paragraphs[0]
    p.font.bold = True
    p.font.size = Pt(30)
    p.font.color.rgb = AZUL_URUGUAY
    
    # Línea decorativa naranja
    line = slide.shapes.add_shape(1, Inches(0.5), Inches(1.15), Inches(4), Inches(0.05))
    line.fill.solid()
    line.fill.foreground_color.rgb = NARANJA_OBRA
    line.line.fill.background()

# --- INTERFAZ DE CARGA MULTIMODAL ---
st.subheader("📁 Carga de Información (Formatos soportados: Word, TXT, Excel, JPG, MP4, MP3)")
col1, col2, col3 = st.columns(3)

with col1:
    f_texto = st.file_uploader("Documento de Texto o Word", type=["docx", "txt"])
    f_excel = st.file_uploader("Planilla Financiera Excel", type=["xlsx"])

with col2:
    f_multimedia = st.file_uploader("Audio o Video de la charla", type=["mp3", "wav", "mp4", "m4a"])
    f_notas = st.text_area("Notas manuales o transcripción:", height=100)

with col3:
    f_imagen = st.file_uploader("Imagen de Fachada o Render", type=["jpg", "png", "jpeg"])

# --- PROCESO PRINCIPAL ---
if st.button("🚀 GENERAR PRESENTACIÓN EJECUTIVA"):
    if not api_key:
        st.error("Por favor, introduce tu API Key de Google.")
    else:
        try:
            genai.configure(api_key=api_key)
            # Auto-selección de modelo para evitar error 404
            model = genai.GenerativeModel('gemini-1.5-flash')
            
            with st.spinner("Analizando información multimodal y diseñando diapositivas..."):
                
                entradas_ia = []
                # 1. Recopilar todo el texto de forma segura
                contexto_texto = f"PROYECTO: {tipo_negocio}\n"
                if f_texto: contexto_texto += f"CONTENIDO DOC: {procesar_archivo_texto(f_texto)}\n"
                if f_excel: contexto_texto += procesar_excel_seguro(f_excel)
                if f_notas: contexto_texto += f"NOTAS EXTRAS: {f_notas}\n"
                
                entradas_ia.append(contexto_texto)

                # 2. Procesar Imagen
                if f_imagen:
                    img = Image.open(f_imagen)
                    entradas_ia.append(img)

                # 3. Procesar Multimedia
                if f_multimedia:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f_multimedia.name)[1]) as tmp:
                        tmp.write(f_multimedia.read())
                        tmp_path = tmp.name
                    
                    g_file = genai.upload_file(path=tmp_path)
                    while g_file.state.name == "PROCESSING":
                        time.sleep(2)
                        g_file = genai.get_file(g_file.name)
                    entradas_ia.append(g_file)

                # 4. Prompt Maestro (KPIs y Requerimientos Uruguay)
                prompt = """
                Analiza toda la información (texto, excel, imagen y multimedia). 
                Crea una presentación comercial de 9 diapositivas para inversores.
                Usa el separador '#### SLIDE ####' entre diapositivas.
                
                Formato de respuesta:
                #### SLIDE ####
                TÍTULO: [Título de la diapositiva]
                CONTENIDO: [Puntos clave separados por asteriscos *]

                REQUERIMIENTOS:
                - Incluye KPIs: ROTE, ROI, Cost to Income y Punto de Equilibrio.
                - Separa flujos en Moneda Nacional (UYU) y Dólares (USD).
                - Menciona Beneficios Fiscales: Vivienda Promovida (Ley 18.795).
                - Diapositivas: Portada, Visión, Ubicación, El Negocio, Inversión y Gastos, Beneficios Fiscales, Rentabilidad (KPIs), Diseño Sugerido (colores y tipos), Cierre.
                """

                res = model.generate_content([prompt] + entradas_ia)
                texto_ia = res.text

                # 5. CONSTRUCCIÓN DEL PPTX
                prs = Presentation()
                bloques = texto_ia.split("#### SLIDE ####")
                
                for bloque in bloques:
                    if "TÍTULO:" in bloque and "CONTENIDO:" in bloque:
                        try:
                            # Extraer datos
                            titulo_s = bloque.split("CONTENIDO:")[0].replace("TÍTULO:", "").strip()
                            contenido_s = bloque.split("CONTENIDO:")[1].strip()
                            
                            slide = prs.slides.add_slide(prs.slide_layouts[1])
                            aplicar_diseno(slide, titulo_s.replace("**", ""))
                            
                            tf = slide.placeholders[1].text_frame
                            tf.word_wrap = True
                            for punto in contenido_s.split("*"):
                                if len(punto.strip()) > 3:
                                    p = tf.add_paragraph()
                                    p.text = punto.strip().replace("**", "")
                                    p.font.size = Pt(17)
                                    p.font.color.rgb = GRIS_TEXTO
                        except: continue

                # Inserción de imagen en portada si existe
                if f_imagen:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
                        img.save(tmp_img.name)
                        prs.slides[0].shapes.add_picture(tmp_img.name, Inches(6.2), Inches(1.5), width=Inches(3.2))

                # 6. Descarga
                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                
                st.success("✅ ¡Presentación Completa Generada!")
                st.download_button("📥 Descargar Propuesta Comercial", buf, "Propuesta_Inversor_Uruguay.pptx")
                
                if f_multimedia: os.remove(tmp_path)

        except Exception as e:
            st.error(f"Error técnico detectado: {e}")
