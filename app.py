import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import io, os, tempfile, re, time
from docx import Document
import google.generativeai as genai
from PIL import Image

# --- ESTÉTICA PREMIUM ---
COLOR_BG = RGBColor(15, 15, 15)       # Negro Grafito
COLOR_GOLD = RGBColor(197, 160, 82)   # Dorado Champagne
COLOR_TEXT = RGBColor(230, 230, 230)  # Blanco Humo

st.set_page_config(page_title="Inversor Pro - Uruguay", layout="wide")
st.title("🏛️ Generador de Propuestas: Inversión Inmobiliaria")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input("Introduce tu Google API Key", type="password")
    st.info("Esta versión detecta automáticamente el modelo disponible en tu cuenta.")

# --- FUNCIONES DE LECTURA ---
def leer_documento_seguro(file):
    nombre = file.name.lower()
    try:
        if nombre.endswith('.docx'):
            doc = Document(file)
            return "\n".join([p.text for p in doc.paragraphs])
        else:
            return file.read().decode("utf-8")
    except Exception as e:
        return f"Error leyendo {nombre}: {e}"

def limpiar_texto(t):
    """Limpia asteriscos, numerales y etiquetas de control"""
    t = t.replace("**", "").replace("#", "").replace("`", "").strip()
    return re.sub(r'^(TITULO|TEXTO|DIAPOSITIVA|SLIDE|CUERPO|CONTENIDO):?\s*', '', t, flags=re.IGNORECASE)

def crear_slide_maquetado(prs, titulo, contenido, imagen=None):
    """Crea una diapositiva con diseño premium y manejo de imagen"""
    slide = prs.slides.add_slide(prs.slide_layouts[1]) # Layout de Titulo y Objetos
    
    # Fondo
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = COLOR_BG
    
    # Titulo
    title_shape = slide.shapes.title
    title_shape.text = limpiar_texto(titulo).upper()
    title_para = title_shape.text_frame.paragraphs[0]
    title_para.font.size = Pt(28)
    title_para.font.bold = True
    title_para.font.color.rgb = COLOR_GOLD
    title_para.alignment = PP_ALIGN.LEFT

    # Cuerpo de texto
    body_shape = slide.placeholders[1]
    # Si hay imagen, achicamos el texto a la mitad
    if imagen:
        body_shape.width = Inches(5)
    
    tf = body_shape.text_frame
    tf.word_wrap = True
    
    # Procesar contenido por lineas
    for linea in contenido.split('\n'):
        linea_limpia = limpiar_texto(linea)
        if len(linea_limpia) > 3:
            p = tf.add_paragraph()
            p.text = "• " + linea_limpia
            p.font.size = Pt(16)
            p.font.color.rgb = COLOR_TEXT
            p.space_after = Pt(10)

    # Insertar Imagen si existe
    if imagen:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
            imagen.save(tmp_img.name)
            slide.shapes.add_picture(tmp_img.name, Inches(5.5), Inches(1.5), height=Inches(4.5))

# --- INTERFAZ ---
st.subheader("📸 Carga de Datos del Proyecto")
c1, c2, c3 = st.columns(3)
with c1:
    f_foto = st.file_uploader("Fachada o Render", type=["jpg", "png", "jpeg"])
    f_xlsx = st.file_uploader("Excel Financiero", type=["xlsx"])
with c2:
    f_doc = st.file_uploader("Documento (Word o TXT)", type=["docx", "txt"])
    f_notas = st.text_area("Notas de la reforma:", height=100)
with c3:
    f_media = st.file_uploader("Audio o Video", type=["mp3", "mp4", "wav"])

# --- PROCESAMIENTO ---
if st.button("🏗️ GENERAR PRESENTACIÓN"):
    if not api_key:
        st.error("Falta la API Key.")
    else:
        try:
            genai.configure(api_key=api_key)
            # Detección de modelo
            available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            model_name = next((m for m in available_models if '1.5-flash' in m), available_models[0])
            model = genai.GenerativeModel(model_name)
            
            with st.spinner("Analizando información y diseñando..."):
                # Recopilar contexto
                ctx = "PROYECTO: Reciclaje Inmobiliario Montevideo\n"
                if f_doc: ctx += f"TEXTO: {leer_documento_seguro(f_doc)}\n"
                if f_notas: ctx += f"NOTAS: {f_notas}\n"
                if f_xlsx:
                    df = pd.read_excel(f_xlsx)
                    ctx += f"EXCEL: {df.to_string()}\n"

                parts = [ctx]
                if f_foto:
                    img = Image.open(f_foto)
                    parts.append(img)
                if f_media:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f_media.name)[1]) as tmp:
                        tmp.write(f_media.read())
                        t_path = tmp.name
                    gf = genai.upload_file(path=t_path)
                    while gf.state.name == "PROCESSING":
                        time.sleep(2)
                        gf = genai.get_file(gf.name)
                    parts.append(gf)

                # Prompt Simplificado pero Directo
                prompt = """
                Analiza todo el material. Genera el contenido para 7 diapositivas de inversión.
                Para CADA diapositiva, escribe:
                SLIDE: [Título de la diapositiva]
                CONTENIDO:
                [Punto 1]
                [Punto 2]
                [Punto 3]
                
                Contenido: 1. Portada, 2. Visión Arquitectónica, 3. Ubicación, 4. Inversión USD, 
                5. KPIs (ROTE, ROI, Punto Equilibrio), 6. Beneficios Fiscales Uruguay, 7. Cierre.
                """
                
                res = model.generate_content([prompt] + parts)
                raw_text = res.text

                # --- CONSTRUCCIÓN PPTX ---
                prs = Presentation()
                
                # Dividir por la palabra SLIDE
                bloques = re.split(r'SLIDE:', raw_text, flags=re.IGNORECASE)
                
                # Si no hay bloques, meter todo el texto en una sola hoja (Seguridad)
                if len(bloques) <= 1:
                    crear_slide_maquetado(prs, "Resumen del Proyecto", raw_text, f_foto if f_foto else None)
                else:
                    for i, bloque in enumerate(bloques):
                        if len(bloque.strip()) > 10:
                            # Separar Título de Contenido
                            secciones = re.split(r'CONTENIDO:', bloque, flags=re.IGNORECASE)
                            titulo = secciones[0].strip()
                            contenido = secciones[1].strip() if len(secciones) > 1 else ""
                            
                            # Solo poner la foto en la Diapositiva 1 o 2
                            foto_a_usar = img if (i <= 2 and f_foto) else None
                            crear_slide_maquetado(prs, titulo, contenido, foto_a_usar)

                # Descarga
                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                st.success("✅ ¡Presentación Generada!")
                st.download_button("📥 DESCARGAR PPT PROFESIONAL", buf, "Propuesta_Final.pptx")
                
                if f_media: os.remove(t_path)
                
                with st.expander("Ver Reporte de IA (Auditoría)"):
                    st.write(raw_text)

        except Exception as e:
            st.error(f"Error técnico: {e}")
