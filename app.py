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
COLOR_GOLD = RGBColor(197, 160, 82)   # Dorado
COLOR_TEXT = RGBColor(230, 230, 230)  # Blanco Humo

st.set_page_config(page_title="Inversor Pro - Montevideo", layout="wide")
st.title("🏛️ Business Case: Reciclaje Inmobiliario")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input("Introduce tu Google API Key", type="password")
    st.info("Sugerencia: Si tienes un Render hecho por IA, súbelo en el campo de fotos.")

# --- FUNCIONES DE LECTURA SEGURA (SOLUCIÓN AL ERROR ZIP) ---
def leer_documento_seguro(file):
    """Detecta si es Word o TXT y lee correctamente"""
    nombre = file.name.lower()
    try:
        if nombre.endswith('.docx'):
            doc = Document(file)
            return "\n".join([p.text for p in doc.paragraphs])
        else:
            # Es un archivo de texto plano (TXT)
            return file.read().decode("utf-8")
    except Exception as e:
        return f"Error al leer el archivo {nombre}: {e}"

def limpiar_texto(t):
    res = t.replace("**", "").replace("`", "").replace("#", "").strip()
    for tag in ["TITULO", "TEXTO", "DIAPOSITIVA", "CUERPO", "SLIDE", "CONTENIDO"]:
        res = re.sub(f"^{tag}:?", "", res, flags=re.IGNORECASE).strip()
    return res

def disenar_diapositiva(slide, titulo_texto):
    """Maquetación de lujo"""
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = COLOR_BG
    
    # Barra dorada superior
    line = slide.shapes.add_shape(1, 0, 0, Inches(10), Inches(0.1))
    line.fill.solid()
    line.fill.foreground_color.rgb = COLOR_GOLD
    line.line.fill.background()

    # Título
    title_shape = slide.shapes.title
    title_shape.text = limpiar_texto(titulo_texto).upper()
    tf = title_shape.text_frame
    p = tf.paragraphs[0]
    p.font.size = Pt(30)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD
    p.alignment = PP_ALIGN.LEFT

# --- INTERFAZ DE CARGA ---
st.subheader("📸 Archivos del Proyecto")
col1, col2, col3 = st.columns(3)

with col1:
    f_foto = st.file_uploader("Subir Fachada o Render", type=["jpg", "png", "jpeg"])
    f_xlsx = st.file_uploader("Planilla Excel", type=["xlsx"])

with col2:
    f_doc = st.file_uploader("Documento (Word o TXT)", type=["docx", "txt"])
    f_notas = st.text_area("Detalles de la reforma:", height=100)

with col3:
    f_media = st.file_uploader("Audio o Video", type=["mp3", "mp4", "wav"])

# --- PROCESAMIENTO ---
if st.button("🏗️ GENERAR PRESENTACIÓN"):
    if not api_key:
        st.error("Falta la API Key.")
    else:
        try:
            genai.configure(api_key=api_key)
            # Buscar modelo disponible
            models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            model_name = next((m for m in models if '1.5-flash' in m), models[0])
            model = genai.GenerativeModel(model_name)
            
            with st.spinner("Analizando y Diseñando..."):
                # Recolectar datos de forma segura
                contexto = "PROYECTO: Reciclaje Inmobiliario Montevideo\n"
                if f_doc:
                    contexto += f"CONTENIDO DOC: {leer_documento_seguro(f_doc)}\n"
                if f_notas:
                    contexto += f"NOTAS: {f_notas}\n"
                if f_xlsx:
                    try:
                        df = pd.read_excel(f_xlsx)
                        contexto += f"DATOS EXCEL: {df.to_string()}\n"
                    except: contexto += "Error leyendo Excel.\n"

                prompt_parts = [contexto]
                
                if f_foto:
                    img = Image.open(f_foto)
                    prompt_parts.append(img)
                
                if f_media:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f_media.name)[1]) as tmp:
                        tmp.write(f_media.read())
                        tmp_path = tmp.name
                    g_file = genai.upload_file(path=tmp_path)
                    while g_file.state.name == "PROCESSING":
                        time.sleep(2)
                        g_file = genai.get_file(g_file.name)
                    prompt_parts.append(g_file)

                prompt_maestro = """
                Eres un experto en inversión y arquitectura. Analiza todo.
                Genera el contenido para 8 diapositivas profesionales. 
                Separador entre diapositivas: [S]
                Formato: TITULO: (nombre) | TEXTO: (punto 1 * punto 2 * punto 3)
                
                Contenido: 
                1. Portada.
                2. Visión Arquitectónica (Analiza la fachada y propone una reforma).
                3. Mercado (Barrio Goes/Reducto).
                4. Análisis Financiero USD.
                5. KPIs: ROTE, ROI y Punto de Equilibrio.
                6. Eficiencia: Cost to Income.
                7. Beneficios Fiscales (Vivienda Promovida).
                8. Conclusión.
                """
                
                prompt_parts.insert(0, prompt_maestro)
                res = model.generate_content(prompt_parts)
                raw_text = res.text

                # --- CREAR PPTX ---
                prs = Presentation()
                bloques = raw_text.split("[S]")
                
                for i, bloque in enumerate(bloques):
                    if "|" in bloque:
                        partes = bloque.split("|")
                        t_slide = partes[0].replace("TITULO:", "").strip()
                        c_slide = partes[1].replace("TEXTO:", "").strip() if len(partes) > 1 else ""
                        
                        slide = prs.slides.add_slide(prs.slide_layouts[1])
                        disenar_diapositiva(slide, t_slide)
                        
                        # Texto a la izquierda
                        body_shape = slide.placeholders[1]
                        # Ajustamos el tamaño del cuadro de texto para que quepa la imagen a la derecha
                        body_shape.width = Inches(5.5)
                        tf = body_shape.text_frame
                        tf.word_wrap = True
                        
                        for p_text in c_slide.split("*"):
                            if len(p_text.strip()) > 2:
                                p = tf.add_paragraph()
                                p.text = "• " + limpiar_texto(p_text)
                                p.font.size = Pt(16)
                                p.font.color.rgb = COLOR_TEXT
                                p.space_after = Pt(8)
                        
                        # Si hay foto, ponerla a la derecha en Portada o Visión
                        if (i == 0 or i == 1) and f_foto:
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
                                img.save(tmp_img.name)
                                slide.shapes.add_picture(tmp_img.name, Inches(6), Inches(1.5), height=Inches(4.5))

                # Descarga
                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                st.success("✅ ¡Presentación Premium Generada!")
                st.download_button("📥 Descargar Propuesta Comercial", buf, "Propuesta_Vilardebo_Pro.pptx")
                
                if f_media: os.remove(tmp_path)

        except Exception as e:
            st.error(f"Error técnico detectado: {e}")
