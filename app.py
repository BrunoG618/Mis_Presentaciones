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

# --- CONFIGURACIÓN DE ESTILO ---
COLOR_BG = RGBColor(30, 30, 30)       # Fondo Gris Oscuro Premium
COLOR_GOLD = RGBColor(212, 175, 55)   # Dorado para Títulos/Acentos
COLOR_WHITE = RGBColor(255, 255, 255) # Blanco para texto

st.set_page_config(page_title="Presentaciones Comerciales Pro", layout="wide")
st.title("🏆 Generador de Propuestas Comerciales Premium")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input("Introduce tu Google API Key", type="password")
    st.info("Esta versión genera un diseño oscuro elegante para inversores.")

# --- FUNCIONES DE LIMPIEZA Y DISEÑO ---
def limpiar_cadena(t):
    """Elimina etiquetas de la IA y símbolos innecesarios"""
    tags = ["DIAPOSITIVA", "TITULO", "TÍTULO", "CONTENIDO", "CUERPO", "SLIDE"]
    res = t.replace("**", "").replace("`", "").strip()
    for tag in tags:
        res = re.sub(f"^{tag}:?", "", res, flags=re.IGNORECASE).strip()
    return res

def disenar_slide(slide, titulo_texto):
    """Aplica el look & feel premium"""
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = COLOR_BG
    
    # Título dorado y centrado
    title_shape = slide.shapes.title
    title_shape.text = limpiar_cadena(titulo_texto).upper()
    tf = title_shape.text_frame
    p = tf.paragraphs[0]
    p.font.size = Pt(34)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD
    p.alignment = PP_ALIGN.CENTER

# --- INTERFAZ ---
col1, col2, col3 = st.columns(3)
with col1:
    f_doc = st.file_uploader("Documento (Word/TXT)", type=["docx", "txt"])
    f_xlsx = st.file_uploader("Planilla Excel", type=["xlsx"])
with col2:
    f_foto = st.file_uploader("Foto de la Propiedad", type=["jpg", "png", "jpeg"])
    f_media = st.file_uploader("Audio/Video", type=["mp3", "mp4", "wav"])
with col3:
    f_notas = st.text_area("Notas extra:", height=150)

# --- PROCESAMIENTO ---
if st.button("🚀 GENERAR PROPUESTA DE ALTO IMPACTO"):
    if not api_key:
        st.error("Falta la API Key.")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-1.5-flash')
            
            with st.spinner("La IA está analizando y diseñando tu presentación..."):
                # Recopilar contexto
                ctx = f"PROYECTO: Reciclaje Inmobiliario\nNotas: {f_notas}\n"
                if f_doc:
                    doc = Document(f_doc)
                    ctx += "\n".join([p.text for p in doc.paragraphs])
                if f_xlsx:
                    df = pd.read_excel(f_xlsx)
                    ctx += f"\nFINANZAS: {df.to_string()}"

                inputs = [ctx]
                if f_foto: inputs.append(Image.open(f_foto))
                
                if f_media:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f_media.name)[1]) as tmp:
                        tmp.write(f_media.read())
                        t_path = tmp.name
                    gf = genai.upload_file(path=t_path)
                    while gf.state.name == "PROCESSING":
                        time.sleep(2)
                        gf = genai.get_file(gf.name)
                    inputs.append(gf)

                # PROMPT DE DISEÑO
                prompt = """
                Analiza todo el material. Eres un Director de Negocios e Inversiones.
                Genera el contenido para 8 diapositivas. 
                IMPORTANTE: Responde SOLO con el contenido de las diapositivas usando este separador: [S]
                En cada diapositiva, usa este formato: 
                TITULO: (nombre del titulo)
                TEXTO: (punto 1 | punto 2 | punto 3)
                
                Incluye: 1. Portada, 2. Concepto y Diseño (analiza la foto), 3. Ubicación, 4. Inversión Inicial, 
                5. KPIs (ROTE, ROI, Punto Equilibrio), 6. Gastos vs Ingresos, 7. Beneficios Fiscales Uruguay, 8. Cierre.
                """
                
                res = model.generate_content([prompt] + inputs)
                raw_text = res.text

                # --- CONSTRUCCIÓN PPTX ---
                prs = Presentation()
                
                # Diapositiva de Portada Especial
                bloques = raw_text.split("[S]")
                
                for i, bloque in enumerate(bloques):
                    if "TITULO:" in bloque:
                        # Extraer Título y Texto
                        partes = bloque.split("TEXTO:")
                        t_slide = partes[0].replace("TITULO:", "").strip()
                        c_slide = partes[1].strip() if len(partes) > 1 else ""
                        
                        slide = prs.slides.add_slide(prs.slide_layouts[1])
                        disenar_slide(slide, t_slide)
                        
                        # Cuadro de texto
                        body_shape = slide.placeholders[1]
                        tf = body_shape.text_frame
                        tf.word_wrap = True
                        
                        for p_text in c_slide.split("|"):
                            clean_p = limpiar_cadena(p_text)
                            if len(clean_p) > 2:
                                p = tf.add_paragraph()
                                p.text = "• " + clean_p
                                p.font.size = Pt(18)
                                p.font.color.rgb = COLOR_WHITE
                                p.space_after = Pt(10)
                        
                        # Si es la diapositiva de Concepto/Portada y hay foto, ponerla
                        if (i == 0 or "CONCEPTO" in t_slide.upper()) and f_foto:
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
                                Image.open(f_foto).save(tmp_img.name)
                                slide.shapes.add_picture(tmp_img.name, Inches(6), Inches(2), height=Inches(4))

                # Descarga
                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                st.success("✅ ¡Presentación Premium Generada!")
                st.download_button("📥 Descargar Propuesta Comercial", buf, "Propuesta_V4_Premium.pptx")
                
                if f_media: os.remove(t_path)

        except Exception as e:
            st.error(f"Error detectado: {e}")
