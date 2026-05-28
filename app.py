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

# --- ESTÉTICA DE ALTO NIVEL ---
COLOR_BG = RGBColor(15, 15, 15)       # Negro Grafito (Elegancia total)
COLOR_GOLD = RGBColor(197, 160, 82)   # Dorado Champagne para acentos
COLOR_TEXT = RGBColor(230, 230, 230)  # Blanco Humo para lectura

st.set_page_config(page_title="Inversor Pro - Montevideo", layout="wide")
st.title("🏛️ Business Case: Reciclaje Inmobiliario Premium")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input("Introduce tu Google API Key", type="password")
    st.info("Sugerencia: Para el render, usa una herramienta de 'AI Architectural Design' y sube el resultado aquí.")

# --- FUNCIONES DE LIMPIEZA ---
def limpiar_texto(t):
    res = t.replace("**", "").replace("`", "").replace("#", "").strip()
    # Elimina palabras de control de la IA
    for tag in ["TITULO", "TEXTO", "DIAPOSITIVA", "CUERPO", "SLIDE", "CONTENIDO"]:
        res = re.sub(f"^{tag}:?", "", res, flags=re.IGNORECASE).strip()
    return res

def disenar_diapositiva(slide, titulo_texto):
    """Diseño de diapositiva con estilo de estudio de arquitectura"""
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = COLOR_BG
    
    # Línea decorativa superior dorada
    line = slide.shapes.add_shape(1, 0, 0, Inches(10), Inches(0.1))
    line.fill.solid()
    line.fill.foreground_color.rgb = COLOR_GOLD
    line.line.fill.background()

    # Título
    title_shape = slide.shapes.title
    title_shape.text = limpiar_texto(titulo_texto).upper()
    tf = title_shape.text_frame
    p = tf.paragraphs[0]
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD
    p.alignment = PP_ALIGN.LEFT

# --- INTERFAZ ---
col1, col2 = st.columns(2)
with col1:
    st.subheader("📸 Visuales del Proyecto")
    f_foto = st.file_uploader("Subir Fachada (Actual o Render AI)", type=["jpg", "png", "jpeg"])
    f_xlsx = st.file_uploader("Excel Financiero", type=["xlsx"])
with col2:
    st.subheader("📄 Datos y Charlas")
    f_doc = st.file_uploader("Documento Word/TXT", type=["docx", "txt"])
    f_notas = st.text_area("Detalles de la reforma (ej: acabados, materiales):", height=150)

# --- PROCESAMIENTO ---
if st.button("🏗️ GENERAR PROPUESTA EJECUTIVA"):
    if not api_key:
        st.error("Introduce tu API Key.")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-1.5-flash')
            
            with st.spinner("La IA está actuando como Arquitecto y Analista..."):
                # Preparar Contexto
                ctx = f"PROYECTO: Reciclaje Vivienda Montevideo\nNotas de Obra: {f_notas}\n"
                if f_doc:
                    ctx += "\n".join([p.text for p in Document(f_doc).paragraphs])
                if f_xlsx:
                    ctx += f"\nDATOS FINANCIEROS: {pd.read_excel(f_xlsx).to_string()}"

                inputs = [ctx]
                if f_foto: 
                    img_pil = Image.open(f_foto)
                    inputs.append(img_pil)
                
                # Prompt con enfoque en "Venta de Visión"
                prompt = """
                Actúa como un Director de Desarrollo Inmobiliario. 
                Si hay una foto, descríbela y propón una intervención moderna y rentable.
                Genera 8 diapositivas profesionales. 
                Separa cada una con: [S]
                Usa el formato: TITULO: (nombre) | TEXTO: (punto 1 * punto 2 * punto 3)
                
                Contenido Sugerido:
                1. Portada.
                2. Visión Arquitectónica (Análisis de la fachada y propuesta estética).
                3. El Mercado: Barrio Goes/Reducto.
                4. Plan Financiero: Inversión Inicial y Reformas.
                5. KPIs: ROTE (Return on Tangible Equity), ROI y Punto de Equilibrio.
                6. Eficiencia Operativa: Cost to Income.
                7. Beneficios Fiscales: Ley de Vivienda Promovida Uruguay.
                8. Conclusión y Branding del Proyecto.
                """
                
                res = model.generate_content([prompt] + inputs)
                raw_text = res.text

                # --- CONSTRUCCIÓN PPTX ---
                prs = Presentation()
                bloques = raw_text.split("[S]")
                
                for i, bloque in enumerate(bloques):
                    if "|" in bloque:
                        partes = bloque.split("|")
                        t_slide = partes[0].replace("TITULO:", "").strip()
                        c_slide = partes[1].replace("TEXTO:", "").strip() if len(partes) > 1 else ""
                        
                        slide = prs.slides.add_slide(prs.slide_layouts[1])
                        disenar_diapositiva(slide, t_slide)
                        
                        # Texto
                        body_shape = slide.placeholders[1]
                        tf = body_shape.text_frame
                        tf.word_wrap = True
                        
                        for p_text in c_slide.split("*"):
                            if len(p_text.strip()) > 2:
                                p = tf.add_paragraph()
                                p.text = "• " + limpiar_texto(p_text)
                                p.font.size = Pt(18)
                                p.font.color.rgb = COLOR_TEXT
                                p.space_after = Pt(12)
                        
                        # Si hay foto, ponerla en la diapositiva de Portada o Arquitectura
                        if (i == 0 or i == 1) and f_foto:
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                                img_pil.save(tmp.name)
                                # Colocar a la derecha de forma elegante
                                slide.shapes.add_picture(tmp.name, Inches(6.2), Inches(1.8), height=Inches(4.5))

                # Descarga
                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                st.success("✅ ¡Propuesta de Inversión Generada!")
                st.download_button("📥 Descargar Propuesta Comercial", buf, "Propuesta_Inmobiliaria_Pro.pptx")

        except Exception as e:
            st.error(f"Error técnico detectado: {e}")
