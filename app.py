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

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador Pro Universal", layout="wide")
st.title("🏆 Business Case Generator: Presentaciones de Inversión")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Configuración")
    api_key = st.text_input("Introduce tu Google API Key", type="password")
    moneda_local = st.text_input("Moneda Local (ej: UYU, ARS)", value="UYU")
    st.info("KPIs: ROTE, ROI, Cost to Income, Punto de Equilibrio.")

# --- FUNCIONES DE LECTURA Y ESTILO (CORREGIDAS) ---

def hex_to_rgb(hex_color):
    try:
        hex_color = hex_color.lstrip('#')
        return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))
    except:
        return RGBColor(31, 119, 180) # Azul corporativo por defecto

def leer_archivo_seguro(file):
    """Evita el error de Zip diferenciando Word de TXT"""
    nombre = file.name.lower()
    try:
        if nombre.endswith('.docx'):
            doc = Document(file)
            return "\n".join([p.text for p in doc.paragraphs])
        else:
            return file.read().decode("utf-8", errors="ignore")
    except Exception as e:
        return f"Error leyendo {nombre}: {e}"

def aplicar_estilo_slide(slide, titulo_texto, color_acento_hex):
    """CORRECCIÓN: El fondo se aplica al objeto 'slide', no a 'presentation'"""
    # 1. Fondo Oscuro
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(15, 15, 15)
    
    color_acento = hex_to_rgb(color_acento_hex)
    
    # 2. Título
    title_shape = slide.shapes.title
    title_shape.text = titulo_texto.upper()
    title_para = title_shape.text_frame.paragraphs[0]
    title_para.font.size = Pt(32)
    title_para.font.bold = True
    title_para.font.color.rgb = color_acento
    title_para.alignment = PP_ALIGN.LEFT

# --- INTERFAZ (TODOS LOS CAMPOS REQUERIDOS) ---
st.subheader("📥 Carga de Datos del Proyecto")
c1, c2, c3 = st.columns(3)
with c1:
    f_doc = st.file_uploader("1. Conversación (Word o TXT)", type=["docx", "txt"])
    f_xlsx = st.file_uploader("2. Planilla Excel Financiera", type=["xlsx"])
with c2:
    f_media = st.file_uploader("3. Audio o Video", type=["mp3", "mp4", "wav", "m4a"])
    f_foto = st.file_uploader("4. Imagen (Fachada, Logo o Producto)", type=["jpg", "png", "jpeg"])
with c3:
    f_notas = st.text_area("5. Notas y Contexto:", height=160, placeholder="Explica el negocio aquí (Lavaderos, Apps, Inmuebles, etc.)")

# --- PROCESAMIENTO ---
if st.button("🚀 GENERAR PRESENTACIÓN COMERCIAL"):
    if not api_key:
        st.error("Por favor, introduce tu API Key.")
    else:
        try:
            genai.configure(api_key=api_key)
            
            # --- AUTO-DETECCIÓN DE MODELO (EVITA ERROR 404) ---
            model_list = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            model_name = next((m for m in model_list if '1.5-flash' in m), model_list[0])
            model = genai.GenerativeModel(model_name)
            
            with st.spinner("Analizando modelo de negocio y diseñando propuesta..."):
                # Unificar datos
                contexto = f"MONEDA LOCAL: {moneda_local}\n"
                if f_notas: contexto += f"NOTAS: {f_notas}\n"
                if f_doc: contexto += f"DOCUMENTO: {leer_archivo_seguro(f_doc)}\n"
                if f_xlsx:
                    df = pd.read_excel(f_xlsx)
                    contexto += f"EXCEL: {df.to_string()}\n"
                
                inputs_ia = [contexto]
                if f_foto: inputs_ia.append(Image.open(f_foto))
                
                # Manejo de Multimedia
                t_path = None
                if f_media:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f_media.name)[1]) as tmp:
                        tmp.write(f_media.read())
                        t_path = tmp.name
                    gf = genai.upload_file(path=t_path)
                    while gf.state.name == "PROCESSING":
                        time.sleep(2)
                        gf = genai.get_file(gf.name)
                    inputs_ia.append(gf)

                # PROMPT UNIVERSAL PARA CUALQUIER NEGOCIO
                prompt_maestro = f"""
                Actúa como un Consultor de Inversiones. Analiza este proyecto sin importar el sector.
                1. Detecta el sector y elige un color HEX vibrante para el diseño.
                2. Calcula KPIs: ROTE, ROI, Cost to Income y Punto de Equilibrio.
                3. Proyecta flujos en {moneda_local} y USD.
                4. Sugiere Tipografías y Colores adecuados.
                
                RESPONDE EXCLUSIVAMENTE CON ESTE FORMATO:
                COLOR: #HEX
                [SLIDE] Título | Punto 1 * Punto 2 * Punto 3
                
                Genera 10 diapositivas que cubran: Portada, Oportunidad, Público, Características, Zonas, Inversión Inicial, Ingresos/Gastos, Beneficios Fiscales, KPIs Financieros y Cierre.
                """
                
                res = model.generate_content([prompt_maestro] + inputs_ia)
                texto_ia = res.text

                # --- CONSTRUCCIÓN DEL PPTX (CORREGIDA) ---
                prs = Presentation()
                
                # Extraer color
                match_c = re.search(r'COLOR:\s*(#[0-9A-Fa-f]{6})', texto_ia)
                accent_hex = match_c.group(1) if match_c else "#D4AF37"
                
                # Dividir por slides
                bloques = texto_ia.split("[SLIDE]")
                for i, bloque in enumerate(bloques):
                    if "|" in bloque:
                        secciones = bloque.split("|")
                        titulo_s = secciones[0].replace(f"COLOR: {accent_hex}", "").strip()
                        contenido_s = secciones[1].strip()
                        
                        # Crear Slide y aplicar fondo e identidad
                        slide = prs.slides.add_slide(prs.slide_layouts[1])
                        aplicar_estilo_slide(slide, titulo_s, accent_hex)
                        
                        # Cuerpo de texto
                        body_shape = slide.placeholders[1]
                        # Si es el slide 1 o 2 y hay foto, dejar espacio para la imagen
                        if (i <= 2 and f_foto):
                            body_shape.width = Inches(5.2)
                        
                        tf = body_shape.text_frame
                        tf.word_wrap = True
                        for punto in contenido_s.split("*"):
                            if len(punto.strip()) > 3:
                                p = tf.add_paragraph()
                                p.text = "• " + punto.strip().replace("**", "")
                                p.font.size = Pt(18)
                                p.font.color.rgb = RGBColor(240, 240, 240)
                                p.space_after = Pt(10)
                        
                        # Insertar Imagen si corresponde
                        if (i <= 2 and f_foto):
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
                                Image.open(f_foto).save(tmp_img.name)
                                slide.shapes.add_picture(tmp_img.name, Inches(5.5), Inches(1.5), height=Inches(4.5))

                # Exportar
                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                st.success("✅ ¡Presentación Generada!")
                st.download_button("📥 Descargar PowerPoint", buf, "Propuesta_Comercial.pptx")
                
                if t_path and os.path.exists(t_path): os.remove(t_path)

        except Exception as e:
            st.error(f"Error técnico detectado: {e}")
