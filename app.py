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
st.title("🏆 Business Case Generator: Presentaciones Premium")

# --- CONEXIÓN AUTOMÁTICA CON API KEY (SECRETS) ---
# Primero busca en Secrets, si no está, permite ponerla en la barra lateral
api_key = st.secrets.get("GOOGLE_API_KEY") or st.sidebar.text_input("Introduce tu Google API Key", type="password")
moneda_local = st.sidebar.text_input("Moneda Local (ej: UYU, ARS)", value="UYU")

# --- FUNCIONES TÉCNICAS ---
def hex_to_rgb(hex_color):
    try:
        hex_color = hex_color.lstrip('#')
        return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))
    except: return RGBColor(31, 119, 180)

def leer_archivo_seguro(file):
    nombre = file.name.lower()
    try:
        if nombre.endswith('.docx'):
            doc = Document(file)
            return "\n".join([p.text for p in doc.paragraphs])
        return file.read().decode("utf-8", errors="ignore")
    except: return "Error leyendo archivo."

def aplicar_estilo_slide(slide, titulo_texto, color_acento_hex):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(15, 15, 15)
    color_acento = hex_to_rgb(color_acento_hex)
    title_shape = slide.shapes.title
    title_shape.text = titulo_texto.upper()
    title_para = title_shape.text_frame.paragraphs[0]
    title_para.font.size = Pt(32)
    title_para.font.bold = True
    title_para.font.color.rgb = color_acento
    title_para.alignment = PP_ALIGN.LEFT

# --- INTERFAZ ---
st.subheader("📥 Carga Multimodal de Proyecto")
c1, c2, c3 = st.columns(3)
with c1:
    f_doc = st.file_uploader("1. Conversación (Word o TXT)", type=["docx", "txt"])
    f_xlsx = st.file_uploader("2. Planilla Excel Financiera", type=["xlsx"])
with c2:
    f_media = st.file_uploader("3. Audio o Video", type=["mp3", "mp4", "wav", "m4a"])
    f_foto = st.file_uploader("4. Imagen (Fachada o Logo)", type=["jpg", "png", "jpeg"])
with c3:
    f_notas = st.text_area("5. Notas y Contexto:", height=160, placeholder="Define el negocio aquí...")

# --- PROCESAMIENTO ---
if st.button("🚀 GENERAR PRESENTACIÓN PREMIUM"):
    if not api_key:
        st.error("⚠️ Configura la API Key en Settings > Secrets o en la barra lateral.")
    else:
        try:
            genai.configure(api_key=api_key)
            modelos_disponibles = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            modelo_nombre = next((m for m in modelos_disponibles if '1.5-flash' in m), modelos_disponibles[0])
            model = genai.GenerativeModel(modelo_nombre)
            
            with st.spinner("Analizando y elevando la calidad del Business Case..."):
                contexto = f"MONEDA LOCAL: {moneda_local}\n"
                if f_notas: contexto += f"NOTAS: {f_notas}\n"
                if f_doc: contexto += f"DOCUMENTO: {leer_archivo_seguro(f_doc)}\n"
                if f_xlsx: contexto += f"EXCEL: {pd.read_excel(f_xlsx).to_string()}\n"
                
                inputs_ia = [contexto]
                if f_foto: inputs_ia.append(Image.open(f_foto))
                
                t_path = None
                if f_media:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f_media.name)[1]) as tmp:
                        tmp.write(f_media.read())
                        t_path = tmp.name
                    gf = genai.upload_file(path=t_path)
                    while gf.state.name == "PROCESSING": time.sleep(2); gf = genai.get_file(gf.name)
                    inputs_ia.append(gf)

                # PROMPT DE ALTA CALIDAD
                prompt_maestro = f"""
                Actúa como un Consultor de Inversiones de Wall Street. Analiza este proyecto con profundidad.
                1. Detecta el sector y elige un color HEX vibrante y sofisticado.
                2. Calcula con precisión o estima basándote en la lógica financiera: ROTE, ROI, Cost to Income y Punto de Equilibrio.
                3. Proyecta flujos financieros detallados en {moneda_local} y USD.
                4. Sugiere una tipografía moderna (como Montserrat o Montserrat Bold) y una paleta de colores complementaria.
                
                FORMATO DE RESPUESTA:
                COLOR: #HEX
                [SLIDE] Título | Punto Detallado 1 * Punto Detallado 2 * Punto Detallado 3 * Punto Detallado 4
                
                Genera 10 diapositivas: Portada, Visión Estratégica, Público y Mercado, Características del Negocio, 
                Zonas de Inversión, Costos de Inversión inicial, Flujos de Ingresos y Gastos, 
                Beneficios Fiscales (Ley de Vivienda Promovida/COMAP si es Uruguay), 
                KPIs y Rentabilidad Proyectada, Conclusión y Siguientes Pasos.
                """
                
                res = model.generate_content([prompt_maestro] + inputs_ia)
                texto_ia = res.text

                # --- CONSTRUCCIÓN PPTX ---
                prs = Presentation()
                match_c = re.search(r'COLOR:\s*(#[0-9A-Fa-f]{6})', texto_ia)
                accent_hex = match_c.group(1) if match_c else "#D4AF37"
                
                bloques = texto_ia.split("[SLIDE]")
                for i, bloque in enumerate(bloques):
                    if "|" in bloque:
                        partes = bloque.split("|")
                        titulo_s = partes[0].replace(f"COLOR: {accent_hex}", "").strip()
                        contenido_s = partes[1].strip()
                        
                        slide = prs.slides.add_slide(prs.slide_layouts[1])
                        aplicar_estilo_slide(slide, titulo_s, accent_hex)
                        
                        body_shape = slide.placeholders[1]
                        if (i <= 2 and f_foto): body_shape.width = Inches(5.2)
                        
                        tf = body_shape.text_frame
                        tf.word_wrap = True
                        for punto in contenido_s.split("*"):
                            if len(punto.strip()) > 3:
                                p = tf.add_paragraph()
                                p.text = "• " + punto.strip().replace("**", "")
                                p.font.size = Pt(18)
                                p.font.color.rgb = RGBColor(235, 235, 235)
                                p.space_after = Pt(10)
                        
                        if (i <= 2 and f_foto):
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
                                Image.open(f_foto).save(tmp_img.name)
                                slide.shapes.add_picture(tmp_img.name, Inches(5.5), Inches(1.5), height=Inches(4.5))

                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                st.success("✅ ¡Presentación Premium Generada!")
                st.download_button("📥 Descargar PowerPoint", buf, "Propuesta_Comercial_Premium.pptx")
                
                if t_path and os.path.exists(t_path): os.remove(t_path)

        except Exception as e:
            st.error(f"Error técnico detectado: {e}")
