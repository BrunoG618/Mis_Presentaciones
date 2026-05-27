import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import io
from docx import Document
import openai

# --- CONFIGURACIÓN ESTÉTICA ---
COLOR_FONDO = RGBColor(240, 242, 246)
COLOR_TEXTO = RGBColor(44, 62, 80)
COLOR_ACENTO = RGBColor(31, 119, 180)

st.set_page_config(page_title="Generador Pro Inversiones", layout="wide")
st.title("🏆 Presentaciones Comerciales de Alto Impacto")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("Configuración de Marca")
    api_key = st.text_input("Introduce tu OpenAI API Key", type="password")
    tipo_negocio = st.selectbox("Sector", ["Reciclaje Inmobiliario", "Retail", "Gastronomía", "Tech"])
    estilo = st.selectbox("Mood de la Presentación", ["Industrial/Arquitectura", "Corporativo/Serio", "Moderno/Limpio"])

# --- FUNCIONES DE APOYO ---
def configurar_slide(slide, titulo_texto):
    """Aplica diseño visual al slide"""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.foreground_color.rgb = RGBColor(255, 255, 255)
    
    title = slide.shapes.title
    title.text = titulo_texto
    title_para = title.text_frame.paragraphs[0]
    title_para.font.bold = True
    title_para.font.size = Pt(32)
    title_para.font.color.rgb = COLOR_ACENTO

def extraer_texto_word(file):
    doc = Document(file)
    return "\n".join([p.text for p in doc.paragraphs])

# --- INTERFAZ ---
col1, col2 = st.columns(2)
with col1:
    documento = st.file_uploader("Subir Documento (Word/TXT)", type=["docx", "txt"])
with col2:
    notas = st.text_area("Notas o Transcripción Manual:", height=150)

# --- PROMPT AVANZADO ---
if st.button("🚀 GENERAR PRESENTACIÓN EJECUTIVA"):
    if not api_key:
        st.error("Falta la API Key")
    elif not documento and not notas:
        st.warning("Sin datos para analizar")
    else:
        with st.spinner("Analizando proyecto y diseñando propuesta visual..."):
            openai.api_key = api_key
            texto_final = (notas if notas else "") + (extraer_texto_word(documento) if documento else "")
            
            prompt = f"""
            Actúa como un Senior Business Analyst. Analiza este proyecto de {tipo_negocio} en Uruguay.
            DATOS: {texto_final}
            
            Genera el contenido para 7 diapositivas. Para cada una, usa este formato exacto:
            SLIDE 1: [Título] | [Contenido detallado]
            SLIDE 2: [Título] | [Contenido detallado]
            ...etc.

            Estructura obligatoria:
            1. Portada Impactante.
            2. Resumen Ejecutivo (Visión y Oportunidad).
            3. Análisis de Mercado y Ubicación (detallar Reducto/Goes si aplica).
            4. Cuadro Financiero Detallado (Inversión inicial vs Ventas esperadas).
            5. Rentabilidad y KPIs (ROTE, ROI, Punto de Equilibrio, Cost to Income).
            6. Beneficios Fiscales (Vivienda Promovida/COMAP con exoneraciones específicas).
            7. Conclusión y Próximos Pasos.
            
            Usa un tono profesional, persuasivo y muy orientado a números.
            """

            try:
                response = openai.chat.completions.create(
                    model="gpt-4o",
                    messages=[{"role": "user", "content": prompt}]
                )
                analisis = response.choices[0].message.content
                
                # --- CONSTRUCCIÓN DEL PPTX ---
                prs = Presentation()
                
                # Parsear el texto por "SLIDE"
                slides_data = analisis.split("SLIDE")
                
                for item in slides_data:
                    if "|" in item:
                        partes = item.split("|")
                        titulo_s = partes[0].strip().replace(":", "").replace("1", "").replace("2", "").replace("3", "").replace("4", "").replace("5", "").replace("6", "").replace("7", "")
                        contenido_s = partes[1].strip()
                        
                        # Crear diapositiva
                        slide_layout = prs.slide_layouts[1] # Título y Contenido
                        slide = prs.slides.add_slide(slide_layout)
                        configurar_slide(slide, titulo_s)
                        
                        # Añadir texto con formato
                        tf = slide.placeholders[1].text_frame
                        tf.word_wrap = True
                        for linea in contenido_s.split(". "):
                            p = tf.add_paragraph()
                            p.text = "• " + linea.strip()
                            p.font.size = Pt(16)
                            p.space_after = Pt(10)

                # Guardar
                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                
                st.success("✅ ¡Presentación Profesional Generada!")
                st.download_button("📥 DESCARGAR PPT PARA INVERSORES", buf, "Propuesta_Comercial_V1.pptx")
                
                with st.expander("Ver Reporte de IA"):
                    st.write(analisis)
            except Exception as e:
                st.error(f"Error: {e}")
