import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import openai
import io

# CONFIGURACIÓN DE LA PÁGINA
st.set_page_config(page_title="Generador de Proyectos de Inversión", layout="wide")
st.title("🚀 Presentaciones Comerciales Inteligentes")
st.subheader("Convierte ideas y datos en propuestas para inversores")

# BARRA LATERAL - CONFIGURACIÓN
with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input("Introduce tu OpenAI API Key", type="password")
    tipo_negocio = st.selectbox("Tipo de Proyecto", ["Reciclaje Inmobiliario", "Retail/Comercio", "Servicios/SaaS", "Gastronomía", "Otro"])
    estilo_visual = st.select_slider("Estilo Visual", options=["Clásico", "Moderno", "Industrial", "Elegante"])

# CARGA DE ARCHIVOS
col1, col2 = st.columns(2)
with col1:
    audio_file = st.file_upload("Subir Audio o Video (Conversación)", type=["mp3", "mp4", "wav", "m4a"])
    excel_file = st.file_upload("Subir Planilla Excel (Costos/Ingresos)", type=["xlsx"])
with col2:
    texto_extra = st.text_area("Notas adicionales o transcripción manual", placeholder="Pega aquí puntos clave que no estén en los archivos...")

# LÓGICA DE PROCESAMIENTO
if st.button("✨ GENERAR PRESENTACIÓN"):
    if not api_key:
        st.error("Por favor, introduce tu API Key de OpenAI.")
    else:
        openai.api_key = api_key
        with st.spinner("Analizando datos y diseñando diapositivas..."):
            
            # 1. Procesar Audio (si existe)
            transcripcion = ""
            if audio_file:
                # Aquí se llamaría a Whisper API
                transcripcion = "Datos extraídos del audio (Simulado)..."

            # 2. Análisis con IA
            prompt = f"""
            Analiza este proyecto de {tipo_negocio}. 
            Datos: {texto_extra} {transcripcion}
            Devuelve un JSON con: Título, Introducción, Público Objetivo, 
            Inversión Inicial, Flujo Mensual (Moneda local y USD), 
            Beneficios Fiscales en Uruguay, ROTE, ROI y Punto de Equilibrio.
            Sugiere una paleta de colores {estilo_visual}.
            """
            # Aquí llamamos a GPT-4o
            # (Simulamos respuesta para el ejemplo)
            res = {
                "titulo": "Reciclaje Vilardebó 1658",
                "kpis": {"ROTE": "38.3%", "Punto Equilibrio": "4.3 unidades", "Inversión": "USD 426k"},
                "color_hex": (44, 62, 80) # Azul medianoche
            }

            # 3. CREAR POWERPOINT
            prs = Presentation()
            
            # Diapositiva 1: Portada
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            slide.shapes.title.text = res["titulo"]
            slide.placeholders[1].text = "Propuesta de Inversión Estratégica"
            
            # Diapositiva 2: KPIs
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = "Métricas de Rentabilidad"
            rows, cols = 4, 2
            table = slide.shapes.add_table(rows, cols, Inches(1), Inches(2), Inches(8), Inches(2)).table
            for i, (k, v) in enumerate(res["kpis"].items()):
                table.cell(i, 0).text = k
                table.cell(i, 1).text = str(v)

            # GUARDAR Y DESCARGAR
            pptx_io = io.BytesIO()
            prs.save(pptx_io)
            pptx_io.seek(0)

            st.success("¡Presentación generada con éxito!")
            st.download_button(
                label="📥 Descargar PowerPoint",
                data=pptx_io,
                file_name="Propuesta_Inversion.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
