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
    # AQUÍ ESTABA EL ERROR: debe ser file_uploader
    audio_file = st.file_uploader("Subir Audio o Video (Conversación)", type=["mp3", "mp4", "wav", "m4a"])
    excel_file = st.file_uploader("Subir Planilla Excel (Costos/Ingresos)", type=["xlsx"])
with col2:
    texto_extra = st.text_area("Notas adicionales o transcripción manual", placeholder="Pega aquí puntos clave que no estén en los archivos...")

# LÓGICA DE PROCESAMIENTO
if st.button("✨ GENERAR PRESENTACIÓN"):
    if not api_key:
        st.error("Por favor, introduce tu API Key de OpenAI.")
    else:
        openai.api_key = api_key
        with st.spinner("Analizando datos y diseñando diapositivas..."):
            
            # 1. Simulación de procesamiento
            # En una versión avanzada aquí iría la conexión real con GPT-4 y Whisper
            res = {
                "titulo": "Propuesta de Inversión",
                "kpis": {"Rentabilidad": "Alta", "Riesgo": "Controlado", "Mercado": "Creciente"},
            }

            # 2. CREAR POWERPOINT
            prs = Presentation()
            
            # Diapositiva 1: Portada
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            slide.shapes.title.text = "PROYECTO: " + tipo_negocio
            slide.placeholders[1].text = "Análisis Estratégico para Inversores"
            
            # Diapositiva 2: KPIs
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = "Métricas Principales"
            rows, cols = 4, 2
            table = slide.shapes.add_table(rows, cols, Inches(1), Inches(2), Inches(8), Inches(2)).table
            
            # Rellenar con datos básicos
            table.cell(0, 0).text = "Concepto"
            table.cell(0, 1).text = "Detalle"
            table.cell(1, 0).text = "Tipo de Negocio"
            table.cell(1, 1).text = tipo_negocio
            table.cell(2, 0).text = "Estilo Visual"
            table.cell(2, 1).text = estilo_visual

            # GUARDAR Y DESCARGAR
            pptx_io = io.BytesIO()
            prs.save(pptx_io)
            pptx_io.seek(0)

            st.success("¡Presentación generada con éxito!")
            st.download_button(
                label="📥 Descargar PowerPoint",
                data=pptx_io,
                file_name="Propuesta_Comercial.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            
