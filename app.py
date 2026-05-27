import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
import io

# 1. Configuración de la página
st.set_page_config(page_title="Presentaciones Comerciales", layout="wide")
st.title("💼 Generador de Proyectos de Inversión")
st.write("Carga tus datos y descarga una presentación profesional para tus inversores.")

# 2. Panel Lateral
with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input("OpenAI API Key", type="password")
    tipo_negocio = st.selectbox("Tipo de Proyecto", ["Reciclaje Inmobiliario", "Retail", "Tecnología", "Gastronomía", "Otro"])
    estilo = st.selectbox("Estilo Visual", ["Industrial", "Moderno", "Corporativo"])

# 3. Entrada de datos
col1, col2 = st.columns(2)
with col1:
    audio = st.file_uploader("Subir Audio/Video de la charla", type=["mp3", "mp4", "wav", "m4a"])
    excel = st.file_uploader("Subir Planilla Excel", type=["xlsx"])
with col2:
    notas = st.text_area("Notas adicionales", placeholder="Escribe aquí puntos clave...", height=150)

# 4. Botón principal
if st.button("🚀 GENERAR PRESENTACIÓN"):
    if not api_key:
        st.warning("⚠️ Por favor, introduce tu API Key en la barra lateral.")
    else:
        with st.spinner("Analizando y diseñando..."):
            # Generamos el archivo PowerPoint
            prs = Presentation()
            
            # Diapositiva 1: Portada
            slide1 = prs.slides.add_slide(prs.slide_layouts[0])
            slide1.shapes.title.text = f"PROYECTO: {tipo_negocio}"
            slide1.placeholders[1].text = "Análisis Estratégico y Financiero para Inversores"
            
            # Diapositiva 2: Métricas (Estructura Básica)
            slide2 = prs.slides.add_slide(prs.slide_layouts[5])
            slide2.shapes.title.text = "Indicadores Principales"
            
            # Creamos una tabla simple
            table = slide2.shapes.add_table(3, 2, Inches(1), Inches(2), Inches(8), Inches(2)).table
            table.cell(0, 0).text = "Concepto"
            table.cell(0, 1).text = "Detalle"
            table.cell(1, 0).text = "Sector"
            table.cell(1, 1).text = tipo_negocio
            table.cell(2, 0).text = "Ubicación / Contexto"
            table.cell(2, 1).text = "Uruguay (Beneficios Fiscales Aplicables)"

            # Preparamos el archivo para descargar
            binary_output = io.BytesIO()
            prs.save(binary_output)
            binary_output.seek(0)
            
            st.success("✅ ¡Presentación lista!")
            
            # Botón de descarga
            st.download_button(
                label="📥 Descargar PowerPoint",
                data=binary_output,
                file_name="Propuesta_Inversion.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

st.info("Nota: Esta es la versión base. Una vez que verifiques que funciona, activaremos la conexión total con la IA.")
