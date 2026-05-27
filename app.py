import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
import io
from docx import Document # Librería para leer Word

# 1. Configuración de la página
st.set_page_config(page_title="Presentaciones Comerciales", layout="wide")
st.title("💼 Generador de Proyectos de Inversión")
st.write("Sube audios, Excels o documentos de texto (Word/TXT) para generar tu propuesta.")

# 2. Panel Lateral
with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input("OpenAI API Key", type="password")
    tipo_negocio = st.selectbox("Tipo de Proyecto", ["Reciclaje Inmobiliario", "Retail", "Tecnología", "Gastronomía", "Otro"])
    estilo = st.selectbox("Estilo Visual", ["Industrial", "Moderno", "Corporativo"])

# 3. Entrada de datos (Multimodal)
col1, col2 = st.columns(2)
with col1:
    st.subheader("📁 Archivos de Origen")
    audio = st.file_uploader("Opción A: Audio/Video de la charla", type=["mp3", "mp4", "wav", "m4a"])
    documento = st.file_uploader("Opción B: Documento de Texto (Word o TXT)", type=["docx", "txt"])
    excel = st.file_uploader("Opción C: Planilla Excel (Datos Financieros)", type=["xlsx"])

with col2:
    st.subheader("📝 Notas y Transcripción")
    notas = st.text_area("Copia y pega aquí la conversación o notas adicionales:", placeholder="Escribe o pega aquí el texto de tu proyecto...", height=280)

# 4. Función para leer el contenido de los archivos subidos
def extraer_texto():
    texto_final = notas
    if documento is not None:
        if documento.name.endswith(".docx"):
            doc = Document(documento)
            texto_final += "\n" + "\n".join([para.text for para in doc.paragraphs])
        elif documento.name.endswith(".txt"):
            texto_final += "\n" + documento.read().decode("utf-8")
    return texto_final

# 5. Botón principal
if st.button("🚀 GENERAR PRESENTACIÓN"):
    if not api_key:
        st.warning("⚠️ Por favor, introduce tu API Key en la barra lateral.")
    else:
        with st.spinner("Analizando información y creando diapositivas..."):
            contenido_proyecto = extraer_texto()
            
            # Generamos el archivo PowerPoint
            prs = Presentation()
            
            # Diapositiva 1: Portada
            slide1 = prs.slides.add_slide(prs.slide_layouts[0])
            slide1.shapes.title.text = f"PROYECTO: {tipo_negocio}"
            slide1.placeholders[1].text = "Propuesta de Inversión Generada Automáticamente"
            
            # Diapositiva 2: Resumen del contenido detectado
            slide2 = prs.slides.add_slide(prs.slide_layouts[1])
            slide2.shapes.title.text = "Resumen del Proyecto"
            tf = slide2.placeholders[1].text_frame
            tf.text = "Información analizada con éxito."
            tf.add_paragraph().text = f"Fuente: {documento.name if documento else 'Texto/Notas'}"
            tf.add_paragraph().text = f"Sector: {tipo_negocio}"
            
            # Diapositiva 3: Tabla Financiera (Básica por ahora)
            slide3 = prs.slides.add_slide(prs.slide_layouts[5])
            slide3.shapes.title.text = "Estructura Financiera Estimada"
            table = slide3.shapes.add_table(3, 2, Inches(1), Inches(2), Inches(8), Inches(2)).table
            table.cell(0, 0).text = "Concepto"
            table.cell(0, 1).text = "Detalle"
            table.cell(1, 0).text = "Ubicación"
            table.cell(1, 1).text = "Montevideo, Uruguay"
            table.cell(2, 0).text = "Incentivos"
            table.cell(2, 1).text = "Vivienda Promovida / COMAP"

            # Preparamos el archivo para descargar
            binary_output = io.BytesIO()
            prs.save(binary_output)
            binary_output.seek(0)
            
            st.success("✅ ¡Presentación generada!")
            st.download_button(
                label="📥 Descargar PowerPoint",
                data=binary_output,
                file_name="Proyecto_Inversion.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
