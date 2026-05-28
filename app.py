import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import io
import os
import tempfile
from docx import Document
import google.generativeai as genai
import time

# --- 1. CONFIGURACIÓN DE ESTILO ---
AZUL_CORP = RGBColor(0, 51, 102) # Azul Inversor
GRIS_SUAVE = RGBColor(245, 245, 245)

st.set_page_config(page_title="Presentaciones Pro Uruguay", layout="wide")
st.title("🏆 Generador de Propuestas Comerciales")

# --- 2. BARRA LATERAL ---
with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input("Introduce tu Google API Key", type="password")
    tipo_negocio = st.selectbox("Sector del Proyecto", ["Reciclaje Inmobiliario", "Retail", "Gastronomía", "Tecnología", "Otro"])

# --- 3. FUNCIONES DE LIMPIEZA ---
def limpiar_texto(t):
    """Elimina asteriscos de negrita y etiquetas molestas"""
    return t.replace("**", "").replace("Título:", "").replace("Contenido:", "").strip()

def agregar_estilo_slide(slide):
    """Agrega una barra decorativa azul a la izquierda para que no sea blanco total"""
    shape = slide.shapes.add_shape(1, 0, 0, Inches(0.2), Inches(7.5)) # Rectángulo fino
    shape.fill.solid()
    shape.fill.foreground_color.rgb = AZUL_CORP
    shape.line.fill.background()

# --- 4. INTERFAZ DE CARGA ---
col1, col2 = st.columns(2)
with col1:
    f_doc = st.file_uploader("Documento (Word o TXT)", type=["docx", "txt"])
    f_xlsx = st.file_uploader("Excel Financiero", type=["xlsx"])
    f_media = st.file_uploader("Audio o Video", type=["mp3", "wav", "mp4", "m4a"])
with col2:
    f_notas = st.text_area("Notas o detalles del proyecto:", height=250)

# --- 5. PROCESAMIENTO ---
if st.button("🚀 GENERAR PRESENTACIÓN PROFESIONAL"):
    if not api_key:
        st.error("Introduce tu API Key.")
    else:
        try:
            genai.configure(api_key=api_key)
            modelos = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            seleccionado = next((m for m in modelos if '1.5-flash' in m), modelos[0])
            model = genai.GenerativeModel(seleccionado)

            with st.spinner("Analizando y Diseñando la Propuesta..."):
                # Recolección de datos
                texto_base = f"PROYECTO: {tipo_negocio}\n"
                if f_notas: texto_base += f"NOTAS: {f_notas}\n"
                if f_doc:
                    if f_doc.name.endswith(".docx"):
                        texto_base += "\n".join([p.text for p in Document(f_doc).paragraphs])
                    else: texto_base += f_doc.read().decode("utf-8")
                
                # Prompt Reforzado para evitar basura visual
                prompt = """
                Analiza los datos y genera 7 diapositivas de inversión. 
                IMPORTANTE: NO uses negritas (asteriscos). 
                Usa el formato EXACTO:
                [INICIO_DIAPOSITIVA]
                TÍTULO: Escribe aquí el título
                CUERPO: Punto 1 | Punto 2 | Punto 3 | Punto 4
                [FIN_DIAPOSITIVA]
                
                Contenido: 1. Portada, 2. Visión, 3. Ubicación/Mercado (Montevideo), 4. Inversión (USD), 5. KPIs (ROTE, Punto Equilibrio), 6. Beneficios Fiscales, 7. Cierre.
                """

                res = model.generate_content([prompt, texto_base])
                texto_ia = res.text

                # --- CONSTRUCCIÓN DEL PPTX ---
                prs = Presentation()
                
                bloques = texto_ia.split("[INICIO_DIAPOSITIVA]")
                
                for bloque in bloques:
                    if "TÍTULO:" in bloque and "CUERPO:" in bloque:
                        # Extraer Título y Cuerpo
                        try:
                            parte_titulo = bloque.split("CUERPO:")[0].replace("TÍTULO:", "").strip()
                            parte_cuerpo = bloque.split("CUERPO:")[1].replace("[FIN_DIAPOSITIVA]", "").strip()
                            
                            # Crear Slide
                            slide_layout = prs.slide_layouts[1] # Título y Contenido
                            slide = prs.slides.add_slide(slide_layout)
                            agregar_estilo_slide(slide)
                            
                            # Llenar Título
                            title_shape = slide.shapes.title
                            title_shape.text = limpiar_texto(parte_titulo)
                            title_shape.text_frame.paragraphs[0].font.color.rgb = AZUL_CORP
                            title_shape.text_frame.paragraphs[0].font.bold = True
                            
                            # Llenar Cuerpo
                            body_shape = slide.placeholders[1]
                            tf = body_shape.text_frame
                            tf.word_wrap = True
                            
                            puntos = parte_cuerpo.split("|")
                            for punto in puntos:
                                if len(punto.strip()) > 2:
                                    p = tf.add_paragraph()
                                    p.text = limpiar_texto(punto)
                                    p.font.size = Pt(18)
                                    p.level = 0
                        except:
                            continue

                # Guardar
                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                
                st.success("✅ ¡Presentación Profesional Generada!")
                st.download_button("📥 DESCARGAR POWERPOINT", buf, "Propuesta_Comercial_V3.pptx")

        except Exception as e:
            st.error(f"Error técnico: {e}")
