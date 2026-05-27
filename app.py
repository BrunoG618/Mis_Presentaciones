import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import io
from docx import Document
import openai

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Fábrica de Inversiones Pro", layout="wide")
st.title("🚀 Presentaciones Comerciales Multimodales")
st.write("Sube audios de reuniones, Excels financieros o documentos Word para generar tu propuesta.")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input("Introduce tu OpenAI API Key", type="password")
    tipo_negocio = st.selectbox("Sector del Proyecto", ["Reciclaje Inmobiliario", "Retail", "Gastronomía", "Tecnología", "Otro"])
    st.divider()
    st.info("Esta herramienta usa Whisper para audio, Pandas para Excel y GPT-4o para el diseño de la presentación.")

# --- FUNCIONES DE PROCESAMIENTO ---

def extraer_texto_word(file):
    doc = Document(file)
    return "\n".join([p.text for p in doc.paragraphs])

def procesar_excel(file):
    df = pd.read_excel(file)
    # Convertimos el resumen del excel a texto para que la IA lo entienda
    return f"Resumen de Excel: {df.to_string()}"

def transcribir_audio(file, key):
    openai.api_key = key
    # Guardamos temporalmente el archivo para enviarlo a Whisper
    transcription = openai.audio.transcriptions.create(
        model="whisper-1", 
        file=file
    )
    return transcription.text

def configurar_slide(slide, titulo_texto):
    title = slide.shapes.title
    title.text = titulo_texto
    title_para = title.text_frame.paragraphs[0]
    title_para.font.bold = True
    title_para.font.size = Pt(30)
    title_para.font.color.rgb = RGBColor(31, 119, 180)

# --- INTERFAZ DE CARGA (3 COLUMNAS) ---
col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("📄 Documentos")
    doc_file = st.file_uploader("Subir Word o Texto", type=["docx", "txt"])
    excel_file = st.file_uploader("Subir Planilla Excel", type=["xlsx"])

with col2:
    st.subheader("🎙️ Multimedia")
    audio_file = st.file_uploader("Subir Audio o Video de la charla", type=["mp3", "mp4", "wav", "m4a"])

with col3:
    st.subheader("✍️ Notas")
    notas = st.text_area("Notas manuales o transcripción:", height=200)

# --- BOTÓN DE ACCIÓN ---
if st.button("🔥 GENERAR PRESENTACIÓN COMPLETA"):
    if not api_key:
        st.error("Por favor, introduce tu API Key.")
    else:
        openai.api_key = api_key
        with st.spinner("Procesando archivos y analizando rentabilidad..."):
            
            try:
                # 1. Recolectar información de todas las fuentes
                contexto_total = f"Tipo de Negocio: {tipo_negocio}\n"
                
                if notas:
                    contexto_total += f"Notas: {notas}\n"
                
                if doc_file:
                    if doc_file.name.endswith(".docx"):
                        contexto_total += f"Contenido Word: {extraer_texto_word(doc_file)}\n"
                    else:
                        contexto_total += f"Contenido TXT: {doc_file.read().decode('utf-8')}\n"
                
                if excel_file:
                    contexto_total += f"Datos Financieros Excel: {procesar_excel(excel_file)}\n"
                
                if audio_file:
                    st.info("Transcribiendo audio... esto puede tardar un momento.")
                    contexto_total += f"Transcripción de Audio: {transcribir_audio(audio_file, api_key)}\n"

                # 2. IA analiza todo y genera los slides
                prompt = f"""
                Actúa como un experto en inversiones inmobiliarias y comerciales en Uruguay.
                Analiza la siguiente información acumulada:
                ---
                {contexto_total}
                ---
                Crea una presentación comercial de 7 diapositivas con este formato exacto:
                SLIDE | TÍTULO | CONTENIDO (en puntos clave)

                Asegúrate de incluir:
                1. Introducción y Visión.
                2. Ubicación y Mercado (Menciona Montevideo/Uruguay).
                3. Costos e Inversión (en USD y moneda local si aparece).
                4. Flujos Mensuales y KPIs (Calcula ROTE, ROI, Punto de Equilibrio y Cost to Income).
                5. Beneficios Fiscales (Ley de Vivienda Promovida/COMAP).
                6. Rol del Administrador y Beneficios antes de impuestos.
                7. Conclusión y Cierre.
                """

                response = openai.chat.completions.create(
                    model="gpt-4o",
                    messages=[{"role": "user", "content": prompt}]
                )
                analisis = response.choices[0].message.content

                # 3. Construir el PowerPoint
                prs = Presentation()
                slides_raw = analisis.split("SLIDE")
                
                for s in slides_raw:
                    if "|" in s:
                        partes = s.split("|")
                        titulo = partes[1].strip()
                        cuerpo = partes[2].strip()
                        
                        slide = prs.slides.add_slide(prs.slide_layouts[1])
                        configurar_slide(slide, titulo)
                        
                        tf = slide.placeholders[1].text_frame
                        tf.word_wrap = True
                        for punto in cuerpo.split(". "):
                            if len(punto) > 3:
                                p = tf.add_paragraph()
                                p.text = "• " + punto.strip()
                                p.font.size = Pt(18)
                                p.space_after = Pt(12)

                # 4. Descarga
                buf = io.BytesIO()
                prs.save(buf)
                buf.seek(0)
                
                st.success("✅ ¡Presentación generada con todos los datos!")
                st.download_button("📥 Descargar Presentación Pro", buf, "Propuesta_Inversor.pptx")
                
                with st.expander("Ver contenido analizado por la IA"):
                    st.write(analisis)

            except Exception as e:
                st.error(f"Se produjo un error al procesar: {e}")
