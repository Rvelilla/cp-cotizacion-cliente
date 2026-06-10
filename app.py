import streamlit as st
import os
import tempfile
import base64
from procesador_pdf import procesar_cotizacion

st.set_page_config(page_title="Generador de Cotizaciones - Carrocerías Panamericana", layout="centered")

# Renderizado de Logo y Título alineados horizontalmente
try:
    with open("assets/cp-logo.png", "rb") as f:
        logo_b64 = base64.b64encode(f.read()).decode()
    st.markdown(
        f"""
        <div style="display: flex; align-items: center; gap: 15px; margin-bottom: 20px;">
            <img src="data:image/png;base64,{logo_b64}" width="60">
            <h1 style="margin: 0; font-size: 2.2rem; color: #31333F;">Generador de Cotizaciones</h1>
        </div>
        """,
        unsafe_allow_html=True
    )
except Exception:
    # Fallback en caso de error con el archivo de imagen
    st.title("Generador de Cotizaciones")

st.write("Sube la Cotización en PDF.")

asesores = {
    "Claudia Velilla": "Dir. Mercadeo",
    "Juan Pablo Jimenez": "Asesor Comercial"
}

nombre_asesor = st.selectbox("Selecciona el Asesor:", list(asesores.keys()))
cargo_asesor = asesores[nombre_asesor]

uploaded_file = st.file_uploader("Sube el PDF aquí", type="pdf")

if uploaded_file is not None:
    if not os.path.exists("plantillas/plantilla_cotizacion.docx"):
        st.error("⚠️ Falta la plantilla de Cotización.")
    else:
        if st.button("Generar Cotización"):
            with st.spinner("Analizando PDF..."):
                # Crear un archivo temporal seguro
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                    tmp_file.write(uploaded_file.getbuffer())
                    temp_pdf_path = tmp_file.name
                
                try:
                    word_path = procesar_cotizacion(temp_pdf_path, nombre_asesor, cargo_asesor)
                    nombre_archivo_salida = os.path.basename(word_path)

                    with open(word_path, "rb") as f:
                        st.download_button(
                            label="📥 Descargar Cotización en Word",
                            data=f,
                            file_name=nombre_archivo_salida,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    st.success("¡Cotización generada exitosamente!")
                except Exception as e:
                    st.error(f"Error al procesar: {e}")
                finally:
                    if os.path.exists(temp_pdf_path):
                        os.remove(temp_pdf_path)
                    if 'word_path' in locals() and os.path.exists(word_path):
                        os.remove(word_path)
