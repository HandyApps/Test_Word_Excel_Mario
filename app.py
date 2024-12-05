import streamlit as st
import pandas as pd
from docx import Document
import os

PASSWORD = st.secrets["general"]["password"]
# Solicitar contraseña al inicio
st.title("Dame algo para continuar")
password = st.text_input("Introduce lo que tu sabes para continuar:", type="password")

if password == "":
    st.info("Esperando...")
    st.stop()

if password != PASSWORD:  # Cambia 'tu_contraseña' por la contraseña deseada
    st.error("No es correcto. No puedes continuar.")
    st.stop()

def process_word_to_excel(word_file, output_excel_path):
    doc = Document(word_file)
    data = []

    lines = [paragraph for paragraph in doc.paragraphs]
    non_empty_lines = []
    empty_line_count = 0

    for line in lines:
        if not line.text.strip():
            empty_line_count += 1
            if empty_line_count >= 11:
                st.error("Deteniendo procesamiento debido a 11 líneas vacías consecutivas.")
                return None
        else:
            empty_line_count = 0
            non_empty_lines.append(line)

    for i in range(0, len(non_empty_lines), 5):
        group = non_empty_lines[i:i+5]
        if len(group) == 5:
            question = group[0].text.strip()
            answers = [p.text.strip() for p in group[1:5]]
            correct_answer_index = 0
            for idx, paragraph in enumerate(group[1:5]):
                if any(run.bold for run in paragraph.runs):
                    correct_answer_index = idx + 1
                elif paragraph.style.name.startswith("Heading"):
                    correct_answer_index = idx + 1
            data.append({
                "Numero": len(data) + 1,
                "Pregunta": question,
                "Respuesta A": answers[0],
                "Respuesta B": answers[1],
                "Respuesta C": answers[2],
                "Respuesta D": answers[3],
                "Respuesta Correcta": correct_answer_index
            })

    if data:
        df = pd.DataFrame(data)
        df.to_excel(output_excel_path, index=False)
        return output_excel_path
    else:
        st.warning("No se encontraron preguntas procesables en el documento.")
        return None

# Streamlit UI
st.title("Procesa Archivo")
uploaded_file = st.file_uploader("Sube tu archivo", type="docx")

if uploaded_file:
    
    excel_name = uploaded_file.name.replace(".docx", ".xlsx")
    with open(uploaded_file.name, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # Procesar el archivo
    result = process_word_to_excel(uploaded_file.name, excel_name)

    if result:
        st.success("Archivo procesado correctamente.")
        st.download_button("Descargar Excel", open(result, "rb"), file_name=excel_name)
    else:
        st.error("Hubo un problema procesando el archivo.")
