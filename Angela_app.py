import streamlit as st
import pandas as pd
import pdfplumber
import re  # Para usar Expresiones Regulares (Regex)
from datetime import datetime  # Para formatear la fecha
import locale  # Para forzar el idioma español en la fecha
import io

# ===============================================
# FUNCIÓN DE EXTRACCIÓN (Lógica de Negocio)
# ===============================================


def extract_data_from_pdf(pdf_file):
    """Extrae el Nombre, la Fecha, el Número, el Total y la Descripción del PDF."""

    # Intentar establecer el idioma español para manejar el nombre del mes ("Marzo")
    try:
        locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
    except locale.Error:
        try:
            locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')
        except locale.Error:
            pass

    with pdfplumber.open(pdf_file) as pdf:
        first_page = pdf.pages[0]
        text = first_page.extract_text()

        # ⚠️ SOLUCIÓN CRÍTICA: Limpiar el texto de caracteres problemáticos
        text = text.replace('\n', ' ').replace('\r', ' ')
        text = re.sub(r'\s+', ' ', text).strip()

    # --- LÓGICA DE EXTRACCIÓN CON REGEX ACTUALIZADA ---

    # 1. CLIENTE (Busca 'SR.(A)' y captura lo que sigue en la misma línea)
    # Patrón: SR.(A) o SR.A o SR(A), seguido de espacios y luego el nombre.
    client_match = re.search(
        r"SR\.\(?A\)?[\s:]*([^\n\r]+?)(?:\s+RUT|[\n\r]|$)", text, re.IGNORECASE)

    # Intenta capturar lo que sigue después del patrón, eliminando posibles espacios iniciales
    extracted_name = client_match.group(
        1).strip() if client_match else "No encontrado"

    # 2. NÚMERO (Busca 'N° :' o 'N°', y captura dígitos)
    number_match = re.search(
        r"N°\s*:\s*(\d+)", text, re.IGNORECASE)
    extracted_number = number_match.group(
        1).strip() if number_match else "No encontrado"

    # 3. FECHA (Busca 'Fecha de Emisión :' y captura el día, mes y año)
    date_match = re.search(
        r"Fecha\s+de\s+Emisión\s*:\s*(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})", text, re.IGNORECASE)

    extracted_date = "Error de Formato"
    if date_match:
        try:
            # Reconstruye la cadena para que datetime la entienda (e.g., "20 de Marzo de 2020")
            date_str = f"{date_match.group(1)} de {date_match.group(2)} de {date_match.group(3)}"
            date_obj = datetime.strptime(date_str, '%d de %B de %Y')
            # Formato DD-MM-AA
            extracted_date = date_obj.strftime('%d-%m-%y')
        except Exception:
            extracted_date = "Error de Formato"

    # 4. TOTAL (PESOS) (Busca 'Total Cuenta Única Telefónica $ ' y captura el número con puntos)
    # Patrón: Busca la frase, ignora el '$', y captura el número con puntos o comas.
    total_match = re.search(
        r"Total\s+Cuenta\s+Única\s+Telefónica\s+\$\s*([\d\.,]+)", text, re.IGNORECASE)
    extracted_total = total_match.group(1) if total_match else "No encontrado"

    # 5. DESCRIPCIÓN (Se mantiene la lógica general o se establece como vacía/fija si no hay patrón)
    # Ya que no se proporcionó un nuevo patrón de descripción, se deja en "Factura Telefónica"
    extracted_description = "Factura Telefónica"

    # --- FIN DE LA LÓGICA DE EXTRACCIÓN ---

    # Esta estructura no cambia, define las columnas de salida
    data = [
        {
            "CLIENT": extracted_name,
            "DATE": extracted_date,
            "NUMBER": extracted_number,
            "DOLLARS": "",
            "PESOS": extracted_total,
            "EUROS": "",
            "DESCRIPTION": extracted_description
        }
    ]

    return data

# ===============================================
# INTERFAZ STREAMLIT (Lógica de la Aplicación Web)
# ===============================================


def main():
    st.set_page_config(page_title="PDF a Excel Simple")
    st.title("📄 Extracción Automática de PDF a Excel")
    st.subheader("Paso 1: Cargar el Archivo PDF")

    # Componente para subir el archivo PDF
    uploaded_pdf = st.file_uploader(
        "Sube el archivo PDF (Factura Telefónica):",
        type=["pdf"],
        accept_multiple_files=False
    )

    if uploaded_pdf is not None:
        st.success(f"Archivo cargado: **{uploaded_pdf.name}**")

        if st.button("Procesar y Generar Nuevo Excel"):
            st.info("Extrayendo datos y generando archivo...")

            try:
                pdf_data = io.BytesIO(uploaded_pdf.getvalue())
                extracted_data = extract_data_from_pdf(pdf_data)

                # Usamos el DataFrame para asegurar el orden y las columnas
                df = pd.DataFrame(extracted_data, columns=[
                    "CLIENT", "DATE", "NUMBER", "DOLLARS", "PESOS", "EUROS", "DESCRIPTION"])

                st.subheader("✅ Datos Extraídos (Vista Previa)")
                st.dataframe(df)

                # B. Crear el archivo Excel en memoria
                output = io.BytesIO()

                # Función para limpiar el Total (quita el punto o coma)
                def clean_total(x):
                    if isinstance(x, str):
                        # Quitar todos los puntos y reemplazar la última coma por un punto decimal si existe
                        return float(x.replace('.', '').replace(',', '.')) if re.match(r'^[\d\.,]+$', x) else x
                    return x

                # Aplicamos la limpieza a la columna PESOS
                df['PESOS'] = df['PESOS'].apply(clean_total)

                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False,
                                sheet_name='Datos Factura')
                output.seek(0)

                # C. Botón de descarga
                st.subheader("⬇️ Archivo Excel Generado")
                st.download_button(
                    label="Descargar Excel de Factura",
                    data=output.read(),
                    file_name=f"Factura_{df['NUMBER'].iloc[0]}_Extraída.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.balloons()

            except Exception as e:
                st.error(
                    f"Ocurrió un error al procesar el archivo. Error: {e}")


if __name__ == "__main__":
    main()
