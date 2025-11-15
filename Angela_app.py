import streamlit as st
import pandas as pd
import pdfplumber
import re  # Para usar Expresiones Regulares (Regex)
from datetime import datetime  # Para formatear la fecha
import locale  # Para forzar el idioma español en la fecha
import io
import xlsxwriter

# ⚠️ CORRECCIÓN CLAVE 1: Mapeo de meses
# Mapeo de meses en español a inglés para evitar problemas de 'locale' en servidores Linux
MONTH_MAPPING = {
    'enero': 'January', 'febrero': 'February', 'marzo': 'March',
    'abril': 'April', 'mayo': 'May', 'junio': 'June',
    'julio': 'July', 'agosto': 'August', 'septiembre': 'September',
    'octubre': 'October', 'noviembre': 'November', 'diciembre': 'December'
}

# Se intenta el locale, pero no es crucial si falla.
try:
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')
    except locale.Error:
        pass  # Si falla, el mapeo de meses lo soluciona.


# ===============================================
# FUNCIÓN DE EXTRACCIÓN (Lógica de Negocio)
# ===============================================

def extract_data_from_pdf(pdf_file):
    """Extrae el Nombre, la Fecha, el Número, el Total y la Descripción del PDF."""

    # Usamos try/except para manejar errores de archivos no válidos individualmente
    try:
        with pdfplumber.open(pdf_file) as pdf:
            first_page = pdf.pages[0]
            text = first_page.extract_text()

            # Limpieza crítica del texto
            text = text.replace('\n', ' ').replace('\r', ' ')
            text = re.sub(r'\s+', ' ', text).strip()

        # --- LÓGICA DE EXTRACCIÓN CON REGEX ---

        # 1. CLIENTE
        client_match = re.search(
            r"SR\.\(?A\)?[\s:]*([^\n\r]+?)(?:\s+RUT|[\n\r]|$)", text, re.IGNORECASE)
        extracted_name = client_match.group(
            1).strip() if client_match else "No encontrado"

        # 2. NÚMERO
        number_match = re.search(
            r"N°\s*:\s*(\d+)", text, re.IGNORECASE)
        extracted_number = number_match.group(
            1).strip() if number_match else "No encontrado"

        # 3. FECHA (Regex más robusta)
        date_match = re.search(
            r"Fecha\s+(?:de\s+)?Emisi[óo]n\s*:\s*(\d{1,2})\s+de\s+(\w+)\s+(?:del|de)\s+(\d{4})",
            text,
            re.IGNORECASE
        )
        extracted_date = "Error de Formato (Inicial)"

        if date_match:
            # ⚠️ Bloque try/except correctamente indentado y cerrado.
            try:
                day = date_match.group(1)
                month_es = date_match.group(2).lower()
                year = date_match.group(3)

                # APLICACIÓN DEL MAPEO DE MESES
                month_en = MONTH_MAPPING.get(month_es, month_es)

                # Crear la cadena de fecha usando el mes en inglés/mapeado
                date_str = f"{day} de {month_en} de {year}"

                # Intentar el parseo
                date_obj = datetime.strptime(date_str, '%d de %B de %Y')

                extracted_date = date_obj.strftime('%d-%m-%y')

            except Exception:
                # Este except está alineado con el try interno (dentro del if)
                extracted_date = "Error de Formato (Parseo final)"

        # 4. TOTAL (PESOS)
        total_match = re.search(
            r"Total\s+Cuenta\s+Única\s+Telefónica\s+\$\s*([\d\.,]+)", text, re.IGNORECASE)
        extracted_total = total_match.group(
            1) if total_match else "No encontrado"

        # 5. DESCRIPCIÓN
        extracted_description = "Factura Telefónica"

        # ESTA ESTRUCTURA DEBE COINCIDIR CON LA TABLA DE SALIDA
        return {
            "CLIENT": extracted_name,
            "DATE": extracted_date,
            "NUMBER": extracted_number,
            "DOLLARS": "",
            "PESOS": extracted_total,
            "EUROS": "",
            "DESCRIPTION": extracted_description
        }

    # Este es el except del bloque try principal (errores de archivo/pdfplumber)
    except Exception as e:
        # Retorna una fila de error si el archivo no puede ser procesado
        return {
            "CLIENT": f"ERROR: No se pudo procesar - {e}",
            "DATE": "N/A",
            "NUMBER": "N/A",
            "DOLLARS": "N/A",
            "PESOS": "N/A",
            "EUROS": "N/A",
            "DESCRIPTION": "N/A"
        }

# ===============================================
# INTERFAZ STREAMLIT (Lógica de la Aplicación Web)
# ===============================================


def main():
    st.set_page_config(page_title="PDF a Excel Múltiple", layout="wide")
    st.title("📂 Extracción Consolidada de Múltiples PDFs a Excel")
    st.subheader("Paso 1: Cargar Archivos PDF")

    uploaded_pdfs = st.file_uploader(
        "Sube uno o más archivos PDF (Facturas):",
        type=["pdf"],
        accept_multiple_files=True
    )

    if uploaded_pdfs:
        st.success(f"Se cargaron **{len(uploaded_pdfs)}** archivos.")

        if st.button("Procesar y Consolidar en Excel"):

            consolidated_data = []

            with st.spinner(f"Iniciando extracción y consolidación de {len(uploaded_pdfs)} archivos..."):

                # Itera sobre CADA archivo cargado
                for uploaded_pdf in uploaded_pdfs:
                    try:
                        pdf_data = io.BytesIO(uploaded_pdf.getvalue())
                        result = extract_data_from_pdf(pdf_data)

                        # Agrega el nombre del archivo
                        result['FILE_NAME'] = uploaded_pdf.name
                        consolidated_data.append(result)

                    except Exception as e:
                        st.warning(
                            f"No se pudo procesar {uploaded_pdf.name}. Error: {e}")

                        consolidated_data.append({
                            "CLIENT": f"ERROR FATAL: {uploaded_pdf.name}",
                            "DATE": "N/A",
                            "NUMBER": "N/A",
                            "DOLLARS": "N/A",
                            "PESOS": "N/A",
                            "EUROS": "N/A",
                            "DESCRIPTION": "N/A",
                            "FILE_NAME": uploaded_pdf.name
                        })

            # A. Crear el DataFrame final
            column_order = ["FILE_NAME", "CLIENT", "DATE",
                            "NUMBER", "DOLLARS", "PESOS", "EUROS", "DESCRIPTION"]
            df = pd.DataFrame(consolidated_data, columns=column_order)

            st.subheader("✅ Datos Consolidados (Vista Previa)")
            st.dataframe(df, width='stretch')

            # B. Crear el archivo Excel en memoria
            output = io.BytesIO()

            def clean_total(x):
                if isinstance(x, str):
                    if re.match(r'^[\d\.,]+$', x):
                        # Reemplazamos todos los puntos y la última coma por un punto decimal
                        return float(x.replace('.', '').replace(',', '.'))
                    return x
                return x

            df['PESOS'] = df['PESOS'].apply(clean_total)

            # Uso de xlsxwriter (instalado via requirements.txt)
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Datos Facturas')
            output.seek(0)

            # C. Botón de descarga
            st.subheader("⬇️ Archivo Excel Consolidado Generado")
            st.download_button(
                label="Descargar Excel de Facturas",
                data=output.read(),
                file_name=f"Facturas_Consolidadas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_button"
            )

            st.balloons()


if __name__ == "__main__":
    main()
