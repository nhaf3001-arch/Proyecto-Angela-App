import streamlit as st
import pandas as pd
import pdfplumber
import re  # Para usar Expresiones Regulares (Regex)
from datetime import datetime  # Para formatear la fecha
import locale  # Para forzar el idioma español en la fecha
import io
import xlsxwriter


# ===============================================
# FUNCIÓN DE EXTRACCIÓN (Lógica de Negocio)
# ===============================================


def extract_data_from_pdf(pdf_file):
    """Extrae el Nombre, la Fecha, el Número, el Total y la Descripción del PDF."""

    # Intentar establecer el idioma español para manejar el nombre del mes
    try:
        locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
    except locale.Error:
        try:
            locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')
        except locale.Error:
            pass

    # Usamos try/except para manejar errores de archivos no válidos individualmente
    try:
        with pdfplumber.open(pdf_file) as pdf:
            first_page = pdf.pages[0]
            text = first_page.extract_text()

            # Limpieza crítica del texto
            text = text.replace('\n', ' ').replace('\r', ' ')
            text = re.sub(r'\s+', ' ', text).strip()

        # --- LÓGICA DE EXTRACCIÓN CON REGEX (SE MANTIENE LA LÓGICA) ---

        # 1. CLIENTE (SR.(A) NATHALY HAYDEE ALARCON FERRES)
        client_match = re.search(
            r"SR\.\(?A\)?[\s:]*([^\n\r]+?)(?:\s+RUT|[\n\r]|$)", text, re.IGNORECASE)
        extracted_name = client_match.group(
            1).strip() if client_match else "No encontrado"

        # 2. NÚMERO (N° : 134877688)
        number_match = re.search(
            r"N°\s*:\s*(\d+)", text, re.IGNORECASE)
        extracted_number = number_match.group(
            1).strip() if number_match else "No encontrado"

        # 3. FECHA (Fecha de Emisión : 20 de Marzo de 2020 -> 20-03-20)
        date_match = re.search(
            r"Fecha\s+de\s+Emisión\s*:\s*(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})", text, re.IGNORECASE)
        extracted_date = "Error de Formato"
        if date_match:
            try:
                date_str = f"{date_match.group(1)} de {date_match.group(2)} de {date_match.group(3)}"
                date_obj = datetime.strptime(date_str, '%d de %B de %Y')
                extracted_date = date_obj.strftime('%d-%m-%y')
            except Exception:
                extracted_date = "Error de Formato"

        # 4. TOTAL (PESOS) (Total Cuenta Única Telefónica $ 20.586)
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
    st.set_page_config(page_title="PDF a Excel Múltiple")
    st.title("📂 Extracción Consolidada de Múltiples PDFs a Excel")
    st.subheader("Paso 1: Cargar Archivos PDF")

    # ⚠️ CAMBIO CLAVE: Cambiar a accept_multiple_files=True
    uploaded_pdfs = st.file_uploader(
        "Sube uno o más archivos PDF (Facturas):",
        type=["pdf"],
        accept_multiple_files=True
    )

    if uploaded_pdfs:
        st.success(f"Se cargaron **{len(uploaded_pdfs)}** archivos.")

        if st.button("Procesar y Consolidar en Excel"):
            st.info(
                f"Iniciando extracción y consolidación de {len(uploaded_pdfs)} archivos...")

            # Lista para almacenar los resultados de TODAS las facturas
            consolidated_data = []

            # ⚠️ CAMBIO CLAVE: Iterar sobre cada archivo cargado
            for uploaded_pdf in uploaded_pdfs:
                try:
                    # Convertimos el archivo cargado a un objeto de memoria
                    pdf_data = io.BytesIO(uploaded_pdf.getvalue())

                    # Extraer datos del PDF y agregar al resultado consolidado
                    result = extract_data_from_pdf(pdf_data)
                    # Agregamos el nombre del archivo al resultado para referencia
                    result['FILE_NAME'] = uploaded_pdf.name
                    consolidated_data.append(result)

                except Exception as e:
                    st.warning(
                        f"No se pudo procesar {uploaded_pdf.name}. Error: {e}")
                    # Agregar una fila de error explícita si el fallo ocurre fuera de la función
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

            # A. Crear el DataFrame final con todos los datos
            column_order = ["FILE_NAME", "CLIENT", "DATE",
                            "NUMBER", "DOLLARS", "PESOS", "EUROS", "DESCRIPTION"]
            df = pd.DataFrame(consolidated_data, columns=column_order)

            st.subheader("✅ Datos Consolidados (Vista Previa)")
            st.dataframe(df)

            # B. Crear el archivo Excel en memoria
            output = io.BytesIO()

            # Función para limpiar el Total (se mantiene la lógica)
            def clean_total(x):
                if isinstance(x, str):
                    if re.match(r'^[\d\.,]+$', x):
                        # Quitar todos los puntos y reemplazar la última coma por un punto decimal si existe
                        return float(x.replace('.', '').replace(',', '.'))
                    return x
                return x

            # Aplicamos la limpieza a la columna PESOS
            df['PESOS'] = df['PESOS'].apply(clean_total)

            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Datos Facturas')
            output.seek(0)

            # C. Botón de descarga
            st.subheader("⬇️ Archivo Excel Consolidado Generado")
            st.download_button(
                label="Descargar Excel de Facturas",
                data=output.read(),
                file_name=f"Facturas_Consolidadas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.balloons()


if __name__ == "__main__":
    main()
