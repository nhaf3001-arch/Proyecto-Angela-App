import streamlit as st
import pandas as pd
import pdfplumber
import io  # Para manejar archivos en memoria
import re  # ¡NUEVO! Para usar Expresiones Regulares (Regex)
from datetime import datetime  # ¡NUEVO! Para formatear la fecha
import openpyxl  # ¡NUEVO! Para trabajar con tu plantilla Excel existente

# ===============================================
# FUNCIÓN DE EXTRACCIÓN (Lógica de Negocio)
# ===============================================


def extract_data_from_pdf(pdf_file):
    """Extrae el Nombre, la Fecha, el Número, el Total y la Descripción del PDF."""

    with pdfplumber.open(pdf_file) as pdf:
        first_page = pdf.pages[0]
        text = first_page.extract_text()

    # --- LÓGICA DE EXTRACCIÓN CON REGEX ---

    # 1. Nombre del Cliente (Señor(es): [NOMBRE]...)
    client_match = re.search(r"SEÑOR\(ES\):\s*(.+)\n", text)
    extracted_name = client_match.group(
        1).strip() if client_match else "No encontrado"

    # 2. Número de Factura (Nº228)
    number_match = re.search(
        r"FACTURA ELECTRONICA\s*Nº(\d+)", text, re.IGNORECASE)
    extracted_number = number_match.group(
        1).strip() if number_match else "No encontrado"

    # 3. Fecha de Emisión (Fecha Emision: 14 de Agosto del 2025)
    # Patrón: un dígito o dos, un espacio, varias letras (mes), ' del ', cuatro dígitos (año)
    date_match = re.search(
        r"Fecha Emision:\s*(\d{1,2}\s+\w+\s+del\s+\d{4})", text, re.IGNORECASE)
    date_str = date_match.group(1) if date_match else "No encontrado"

    # Convertir la fecha al formato DD-MM-AA (Ej: "14 de Agosto del 2025" -> "14-08-25")
    try:
        # Importante: El nombre del mes debe estar en el idioma de la configuración local de Python.
        # Si da error, el mes podría necesitar ser traducido. Por ahora, asumimos que funciona.
        date_obj = datetime.strptime(date_str, '%d de %B del %Y')
        extracted_date = date_obj.strftime('%d-%m-%y')
    except:
        extracted_date = "Error de Formato"

    # 4. Total (TOTAL $ 7.725.844)
    # Busca el total después de 'TOTAL $ ' y captura los dígitos y puntos.
    total_match = re.search(r"TOTAL\s*\$\s*([\d\.]+)", text)
    extracted_total = total_match.group(
        1) if total_match else "No encontrado"  # Ya no quitamos el punto aquí

    # 5. Descripción (Buscar las líneas de detalle: SV_65000 y CW_DRIV)
    description_match = re.findall(r"-\s*(\w+)", text)
    extracted_description = " + ".join(
        description_match) if description_match else "No encontrado"

    # --- FIN DE LA LÓGICA DE EXTRACCIÓN ---

    # Preparamos los datos
    data = [
        {
            "Cliente": extracted_name,
            "Fecha": extracted_date,
            "Numero": extracted_number,
            "Total": extracted_total,  # Incluye puntos para el reemplazo
            "Descripcion": extracted_description
        }
    ]

    return data

# ===============================================
# INTERFAZ STREAMLIT (Lógica de la Aplicación Web)
# ===============================================


def main():
    st.set_page_config(page_title="Automatización de PDF a Excel")
    st.title("📄 Automatización de Extracción e Inserción en Excel")
    st.subheader("Primer Paso: Cargar el Archivo PDF y la Plantilla Excel")

    # Contenedores para subir los dos archivos
    col1, col2 = st.columns(2)

    with col1:
        uploaded_pdf = st.file_uploader(
            "Sube el archivo PDF (Factura):",
            type=["pdf"]
        )

    with col2:
        uploaded_excel = st.file_uploader(
            "Sube tu Plantilla Excel (xlsx):",
            type=["xlsx"]
        )

    # Solo procesamos si ambos archivos están cargados
    if uploaded_pdf is not None and uploaded_excel is not None:
        st.success(
            f"Archivos listos. PDF: **{uploaded_pdf.name}**, Excel: **{uploaded_excel.name}**")

        if st.button("Procesar, Insertar Datos y Generar Nuevo Excel"):
            st.info("Procesando la información e insertando datos...")

            try:
                # --- A. Extracción de Datos del PDF ---
                pdf_data = io.BytesIO(uploaded_pdf.getvalue())
                extracted_data = extract_data_from_pdf(pdf_data)

                # Suponemos que solo hay un conjunto de datos (una factura)
                data_to_insert = extracted_data[0]

                # Muestra los datos extraídos para verificación
                st.subheader("✅ Datos Extraídos (Verificación)")
                st.json(data_to_insert)

                # --- B. Carga y Modificación del Excel con openpyxl ---

                # 1. Cargar el libro de trabajo (workbook) desde el archivo subido
                wb = openpyxl.load_workbook(uploaded_excel)
                ws = wb.active  # Seleccionamos la hoja activa (la primera)

                # 2. Encontrar la primera fila vacía para insertar
                insert_row = 15
                # Busca la primera fila donde la Columna C (Cliente) esté vacía
                while ws[f'C{insert_row}'].value is not None:
                    insert_row += 1

                st.info(
                    f"Se insertarán los datos en la Fila: **{insert_row}** de la hoja activa.")

                # 3. Mapear las columnas según tu ejemplo final

                # Columna C: Nombre del Cliente
                # ENLASA GENERACION CHILE S.A
                ws[f'C{insert_row}'] = data_to_insert["Cliente"]

                # Columna D: Fecha
                ws[f'D{insert_row}'] = data_to_insert["Fecha"]  # 14-08-25

                # Columna E: Número de Factura
                ws[f'E{insert_row}'] = data_to_insert["Numero"]  # 228

                # Columna I: Total (PESOS) <-- ¡CORREGIDO!
                # Quitar puntos y convertir a número (entero o flotante) para que Excel lo interprete como valor.
                try:
                    # El valor es "7.725.844". En Chile/España el punto es el separador de miles.
                    # Lo quitamos para obtener 7725844 y que Python lo vea como un número entero.
                    total_str_clean = data_to_insert["Total"].replace('.', '')
                    total_value = int(total_str_clean)
                    ws[f'I{insert_row}'] = total_value
                except ValueError:
                    # Si falla, se escribe el texto tal cual (debería ser raro, pero es un fallback)
                    ws[f'I{insert_row}'] = data_to_insert["Total"]

                # Columna K: Descripción
                # SV_65000 + CW_DRIV
                ws[f'K{insert_row}'] = data_to_insert["Descripcion"]

                # Columna F: Fecha del PO (Dejamos vacío, puedes añadirlo si lo extraes)
                # Columna G: Número del PO (Dejamos vacío)

                # 4. Guardar el Workbook modificado en un buffer de memoria
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)

                # C. Botón de descarga
                st.subheader("⬇️ Listo para Descargar")
                st.download_button(
                    label="Descargar Plantilla Excel ACTUALIZADA",
                    data=output.read(),
                    file_name=f"Plantilla_Actualizada_Factura_{data_to_insert['Numero']}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.balloons()

            except Exception as e:
                st.error(
                    f"Ocurrió un error al procesar o escribir el archivo. Por favor, verifica la integridad de tus archivos. Error: {e}")


if __name__ == "__main__":
    main()
