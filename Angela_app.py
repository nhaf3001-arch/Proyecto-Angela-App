import streamlit as st
import pandas as pd
import pdfplumber
import io  # Para manejar archivos en memoria
import re  # Para usar Expresiones Regulares (Regex)
from datetime import datetime  # Para formatear la fecha
import locale  # Para forzar el idioma español en la fecha

# ===============================================
# FUNCIÓN DE EXTRACCIÓN (Lógica de Negocio)
# ===============================================


def extract_data_from_pdf(pdf_file):
    """Extrae el Nombre, la Fecha, el Número, el Total y la Descripción del PDF."""

    # Intentar establecer el idioma español para manejar el nombre del mes ("Agosto")
    try:
        # Intenta la configuración para Linux/Mac
        locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
    except locale.Error:
        try:
            # Intenta la configuración para Windows
            locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')
        except locale.Error:
            # Si falla, continuará, aunque la fecha podría fallar si el sistema no soporta el idioma.
            pass

    with pdfplumber.open(pdf_file) as pdf:
        first_page = pdf.pages[0]
        # Extraemos el texto completo para las búsquedas
        text = first_page.extract_text()

    # --- LÓGICA DE EXTRACCIÓN CON REGEX CORREGIDA ---

    # 1. CLIENTE (Ej: ENLASA GENERACION CHILE S.A.)
    client_match = re.search(r"SEÑOR\(ES\):\s*([^\n]+)", text)
    extracted_name = client_match.group(
        1).strip() if client_match else "No encontrado"

    # 2. NÚMERO (Ej: 228)
    number_match = re.search(
        r"FACTURA ELECTRONICA\s*Nº(\d+)", text, re.IGNORECASE)
    extracted_number = number_match.group(
        1).strip() if number_match else "No encontrado"

    # 3. FECHA (Ej: 14-08-25)
    # Patrón: día(1 o 2 dígitos) + mes(letras) + año(4 dígitos)
    date_match = re.search(
        r"Fecha Emision:\s*(\d{1,2})\s+de\s+(\w+)\s+del\s+(\d{4})", text, re.IGNORECASE)

    extracted_date = "Error de Formato"
    if date_match:
        try:
            # Reconstruye: "14 de Agosto del 2025"
            date_str = f"{date_match.group(1)} de {date_match.group(2)} del {date_match.group(3)}"
            date_obj = datetime.strptime(date_str, '%d de %B del %Y')
            extracted_date = date_obj.strftime('%d-%m-%y')  # Formato DD-MM-AA
        except Exception:
            extracted_date = "Error de Formato"

    # 4. TOTAL (PESOS) (Ej: 7.725.844)
    # Buscamos el total después de 'TOTAL $ ' y capturamos los dígitos y puntos.
    total_match = re.search(r"TOTAL\s*\$\s*([\d\.]+)", text)
    extracted_total = total_match.group(1) if total_match else "No encontrado"

    # 5. DESCRIPCIÓN (Ej: SV_65000 + CW_DRIV)
    # Buscamos los códigos de servicio
    description_match = re.findall(r"-\s*(\w+)", text)
    extracted_description = " + ".join(
        description_match) if description_match else "No encontrado"

    # --- FIN DE LA LÓGICA DE EXTRACCIÓN ---

    # ESTA ESTRUCTURA DEBE COINCIDIR CON LA TABLA DE SALIDA QUE PEDISTE
    data = [
        {
            "CLIENT": extracted_name,
            "DATE": extracted_date,
            "NUMBER": extracted_number,
            "DOLLARS": "",             # Columna vacía
            "PESOS": extracted_total,  # Total extraído
            "EUROS": "",               # Columna vacía
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
        "Sube el archivo PDF (Factura):",
        type=["pdf"],
        accept_multiple_files=False
    )

    if uploaded_pdf is not None:
        st.success(f"Archivo cargado: **{uploaded_pdf.name}**")

        if st.button("Procesar y Generar Nuevo Excel"):
            st.info("Extrayendo datos y generando archivo...")

            try:
                # Convertimos el archivo cargado a un objeto de memoria
                pdf_data = io.BytesIO(uploaded_pdf.getvalue())

                # A. Extraer datos y crear el DataFrame de Pandas
                extracted_data = extract_data_from_pdf(pdf_data)

                # Usamos el DataFrame para asegurar el orden y las columnas
                df = pd.DataFrame(extracted_data, columns=[
                                  "CLIENT", "DATE", "NUMBER", "DOLLARS", "PESOS", "EUROS", "DESCRIPTION"])

                st.subheader("✅ Datos Extraídos (Vista Previa)")
                st.dataframe(df)  # Mostrar los datos extraídos

                # B. Crear el archivo Excel en memoria
                output = io.BytesIO()

                # Función para limpiar el Total antes de guardarlo en el Excel (quita el punto)
                def clean_total(x):
                    # Solo intenta limpiar si no es una cadena vacía o "No encontrado"
                    if isinstance(x, str) and x.replace('.', '', 1).isdigit():
                        try:
                            # Convierte el string "7.725.844" a número 7725844
                            return float(x.replace('.', ''))
                        except:
                            return x  # Retorna el texto si hay error
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
