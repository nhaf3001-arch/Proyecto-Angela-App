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
        # ⚠️ SOLUCIÓN CRÍTICA: Limpiar el texto de caracteres problemáticos
        # 1. Reemplaza saltos de línea y retornos de carro por un solo espacio.
        text = text.replace('\n', ' ').replace('\r', ' ')
        # 2. Reemplaza múltiples espacios por un solo espacio.
        text = re.sub(r'\s+', ' ', text).strip()

        # Nota: El re.sub necesita importar 're', que ya está arriba.

    # --- LÓGICA DE EXTRACCIÓN CON REGEX CORREGIDA ---

    # 1. CLIENTE (Más flexible: busca 'SEÑOR(ES):' y captura la línea siguiente)
    # Usamos un patrón más simple para evitar problemas con saltos de línea inmediatos
    client_match = re.search(
        r"SEÑOR\(ES\):[\s]*([^\n\r]+)", text, re.IGNORECASE)
    extracted_name = client_match.group(
        1).strip() if client_match else "No encontrado"

    # 2. NÚMERO (Busca 'Nº' y captura dígitos, ignorando espacios y mayúsculas)
    number_match = re.search(
        r"Nº[\s]*(\d+)", text, re.IGNORECASE)
    extracted_number = number_match.group(
        1).strip() if number_match else "No encontrado"

    # 3. FECHA (Busca 'Fecha Emision:' y captura el día, mes y año)
    date_match = re.search(
        r"Fecha Emision:[\s]*(\d{1,2})\s+de\s+(\w+)\s+del\s+(\d{4})", text, re.IGNORECASE)

    extracted_date = "Error de Formato"
    if date_match:
        try:
            # Reconstruye la cadena para que datetime la entienda
            date_str = f"{date_match.group(1)} de {date_match.group(2)} del {date_match.group(3)}"
            date_obj = datetime.strptime(date_str, '%d de %B del %Y')
            extracted_date = date_obj.strftime('%d-%m-%y')  # Formato DD-MM-AA
        except Exception:
            extracted_date = "Error de Formato"

    # 4. TOTAL (PESOS) (Busca 'TOTAL $' y captura el número con puntos)
    # Usa [\s\S]*? para capturar cualquier cosa entre TOTAL y el valor, en caso de saltos de línea
    total_match = re.search(r"TOTAL[\s\S]*?\$\s*([\d\.]+)", text)
    extracted_total = total_match.group(1) if total_match else "No encontrado"

    # 5. DESCRIPCIÓN (Busca las líneas de código/descripción SV_65000 y CW_DRIV)
    # Basado en los códigos cortos del detalle de la factura
    description_codes = re.findall(r"-\s*(\w{2,}\_\w{2,})", text)
    extracted_description = " + ".join(
        description_codes) if description_codes else "No encontrado"

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
