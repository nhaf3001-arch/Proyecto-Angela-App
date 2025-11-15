import streamlit as st
import pandas as pd
import pdfplumber
import re  # Para usar Expresiones Regulares (Regex)
from datetime import datetime  # Para formatear la fecha
import locale  # Para forzar el idioma español en la fecha
import io
import xlsxwriter

# ⚠️ CONFIGURACIÓN GLOBAL (Mapeo de meses y Locale)
# Se mantiene fuera de la clase ya que son constantes de configuración
MONTH_MAPPING = {
    'enero': 'January', 'febrero': 'February', 'marzo': 'March',
    'abril': 'April', 'mayo': 'May', 'junio': 'June',
    'julio': 'July', 'agosto': 'August', 'septiembre': 'September',
    'octubre': 'October', 'noviembre': 'November', 'diciembre': 'December'
}

# Se intenta configurar el locale.
try:
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')
    except locale.Error:
        pass


# ===============================================
# CLASE DE EXTRACCIÓN (PROGRAMACIÓN ORIENTADA A OBJETOS)
# ===============================================

class FacturaExtractor:
    """
    Encapsula la lógica y las reglas de extracción para un tipo de documento.
    """

    # REGLAS DE EXTRACCIÓN: Ahora son atributos de la clase.
    EXTRACTION_RULES = {
        "CLIENT": [
            # Regla 1 (Original): Busca SR(A): NOMBRE...
            r"SR\.\(?A\)?[\s:]*([^\n\r]+?)(?:\s+RUT|[\n\r]|$)",
            # Regla 2 (Flexible): Busca cualquier nombre después de "Hola" o sin prefijo formal
            r"(?:SR\.\(?A\)?|Hola|Estimado\s*:\s*)?([^\n\r]+?)(?:\s+RUT|[\n\r]|$)"
        ],

        "NUMBER": [
            # Regla 1 (Única): Busca N°: 12345
            r"N°\s*:\s*(\d+)"
        ],

        "DATE": [
            # Regla A (Original/Larga): 10 de Febrero de 2020
            {"regex": r"Fecha\s+(?:de\s+)?Emisi[óo]n\s*:\s*(\d{1,2})\s+de\s+(\w+)\s+(?:del|de)\s+(\d{4})",
             "format": "LONG_FORMAT"},
            # Regla B (Nueva/Corta): 10-02-20 o 10/02/2020
            {"regex": r"Fecha\s*:\s*(\d{1,2})[\s\-\/](\d{1,2})[\s\-\/](\d{2,4})",
             "format": "DD_MM_YY"}
        ],

        "TOTAL": [
            # Regla 1 (Única): Busca Total Cuenta Única Telefónica $ 123.456
            r"Total\s+Cuenta\s+Única\s+Telefónica\s+\$\s*([\d\.,]+)"
        ]
    }

    def __init__(self, pdf_file):
        """Inicializa el extractor leyendo y limpiando el texto del PDF."""
        try:
            with pdfplumber.open(pdf_file) as pdf:
                first_page = pdf.pages[0]
                text = first_page.extract_text()

                # Limpieza crítica del texto
                text = text.replace('\n', ' ').replace('\r', ' ')
                self.text = re.sub(r'\s+', ' ', text).strip()
        except Exception as e:
            self.text = ""
            st.warning(f"Error al cargar texto del PDF: {e}")

    def _parse_date(self, date_match, date_format_type):
        """
        Método privado para parsear la fecha basándose en el tipo de formato.
        Utiliza el mapeo global MONTH_MAPPING.
        """
        extracted_date = "Error de Formato (Parseo)"

        if date_format_type == "LONG_FORMAT":
            try:
                day = date_match.group(1)
                month_es = date_match.group(2)
                year = date_match.group(3)

                # Intento con locale y fallback con mapeo
                try:
                    date_str = f"{day} de {month_es} de {year}"
                    date_obj = datetime.strptime(date_str, '%d de %B de %Y')
                except ValueError:
                    month_es_lower = month_es.lower()
                    month_en = MONTH_MAPPING.get(month_es_lower, month_es)
                    date_str = f"{day} de {month_en} de {year}"
                    date_obj = datetime.strptime(date_str, '%d de %B de %Y')

                extracted_date = date_obj.strftime('%d-%m-%y')

            except Exception:
                extracted_date = "Error de Formato (Largo Fallido)"

        elif date_format_type == "DD_MM_YY":
            try:
                day = date_match.group(1).zfill(2)
                month = date_match.group(2).zfill(2)
                year = date_match.group(3)
                # Asegurar año de 4 dígitos si viene de 2
                if len(year) == 2:
                    year = f"20{year}"

                date_str = f"{day}-{month}-{year}"
                date_obj = datetime.strptime(date_str, '%d-%m-%Y')
                extracted_date = date_obj.strftime('%d-%m-%y')

            except Exception:
                extracted_date = "Error de Formato (Corto Fallido)"

        return extracted_date

    def _try_find(self, field_name):
        """
        Método privado que prueba secuencialmente los patrones para un campo.
        """
        patterns = self.EXTRACTION_RULES.get(field_name, [])

        for pattern in patterns:
            if isinstance(pattern, dict):
                # Para reglas complejas como la Fecha
                regex = pattern.get("regex")
            else:
                # Para reglas sencillas (Cliente, Número, Total)
                regex = pattern

            # Buscamos en el texto limpio del PDF
            match = re.search(regex, self.text, re.IGNORECASE)
            if match:
                # Si el patrón es simple, devolvemos el grupo 1, el objeto match y el patrón.
                result = match.group(1).strip() if len(
                    match.groups()) > 0 else ""
                return result, match, pattern

        # Si no se encuentra ninguna coincidencia
        return "No encontrado", None, None

    def extract_all(self):
        """Método principal que ejecuta todas las extracciones."""

        # 1. CLIENTE
        extracted_name, _, _ = self._try_find("CLIENT")

        # 2. NÚMERO
        extracted_number, _, _ = self._try_find("NUMBER")

        # 3. FECHA
        extracted_date = "No encontrado"
        _, date_match, date_rule = self._try_find("DATE")

        if date_match and date_rule:
            extracted_date = self._parse_date(date_match, date_rule["format"])

        # 4. TOTAL
        extracted_total, _, _ = self._try_find("TOTAL")

        # 5. DESCRIPCIÓN
        extracted_description = "Factura Telefónica"

        # Retorna el diccionario de resultados
        return {
            "CLIENT": extracted_name,
            "DATE": extracted_date,
            "NUMBER": extracted_number,
            "DOLLARS": "",
            "PESOS": extracted_total,
            "EUROS": "",
            "DESCRIPTION": extracted_description
        }


# ===============================================
# FUNCIÓN DE ENTRADA (Wrapper)
# ===============================================

def extract_data_from_pdf(pdf_file):
    """
    Función de entrada que crea una instancia del extractor
    y llama a su método principal para obtener los datos.
    """
    try:
        extractor = FacturaExtractor(pdf_file)
        return extractor.extract_all()

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
                        # Llama a la función wrapper, que ahora usa la clase OOP
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
                    # Maneja el caso de "No encontrado"
                    if x in ["No encontrado", "N/A"]:
                        return x

                    # Reemplazamos todos los puntos y la última coma por un punto decimal
                    cleaned_x = x.replace('.', '')
                    cleaned_x = cleaned_x.replace(',', '.')

                    try:
                        return float(cleaned_x)
                    except ValueError:
                        return x

                return x

            df['PESOS'] = df['PESOS'].apply(clean_total)

            # Uso de xlsxwriter
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
