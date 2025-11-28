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
    # REGLAS DE EXTRACCIÓN: Priorizando las reglas del SII (Documentos Chile)
    EXTRACTION_RULES = {
        "CLIENT": [
            # Regla 1 (MÁXIMA PRECISIÓN para Razón Social):
            # Busca SEÑOR(ES): y captura CUALQUIER COSA de forma perezosa ([^\n\r]+?),
            # hasta que ve R.U.T., GIRO, DIRECCIÓN o FECHA.
            r"(?:SEÑOR\s*\(?ES\)?\s*:\s*)([^\n\r]+?)(?=\s*(?:R\.?U\.?T\.|GIRO|DIRECCI[ÓO]N|FECHA|COMUNA|[\n\r]|$))",
            # Regla 2 (Fallback si no hay R.U.T. cerca): Busca SR(A): NOMBRE...
            r"(?:SR\.\(?A\)?[\s:]*)([^\n\r]+?)(?:\s+RUT|[\n\r]|$)",
            # Regla 3 (Flexible): Fallback por si no tiene prefijo formal
            r"(?:SR\.\(?A\)?|Hola|Estimado\s*:\s*)?([^\n\r]+?)(?:\s+RUT|[\n\r]|$)"
        ],

        "NUMBER": [
            # Regla 1 (SII - Prioridad): Busca N° o Nº seguido del número (e.g., Nº27).
            r"N[°º]\s*:\s*(\d+)",
            r"N[°º]\s*(\d+)",
            # Regla 2 (Original): Busca N°: 12345
            r"N°\s*:\s*(\d+)"
        ],

        "DATE": [
            # Regla A (Original/Larga - SII Compatible): Fecha Emision: 09 de Octubre del 2025
            {"regex": r"Fecha\s+(?:de\s+)?Emisi[óo]n\s*:\s*(\d{1,2})\s+de\s+(\w+)\s+(?:del|de)\s+(\d{4})",
             "format": "LONG_FORMAT"},
            # Regla B (Nueva/Corta): 10-02-20 o 10/02/2020
            {"regex": r"Fecha\s*:\s*(\d{1,2})[\s\-\/](\d{1,2})[\s\-\/](\d{2,4})",
             "format": "DD_MM_YY"}
        ],

        "TOTAL": [
            # Regla 1 (SII - Prioridad): Busca el TOTAL del documento (e.g., TOTAL $ 14.280.000)
            r"TOTAL\s+\$\s*([\d\.,]+)",
            # Regla 2 (Original): Busca Total Cuenta Única Telefónica $ 123.456
            r"Total\s+Cuenta\s+Única\s+Telefónica\s+\$\s*([\d\.,]+)"
        ],

        "DESCRIPTION": [
            # --- REGLAS MODIFICADAS: PRIORIZANDO EL TIPO DE DOCUMENTO ---
            # Regla 1 (MÁXIMA PRIORIDAD - Boleta/Factura Electrónica):
            # Busca explícitamente la etiqueta "BOLETA ELECTRONICA" o "FACTURA ELECTRONICA"
            r"(BOLETA\s+ELECTRONICA)",
            # Regla 2 (Prioridad Media - Guía de Despacho):
            r"(GUIA\s+DE\s+DESPACHO\s+ELECTRONICA)",
            # Regla 3 (Fallback - Código de Producto, e.g., SAT-DUST):
            # Busca códigos alfanuméricos con guion.
            r"([A-Z0-9]{2,}[-][A-Z0-9]{2,})",
            # Regla 4 (Fallback - Código/Texto Corto): Busca palabras clave que parezcan SKU (e.g., SATDUST)
            r"\b([A-Z]{3,}\d{2,})\b",
            # Regla 5 (Fallback genérico del SII, como el que se estaba capturando antes):
            r"(SII[^\n\r]+SANTIAGO)",
            # --- FIN REGLAS MODIFICADAS ---
        ]
    }

    def __init__(self, pdf_file):
        """Inicializa el extractor leyendo y limpiando el texto del PDF."""

        try:
            with pdfplumber.open(pdf_file) as pdf:
                # Extraer texto de todas las páginas para mayor robustez
                text = "".join(page.extract_text() for page in pdf.pages)
                # Limpieza crítica del texto
                # Se eliminan saltos de línea y se reduce el espacio múltiple a un solo espacio.
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
                    # Fallback usando el mapeo de meses a inglés
                    month_es_lower = month_es.lower()
                    month_en = MONTH_MAPPING.get(month_es_lower, month_es)
                    # Usamos 'of' y luego intentamos parsear con la versión en inglés del mes
                    date_str = f"{day} of {month_en} of {year}"
                    date_obj = datetime.strptime(date_str, '%d of %B of %Y')
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
                # Para reglas sencillas (Cliente, Número, Total, Descripción)
                regex = pattern
            # Buscamos en el texto limpio del PDF
            search_flags = 0
            # Se usa re.IGNORECASE para todos, excepto para DESCRIPTION (donde BOLETA/GUIA deben ser precisos)
            if field_name not in ["DESCRIPTION"]:
                search_flags = re.IGNORECASE
            match = re.search(regex, self.text, search_flags)
            if match:
                # Si el patrón es simple, devolvemos el grupo 1, el objeto match y el patrón.
                result = match.group(1).strip() if len(
                    match.groups()) > 0 else ""
                # --- LIMPIEZA CRÍTICA PARA CLIENTE ---
                if field_name == "CLIENT":
                    # 1. Eliminar explícitamente cualquier prefijo de cortesía/etiqueta
                    # Esto garantiza que el nombre quede solo.
                    result = re.sub(
                        r"^(SEÑOR\s*\(?ES\)?\s*:\s*|SR\.\(?A\)?[\s:]*)", "", result, flags=re.IGNORECASE).strip()
                    # 2. Eliminar cualquier R.U.T. o texto que se haya colado al final,
                    # usando una detención en R.U.T. si la regex falló.
                    result = re.sub(r"\s*R\.?U\.?T\..*$", "",
                                    result, flags=re.IGNORECASE).strip()
                    # 3. Limpieza de cualquier carácter residual (como dos puntos o espacios al final)
                    result = result.replace(':', '').strip()
                return result, match, pattern
        # Si no se encuentra ninguna coincidencia
        return "No encontrado", None, None

    def extract_all(self):
        """Método principal que ejecuta todas las extracciones."""

        # 1. CLIENTE (Limpieza adicional en _try_find)
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
        # Se extrae el primer patrón encontrado (que ahora prioriza el tipo de documento)
        extracted_description, _, _ = self._try_find("DESCRIPTION")

        # Fallback si no encuentra ninguna de las etiquetas
        if extracted_description == "No encontrado":
            extracted_description = "Documento Genérico (Default)"
            pass

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
                        if x in ["No encontrado", "N/A", "Documento Genérico (Default)"]:
                            return x
                        # Elimina separadores de miles (puntos)
                        cleaned_x = x.replace('.', '')
                        # Convierte la coma a punto decimal (aunque se espera que solo haya puntos para miles)
                        cleaned_x = cleaned_x.replace(',', '.')
                        try:
                            # Intentar convertir a float
                            return float(cleaned_x)
                        except ValueError:
                            # Si falla la conversión, devolver el valor original como string
                            return x
                    return x
                # Limpiamos solo la columna PESOS para convertirla a número
                df['PESOS'] = df['PESOS'].apply(clean_total)
                # Uso de xlsxwriter
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False,
                                sheet_name='Datos Facturas')
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
