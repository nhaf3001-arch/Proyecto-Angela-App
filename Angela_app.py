import streamlit as st
import pandas as pd
import pdfplumber
import re  # Para usar Expresiones Regulares (Regex)
from datetime import datetime  # Para formatear la fecha
import locale  # Para forzar el idioma español en la fecha
import io
import xlsxwriter
# Nueva librería para leer archivos Word (.docx)
import docx

# ⚠️ CONFIGURACIÓN GLOBAL (Mapeo de meses y Locale)
# Se mantiene fuera de la clase ya que son constantes de configuración
MONTH_MAPPING = {
    'enero': 'January', 'febrero': 'February', 'marzo': 'March',
    'abril': 'April', 'mayo': 'May', 'junio': 'June',
    'julio': 'July', 'agosto': 'August', 'septiembre': 'September',
    'octubre': 'October', 'noviembre': 'November', 'diciembre': 'December'
}

# Reglas de Extracción Centralizadas (para PDF y DOCX)
EXTRACTION_RULES = {
    "CLIENT": [
        # Regla 1 (MÁXIMA PRECISIÓN para Razón Social):
        r"(?:SEÑOR\s*\(?ES\)?\s*:\s*)([^\n\r]+?)(?=\s*(?:R\.?U\.?T\.|GIRO|DIRECCI[ÓO]N|FECHA|COMUNA|[\n\r]|$))",
        # Regla 2 (Fallback si no hay R.U.T. cerca): Busca SR(A): NOMBRE...
        r"(?:SR\.\(?A\)?[\s:]*)([^\n\r]+?)(?:\s+RUT|[\n\r]|$)",
        # Regla 3 (Flexible): Fallback por si no tiene prefijo formal
        r"(?:SR\.\(?A\)?|Hola|Estimado\s*:\s*)?([^\n\r]+?)(?:\s+RUT|[\n\r]|$)"
    ],
    "NUMBER": [r"N[°º]\s*:\s*(\d+)", r"N[°º]\s*(\d+)", r"N°\s*:\s*(\d+)"],
    "DATE": [
        {"regex": r"Fecha\s+(?:de\s+)?Emisi[óo]n\s*:\s*(\d{1,2})\s+de\s+(\w+)\s+(?:del|de)\s+(\d{4})",
         "format": "LONG_FORMAT"},
        {"regex": r"Fecha\s*:\s*(\d{1,2})[\s\-\/](\d{1,2})[\s\-\/](\d{2,4})",
         "format": "DD_MM_YY"}
    ],
    "TOTAL": [r"TOTAL\s+\$\s*([\d\.,]+)", r"Total\s+Cuenta\s+Única\s+Telefónica\s+\$\s*([\d\.,]+)"],
    "DESCRIPTION": [r"(BOLETA\s+ELECTRONICA)", r"(GUIA\s+DE\s+DESPACHO\s+ELECTRONICA)", r"([A-Z0-9]{2,}[-][A-Z0-9]{2,})", r"\b([A-Z]{3,}\d{2,})\b", r"(SII[^\n\r]+SANTIAGO)"]
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
# FUNCIONES AUXILIARES DE EXTRACCIÓN Y LIMPIEZA
# ===============================================

def _find_client_in_text(text, rules):
    """ Busca el nombre del cliente usando las reglas de FacturaExtractor. """
    patterns = rules.get("CLIENT", [])
    for pattern in patterns:
        search_flags = re.IGNORECASE
        match = re.search(pattern, text, search_flags)
        if match:
            result = match.group(1).strip() if len(match.groups()) > 0 else ""
            # --- LIMPIEZA CRÍTICA ---
            result = re.sub(
                r"^(SEÑOR\s*\(?ES\)?\s*:\s*|SR\.\(?A\)?[\s:]*)", "", result, flags=re.IGNORECASE).strip()
            result = re.sub(r"\s*R\.?U\.?T\..*$", "",
                            result, flags=re.IGNORECASE).strip()
            result = result.replace(':', '').strip()
            return result
    return "No encontrado"


# ===============================================
# CLASE DE EXTRACCIÓN PDF
# ===============================================

class FacturaExtractor:
    """ Encapsula la lógica y las reglas de extracción para un tipo de documento PDF. """

    def __init__(self, pdf_file):
        """Inicializa el extractor leyendo y limpiando el texto del PDF."""

        try:
            with pdfplumber.open(pdf_file) as pdf:
                text = "".join(page.extract_text() for page in pdf.pages)
                text = text.replace('\n', ' ').replace('\r', ' ')
                self.text = re.sub(r'\s+', ' ', text).strip()
        except Exception as e:
            self.text = ""
            st.warning(f"Error al cargar texto del PDF: {e}")

    def _parse_date(self, date_match, date_format_type):
        """ Método privado para parsear la fecha basándose en el tipo de formato. """
        extracted_date = "Error de Formato (Parseo)"
        if date_format_type == "LONG_FORMAT":
            try:
                day = date_match.group(1)
                month_es = date_match.group(2)
                year = date_match.group(3)
                try:
                    date_str = f"{day} de {month_es} de {year}"
                    date_obj = datetime.strptime(date_str, '%d de %B de %Y')
                except ValueError:
                    month_es_lower = month_es.lower()
                    month_en = MONTH_MAPPING.get(month_es_lower, month_es)
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
                if len(year) == 2:
                    year = f"20{year}"
                date_str = f"{day}-{month}-{year}"
                date_obj = datetime.strptime(date_str, '%d-%m-%Y')
                extracted_date = date_obj.strftime('%d-%m-%y')
            except Exception:
                extracted_date = "Error de Formato (Corto Fallido)"
        return extracted_date

    def _try_find(self, field_name):
        """ Método privado que prueba secuencialmente los patrones para un campo. """
        patterns = EXTRACTION_RULES.get(field_name, [])
        for pattern in patterns:
            if isinstance(pattern, dict):
                regex = pattern.get("regex")
            else:
                regex = pattern

            search_flags = re.IGNORECASE if field_name not in [
                "DESCRIPTION"] else 0
            match = re.search(regex, self.text, search_flags)

            if match:
                result = match.group(1).strip() if len(
                    match.groups()) > 0 else ""
                # La limpieza crítica de CLIENTE ahora se hace en _find_client_in_text.
                return result, match, pattern
        return "No encontrado", None, None

    def extract_all(self):
        """Método principal que ejecuta todas las extracciones."""

        # 1. CLIENTE (Usando la función externa)
        extracted_name = _find_client_in_text(self.text, EXTRACTION_RULES)

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
        extracted_description, _, _ = self._try_find("DESCRIPTION")

        if extracted_description == "No encontrado":
            extracted_description = "Documento Genérico (Default)"

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
# FUNCIÓN DE EXTRACCIÓN DOCX
# ===============================================

def extract_data_from_docx(docx_file):
    """
    Extrae el número y fecha de cotización de un archivo DOCX.
    """
    # Patrón: COTIZACIÓN # <CÓDIGO>/<TEXTO>, <FECHA>
    # Simplificado, ya no necesitamos extraer CLIENT para la fusión forzada.
    QUOTE_PATTERN = r"COTIZACI[ÓO]N\s*#\s*([A-Z0-9]+)\/?[A-Z]*,\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})"

    extracted_quotation = {
        "QUOTATION_NUMBER": "No encontrado",
        "QUOTATION_DATE": "No encontrado",
        # El campo CLIENT ya no se extrae ni se necesita para la fusión secuencial.
    }

    try:
        # Load the document from the in-memory BytesIO object
        document = docx.Document(docx_file)

        # Leemos todo el texto del documento
        full_text = []
        for paragraph in document.paragraphs:
            full_text.append(paragraph.text)
        # Unir y limpiar el texto
        text = " ".join(full_text)
        text = re.sub(r'\s+', ' ', text).strip()

        # 1. Extracción del Número y Fecha de Cotización
        match = re.search(QUOTE_PATTERN, text, re.IGNORECASE)
        if match:
            quotation_number = match.group(1).strip()
            quotation_date_raw = match.group(2).strip()
            formatted_date = quotation_date_raw

            # --- NUEVA LÓGICA DE LIMPIEZA ---
            # Eliminar el prefijo "CB" si existe, seguido de dígitos.
            quotation_number = re.sub(
                r"^CB", "", quotation_number, flags=re.IGNORECASE).strip()

            try:  # Intentar formatear la fecha
                date_obj = datetime.strptime(quotation_date_raw, '%d/%m/%Y')
                formatted_date = date_obj.strftime('%d-%m-%y')
            except ValueError:
                try:
                    date_obj = datetime.strptime(
                        quotation_date_raw, '%d-%m-%Y')
                    formatted_date = date_obj.strftime('%d-%m-%y')
                except ValueError:
                    try:
                        date_obj = datetime.strptime(
                            quotation_date_raw, '%d/%m/%y')
                        formatted_date = date_obj.strftime('%d-%m-%y')
                    except ValueError:
                        pass  # Si falla, se deja la cadena original.

            extracted_quotation["QUOTATION_NUMBER"] = quotation_number
            extracted_quotation["QUOTATION_DATE"] = formatted_date

        return extracted_quotation

    except ImportError:
        st.error("Error: La librería 'python-docx' (import docx) no está instalada. Es necesaria para procesar archivos Word.")
        return extracted_quotation
    except Exception as e:
        st.warning(
            f"Error al procesar el archivo DOCX: {docx_file.name}. Detalles: {e}")
        return extracted_quotation


# ===============================================
# INTERFAZ STREAMLIT (Lógica de la Aplicación Web)
# ===============================================


def extract_data_from_pdf(pdf_file):
    """ Función wrapper para la extracción de PDF. """
    try:
        extractor = FacturaExtractor(pdf_file)
        return extractor.extract_all()
    except Exception as e:
        return {
            "CLIENT": f"ERROR: No se pudo procesar - {e}",
            "DATE": "N/A", "NUMBER": "N/A", "DOLLARS": "N/A",
            "PESOS": "N/A", "EUROS": "N/A", "DESCRIPTION": "N/A"
        }


def main():
    st.set_page_config(page_title="PDF y DOCX a Excel Múltiple", layout="wide")
    st.title("📂 Extracción Consolidada de Facturas y Cotizaciones a Excel")

    # === UPLOADERS ===
    st.subheader("Paso 1: Cargar Facturas (PDF)")
    uploaded_pdfs = st.file_uploader(
        "Sube uno o más archivos PDF (Facturas):",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_uploader"
    )

    st.subheader("Paso 2: Cargar Cotizaciones (DOCX)")
    uploaded_docs = st.file_uploader(
        "Sube uno o más archivos Word (.docx) (Se fusionarán por orden de carga con los PDFs):",
        type=["docx"],
        accept_multiple_files=True,
        key="docx_uploader"
    )

    # === PROCESAMIENTO ===
    if uploaded_pdfs or uploaded_docs:
        if st.button("Procesar y Consolidar en Excel", type="primary"):

            # Almacenamiento consolidado. La clave es el CLIENTE
            all_data = {}

            # NUEVO: Lista para mantener el orden de las claves de los clientes
            # (El orden de los PDFs define el orden de las filas)
            pdf_client_keys = []

            # --- 1. PROCESAR PDFs (Fuente principal de filas) ---
            if uploaded_pdfs:
                with st.spinner(f"Iniciando extracción de {len(uploaded_pdfs)} Facturas (PDF)..."):
                    for uploaded_pdf in uploaded_pdfs:
                        try:
                            pdf_data = io.BytesIO(uploaded_pdf.getvalue())
                            result = extract_data_from_pdf(pdf_data)

                            # ⚠️ CLAVE DE FUSIÓN: Nombre del Cliente (normalizado)
                            merge_key = result["CLIENT"].upper().strip()

                            # Solo agregamos si se encontró el cliente en el PDF
                            if merge_key != "NO ENCONTRADO" and merge_key != "":

                                # Si ya existe el cliente (duplicado), le agregamos un sufijo para que sea único
                                unique_merge_key = merge_key
                                count = 1
                                while unique_merge_key in all_data:
                                    unique_merge_key = f"{merge_key}_{count}"
                                    count += 1

                                # Guardar la clave única para usarla en la fusión DOCX
                                pdf_client_keys.append(unique_merge_key)

                                # Añadimos placeholders para las columnas de cotización
                                placeholder_quote_data = {
                                    "QUOTATION_NUMBER": "No DOCX adjunto",
                                    "QUOTATION_DATE": "No DOCX adjunto"
                                }
                                all_data[unique_merge_key] = {
                                    **result, **placeholder_quote_data, "FILE_NAME": uploaded_pdf.name}
                            else:
                                st.warning(
                                    f"PDF ignorado: {uploaded_pdf.name}. No se pudo extraer el CLIENTE para generar la fila.")

                        except Exception as e:
                            st.warning(
                                f"Error en PDF {uploaded_pdf.name}: {e}")

            # --- 2. PROCESAR DOCX (Fusión SECUENCIAL forzada) ---
            if uploaded_docs:
                with st.spinner(f"Iniciando extracción de {len(uploaded_docs)} Cotizaciones (DOCX) y fusionando por orden..."):

                    # Iteramos sobre los DOCXs, y usamos el índice para obtener la clave del PDF correspondiente
                    # Si hay 3 PDFs y 2 DOCXs, solo se fusionan los 2 primeros PDFs.
                    num_docs_to_process = min(
                        len(uploaded_docs), len(pdf_client_keys))

                    for i in range(num_docs_to_process):
                        uploaded_doc = uploaded_docs[i]
                        # Obtenemos la clave del PDF correspondiente
                        pdf_key_to_update = pdf_client_keys[i]

                        try:
                            doc_data = io.BytesIO(uploaded_doc.getvalue())
                            quote_result = extract_data_from_docx(doc_data)

                            # ¡FUSIÓN EXITOSA FORZADA! Actualizamos la fila del PDF usando la clave secuencial
                            all_data[pdf_key_to_update].update(quote_result)

                            original_client_name = all_data[pdf_key_to_update]['CLIENT']
                            st.success(
                                f"DOCX fusionado (Secuencial): {uploaded_doc.name} se consolidó con la fila del PDF '{original_client_name}'.")

                        except KeyError:
                            # Esto no debería pasar si pdf_key_to_update está en pdf_client_keys
                            st.error(
                                f"Error interno: La clave '{pdf_key_to_update}' no existe en all_data.")
                        except Exception as e:
                            st.warning(
                                f"Error en DOCX {uploaded_doc.name} (Fallo Secuencial): {e}")

                    if len(uploaded_docs) > len(pdf_client_keys):
                        st.info(
                            f"Se ignoraron {len(uploaded_docs) - len(pdf_client_keys)} DOCXs porque no había más PDFs para fusionar.")

            # --- 3. CONSOLIDAR DATAFRAME ---

            # Mapear el diccionario de resultados a una lista, manteniendo el orden de las claves.
            consolidated_data = [all_data[key] for key in pdf_client_keys]

            # A. Crear el DataFrame final
            column_order = [
                "FILE_NAME", "CLIENT", "DATE", "NUMBER",
                "QUOTATION_NUMBER", "QUOTATION_DATE",
                "DOLLARS", "PESOS", "EUROS", "DESCRIPTION"
            ]
            df = pd.DataFrame(consolidated_data, columns=column_order)

            # Limpiar claves de sufijos si se duplicaron
            df['CLIENT'] = df['CLIENT'].apply(lambda x: x.split('_')[0])

            st.subheader("✅ Datos Consolidados (Vista Previa)")
            st.dataframe(df, width='stretch')

            # B. Crear el archivo Excel en memoria
            output = io.BytesIO()

            def clean_total(x):
                if isinstance(x, str):
                    if x in ["No encontrado", "N/A", "Documento Genérico (Default)", "No DOCX adjunto"]:
                        return x
                    # Remueve el punto como separador de miles
                    cleaned_x = x.replace('.', '')
                    # Reemplaza la coma por punto para decimales (formato float)
                    cleaned_x = cleaned_x.replace(',', '.')
                    try:
                        return float(cleaned_x)
                    except ValueError:
                        return x
                # Para números directos, los devuelve tal cual
                return x

            # Aplicar limpieza a la columna PESOS
            df['PESOS'] = df['PESOS'].apply(clean_total)

            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False,
                            sheet_name='Datos Consolidación')
            output.seek(0)

            # C. Botón de descarga
            st.subheader("⬇️ Archivo Excel Consolidado Generado")
            st.download_button(
                label="Descargar Excel Consolidado",
                data=output.read(),
                file_name=f"Consolidado_Facturas_Cotizaciones_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_button"
            )
            st.balloons()


if __name__ == "__main__":
    main()
