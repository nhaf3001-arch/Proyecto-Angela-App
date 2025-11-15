📂 Extracción Consolidada de Facturas (PDF a Excel)

Esta aplicación Streamlit permite a los usuarios cargar múltiples archivos PDF de facturas telefónicas, extraer automáticamente datos clave de cada documento (Cliente, Fecha, Número, Total) y consolidar toda la información en un único archivo Excel para su fácil gestión y análisis.

🌟 Características

Procesamiento por Lotes: Permite subir múltiples archivos PDF a la vez.

Extracción de Datos: Utiliza expresiones regulares (regex) para extraer campos específicos como el nombre del cliente, número de factura, fecha de emisión y el total en pesos chilenos.

Compatibilidad Dual (Local/Cloud): El manejo de fechas está optimizado para funcionar correctamente tanto en entornos de desarrollo local (con configuraciones regionales en español) como en plataformas de despliegue en la nube como Streamlit Cloud (que utilizan configuraciones regionales en inglés).

Salida Consolidada: Genera un archivo Excel (.xlsx) con una fila por cada PDF procesado.

🛠️ Instalación y Requisitos

Para ejecutar la aplicación localmente, necesitas tener Python instalado.

1. Requisitos de Python

Asegúrate de tener instaladas las bibliotecas necesarias. Los requisitos se encuentran en el archivo requirements.txt.

# Asegúrate de tener Python instalado (versión 3.8+)
python -m venv venv
source venv/bin/activate  # En Linux/macOS
venv\Scripts\activate     # En Windows

# Instalar dependencias
pip install -r requirements.txt



2. Archivos del Proyecto

El proyecto se compone de los siguientes archivos principales:

Archivo

Descripción

Angela_app.py

El código fuente de la aplicación Streamlit y la lógica de extracción.

requirements.txt

Lista de dependencias de Python necesarias.

README.md

Este archivo de documentación.

🚀 Uso de la Aplicación

Ejecutar la Aplicación: Abre tu terminal, activa tu entorno virtual y ejecuta el siguiente comando:

streamlit run Angela_app.py



Esto abrirá la aplicación en tu navegador web.

Cargar PDFs: En la interfaz de Streamlit, haz clic en el botón para subir archivos. Selecciona todos los archivos PDF de facturas que deseas procesar.

Procesar: Haz clic en el botón "Procesar y Consolidar en Excel". La aplicación iterará sobre cada archivo cargado, intentará extraer los datos y mostrará una vista previa en la tabla de Datos Consolidados.

Descargar: Una vez completado el procesamiento, haz clic en el botón "Descargar Excel de Facturas" para obtener el archivo Facturas_Consolidadas_YYYYMMDD_HHMMSS.xlsx con todos los datos.

⚙️ Lógica de Extracción (Regex)

La función extract_data_from_pdf utiliza las siguientes expresiones regulares para identificar los campos en los documentos:

Campo

Patrón Regex

Descripción

CLIENT

`r"SR.$?A$?[\s:]*([^\n\r]+?)(?:\s+RUT

[\n\r]

NUMBER

r"N°\s*:\s*(\d+)"

Busca la secuencia de dígitos después de "N° :".

DATE

`r"Fecha\s+(?:de\s+)?Emisi[óo]n\s*:\s*(\d{1,2})\s+de\s+(\w+)\s+(?:del

de)\s+(\d{4})"`

PESOS

r"Total\s+Cuenta\s+Única\s+Telefónica\s+\$\s*([\d\.,]+)"

Captura el valor numérico (incluyendo puntos y comas) asociado al total.

Cualquier PDF que no siga la estructura esperada para estos campos será marcado con "No encontrado" o "ERROR".