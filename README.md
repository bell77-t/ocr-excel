# 📊 Extractor Universal de Planillas y Tablas a Excel (OCR + IA Visión)

Una aplicación web desarrollada con **Streamlit**, **OpenCV**, **Tesseract OCR**, **Google Gemini Vision API** y **OpenPyXL** para digitalizar cualquier tipo de planilla, tabla impresa o cuaderno de campo manuscrito, transformándolo en un archivo Excel profesional, ordenado y estilizado.

---

## 🚀 Características Principales

* **🤖 Modo Inteligente con Visión por IA (Universal):**
  * Lee **escritura manuscrita** (a mano) y texto impreso.
  * Detecta automáticamente cualquier título de documento y columnas variables (Riego, Agronomía, Asistencias, Lockers, Inventarios, etc.).
  * Reconoce decimales, operaciones matemáticas (ej. `90+10=100`), porcentajes y fechas.
  * Extrae y organiza de forma grande y destacada las **Mediciones Adicionales, Excedentes y Observaciones**.
* **⚡ Modo Local Offline (Tesseract OCR):**
  * Procesamiento rápido y sin internet para listas tabulares estándar.
  * Corrección automática de inclinación (*Deskewing*) y binarización adaptativa Otsu.
* **✨ Excel con Diseño Ejecutivo y Profesional:**
  * Cabecera institucional con fecha y total de registros.
  * Encabezados corporativos en azul marino (`#1F4E78`) con texto blanco en negrita.
  * Filas con efecto cebra (`#F2F5F9` / `#FFFFFF`) para fácil lectura.
  * Bordes suaves y autoajuste inteligente del ancho de columnas.
  * Sección destacada y amplia para anotaciones y excedentes.
* **🔒 Seguridad y Persistencia de API Key:**
  * Guardado local seguro para no tener que ingresar la llave en cada uso.
  * Bloqueo visual en la interfaz para protegerla contra edición accidental.

---

## 🛠️ Requisitos e Instalación

### 1. Clonar el repositorio
```bash
git clone https://github.com/TU_USUARIO/TU_REPOSITORIO.git
cd TU_REPOSITORIO
```

### 2. Instalar dependencias de Python
```bash
pip install -r requirements.txt
```

### 3. (Opcional para modo offline) Instalar Tesseract OCR
Si deseas usar el modo local sin internet, descarga e instala [Tesseract OCR para Windows](https://github.com/UB-Mannheim/tesseract/wiki) en la ruta por defecto:
`C:\Program Files\Tesseract-OCR\tesseract.exe`

---

## ▶️ Uso de la Aplicación

## 🔑 Configuración de la API Key de Gemini

Tienes dos formas muy sencillas de configurar tu API Key:

1. **Directamente en el código:** Abre `app.py` y en la primera línea pega tu clave:
   ```python
   MI_API_KEY_DIRECTA = "AIzaSy..."  # <-- Pega tu llave aquí
   ```
2. **Desde la interfaz web:** Al abrir la aplicación en tu navegador, ingresa tu API Key en la barra lateral izquierda y presiona **"Guardar y Proteger Clave"**.

---

## 📂 Estructura Limpia del Proyecto

```text
ocr-excel/
├── app.py              # Aplicación principal Web (Streamlit + IA Visión + Decisión de Riego)
├── main.py             # Script local para procesamiento rápido por consola
├── Hoja.jpeg           # Imagen de prueba (Planilla de ejemplo)
├── requirements.txt    # Dependencias del proyecto
├── README.md           # Documentación del proyecto
├── .gitignore          # Filtro de seguridad para archivos privados (.env, claves, temporales)
└── config_secret.json  # Almacén local seguro de tu API Key (no se sube a GitHub)
```

---

## 📄 Licencia
Distribuido bajo la Licencia MIT.
