import streamlit as st
import cv2
import pytesseract
import pandas as pd
import numpy as np
import re
import io
import json
import os
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- GESTIÓN SEGURA Y PERSISTENTE DE LA API KEY ---
CONFIG_FILE = "config_secret.json"

def cargar_api_key():
    """Carga la API key desde config_secret.json, .env o variable de entorno."""
    key = os.environ.get("GEMINI_API_KEY", "")
    if key and key.strip():
        return key.strip()
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("GEMINI_API_KEY", "").strip()
        except Exception:
            pass
    if os.path.exists(".env"):
        try:
            with open(".env", "r", encoding="utf-8") as f:
                for line in f:
                    if line.startswith("GEMINI_API_KEY="):
                        return line.split("=", 1)[1].strip().strip('"').strip("'")
        except Exception:
            pass
    return ""

def guardar_api_key(nueva_key):
    """Guarda permanentemente la API key en config_secret.json y .env."""
    nueva_key = nueva_key.strip()
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"GEMINI_API_KEY": nueva_key}, f, indent=2)
        with open(".env", "w", encoding="utf-8") as f:
            f.write(f'GEMINI_API_KEY="{nueva_key}"\n')
    except Exception:
        pass

# Ruta a Tesseract en Windows
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

st.set_page_config(
    page_title="Extractor Universal de Tablas a Excel",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Extractor Universal de Planillas y Tablas a Excel")
st.write("Convierte **cualquier foto** (manuscrita o impresa, con cualquier título o número de columnas) en un archivo **Excel profesional y editable**.")

# --- BARRA LATERAL: Configuración y Seguridad ---
api_key_guardada = cargar_api_key()

with st.sidebar:
    st.header("⚙️ Motor de Reconocimiento")
    modo = st.radio(
        "Selecciona el método de extracción:",
        [
            "🤖 IA con Visión (Recomendado)",
            "⚡ Motor Local Tesseract (Offline)"
        ],
        index=0,
        help="La IA con Visión entiende escritura a mano, cualquier título y operaciones matemáticas. Tesseract funciona sin internet para tablas impresas simples."
    )
    
    api_key_activa = api_key_guardada

    if "IA" in modo:
        st.markdown("---")
        st.subheader("🔑 Seguridad de la API Key")
        
        if api_key_guardada:
            st.success("🔒 **Clave API Guardada y Protegida**")
            # Mostrar solo los últimos 4 caracteres para confirmación sin exponerla
            preview_key = f"••••••••••••••••••••••••{api_key_guardada[-4:]}" if len(api_key_guardada) >= 4 else "••••••••••••"
            st.text_input("Estado de la Llave:", value=preview_key, disabled=True, help="Tu llave está guardada de forma segura en el sistema y protegida contra edición accidental.")
            
            with st.expander("⚙️ Reconfigurar / Cambiar Clave"):
                nueva_llave = st.text_input("Ingresar nueva clave:", type="password", key="input_nueva_llave")
                if st.button("💾 Actualizar y Guardar Nueva Clave"):
                    if nueva_llave.strip():
                        guardar_api_key(nueva_llave)
                        st.success("¡Clave actualizada con éxito! Recargando...")
                        st.rerun()
        else:
            st.warning("⚠️ No hay ninguna Clave API guardada.")
            clave_ingresada = st.text_input(
                "Ingresa tu Google Gemini API Key:",
                type="password",
                help="Obtén tu clave gratis en https://aistudio.google.com"
            )
            if st.button("💾 Guardar y Proteger Clave Permanentemente"):
                if clave_ingresada.strip():
                    guardar_api_key(clave_ingresada)
                    st.success("¡Clave guardada y protegida con éxito!")
                    st.rerun()
                else:
                    st.error("Por favor escribe una clave válida.")
            api_key_activa = clave_ingresada

st.markdown("---")

# Subir archivo
col_izq, col_der = st.columns([1, 1])

with col_izq:
    archivo_subido = st.file_uploader("📂 Sube la foto del documento (JPG, PNG)", type=["jpg", "jpeg", "png"])
    nombre_archivo = st.text_input("Nombre para el archivo Excel (opcional):", value="Reporte_Digitalizado")

if archivo_subido is not None:
    with col_der:
        st.image(archivo_subido, caption="Vista previa del documento", use_container_width=True)

# --- FUNCIONES DE EXTRACCIÓN ---

def extraer_con_ia_vision(bytes_imagen, api_key):
    """Extrae cualquier tabla (manuscrita o impresa) usando Gemini Vision y devuelve un DataFrame estructurado."""
    from google import genai
    from google.genai import types

    client = genai.Client(api_key=api_key)
    
    prompt = """
    Eres un sistema experto en digitalización de tablas, planillas y cuadernos de campo (tanto manuscritos como impresos).
    Analiza la imagen minuciosamente y extrae toda la información tabular y las anotaciones adicionales.

    Instrucciones específicas:
    1. Detecta automáticamente todos los encabezados y columnas de la tabla.
    2. Si hay dos tablas paralelas lado a lado con un índice común (por ejemplo V1 a V26 con Entrada y Drenaje), únelas en una sola fila por índice con todas sus columnas (ej: V, CE_Entrada, PH_Entrada, V_Aforo_Entrada, V_R, CE_Drenaje, PH_Drenaje, V_Aforo_Drenaje, Porcentaje, etc.).
    3. Conserva números decimales, porcentajes, operaciones escritas (como '90+10=100') y valores especiales como asteriscos (*).
    4. Identifica el título temático del documento y fechas si existen (ej. '01/09/26', 'Ref Calcio', etc.).
    5. NOTAS Y EXCEDENTES AL PIE: Identifica minuciosamente todas las anotaciones, mediciones adicionales, excedentes, fórmulas o concentraciones escritas abajo o en los márgenes (ej: 'V10 = 18 conc 402sc', 'V4: 3.0 6.7 (21)', 'Ref Calcio...', etc.). Desglósalas ordenadamente en un arreglo de strings, una por cada línea lógica o bloque relevante.

    Debes responder EXCLUSIVAMENTE con un JSON válido con la siguiente estructura exacta:
    {
      "titulo": "Título descriptivo del reporte",
      "fecha": "Fecha detectada o vacía",
      "columnas": ["Columna1", "Columna2", "Columna3", ...],
      "filas": [
        ["val1", "val2", "val3", ...],
        ["val1", "val2", "val3", ...]
      ],
      "notas": [
        "Línea 1 de medición adicional / excedente",
        "Línea 2 de concentración o referencia"
      ]
    }
    """

    # 1. Obtener dinámicamente los modelos disponibles para la API Key
    modelos_disponibles = []
    try:
        for m in client.models.list():
            nombre = getattr(m, 'name', '') or getattr(m, 'model', '')
            if nombre:
                nombre_limpio = nombre.replace('models/', '')
                if 'gemini' in nombre_limpio.lower() and 'embed' not in nombre_limpio.lower():
                    if 'flash' in nombre_limpio.lower():
                        modelos_disponibles.insert(0, nombre_limpio)
                    else:
                        modelos_disponibles.append(nombre_limpio)
    except Exception:
        pass

    # Si la lista dinámica falló o no tiene elementos, usar lista estándar de fallback
    if not modelos_disponibles:
        modelos_disponibles = [
            'gemini-2.5-flash',
            'gemini-2.0-flash',
            'gemini-3.6-flash',
            'gemini-1.5-flash-latest',
            'gemini-1.5-flash',
            'gemini-1.5-pro'
        ]

    response_text = None
    ultimo_error = None

    # Intentar con google.genai
    for nombre_modelo in modelos_disponibles:
        try:
            resp = client.models.generate_content(
                model=nombre_modelo,
                contents=[
                    types.Part.from_bytes(
                        data=bytes_imagen,
                        mime_type='image/jpeg'
                    ),
                    prompt
                ],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.1
                )
            )
            if resp and resp.text:
                response_text = resp.text
                break
        except Exception as e:
            ultimo_error = e
            continue

    # Fallback con google.generativeai si google.genai dio error
    if not response_text:
        try:
            import google.generativeai as gai
            gai.configure(api_key=api_key)
            for m_gai in ['gemini-1.5-flash', 'gemini-1.5-pro', 'gemini-2.0-flash']:
                try:
                    modelo_gai = gai.GenerativeModel(m_gai)
                    resp_gai = modelo_gai.generate_content(
                        contents=[
                            {'mime_type': 'image/jpeg', 'data': bytes_imagen},
                            prompt
                        ]
                    )
                    if resp_gai and resp_gai.text:
                        response_text = resp_gai.text
                        break
                except Exception as e_gai:
                    ultimo_error = e_gai
                    continue
        except Exception:
            pass

    if not response_text:
        raise Exception(f"No se pudo consultar el modelo de Gemini. Detalle: {ultimo_error}")

    # Limpiar formato Markdown ```json si aparece
    texto_limpio = response_text.strip()
    if texto_limpio.startswith("```"):
        texto_limpio = re.sub(r'^```(?:json)?\s*', '', texto_limpio)
        texto_limpio = re.sub(r'\s*```$', '', texto_limpio)

    data = json.loads(texto_limpio)
    titulo = data.get("titulo", "REPORTE DIGITALIZADO")
    fecha = data.get("fecha", "")
    columnas = data.get("columnas", [])
    filas = data.get("filas", [])
    notas = data.get("notas", "")

    df = pd.DataFrame(filas, columns=columnas)
    return df, titulo, fecha, notas

def enderezar_imagen(gris):
    _, thresh = cv2.threshold(gris, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    coords = np.column_stack(np.where(thresh > 0))
    if len(coords) > 0:
        angle = cv2.minAreaRect(coords)[-1]
        if angle < -45:
            angle = -(90 + angle)
        elif angle > 45:
            angle = 90 - angle
        else:
            angle = -angle
        if 0.5 < abs(angle) < 15:
            (h, w) = gris.shape[:2]
            centro = (w // 2, h // 2)
            M = cv2.getRotationMatrix2D(centro, angle, 1.0)
            gris = cv2.warpAffine(gris, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
    return gris

def extraer_con_tesseract(bytes_imagen):
    """Extractor local optimizado para listas impresas (Tesseract)."""
    nparr = np.frombuffer(bytes_imagen, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)

    gris = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gris = enderezar_imagen(gris)
    gris_2x = cv2.resize(gris, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_CUBIC)
    _, thresh = cv2.threshold(gris_2x, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    config = r'--oem 3 --psm 6 -l spa'
    texto_crudo = pytesseract.image_to_string(thresh, config=config)

    registros = []
    ultimo_num = 0

    for linea in texto_crudo.splitlines():
        linea = linea.strip()
        if not linea or "LOCKER" in linea.upper():
            continue

        linea_limpia = re.sub(r'[\[\]\|\_\—\-\~\{\}\:\;\¿\¡\°\º\(\)\<\>\"\=\'\?\*\+\#\$\%]', ' ', linea)
        linea_limpia = " ".join(linea_limpia.split())

        if len(linea_limpia) < 4:
            continue

        nums = re.findall(r'\b\d{1,3}\b', linea_limpia)
        num_elegido = None
        for n_str in nums:
            n = int(n_str)
            candidatos = [n]
            if n > 50:
                candidatos.extend([n // 10, n % 100])
                if len(n_str) >= 2:
                    candidatos.append(int(n_str[:2]))
            for cand in candidatos:
                if ultimo_num < cand <= ultimo_num + 3:
                    num_elegido = cand
                    break
            if num_elegido is not None:
                break

        if num_elegido is None:
            num_elegido = ultimo_num + 1

        m_letras = re.search(r'[A-ZÁÉÍÓÚÑa-záéíóúñ]{3,}.*', linea_limpia)
        if not m_letras:
            continue
        nombre = m_letras.group(0)
        nombre = re.sub(r'^(ENTER|\d+)\s+', '', nombre, flags=re.IGNORECASE).strip()
        nombre = re.sub(r'^[a-zA-Z]{1,2}\s+', '', nombre).strip()
        nombre = re.sub(r'\b(NO|OK|O|AA|TN|WD|TS|OS)\b$', '', nombre, flags=re.IGNORECASE).strip()
        nombre = re.sub(r'[\d\W]+$', '', nombre).strip()
        nombre = re.sub(r'\s+[A-Z0-9]{1,2}$', '', nombre).strip()
        nombre = " ".join(nombre.split()).upper()

        palabras = nombre.split()
        if len(nombre) >= 4 and len(palabras) >= 2 and any(c.isalpha() for c in nombre):
            ultimo_num = num_elegido
            registros.append({
                "N°": num_elegido,
                "DESCRIPCIÓN / NOMBRE": nombre
            })

    df = pd.DataFrame(registros)
    if not df.empty:
        df = df.drop_duplicates(subset=["N°"]).sort_values(by="N°")
    return df, "CONTROL Y ASIGNACIÓN", "", ""

# --- GENERADOR DE EXCEL PROFESIONAL ADAPTABLE A CUALQUIER TABLA ---

def generar_excel_estilizado(df: pd.DataFrame, titulo: str, fecha: str = "", notas=None) -> io.BytesIO:
    """Genera un archivo Excel profesional adaptándose a cualquier cantidad y tipo de columnas con sección destacada de excedentes/notas."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Datos", startrow=3)
        ws = writer.sheets["Datos"]
        ws.views.sheetView[0].showGridLines = True

        num_cols = max(len(df.columns), 4)
        col_fin = get_column_letter(len(df.columns))
        col_fin_notas = get_column_letter(num_cols)

        # 1. Título principal
        ws.merge_cells(f"A1:{col_fin}1")
        ws["A1"] = titulo.upper()
        ws["A1"].font = Font(name="Segoe UI", size=14, bold=True, color="1F4E78")
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30

        # 2. Subtítulo con fecha y total
        fecha_texto = f"Fecha: {fecha}  |  " if fecha else ""
        fecha_gen = datetime.now().strftime("%d/%m/%Y %H:%M")
        ws.merge_cells(f"A2:{col_fin}2")
        ws["A2"] = f"{fecha_texto}Total Registros: {len(df)}  |  Generado: {fecha_gen}"
        ws["A2"].font = Font(name="Segoe UI", size=10, italic=True, color="595959")
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 20

        # 3. Encabezados de columna
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")

        for col_num in range(1, len(df.columns) + 1):
            cell = ws.cell(row=4, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[4].height = 28

        # 4. Bordes y estilo cebra
        thin_border = Border(
            left=Side(style='thin', color="D9D9D9"),
            right=Side(style='thin', color="D9D9D9"),
            top=Side(style='thin', color="D9D9D9"),
            bottom=Side(style='thin', color="D9D9D9")
        )
        zebra_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
        white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

        start_row = 5
        for row_idx, row in enumerate(ws.iter_rows(min_row=start_row, max_row=start_row + len(df) - 1, min_col=1, max_col=len(df.columns))):
            ws.row_dimensions[row[0].row].height = 22
            es_par = (row_idx % 2 == 1)
            for col_idx, cell in enumerate(row):
                cell.border = thin_border
                cell.font = Font(name="Segoe UI", size=10)
                cell.fill = zebra_fill if es_par else white_fill
                
                # Formato inteligente
                val_str = str(cell.value or "")
                if len(val_str) <= 8 or col_idx == 0:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

        # 5. Autoajustar ancho de columnas
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row >= 4 and cell.row < start_row + len(df) and cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max(max_len + 5, 14)

        # 6. SECCIÓN DESTACADA DE EXCEDENTES Y NOTAS (Grande, ordenada y elegante)
        # Normalizar notas a lista de strings
        lista_notas = []
        if isinstance(notas, list):
            lista_notas = [str(n).strip() for n in notas if str(n).strip()]
        elif isinstance(notas, str) and notas.strip():
            lista_notas = [line.strip() for line in notas.splitlines() if line.strip()]

        if lista_notas:
            fila_separador = start_row + len(df) + 1
            ws.row_dimensions[fila_separador].height = 14
            
            fila_encabezado_notas = fila_separador + 1
            ws.merge_cells(f"A{fila_encabezado_notas}:{col_fin_notas}{fila_encabezado_notas}")
            ws[f"A{fila_encabezado_notas}"] = "📌 MEDICIONES ADICIONALES, EXCEDENTES Y OBSERVACIONES"
            ws[f"A{fila_encabezado_notas}"].font = Font(name="Segoe UI", size=12, bold=True, color="FFFFFF")
            ws[f"A{fila_encabezado_notas}"].fill = PatternFill(start_color="2A4D69", end_color="2A4D69", fill_type="solid")
            ws[f"A{fila_encabezado_notas}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws.row_dimensions[fila_encabezado_notas].height = 28

            fila_actual = fila_encabezado_notas + 1
            nota_fill_par = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
            nota_fill_impar = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

            for idx, nota_texto in enumerate(lista_notas):
                ws.merge_cells(f"A{fila_actual}:{col_fin_notas}{fila_actual}")
                celda = ws[f"A{fila_actual}"]
                celda.value = f"  •   {nota_texto}"
                celda.font = Font(name="Segoe UI", size=11, color="1E293B", bold=False)
                celda.fill = nota_fill_par if idx % 2 == 0 else nota_fill_impar
                celda.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                celda.border = thin_border
                ws.row_dimensions[fila_actual].height = 26
                fila_actual += 1

    buffer.seek(0)
    return buffer

# --- ACCIÓN PRINCIPAL ---

if archivo_subido is not None:
    if st.button("🚀 Extraer Datos y Generar Excel", type="primary"):
        bytes_foto = archivo_subido.read()
        
        with st.spinner("Procesando y reconociendo estructura de la tabla y excedentes..."):
            try:
                if "IA" in modo:
                    if not api_key_activa:
                        st.error("⚠️ Para usar el modo IA por favor ingresa tu API Key en la barra lateral izquierda.")
                        st.stop()
                    df_resultado, titulo_doc, fecha_doc, notas_doc = extraer_con_ia_vision(bytes_foto, api_key_activa)
                else:
                    df_resultado, titulo_doc, fecha_doc, notas_doc = extraer_con_tesseract(bytes_foto)

                if not df_resultado.empty:
                    st.success(f"¡Se detectaron exitosamente **{len(df_resultado)} filas** y **{len(df_resultado.columns)} columnas**!")
                    
                    st.subheader(f"📋 {titulo_doc}")
                    if fecha_doc:
                        st.caption(f"📅 Fecha detectada: **{fecha_doc}**")

                    # Tabla interactiva editable por el usuario
                    st.markdown("##### 📊 Tabla Principal de Datos")
                    st.caption("💡 *Puedes hacer doble clic en cualquier celda para corregirla antes de descargar:*")
                    df_editado = st.data_editor(df_resultado, use_container_width=True, height=400)

                    # Sección destacada de Excedentes y Notas
                    notas_formateadas = []
                    if isinstance(notas_doc, list):
                        notas_formateadas = notas_doc
                    elif isinstance(notas_doc, str) and notas_doc.strip():
                        notas_formateadas = [n.strip() for n in notas_doc.splitlines() if n.strip()]

                    st.markdown("---")
                    st.markdown("##### 📌 Mediciones Adicionales, Excedentes y Observaciones")
                    texto_notas_default = "\n".join([f"• {n}" if not n.startswith("•") else n for n in notas_formateadas])
                    
                    notas_editadas = st.text_area(
                        "Edita o añade notas y excedentes (cada línea aparecerá como un punto en el Excel):",
                        value=texto_notas_default,
                        height=140,
                        help="Cada renglón que escribas aquí se incluirá formateado en grande y ordenado en el archivo Excel final."
                    )

                    notas_finales = [n.replace("•", "").strip() for n in notas_editadas.splitlines() if n.strip()]

                    # Generar Excel con diseño profesional
                    buffer_excel = generar_excel_estilizado(df_editado, titulo_doc, fecha_doc, notas_finales)
                    nombre_final = f"{nombre_archivo.strip() or 'Reporte'}.xlsx"

                    st.download_button(
                        label="📥 Descargar archivo Excel Profesional (con Excedentes)",
                        data=buffer_excel,
                        file_name=nombre_final,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("No se pudieron reconocer datos en la imagen. Verifica que la foto sea legible.")
            except Exception as e:
                st.error(f"Ocurrió un error al procesar la imagen: {e}")
