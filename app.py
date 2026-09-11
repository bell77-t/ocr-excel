# ==============================================================================
# 🔑 PEGA TU API KEY DE GOOGLE GEMINI AQUÍ DIRECTAMENTE:
# ==============================================================================
MI_API_KEY_DIRECTA = ""  # <-- PEGA TU LLAVE AQUÍ (Ejemplo: "AIzaSy...")
# ==============================================================================

# Intentar cargar desde el archivo config_api_key.py si existe
try:
    from config_api_key import GEMINI_API_KEY as KEY_FROM_CONFIG
except Exception:
    KEY_FROM_CONFIG = ""

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
    """Carga la API key con la siguiente prioridad:
    1. Variable directa en este archivo (MI_API_KEY_DIRECTA).
    2. Archivo config_api_key.py (GEMINI_API_KEY).
    3. Variable de entorno del sistema (GEMINI_API_KEY).
    4. Archivo config_secret.json.
    5. Archivo .env.
    """
    # 1. Directa en app.py
    if MI_API_KEY_DIRECTA and MI_API_KEY_DIRECTA.strip():
        return MI_API_KEY_DIRECTA.strip()

    # 2. Desde config_api_key.py
    if KEY_FROM_CONFIG and KEY_FROM_CONFIG.strip():
        return KEY_FROM_CONFIG.strip()

    # 3. Variable de entorno
    key = os.environ.get("GEMINI_API_KEY", "")
    if key and key.strip():
        return key.strip()

    # 3. Archivo config_secret.json
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                val = data.get("GEMINI_API_KEY", "").strip()
                if val:
                    return val
        except Exception:
            pass

    # 4. Archivo .env
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

# --- MAPEADOR AUTOMÁTICO A DECISIÓN DE RIEGO ---

def mapear_decision_riego(df_foto: pd.DataFrame, fecha_str: str = "") -> pd.DataFrame:
    """Transforma los datos leídos de la foto (Entrada, Drenaje, Válvulas) en el formato oficial de DECISIÓN DE RIEGO."""
    filas_dec = []
    
    col_map = {re.sub(r'[^A-Z0-9%]', '', str(c).upper()): c for c in df_foto.columns}

    # Detectar columnas clave
    c_val = next((col_map[k] for k in col_map if 'V' in k and ('VAL' in k or k == 'V' or (k.startswith('V') and not 'AFORO' in k))), df_foto.columns[0] if len(df_foto.columns) > 0 else None)
    c_ce_ent = next((col_map[k] for k in col_map if 'CE' in k and ('ENT' in k or not 'DREN' in k)), None)
    c_ce_dren = next((col_map[k] for k in col_map if 'CE' in k and 'DREN' in k), c_ce_ent)
    c_ph_ent = next((col_map[k] for k in col_map if 'PH' in k and ('ENT' in k or not 'DREN' in k)), None)
    c_ph_dren = next((col_map[k] for k in col_map if 'PH' in k and 'DREN' in k), c_ph_ent)
    c_vaforo = next((col_map[k] for k in col_map if 'VAFORO' in k or 'AFORO' in k or 'VOLEJ' in k), None)
    c_vaforo_dren = next((col_map[k] for k in col_map if 'AFORO' in k and 'DREN' in k), None)
    c_dr = next((col_map[k] for k in col_map if '%' in k or 'DREN' in k or 'PORC' in k or 'DR' in k), None)

    # Calcular número de semana estimado
    semana_val = "36"
    if fecha_str:
        nums = re.findall(r'\b\d{1,4}\b', fecha_str)
        if len(nums) >= 2:
            try:
                dia, mes = int(nums[0]), int(nums[1])
                semana_val = str(min(52, max(1, (mes - 1) * 4 + dia // 7 + 1)))
            except Exception:
                pass

    for idx, row in df_foto.iterrows():
        v_raw = str(row[c_val] if c_val in row else idx + 1).strip()
        m_v = re.search(r'\d+', v_raw)
        num_val = int(m_v.group(0)) if m_v else (idx + 1)

        # CE y pH (priorizar medición de drenaje si existe para semáforo agronómico)
        ce_val = row[c_ce_dren] if (c_ce_dren and c_ce_dren in row) else ""
        if ce_val == "" or ce_val == "*" or pd.isna(ce_val):
            ce_val = row[c_ce_ent] if (c_ce_ent and c_ce_ent in row) else "*"

        ph_val = row[c_ph_dren] if (c_ph_dren and c_ph_dren in row) else ""
        if ph_val == "" or ph_val == "*" or pd.isna(ph_val):
            ph_val = row[c_ph_ent] if (c_ph_ent and c_ph_ent in row) else "*"

        # Convertir CE y pH a float si es posible
        try:
            ce_val = float(str(ce_val).replace(',', '.'))
        except Exception:
            pass

        try:
            ph_val = float(str(ph_val).replace(',', '.'))
        except Exception:
            pass

        # Vol Ejecutado
        vol_ej_val = row[c_vaforo] if (c_vaforo and c_vaforo in row) else ""
        try:
            vol_ej_val = float(str(vol_ej_val).replace(',', '.'))
        except Exception:
            pass

        # % Drenaje
        dr_val = row[c_dr] if (c_dr and c_dr in row) else ""
        if isinstance(dr_val, str) and "%" in dr_val:
            try:
                dr_num = float(dr_val.replace("%", "").replace(",", ".").strip()) / 100.0
            except Exception:
                dr_num = dr_val
        else:
            try:
                dr_num = float(str(dr_val).replace(',', '.'))
                if dr_num > 1.0 and dr_num <= 100.0:
                    dr_num = dr_num / 100.0
            except Exception:
                dr_num = dr_val

        # Observaciones
        obs = ""
        if c_vaforo_dren and c_vaforo_dren in row and str(row[c_vaforo_dren]).strip() not in ["", "*", "nan"]:
            obs = f"Aforo drenaje: {str(row[c_vaforo_dren]).strip()}"

        filas_dec.append({
            "SEMANA": semana_val,
            "FECHA": fecha_str or datetime.now().strftime("%Y-%m-%d"),
            "PRODUCTO": "ST" if num_val > 15 else "GB",
            "SECTOR": f"SECTOR {(num_val - 1) // 6 + 1}",
            "BLOQ.": str((num_val - 1) % 6 + 1),
            "VAL": num_val,
            "VOL_PROG_LT": 160 if num_val > 15 else 180,
            "#Puls": 3 if num_val > 15 else 4,
            "VOL.EJ": vol_ej_val,
            "CE": ce_val,
            "PH": ph_val,
            "%DR": dr_num,
            "OBSERVACIONES": obs
        })

    return pd.DataFrame(filas_dec)

# --- GENERADOR MULTI-PESTAÑA DE EXCEL CON SEMÁFOROS Y FUNCIONES ---

def generar_excel_estilizado(df_foto: pd.DataFrame, titulo: str, fecha: str = "", notas=None) -> io.BytesIO:
    """Genera un archivo Excel profesional completo con pestañas DECISION_RIEGO, DASHBOARD, PARAMETROS y DATOS_FOTO."""
    buffer = io.BytesIO()
    
    # 1. Mapear datos a Decisión de Riego
    df_decision = mapear_decision_riego(df_foto, fecha)

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        # -------------------------------------------------------------
        # HOJA 1: DECISION_RIEGO (La pestaña principal con semáforos)
        # -------------------------------------------------------------
        wb = writer.book
        ws_dec = wb.create_sheet(title="DECISION_RIEGO")
        ws_dec.views.sheetView[0].showGridLines = True

        # Estilos comunes
        thin_border = Border(
            left=Side(style='thin', color="D9D9D9"),
            right=Side(style='thin', color="D9D9D9"),
            top=Side(style='thin', color="D9D9D9"),
            bottom=Side(style='thin', color="D9D9D9")
        )

        # Fila 1: Barra de navegación
        ws_dec["A1"] = '📊 Dashboard'
        ws_dec["C1"] = '🌱 Decisión Riego'
        ws_dec["F1"] = '📸 Datos Foto'
        ws_dec["I1"] = '⚙️ Parámetros Nutrición'
        for nav_col in ["A1", "C1", "F1", "I1"]:
            ws_dec[nav_col].font = Font(name="Segoe UI", size=9, bold=True, color="1F4E78")

        # Fila 3 y 4: Tarjetas KPI Superiores
        kpis = [
            ("B3", "B4", "Total Válvulas", f"=COUNTA(A7:A{6 + len(df_decision)})", "1F4E78"),
            ("E3", "E4", "Vol. Total Ejecutado (LT)", f"=SUM(K7:K{6 + len(df_decision)})", "107C41"),
            ("H3", "H4", "Promedio pH", f"=AVERAGE(M7:M{6 + len(df_decision)})", "8E44AD"),
            ("K3", "K4", "Promedio CE (dS/m)", f"=AVERAGE(L7:L{6 + len(df_decision)})", "D35400"),
            ("N3", "N4", "Promedio % Drenaje", f"=AVERAGE(N7:N{6 + len(df_decision)})", "2980B9")
        ]

        for title_cell, val_cell, kpi_title, kpi_formula, color_hex in kpis:
            # Título KPI
            ws_dec[title_cell] = kpi_title
            ws_dec[title_cell].font = Font(name="Segoe UI", size=8, bold=True, color="595959")
            ws_dec[title_cell].alignment = Alignment(horizontal="center", vertical="center")
            
            # Valor KPI
            ws_dec[val_cell] = kpi_formula
            ws_dec[val_cell].font = Font(name="Segoe UI", size=13, bold=True, color=color_hex)
            ws_dec[val_cell].alignment = Alignment(horizontal="center", vertical="center")
            ws_dec[val_cell].fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
            ws_dec[val_cell].border = thin_border
            if "%" in kpi_title:
                ws_dec[val_cell].number_format = "0.0%"
            elif "Vol" in kpi_title or "Total" in kpi_title:
                ws_dec[val_cell].number_format = "#,##0"
            else:
                ws_dec[val_cell].number_format = "0.00"

        ws_dec.row_dimensions[3].height = 16
        ws_dec.row_dimensions[4].height = 24

        # Fila 6: Encabezados de la Tabla
        headers_dec = [
            "SEMANA", "FECHA", "PRODUCTO", "SECTOR", "BLOQ.", "VAL", 
            "VOL_PROG_LT", "#Puls", "Lt/Pul", "Vol. Esperado", "VOL.EJ", 
            "CE", "PH", "%DR", "DE. RIEGO", "10% MAS", "10% MENOS", "OBSERVACIONES"
        ]

        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")

        for col_idx, h_text in enumerate(headers_dec, 1):
            c = ws_dec.cell(row=6, column=col_idx, value=h_text)
            c.fill = header_fill
            c.font = header_font
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws_dec.row_dimensions[6].height = 28

        # Filas 7+: Datos y Fórmulas
        start_row = 7
        for i, r_data in df_decision.iterrows():
            curr_row = start_row + i
            ws_dec.row_dimensions[curr_row].height = 20
            
            val_num = r_data["VAL"]
            vol_prog = r_data["VOL_PROG_LT"]
            pulsos = r_data["#Puls"]
            vol_ej = r_data["VOL.EJ"]
            ce_val = r_data["CE"]
            ph_val = r_data["PH"]
            dr_val = r_data["%DR"]

            # Escribir valores
            ws_dec.cell(row=curr_row, column=1, value=r_data["SEMANA"]).alignment = Alignment(horizontal="center")
            ws_dec.cell(row=curr_row, column=2, value=r_data["FECHA"]).alignment = Alignment(horizontal="center")
            ws_dec.cell(row=curr_row, column=3, value=r_data["PRODUCTO"]).alignment = Alignment(horizontal="center")
            ws_dec.cell(row=curr_row, column=4, value=r_data["SECTOR"]).alignment = Alignment(horizontal="center")
            ws_dec.cell(row=curr_row, column=5, value=r_data["BLOQ."]).alignment = Alignment(horizontal="center")
            ws_dec.cell(row=curr_row, column=6, value=val_num).alignment = Alignment(horizontal="center")
            ws_dec.cell(row=curr_row, column=7, value=vol_prog).alignment = Alignment(horizontal="center")
            ws_dec.cell(row=curr_row, column=8, value=pulsos).alignment = Alignment(horizontal="center")
            
            # Fórmulas
            ws_dec.cell(row=curr_row, column=9, value=f"=IF(H{curr_row}>0, G{curr_row}/H{curr_row}, 0)").alignment = Alignment(horizontal="center")
            ws_dec.cell(row=curr_row, column=10, value=f"=G{curr_row}").alignment = Alignment(horizontal="center")
            
            # VOL.EJ
            c_volej = ws_dec.cell(row=curr_row, column=11, value=vol_ej)
            c_volej.alignment = Alignment(horizontal="center")
            
            # CE
            c_ce = ws_dec.cell(row=curr_row, column=12, value=ce_val)
            c_ce.alignment = Alignment(horizontal="center")
            if isinstance(ce_val, (int, float)):
                c_ce.number_format = "0.00"

            # PH
            c_ph = ws_dec.cell(row=curr_row, column=13, value=ph_val)
            c_ph.alignment = Alignment(horizontal="center")
            if isinstance(ph_val, (int, float)):
                c_ph.number_format = "0.0"

            # %DR
            c_dr = ws_dec.cell(row=curr_row, column=14, value=dr_val)
            c_dr.alignment = Alignment(horizontal="center")
            if isinstance(dr_val, (int, float)):
                c_dr.number_format = "0.0%"

            # DE. RIEGO (Pulsos de Decisión Inteligente)
            ws_dec.cell(row=curr_row, column=15, value=f'=IF(N{curr_row}="*", "*", IF(N{curr_row}<0.2, 4, IF(N{curr_row}>0.45, 2, 3)))').alignment = Alignment(horizontal="center")
            
            # 10% MAS y 10% MENOS
            ws_dec.cell(row=curr_row, column=16, value=f"=ROUND(G{curr_row}*1.1, 0)").alignment = Alignment(horizontal="center")
            ws_dec.cell(row=curr_row, column=17, value=f"=ROUND(G{curr_row}*0.9, 0)").alignment = Alignment(horizontal="center")
            
            # OBSERVACIONES
            ws_dec.cell(row=curr_row, column=18, value=r_data["OBSERVACIONES"]).alignment = Alignment(horizontal="left")

            # Bordes y fuente
            for col_c in range(1, 19):
                cell_item = ws_dec.cell(row=curr_row, column=col_c)
                cell_item.border = thin_border
                cell_item.font = Font(name="Segoe UI", size=9)

            # SEMÁFOROS (Colores visuales automáticos)
            # 1. Semáforo VOL.EJ vs 10% MAS / 10% MENOS
            if isinstance(vol_ej, (int, float)) and isinstance(vol_prog, (int, float)):
                if vol_ej < vol_prog * 0.9:
                    c_volej.fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid") # Rojo (Déficit)
                elif vol_ej > vol_prog * 1.1:
                    c_volej.fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid") # Azul (Exceso)
                else:
                    c_volej.fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid") # Verde (Óptimo)

            # 2. Semáforo %DR (20% - 45% Óptimo)
            if isinstance(dr_val, (int, float)):
                if dr_val < 0.20:
                    c_dr.fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid") # Rojo (<20%)
                elif dr_val > 0.45:
                    c_dr.fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid") # Azul (>45%)
                else:
                    c_dr.fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid") # Verde (20-45%)

            # 3. Semáforo CE (1.4 - 2.2 dS/m Óptimo)
            if isinstance(ce_val, (int, float)):
                if ce_val < 1.4:
                    c_ce.fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid") # Amarillo (Baja nutrición)
                elif ce_val > 2.2:
                    c_ce.fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid") # Rojo (Exceso sales)
                else:
                    c_ce.fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid") # Verde (1.4 - 2.2)

            # 4. Semáforo pH (5.8 - 6.5 Óptimo)
            if isinstance(ph_val, (int, float)):
                if ph_val < 5.8:
                    c_ph.fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid") # Rojo (Ácido)
                elif ph_val > 6.5:
                    c_ph.fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid") # Amarillo (Alcalino)
                else:
                    c_ph.fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid") # Verde (Óptimo)

        # Autoajustar anchos en DECISION_RIEGO
        for col in ws_dec.columns:
            max_l = max(len(str(cell.value or '')) for cell in col if cell.row >= 6)
            col_letter = get_column_letter(col[0].column)
            ws_dec.column_dimensions[col_letter].width = max(max_l + 4, 12)

        # -------------------------------------------------------------
        # HOJA 2: DATOS_FOTO_DETALLADOS (La tabla exacta de la libreta)
        # -------------------------------------------------------------
        df_foto.to_excel(writer, index=False, sheet_name="DATOS_FOTO_DETALLADOS", startrow=3)
        ws_foto = writer.sheets["DATOS_FOTO_DETALLADOS"]
        ws_foto.views.sheetView[0].showGridLines = True

        num_cols_f = len(df_foto.columns)
        col_fin_f = get_column_letter(num_cols_f)
        col_fin_notas = get_column_letter(max(num_cols_f, 6))

        # Título
        ws_foto.merge_cells(f"A1:{col_fin_f}1")
        ws_foto["A1"] = f"{titulo.upper()} (DATOS CRUDOS DIGITALIZADOS)"
        ws_foto["A1"].font = Font(name="Segoe UI", size=13, bold=True, color="1F4E78")
        ws_foto["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws_foto.row_dimensions[1].height = 28

        # Subtítulo
        fecha_texto = f"Fecha: {fecha}  |  " if fecha else ""
        fecha_gen = datetime.now().strftime("%d/%m/%Y %H:%M")
        ws_foto.merge_cells(f"A2:{col_fin_f}2")
        ws_foto["A2"] = f"{fecha_texto}Total Registros: {len(df_foto)}  |  Generado: {fecha_gen}"
        ws_foto["A2"].font = Font(name="Segoe UI", size=9, italic=True, color="595959")
        ws_foto["A2"].alignment = Alignment(horizontal="center", vertical="center")
        ws_foto.row_dimensions[2].height = 18

        # Encabezados
        for col_num in range(1, num_cols_f + 1):
            cell = ws_foto.cell(row=4, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws_foto.row_dimensions[4].height = 26

        # Filas cebra
        zebra_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
        white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        
        start_row_f = 5
        for row_idx, row in enumerate(ws_foto.iter_rows(min_row=start_row_f, max_row=start_row_f + len(df_foto) - 1, min_col=1, max_col=num_cols_f)):
            ws_foto.row_dimensions[row[0].row].height = 20
            es_par = (row_idx % 2 == 1)
            for col_idx, cell in enumerate(row):
                cell.border = thin_border
                cell.font = Font(name="Segoe UI", size=9)
                cell.fill = zebra_fill if es_par else white_fill
                val_s = str(cell.value or "")
                cell.alignment = Alignment(horizontal="center" if len(val_s) <= 8 or col_idx == 0 else "left", vertical="center")

        # Autoajustar ancho de foto
        for col in ws_foto.columns:
            max_l = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row >= 4 and cell.row < start_row_f + len(df_foto) and cell.value is not None:
                    max_l = max(max_l, len(str(cell.value)))
            ws_foto.column_dimensions[col_letter].width = max(max_l + 4, 12)

        # Sección destacada de Excedentes y Notas al pie
        lista_notas = []
        if isinstance(notas, list):
            lista_notas = [str(n).strip() for n in notas if str(n).strip()]
        elif isinstance(notas, str) and notas.strip():
            lista_notas = [line.strip() for line in notas.splitlines() if line.strip()]

        if lista_notas:
            fila_separador = start_row_f + len(df_foto) + 1
            ws_foto.row_dimensions[fila_separador].height = 14
            
            fila_encabezado_notas = fila_separador + 1
            ws_foto.merge_cells(f"A{fila_encabezado_notas}:{col_fin_notas}{fila_encabezado_notas}")
            ws_foto[f"A{fila_encabezado_notas}"] = "📌 MEDICIONES ADICIONALES, EXCEDENTES Y OBSERVACIONES"
            ws_foto[f"A{fila_encabezado_notas}"].font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
            ws_foto[f"A{fila_encabezado_notas}"].fill = PatternFill(start_color="2A4D69", end_color="2A4D69", fill_type="solid")
            ws_foto[f"A{fila_encabezado_notas}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws_foto.row_dimensions[fila_encabezado_notas].height = 26

            fila_actual = fila_encabezado_notas + 1
            nota_fill_par = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
            nota_fill_impar = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

            for idx, nota_texto in enumerate(lista_notas):
                ws_foto.merge_cells(f"A{fila_actual}:{col_fin_notas}{fila_actual}")
                celda = ws_foto[f"A{fila_actual}"]
                celda.value = f"  •   {nota_texto}"
                celda.font = Font(name="Segoe UI", size=10, color="1E293B", bold=False)
                celda.fill = nota_fill_par if idx % 2 == 0 else nota_fill_impar
                celda.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                celda.border = thin_border
                ws_foto.row_dimensions[fila_actual].height = 24
                fila_actual += 1

        # -------------------------------------------------------------
        # HOJA 3: DASHBOARD_EJECUTIVO (Resumen de Balance y Cumplimiento)
        # -------------------------------------------------------------
        ws_dash = wb.create_sheet(title="DASHBOARD_EJECUTIVO")
        ws_dash.views.sheetView[0].showGridLines = True
        
        ws_dash["A1"] = '📊 DASHBOARD EJECUTIVO Y CONTROL DE FERTIRRIEGO'
        ws_dash["A1"].font = Font(name="Segoe UI", size=14, bold=True, color="1F4E78")
        ws_dash.row_dimensions[1].height = 26

        # Tarjetas de resumen
        ws_dash["A3"] = "💧 CONSUMO TOTAL (M³)"
        ws_dash["A4"] = f"=SUM('DECISION_RIEGO'!K7:K{6 + len(df_decision)})/1000"
        ws_dash["A4"].number_format = "#,##0.0"

        ws_dash["D3"] = "🎯 % CUMPLIMIENTO GLOBAL"
        ws_dash["D4"] = f"=SUM('DECISION_RIEGO'!K7:K{6 + len(df_decision)})/SUM('DECISION_RIEGO'!G7:G{6 + len(df_decision)})"
        ws_dash["D4"].number_format = "0.0%"

        ws_dash["G3"] = "🧪 PROMEDIO pH"
        ws_dash["G4"] = f"=AVERAGE('DECISION_RIEGO'!M7:M{6 + len(df_decision)})"
        ws_dash["G4"].number_format = "0.0"

        ws_dash["J3"] = "⚡ PROMEDIO C.E. (dS/m)"
        ws_dash["J4"] = f"=AVERAGE('DECISION_RIEGO'!L7:L{6 + len(df_decision)})"
        ws_dash["J4"].number_format = "0.00"

        for col_d in ["A", "D", "G", "J"]:
            ws_dash[f"{col_d}3"].font = Font(name="Segoe UI", size=8, bold=True, color="595959")
            ws_dash[f"{col_d}4"].font = Font(name="Segoe UI", size=14, bold=True, color="1F4E78")
            ws_dash[f"{col_d}4"].fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
            ws_dash[f"{col_d}4"].alignment = Alignment(horizontal="center", vertical="center")
            ws_dash[f"{col_d}4"].border = thin_border

        ws_dash.row_dimensions[3].height = 16
        ws_dash.row_dimensions[4].height = 25

        # -------------------------------------------------------------
        # HOJA 4: PARAMETROS_NUTRICION (Reglas de Semáforos y Parámetros)
        # -------------------------------------------------------------
        ws_param = wb.create_sheet(title="PARAMETROS_NUTRICION")
        ws_param.views.sheetView[0].showGridLines = True

        ws_param["A1"] = "⚙️ PARÁMETROS AGRONÓMICOS, NUTRICIÓN Y CÓDIGOS DE DECISIÓN DE RIEGO"
        ws_param["A1"].font = Font(name="Segoe UI", size=13, bold=True, color="1F4E78")

        headers_p = ["VARIABLE", "ROJO (DÉFICIT / ALERTA)", "VERDE (RANGO ÓPTIMO)", "AZUL / AMARILLO (EXCESO)", "UNIDAD"]
        for col_p, h_p in enumerate(headers_p, 1):
            cell_p = ws_param.cell(row=3, column=col_p, value=h_p)
            cell_p.fill = header_fill
            cell_p.font = header_font
            cell_p.alignment = Alignment(horizontal="center", vertical="center")
        ws_param.row_dimensions[3].height = 24

        param_rows = [
            ("VOL. EJ (Volumen Ejecutado)", "< 10% MENOS (Déficit / Gotero tapado)", "Entre 10% MENOS y 10% MAS", "> 10% MAS (Exceso / Fuga)", "Litros"),
            ("Porcentaje de Drenaje (%DR)", "< 20.0% (Riesgo Salinización)", "20.0% - 45.0% (Lavado Óptimo)", "> 45.0% (Desperdicio / Lixiviación)", "%"),
            ("pH Solución / Drenaje", "< 5.8 (Acidosis)", "5.8 - 6.5 (Asimilación Óptima)", "> 6.5 (Alcalosis / Bloqueo)", "pH"),
            ("Conductividad Eléctrica (CE)", "< 1.4 dS/m (Baja Nutrición)", "1.4 - 2.2 dS/m (Nutrición Balanceada)", "> 2.2 dS/m (Exceso de Sales)", "dS/m"),
        ]

        for p_idx, p_data in enumerate(param_rows, 4):
            for p_c, p_val in enumerate(p_data, 1):
                c_item = ws_param.cell(row=p_idx, column=p_c, value=p_val)
                c_item.font = Font(name="Segoe UI", size=9)
                c_item.border = thin_border
                c_item.alignment = Alignment(horizontal="center" if p_c > 1 else "left", vertical="center")
            ws_param.row_dimensions[p_idx].height = 20

        # Autoajustar ancho en parámetros
        for col in ws_param.columns:
            max_l = max(len(str(cell.value or '')) for cell in col if cell.row >= 3)
            col_letter = get_column_letter(col[0].column)
            ws_param.column_dimensions[col_letter].width = max(max_l + 4, 15)

    buffer.seek(0)
    return buffer

# --- ACCIÓN PRINCIPAL ---

if archivo_subido is not None:
    if st.button("🚀 Extraer Datos y Generar Excel con Decisión de Riego", type="primary"):
        bytes_foto = archivo_subido.read()
        
        with st.spinner("Procesando imagen, mapeando indicadores y calculando semáforos agronómicos..."):
            try:
                if "IA" in modo:
                    if not api_key_activa:
                        st.error("⚠️ Para usar el modo IA por favor ingresa tu API Key en la barra lateral izquierda.")
                        st.stop()
                    df_resultado, titulo_doc, fecha_doc, notas_doc = extraer_con_ia_vision(bytes_foto, api_key_activa)
                else:
                    df_resultado, titulo_doc, fecha_doc, notas_doc = extraer_con_tesseract(bytes_foto)

                if not df_resultado.empty:
                    st.success(f"¡Se detectaron exitosamente **{len(df_resultado)} filas** y se generaron los semáforos agronómicos!")
                    
                    # Mapear a Decisión de Riego
                    df_decision_vista = mapear_decision_riego(df_resultado, fecha_doc)

                    # Pestañas interactivas en la App
                    tab_dec, tab_foto, tab_notas = st.tabs([
                        "🌱 Hoja Decisión de Riego (Semáforos y Funciones)", 
                        "📸 Datos Crudos de la Foto", 
                        "📌 Excedentes y Observaciones"
                    ])

                    with tab_dec:
                        st.markdown("##### 🚦 Tabla Oficial de Decisión de Riego")
                        st.caption("💡 *Incluye fórmulas automáticas (Lt/Pul, DE. RIEGO, 10% MAS/MENOS) y semáforos por rangos:*")
                        
                        # Mostrar métricas resumen
                        m1, m2, m3, m4 = st.columns(4)
                        m1.metric("Total Válvulas", len(df_decision_vista))
                        
                        # Promedios numéricos seguros
                        ce_nums = pd.to_numeric(df_decision_vista["CE"], errors='coerce').dropna()
                        ph_nums = pd.to_numeric(df_decision_vista["PH"], errors='coerce').dropna()
                        dr_nums = pd.to_numeric(df_decision_vista["%DR"], errors='coerce').dropna()

                        if not ce_nums.empty:
                            m2.metric("Promedio C.E.", f"{ce_nums.mean():.2f} dS/m")
                        if not ph_nums.empty:
                            m3.metric("Promedio pH", f"{ph_nums.mean():.1f}")
                        if not dr_nums.empty:
                            m4.metric("Promedio Drenaje", f"{dr_nums.mean() * 100:.1f}%" if dr_nums.mean() <= 1 else f"{dr_nums.mean():.1f}%")

                        df_decision_editado = st.data_editor(df_decision_vista, use_container_width=True, height=380)

                    with tab_foto:
                        st.markdown("##### 📋 Datos Extraídos Directamente de la Imagen")
                        df_foto_editado = st.data_editor(df_resultado, use_container_width=True, height=380)

                    with tab_notas:
                        st.markdown("##### 📌 Mediciones Adicionales, Excedentes y Observaciones al Pie")
                        notas_formateadas = []
                        if isinstance(notas_doc, list):
                            notas_formateadas = notas_doc
                        elif isinstance(notas_doc, str) and notas_doc.strip():
                            notas_formateadas = [n.strip() for n in notas_doc.splitlines() if n.strip()]

                        texto_notas_default = "\n".join([f"• {n}" if not n.startswith("•") else n for n in notas_formateadas])
                        notas_editadas = st.text_area(
                            "Edita o añade notas y excedentes (aparecerán con formato grande y ordenado en el Excel):",
                            value=texto_notas_default,
                            height=140
                        )
                        notas_finales = [n.replace("•", "").strip() for n in notas_editadas.splitlines() if n.strip()]

                    # Generar Excel completo con todas las pestañas y semáforos
                    buffer_excel = generar_excel_estilizado(df_foto_editado, titulo_doc, fecha_doc, notas_finales)
                    nombre_final = f"{nombre_archivo.strip() or 'Reporte_Decision_Riego'}.xlsx"

                    st.markdown("---")
                    st.download_button(
                        label="📥 Descargar Libro Excel Completo (con Semáforos, Decisión de Riego y Dashboard)",
                        data=buffer_excel,
                        file_name=nombre_final,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                else:
                    st.error("No se pudieron reconocer datos en la imagen. Verifica que la foto sea legible.")
            except Exception as e:
                st.error(f"Ocurrió un error al procesar la imagen: {e}")
