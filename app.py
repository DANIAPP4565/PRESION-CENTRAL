# app_pac.py
# App Streamlit para informe de Presión Aórtica Central (PAC)
# Importa PDF tipo MODELO PAC y archivos TXT estructurados o sin formato de métricas, extrae variables, genera historial Excel y PDF médico integrado.

import io, re, math, tempfile
from datetime import datetime
from pathlib import Path

import numpy as np


def safe_trapezoid_integral(y, x):
    """Integración trapezoidal compatible con NumPy nuevo y antiguo.
    Evita AttributeError cuando np.trapz no está disponible.
    """
    y = np.asarray(y, dtype=float)
    x = np.asarray(x, dtype=float)
    if y.size < 2 or x.size < 2:
        return 0.0
    y = np.nan_to_num(y, nan=0.0, posinf=0.0, neginf=0.0)
    x = np.nan_to_num(x, nan=0.0, posinf=0.0, neginf=0.0)
    if hasattr(np, "trapezoid"):
        return float(np.trapezoid(y, x))
    if hasattr(np, "trapz"):
        return float(np.trapz(y, x))
    return float(np.sum((y[1:] + y[:-1]) * 0.5 * np.diff(x)))
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

try:
    import pdfplumber
except Exception:
    pdfplumber = None

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image,
    PageBreak, KeepTogether
)
from reportlab.pdfbase.pdfmetrics import stringWidth

st.set_page_config(page_title="PAC IA - Presión Aórtica Central", layout="wide")

APP_TITLE = "PAC IA - Informe médico integrado de Presión Aórtica Central"
HISTORIAL_FILE = Path("historial_pac.xlsx")

BIBLIOGRAFIA = [
    "Agabiti-Rosei E, Mancia G, O'Rourke MF, et al. Central blood pressure measurements and antihypertensive therapy: a consensus document. Hypertension. 2007;50:154-160.",
    "Zócalo Y, Bia D. Presión aórtica central y parámetros clínicos derivados de la onda del pulso: evaluación no invasiva en la práctica clínica. Rev Urug Cardiol. 2014;29:215-230.",
    "Herbert A, Cruickshank JK, Laurent S, Boutouyrie P, et al. Establishing reference values for central blood pressure and its amplification. Eur Heart J. 2014;35:3122-3133.",
    "Westerhof BE, Guelen I, Westerhof N, Karemaker JM, Avolio A. Quantification of wave reflection in the human aorta from pressure alone. Hypertension. 2006;48:595-601.",
    "Norton GR, An DW, Aparicio LS, et al. Mortality and cardiovascular end points in relation to the aortic pulse wave components. Hypertension. 2024;81:1065-1075.",
    "Huang QF, An DW, Aparicio LS, et al. An outcome-driven threshold for pulse pressure amplification. Hypertension Research. 2024."
]

CENTRAL_SBP_TABLE = {
    "Óptimo": {"F": 108, "M": 97},
    "Normal": {"F": 123, "M": 116},
    "Normal alta": {"F": 133, "M": 126},
    "Etapa 1": {"F": 143, "M": 137},
    "Etapa 2": {"F": 161, "M": 154},
    "Etapa 3": {"F": 183, "M": 173},
    "ISH": {"F": 147, "M": 140},
}

def safe_text(x):
    if x is None:
        return ""
    return str(x).replace("\x00", "").strip()

def to_float(x):
    if x is None or str(x).strip() == "":
        return np.nan
    s = str(x).replace(",", ".")
    s = re.sub(r"[^0-9.\-+]", "", s)
    try:
        return float(s)
    except Exception:
        return np.nan



def is_missing_value(x):
    """True para valores vacíos/NaN usados por el parser."""
    if x is None:
        return True
    try:
        if pd.isna(x):
            return True
    except Exception:
        pass
    if isinstance(x, str) and x.strip().lower() in ["", "nan", "none", "---"]:
        return True
    return False

def set_if_present(data, key, value, overwrite_zero=False):
    """Carga un valor si es válido. Permite reemplazar 0 cuando viene de falla de importación."""
    if is_missing_value(value):
        return
    cur = data.get(key, np.nan)
    if is_missing_value(cur) or (overwrite_zero and to_float(cur) == 0 and to_float(value) != 0):
        data[key] = value

def extract_pdf_text(pdf_bytes):
    text = ""
    if pdfplumber is not None:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                text += "\n" + (page.extract_text() or "")
    elif fitz is not None:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        text = "\n".join([p.get_text() for p in doc])
    return text



def decode_uploaded_text(uploaded_file):
    """Lee TXT/CSV exportados por equipos/software sin fallar por codificación.
    Soporta UTF-8, UTF-8 BOM, Windows-1252, Latin-1 e ISO-8859-1.
    """
    if uploaded_file is None:
        return ""
    try:
        uploaded_file.seek(0)
    except Exception:
        pass
    raw = uploaded_file.read()
    if isinstance(raw, str):
        return raw
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1", "iso-8859-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
        except Exception:
            continue
    return raw.decode("latin-1", errors="replace")

def parse_metric_txt(text):
    """
    Importador flexible de métricas desde .txt.
    Acepta TXT estructurado y TXT sin formato: texto pegado desde PDF, OCR,
    saltos de línea rotos, bloques tipo Exxer/MODELO PAC, o pares etiqueta-valor.
    Ejemplos válidos:
      Paciente Juan Perez PAS central 126 PAD central 90 PP central 36
      Radial Central 139 126 89 90 108 108 50 36
      PAS PAD PAM PP / Radial 139 89 108 50 / Central 126 90 108 36
    """
    txt = safe_text(text)
    # Normalización suave: preserva saltos para el parser MODELO PAC y agrega una copia canónica.
    data = parse_model_pac(txt)

    # Copia canónica para TXT sin formato: corrige OCR frecuente y permite búsqueda global.
    canon = txt.replace("\t", " ")
    canon = canon.replace("º", "o").replace("°", "o")
    canon = re.sub(r"[;|]+", "\n", canon)
    canon = re.sub(r"(?<=\d),(?=\d)", ".", canon)  # coma decimal
    canon = re.sub(r" {2,}", " ", canon)
    flat = re.sub(r"\s+", " ", canon).strip()

    def first_number_after(pattern, source=flat, window=80):
        m = re.search(pattern, source, re.I)
        if not m:
            return np.nan
        frag = source[m.end():m.end()+window]
        n = re.search(r"[-+]?\d+(?:[.,]\d+)?", frag)
        return to_float(n.group(0)) if n else np.nan

    def first_text_after(pattern, source=flat, window=90):
        m = re.search(pattern, source, re.I)
        if not m:
            return ""
        frag = source[m.end():m.end()+window]
        frag = re.split(r"\b(?:Número|Numero|Estudio|Fecha|Hora|Médico|Medico|Sede|Edad|Sexo|Peso|Altura|IMC|Diagnóstico|Diagnostico|Medicación|Medicacion|PAS|PAD|PAM|PP|FC)\b", frag, maxsplit=1, flags=re.I)[0]
        return safe_text(re.sub(r"^\s*[:#=\-]*\s*", "", frag))

    label_map = {
        "paciente": [r"paciente", r"nombre del paciente", r"apellido y nombre"],
        "estudio": [r"estudio", r"numero de estudio", r"número de estudio"],
        "fecha": [r"fecha"],
        "hora": [r"hora"],
        "hc": [r"h\.?c\.?"] ,
        "diagnostico_previo": [r"diagnostico", r"diagnóstico"],
        "medicacion": [r"medicacion", r"medicación"],
        "edad": [r"edad"],
        "sexo": [r"sexo"],
        "peso": [r"peso"],
        "altura": [r"altura", r"talla"],
        "imc": [r"imc"],
        "sc": [r"sc", r"superficie corporal"],
        "fc": [r"fc", r"frecuencia cardiaca", r"frecuencia cardíaca"],
        "pas_radial": [r"pas radial", r"pas braquial", r"sistolica radial", r"sistólica radial"],
        "pad_radial": [r"pad radial", r"pad braquial", r"diastolica radial", r"diastólica radial"],
        "pam_radial": [r"pam radial", r"pam braquial", r"media radial"],
        "pp_radial": [r"pp radial", r"pp braquial", r"presion de pulso radial", r"presión de pulso radial"],
        "pas_central": [r"pas central", r"pas aortica", r"pas aórtica", r"presion sistolica central", r"presión sistólica central", r"csbp"],
        "pad_central": [r"pad central", r"pad aortica", r"pad aórtica", r"presion diastolica central", r"presión diastólica central"],
        "pam_central": [r"pam central", r"pam aortica", r"pam aórtica", r"presion arterial media central", r"presión arterial media central"],
        "pp_central": [r"pp central", r"pp aortica", r"pp aórtica", r"presion de pulso central", r"presión de pulso central"],
        "au": [r"au", r"aumentacion", r"aumentación", r"augmentation"],
        "iau": [r"iau", r"aix", r"indice de aumentacion", r"índice de aumentación", r"augmentation index"],
        "rvse": [r"rvse", r"sevr", r"relacion de viabilidad", r"relación de viabilidad"],
        "pe": [r"pe", r"periodo eyectivo", r"período eyectivo", r"ejection duration"],
        "apc": [r"apc", r"amplificacion periferico central", r"amplificación periférico central"],
    }

    for key, labels in label_map.items():
        current = data.get(key, "")
        current_missing = (current == "" or (isinstance(current, float) and np.isnan(current)))
        if not current_missing and key not in ["sexo"]:
            continue
        for lab in labels:
            # Captura hasta fin de línea, coma, punto y coma o separador largo.
            pat = re.compile(rf"(?im)^\s*{lab}\s*(?:[:=\-]|\s)\s*([^\n;|]+)")
            m = pat.search(canon)
            if m:
                val = safe_text(m.group(1))
                if key in ["paciente","estudio","fecha","hora","hc","diagnostico_previo","medicacion","sexo"]:
                    data[key] = val.upper()[:1] if key == "sexo" else val
                else:
                    data[key] = to_float(val)
                break

    # Formato compacto frecuente: "PAS central 126" sin dos puntos.
    compact_patterns = {
        "pas_central": r"PAS\s+(?:CENTRAL|AORTICA|AÓRTICA)\s+([0-9]+(?:[,.][0-9]+)?)",
        "pad_central": r"PAD\s+(?:CENTRAL|AORTICA|AÓRTICA)\s+([0-9]+(?:[,.][0-9]+)?)",
        "pam_central": r"PAM\s+(?:CENTRAL|AORTICA|AÓRTICA)\s+([0-9]+(?:[,.][0-9]+)?)",
        "pp_central": r"PP\s+(?:CENTRAL|AORTICA|AÓRTICA)\s+([0-9]+(?:[,.][0-9]+)?)",
        "pas_radial": r"PAS\s+(?:RADIAL|BRAQUIAL)\s+([0-9]+(?:[,.][0-9]+)?)",
        "pad_radial": r"PAD\s+(?:RADIAL|BRAQUIAL)\s+([0-9]+(?:[,.][0-9]+)?)",
        "pam_radial": r"PAM\s+(?:RADIAL|BRAQUIAL)\s+([0-9]+(?:[,.][0-9]+)?)",
        "pp_radial": r"PP\s+(?:RADIAL|BRAQUIAL)\s+([0-9]+(?:[,.][0-9]+)?)",
    }
    for key, pat in compact_patterns.items():
        cur = data.get(key, np.nan)
        if isinstance(cur, float) and np.isnan(cur):
            m = re.search(pat, canon, re.I)
            if m:
                data[key] = to_float(m.group(1))

    # --- TXT sin formato: extracción global por cercanía de etiqueta ---
    global_num_patterns = {
        "edad": r"\bEdad\b",
        "peso": r"\bPeso\b",
        "altura": r"\b(?:Altura|Talla)\b",
        "imc": r"\bIMC\b",
        "sc": r"\b(?:SC|Superficie\s+corporal)\b",
        "fc": r"\b(?:FC|Frecuencia\s+card[ií]aca)\b",
        "pas_radial": r"\bPAS\b[^\n]{0,25}\b(?:Radial|Braquial|Perif[eé]rica)\b|\b(?:Radial|Braquial|Perif[eé]rica)\b[^\n]{0,25}\bPAS\b",
        "pad_radial": r"\bPAD\b[^\n]{0,25}\b(?:Radial|Braquial|Perif[eé]rica)\b|\b(?:Radial|Braquial|Perif[eé]rica)\b[^\n]{0,25}\bPAD\b",
        "pam_radial": r"\bPAM\b[^\n]{0,25}\b(?:Radial|Braquial|Perif[eé]rica)\b|\b(?:Radial|Braquial|Perif[eé]rica)\b[^\n]{0,25}\bPAM\b",
        "pp_radial": r"\bPP\b[^\n]{0,25}\b(?:Radial|Braquial|Perif[eé]rica)\b|\b(?:Radial|Braquial|Perif[eé]rica)\b[^\n]{0,25}\bPP\b",
        "pas_central": r"\bPAS\b[^\n]{0,30}\b(?:Central|A[oó]rtica)\b|\b(?:Central|A[oó]rtica)\b[^\n]{0,30}\bPAS\b|\bcSBP\b",
        "pad_central": r"\bPAD\b[^\n]{0,30}\b(?:Central|A[oó]rtica)\b|\b(?:Central|A[oó]rtica)\b[^\n]{0,30}\bPAD\b",
        "pam_central": r"\bPAM\b[^\n]{0,30}\b(?:Central|A[oó]rtica)\b|\b(?:Central|A[oó]rtica)\b[^\n]{0,30}\bPAM\b",
        "pp_central": r"\bPP\b[^\n]{0,30}\b(?:Central|A[oó]rtica)\b|\b(?:Central|A[oó]rtica)\b[^\n]{0,30}\bPP\b",
        "au": r"\b(?:Au|Aumentaci[oó]n\s+a[oó]rtica)\b",
        "iau": r"\b(?:IAu|AIx|[IÍ]ndice\s+de\s+aumentaci[oó]n)\b",
        "rvse": r"\b(?:RVSE|SEVR|Viabilidad\s+Sub[- ]?Endoc[aá]rdica)\b",
        "pe": r"\b(?:PE|Per[ií]odo\s+Eyectivo)\b",
        "apc": r"\b(?:APC|Amplificaci[oó]n\s+Perif[eé]rico[- ]?Central)\b",
    }
    for key, pat in global_num_patterns.items():
        cur = data.get(key, np.nan)
        missing = cur == "" or (isinstance(cur, float) and np.isnan(cur))
        if missing:
            val = first_number_after(pat)
            if not np.isnan(val):
                data[key] = val

    # Campos de texto en TXT sin formato.
    for key, pat in {
        "paciente": r"(?:Nombre\s+del\s+paciente|Paciente)",
        "estudio": r"(?:N[uú]mero\s+de\s+estudio|Estudio\s*#?|Estudio)",
        "fecha": r"\bFecha\b",
        "hora": r"\bHora\b",
        "hc": r"\bH\.?C\.?\s*#?",
        "diagnostico_previo": r"\bDiagn[oó]stico\b",
        "medicacion": r"\bMedicaci[oó]n\b",
        "sexo": r"\bSexo\b",
    }.items():
        cur = data.get(key, "")
        if cur == "" or (isinstance(cur, float) and np.isnan(cur)):
            val = first_text_after(pat)
            if val:
                data[key] = val.upper()[:1] if key == "sexo" else val

    # --- Bloques numéricos sin formato tipo Exxer ---
    # Caso 1: "Radial Central 139 126 89 90 108 108 50 36".
    m = re.search(r"Radial\s+Central\s+((?:[-+]?\d+(?:[.,]\d+)?\s+){7,12}[-+]?\d+(?:[.,]\d+)?)", flat, re.I)
    if m:
        nums = [to_float(x) for x in re.findall(r"[-+]?\d+(?:[.,]\d+)?", m.group(1))]
        if len(nums) >= 8:
            # En MODELO PAC el orden visual suele ser PAS, PAD, PP, PAM por pares radial/central.
            data["pas_radial"], data["pas_central"] = nums[0], nums[1]
            data["pad_radial"], data["pad_central"] = nums[2], nums[3]
            # Si el tercer par parece PAM (>PP), intercambiar con PP.
            if nums[4] > 70 and nums[6] < 80:
                data["pam_radial"], data["pam_central"] = nums[4], nums[5]
                data["pp_radial"], data["pp_central"] = nums[6], nums[7]
            else:
                data["pp_radial"], data["pp_central"] = nums[4], nums[5]
                data["pam_radial"], data["pam_central"] = nums[6], nums[7]

    # Caso 2: tabla con encabezado PAS PAD PAM PP y filas Radial/Central.
    m = re.search(r"PAS\s+PAD\s+PAM\s+PP.*?Radial\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+).*?Central\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)", flat, re.I)
    if m:
        vals = [to_float(x) for x in m.groups()]
        data.update({
            "pas_radial": vals[0], "pad_radial": vals[1], "pam_radial": vals[2], "pp_radial": vals[3],
            "pas_central": vals[4], "pad_central": vals[5], "pam_central": vals[6], "pp_central": vals[7],
        })

    # Caso 3: lista de parámetros centrales después de "PAS PP Au IAu ...".
    m = re.search(r"PAS\s+PP\s+Au\s+IAu\s+(?:RVSE\s+)?PE\s+APC?\s+((?:[-+]?\d+(?:[.,]\d+)?\s+){4,8}[-+]?\d+(?:[.,]\d+)?)", flat, re.I)
    if m:
        nums = [to_float(x) for x in re.findall(r"[-+]?\d+(?:[.,]\d+)?", m.group(1))]
        if len(nums) >= 5:
            data["pas_central"] = nums[0]
            data["pp_central"] = nums[1]
            data["au"] = nums[2]
            data["iau"] = nums[3]
            if len(nums) == 5:
                data["pe"] = nums[4]
            elif len(nums) >= 6:
                data["rvse"] = nums[4]
                data["pe"] = nums[5]
                if len(nums) >= 7:
                    data["apc"] = nums[6]

    # Si falta PAM, calcularla. Si falta PP, calcularla.
    for pref in ("radial", "central"):
        pas = to_float(data.get(f"pas_{pref}")); pad = to_float(data.get(f"pad_{pref}"))
        if (isinstance(data.get(f"pp_{pref}"), float) and np.isnan(data.get(f"pp_{pref}"))) and not np.isnan(pas) and not np.isnan(pad):
            data[f"pp_{pref}"] = pas - pad
        if (isinstance(data.get(f"pam_{pref}"), float) and np.isnan(data.get(f"pam_{pref}"))) and not np.isnan(pas) and not np.isnan(pad):
            data[f"pam_{pref}"] = pad + (pas - pad) / 3
    return data


def render_pdf_page_png(pdf_bytes, page_index=1, zoom=2.0):
    if fitz is None:
        return None
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    if len(doc) == 0:
        return None
    page_index = min(max(page_index, 0), len(doc)-1)
    pix = doc[page_index].get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    return pix.tobytes("png")

def find_after(label, text, default=""):
    pat = re.compile(label + r"\s*[:#]?\s*([^\n]+)", re.I)
    m = pat.search(text)
    return safe_text(m.group(1)) if m else default


def parse_metric_csv(uploaded_file):
    """Importa métricas del estudio desde CSV/TXT tabular o clave-valor.

    Corrige dos problemas frecuentes:
    1) UnicodeDecodeError: nunca pasa bytes crudos a pandas; primero decodifica con fallback.
    2) Au/IAu no detectados: interpreta columnas, filas clave-valor y bloques verticales tipo Exxer.
    """
    txt = decode_uploaded_text(uploaded_file)
    if not txt.strip():
        return {}

    txt_norm = txt.replace("\ufeff", "")

    def norm_col(x):
        x = safe_text(x).lower()
        x = x.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u").replace("ü", "u").replace("ñ", "n")
        x = re.sub(r"[^a-z0-9]+", "_", x).strip("_")
        return x

    col_map = {
        "paciente": ["paciente", "nombre_paciente", "nombre_del_paciente", "apellido_nombre", "apellido_y_nombre"],
        "estudio": ["estudio", "numero_estudio", "nro_estudio", "id_estudio"],
        "fecha": ["fecha"],
        "hora": ["hora"],
        "edad": ["edad"],
        "sexo": ["sexo", "genero"],
        "peso": ["peso", "peso_kg"],
        "altura": ["altura", "talla", "altura_cm"],
        "imc": ["imc", "bmi"],
        "sc": ["sc", "superficie_corporal"],
        "fc": ["fc", "frecuencia_cardiaca", "frecuencia_cardíaca", "hr"],
        "pas_radial": ["pas_radial", "pas_braquial", "sbp_radial", "sbp_braquial", "radial_pas", "braquial_pas", "pas_periferica"],
        "pad_radial": ["pad_radial", "pad_braquial", "dbp_radial", "dbp_braquial", "radial_pad", "braquial_pad", "pad_periferica"],
        "pam_radial": ["pam_radial", "pam_braquial", "map_radial", "map_braquial", "radial_pam", "braquial_pam", "pam_periferica"],
        "pp_radial": ["pp_radial", "pp_braquial", "pulse_pressure_radial", "radial_pp", "braquial_pp", "pp_periferica"],
        "pas_central": ["pas_central", "pas_aortica", "pas_aórtica", "csbp", "central_pas", "aortica_pas", "aortic_sbp"],
        "pad_central": ["pad_central", "pad_aortica", "pad_aórtica", "cdbp", "central_pad", "aortica_pad", "aortic_dbp"],
        "pam_central": ["pam_central", "pam_aortica", "pam_aórtica", "cmap", "central_pam", "aortica_pam", "aortic_map"],
        "pp_central": ["pp_central", "pp_aortica", "pp_aórtica", "cpp", "central_pp", "aortica_pp", "aortic_pp"],
        "au": ["au", "aumentacion", "aumentacion_aortica", "aumentacion_aortica_central", "aumentacion_central", "augmentation", "augmentation_pressure", "augmented_pressure", "ap"],
        "iau": ["iau", "iauc", "aix", "aix_75", "caix", "indice_aumentacion", "indice_de_aumentacion", "indice_aumentacion_central", "indice_de_aumentacion_central", "augmentation_index"],
        "rvse": ["rvse", "sevr", "relacion_viabilidad_subendocardica", "relacion_de_viabilidad_sub_endocardica"],
        "pe": ["pe", "periodo_eyectivo", "periodo_eyectivo_pct", "ejection_duration", "ed"],
        "apc": ["apc", "amplificacion_periferico_central", "amplificacion_p_c", "amplificacion"],
        "medicacion": ["medicacion", "medicación", "tratamiento"],
        "diagnostico_previo": ["diagnostico", "diagnóstico", "diagnostico_previo"],
    }
    alias_to_target = {}
    for target, aliases in col_map.items():
        for a in aliases:
            alias_to_target[norm_col(a)] = target

    def parse_value_for_target(target, val):
        if target in ["paciente", "estudio", "fecha", "hora", "sexo", "medicacion", "diagnostico_previo"]:
            return safe_text(val).upper()[:1] if target == "sexo" else safe_text(val)
        return to_float(val)

    def extract_from_key_value_df(df):
        out = {}
        if df is None or df.empty or df.shape[1] < 2:
            return out
        for _, r in df.iterrows():
            key_raw = r.iloc[0]
            val = r.iloc[1]
            nk = norm_col(key_raw)
            if nk in alias_to_target:
                out[alias_to_target[nk]] = parse_value_for_target(alias_to_target[nk], val)
            else:
                # Buscar alias contenido dentro de etiquetas largas como "Indice de Aumentacion Central (%)".
                for alias, target in alias_to_target.items():
                    if alias and (nk == alias or nk.startswith(alias) or alias in nk):
                        out[target] = parse_value_for_target(target, val)
                        break
        return out

    attempts = [
        dict(sep=None, engine="python", decimal=","),
        dict(sep=";", engine="python", decimal=","),
        dict(sep=",", engine="python", decimal="."),
        dict(sep=r"\t+", engine="python", decimal=","),
        dict(sep=r"\s+", engine="python", decimal=","),
    ]

    best_out = {}
    best_df = None
    for header in [0, None]:
        for kwargs in attempts:
            try:
                cand = pd.read_csv(io.StringIO(txt_norm), header=header, **kwargs)
                if cand is None or cand.empty:
                    continue
                best_df = cand

                # 1) CSV clave-valor, incluso sin encabezado.
                kv = extract_from_key_value_df(cand)
                best_out.update({k: v for k, v in kv.items() if not is_missing_value(v)})

                # 2) CSV tabular: nombres de columnas = variables, primera fila = valores.
                if header == 0:
                    norm_to_original = {norm_col(c): c for c in cand.columns}
                    r0 = cand.iloc[0]
                    for target, aliases in col_map.items():
                        for alias in aliases:
                            na = norm_col(alias)
                            if na in norm_to_original:
                                best_out[target] = parse_value_for_target(target, r0[norm_to_original[na]])
                                break
                            # Columnas largas que contienen el alias.
                            hit = next((orig for ncol, orig in norm_to_original.items() if na and (ncol == na or ncol.startswith(na) or na in ncol)), None)
                            if hit is not None:
                                best_out[target] = parse_value_for_target(target, r0[hit])
                                break
                if len(best_out) >= 2:
                    break
            except Exception:
                continue
        if len(best_out) >= 2:
            break

    # 3) Siempre complementar con parser de texto libre/vertical para capturar Au e IAu si faltan.
    txt_out = parse_metric_txt(txt_norm)
    for k, v in txt_out.items():
        if not is_missing_value(v):
            # Reemplaza Au/IAu/RVSE/PE si venían 0 por falla de lectura o si estaban vacíos.
            overwrite_zero = k in ["au", "iau", "rvse", "pe", "apc"]
            cur = best_out.get(k, np.nan)
            if is_missing_value(cur) or (overwrite_zero and to_float(cur) == 0 and to_float(v) != 0):
                best_out[k] = v

    # 4) Si aún no hubo mapeo útil, unir todo el CSV como texto y reintentar.
    if not best_out and best_df is not None:
        joined = "\n".join([" ".join(map(str, row)) for row in best_df.astype(str).values.tolist()])
        best_out = parse_metric_txt(joined)

    # Derivaciones fisiológicas útiles si el CSV no trae PP/PAM.
    if not np.isnan(to_float(best_out.get("pas_radial", np.nan))) and not np.isnan(to_float(best_out.get("pad_radial", np.nan))):
        if is_missing_value(best_out.get("pp_radial", np.nan)):
            best_out["pp_radial"] = to_float(best_out["pas_radial"]) - to_float(best_out["pad_radial"])
        if is_missing_value(best_out.get("pam_radial", np.nan)):
            best_out["pam_radial"] = to_float(best_out["pad_radial"]) + (to_float(best_out["pp_radial"]) / 3.0)
    if not np.isnan(to_float(best_out.get("pas_central", np.nan))) and not np.isnan(to_float(best_out.get("pad_central", np.nan))):
        if is_missing_value(best_out.get("pp_central", np.nan)):
            best_out["pp_central"] = to_float(best_out["pas_central"]) - to_float(best_out["pad_central"])
        if is_missing_value(best_out.get("pam_central", np.nan)):
            best_out["pam_central"] = to_float(best_out["pad_central"]) + (to_float(best_out["pp_central"]) / 3.0)
    return best_out

def parse_model_pac(text):
    # Parser orientado al patrón MODELO PAC / Exxer. Incluye formatos con etiquetas verticales.
    lines = [safe_text(x) for x in text.splitlines() if safe_text(x)]
    joined = "\n".join(lines)
    flat = re.sub(r"\s+", " ", joined.replace(",", ".")).strip()
    data = {
        "paciente": find_after(r"Paciente|Nombre del paciente", joined),
        "estudio": find_after(r"Estudio|Número de estudio", joined),
        "fecha": find_after(r"Fecha", joined),
        "hora": find_after(r"Hora", joined),
        "hc": find_after(r"H\.?C\.", joined),
        "diagnostico_previo": find_after(r"Diagnóstico", joined),
        "medicacion": find_after(r"Medicación", joined),
        "edad": to_float(find_after(r"Edad", joined)),
        "sexo": find_after(r"Sexo", joined).upper()[:1],
        "peso": to_float(find_after(r"Peso", joined)),
        "altura": to_float(find_after(r"Altura", joined)),
        "imc": to_float(find_after(r"IMC", joined)),
        "sc": to_float(find_after(r"SC", joined)),
        "fc": to_float(find_after(r"FC", joined)),
    }
    # Patrón tabla Radial/Central: en MODELO PAC suele venir PAS, PAD, PP, PAM por pares.
    m = re.search(r"Radial\s+Central\s+((?:[-+]?\d+(?:[.,]\d+)?\s+){7,12}[-+]?\d+(?:[.,]\d+)?)", flat, re.I)
    if m:
        vals = [to_float(x) for x in re.findall(r"[-+]?\d+(?:[.,]\d+)?", m.group(1))]
        if len(vals) >= 8:
            data.update({
                "pas_radial": vals[0], "pas_central": vals[1],
                "pad_radial": vals[2], "pad_central": vals[3],
            })
            # Diferencia entre softwares: a veces el 3er par es PP y a veces PAM.
            if vals[4] > 70 and vals[6] < 80:
                data.update({"pam_radial": vals[4], "pam_central": vals[5], "pp_radial": vals[6], "pp_central": vals[7]})
            else:
                data.update({"pp_radial": vals[4], "pp_central": vals[5], "pam_radial": vals[6], "pam_central": vals[7]})

    # Parámetros centrales alrededor de PAS PP Au IAu RVSE PE APC.
    # El PDF/CSV puede traer etiquetas verticales y luego los valores: PAS, PP, Au, IAu, RVSE, PE, APC.
    block = ""
    mb = re.search(r"Par[aá]metros hemodin[aá]micos centrales(.{0,600})", flat, re.I)
    if mb:
        block = mb.group(1)
    else:
        block = flat

    # Patrón con etiquetas y seis valores: PAS, PP, Au, IAu, RVSE, PE.
    m2 = re.search(
        r"PAS\s+PP\s+Au\s+IAu(?:\s+RVSE)?\s+PE(?:\s+APC)?\s+"
        r"([-+]?\d+(?:\.\d+)?)\s+([-+]?\d+(?:\.\d+)?)\s+([-+]?\d+(?:\.\d+)?)\s+([-+]?\d+(?:\.\d+)?)\s+([-+]?\d+(?:\.\d+)?)\s+([-+]?\d+(?:\.\d+)?)",
        block, re.I)
    if m2:
        vals = [to_float(x) for x in m2.groups()]
        data.update({
            "pas_central": vals[0],
            "pp_central": vals[1],
            "au": vals[2],
            "iau": vals[3],
            "rvse": vals[4],
            "pe": vals[5],
        })

    # Fallback por cercanía de etiqueta: Au +6, IAu +18, RVSE 142, PE 38.1.
    def val_near(label_regex, source=block):
        m = re.search(label_regex + r"\s*[:=]?\s*([-+]?\d+(?:[.,]\d+)?)", source, re.I)
        return to_float(m.group(1)) if m else np.nan
    for key, lab in {
        "au": r"\bAu\b|Aumentaci[oó]n(?:\s+A[oó]rtica)?(?:\s+Central)?",
        "iau": r"\bIAu\b|\bAIx\b|[IÍ]ndice\s+de\s+Aumentaci[oó]n(?:\s+Central)?",
        "rvse": r"\bRVSE\b|Viabilidad\s+Sub[- ]?Endoc[aá]rdica",
        "pe": r"\bPE\b|Per[ií]odo\s+Eyectivo",
        "apc": r"\bAPC\b|Amplificaci[oó]n\s+Perif[eé]rico[- ]?Central",
    }.items():
        v = val_near(lab)
        if not np.isnan(v):
            set_if_present(data, key, v, overwrite_zero=True)

    # Fallback específico cuando aparecen etiquetas verticales y luego números.
    if (is_missing_value(data.get("au")) or to_float(data.get("au")) == 0) and re.search(r"\bAu\b", block, re.I) and re.search(r"\bIAu\b", block, re.I):
        nums = [to_float(x) for x in re.findall(r"[-+]?\d+(?:[.,]\d+)?", block)]
        # Buscar la secuencia fisiológica: PAS central, PP central, Au, IAu, RVSE, PE.
        for i in range(0, max(0, len(nums)-5)):
            pas, pp, au, iau, rvse, pe = nums[i:i+6]
            if 70 <= pas <= 250 and 10 <= pp <= 120 and -50 <= au <= 80 and -80 <= iau <= 120 and 50 <= rvse <= 300 and 10 <= pe <= 80:
                data.update({"pas_central": pas, "pp_central": pp, "au": au, "iau": iau, "rvse": rvse, "pe": pe})
                break

    data.setdefault("pas_radial", np.nan); data.setdefault("pad_radial", np.nan)
    data.setdefault("pam_radial", np.nan); data.setdefault("pp_radial", np.nan)
    data.setdefault("pas_central", np.nan); data.setdefault("pad_central", np.nan)
    data.setdefault("pam_central", np.nan); data.setdefault("pp_central", np.nan)
    data.setdefault("au", np.nan); data.setdefault("iau", np.nan); data.setdefault("rvse", np.nan); data.setdefault("pe", np.nan); data.setdefault("apc", np.nan)
    return data

def brachial_bp_category(pas, pad):
    if np.isnan(pas) or np.isnan(pad): return "No clasificable"
    if pas >= 180 or pad >= 110: return "Etapa 3"
    if pas >= 160 or pad >= 100: return "Etapa 2"
    if pas >= 140 or pad >= 90: return "Etapa 1"
    if pas >= 130 or pad >= 85: return "Normal alta"
    if pas >= 120 or pad >= 80: return "Normal"
    return "Óptimo"

def estimate_central_percentile(row):
    """Percentilo operativo estimado de PAS central.

    La app usa una aproximación clínica por edad/sexo cuando no se dispone de
    tablas completas embebidas del dispositivo. El objetivo es orientar el
    informe; no reemplaza tablas percentilares originales específicas por
    población, dispositivo o calibración.
    """
    csbp = to_float(row.get("pas_central"))
    age = to_float(row.get("edad"))
    sex = (str(row.get("sexo") or "M").upper()[:1])
    if np.isnan(csbp):
        return np.nan, "No clasificable", {"p50": np.nan, "p75": np.nan, "p90": np.nan, "p95": np.nan}
    if np.isnan(age) or age <= 0:
        age = 60.0

    # Umbrales operativos adultos: suben gradualmente con la edad y son algo mayores en varones.
    # Para <22 años se aplica una banda más baja compatible con tablas pediátricas/jóvenes.
    if age < 22:
        base_p50 = 95 + 0.45 * age
        sex_shift = 3 if sex == "M" else 0
        p50 = base_p50 + sex_shift
        p75, p90, p95 = p50 + 8, p50 + 16, p50 + 22
    else:
        decade = max(0, min((age - 30) / 10.0, 6))
        p50 = (111 if sex == "F" else 116) + 2.2 * decade
        p75 = p50 + 10
        p90 = p50 + 20
        p95 = p50 + 27

    # Ajuste de seguridad por categoría braquial: si la PA braquial ya está en categoría alta,
    # no permitir que el P50 operativo quede por debajo de la tabla de clasificación existente.
    pSBP = to_float(row.get("pas_radial")); pDBP = to_float(row.get("pad_radial"))
    cat = brachial_bp_category(pSBP, pDBP)
    ref_cat = CENTRAL_SBP_TABLE.get(cat, {}).get(sex, np.nan)
    if not np.isnan(ref_cat):
        p50 = max(p50, ref_cat)
        p75 = max(p75, p50 + 8)
        p90 = max(p90, p50 + 16)
        p95 = max(p95, p50 + 22)

    if csbp < p50:
        pct, cls = 40.0, "Presión central no aumentada"
    elif csbp < p75:
        pct = 50 + (csbp - p50) / max(p75 - p50, 1) * 25
        cls = "Presión central limítrofe / por encima de P50"
    elif csbp < p90:
        pct = 75 + (csbp - p75) / max(p90 - p75, 1) * 15
        cls = "Presión central aumentada"
    elif csbp < p95:
        pct = 90 + (csbp - p90) / max(p95 - p90, 1) * 5
        cls = "Presión central claramente aumentada"
    else:
        pct = min(99.0, 95 + (csbp - p95) / 10.0)
        cls = "Presión central severamente aumentada"
    return float(pct), cls, {"p50": p50, "p75": p75, "p90": p90, "p95": p95}


def central_diagnosis(row):
    cSBP = to_float(row.get("pas_central")); cPP = to_float(row.get("pp_central"))
    pSBP = to_float(row.get("pas_radial")); pDBP = to_float(row.get("pad_radial"))
    sex = (row.get("sexo") or "M").upper()[:1]
    cat = brachial_bp_category(pSBP, pDBP)
    ref = CENTRAL_SBP_TABLE.get(cat, {}).get(sex, np.nan)
    amp_sbp = pSBP - cSBP if not np.isnan(pSBP) and not np.isnan(cSBP) else np.nan
    ppa = (to_float(row.get("pp_radial")) / cPP) if cPP and not np.isnan(cPP) and cPP != 0 else np.nan
    pct, pct_class, pct_refs = estimate_central_percentile(row)

    if np.isnan(cSBP):
        dx = "No clasificable por falta de PAS central."
    elif cSBP >= 130 or (not np.isnan(pct) and pct >= 75):
        dx = f"{pct_class}: PAS central {cSBP:.0f} mmHg, percentilo operativo estimado P{pct:.0f}."
    elif not np.isnan(ref) and cSBP > ref:
        dx = f"{pct_class}: PAS central por encima de referencia central operativa para categoría braquial {cat}; percentilo estimado P{pct:.0f}."
    else:
        dx = f"{pct_class}: PAS central {cSBP:.0f} mmHg, percentilo operativo estimado P{pct:.0f}."

    risk = []
    if not np.isnan(ppa) and ppa < 1.30: risk.append("amplificación de presión de pulso reducida (<1,30)")
    if not np.isnan(cPP) and cPP >= 50: risk.append("presión de pulso central aumentada")
    iau_val = to_float(row.get("iau"))
    if not np.isnan(iau_val) and iau_val >= 25: risk.append("índice de aumentación elevado")
    if not np.isnan(pct) and pct >= 90: risk.append("PAS central sobre P90")
    return dx, cat, ref, amp_sbp, ppa, "; ".join(risk) if risk else "sin señales hemodinámicas mayores agregadas"

def make_waveform(row, n=512):
    """Genera una curva central personalizada cuando no se importa curva real.

    IMPORTANTE: antes la curva sintética tenía siempre la misma morfología, por eso
    los gráficos de separación de ondas y el análisis armónico podían verse iguales
    entre estudios. Ahora la forma cambia con PAS/PAD/PP central, Au, IAu, PE, FC,
    edad y amplificación periférico-central. Si se adjunta CSV/TXT de curva real,
    esta función no se usa.
    """
    cSBP = to_float(row.get("pas_central")); cDBP = to_float(row.get("pad_central"))
    cPP = to_float(row.get("pp_central")); iau = to_float(row.get("iau")); au = to_float(row.get("au"))
    pe_pct = to_float(row.get("pe")); fc = to_float(row.get("fc")); age = to_float(row.get("edad"))
    pp_radial = to_float(row.get("pp_radial"))

    if np.isnan(cSBP) or cSBP <= 0: cSBP = 120.0
    if np.isnan(cDBP) or cDBP <= 0: cDBP = 80.0
    if np.isnan(cPP) or cPP <= 0: cPP = max(cSBP - cDBP, 20.0)
    if cSBP <= cDBP: cSBP = cDBP + cPP
    if np.isnan(iau): iau = 15.0
    if np.isnan(au): au = max(cPP * iau / 100.0, 0.0)
    if np.isnan(pe_pct) or pe_pct <= 0: pe_pct = 38.0
    if np.isnan(fc) or fc <= 0: fc = 70.0
    if np.isnan(age) or age <= 0: age = 55.0

    # Eje temporal: usa el ciclo cardíaco por FC, limitado para que el gráfico sea comparable.
    cycle_ms = float(np.clip(60000.0 / fc, 650.0, 1200.0))
    t_ms = np.linspace(0, cycle_ms, n)
    t = t_ms / cycle_ms

    # Marcadores fisiológicos personalizados. Mayor IAu/Au/edad adelanta y aumenta la onda reflejada.
    eject_frac = float(np.clip((pe_pct / 100.0) * (cycle_ms / 1000.0), 0.26, 0.48))
    ppa = (pp_radial / cPP) if not np.isnan(pp_radial) and cPP > 0 else 1.30
    reflection_gain = float(np.clip(0.16 + max(iau, -10) / 95.0 + max(au, 0) / max(cPP * 2.5, 1), 0.10, 0.72))
    if not np.isnan(ppa) and ppa < 1.30:
        reflection_gain += min(0.18, (1.30 - ppa) * 0.35)
    reflection_gain = float(np.clip(reflection_gain, 0.10, 0.78))

    primary_center = float(np.clip(0.18 + (70.0 - fc) / 600.0, 0.14, 0.25))
    primary_width = float(np.clip(0.060 + (pe_pct - 30.0) / 800.0, 0.050, 0.105))
    reflected_center = float(np.clip(0.46 - (iau / 180.0) - (age - 55.0) / 900.0, 0.27, 0.58))
    reflected_width = float(np.clip(0.085 - min(max(iau, 0), 50) / 1000.0, 0.045, 0.110))

    primary = np.exp(-((t-primary_center)/primary_width)**2)
    reflected = reflection_gain * np.exp(-((t-reflected_center)/reflected_width)**2)
    shoulder = 0.10 * np.exp(-((t-(primary_center+0.10))/0.055)**2)
    diast = float(np.clip(0.18 - reflection_gain/8.0, 0.08, 0.22)) * np.exp(-((t-0.70)/0.22)**2)
    notch = -0.050 * np.exp(-((t-eject_frac)/0.025)**2)

    raw = primary + reflected + shoulder + diast + notch
    raw = raw - np.nanmin(raw)
    denom = np.nanmax(raw) - np.nanmin(raw)
    if denom <= 0: denom = 1.0
    p = cDBP + (cSBP - cDBP) * raw / denom

    out = pd.DataFrame({"tiempo_ms": t_ms, "presion_central_mmHg": p})
    out.attrs["source"] = "sintetica_personalizada"
    return out



def read_text_bytes_robust(uploaded_file):
    """Devuelve texto desde un archivo subido sin fallar por codificación.
    Soporta UTF-8, UTF-8 BOM, Latin-1, Windows-1252 y archivos generados por equipos antiguos.
    """
    if uploaded_file is None:
        return ""
    try:
        uploaded_file.seek(0)
    except Exception:
        pass
    raw = uploaded_file.read()
    if isinstance(raw, str):
        return raw
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1", "iso-8859-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
        except Exception:
            continue
    return raw.decode("latin-1", errors="replace")


def read_wave_file(uploaded_file):
    """Lee curva de presión desde CSV/TXT con codificación robusta.

    Acepta archivos con separadores por coma, punto y coma, tabulación, espacios
    o TXT sin encabezado. Evita el error UnicodeDecodeError de pandas leyendo
    primero los bytes y decodificando con fallback Latin-1/Windows-1252.
    """
    txt = read_text_bytes_robust(uploaded_file)
    if not txt.strip():
        raise ValueError("El archivo está vacío o no pudo leerse.")

    # Normalización: coma decimal solo cuando está entre números, preservando separadores.
    txt_norm = txt.replace("\ufeff", "")
    txt_norm = re.sub(r"(?<=\d),(?=\d)", ".", txt_norm)

    # Primer intento: separador automático amplio.
    attempts = [
        dict(sep=None, engine="python", comment="#"),
        dict(sep=r"[;\t]+", engine="python", comment="#"),
        dict(sep=r"[,;\t ]+", engine="python", comment="#"),
    ]
    last_error = None
    df = None
    for kwargs in attempts:
        try:
            candidate = pd.read_csv(io.StringIO(txt_norm), **kwargs)
            if candidate.shape[1] >= 2:
                df = candidate
                break
        except Exception as e:
            last_error = e

    # Fallback para TXT totalmente sin formato: extrae pares numéricos por línea.
    if df is None or df.shape[1] < 2:
        pairs = []
        for line in txt_norm.splitlines():
            nums = re.findall(r"[-+]?\d+(?:\.\d+)?", line)
            if len(nums) >= 2:
                pairs.append((nums[0], nums[1]))
        if len(pairs) >= 2:
            df = pd.DataFrame(pairs, columns=["tiempo_ms", "presion_central_mmHg"])
        else:
            raise ValueError(f"No se reconocieron dos columnas numéricas de curva. Último error: {last_error}")

    out = df.iloc[:, :2].copy()
    out.columns = ["tiempo_ms", "presion_central_mmHg"]
    out["tiempo_ms"] = pd.to_numeric(out["tiempo_ms"].astype(str).str.replace(",", ".", regex=False), errors="coerce")
    out["presion_central_mmHg"] = pd.to_numeric(out["presion_central_mmHg"].astype(str).str.replace(",", ".", regex=False), errors="coerce")
    out = out.dropna().sort_values("tiempo_ms").drop_duplicates("tiempo_ms")

    # Si el archivo trae índice de muestra 0..N en lugar de ms, lo escala a 0-1000 ms.
    if len(out) >= 20 and out["tiempo_ms"].max() <= len(out) + 5:
        out["tiempo_ms"] = np.linspace(0, 1000, len(out))

    if len(out) < 20:
        raise ValueError("La curva importada tiene menos de 20 puntos válidos. Revise que el TXT/CSV tenga tiempo y presión.")
    return out


def harmonic_analysis(wave_df):
    y = wave_df.iloc[:,1].astype(float).to_numpy()
    t = wave_df.iloc[:,0].astype(float).to_numpy()
    y0 = y - np.mean(y)
    fft = np.fft.rfft(y0)
    amp = np.abs(fft) / len(y0) * 2
    freq = np.fft.rfftfreq(len(y0), d=np.mean(np.diff(t))/1000.0)
    df = pd.DataFrame({"frecuencia_hz": freq, "amplitud": amp})
    df["energia_relativa_%"] = (df["amplitud"]**2) / np.sum(df["amplitud"]**2) * 100
    df = df.iloc[1:11].reset_index(drop=True)
    return df

def fig_to_png(fig):
    buf = io.BytesIO(); fig.savefig(buf, format="png", dpi=180, bbox_inches="tight"); plt.close(fig); buf.seek(0); return buf

def plot_pressure_comparison(row):
    labels = ["PAS", "PAD", "PAM", "PP"]
    radial = [row.get("pas_radial"), row.get("pad_radial"), row.get("pam_radial"), row.get("pp_radial")]
    central = [row.get("pas_central"), row.get("pad_central"), row.get("pam_central"), row.get("pp_central")]
    x = np.arange(len(labels)); w = 0.35
    fig, ax = plt.subplots(figsize=(7,4))
    ax.bar(x-w/2, radial, w, label="Radial/Braquial")
    ax.bar(x+w/2, central, w, label="Central")
    ax.set_xticks(x); ax.set_xticklabels(labels); ax.set_ylabel("mmHg"); ax.set_title("Presiones periféricas vs centrales"); ax.legend(); ax.grid(axis="y", alpha=.25)
    return fig_to_png(fig)

def plot_waveform(wave_df):
    """Gráfico principal de la onda de presión aórtica central en color rojo."""
    fig, ax = plt.subplots(figsize=(7,4))
    ax.plot(wave_df.iloc[:,0], wave_df.iloc[:,1], color="red", linewidth=2.4)
    ax.set_xlabel("Tiempo (ms)"); ax.set_ylabel("Presión central (mmHg)"); ax.set_title("Onda de presión aórtica central")
    ax.grid(alpha=.25)
    return fig_to_png(fig)


def estimate_aortic_flow_and_waves(wave_df, row):
    """
    Estima flujo aórtico y separa onda anterógrada/retrógrada desde presión sola.
    Usa un flujo eyectivo estimado cuando no se dispone de flujo medido.
    """
    t = wave_df.iloc[:, 0].astype(float).to_numpy()
    p = wave_df.iloc[:, 1].astype(float).to_numpy()
    if len(t) < 5:
        return pd.DataFrame(), {}
    t0 = t - np.nanmin(t)
    cycle_ms = np.nanmax(t0) if np.nanmax(t0) > 0 else 1000.0
    tn = t0 / cycle_ms * 1000.0

    dbp = to_float(row.get("pad_central")); sbp = to_float(row.get("pas_central")); pp = to_float(row.get("pp_central"))
    fc = to_float(row.get("fc")); sc = to_float(row.get("sc")); pe_pct = to_float(row.get("pe"))
    if np.isnan(dbp): dbp = float(np.nanmin(p))
    if np.isnan(sbp): sbp = float(np.nanmax(p))
    if np.isnan(pp) or pp <= 0: pp = max(sbp - dbp, 1.0)
    if np.isnan(fc) or fc <= 0: fc = 70.0
    if np.isnan(sc) or sc <= 0: sc = 1.8
    ejection_ms = 330.0 if np.isnan(pe_pct) or pe_pct <= 0 else float(np.clip(pe_pct * 10.0, 250.0, 450.0))

    flow_shape = np.zeros_like(tn, dtype=float)
    mask = (tn >= 0) & (tn <= ejection_ms)
    x = np.zeros_like(tn, dtype=float)
    x[mask] = tn[mask] / ejection_ms
    flow_shape[mask] = (np.sin(np.pi * x[mask]) ** 1.35) * np.exp(-0.30 * x[mask])
    if np.nanmax(flow_shape) <= 0:
        flow_shape[mask] = np.sin(np.pi * x[mask])

    sv_ml = float(np.clip(42.0 * sc, 45.0, 115.0))
    area_sec = safe_trapezoid_integral(flow_shape, tn / 1000.0)
    q_ml_s = flow_shape * (sv_ml / area_sec) if area_sec > 0 else flow_shape * 350.0
    gc_l_min = sv_ml * fc / 1000.0

    p_excess = np.maximum(p - dbp, 0)
    q_norm = q_ml_s / np.nanmax(q_ml_s) if np.nanmax(q_ml_s) > 0 else flow_shape
    zq = 0.72 * max(float(np.nanmax(p_excess)), 1.0)
    pf = 0.5 * (p_excess + zq * q_norm)
    pb = np.maximum(p_excess - pf, 0)
    total = pf + pb
    factor = np.divide(p_excess, total, out=np.ones_like(p_excess), where=total>0)
    pf = pf * factor
    pb = pb * factor

    i_pf = int(np.nanargmax(pf)) if len(pf) else 0
    i_pb = int(np.nanargmax(pb)) if len(pb) else 0
    pf_peak = float(np.nanmax(pf)) if len(pf) else np.nan
    pb_peak = float(np.nanmax(pb)) if len(pb) else np.nan
    rm = pb_peak / pf_peak if pf_peak and pf_peak > 0 else np.nan
    ri = pb_peak / (pf_peak + pb_peak) if (pf_peak + pb_peak) > 0 else np.nan
    metrics = {
        "flujo_pico_ml_s": float(np.nanmax(q_ml_s)),
        "tiempo_flujo_pico_ms": float(tn[int(np.nanargmax(q_ml_s))]),
        "volumen_sistolico_ml": sv_ml,
        "gasto_cardiaco_l_min": gc_l_min,
        "pf_pico_mmhg": pf_peak,
        "pb_pico_mmhg": pb_peak,
        "t_pf_ms": float(tn[i_pf]),
        "t_pb_ms": float(tn[i_pb]),
        "rm": float(rm) if not np.isnan(rm) else np.nan,
        "ri": float(ri) if not np.isnan(ri) else np.nan,
        "tfor_tref": float(tn[i_pf] / tn[i_pb]) if tn[i_pb] > 0 else np.nan,
        "ejection_ms": ejection_ms,
    }
    out = pd.DataFrame({
        "tiempo_ms": tn,
        "presion_total_exceso_mmhg": p_excess,
        "onda_anterograda_pf_mmhg": pf,
        "onda_retrograda_pb_mmhg": pb,
        "flujo_aortico_estimado_ml_s": q_ml_s,
    })
    return out, metrics


def plot_wave_separation(wave_df, row):
    sep, metrics = estimate_aortic_flow_and_waves(wave_df, row)
    fig, ax = plt.subplots(figsize=(7,4))
    if sep.empty:
        ax.text(0.5, 0.5, "No disponible", ha="center", va="center")
    else:
        ax.plot(sep["tiempo_ms"], sep["presion_total_exceso_mmhg"], color="green", linewidth=2.0, label="Onda medida (P total)")
        ax.plot(sep["tiempo_ms"], sep["onda_anterograda_pf_mmhg"], color="blue", linewidth=2.0, label="Onda anterógrada (Pf)")
        ax.plot(sep["tiempo_ms"], sep["onda_retrograda_pb_mmhg"], color="red", linestyle="--", linewidth=2.0, label="Onda retrógrada (Pb)")
        ax.legend(fontsize=8)
    ax.set_xlabel("Tiempo (ms)"); ax.set_ylabel("Presión sobre PAD (mmHg)")
    ax.set_title("Separación estimada de ondas de presión")
    ax.grid(alpha=.25)
    return fig_to_png(fig)


def plot_aortic_flow(wave_df, row):
    sep, metrics = estimate_aortic_flow_and_waves(wave_df, row)
    fig, ax = plt.subplots(figsize=(7,4))
    if sep.empty:
        ax.text(0.5, 0.5, "No disponible", ha="center", va="center")
    else:
        ax.plot(sep["tiempo_ms"], sep["flujo_aortico_estimado_ml_s"], color="purple", linewidth=2.2)
        ax.axvline(metrics.get("ejection_ms", np.nan), linestyle="--", linewidth=1)
    ax.set_xlabel("Tiempo (ms)"); ax.set_ylabel("Flujo (mL/s)")
    ax.set_title("Onda de flujo aórtico estimada")
    ax.grid(alpha=.25)
    return fig_to_png(fig)

def plot_harmonics(hdf):
    fig, ax = plt.subplots(figsize=(7,4))
    ax.bar(range(1, len(hdf)+1), hdf["energia_relativa_%"])
    ax.set_xlabel("Armónico"); ax.set_ylabel("Energía relativa (%)"); ax.set_title("Análisis armónico de la onda de presión central")
    ax.grid(axis="y", alpha=.25)
    return fig_to_png(fig)

def plot_clinical_gauges(row, ppa):
    metrics = {
        "PAS central": (to_float(row.get("pas_central")), 130),
        "PP central": (to_float(row.get("pp_central")), 50),
        "IAu": (to_float(row.get("iau")), 25),
        "PPA": (ppa, 1.30),
    }
    fig, ax = plt.subplots(figsize=(7,3.8))
    names = list(metrics.keys())
    vals = [metrics[k][0] for k in names]
    refs = [metrics[k][1] for k in names]
    score = []
    for n,v,r in zip(names, vals, refs):
        if np.isnan(v): score.append(np.nan)
        elif n == "PPA": score.append(v/r)
        else: score.append(v/r)
    ax.barh(names, score)
    ax.axvline(1, linestyle="--", linewidth=1)
    ax.set_xlabel("Relación valor / umbral")
    ax.set_title("Semaforización clínica de parámetros centrales")
    ax.grid(axis="x", alpha=.25)
    return fig_to_png(fig)


def interpret_wave_separation(metrics, row):
    """Redacción clínica automática de separación de ondas según patrones publicados.
    Usa RM=Pb/Pf y RI=Pb/(Pf+Pb) (Westerhof), y Tfor/Tref como descriptor temporal
    con valor pronóstico explorado en IDCARS/Norton.
    """
    if not metrics:
        return "No se pudo realizar interpretación de separación de ondas por ausencia de curva válida."
    pf = metrics.get("pf_pico_mmhg", np.nan)
    pb = metrics.get("pb_pico_mmhg", np.nan)
    rm = metrics.get("rm", np.nan)
    ri = metrics.get("ri", np.nan)
    tpf = metrics.get("t_pf_ms", np.nan)
    tpb = metrics.get("t_pb_ms", np.nan)
    ratio = metrics.get("tfor_tref", np.nan)
    iau = to_float(row.get("iau"))
    au = to_float(row.get("au"))

    patrones = []
    if not np.isnan(rm):
        if rm >= 0.55:
            patrones.append("magnitud de reflexión elevada (RM ≥0,55), compatible con onda retrógrada de gran contribución a la carga pulsátil central")
        elif rm >= 0.35:
            patrones.append("magnitud de reflexión intermedia (RM 0,35-0,54)")
        else:
            patrones.append("magnitud de reflexión baja o conservada (RM <0,35)")
    if not np.isnan(ri):
        if ri >= 0.35:
            patrones.append("índice de reflexión elevado (RI ≥0,35)")
        elif ri >= 0.25:
            patrones.append("índice de reflexión intermedio (RI 0,25-0,34)")
        else:
            patrones.append("índice de reflexión bajo (RI <0,25)")
    if not np.isnan(tpb):
        if tpb <= 250:
            patrones.append("retorno precoz de la onda reflejada, sugerente de mayor rigidez arterial efectiva o sitios de reflexión más próximos")
        elif tpb <= 350:
            patrones.append("retorno sistólico medio de la onda reflejada")
        else:
            patrones.append("retorno tardío de la onda reflejada, perfil más compatible con mejor complacencia arterial")
    if not np.isnan(ratio):
        if ratio >= 0.60:
            patrones.append("relación Tfor/Tref elevada, descriptor temporal asociado en estudios poblacionales a mayor carga pulsátil y riesgo")
        else:
            patrones.append("relación Tfor/Tref no elevada")
    if not np.isnan(iau):
        if iau >= 25:
            patrones.append("IAu elevado, consistente con aumento de aumentación central")
        elif iau >= 10:
            patrones.append("IAu intermedio")
        else:
            patrones.append("IAu bajo o conservado")

    conclusion = "La separación de ondas estimada muestra Pf pico "
    conclusion += f"{pf:.1f} mmHg y Pb pico {pb:.1f} mmHg. " if not np.isnan(pf) and not np.isnan(pb) else "valores Pf/Pb no disponibles. "
    conclusion += "Interpretación: " + "; ".join(patrones) + ". "
    conclusion += "Estos parámetros deben interpretarse como estimaciones derivadas de presión sola cuando no hay flujo aórtico medido; su utilidad principal es integrar carga anterógrada, componente reflejado, temporización de la reflexión y poscarga ventricular."
    return conclusion


def interpret_harmonic_analysis(hdf):
    """Interpreta la distribución energética armónica de la onda central."""
    if hdf is None or hdf.empty:
        return "No se pudo interpretar el análisis armónico por ausencia de datos espectrales."
    energy = hdf["energia_relativa_%"].astype(float).to_numpy()
    freq = hdf["frecuencia_hz"].astype(float).to_numpy()
    total = np.nansum(energy)
    if total <= 0:
        return "El espectro no mostró energía suficiente para interpretación armónica confiable."
    e1 = energy[0] if len(energy) > 0 else np.nan
    e2e3 = np.nansum(energy[1:3]) if len(energy) >= 3 else np.nan
    e_high = np.nansum(energy[3:]) if len(energy) > 3 else 0.0
    centroid = float(np.nansum(freq * energy) / total)
    cum = np.cumsum(energy) / total * 100
    bf_idx = int(np.argmax(cum >= 95)) if np.any(cum >= 95) else len(freq)-1
    break_freq = float(freq[bf_idx]) if len(freq) else np.nan

    if e_high >= 25:
        patron = "aumento relativo de armónicos altos, compatible con onda más angosta, mayor contenido de alta frecuencia y posible rigidez/reflexión temprana"
    elif e1 >= 65:
        patron = "predominio del armónico fundamental, compatible con morfología más suave y menor complejidad espectral"
    else:
        patron = "distribución armónica intermedia, sin predominio extremo de alta frecuencia"
    if centroid >= 4:
        patron += "; centroide espectral alto, lo que refuerza la sospecha de mayor carga pulsátil de alta frecuencia"
    elif centroid < 2.5:
        patron += "; centroide espectral bajo, compatible con menor componente pulsátil de alta frecuencia"

    return (f"El análisis armónico muestra energía del primer armónico {e1:.1f}%, energía de segundo-tercer armónico {e2e3:.1f}% "
            f"y energía en armónicos altos {e_high:.1f}%. La frecuencia de ruptura operativa, definida por el 95% de energía acumulada, "
            f"es {break_freq:.2f} Hz; centroide espectral {centroid:.2f} Hz. Interpretación: {patron}.")


def build_integrated_clinical_text(row, wave_metrics, hdf):
    dx, cat, ref, amp_sbp, ppa, risk = central_diagnosis(row)
    pct, pct_class, pct_refs = estimate_central_percentile(row)
    csbp = to_float(row.get("pas_central")); cpp = to_float(row.get("pp_central"))
    ptxt = "no disponible" if np.isnan(pct) else f"P{pct:.0f}"
    ref_txt = f"P50 {pct_refs['p50']:.0f}, P75 {pct_refs['p75']:.0f}, P90 {pct_refs['p90']:.0f}, P95 {pct_refs['p95']:.0f} mmHg" if not np.isnan(pct_refs.get('p50', np.nan)) else "referencias no disponibles"
    central_txt = (f"Diagnóstico de presión central: {pct_class}. La PAS central medida fue {csbp:.0f} mmHg "
                   f"y corresponde aproximadamente al percentilo operativo {ptxt} para edad/sexo/categoría braquial. "
                   f"Referencias usadas por la app: {ref_txt}. PP central: {cpp:.0f} mmHg. "
                   f"Amplificación periférico-central de PAS: {amp_sbp:.1f} mmHg; PPA: {ppa:.2f}. Perfil agregado: {risk}.")
    return central_txt, interpret_wave_separation(wave_metrics, row), interpret_harmonic_analysis(hdf)

def build_pdf(row, wave_df, hdf, screenshot_png=None):
    dx, cat, ref, amp_sbp, ppa, risk = central_diagnosis(row)
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=14*mm, leftMargin=14*mm, topMargin=12*mm, bottomMargin=12*mm)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Small", fontSize=8, leading=10))
    styles.add(ParagraphStyle(name="TitlePAC", fontSize=16, leading=20, alignment=1, textColor=colors.HexColor("#17365D")))
    story = []
    story.append(Paragraph("PRESIÓN AÓRTICA CENTRAL - INFORME MÉDICO INTEGRADO", styles["TitlePAC"]))
    story.append(Spacer(1, 5*mm))
    datos = [
        ["Paciente", row.get("paciente",""), "Estudio", row.get("estudio","")],
        ["Fecha", row.get("fecha",""), "Hora", row.get("hora","")],
        ["Edad", row.get("edad",""), "Sexo", row.get("sexo","")],
        ["Peso", row.get("peso",""), "Altura", row.get("altura","")],
        ["IMC", row.get("imc",""), "Medicación", row.get("medicacion","")],
    ]
    story.append(Table(datos, colWidths=[28*mm,55*mm,28*mm,55*mm], style=[("GRID",(0,0),(-1,-1),.25,colors.grey), ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#EAF2F8")), ("FONT",(0,0),(-1,-1),"Helvetica",8)]))
    story.append(Spacer(1, 4*mm))
    vals = [["Variable", "Radial/Braquial", "Central", "Unidad"],
            ["PAS", row.get("pas_radial"), row.get("pas_central"), "mmHg"],
            ["PAD", row.get("pad_radial"), row.get("pad_central"), "mmHg"],
            ["PAM", row.get("pam_radial"), row.get("pam_central"), "mmHg"],
            ["PP", row.get("pp_radial"), row.get("pp_central"), "mmHg"],
            ["FC", row.get("fc"), "", "lpm"],
            ["Au", "", row.get("au"), "mmHg"],
            ["IAu", "", row.get("iau"), "%"],
            ["RVSE", "", row.get("rvse"), "%"],
            ["PE", "", row.get("pe"), "%"]]
    story.append(Paragraph("Valores hemodinámicos centrales", styles["Heading2"]))
    story.append(Table(vals, colWidths=[40*mm,42*mm,42*mm,30*mm], style=[("GRID",(0,0),(-1,-1),.25,colors.grey), ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#D9EAF7")), ("FONT",(0,0),(-1,-1),"Helvetica",8)]))
    story.append(Spacer(1, 4*mm))
    wave_sep_df, wave_sep_metrics = estimate_aortic_flow_and_waves(wave_df, row)
    central_txt, wave_txt, harmonic_txt = build_integrated_clinical_text(row, wave_sep_metrics, hdf)
    story.append(Paragraph("Conclusión clínica integrada", styles["Heading2"]))
    story.append(Paragraph(central_txt, styles["BodyText"]))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph("Interpretación de separación de ondas", styles["Heading3"]))
    story.append(Paragraph(wave_txt, styles["BodyText"]))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph("Interpretación del análisis armónico", styles["Heading3"]))
    story.append(Paragraph(harmonic_txt, styles["BodyText"]))
    story.append(Spacer(1, 4*mm))
    for title, png in [
        ("Gráfico comparativo de presiones", plot_pressure_comparison(row)),
        ("Onda de presión aórtica central", plot_waveform(wave_df)),
        ("Separación estimada de onda anterógrada y retrógrada", plot_wave_separation(wave_df, row)),
        ("Onda de flujo aórtico estimada", plot_aortic_flow(wave_df, row)),
        ("Análisis armónico", plot_harmonics(hdf)),
        ("Semaforización clínica", plot_clinical_gauges(row, ppa)),
    ]:
        story.append(KeepTogether([Paragraph(title, styles["Heading3"]), Image(png, width=170*mm, height=90*mm)]))
    if wave_sep_metrics:
        story.append(Paragraph("Parámetros derivados de separación de ondas y flujo estimado", styles["Heading2"]))
        mtab = [["Parámetro", "Valor", "Unidad"],
                ["Pf pico", f"{wave_sep_metrics.get('pf_pico_mmhg', np.nan):.1f}", "mmHg"],
                ["Pb pico", f"{wave_sep_metrics.get('pb_pico_mmhg', np.nan):.1f}", "mmHg"],
                ["RM Pb/Pf", f"{wave_sep_metrics.get('rm', np.nan):.2f}", ""],
                ["RI Pb/(Pf+Pb)", f"{wave_sep_metrics.get('ri', np.nan):.2f}", ""],
                ["Tfor/Tref", f"{wave_sep_metrics.get('tfor_tref', np.nan):.2f}", ""],
                ["Flujo pico estimado", f"{wave_sep_metrics.get('flujo_pico_ml_s', np.nan):.0f}", "mL/s"],
                ["Volumen sistólico estimado", f"{wave_sep_metrics.get('volumen_sistolico_ml', np.nan):.1f}", "mL"],
                ["Gasto cardíaco estimado", f"{wave_sep_metrics.get('gasto_cardiaco_l_min', np.nan):.2f}", "L/min"]]
        story.append(Table(mtab, colWidths=[70*mm,40*mm,35*mm], style=[("GRID",(0,0),(-1,-1),.25,colors.grey), ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#D9EAF7")), ("FONT",(0,0),(-1,-1),"Helvetica",8)]))
        story.append(Paragraph(wave_txt, styles["BodyText"]))
        story.append(Paragraph("Nota: flujo y separación de ondas son estimaciones derivadas de presión sola con modelo de flujo eyectivo, útiles para interpretación clínica integrada cuando no existe medición directa de flujo aórtico.", styles["Small"]))
    story.append(PageBreak())
    story.append(Paragraph("Análisis armónico de la onda de presión central", styles["Heading2"]))
    story.append(Paragraph("Se calcula por transformada rápida de Fourier sobre la onda central importada o, si no se adjunta curva digitalizada, sobre una curva sintética calibrada con la PAS/PAD central. Su utilidad clínica es describir la distribución de energía pulsátil, los componentes de alta frecuencia y la posible contribución de rigidez arterial/reflexiones tempranas.", styles["BodyText"]))
    htab = [["Armónico", "Frecuencia (Hz)", "Amplitud", "Energía relativa (%)"]] + [[i+1, f"{r.frecuencia_hz:.2f}", f"{r.amplitud:.2f}", f"{r.energia_relativa_:.2f}" if False else f"{r['energia_relativa_%']:.2f}"] for i,r in hdf.iterrows()]
    story.append(Table(htab, colWidths=[28*mm,42*mm,42*mm,48*mm], style=[("GRID",(0,0),(-1,-1),.25,colors.grey), ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#D9EAF7")), ("FONT",(0,0),(-1,-1),"Helvetica",8)]))
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(harmonic_txt, styles["BodyText"]))
    if screenshot_png:
        story.append(PageBreak())
        story.append(Paragraph("CAPTURA PANTALLA DE MEDICIONES - segunda hoja del estudio original", styles["Heading2"]))
        story.append(Image(io.BytesIO(screenshot_png), width=170*mm, height=230*mm))
    story.append(PageBreak())
    story.append(Paragraph("Referencias bibliográficas", styles["Heading2"]))
    for i, refb in enumerate(BIBLIOGRAFIA, 1):
        story.append(Paragraph(f"{i}. {refb}", styles["Small"]))
    doc.build(story)
    buf.seek(0)
    return buf.getvalue()

def save_history(row):
    new = pd.DataFrame([row])
    if HISTORIAL_FILE.exists():
        old = pd.read_excel(HISTORIAL_FILE)
        out = pd.concat([old, new], ignore_index=True)
    else:
        out = new
    out.to_excel(HISTORIAL_FILE, index=False)
    return out

st.title(APP_TITLE)
st.caption("Importación tipo MODELO PAC en PDF o métricas en TXT estructurado/sin formato, informe PDF integrado, captura de segunda hoja, historial Excel y análisis armónico.")

with st.sidebar:
    st.header("1) Importar estudio")
    pdf_file = st.file_uploader("PDF original PAC / Exxer", type=["pdf"])
    txt_metrics_file = st.file_uploader("Opcional: CSV/TXT con métricas del estudio", type=["csv", "txt"])
    wave_file = st.file_uploader("Opcional: CSV/TXT curva central (tiempo_ms, presion_mmHg)", type=["csv", "txt"], key="wave_file")
    st.info("Si la extracción automática no detecta algún dato, corríjalo manualmente antes de generar el PDF.")

base = {}
screenshot = None
if pdf_file:
    pdf_bytes = pdf_file.read()
    text = extract_pdf_text(pdf_bytes)
    base = parse_model_pac(text)
    screenshot = render_pdf_page_png(pdf_bytes, page_index=1)
else:
    base = parse_model_pac("")

# El TXT puede usarse solo o complementar/corregir lo extraído del PDF.
if txt_metrics_file:
    try:
        name = getattr(txt_metrics_file, "name", "").lower()
        if name.endswith(".csv"):
            txt_data = parse_metric_csv(txt_metrics_file)
            msg_tipo = "CSV"
        else:
            txt_text = decode_uploaded_text(txt_metrics_file)
            txt_data = parse_metric_txt(txt_text)
            msg_tipo = "TXT"
        for k, v in txt_data.items():
            missing_in_base = (k not in base) or base.get(k) == "" or (isinstance(base.get(k), float) and np.isnan(base.get(k)))
            value_available = not (v == "" or (isinstance(v, float) and np.isnan(v)))
            if value_available and (missing_in_base or txt_metrics_file is not None):
                base[k] = v
        st.sidebar.success(f"{msg_tipo} de métricas importado con codificación robusta y aplicado a las variables correspondientes.")
    except Exception as e:
        st.sidebar.error(f"No se pudo importar el archivo de métricas CSV/TXT: {e}")

st.subheader("Datos extraídos / edición manual")
cols = st.columns(4)
fields = ["paciente","estudio","fecha","hora","edad","sexo","peso","altura","imc","pas_radial","pad_radial","pam_radial","pp_radial","pas_central","pad_central","pam_central","pp_central","fc","au","iau","rvse","pe","medicacion","diagnostico_previo"]
row = {}
for i, f in enumerate(fields):
    with cols[i%4]:
        val = base.get(f, "")
        if f in ["paciente","estudio","fecha","hora","sexo","medicacion","diagnostico_previo"]:
            row[f] = st.text_input(f, value="" if pd.isna(val) else str(val))
        else:
            row[f] = st.number_input(f, value=0.0 if pd.isna(val) else float(val), step=1.0, format="%.2f")

if wave_file:
    try:
        wave_df = read_wave_file(wave_file)
        st.success(f"Curva importada: {len(wave_df)} puntos válidos.")
    except Exception as e:
        st.error(f"No se pudo importar la curva CSV/TXT: {e}")
        wave_df = make_waveform(row)
else:
    wave_df = make_waveform(row)

hdf = harmonic_analysis(wave_df)
dx, cat, ref, amp_sbp, ppa, risk = central_diagnosis(row)

st.subheader("Vista clínica previa")
st.write(dx)
st.write(f"Categoría braquial: {cat} | Amplificación PAS: {amp_sbp:.1f} mmHg | PPA: {ppa:.2f} | Perfil: {risk}")

g1, g2 = st.columns(2)
with g1:
    st.image(plot_pressure_comparison(row), caption="Presiones periféricas vs centrales")
    st.image(plot_harmonics(hdf), caption="Armónicos de la onda central")
with g2:
    st.image(plot_waveform(wave_df), caption="Onda central en rojo")
    st.image(plot_clinical_gauges(row, ppa), caption="Semaforización clínica")

st.subheader("Separación de ondas y flujo aórtico estimado")
sep_df, sep_metrics = estimate_aortic_flow_and_waves(wave_df, row)
sg1, sg2 = st.columns(2)
with sg1:
    st.image(plot_wave_separation(wave_df, row), caption="Onda anterógrada Pf, retrógrada Pb y onda total")
with sg2:
    st.image(plot_aortic_flow(wave_df, row), caption="Curva de flujo aórtico estimada")
if sep_metrics:
    st.dataframe(pd.DataFrame([sep_metrics]), use_container_width=True)

central_txt, wave_txt, harmonic_txt = build_integrated_clinical_text(row, sep_metrics, hdf)
st.subheader("Redacción automática para informe médico integrado")
st.markdown("**Diagnóstico de presión central y percentilo**")
st.write(central_txt)
st.markdown("**Interpretación de separación de ondas**")
st.write(wave_txt)
st.markdown("**Interpretación del análisis armónico**")
st.write(harmonic_txt)

st.subheader("Historial y exportación")
if st.button("Guardar en historial"):
    hist = save_history(row)
    st.success(f"Registro guardado. Total: {len(hist)} estudios.")

if HISTORIAL_FILE.exists():
    hist = pd.read_excel(HISTORIAL_FILE)
    st.dataframe(hist, use_container_width=True)
    st.download_button("Descargar historial Excel", HISTORIAL_FILE.read_bytes(), file_name="historial_pac.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

pdf_bytes_out = build_pdf(row, wave_df, hdf, screenshot)
st.download_button("Generar y descargar PDF médico integrado", pdf_bytes_out, file_name=f"PAC_IA_{row.get('paciente','paciente').replace(' ','_')}.pdf", mime="application/pdf")
