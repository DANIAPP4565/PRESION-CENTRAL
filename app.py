# app_pac.py
# App Streamlit para informe de Presión Aórtica Central (PAC)
# Importa PDF tipo MODELO PAC, extrae variables, genera historial Excel y PDF médico integrado.

import io, re, math, tempfile
from datetime import datetime
from pathlib import Path

import numpy as np
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

def ensure_download_bytes(obj):
    """Convierte salida de PDF a bytes válidos para st.download_button."""
    if obj is None:
        return b""
    if isinstance(obj, bytes):
        return obj
    if isinstance(obj, bytearray):
        return bytes(obj)
    if hasattr(obj, "getvalue"):
        val = obj.getvalue()
        if isinstance(val, str):
            return val.encode("latin-1", errors="ignore")
        return bytes(val)
    if isinstance(obj, str):
        # Si es una ruta a archivo, leer bytes. Si es contenido, codificar.
        try:
            p = Path(obj)
            if p.exists() and p.is_file():
                return p.read_bytes()
        except Exception:
            pass
        return obj.encode("latin-1", errors="ignore")
    try:
        return bytes(obj)
    except Exception:
        return str(obj).encode("latin-1", errors="ignore")


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




def is_physiologic_waveform(df, row=None):
    """Valida que la curva tenga morfología fisiológica y no sea una tabla de métricas mal leída."""
    try:
        if df is None or len(df) < 20:
            return False, "menos de 20 puntos útiles"
        t = pd.to_numeric(df.iloc[:, 0], errors="coerce").to_numpy(dtype=float)
        p = pd.to_numeric(df.iloc[:, 1], errors="coerce").to_numpy(dtype=float)
        ok = np.isfinite(t) & np.isfinite(p)
        t, p = t[ok], p[ok]
        if len(p) < 20:
            return False, "menos de 20 pares tiempo-presión"
        if np.nanmax(t) - np.nanmin(t) < 250:
            return False, "duración menor a 250 ms"
        pmin, pmax = float(np.nanmin(p)), float(np.nanmax(p))
        pp = pmax - pmin
        if not (35 <= pmin <= 140 and 70 <= pmax <= 240 and 10 <= pp <= 120):
            return False, f"rango no fisiológico: mínimo {pmin:.1f}, máximo {pmax:.1f}"
        # Evitar curvas con saltos erráticos como las de la captura.
        dif = np.abs(np.diff(p))
        if len(dif) > 0 and np.nanpercentile(dif, 95) > max(18, 0.45 * pp):
            return False, "saltos bruscos no compatibles con onda de presión central"
        # Debe tener un ascenso sistólico claro.
        peak_i = int(np.nanargmax(p))
        if peak_i < 2 or peak_i > int(len(p) * 0.75):
            return False, "pico sistólico mal ubicado"
        if (pmax - p[0]) < 0.35 * pp:
            return False, "no hay ascenso sistólico claro"
        return True, "curva fisiológica"
    except Exception as e:
        return False, str(e)


def calibrate_waveform_to_metrics(wave_df, row):
    """Escala la curva para que coincida exactamente con PAD/PAS central del estudio."""
    df = wave_df.copy()
    df.columns = ["tiempo_ms", "presion_central_mmHg"]
    df = df.replace([np.inf, -np.inf], np.nan).dropna()
    df = df.sort_values("tiempo_ms").drop_duplicates("tiempo_ms").reset_index(drop=True)

    pas = to_float(row.get("pas_central"))
    pad = to_float(row.get("pad_central"))
    if np.isnan(pas) or pas <= 0:
        pas = float(df["presion_central_mmHg"].max())
    if np.isnan(pad) or pad <= 0:
        pad = float(df["presion_central_mmHg"].min())
    if pas <= pad:
        pas = pad + max(25.0, to_float(row.get("pp_central")) if not np.isnan(to_float(row.get("pp_central"))) else 35.0)

    y = pd.to_numeric(df["presion_central_mmHg"], errors="coerce").to_numpy(dtype=float)
    y = pd.Series(y).interpolate(limit_direction="both").to_numpy(dtype=float)

    # Suavizado conservador para preservar morfología y quitar dientes del TXT/CSV.
    if len(y) >= 9:
        win = max(5, min(31, (len(y)//20)*2 + 1))
        y = pd.Series(y).rolling(win, center=True, min_periods=1).median().to_numpy()
        y = pd.Series(y).rolling(win, center=True, min_periods=1).mean().to_numpy()

    ymin, ymax = float(np.nanmin(y)), float(np.nanmax(y))
    if ymax - ymin < 1:
        return make_waveform(row)

    ycal = pad + (y - ymin) * (pas - pad) / (ymax - ymin)

    t = pd.to_numeric(df["tiempo_ms"], errors="coerce").to_numpy(dtype=float)
    t = pd.Series(t).interpolate(limit_direction="both").to_numpy(dtype=float)
    if np.nanmax(t) - np.nanmin(t) <= 0:
        t = np.linspace(0, 1000, len(ycal))
    else:
        t = (t - np.nanmin(t)) / (np.nanmax(t) - np.nanmin(t)) * 1000.0

    return pd.DataFrame({"tiempo_ms": t, "presion_central_mmHg": ycal})


def read_curve_file_robust(uploaded_file, row=None):
    """Lee CSV/TXT de curva y rechaza tablas de métricas mal interpretadas como curva.

    Acepta columnas nombradas como tiempo/time/ms y presión/pressure/PAC/mmHg.
    También acepta TXT/CSV sin encabezado si contiene una serie real de al menos 20-50 puntos.
    Si el archivo no contiene una curva fisiológica, se usa curva sintética calibrada.
    """
    raw = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()

    text = None
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1", "iso-8859-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = raw.decode("latin-1", errors="replace")

    text = text.replace("\x00", "").replace("\ufeff", "").strip()
    if not text:
        raise ValueError("El archivo de curva está vacío.")

    candidates = []

    # CSV/TXT estructurado con separadores habituales.
    for sep in (None, ";", ",", "\t", r"\s+"):
        try:
            df = pd.read_csv(io.StringIO(text), sep=sep, engine="python", decimal=",", on_bad_lines="skip")
            if df is not None and df.shape[0] >= 20:
                try:
                    candidates.append(normalize_wave_dataframe(df))
                except Exception:
                    pass
        except Exception:
            pass

    # Sin encabezado: solo si hay muchos pares y tiempo monótono.
    numeric_rows = []
    for line in text.splitlines():
        nums = re.findall(r"[-+]?\d+(?:[\.,]\d+)?", line)
        vals = [to_float(n) for n in nums]
        vals = [v for v in vals if not np.isnan(v)]
        if vals:
            numeric_rows.append(vals)

    pair_rows = [r[:2] for r in numeric_rows if len(r) >= 2]
    if len(pair_rows) >= 20:
        try:
            candidates.append(normalize_wave_dataframe(pd.DataFrame(pair_rows, columns=["tiempo_ms", "presion_mmHg"])))
        except Exception:
            pass

    # Una columna / vector de presión: solo si hay suficientes puntos fisiológicos.
    flat = []
    for r in numeric_rows:
        if len(r) == 1:
            flat.append(r[0])
    if len(flat) >= 50:
        pressure_like = [v for v in flat if 35 <= v <= 240]
        if len(pressure_like) >= 50:
            candidates.append(pd.DataFrame({"tiempo_ms": np.linspace(0, 1000, len(pressure_like)), "presion_central_mmHg": pressure_like}))

    errors = []
    for cand in candidates:
        cand = calibrate_waveform_to_metrics(cand, row or {})
        ok, msg = is_physiologic_waveform(cand, row)
        if ok:
            return cand
        errors.append(msg)

    raise ValueError("El archivo no contiene una curva de presión central fisiológica reconocible. " + ("; ".join(errors[:3]) if errors else ""))


def normalize_wave_dataframe(df):
    """Normaliza cualquier tabla de curva a tiempo_ms/presion_central_mmHg."""
    df = df.copy()

    if df.shape[1] == 1:
        col = df.columns[0]
        joined = "\n".join(df[col].astype(str).tolist())
        for sep in (";", ",", "\t", r"\s+"):
            try:
                tmp = pd.read_csv(io.StringIO(joined), sep=sep, engine="python", header=None, on_bad_lines="skip")
                if tmp.shape[1] >= 2 and tmp.shape[0] >= 20:
                    df = tmp
                    break
            except Exception:
                pass

    df.columns = [str(c).strip().lower() for c in df.columns]

    num = pd.DataFrame()
    for c in df.columns:
        num[c] = df[c].map(to_float)

    valid_numeric_cols = [c for c in num.columns if num[c].notna().sum() >= 20]
    if len(valid_numeric_cols) < 2:
        # una sola columna con presión
        valid_one = [c for c in num.columns if num[c].notna().sum() >= 50]
        if len(valid_one) == 1:
            pressure = num[valid_one[0]].dropna().astype(float)
            pressure = pressure[(pressure >= 35) & (pressure <= 240)]
            if len(pressure) >= 50:
                return pd.DataFrame({"tiempo_ms": np.linspace(0, 1000, len(pressure)), "presion_central_mmHg": pressure.values})
        raise ValueError("No se detectaron columnas de curva con suficientes puntos.")

    time_keys = ["tiempo", "time", "ms", "mseg", "miliseg", "seg", "sec", "x"]
    press_keys = ["pres", "pressure", "pao", "pac", "central", "mmhg", "aort", "ao", "y"]

    time_candidates = [c for c in valid_numeric_cols if any(k in str(c).lower() for k in time_keys)]
    pressure_candidates = [c for c in valid_numeric_cols if any(k in str(c).lower() for k in press_keys)]

    if time_candidates:
        tcol = time_candidates[0]
    else:
        # Escoger columna más monótona y con rango compatible con tiempo.
        scores = {}
        for c in valid_numeric_cols:
            s = num[c].dropna().astype(float).to_numpy()
            diffs = np.diff(s)
            mono = np.mean(diffs >= 0)
            rng = np.nanmax(s) - np.nanmin(s)
            scores[c] = mono + (0.5 if rng >= 250 else 0)
        tcol = max(scores, key=scores.get)

    pcol = None
    for c in pressure_candidates:
        if c != tcol:
            pcol = c
            break

    if pcol is None:
        # Escoger columna con rango de presión fisiológico.
        ranked = []
        for c in valid_numeric_cols:
            if c == tcol:
                continue
            s = num[c].dropna().astype(float)
            med = float(np.nanmedian(s)); rng = float(np.nanmax(s)-np.nanmin(s))
            score = 0
            if 50 <= med <= 160: score += 2
            if 10 <= rng <= 120: score += 2
            ranked.append((score, c))
        if not ranked:
            raise ValueError("No se detectó columna de presión.")
        pcol = sorted(ranked, reverse=True)[0][1]

    out = pd.DataFrame({"tiempo_ms": num[tcol], "presion_central_mmHg": num[pcol]}).dropna()
    out = out.sort_values("tiempo_ms").drop_duplicates("tiempo_ms").reset_index(drop=True)

    if len(out) < 20:
        raise ValueError("Curva con menos de 20 puntos válidos.")

    if out["tiempo_ms"].max() <= 5:
        out["tiempo_ms"] *= 1000.0

    tmin, tmax = out["tiempo_ms"].min(), out["tiempo_ms"].max()
    if tmax - tmin <= 0:
        out["tiempo_ms"] = np.linspace(0, 1000, len(out))
    elif tmax > 5000 or tmax < 250:
        out["tiempo_ms"] = (out["tiempo_ms"] - tmin) / (tmax - tmin) * 1000.0

    return out

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

def parse_model_pac(text):
    # Parser orientado al patrón MODELO PAC / Exxer. Incluye fallback manual en la interfaz.
    lines = [safe_text(x) for x in text.splitlines() if safe_text(x)]
    joined = "\n".join(lines)
    data = {
        "paciente": find_after(r"Paciente|Nombre del paciente", joined),
        "estudio": find_after(r"Estudio|Número de estudio", joined),
        "fecha": find_after(r"Fecha", joined),
        "hora": find_after(r"Hora", joined),
        "hc": find_after(r"H\.C\.", joined),
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
    # Patrón de tabla Radial Central del modelo: PAS PAD PAM PP en filas.
    m = re.search(r"Radial\s+Central\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)", joined, re.S|re.I)
    if m:
        vals = list(map(float, m.groups()))
        data.update({
            "pas_radial": vals[0], "pas_central": vals[1],
            "pad_radial": vals[2], "pad_central": vals[3],
            "pp_radial": vals[4], "pp_central": vals[5],
            "pam_radial": vals[6], "pam_central": vals[7],
        })
    # Parámetros centrales alrededor de PAS PP Au IAu RVSE PE APC.
    m2 = re.search(r"Parámetros hemodinámicos centrales.*?PAS\s+PP\s+Au\s+IAu\s+PE\s+(\d+)\s+(\d+)\s+([+\-]?\d+)\s+([+\-]?\d+)\s+(\d+)\s+([\d.,]+)", joined, re.S|re.I)
    if m2:
        data.update({
            "pas_central": to_float(m2.group(1)),
            "pp_central": to_float(m2.group(2)),
            "au": to_float(m2.group(3)),
            "iau": to_float(m2.group(4)),
            "rvse": to_float(m2.group(5)),
            "pe": to_float(m2.group(6)),
        })
    data.setdefault("pas_radial", np.nan); data.setdefault("pad_radial", np.nan)
    data.setdefault("pam_radial", np.nan); data.setdefault("pp_radial", np.nan)
    data.setdefault("pas_central", np.nan); data.setdefault("pad_central", np.nan)
    data.setdefault("pam_central", np.nan); data.setdefault("pp_central", np.nan)
    data.setdefault("au", np.nan); data.setdefault("iau", np.nan); data.setdefault("rvse", np.nan); data.setdefault("pe", np.nan)
    return data

def brachial_bp_category(pas, pad):
    if np.isnan(pas) or np.isnan(pad): return "No clasificable"
    if pas >= 180 or pad >= 110: return "Etapa 3"
    if pas >= 160 or pad >= 100: return "Etapa 2"
    if pas >= 140 or pad >= 90: return "Etapa 1"
    if pas >= 130 or pad >= 85: return "Normal alta"
    if pas >= 120 or pad >= 80: return "Normal"
    return "Óptimo"

def central_diagnosis(row):
    cSBP = to_float(row.get("pas_central")); cPP = to_float(row.get("pp_central"));
    pSBP = to_float(row.get("pas_radial")); pDBP = to_float(row.get("pad_radial"));
    sex = (row.get("sexo") or "M").upper()[:1]
    cat = brachial_bp_category(pSBP, pDBP)
    ref = CENTRAL_SBP_TABLE.get(cat, {}).get(sex, np.nan)
    amp_sbp = pSBP - cSBP if not np.isnan(pSBP) and not np.isnan(cSBP) else np.nan
    ppa = (to_float(row.get("pp_radial")) / cPP) if cPP and not np.isnan(cPP) and cPP != 0 else np.nan
    if np.isnan(cSBP):
        dx = "No clasificable por falta de PAS central."
    elif cSBP >= 130:
        dx = "Hipertensión central probable / PAS aórtica central elevada."
    elif not np.isnan(ref) and cSBP > ref:
        dx = f"PAS central por encima del percentil 50 de referencia para la categoría braquial {cat}."
    else:
        dx = "Sin hipertensión central por el umbral operativo de la app."
    risk = []
    if not np.isnan(ppa) and ppa < 1.30: risk.append("amplificación de presión de pulso reducida (<1,30)")
    if not np.isnan(cPP) and cPP >= 50: risk.append("presión de pulso central aumentada")
    if not np.isnan(row.get("iau", np.nan)) and row.get("iau") >= 25: risk.append("índice de aumentación elevado")
    return dx, cat, ref, amp_sbp, ppa, "; ".join(risk) if risk else "sin señales hemodinámicas mayores agregadas"


def _norm01_from_value(value, low, high, default=0.5):
    """Normaliza un valor clínico a 0-1 para modular la morfología de la onda."""
    v = to_float(value)
    if np.isnan(v):
        return float(default)
    return float(np.clip((v - low) / max(high - low, 1e-6), 0, 1))


def make_waveform(row, n=512):
    """Curva sintética fisiológica individualizada y calibrada con PAS/PAD central.

    La versión previa generaba ondas demasiado parecidas porque usaba tiempos y amplitudes
    casi fijos. Esta versión modula forma, pico sistólico, hombro, cola diastólica y contenido
    armónico según PP central, Au, IAu, PE, FC y amplificación periférico-central.
    """
    cSBP = to_float(row.get("pas_central"))
    cDBP = to_float(row.get("pad_central"))
    pp_c = to_float(row.get("pp_central"))
    pSBP = to_float(row.get("pas_radial"))
    pp_r = to_float(row.get("pp_radial"))
    au = to_float(row.get("au"))
    iau = to_float(row.get("iau"))
    pe = to_float(row.get("pe"))
    fc = to_float(row.get("fc"))

    if np.isnan(cSBP) or cSBP <= 0:
        cSBP = 120.0
    if np.isnan(cDBP) or cDBP <= 0:
        cDBP = 80.0
    if cSBP <= cDBP:
        cSBP = cDBP + (pp_c if not np.isnan(pp_c) and pp_c > 10 else 35.0)

    pp = max(cSBP - cDBP, 12.0)
    t = np.linspace(0, 1, n)

    stiffness = _norm01_from_value(iau, 0, 45, 0.45)
    au_rel = 0.0 if np.isnan(au) else float(np.clip(au / max(pp, 1.0), -0.25, 0.55))
    pp_load = _norm01_from_value(pp, 25, 80, 0.45)
    fc_load = _norm01_from_value(fc, 55, 105, 0.45)
    pe_load = _norm01_from_value(pe, 25, 45, 0.50)
    amp_sbp = 0.0 if np.isnan(pSBP) else float(np.clip((pSBP - cSBP) / 35.0, -0.25, 1.0))
    ppa = np.nan
    if not np.isnan(pp_r) and pp > 0:
        ppa = pp_r / pp
    amp_load = 0.5 if np.isnan(ppa) else float(np.clip((1.70 - ppa) / 0.70, 0, 1))  # menor PPA = más carga central

    # Tiempos: rigidez/IAu alto => retorno más temprano y hombro más próximo al pico.
    peak_t = 0.18 + 0.055 * (1 - fc_load) + 0.025 * (1 - pe_load)
    peak_t = float(np.clip(peak_t, 0.16, 0.27))
    refl_t = 0.46 - 0.13 * stiffness - 0.04 * amp_load + 0.03 * (1 - pe_load)
    refl_t = float(np.clip(refl_t, peak_t + 0.08, 0.54))
    notch_t = float(np.clip(0.34 + 0.18 * pe_load, 0.34, 0.58))

    # Amplitudes: Au/IAu/PP/PPA modifican claramente el hombro reflejado.
    refl_amp = 0.12 + 0.35 * stiffness + 0.20 * max(au_rel, 0) + 0.12 * pp_load + 0.10 * amp_load
    refl_amp = float(np.clip(refl_amp, 0.08, 0.72))
    primary_width = 0.075 + 0.04 * (1 - stiffness) + 0.015 * pe_load
    refl_width = 0.085 + 0.07 * (1 - stiffness) + 0.025 * pe_load
    tail_amp = 0.10 + 0.16 * (1 - stiffness) + 0.05 * (1 - pp_load)

    primary = np.exp(-((t - peak_t) / primary_width) ** 2)
    reflected = refl_amp * np.exp(-((t - refl_t) / refl_width) ** 2)
    notch = -0.07 * (1 - stiffness) * np.exp(-((t - notch_t) / 0.040) ** 2)
    diastolic_tail = tail_amp * np.exp(-((t - 0.70) / (0.22 + 0.07*(1-stiffness))) ** 2)
    runoff = (0.04 + 0.07 * (1 - stiffness)) * (1 - t)

    # Componente armónico leve: evita ondas idénticas cuando los parámetros son cercanos.
    harmonic_shape = 0.018 * (pp_load - 0.5) * np.sin(2*np.pi*3*t) + 0.012 * (stiffness - 0.5) * np.sin(2*np.pi*5*t)

    raw = primary + reflected + notch + diastolic_tail + runoff + harmonic_shape
    raw = pd.Series(raw).rolling(5, center=True, min_periods=1).mean().to_numpy()
    p = cDBP + pp * (raw - raw.min()) / max(raw.max() - raw.min(), 1e-6)

    # Fijar exactamente PAD y PAS del estudio.
    p = cDBP + (p - p.min()) * (cSBP - cDBP) / max(p.max() - p.min(), 1e-6)
    return pd.DataFrame({"tiempo_ms": t * 1000, "presion_central_mmHg": p})

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

def _apply_professional_axes(ax, title=None, xlabel=None, ylabel=None):
    """Aplica formato profesional homogéneo a los gráficos clínicos."""
    ax.set_facecolor("#FFFFFF")
    ax.figure.patch.set_facecolor("#FFFFFF")
    if title:
        ax.set_title(title, fontsize=11, fontweight="bold", color="#12355B", pad=10)
    if xlabel:
        ax.set_xlabel(xlabel, fontsize=9, color="#263238")
    if ylabel:
        ax.set_ylabel(ylabel, fontsize=9, color="#263238")
    ax.tick_params(axis="both", labelsize=8, colors="#263238")
    ax.grid(True, alpha=0.22, linewidth=0.6)
    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)
    ax.spines["left"].set_color("#B0BEC5")
    ax.spines["bottom"].set_color("#B0BEC5")

def fig_to_png(fig):
    fig.tight_layout(pad=1.15)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf

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



def estimate_wave_separation(wave_df, row):
    """Separación estimada de onda anterógrada (Pf) y retrógrada (Pb) desde presión central.

    Versión reforzada:
    - Evita curvas iguales entre pacientes usando un vector morfológico individual: PP central,
      Au, IAu, PE, FC, PPA y amplificación PAS periférico-central.
    - Pf/Pb quedan superpuestas sobre la línea diastólica de la presión aórtica central.
    - La onda de flujo aórtico deja de ser triangular en punta y usa una forma beta/ejectiva
      redondeada con ascenso y descenso fisiológicos.
    """
    df = wave_df.copy()
    t = pd.to_numeric(df.iloc[:, 0], errors="coerce").to_numpy(dtype=float)
    p = pd.to_numeric(df.iloc[:, 1], errors="coerce").to_numpy(dtype=float)
    ok = np.isfinite(t) & np.isfinite(p)
    t, p = t[ok], p[ok]

    if len(p) < 20:
        df = make_waveform(row)
        t = df["tiempo_ms"].to_numpy(dtype=float)
        p = df["presion_central_mmHg"].to_numpy(dtype=float)

    t_norm = (t - np.nanmin(t)) / max(np.nanmax(t)-np.nanmin(t), 1e-6) * 1000.0
    t0 = np.linspace(0, 1000, 640)
    p0 = np.interp(t0, t_norm, p)
    p0 = pd.Series(p0).rolling(9, center=True, min_periods=1).mean().to_numpy()

    pad = to_float(row.get("pad_central"))
    pas = to_float(row.get("pas_central"))
    if np.isnan(pad): pad = float(np.nanmin(p0))
    if np.isnan(pas): pas = float(np.nanmax(p0))
    if pas <= pad:
        pas = float(np.nanmax(p0)); pad = float(np.nanmin(p0))
    pp = max(pas - pad, 1.0)

    p0 = pad + (p0 - np.nanmin(p0)) * (pas - pad) / max(np.nanmax(p0)-np.nanmin(p0), 1e-6)
    excess = np.clip(p0 - pad, 0, None)

    au = to_float(row.get("au"))
    iau = to_float(row.get("iau"))
    pe_pct = to_float(row.get("pe"))
    fc = to_float(row.get("fc"))
    pSBP = to_float(row.get("pas_radial"))
    pp_r = to_float(row.get("pp_radial"))

    stiffness = _norm01_from_value(iau, -10, 45, 0.42)
    au_rel = 0.0 if np.isnan(au) else float(np.clip(au / max(pp, 1.0), -0.35, 0.65))
    pp_load = _norm01_from_value(pp, 22, 82, 0.45)
    fc_load = _norm01_from_value(fc, 50, 110, 0.45)
    cycle_ms = 60000.0 / fc if not np.isnan(fc) and fc > 20 else 1000.0
    ppa = np.nan if np.isnan(pp_r) or pp <= 0 else pp_r / pp
    amp_load = 0.50 if np.isnan(ppa) else float(np.clip((1.70 - ppa) / 0.75, 0, 1))
    amp_sbp = 0.0 if np.isnan(pSBP) else float(np.clip((pSBP - pas) / 40.0, -0.35, 1.15))

    # Evitar valores de PE demasiado fijos: la duración eyectiva cambia con PE/FC y carga.
    if not np.isnan(pe_pct) and pe_pct > 10:
        ej_ms = float(np.clip(cycle_ms * pe_pct / 100.0, 210, 460))
    else:
        ej_ms = float(np.clip(300 + 70*(1-fc_load) + 35*(1-stiffness), 240, 430))

    peak_i = int(np.nanargmax(p0))
    t_peak = float(t0[peak_i])

    # Retorno reflejado individualizado: IAu/Au altos + PPA baja => retorno más precoz y más prominente.
    tref = t_peak + 180 - 2.6*(0 if np.isnan(iau) else iau) - 42*amp_load + 24*(1-pp_load) - 20*max(au_rel, 0)
    tref = float(np.clip(tref, t_peak + 48, min(560, ej_ms + 180)))

    # Pico retrógrado con fuerte dependencia de Au/IAu/PPA/PP. Permite diferencias visuales claras.
    if not np.isnan(au):
        pb_peak = 0.08*pp + max(au, -0.12*pp) + 0.24*pp*stiffness + 0.13*pp*amp_load + 0.04*pp_load*pp
    elif not np.isnan(iau):
        pb_peak = pp * (0.09 + iau/90.0 + 0.12*amp_load + 0.05*pp_load)
    else:
        pb_peak = pp * (0.18 + 0.16*amp_load + 0.10*stiffness)
    pb_peak = float(np.clip(pb_peak, 0.05*pp, 0.72*pp))

    # Pb: no una gaussiana única fija; combina hombro sistólico y cola tardía con anchura variable.
    pb_width = float(np.clip(82 + 92*(1-stiffness) + 28*(1-pp_load), 72, 190))
    pb = pb_peak * np.exp(-((t0 - tref) / pb_width) ** 2)
    tail_center = 610 + 90*(1-stiffness) - 40*fc_load
    pb += (0.06 + 0.13*(1-stiffness)) * pb_peak * np.exp(-((t0 - tail_center) / (220 + 70*(1-stiffness))) ** 2)

    # Pf: componente anterógrado ancho/estrecho según FC, PE y rigidez; conserva la silueta de la curva madre.
    pf = np.clip(excess - pb, 0, None)
    pf_peak_target = pp * float(np.clip(0.84 - 0.30*stiffness + 0.08*amp_sbp + 0.07*(1-amp_load), 0.48, 0.95))
    pf_width = float(np.clip(68 + 58*(1-stiffness) + 0.10*ej_ms - 22*fc_load, 62, 160))
    pf_template = pf_peak_target * np.exp(-((t0 - t_peak) / pf_width) ** 2)
    pf_template += 0.05 * pf_peak_target * np.exp(-((t0 - (t_peak + 135 + 40*(1-stiffness))) / 175.0) ** 2)
    # Mezcla adaptativa: si la resta queda pobre, predomina la plantilla individual.
    if np.nanmax(pf) < 0.32*pp:
        pf = pf_template
    else:
        pf = 0.72*pf + 0.28*pf_template

    # Escala sin borrar la relación Pf/Pb. La suma puede no coincidir perfecto, pero conserva PAS/PAD y diferencias.
    summed = pf + pb
    if np.nanmax(summed) > 0:
        scale = pp / max(np.nanmax(summed), 1e-6)
        pf *= scale
        pb *= scale

    pf_abs = pad + pf
    pb_abs = pad + pb
    p_model = pad + np.clip(pf + pb, 0, None)
    p_model = pad + (p_model - np.nanmin(p_model)) * (pas - pad) / max(np.nanmax(p_model)-np.nanmin(p_model), 1e-6)

    pf_peak = float(np.nanmax(pf))
    pb_peak = float(np.nanmax(pb))
    tfor = float(t0[int(np.nanargmax(pf))])
    tref_m = float(t0[int(np.nanargmax(pb))])
    rm = pb_peak / pf_peak if pf_peak > 0 else np.nan
    ri = pb_peak / (pf_peak + pb_peak) if (pf_peak + pb_peak) > 0 else np.nan
    t_ratio = tfor / tref_m if tref_m > 0 else np.nan

    # Flujo aórtico redondeado tipo eyección, no triangular. Usa forma beta con parámetros variables.
    flow = np.zeros_like(t0)
    ej_start = max(0.0, t_peak - (100 + 18*(1-stiffness)))
    ej_end = float(np.clip(ej_ms, max(t_peak+120, 250), 540))
    eject = (t0 >= ej_start) & (t0 <= ej_end)
    u = np.zeros_like(t0)
    u[eject] = (t0[eject] - ej_start) / max(ej_end - ej_start, 1e-6)
    # alpha/beta modulan punta: más rigidez/FC = pico algo más temprano, pero siempre romo.
    alpha = 2.05 + 0.55*(1-stiffness) + 0.15*(1-fc_load)
    beta = 3.15 + 0.65*stiffness + 0.25*fc_load
    flow_shape = np.zeros_like(t0)
    flow_shape[eject] = (u[eject] ** (alpha-1)) * ((1-u[eject]) ** (beta-1))
    # Meseta sistólica suave para evitar aspecto en punta.
    flow_shape[eject] = pd.Series(flow_shape[eject]).rolling(23, center=True, min_periods=1).mean().to_numpy()
    if np.nanmax(flow_shape) > 0:
        flow_shape = flow_shape / np.nanmax(flow_shape)
    qp = np.clip(210 + 3.4*pp + (0 if np.isnan(fc) else 0.75*fc) - 36*stiffness + 18*(1-amp_load), 170, 560)
    flow = flow_shape * qp
    flow = pd.Series(flow).rolling(17, center=True, min_periods=1).mean().to_numpy()

    sep_df = pd.DataFrame({
        "tiempo_ms": t0,
        "presion_total_mmHg": p0,
        "presion_modelo_pf_pb_mmHg": p_model,
        "onda_anterograda_pf": pf,
        "onda_retrograda_pb": pb,
        "onda_anterograda_pf_abs": pf_abs,
        "onda_retrograda_pb_abs": pb_abs,
        "flujo_aortico_estimado_ml_s": flow,
    })

    metrics = {
        "pf_pico": pf_peak,
        "pb_pico": pb_peak,
        "tfor_ms": tfor,
        "tref_ms": tref_m,
        "rm": rm,
        "ri": ri,
        "tfor_tref": t_ratio,
        "qp_ml_s": float(np.nanmax(flow)),
        "pe_ms": float(ej_end - ej_start),
        "ppa": ppa,
        "rigidez_modelo": stiffness,
        "t_ej_inicio_ms": float(ej_start),
        "t_ej_fin_ms": float(ej_end),
    }
    return sep_df, metrics


def interpret_wave_separation(sep_metrics):
    rm = sep_metrics.get("rm", np.nan)
    ri = sep_metrics.get("ri", np.nan)
    tref = sep_metrics.get("tref_ms", np.nan)
    ratio = sep_metrics.get("tfor_tref", np.nan)
    pb = sep_metrics.get("pb_pico", np.nan)

    parts = []
    if not np.isnan(rm):
        if rm < 0.30:
            parts.append("La magnitud de reflexión (RM) es baja, compatible con onda reflejada poco dominante y menor carga pulsátil central.")
        elif rm < 0.45:
            parts.append("La magnitud de reflexión (RM) es intermedia, compatible con reflexión arterial presente sin predominio marcado.")
        else:
            parts.append("La magnitud de reflexión (RM) es elevada, compatible con mayor contribución de la onda retrógrada a la presión sistólica central y aumento de poscarga pulsátil.")
    if not np.isnan(ri):
        if ri < 0.23:
            parts.append("El índice de reflexión (RI) se encuentra en rango bajo.")
        elif ri < 0.32:
            parts.append("El índice de reflexión (RI) se encuentra en rango intermedio.")
        else:
            parts.append("El índice de reflexión (RI) se encuentra aumentado, sugiriendo mayor carga por onda reflejada.")
    if not np.isnan(tref):
        if tref < 320:
            parts.append("El retorno de la onda retrógrada es precoz, patrón compatible con rigidez arterial aumentada o reflexión periférica temprana.")
        elif tref <= 430:
            parts.append("El tiempo de retorno de la onda reflejada es intermedio.")
        else:
            parts.append("El tiempo de retorno de la onda reflejada es tardío, patrón más compatible con mejor complacencia arterial central.")
    if not np.isnan(ratio):
        if ratio > 0.55:
            parts.append("La relación Tfor/Tref se encuentra relativamente elevada, lo que puede expresar solapamiento temporal entre componente anterógrado y retrógrado.")
        else:
            parts.append("La relación Tfor/Tref no sugiere solapamiento reflejo precoz significativo.")
    if not parts:
        return "No fue posible estimar en forma estable la separación de ondas."
    return " ".join(parts)
def plot_waveform(wave_df):
    fig, ax = plt.subplots(figsize=(7.6, 4.2))
    x = pd.to_numeric(wave_df.iloc[:,0], errors="coerce")
    y = pd.to_numeric(wave_df.iloc[:,1], errors="coerce")
    ax.plot(x, y, color="#B71C1C", linewidth=2.6)
    ax.fill_between(x, y.min(), y, color="#B71C1C", alpha=0.06)
    _apply_professional_axes(ax, "Onda de presión aórtica central", "Tiempo (ms)", "Presión central (mmHg)")
    ax.margins(x=0.01)
    return fig_to_png(fig)


def plot_wave_separation(sep_df):
    """Gráfico clínico integrado: presión aórtica central + Pf/Pb superpuestas."""
    t = pd.to_numeric(sep_df["tiempo_ms"], errors="coerce").to_numpy(dtype=float)
    p_total = pd.to_numeric(sep_df["presion_total_mmHg"], errors="coerce").to_numpy(dtype=float)
    pf_abs = pd.to_numeric(sep_df.get("onda_anterograda_pf_abs", sep_df["onda_anterograda_pf"]), errors="coerce").to_numpy(dtype=float)
    pb_abs = pd.to_numeric(sep_df.get("onda_retrograda_pb_abs", sep_df["onda_retrograda_pb"]), errors="coerce").to_numpy(dtype=float)
    ok = np.isfinite(t) & np.isfinite(p_total) & np.isfinite(pf_abs) & np.isfinite(pb_abs)
    t, p_total, pf_abs, pb_abs = t[ok], p_total[ok], pf_abs[ok], pb_abs[ok]
    pad_base = float(np.nanmin(p_total)) if len(p_total) else 0.0

    fig, ax = plt.subplots(figsize=(8.8, 4.9))
    ax.plot(t, p_total, color="#111111", linewidth=3.0, label="Presión aórtica central completa", zorder=5)
    ax.plot(t, pf_abs, color="#168038", linewidth=2.35, label="Onda anterógrada Pf", zorder=4)
    ax.plot(t, pb_abs, color="#EF6C00", linestyle="--", linewidth=2.35, label="Onda retrógrada Pb", zorder=4)
    ax.fill_between(t, pad_base, pf_abs, alpha=0.07, color="#168038", zorder=2)
    ax.fill_between(t, pad_base, pb_abs, alpha=0.08, color="#EF6C00", zorder=2)
    ax.axhline(pad_base, color="#78909C", linewidth=0.9, alpha=0.8)
    _apply_professional_axes(ax, "Separación de ondas superpuesta a la presión aórtica central", "Tiempo (ms)", "Presión / componentes sobre PAD (mmHg)")
    ax.legend(fontsize=8, loc="upper right", frameon=True, facecolor="white", edgecolor="#CFD8DC")
    ax.margins(x=0.01)
    return fig_to_png(fig)


def plot_aortic_flow(sep_df):
    fig, ax = plt.subplots(figsize=(7.6, 4.2))
    t = sep_df["tiempo_ms"]
    q = sep_df["flujo_aortico_estimado_ml_s"]
    ax.plot(t, q, color="#6A1B9A", linewidth=2.6)
    ax.fill_between(t, 0, q, color="#6A1B9A", alpha=0.08)
    _apply_professional_axes(ax, "Curva estimada de flujo aórtico", "Tiempo (ms)", "Flujo aórtico estimado (mL/s)")
    ax.set_ylim(bottom=0)
    ax.margins(x=0.01)
    return fig_to_png(fig)


def plot_harmonics(hdf):
    fig, ax = plt.subplots(figsize=(7.6, 4.2))
    ax.bar(range(1, len(hdf)+1), hdf["energia_relativa_%"], color="#455A64")
    ax.set_xticks(range(1, len(hdf)+1))
    _apply_professional_axes(ax, "Análisis armónico de la onda de presión central", "Armónico", "Energía relativa (%)")
    ax.grid(axis="y", alpha=.22)
    return fig_to_png(fig)


def plot_pressure_comparison(row):
    labels = ["PAS", "PAD", "PAM", "PP"]
    radial = [row.get("pas_radial"), row.get("pad_radial"), row.get("pam_radial"), row.get("pp_radial")]
    central = [row.get("pas_central"), row.get("pad_central"), row.get("pam_central"), row.get("pp_central")]
    x = np.arange(len(labels)); w = 0.34
    fig, ax = plt.subplots(figsize=(7.6,4.2))
    ax.bar(x-w/2, radial, w, label="Radial/Braquial", color="#607D8B")
    ax.bar(x+w/2, central, w, label="Central", color="#1565C0")
    ax.set_xticks(x); ax.set_xticklabels(labels)
    _apply_professional_axes(ax, "Presiones periféricas vs centrales", "Variable", "mmHg")
    ax.legend(fontsize=8, frameon=True, facecolor="white", edgecolor="#CFD8DC")
    ax.grid(axis="y", alpha=.22)
    return fig_to_png(fig)


def plot_clinical_gauges(row, ppa):
    metrics = {
        "PAS central": (to_float(row.get("pas_central")), 130),
        "PP central": (to_float(row.get("pp_central")), 50),
        "IAu": (to_float(row.get("iau")), 25),
        "PPA": (ppa, 1.30),
    }
    fig, ax = plt.subplots(figsize=(7.6,3.9))
    names = list(metrics.keys())
    vals = [metrics[k][0] for k in names]
    refs = [metrics[k][1] for k in names]
    score = []
    for n,v,r in zip(names, vals, refs):
        if np.isnan(v): score.append(np.nan)
        elif n == "PPA": score.append(v/r)
        else: score.append(v/r)
    ax.barh(names, score, color="#546E7A")
    ax.axvline(1, linestyle="--", linewidth=1.2, color="#B71C1C")
    _apply_professional_axes(ax, "Semaforización clínica de parámetros centrales", "Relación valor / umbral", "")
    ax.grid(axis="x", alpha=.22)
    return fig_to_png(fig)



def interpret_pressure_central_metrics(row, dx, cat, ref, amp_sbp, ppa, risk):
    """Conclusión textual continua del bloque de presión central y métricas."""
    pas_c = to_float(row.get("pas_central")); pad_c = to_float(row.get("pad_central")); pp_c = to_float(row.get("pp_central"))
    pam_c = to_float(row.get("pam_central")); iau = to_float(row.get("iau")); au = to_float(row.get("au")); fc = to_float(row.get("fc"))

    def fmt(v, dec=1):
        try:
            f = float(v)
            if np.isnan(f): return "no disponible"
            return f"{f:.{dec}f}"
        except Exception:
            return "no disponible"

    pressure_flags = []
    if not np.isnan(pas_c):
        if pas_c >= 130:
            pressure_flags.append("PAS central elevada por umbral operativo de 130 mmHg")
        else:
            pressure_flags.append("PAS central por debajo del umbral operativo de 130 mmHg")
    if not np.isnan(pp_c):
        if pp_c >= 50:
            pressure_flags.append("presión de pulso central aumentada")
        else:
            pressure_flags.append("presión de pulso central no aumentada")
    if not np.isnan(iau):
        if iau >= 25:
            pressure_flags.append("IAu elevado, compatible con mayor aumentación sistólica")
        elif iau >= 10:
            pressure_flags.append("IAu intermedio")
        else:
            pressure_flags.append("IAu bajo")
    if not np.isnan(ppa):
        if ppa < 1.30:
            pressure_flags.append("amplificación de presión de pulso reducida")
        else:
            pressure_flags.append("amplificación de presión de pulso conservada")

    return (
        f"El análisis de presión central muestra PAS central {fmt(pas_c,0)} mmHg, PAD central {fmt(pad_c,0)} mmHg, "
        f"PAM central {fmt(pam_c,0)} mmHg y PP central {fmt(pp_c,0)} mmHg. La categoría tensional periférica/braquial es {cat}. "
        f"La amplificación PAS periférico-central es {fmt(amp_sbp,1)} mmHg y la PPA es {fmt(ppa,2)}. "
        f"Au: {fmt(au,1)} mmHg, IAu: {fmt(iau,1)}%, FC: {fmt(fc,0)} lpm. "
        f"Conclusión operativa: {dx}. Perfil hemodinámico agregado: {risk}. "
        f"Síntesis: {'; '.join(pressure_flags) if pressure_flags else 'sin marcadores suficientes para estratificación central completa'}."
    )


def interpret_harmonic_profile(hdf):
    """Conclusión clínica simple del análisis armónico."""
    try:
        df = hdf.copy()
        if df is None or len(df) == 0:
            return "No fue posible calcular un perfil armónico interpretable por falta de datos válidos de la onda central."
        e = pd.to_numeric(df.get("energia_relativa_%"), errors="coerce").to_numpy(dtype=float)
        f = pd.to_numeric(df.get("frecuencia_hz"), errors="coerce").to_numpy(dtype=float)
        a = pd.to_numeric(df.get("amplitud"), errors="coerce").to_numpy(dtype=float)
        if len(e) == 0 or np.all(~np.isfinite(e)):
            return "No fue posible estimar la distribución espectral de energía de la onda central."
        e1 = e[0] if len(e) > 0 and np.isfinite(e[0]) else np.nan
        e2 = e[1] if len(e) > 1 and np.isfinite(e[1]) else np.nan
        e_high = np.nansum(e[3:]) if len(e) > 4 else np.nan
        dom_i = int(np.nanargmax(e)) if np.any(np.isfinite(e)) else 0
        dom_freq = f[dom_i] if len(f) > dom_i and np.isfinite(f[dom_i]) else np.nan
        dom_amp = a[dom_i] if len(a) > dom_i and np.isfinite(a[dom_i]) else np.nan

        if not np.isnan(e_high) and e_high >= 25:
            spectral_note = "mayor contenido de alta frecuencia, hallazgo compatible con morfología más abrupta o mayor complejidad de la onda de presión"
        elif not np.isnan(e_high) and e_high >= 12:
            spectral_note = "contenido de alta frecuencia intermedio"
        else:
            spectral_note = "predominio de componentes armónicos bajos, compatible con onda más suavizada"

        return (
            f"El análisis armónico muestra una energía relativa del primer armónico de {e1:.1f}% "
            f"y del segundo armónico de {e2:.1f}% cuando está disponible. El armónico dominante se ubica en torno a "
            f"{dom_freq:.2f} Hz, con amplitud {dom_amp:.3f}. La suma aproximada de armónicos altos desde el cuarto componente es "
            f"{e_high:.1f}%. Interpretación: {spectral_note}. Este bloque debe interpretarse como complemento morfológico de la onda central, "
            f"no como sustituto de la interpretación clínica de presión central, aumentación y separación de ondas."
        )
    except Exception:
        return "El perfil armónico no pudo interpretarse en forma estable, aunque se conserva el gráfico y la tabla espectral para revisión visual."


def build_continuous_conclusions(row, wave_df, hdf):
    """Devuelve las cuatro conclusiones continuas solicitadas antes de los gráficos."""
    dx, cat, ref, amp_sbp, ppa, risk = central_diagnosis(row)
    sep_df, sep_metrics = estimate_wave_separation(wave_df, row)
    sep_interp = interpret_wave_separation(sep_metrics)

    def fmt(v, dec=1):
        try:
            f = float(v)
            if np.isnan(f): return "no disponible"
            return f"{f:.{dec}f}"
        except Exception:
            return "no disponible"

    c1 = interpret_pressure_central_metrics(row, dx, cat, ref, amp_sbp, ppa, risk)

    c2 = (
        f"La separación de ondas estima un componente anterógrado Pf pico de {fmt(sep_metrics.get('pf_pico', np.nan),1)} mmHg "
        f"y un componente retrógrado Pb pico de {fmt(sep_metrics.get('pb_pico', np.nan),1)} mmHg. "
        f"La magnitud de reflexión RM Pb/Pf es {fmt(sep_metrics.get('rm', np.nan),2)}, el índice de reflexión RI es "
        f"{fmt(sep_metrics.get('ri', np.nan),2)}, Tfor es {fmt(sep_metrics.get('tfor_ms', np.nan),0)} ms, "
        f"Tref es {fmt(sep_metrics.get('tref_ms', np.nan),0)} ms y la relación Tfor/Tref es "
        f"{fmt(sep_metrics.get('tfor_tref', np.nan),2)}. {sep_interp}"
    )

    c3 = interpret_harmonic_profile(hdf)

    rm = sep_metrics.get("rm", np.nan); ri = sep_metrics.get("ri", np.nan)
    pp_c = to_float(row.get("pp_central")); iau = to_float(row.get("iau")); pas_c = to_float(row.get("pas_central"))
    integrated_flags = []
    if not np.isnan(pas_c) and pas_c >= 130:
        integrated_flags.append("presión central elevada")
    if not np.isnan(pp_c) and pp_c >= 50:
        integrated_flags.append("carga pulsátil central aumentada")
    if not np.isnan(iau) and iau >= 25:
        integrated_flags.append("aumentación sistólica elevada")
    if not np.isnan(rm) and rm >= 0.45:
        integrated_flags.append("reflexión de onda aumentada")
    if not np.isnan(ri) and ri >= 0.32:
        integrated_flags.append("mayor contribución retrógrada")
    if not integrated_flags:
        integrated_flags.append("sin marcadores mayores simultáneos de sobrecarga pulsátil central en los parámetros disponibles")

    c4 = (
        "La interpretación final integrada combina presión central, métricas de aumentación, separación Pf/Pb y análisis armónico. "
        f"En este registro predominan los siguientes elementos: {', '.join(integrated_flags)}. "
        "El resultado debe integrarse con edad, sexo, presión braquial, tratamiento, lesión de órgano blanco y contexto clínico. "
        "La lectura más relevante es la concordancia entre presión central, carga pulsátil y temporalidad de la onda reflejada, porque esa combinación orienta el fenotipo vascular y la carga hemodinámica central."
    )

    return [
        ("1. Análisis de presión central y métricas", c1),
        ("2. Análisis de separación de ondas", c2),
        ("3. Análisis de armónicos", c3),
        ("4. Interpretación final integrada", c4),
    ], sep_df, sep_metrics, sep_interp



def classify_central_pressure_phenotype(row, sep_metrics, hdf):
    """Define fenotipo final integrando presión central, separación de ondas y armónicos."""
    pas_c = to_float(row.get("pas_central"))
    pad_c = to_float(row.get("pad_central"))
    pp_c = to_float(row.get("pp_central"))
    au = to_float(row.get("au"))
    iau = to_float(row.get("iau"))
    pas_r = to_float(row.get("pas_radial"))
    pp_r = to_float(row.get("pp_radial"))
    ppa = pp_r / pp_c if not np.isnan(pp_r) and not np.isnan(pp_c) and pp_c > 0 else np.nan
    amp_sbp = pas_r - pas_c if not np.isnan(pas_r) and not np.isnan(pas_c) else np.nan
    rm = sep_metrics.get("rm", np.nan)
    ri = sep_metrics.get("ri", np.nan)
    tref = sep_metrics.get("tref_ms", np.nan)
    pb = sep_metrics.get("pb_pico", np.nan)
    pf = sep_metrics.get("pf_pico", np.nan)

    e_high = np.nan
    e1 = np.nan
    try:
        e = pd.to_numeric(hdf.get("energia_relativa_%"), errors="coerce").to_numpy(dtype=float)
        if len(e) > 0:
            e1 = e[0]
        if len(e) > 4:
            e_high = np.nansum(e[3:])
    except Exception:
        pass

    pressure_score = 0
    wave_score = 0
    harmonic_score = 0
    reasons_pressure = []
    reasons_wave = []
    reasons_harm = []

    if not np.isnan(pas_c):
        if pas_c >= 130:
            pressure_score += 2; reasons_pressure.append("PAS central elevada")
        elif pas_c >= 120:
            pressure_score += 1; reasons_pressure.append("PAS central limítrofe")
        else:
            reasons_pressure.append("PAS central no elevada")
    if not np.isnan(pp_c):
        if pp_c >= 60:
            pressure_score += 2; reasons_pressure.append("PP central marcadamente aumentada")
        elif pp_c >= 50:
            pressure_score += 1; reasons_pressure.append("PP central aumentada")
        else:
            reasons_pressure.append("PP central no aumentada")
    if not np.isnan(iau):
        if iau >= 35:
            pressure_score += 2; reasons_pressure.append("IAu alto")
        elif iau >= 25:
            pressure_score += 1; reasons_pressure.append("IAu aumentado")
        else:
            reasons_pressure.append("IAu no aumentado")
    if not np.isnan(ppa):
        if ppa < 1.20:
            pressure_score += 2; reasons_pressure.append("PPA francamente reducida")
        elif ppa < 1.30:
            pressure_score += 1; reasons_pressure.append("PPA reducida")
        else:
            reasons_pressure.append("PPA conservada")

    if not np.isnan(rm):
        if rm >= 0.50:
            wave_score += 2; reasons_wave.append("RM elevada")
        elif rm >= 0.35:
            wave_score += 1; reasons_wave.append("RM intermedia")
        else:
            reasons_wave.append("RM baja")
    if not np.isnan(ri):
        if ri >= 0.35:
            wave_score += 2; reasons_wave.append("RI aumentado")
        elif ri >= 0.25:
            wave_score += 1; reasons_wave.append("RI intermedio")
        else:
            reasons_wave.append("RI bajo")
    if not np.isnan(tref):
        if tref < 320:
            wave_score += 2; reasons_wave.append("retorno reflejo precoz")
        elif tref <= 430:
            wave_score += 1; reasons_wave.append("retorno reflejo intermedio")
        else:
            reasons_wave.append("retorno reflejo tardío")

    if not np.isnan(e_high):
        if e_high >= 25:
            harmonic_score += 2; reasons_harm.append("alto contenido de armónicos superiores")
        elif e_high >= 12:
            harmonic_score += 1; reasons_harm.append("contenido armónico superior intermedio")
        else:
            reasons_harm.append("predominio de armónicos bajos")
    if not np.isnan(e1):
        if e1 < 35:
            harmonic_score += 1; reasons_harm.append("menor predominio del primer armónico")
        else:
            reasons_harm.append("primer armónico predominante")

    total_score = pressure_score + wave_score + harmonic_score
    high_pressure = pressure_score >= 3
    high_wave = wave_score >= 3
    high_harm = harmonic_score >= 2

    if high_pressure and high_wave and high_harm:
        phenotype = "Fenotipo central rígido-reflectivo con complejidad armónica aumentada"
        clinical = "predomina una carga pulsátil central aumentada, con retorno retrógrado relevante/precoz y mayor complejidad espectral de la onda de presión."
    elif high_pressure and high_wave:
        phenotype = "Fenotipo central rígido-reflectivo"
        clinical = "predomina presión/carga pulsátil central elevada asociada a mayor contribución de onda retrógrada."
    elif high_pressure and high_harm:
        phenotype = "Fenotipo central de carga pulsátil elevada con distorsión armónica"
        clinical = "predomina presión o presión de pulso central elevada con morfología espectral más compleja."
    elif high_wave:
        phenotype = "Fenotipo reflectivo predominante"
        clinical = "la alteración principal se concentra en la separación de ondas, con mayor peso de Pb y/o retorno reflejo más precoz."
    elif high_harm:
        phenotype = "Fenotipo armónico complejo sin sobrecarga central mayor manifiesta"
        clinical = "el hallazgo dominante es morfológico-espectral y debe correlacionarse con la calidad de señal y el contexto vascular."
    elif high_pressure:
        phenotype = "Fenotipo de presión central elevada no reflectivo predominante"
        clinical = "predomina la elevación de presión/carga pulsátil central, sin evidencia fuerte de predominio retrógrado en los parámetros disponibles."
    else:
        phenotype = "Fenotipo central conservado o de bajo impacto pulsátil"
        clinical = "no se identifican marcadores simultáneos mayores de sobrecarga central, reflexión aumentada o complejidad armónica relevante en los datos disponibles."

    table = [
        ["Dominio", "Puntaje", "Elementos considerados"],
        ["Presión central y métricas", str(pressure_score), "; ".join(reasons_pressure) if reasons_pressure else "sin datos suficientes"],
        ["Separación de ondas Pf/Pb", str(wave_score), "; ".join(reasons_wave) if reasons_wave else "sin datos suficientes"],
        ["Armónicos", str(harmonic_score), "; ".join(reasons_harm) if reasons_harm else "sin datos suficientes"],
        ["Puntaje integrado", str(total_score), phenotype],
    ]

    text = (
        f"Fenotipo final: {phenotype}. La integración de presión central, separación de ondas y análisis armónico indica que {clinical} "
        f"Valores integrados: PAS central {pas_c:.0f} mmHg si disponible, PAD central {pad_c:.0f} mmHg si disponible, "
        f"PP central {pp_c:.0f} mmHg si disponible, IAu {iau:.1f}% si disponible, PPA {ppa:.2f} si disponible, "
        f"RM {rm:.2f}, RI {ri:.2f}, Tref {tref:.0f} ms, Pf pico {pf:.1f} mmHg y Pb pico {pb:.1f} mmHg. "
        "Este fenotipo es una clasificación operativa de apoyo y debe integrarse con edad, sexo, presión braquial, tratamiento, riesgo cardiovascular y lesión de órgano blanco."
    )
    text = text.replace("nan mmHg", "no disponible").replace("nan%", "no disponible").replace("nan", "no disponible")
    return phenotype, text, table

def build_pdf(row, wave_df, hdf, screenshot_png=None):
    dx, cat, ref, amp_sbp, ppa, risk = central_diagnosis(row)
    conclusion_blocks, sep_df, sep_metrics, sep_interp = build_continuous_conclusions(row, wave_df, hdf)
    final_phenotype, final_phenotype_text, final_phenotype_table = classify_central_pressure_phenotype(row, sep_metrics, hdf)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        rightMargin=13*mm, leftMargin=13*mm,
        topMargin=23*mm, bottomMargin=15*mm
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="SmallPAC", parent=styles["BodyText"], fontName="Helvetica",
        fontSize=7.5, leading=9.2, textColor=colors.HexColor("#263238")
    ))
    styles.add(ParagraphStyle(
        name="BodyPAC", parent=styles["BodyText"], fontName="Helvetica",
        fontSize=8.7, leading=11.2, textColor=colors.HexColor("#1F2D3D"), spaceAfter=3
    ))
    styles.add(ParagraphStyle(
        name="ConclusionPAC", parent=styles["BodyText"], fontName="Helvetica",
        fontSize=8.4, leading=10.7, textColor=colors.HexColor("#1F2D3D"), spaceAfter=5
    ))
    styles.add(ParagraphStyle(
        name="TitlePAC", parent=styles["Title"], fontName="Helvetica-Bold",
        fontSize=15.5, leading=19, alignment=1, textColor=colors.HexColor("#12355B"), spaceAfter=5
    ))
    styles.add(ParagraphStyle(
        name="SectionPAC", parent=styles["Heading2"], fontName="Helvetica-Bold",
        fontSize=10.5, leading=13, textColor=colors.white, leftIndent=0, spaceAfter=0
    ))
    styles.add(ParagraphStyle(
        name="H3PAC", parent=styles["Heading3"], fontName="Helvetica-Bold",
        fontSize=9, leading=11, textColor=colors.HexColor("#17365D"), spaceAfter=2
    ))
    styles.add(ParagraphStyle(
        name="MiniTitlePAC", parent=styles["Heading3"], fontName="Helvetica-Bold",
        fontSize=8.8, leading=10.5, textColor=colors.HexColor("#12355B"), spaceBefore=2, spaceAfter=1
    ))

    def _fmt(v, dec=1):
        try:
            f = float(v)
            if np.isnan(f):
                return ""
            return f"{f:.{dec}f}"
        except Exception:
            return safe_text(v)

    def _section(title):
        return Table([[Paragraph(title, styles["SectionPAC"])]], colWidths=[184*mm], style=TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#17365D")),
            ("BOX", (0,0), (-1,-1), 0.4, colors.HexColor("#17365D")),
            ("LEFTPADDING", (0,0), (-1,-1), 6),
            ("RIGHTPADDING", (0,0), (-1,-1), 6),
            ("TOPPADDING", (0,0), (-1,-1), 4),
            ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ]))

    def _table_style(header_color="#D9EAF7"):
        return TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor(header_color)),
            ("TEXTCOLOR", (0,0), (-1,0), colors.HexColor("#17365D")),
            ("FONT", (0,0), (-1,0), "Helvetica-Bold", 8),
            ("FONT", (0,1), (-1,-1), "Helvetica", 7.6),
            ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#B0BEC5")),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F7FAFC")]),
            ("LEFTPADDING", (0,0), (-1,-1), 4),
            ("RIGHTPADDING", (0,0), (-1,-1), 4),
            ("TOPPADDING", (0,0), (-1,-1), 3),
            ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ])

    def _header_footer(canvas, doc_obj):
        canvas.saveState()
        width, height = A4
        canvas.setFillColor(colors.HexColor("#12355B"))
        canvas.rect(0, height-17*mm, width, 17*mm, stroke=0, fill=1)
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica-Bold", 9)
        canvas.drawString(13*mm, height-10*mm, "PAC IA | Presión Aórtica Central")
        canvas.setFont("Helvetica", 7.5)
        canvas.drawRightString(width-13*mm, height-10*mm, datetime.now().strftime("%d/%m/%Y"))
        canvas.setFillColor(colors.HexColor("#607D8B"))
        canvas.setFont("Helvetica", 7)
        canvas.drawString(13*mm, 8*mm, "Informe médico integrado - conclusiones continuas antes de gráficos")
        canvas.drawRightString(width-13*mm, 8*mm, f"Página {doc_obj.page}")
        canvas.restoreState()

    story = []
    story.append(Paragraph("PRESIÓN AÓRTICA CENTRAL", styles["TitlePAC"]))
    story.append(Paragraph("Informe médico integrado con conclusiones clínicas continuas y panel gráfico posterior", styles["BodyPAC"]))
    story.append(Spacer(1, 3*mm))

    story.append(_section("1. Datos del paciente y valores principales"))
    datos = [
        ["Paciente", safe_text(row.get("paciente","")), "Estudio", safe_text(row.get("estudio",""))],
        ["Fecha", safe_text(row.get("fecha","")), "Hora", safe_text(row.get("hora",""))],
        ["Edad", _fmt(row.get("edad",""),0), "Sexo", safe_text(row.get("sexo",""))],
        ["Peso", _fmt(row.get("peso",""),1), "Altura", _fmt(row.get("altura",""),1)],
        ["IMC", _fmt(row.get("imc",""),1), "Medicación", safe_text(row.get("medicacion",""))],
    ]
    story.append(Table(datos, colWidths=[28*mm, 59*mm, 28*mm, 69*mm], style=_table_style("#EAF2F8")))
    story.append(Spacer(1, 2*mm))

    vals = [["Variable", "Radial/Braquial", "Central", "Unidad"],
            ["PAS", _fmt(row.get("pas_radial")), _fmt(row.get("pas_central")), "mmHg"],
            ["PAD", _fmt(row.get("pad_radial")), _fmt(row.get("pad_central")), "mmHg"],
            ["PAM", _fmt(row.get("pam_radial")), _fmt(row.get("pam_central")), "mmHg"],
            ["PP", _fmt(row.get("pp_radial")), _fmt(row.get("pp_central")), "mmHg"],
            ["FC", _fmt(row.get("fc"),0), "", "lpm"],
            ["Au", "", _fmt(row.get("au")), "mmHg"],
            ["IAu", "", _fmt(row.get("iau")), "%"],
            ["RVSE", "", _fmt(row.get("rvse")), "%"],
            ["PE", "", _fmt(row.get("pe")), "%"]]
    story.append(Table(vals, colWidths=[42*mm, 47*mm, 47*mm, 30*mm], style=_table_style()))
    story.append(Spacer(1, 3*mm))

    story.append(_section("2. Conclusiones clínicas continuas"))
    # Caja clínica continua: las cuatro conclusiones van juntas, antes de cualquier gráfico.
    conclusion_rows = []
    for title, body in conclusion_blocks:
        conclusion_rows.append([Paragraph(title, styles["MiniTitlePAC"])])
        conclusion_rows.append([Paragraph(body, styles["ConclusionPAC"])])
    story.append(Table(conclusion_rows, colWidths=[184*mm], style=TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#F7FAFC")),
        ("BOX", (0,0), (-1,-1), 0.5, colors.HexColor("#90A4AE")),
        ("INNERGRID", (0,0), (-1,-1), 0.15, colors.HexColor("#ECEFF1")),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 7),
        ("RIGHTPADDING", (0,0), (-1,-1), 7),
        ("TOPPADDING", (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ])))

    story.append(PageBreak())
    story.append(_section("3. Gráficos del informe"))
    story.append(Spacer(1, 2*mm))
    story.append(KeepTogether([
        Paragraph("Presión aórtica central con Pf y Pb superpuestas", styles["H3PAC"]),
        Image(plot_wave_separation(sep_df), width=182*mm, height=101*mm)
    ]))
    story.append(Spacer(1, 2*mm))

    img_w = 88*mm
    img_h = 51*mm
    graph_table = [
        [Paragraph("Presiones periféricas vs centrales", styles["H3PAC"]), Paragraph("Onda de presión aórtica central", styles["H3PAC"])],
        [Image(plot_pressure_comparison(row), width=img_w, height=img_h), Image(plot_waveform(wave_df), width=img_w, height=img_h)],
        [Paragraph("Flujo aórtico estimado", styles["H3PAC"]), Paragraph("Análisis armónico", styles["H3PAC"])],
        [Image(plot_aortic_flow(sep_df), width=img_w, height=img_h), Image(plot_harmonics(hdf), width=img_w, height=img_h)],
        [Paragraph("Semaforización clínica", styles["H3PAC"]), ""],
        [Image(plot_clinical_gauges(row, ppa), width=img_w, height=img_h), ""],
    ]
    story.append(Table(graph_table, colWidths=[92*mm, 92*mm], style=TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("BOX", (0,0), (-1,-1), 0.25, colors.HexColor("#CFD8DC")),
        ("INNERGRID", (0,0), (-1,-1), 0.15, colors.HexColor("#ECEFF1")),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ])))

    story.append(PageBreak())
    story.append(_section("4. Tabla de análisis armónico"))
    harm_table = [["Armónico", "Frecuencia (Hz)", "Amplitud", "Energía relativa (%)"]]
    for i, r in hdf.iterrows():
        harm_table.append([str(i+1), f"{r.get('frecuencia_hz',0):.2f}", f"{r.get('amplitud',0):.3f}", f"{r.get('energia_relativa_%',0):.1f}"])
    story.append(Table(harm_table, colWidths=[30*mm, 45*mm, 45*mm, 55*mm], style=_table_style("#EAF2F8")))
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph(
        "Nota metodológica: la separación Pf/Pb es una estimación clínica no invasiva. En el gráfico integrado, Pf y Pb se muestran sobre la línea diastólica basal para permitir comparación directa con la presión aórtica central completa.",
        styles["SmallPAC"]
    ))

    if screenshot_png:
        story.append(PageBreak())
        story.append(_section("5. Captura pantalla de mediciones"))
        story.append(Spacer(1, 3*mm))
        story.append(Image(io.BytesIO(screenshot_png), width=170*mm, height=220*mm))

    story.append(PageBreak())
    story.append(_section("6. Fenotipo final de presión central"))
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(final_phenotype, styles["H3PAC"]))
    story.append(Paragraph(final_phenotype_text, styles["BodyPAC"]))
    story.append(Spacer(1, 3*mm))
    phenotype_rows = [[Paragraph(str(cell), styles["SmallPAC"]) for cell in row_cells] for row_cells in final_phenotype_table]
    story.append(Table(phenotype_rows, colWidths=[45*mm, 22*mm, 108*mm], style=_table_style("#D9EAF7")))
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph(
        "Criterio operativo: el fenotipo final integra tres dimensiones complementarias: presión/carga pulsátil central, magnitud y temporalidad de la onda retrógrada, y complejidad armónica de la señal. No reemplaza el juicio clínico ni la validación del trazado original.",
        styles["SmallPAC"]
    ))

    story.append(PageBreak())
    story.append(_section("7. Referencias bibliográficas"))
    refs = [
        "Agabiti-Rosei E, et al. Central blood pressure measurements and antihypertensive therapy. Hypertension. 2007.",
        "Zócalo Y, Bia D. Presión aórtica central y parámetros clínicos derivados de la onda del pulso. 2014.",
        "SAHA. Manual de Mecánica Vascular. Grupo de Trabajo de Mecánica Vascular. 2024.",
        "Westerhof BE, et al. Quantification of wave reflection in the human aorta from pressure alone. Hypertension. 2006.",
        "Herbert A, et al. Establishing reference values for central blood pressure and amplification. Eur Heart J. 2014.",
        "Huang QF, et al. Outcome-driven threshold for pulse pressure amplification. Hypertension Research. 2024.",
    ]
    for i, ref_txt in enumerate(refs, 1):
        story.append(Paragraph(f"{i}. {ref_txt}", styles["SmallPAC"]))

    doc.build(story, onFirstPage=_header_footer, onLaterPages=_header_footer)
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
st.caption("Importación tipo MODELO PAC, informe PDF integrado, captura de segunda hoja, historial Excel y análisis armónico.")

with st.sidebar:
    st.header("1) Importar estudio")
    pdf_file = st.file_uploader("PDF original PAC / Exxer", type=["pdf"])
    wave_file = st.file_uploader("Opcional: CSV/TXT curva central (tiempo_ms, presion_mmHg)", type=["csv", "txt"])
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
        wave_df = read_curve_file_robust(wave_file, row)
        st.success("Curva importada correctamente con lector robusto CSV/TXT.")
    except Exception as e:
        st.warning(f"El archivo importado no contiene una curva válida. Se usará curva fisiológica calibrada con las métricas del estudio. Detalle: {e}")
        wave_df = make_waveform(row)
else:
    wave_df = make_waveform(row)

hdf = harmonic_analysis(wave_df)
dx, cat, ref, amp_sbp, ppa, risk = central_diagnosis(row)

st.subheader("Vista clínica previa")
sep_df_preview, sep_metrics_preview = estimate_wave_separation(wave_df, row)
conclusion_blocks_preview, sep_df_preview, sep_metrics_preview, sep_interp_preview = build_continuous_conclusions(row, wave_df, hdf)

summary_cols = st.columns(4)
summary_cols[0].metric("PAS central", f"{to_float(row.get('pas_central')):.0f} mmHg")
summary_cols[1].metric("PP central", f"{to_float(row.get('pp_central')):.0f} mmHg")
summary_cols[2].metric("PPA", f"{ppa:.2f}")
summary_cols[3].metric("RM Pb/Pf", f"{sep_metrics_preview.get('rm', np.nan):.2f}")

st.markdown("### Conclusiones clínicas continuas")
for title, body in conclusion_blocks_preview:
    st.markdown(f"**{title}**")
    st.write(body)

st.markdown("---")
st.markdown("### Gráficos")
st.image(plot_wave_separation(sep_df_preview), caption="Presión aórtica central con onda anterógrada Pf y retrógrada Pb superpuestas", use_container_width=True)

g1, g2 = st.columns(2)
with g1:
    st.image(plot_waveform(wave_df), caption="Onda central", use_container_width=True)
    st.image(plot_aortic_flow(sep_df_preview), caption="Flujo aórtico estimado redondeado", use_container_width=True)
with g2:
    st.image(plot_pressure_comparison(row), caption="Presiones periféricas vs centrales", use_container_width=True)
    st.image(plot_harmonics(hdf), caption="Armónicos de la onda central", use_container_width=True)

st.image(plot_clinical_gauges(row, ppa), caption="Semaforización clínica", use_container_width=True)

final_phenotype_preview, final_phenotype_text_preview, final_phenotype_table_preview = classify_central_pressure_phenotype(row, sep_metrics_preview, hdf)
st.markdown("---")
st.markdown("### Fenotipo final de presión central")
st.success(final_phenotype_preview)
st.write(final_phenotype_text_preview)
st.dataframe(pd.DataFrame(final_phenotype_table_preview[1:], columns=final_phenotype_table_preview[0]), use_container_width=True)

st.subheader("Historial y exportación")
if st.button("Guardar en historial"):
    hist = save_history(row)
    st.success(f"Registro guardado. Total: {len(hist)} estudios.")

if HISTORIAL_FILE.exists():
    hist = pd.read_excel(HISTORIAL_FILE)
    st.dataframe(hist, use_container_width=True)
    st.download_button("Descargar historial Excel", HISTORIAL_FILE.read_bytes(), file_name="historial_pac.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

pdf_bytes_out = build_pdf(row, wave_df, hdf, screenshot)
pdf_download_bytes = ensure_download_bytes(pdf_bytes_out)
if not pdf_download_bytes:
    st.error("No se pudo generar el PDF médico integrado.")
else:
    st.download_button(
        "Generar y descargar PDF médico integrado",
        data=pdf_download_bytes,
        file_name=f"PAC_IA_{str(row.get('paciente','paciente')).replace(' ','_')}.pdf",
        mime="application/pdf"
    )
