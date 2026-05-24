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

try:
    from PIL import Image as PILImage, ImageDraw
except Exception:
    PILImage = None
    ImageDraw = None

try:
    import cv2
except Exception:
    cv2 = None

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


def safe_trapezoid(y, x):
    """Integración compatible con NumPy 2.x/Streamlit Cloud.

    np.trapz fue retirado en algunas versiones recientes de NumPy; por eso
    usamos np.trapezoid si existe y dejamos fallback manual para evitar caídas
    del informe en producción.
    """
    y = np.asarray(y, dtype=float)
    x = np.asarray(x, dtype=float)
    ok = np.isfinite(y) & np.isfinite(x)
    y, x = y[ok], x[ok]
    if len(y) < 2:
        return 0.0
    if hasattr(np, "trapezoid"):
        return float(np.trapezoid(y, x))
    return float(np.sum((y[1:] + y[:-1]) * 0.5 * np.diff(x)))


def format_optional(v, dec=1):
    """Formatea valores numéricos evitando mostrar nan en informes clínicos."""
    try:
        f = float(v)
        if np.isnan(f):
            return "no disponible"
        return f"{f:.{dec}f}"
    except Exception:
        return "no disponible"




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
        raise ValueError("Curva real inválida: amplitud de presión insuficiente para calibración. No se generará curva sintética.")

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
    Modo estricto: si no hay curva real válida, se detiene el análisis de ondas/armónicos y no se usa curva sintética.
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


def _pixmap_to_rgb_array(page, zoom=3.0):
    """Renderiza una página PDF a matriz RGB para digitalización de curvas."""
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if arr.shape[2] >= 3:
        arr = arr[:, :, :3]
    return arr


def _curve_masks_from_rgb(arr):
    """Máscaras candidatas para detectar la curva real dibujada en el PDF.

    Se prioriza curva roja/magenta, típica en reportes PAC. Como respaldo se buscan
    curvas azules/verdes y, solo si no hay color, una máscara oscura más restrictiva.
    """
    rgb = arr.astype(np.int16)
    r, g, b = rgb[:, :, 0], rgb[:, :, 1], rgb[:, :, 2]
    # Curvas coloreadas: evita texto gris/negro y fondos claros.
    red = (r > 115) & (r > g + 35) & (r > b + 35)
    magenta = (r > 105) & (b > 80) & (r > g + 25) & (b > g + 15)
    blue = (b > 115) & (b > r + 35) & (b > g + 20)
    green = (g > 105) & (g > r + 25) & (g > b + 25)
    # Respaldo para curva negra: exige vecindad dentro de un gráfico y evita bordes externos.
    dark = (r < 70) & (g < 70) & (b < 70)
    return [
        ("roja", red | magenta),
        ("azul", blue),
        ("verde", green),
        ("oscura_restrictiva", dark),
    ]


def _clean_mask(mask):
    """Limpieza morfológica opcional sin exigir OpenCV en el entorno."""
    mask = mask.astype(np.uint8)
    if cv2 is None:
        return mask.astype(bool)
    kernel = np.ones((2, 2), np.uint8)
    m = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel, iterations=1)
    m = cv2.morphologyEx(m, cv2.MORPH_CLOSE, kernel, iterations=1)
    return m.astype(bool)


def _best_curve_component(mask, page_shape, color_name):
    """Selecciona el componente más compatible con una curva de presión.

    Puntúa componentes anchos, no demasiado altos, con aspecto horizontal y lejos de
    márgenes extremos. Esto evita seleccionar títulos, logos o textos coloreados.
    """
    h, w = page_shape[:2]
    mask = _clean_mask(mask)
    if mask.sum() < 80:
        return None

    if cv2 is None:
        ys, xs = np.where(mask)
        if len(xs) < 80:
            return None
        return (int(xs.min()), int(ys.min()), int(xs.max()), int(ys.max()), mask)

    n, labels, stats, _ = cv2.connectedComponentsWithStats(mask.astype(np.uint8), connectivity=8)
    best = None
    best_score = -1
    for lab in range(1, n):
        x, y, bw, bh, area = stats[lab]
        if area < 60 or bw < 80 or bh < 8:
            continue
        if bw > 0.95 * w or bh > 0.65 * h:
            continue
        aspect = bw / max(bh, 1)
        if aspect < 1.6:
            continue
        # Para máscara oscura, ser más exigente porque hay mucho texto/ejes.
        if color_name.startswith("oscura") and (bw < 0.20 * w or area < 120 or aspect < 2.5):
            continue
        margin_penalty = 0
        if y < 0.04*h or y+bh > 0.96*h or x < 0.02*w or x+bw > 0.98*w:
            margin_penalty = 0.55
        density = area / max(bw * bh, 1)
        score = (bw * 1.3 + area * 0.35 + aspect * 18) * (1 - margin_penalty) * (0.65 + min(density, 0.45))
        if score > best_score:
            comp = labels == lab
            best = (int(x), int(y), int(x+bw), int(y+bh), comp)
            best_score = score
    return best


def _digitize_curve_from_mask(mask, bbox, row, n_points=512):
    """Convierte píxeles de curva en puntos tiempo-presión.

    La presión se calibra linealmente contra PAS/PAD central reales del estudio. El
    tiempo se escala al ciclo visible 0-1000 ms, preservando la morfología del trazo.
    """
    x0, y0, x1, y1 = bbox
    sub = mask[y0:y1+1, x0:x1+1]
    ys, xs = np.where(sub)
    if len(xs) < 60:
        raise ValueError("No se detectaron suficientes píxeles de curva para digitalizar.")
    # Tomar una presión por columna: mediana de píxeles de trazo en esa columna.
    data = []
    for x in np.unique(xs):
        yy = ys[xs == x]
        if len(yy) == 0:
            continue
        data.append((float(x), float(np.median(yy))))
    if len(data) < 30:
        raise ValueError("La curva digitalizada tiene muy pocas columnas útiles.")
    data = np.asarray(data, dtype=float)
    xpix, ypix = data[:, 0], data[:, 1]

    # Filtro de saltos: conserva el contorno principal del trazo.
    y_s = pd.Series(ypix).rolling(5, center=True, min_periods=1).median().to_numpy()
    bad = np.abs(y_s - np.nanmedian(y_s)) > 4 * max(np.nanstd(y_s), 1.0)
    if np.mean(~bad) > 0.70:
        xpix, y_s = xpix[~bad], y_s[~bad]

    t = (xpix - np.nanmin(xpix)) / max(np.nanmax(xpix) - np.nanmin(xpix), 1e-6) * 1000.0
    # En imagen, menor y = mayor presión. Escala preliminar 0-1 y luego calibra.
    y_norm = (np.nanmax(y_s) - y_s) / max(np.nanmax(y_s) - np.nanmin(y_s), 1e-6)
    tmp = pd.DataFrame({"tiempo_ms": t, "presion_central_mmHg": y_norm})
    tmp = tmp.sort_values("tiempo_ms").drop_duplicates("tiempo_ms")
    t_grid = np.linspace(0, 1000, n_points)
    y_grid = np.interp(t_grid, tmp["tiempo_ms"], tmp["presion_central_mmHg"])

    pas = to_float(row.get("pas_central"))
    pad = to_float(row.get("pad_central"))
    pp = to_float(row.get("pp_central"))
    if np.isnan(pas) or np.isnan(pad) or pas <= pad:
        if not np.isnan(pp) and pp > 10 and not np.isnan(pad):
            pas = pad + pp
        else:
            raise ValueError("Para calibrar la curva digitalizada se requieren PAS/PAD central válidas.")
    p_grid = pad + (y_grid - np.nanmin(y_grid)) * (pas - pad) / max(np.nanmax(y_grid) - np.nanmin(y_grid), 1e-6)
    out = pd.DataFrame({"tiempo_ms": t_grid, "presion_central_mmHg": p_grid})
    ok, msg = is_physiologic_waveform(out, row)
    if not ok:
        raise ValueError("Curva digitalizada no supera validación fisiológica: " + msg)
    return out


def _annotate_digitized_region_png(arr, bbox, label):
    """Genera PNG diagnóstico con el rectángulo usado para digitalizar."""
    if PILImage is None or ImageDraw is None:
        return None
    img = PILImage.fromarray(arr.astype(np.uint8), mode="RGB")
    draw = ImageDraw.Draw(img)
    x0, y0, x1, y1 = bbox
    draw.rectangle([x0, y0, x1, y1], outline=(220, 30, 30), width=5)
    draw.text((x0, max(0, y0-24)), f"Curva digitalizada: {label}", fill=(220, 30, 30))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def _restrict_mask_to_pdf_curve_roi(mask, arr_shape, page_index):
    """Limita la búsqueda al panel real de curva del equipo PAC.

    Corrección por captura aportada:
    - La curva está en la SEGUNDA HOJA.
    - El panel se ubica ARRIBA A LA IZQUIERDA.
    - Debe incluir el gráfico completo "Pulso aórtico y radial promediado" y excluir
      la tabla central/derecha de métricas para no capturar números, barras o textos.

    Coordenadas relativas estimadas sobre la página renderizada:
    x: 0% a 50% de ancho
    y: 7% a 69% de alto
    """
    h, w = arr_shape[:2]
    roi = np.zeros_like(mask, dtype=bool)

    if page_index == 1:
        # Segunda hoja, panel superior izquierdo: gráfico de pulso aórtico/radial.
        # En la imagen ejemplo ocupa aproximadamente desde el margen izquierdo hasta
        # antes de la tabla de datos, y desde debajo del encabezado hasta la base del gráfico.
        x0 = int(0.000 * w)
        x1 = int(0.500 * w)
        y0 = int(0.070 * h)
        y1 = int(0.690 * h)
    else:
        # Respaldo conservador: mismo sector, apenas más amplio si cambia escala.
        x0 = int(0.000 * w)
        x1 = int(0.520 * w)
        y0 = int(0.060 * h)
        y1 = int(0.710 * h)

    roi[y0:y1, x0:x1] = True
    return mask & roi, (x0, y0, x1, y1)

def digitize_curve_from_pdf(pdf_bytes, row, max_pages=4, zoom=3.0, preferred_page_index=1):
    """Extrae/digitaliza la curva real desde la imagen del PDF cuando no hay CSV/TXT.

    Regla principal para MODELO PAC/Exxer:
    - Buscar primero en la segunda hoja del PDF.
    - Limitar la detección al sector superior izquierdo: panel de curva de esa hoja.
    - Calibrar exclusivamente con PAS/PAD central reales del paciente.

    Devuelve: wave_df, debug_png, metadata.
    No genera curvas sintéticas: si no detecta una curva válida, falla con mensaje explícito.
    """
    if fitz is None:
        raise ValueError("PyMuPDF/fitz no está disponible; no se puede renderizar el PDF para digitalizar la curva.")
    if not pdf_bytes:
        raise ValueError("No hay PDF cargado para digitalizar la curva.")
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    if len(doc) == 0:
        raise ValueError("El PDF no contiene páginas.")

    attempts = []

    # Prioridad absoluta: segunda hoja, sector superior izquierdo: panel de curva. Luego, si falla, páginas siguientes
    # también con ROI del panel superior izquierdo. No vuelve a la hoja 1 salvo como último respaldo.
    page_order = []
    if len(doc) > preferred_page_index:
        page_order.append(preferred_page_index)
    page_order += [i for i in range(min(len(doc), max_pages)) if i not in page_order and i != 0]
    if 0 not in page_order and len(doc) > 0:
        page_order.append(0)

    for pi in page_order:
        page = doc[pi]
        arr = _pixmap_to_rgb_array(page, zoom=zoom)
        for color_name, raw_mask in _curve_masks_from_rgb(arr):
            raw_mask, roi_bbox = _restrict_mask_to_pdf_curve_roi(raw_mask, arr.shape, pi)
            comp = _best_curve_component(raw_mask, arr.shape, color_name)
            if comp is None:
                attempts.append(f"página {pi+1} sector superior izquierdo: panel de curva {color_name}: sin componente compatible")
                continue
            x0, y0, x1, y1, comp_mask = comp
            # Expandir un poco para conservar extremos del trazo, no para calibrar contra ejes.
            pad_x = int(max(3, 0.020 * (x1-x0)))
            pad_y = int(max(3, 0.12 * (y1-y0)))
            x0e = max(0, x0 - pad_x); x1e = min(arr.shape[1]-1, x1 + pad_x)
            y0e = max(0, y0 - pad_y); y1e = min(arr.shape[0]-1, y1 + pad_y)
            try:
                wave = _digitize_curve_from_mask(comp_mask, (x0, y0, x1, y1), row)
                debug = _annotate_digitized_region_png(
                    arr,
                    (x0e, y0e, x1e, y1e),
                    f"pág. {pi+1} / sector superior izquierdo: panel de curva / {color_name}"
                )
                meta = {
                    "pagina": pi + 1,
                    "sector": "izquierdo-superior de la hoja",
                    "pagina_preferida": 2,
                    "color_detectado": color_name,
                    "bbox_px": (x0, y0, x1, y1),
                    "roi_px": roi_bbox,
                    "puntos": int(len(wave)),
                    "metodo": "digitalización automática desde segunda hoja, sector superior izquierdo: panel de curva, calibrada con PAS/PAD central",
                }
                return wave, debug, meta
            except Exception as e:
                attempts.append(f"página {pi+1} sector superior izquierdo: panel de curva {color_name}: {e}")
    raise ValueError("No se pudo digitalizar una curva central válida desde la segunda hoja, sector superior izquierdo: panel de curva del PDF. " + " | ".join(attempts[:10]))


def find_after(label, text, default=""):
    pat = re.compile(label + r"\s*[:#]?\s*([^\n]+)", re.I)
    m = pat.search(text)
    return safe_text(m.group(1)) if m else default


# -----------------------------
# Parser robusto de datos PAC
# -----------------------------
def _collapse_spaces(s):
    return re.sub(r"\s+", " ", safe_text(s)).strip()



def _line_text_after_label(line_text, label_regex):
    """Devuelve texto a la derecha de una etiqueta dentro de una línea visual."""
    m = re.search(label_regex, line_text, flags=re.I)
    if not m:
        return ""
    return _collapse_spaces(line_text[m.end():]).strip(" :#-")


def _layout_lines_from_pdf(pdf_bytes, preferred_page_index=1):
    """Extrae líneas visuales con coordenadas usando PyMuPDF.

    Esto corrige el problema típico de pdfplumber en estos informes: mezcla columnas y
    hace que Paciente tome valores de otra etiqueta como 'Número de estudio'.
    """
    if fitz is None or not pdf_bytes:
        return []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page_indices = []
    if len(doc) > preferred_page_index:
        page_indices.append(preferred_page_index)
    page_indices += [i for i in range(len(doc)) if i not in page_indices]
    all_lines = []
    for pi in page_indices:
        page = doc[pi]
        words = page.get_text("words") or []
        # words: x0, y0, x1, y1, text, block, line, word_no
        grouped = {}
        for w in words:
            if len(w) < 8:
                continue
            x0, y0, x1, y1, txt, block, line, word_no = w[:8]
            key = (pi, int(block), int(line))
            grouped.setdefault(key, []).append((float(x0), float(y0), float(x1), float(y1), str(txt)))
        for key, ws in grouped.items():
            ws = sorted(ws, key=lambda z: z[0])
            text = _collapse_spaces(" ".join(w[4] for w in ws))
            if not text:
                continue
            x0 = min(w[0] for w in ws); y0 = min(w[1] for w in ws)
            x1 = max(w[2] for w in ws); y1 = max(w[3] for w in ws)
            all_lines.append({"page": key[0], "x0": x0, "y0": y0, "x1": x1, "y1": y1, "text": text, "words": ws})
    all_lines.sort(key=lambda d: (d["page"], d["y0"], d["x0"]))
    return all_lines


def _numeric_values_in_text(txt):
    return [to_float(x) for x in re.findall(r"[-+]?\d+(?:[\.,]\d+)?", txt)]


def _extract_layout_fields_from_pdf(pdf_bytes):
    """Extractor visual de cabecera y métricas desde el PDF original.

    Prioriza la segunda página porque allí está el panel del estudio mostrado por el usuario.
    Extrae paciente, estudio, demografía, tabla Radial/Central y aumentaciones desde líneas
    visuales, no desde texto plano concatenado.
    """
    out = {}
    lines = _layout_lines_from_pdf(pdf_bytes, preferred_page_index=1)
    if not lines:
        return out

    # 1) Paciente y estudio por línea visual.
    patient_candidates = []
    for ln in lines:
        txt = ln["text"]
        if re.search(r"\bPaciente\b", txt, flags=re.I):
            cand = _line_text_after_label(txt, r"\bPaciente\b")
            cand = re.split(r"\b(?:Estudio\s*#?|N[úu]mero\s+de\s+estudio|Fecha|Hora|Edad|Sexo|Peso|Altura|IMC|SC|Diagn[oó]stico|Medicaci[oó]n)\b", cand, flags=re.I)[0]
            cand = _clean_patient_name(cand)
            if cand and not _is_bad_patient_value(cand):
                # preferir nombres con al menos dos tokens alfabéticos y ubicación derecha/superior del informe
                alpha_tokens = re.findall(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]{2,}", cand)
                score = len(alpha_tokens) + (2 if ln["page"] == 1 else 0) + (1 if ln["y0"] < 180 else 0)
                patient_candidates.append((score, cand))
    if patient_candidates:
        out["paciente"] = sorted(patient_candidates, reverse=True)[0][1]

    for ln in lines:
        txt = ln["text"]
        m = re.search(r"\bEstudio\s*#?\b\s*[:#]?\s*([A-Za-z0-9][A-Za-z0-9_\-/]{0,30})", txt, re.I)
        if m:
            val = _collapse_spaces(m.group(1)).strip(" :#-")
            val = re.split(r"\bH\.?\s*C\.?\b|#|Paciente|Fecha|Hora|Edad|Sexo", val, flags=re.I)[0].strip(" :#-")
            if val and not re.fullmatch(r"(?i)(M|F|Paciente|Fecha|Hora|Edad|Sexo)", val):
                out["estudio"] = val
                break

    # 2) Demografía/antropometría por línea visual.
    def put_num_from_line(label, key):
        for ln in lines:
            txt = ln["text"]
            if re.search(rf"\b{label}\b", txt, flags=re.I):
                valtxt = _line_text_after_label(txt, rf"\b{label}\b")
                vals = _numeric_values_in_text(valtxt)
                if vals:
                    out[key] = vals[0]
                    return
    for label, key in [("Edad", "edad"), ("Peso", "peso"), ("Altura", "altura"), ("IMC", "imc"), ("SC", "sc")]:
        put_num_from_line(label, key)
    for ln in lines:
        txt = ln["text"]
        m = re.search(r"\bSexo\b\s*[:#]?\s*([MF])\b", txt, re.I)
        if m:
            out["sexo"] = m.group(1).upper(); break
    for ln in lines:
        txt = ln["text"]
        m = re.search(r"\bFecha\b\s*[:#]?\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})", txt, re.I)
        if m:
            out["fecha"] = m.group(1); break
    for ln in lines:
        txt = ln["text"]
        m = re.search(r"\bHora\b\s*[:#]?\s*(\d{1,2}:\d{2}(?::\d{2})?)", txt, re.I)
        if m:
            out["hora"] = m.group(1); break

    # 3) Tabla Radial/Central: buscar filas PAS/PAD/PAM/PP con dos primeros números.
    rc_keys = {"PAS": ("pas_radial", "pas_central"), "PAD": ("pad_radial", "pad_central"),
               "PAM": ("pam_radial", "pam_central"), "PP": ("pp_radial", "pp_central")}
    for ln in lines:
        txt = ln["text"]
        for lab, (kr, kc) in rc_keys.items():
            if re.match(rf"^\s*{lab}\b", txt, flags=re.I):
                vals = _numeric_values_in_text(txt)
                # En la tabla Radial/Central, la fila contiene dos valores: radial y central.
                if len(vals) >= 2:
                    # Evitar capturar la sección de parámetros centrales; allí la línea tiene unidad mmHg y +/-.
                    if not re.search(r"\+/-|mmHg|%", txt, flags=re.I) or lab in ["PAS", "PP"] and len(vals) >= 2:
                        # Si está en parámetros centrales, x suele estar más abajo y cerca del título. Se prioriza solo si no hay valor todavía.
                        if kr not in out or kc not in out:
                            out[kr] = vals[0]; out[kc] = vals[1]
    # FC: línea de la tabla principal.
    for ln in lines:
        txt = ln["text"]
        if re.match(r"^\s*FC\b", txt, flags=re.I):
            vals = _numeric_values_in_text(txt)
            if vals:
                out["fc"] = vals[0]
                break

    # 4) Parámetros hemodinámicos centrales / aumentaciones. Buscar filas específicas.
    central_map = {
        "PAS": "pas_central", "PP": "pp_central", "Au": "au", "IAu": "iau",
        "RVSE": "rvse", "PE": "pe", "APC": "apc"
    }
    for ln in lines:
        txt = ln["text"]
        for lab, key in central_map.items():
            if re.match(rf"^\s*{lab}\b", txt, flags=re.I):
                # Las filas centrales suelen traer unidad y/o +/-: 'Au mmHg +8 +/-1'
                if lab in ["Au", "IAu", "RVSE", "PE", "APC"] or re.search(r"mmHg|%|\+/-", txt, flags=re.I):
                    vals = _numeric_values_in_text(txt)
                    if vals:
                        # Para PAS/PP central, no pisar si la tabla radial/central ya extrajo lo correcto salvo que sea sección central con unidad.
                        out[key] = vals[0]
    return out




def _patient_name_strict(candidate):
    """Acepta solo nombres reales; rechaza frases administrativas como 'en posición'."""
    c = _clean_patient_name(candidate)
    if not c:
        return ""
    bad = r"\b(en\s+posici[oó]n|posici[oó]n|realizado|paciente|estudio|fecha|hora|diagn[oó]stico|medicaci[oó]n|n[úu]mero)\b"
    if re.search(bad, c, re.I):
        return ""
    toks = re.findall(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]{2,}", c)
    if len(toks) < 2:
        return ""
    # Preferir nombres en mayúsculas del PDF. Si todo está minúscula, suele ser frase de sección.
    upper_tokens = sum(1 for t in toks if t.upper() == t)
    if upper_tokens < 2 and not any(t.istitle() for t in toks):
        return ""
    return " ".join(toks[:5]).upper()


def _extract_patient_study_by_words(pdf_bytes):
    """Extractor por palabras y coordenadas: toma texto a la derecha de Paciente/Estudio en la misma línea visual."""
    out = {}
    if fitz is None or not pdf_bytes:
        return out
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        page_order = list(range(len(doc)))
        # priorizar segunda página, luego primera
        if len(page_order) > 1:
            page_order = [1, 0] + [i for i in page_order if i not in (0,1)]
        for pi in page_order:
            words = page_order and (doc[pi].get_text("words") or [])
            # agrupar por y aproximada para sort visual incluso si block/line se parte mal
            rows = []
            for w in words:
                if len(w) < 5: continue
                x0,y0,x1,y1,txt = float(w[0]),float(w[1]),float(w[2]),float(w[3]),str(w[4])
                placed=False
                for row in rows:
                    if abs(row[0]-y0) <= 4.0:
                        row[1].append((x0,y0,x1,y1,txt)); placed=True; break
                if not placed:
                    rows.append([y0, [(x0,y0,x1,y1,txt)]])
            for _, rowwords in rows:
                rowwords=sorted(rowwords, key=lambda z:z[0])
                texts=[w[4] for w in rowwords]
                line=_collapse_spaces(" ".join(texts))
                # Estudio #: primer número posterior a Estudio
                if "estudio" in line.lower() and "estudio" not in out:
                    m=re.search(r"\bEstudio\s*#?\s*([0-9A-Za-z][0-9A-Za-z_\-/]*)", line, re.I)
                    if m:
                        val=re.split(r"\bH\.?\s*C\.?\b|#|Paciente|Fecha|Hora|Edad|Sexo", m.group(1), flags=re.I)[0].strip(" :#-")
                        if val and not re.fullmatch(r"(?i)(M|F|Paciente|Fecha|Hora|Edad|Sexo)", val):
                            out["estudio"] = val
                # Paciente: tomar tokens a la derecha de la palabra Paciente hasta próxima etiqueta
                for idx, w in enumerate(rowwords):
                    if re.fullmatch(r"Paciente:?", w[4], re.I):
                        right=[]
                        for ww in rowwords[idx+1:]:
                            tx=ww[4]
                            if re.search(r"(?i)^(Estudio|N[uú]mero|Fecha|Hora|Edad|Sexo|Peso|Altura|IMC|SC|Diagn|Medic|Realizado|Abdomen|Cuello|H\.?C\.?)", tx):
                                break
                            right.append(tx)
                        cand=_patient_name_strict(" ".join(right))
                        if cand:
                            out["paciente"] = cand
                            return out if out.get("estudio") else out
        return out
    except Exception:
        return out




def _page_words_grouped_rows(pdf_bytes, page_index=1, y_tol=4.5):
    """Devuelve filas visuales por coordenadas PyMuPDF para evitar mezclas de pdfplumber."""
    if fitz is None or not pdf_bytes:
        return [], None
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if len(doc) == 0:
            return [], None
        pi = page_index if len(doc) > page_index else 0
        page = doc[pi]
        W, H = float(page.rect.width), float(page.rect.height)
        words = page.get_text("words") or []
        rows = []
        for w in words:
            if len(w) < 5:
                continue
            x0, y0, x1, y1, txt = float(w[0]), float(w[1]), float(w[2]), float(w[3]), str(w[4])
            if not txt.strip():
                continue
            placed = False
            yc = (y0 + y1) / 2.0
            for row in rows:
                if abs(row[0] - yc) <= y_tol:
                    row[1].append((x0, y0, x1, y1, txt)); placed = True; break
            if not placed:
                rows.append([yc, [(x0, y0, x1, y1, txt)]])
        out = []
        for yc, ws in rows:
            ws = sorted(ws, key=lambda z: z[0])
            text = _collapse_spaces(" ".join(w[4] for w in ws))
            out.append({"y": yc, "words": ws, "text": text, "page": pi, "W": W, "H": H})
        out.sort(key=lambda r: (r["y"], min(w[0] for w in r["words"])))
        return out, (W, H, pi)
    except Exception:
        return [], None


def _numbers_from_row_words(rowwords, x_min=None, x_max=None):
    vals = []
    for x0, y0, x1, y1, txt in rowwords:
        if x_min is not None and x0 < x_min:
            continue
        if x_max is not None and x0 > x_max:
            continue
        if re.fullmatch(r"[-+]?\d+(?:[\.,]\d+)?", str(txt).strip()):
            vals.append((x0, to_float(txt)))
    return vals


def _extract_patient_top_right_by_coordinates(pdf_bytes):
    """Lee Paciente/Estudio desde la cabecera visual, evitando 'Realizado con paciente en posición'."""
    out = {}
    rows, meta = _page_words_grouped_rows(pdf_bytes, page_index=1, y_tol=4.5)
    if not rows or not meta:
        return out
    W, H, _ = meta
    # priorizar cabecera superior derecha, pero permitir toda la parte superior.
    for row in rows:
        txt = row["text"]
        if row["y"] > 0.33 * H:
            continue
        if re.search(r"(?i)\bRealizado\b|\bposici[oó]n\b", txt):
            continue
        # Estudio # 684
        if "estudio" not in out:
            m = re.search(r"\bEstudio\s*#?\s*([0-9A-Za-z][0-9A-Za-z_\-/]*)", txt, re.I)
            if m:
                val = re.split(r"\bH\.?\s*C\.?\b|#|Paciente|Fecha|Hora|Edad|Sexo", m.group(1), flags=re.I)[0].strip(" :#-")
                if val and not re.fullmatch(r"(?i)(M|F|Paciente|Fecha|Hora|Edad|Sexo)", val):
                    out["estudio"] = val
        # Paciente ABEL ALEJANDRO SANCHO: solo si la palabra Paciente inicia etiqueta y no frase interna.
        for idx, ww in enumerate(row["words"]):
            tx = ww[4].strip()
            if re.fullmatch(r"Paciente:?", tx, flags=re.I):
                # Rechazar si aparece precedido por 'con', 'del', etc.
                if idx > 0 and re.search(r"(?i)^(con|del|el|la)$", row["words"][idx-1][4]):
                    continue
                right = []
                for w2 in row["words"][idx+1:]:
                    tx2 = w2[4].strip()
                    if re.search(r"(?i)^(Estudio|N[uú]mero|Fecha|Hora|Edad|Sexo|Peso|Altura|IMC|SC|Diagn|Medic|Realizado|Abdomen|Cuello|H\.?C\.?)", tx2):
                        break
                    right.append(tx2)
                cand = _patient_name_strict(" ".join(right))
                if cand:
                    out["paciente"] = cand
                    return out
    return out


def _extract_tables_by_coordinates(pdf_bytes):
    """Extrae Radial/Central y parámetros centrales por coordenadas visuales.

    Diseñado para el formato mostrado: segunda hoja, tabla Radial/Central y bloque
    'Parámetros hemodinámicos centrales'. Evita que PAM y PP se intercambien y
    recupera Au/IAu/RVSE/PE aunque el texto plano se parta por columnas.
    """
    out = {}
    rows, meta = _page_words_grouped_rows(pdf_bytes, page_index=1, y_tol=5.5)
    if not rows or not meta:
        return out
    W, H, _ = meta

    # Buscar encabezados Radial / Central para ubicar columnas.
    radial_x = central_x = header_y = None
    for row in rows:
        if re.search(r"\bRadial\b", row["text"], re.I) and re.search(r"\bCentral\b", row["text"], re.I):
            for w in row["words"]:
                if re.fullmatch(r"Radial", w[4], re.I): radial_x = (w[0]+w[2])/2
                if re.fullmatch(r"Central", w[4], re.I): central_x = (w[0]+w[2])/2
            header_y = row["y"]
            break
    if radial_x is not None and central_x is not None and header_y is not None:
        label_map = {"PAS":("pas_radial","pas_central"), "PAD":("pad_radial","pad_central"), "PAM":("pam_radial","pam_central"), "PP":("pp_radial","pp_central")}
        for row in rows:
            if not (header_y < row["y"] < header_y + 0.28*H):
                continue
            words = row["words"]
            first = words[0][4].strip() if words else ""
            for lab, (kr, kc) in label_map.items():
                if re.fullmatch(lab, first, re.I):
                    nums = _numbers_from_row_words(words, x_min=words[0][2])
                    if len(nums) >= 2:
                        # Asignar por cercanía a los centros de las columnas Radial/Central.
                        nr = min(nums, key=lambda z: abs(z[0]-radial_x))[1]
                        nc = min(nums, key=lambda z: abs(z[0]-central_x))[1]
                        out[kr] = nr; out[kc] = nc
            if re.fullmatch(r"FC", first, re.I):
                nums = _numbers_from_row_words(words, x_min=words[0][2])
                if nums:
                    out["fc"] = nums[0][1]

    # Bloque de parámetros hemodinámicos centrales.
    param_y = None
    for row in rows:
        if re.search(r"Par[aá]metros\s+hemodin[aá]micos\s+centrales", row["text"], re.I):
            param_y = row["y"]; break
    param_labels = {"PAS":"pas_central", "PP":"pp_central", "Au":"au", "IAu":"iau", "RVSE":"rvse", "PE":"pe", "APC":"apc"}
    for row in rows:
        if param_y is not None:
            if not (param_y < row["y"] < param_y + 0.30*H):
                continue
        else:
            if not (row["y"] > 0.55*H and min(w[0] for w in row["words"]) > 0.25*W):
                continue
        words = row["words"]
        text = row["text"].replace("I Au", "IAu")
        # Tomar etiqueta como primera palabra o cerca del inicio de la fila.
        lab = None
        for cand in ["IAu", "RVSE", "APC", "PAS", "PP", "Au", "PE"]:
            if re.search(rf"(^|\s){cand}(\s|$)", text, re.I):
                lab = cand; break
        if not lab:
            continue
        # Buscar primer número situado después de la etiqueta y después de unidad; evita capturar tolerancias +/-.
        label_x1 = None
        for w in words:
            if re.fullmatch(lab, w[4], re.I) or (lab == "IAu" and re.fullmatch(r"I?Au", w[4], re.I)):
                label_x1 = w[2]; break
        if label_x1 is None:
            label_x1 = min(w[0] for w in words)
        nums = _numbers_from_row_words(words, x_min=label_x1)
        if nums:
            # En filas con '+/-', el primer número tras etiqueta/unidad es el valor principal.
            out[param_labels[lab]] = nums[0][1]

    # APC del recuadro: 'APC:' '(1.08)' suele estar en el panel de curva superior izquierdo.
    for row in rows:
        if re.search(r"\bAPC\b", row["text"], re.I):
            nums = _numbers_from_row_words(row["words"])
            if nums:
                # suele ser decimal 1.08; preferir valor entre 0.2 y 3.5
                valid = [v for _, v in nums if 0.2 <= v <= 3.5]
                if valid:
                    out["apc"] = valid[0]
                    break
    return out

def _extract_central_params_global(flat_text):
    """Extrae Au/IAu/RVSE/PE/APC y corrige PAS/PP centrales con patrones globales robustos."""
    out={}
    txt=_collapse_spaces(flat_text).replace("I Au", "IAu").replace("R V S E", "RVSE").replace("P E", "PE")
    # APC aparece en el recuadro superior: APC: (1.08)
    m=re.search(r"\bAPC\b\s*:?\s*\(?\s*([-+]?\d+(?:[\.,]\d+)?)", txt, re.I)
    if m: out["apc"] = to_float(m.group(1))
    # Parámetros centrales: usar aparición con unidad y +/- para evitar tabla radial/central
    patterns={
        "pas_central": r"\bPAS\b\s*mmHg\s*([-+]?\d+(?:[\.,]\d+)?)",
        "pp_central": r"\bPP\b\s*mmHg\s*([-+]?\d+(?:[\.,]\d+)?)",
        "au": r"(?<!I)\bAu\b\s*mmHg\s*([-+]?\d+(?:[\.,]\d+)?)",
        "iau": r"\bIAu\b\s*%\s*([-+]?\d+(?:[\.,]\d+)?)",
        "rvse": r"\bRVSE\b\s*%\s*([-+]?\d+(?:[\.,]\d+)?)",
        "pe": r"\bPE\b\s*%\s*([-+]?\d+(?:[\.,]\d+)?)",
    }
    for k,pat in patterns.items():
        ms=list(re.finditer(pat, txt, re.I))
        if ms:
            # tomar la última suele corresponder a la tabla de parámetros, no a leyenda
            out[k]=to_float(ms[-1].group(1))
    return out


def _repair_swapped_pam_pp(data):
    """Repara intercambios PAM/PP y recalcula PP si no coincide con PAS-PAD.

    En este formato del PDF, es frecuente que la lectura de columnas cargue PAM como PP
    y PP como PAM. La prioridad clínica es: PP = PAS - PAD; PAM aproximada = (PAS+2*PAD)/3.
    """
    pas_r,pad_r,pam_r,pp_r = [to_float(data.get(k)) for k in ("pas_radial","pad_radial","pam_radial","pp_radial")]
    pas_c,pad_c,pam_c,pp_c = [to_float(data.get(k)) for k in ("pas_central","pad_central","pam_central","pp_central")]

    def close(a,b,tol):
        return (not np.isnan(a)) and (not np.isnan(b)) and abs(a-b) <= tol

    if not np.isnan(pas_r) and not np.isnan(pad_r) and pas_r > pad_r:
        pp_exp = pas_r - pad_r
        pam_exp = (pas_r + 2*pad_r) / 3.0
        # Caso clásico: PAM=40 y PP=102 para PAS/PAD 127/87.
        if close(pam_r, pp_exp, 8) and close(pp_r, pam_exp, 15):
            data["pam_radial"], data["pp_radial"] = pp_r, pam_r
            pam_r, pp_r = pp_r, pam_r
        # Si PP no coincide con PAS-PAD, corregir PP; si PAM falta, completar aproximada.
        if np.isnan(pp_r) or abs(pp_r - pp_exp) > 12 or pp_r < 10 or pp_r > 120:
            data["pp_radial"] = pp_exp
        if np.isnan(pam_r) or pam_r < 40 or pam_r > 200:
            data["pam_radial"] = pam_exp

    if not np.isnan(pas_c) and not np.isnan(pad_c) and pas_c > pad_c:
        pp_exp = pas_c - pad_c
        pam_exp = (pas_c + 2*pad_c) / 3.0
        # Caso clásico: PAM=31 y PP=102/80 para PAS/PAD 119/88.
        if close(pam_c, pp_exp, 8) and close(pp_c, pam_exp, 15):
            data["pam_central"], data["pp_central"] = pp_c, pam_c
            pam_c, pp_c = pp_c, pam_c
        if np.isnan(pp_c) or abs(pp_c - pp_exp) > 12 or pp_c < 10 or pp_c > 120:
            data["pp_central"] = pp_exp
        if np.isnan(pam_c) or pam_c < 40 or pam_c > 200:
            data["pam_central"] = pam_exp
    return data

def parse_model_pac_from_pdf(pdf_bytes, fallback_text=""):
    """Parser principal: texto plano + corrección visual por coordenadas del PDF."""
    data = parse_model_pac(fallback_text or "")
    try:
        # 0) Extracción estricta por coordenadas de cabecera y tablas reales del PDF.
        head0 = _extract_patient_top_right_by_coordinates(pdf_bytes)
        for k, v in head0.items():
            if k == "paciente":
                cand = _patient_name_strict(v)
                if cand:
                    data[k] = cand
            elif v not in [None, ""]:
                data[k] = v

        coord_vars = _extract_tables_by_coordinates(pdf_bytes)
        for k, v in coord_vars.items():
            fv = to_float(v)
            if not np.isnan(fv):
                data[k] = fv

        # 1) Corregir paciente/estudio con palabras visuales; evita 'en posición' y 'H.C. #'.
        head = _extract_patient_study_by_words(pdf_bytes)
        for k, v in head.items():
            if k == "paciente":
                if _patient_name_strict(v):
                    data[k] = _patient_name_strict(v)
            elif v not in [None, ""]:
                data[k] = v

        # 2) Corrección visual amplia de variables numéricas.
        layout = _extract_layout_fields_from_pdf(pdf_bytes)
        for k, v in layout.items():
            if k == "paciente":
                cand = _patient_name_strict(v)
                if cand:
                    data[k] = cand
            elif k in ["estudio", "fecha", "hora", "sexo"]:
                if v not in [None, ""]:
                    if k == "estudio":
                        v = re.split(r"\bH\.?\s*C\.?\b|#|Paciente|Fecha|Hora|Edad|Sexo", str(v), flags=re.I)[0].strip(" :#-")
                    data[k] = v
            else:
                fv = to_float(v)
                if not np.isnan(fv):
                    data[k] = fv

        # 3) Aumentaciones y parámetros centrales desde texto global, más robusto que líneas partidas.
        flat = _collapse_spaces(fallback_text or "")
        for k, v in _extract_central_params_global(flat).items():
            if not np.isnan(to_float(v)):
                data[k] = v

        # 4) Reaplicar coordenadas al final: tienen prioridad sobre texto plano mezclado.
        for k, v in coord_vars.items():
            fv = to_float(v)
            if not np.isnan(fv):
                data[k] = fv
        if head0.get("paciente") and _patient_name_strict(head0.get("paciente")):
            data["paciente"] = _patient_name_strict(head0.get("paciente"))

        # 5) Reparar PAM/PP intercambiadas y valores contaminados antes de validar.
        data = _repair_swapped_pam_pp(data)
        data = _validate_and_repair_pac_data(data)
        data = _repair_swapped_pam_pp(data)
        data = _validate_and_repair_pac_data(data)
    except Exception:
        data = _validate_and_repair_pac_data(data)
    return data


def _strip_trailing_labels(value):
    """Corta valores que quedaron contaminados por la etiqueta siguiente."""
    v = _collapse_spaces(value)
    stop = r"\b(Estudio\s*#?|Número\s+de\s+estudio|Paciente|Fecha|Hora|Edad|Sexo|Peso|Altura|IMC|SC|Diagnóstico|Diagnostico|Medicación|Medicacion|Abdomen|Cuello|Realizado|Radial|Central|PAS|PAD|PAM|PP|FC|Parámetros)\b"
    m = re.search(stop, v, flags=re.I)
    if m:
        if m.start() == 0:
            return ""
        v = v[:m.start()].strip()
    return v.strip(" :;,-")


def _value_after_label(flat_text, label_patterns, stop_patterns=None, max_chars=80):
    """Extrae texto posterior a una etiqueta, incluso cuando PDFPlumber une columnas en la misma línea."""
    if isinstance(label_patterns, str):
        label_patterns = [label_patterns]
    if stop_patterns is None:
        stop_patterns = [
            r"Estudio\s*#?", r"Número\s+de\s+estudio", r"Paciente", r"Fecha", r"Hora", r"Edad", r"Sexo",
            r"Peso", r"Altura", r"IMC", r"SC", r"Diagnóstico", r"Diagnostico", r"Medicación", r"Medicacion",
            r"Abdomen", r"Cuello", r"Realizado", r"Radial", r"Central", r"PAS", r"PAD", r"PAM", r"PP", r"FC"
        ]
    stop_re = "|".join(f"(?:{p})" for p in stop_patterns)
    for lab in label_patterns:
        pat = re.compile(rf"(?:^|\b){lab}\s*(?:[:#])?\s*(.*?)(?=\s+(?:{stop_re})\b|$)", re.I)
        m = pat.search(flat_text)
        if m:
            val = _strip_trailing_labels(m.group(1)[:max_chars])
            if val and val not in ["---", "--", "-"]:
                return val
    return ""


def _number_after_label(flat_text, label_patterns, default=np.nan):
    val = _value_after_label(flat_text, label_patterns, max_chars=50)
    m = re.search(r"[-+]?\d+(?:[\.,]\d+)?", val)
    if m:
        return to_float(m.group(0))
    if isinstance(label_patterns, str):
        label_patterns = [label_patterns]
    for lab in label_patterns:
        # Respaldo: primer número razonable inmediatamente luego de la etiqueta.
        pat = re.compile(rf"(?:^|\b){lab}\s*(?:[:#])?\s*([-+]?\d+(?:[\.,]\d+)?)", re.I)
        m = pat.search(flat_text)
        if m:
            return to_float(m.group(1))
    return default


def _normalize_sex(value):
    v = _collapse_spaces(value).upper()
    if re.search(r"\b(F|FEMENINO|MUJER)\b", v):
        return "F"
    if re.search(r"\b(M|MASCULINO|VARON|VARÓN|HOMBRE)\b", v):
        return "M"
    return ""


def _clean_patient_name(value):
    """Limpia y valida nombre de paciente evitando que entren etiquetas del PDF."""
    v = _strip_trailing_labels(value)
    v = re.sub(r"\s{2,}", " ", v).strip(" :;,-")
    blacklist = r"\b(n[úu]mero\s+de\s+estudio|numero\s+de\s+estudio|estudio\s*#?|fecha|hora|edad|sexo|peso|altura|imc|sc|diagn[oó]stico|medicaci[oó]n|radial|central|pas|pad|pam|pp|fc|par[aá]metros|realizado|posici[oó]n|paciente\s+en\s+posici[oó]n|en\s+posici[oó]n)\b"
    if re.search(blacklist, v or "", flags=re.I):
        return ""
    if re.fullmatch(r"(?i)(m|f)", v or ""):
        return ""
    if len(v) < 3 or re.fullmatch(r"[\d\s:;#\-/.]+", v):
        return ""
    # Evitar valores muy largos provenientes de una fila entera del PDF.
    if len(v.split()) > 8:
        return ""
    return v


def _extract_header_by_regex(flat_text):
    """Rescate fuerte de cabecera cuando pdfplumber mezcla columnas.

    Busca patrones reales del PDF tipo Exxer:
    Estudio # 684 / Paciente ABEL ALEJANDRO SANCHO / Edad 54 / Sexo M.
    """
    out = {}
    flat = _collapse_spaces(flat_text)

    m = re.search(r"\bEstudio\s*#?\s*[:#]?\s*([0-9A-Za-z][0-9A-Za-z_\-/]{0,20})", flat, re.I)
    if m:
        val = _collapse_spaces(m.group(1)).strip(" :#-;")
        if val and not re.fullmatch(r"(?i)(m|f|paciente|fecha|hora|edad|sexo)", val):
            out["estudio"] = val

    # Captura el nombre hasta la próxima etiqueta demográfica/administrativa.
    # Se recorren TODAS las ocurrencias de "Paciente" porque algunos PDFs traen primero
    # un encabezado contaminado como "Paciente Número de estudio:".
    patient_patterns = [
        r"\bPaciente\b\s*[:#]?\s*([A-ZÁÉÍÓÚÜÑ][A-ZÁÉÍÓÚÜÑ ,.'\-]{2,80}?)(?=\s+\b(?:Edad|Sexo|Peso|Altura|IMC|SC|Abdomen|Cuello|Diagn[oó]stico|Medicaci[oó]n|Realizado|Fecha|Hora|Estudio)\b|$)",
        r"\bPaciente\b\s*[:#]?\s*([^\n\r]{3,90}?)(?=\s+\b(?:Edad|Sexo|Peso|Altura|IMC|SC|Abdomen|Cuello|Diagn[oó]stico|Medicaci[oó]n|Realizado|Fecha|Hora|Estudio)\b|$)",
    ]
    for pat in patient_patterns:
        for m in re.finditer(pat, flat, re.I):
            cand_raw = m.group(1)
            cand_raw = re.split(r"\b(?:N[úu]mero\s+de\s+estudio|Numero\s+de\s+estudio|Estudio\s*#?|Fecha|Hora|Edad|Sexo|Peso|Altura|IMC|SC|Diagn[oó]stico|Medicaci[oó]n)\b", cand_raw, flags=re.I)[0]
            cand = _clean_patient_name(cand_raw)
            if cand and not _is_bad_patient_value(cand):
                out["paciente"] = cand
                break
        if out.get("paciente"):
            break

    # Demografía y antropometría por patrones tolerantes.
    for key, label in [("edad", "Edad"), ("peso", "Peso"), ("altura", "Altura"), ("imc", "IMC"), ("sc", "SC")]:
        m = re.search(rf"\b{label}\b\s*[:#]?\s*([-+]?\d+(?:[\.,]\d+)?)", flat, re.I)
        if m:
            out[key] = to_float(m.group(1))
    m = re.search(r"\bSexo\b\s*[:#]?\s*([MF])\b", flat, re.I)
    if m:
        out["sexo"] = m.group(1).upper()

    m = re.search(r"\bFecha\b\s*[:#]?\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})", flat, re.I)
    if m:
        out["fecha"] = m.group(1)
    m = re.search(r"\bHora\b\s*[:#]?\s*(\d{1,2}:\d{2}(?::\d{2})?)", flat, re.I)
    if m:
        out["hora"] = m.group(1)
    return out

def _parse_radial_central_table(flat_text, data):
    """Extrae tabla Radial/Central por filas PAS/PAD/PAM/PP/FC. Evita intercambiar PAM con PP."""
    # Patrón más seguro: cada fila con su etiqueta y dos valores.
    def pair(label):
        pat = re.compile(rf"\b{label}\b\s*(?:mmHg|lpm|%)?\s*([-+]?\d+(?:[\.,]\d+)?)\s*(?:\+/-\s*\d+(?:[\.,]\d+)?\s*)?([-+]?\d+(?:[\.,]\d+)?)", re.I)
        m = pat.search(flat_text)
        if m:
            return to_float(m.group(1)), to_float(m.group(2))
        return np.nan, np.nan

    # Recorte preferencial entre Radial/Central y parámetros centrales.
    mblock = re.search(r"Radial\s+Central(.*?)(?:Parámetros\s+hemodinámicos\s+centrales|PAS\s+mmHg|Conclusiones|$)", flat_text, re.I)
    block = mblock.group(1) if mblock else flat_text

    pas_r, pas_c = pair("PAS") if not mblock else pair_from_block(block, "PAS")
    pad_r, pad_c = pair("PAD") if not mblock else pair_from_block(block, "PAD")
    pam_r, pam_c = pair("PAM") if not mblock else pair_from_block(block, "PAM")
    pp_r, pp_c = pair("PP") if not mblock else pair_from_block(block, "PP")

    for k, v in {
        "pas_radial": pas_r, "pas_central": pas_c,
        "pad_radial": pad_r, "pad_central": pad_c,
        "pam_radial": pam_r, "pam_central": pam_c,
        "pp_radial": pp_r, "pp_central": pp_c,
    }.items():
        if not np.isnan(v):
            data[k] = v

    # FC suele ser una sola fila, no radial/central.
    mfc = re.search(r"\bFC\b\s*([-+]?\d+(?:[\.,]\d+)?)", block, re.I)
    if mfc:
        data["fc"] = to_float(mfc.group(1))


def pair_from_block(block, label):
    pat = re.compile(rf"\b{label}\b\s*([-+]?\d+(?:[\.,]\d+)?)\s*([-+]?\d+(?:[\.,]\d+)?)", re.I)
    m = pat.search(block)
    if m:
        return to_float(m.group(1)), to_float(m.group(2))
    return np.nan, np.nan


def _parse_central_parameters(flat_text, data):
    """Extrae parámetros centrales con tolerancia a tablas partidas del PDF.

    Incluye PAS, PP, Au, IAu, RVSE, PE y APC si está disponible.
    La prioridad es la sección 'Parámetros hemodinámicos centrales'.
    """
    msec = re.search(r"Par[aá]metros\s+hemodin[aá]micos\s+centrales(.*?)(?:Conclusiones|Conclusi[oó]n|PAS\s+Presi[oó]n|Pulso|$)", flat_text, re.I)
    sec = msec.group(1) if msec else flat_text
    sec = _collapse_spaces(sec)

    def val(label, aliases=()):
        labels = [label] + list(aliases)
        for lab in labels:
            # Label seguido opcionalmente de unidad y valor con signo. Evita que Au capture IAu usando lookbehind negativo.
            if lab.lower() == "au":
                pat = re.compile(r"(?<!I)\bAu\b\s*(?:mmHg|%)?\s*([-+]?\d+(?:[\.,]\d+)?)", re.I)
            else:
                pat = re.compile(rf"\b{lab}\b\s*(?:mmHg|%|lpm|m/s)?\s*([-+]?\d+(?:[\.,]\d+)?)", re.I)
            mm = pat.search(sec)
            if mm:
                return to_float(mm.group(1))
        return np.nan

    mapping = {
        "pas_central": val("PAS"),
        "pp_central": val("PP"),
        "au": val("Au", aliases=[r"Aumentaci[oó]n\s+A[oó]rtica\s+Central"]),
        "iau": val("IAu", aliases=[r"Indice\s+de\s+Aumentaci[oó]n\s+Central", r"Índice\s+de\s+Aumentaci[oó]n\s+Central"]),
        "rvse": val("RVSE", aliases=[r"Relaci[oó]n\s+de\s+Viabilidad\s+Sub\s*Endoc[aá]rdica"]),
        "pe": val("PE", aliases=[r"Periodo\s+Eyectivo", r"Per[ií]odo\s+Eyectivo"]),
        "apc": val("APC", aliases=[r"Amplificaci[oó]n\s+Perif[eé]rico\s*Central"]),
    }

    # Respaldo por patrón posicional de la tabla central en la captura:
    # PAS mmHg 119 +/-1  PP mmHg 31 +/-2  Au mmHg +8 +/-1  IAu % +25 +/-3  RVSE % 180 +/-3  PE % 32.8
    positional = re.findall(r"\b(PAS|PP|Au|IAu|RVSE|PE|APC)\b\s*(?:mmHg|%)?\s*([-+]?\d+(?:[\.,]\d+)?)", sec, flags=re.I)
    for lab, num in positional:
        key = {"pas":"pas_central", "pp":"pp_central", "au":"au", "iau":"iau", "rvse":"rvse", "pe":"pe", "apc":"apc"}.get(lab.lower())
        if key and np.isnan(to_float(mapping.get(key, np.nan))):
            mapping[key] = to_float(num)

    for k, v in mapping.items():
        if not np.isnan(v):
            data[k] = v


def _is_bad_patient_value(v):
    """Detecta valores que son etiquetas administrativas y no nombres de paciente."""
    s = _collapse_spaces(v)
    if not s:
        return True
    if re.search(r"(?i)\b(Número\s+de\s+estudio|Numero\s+de\s+estudio|Estudio\s*#?|Fecha|Hora|Edad|Sexo|Peso|Altura|IMC|SC|Diagn[oó]stico|Medicaci[oó]n|Radial|Central|Realizado|posici[oó]n|en\s+posici[oó]n)\b", s):
        return True
    if len(s) < 3 or re.fullmatch(r"[\d\s:;#\-/.]+", s):
        return True
    return False


def _parse_header_fields_from_lines(lines):
    """Extrae cabecera paciente/estudio/demografía por líneas, más confiable que texto plano.

    En los PDF tipo Exxer, pdfplumber puede mezclar columnas y hacer que "Paciente"
    tome "Número de estudio" como valor. Este extractor prioriza la estructura visual
    por líneas: Estudio # 684 / Paciente APELLIDO NOMBRE / Edad 54 / Sexo M, etc.
    """
    out = {}
    clean_lines = [_collapse_spaces(x) for x in lines if _collapse_spaces(x)]

    def take_num(label, key, pattern_suffix=r""):
        for ln in clean_lines:
            m = re.search(rf"\b{label}\b\s*[:#]?\s*([-+]?\d+(?:[\.,]\d+)?)\s*{pattern_suffix}", ln, re.I)
            if m:
                out[key] = to_float(m.group(1)); return

    # Estudio: solo aceptar número/código inmediato luego de Estudio #, no letras sueltas como M.
    for ln in clean_lines:
        m = re.search(r"\bEstudio\s*#?\s*[:#]?\s*([A-Za-z0-9][A-Za-z0-9_\-/]{0,30})", ln, re.I)
        if m:
            val = _collapse_spaces(m.group(1)).strip(":#- ")
            val = re.split(r"\bH\.?\s*C\.?\b|#|Paciente|Fecha|Hora|Edad|Sexo", val, flags=re.I)[0].strip(" :#-")
            if val and not re.fullmatch(r"(?i)(M|F|Paciente|Fecha|Hora|Edad|Sexo)", val):
                out["estudio"] = val
                break

    # Paciente: preferir línea que empieza o contiene Paciente + nombre.
    for i, ln in enumerate(clean_lines):
        m = re.search(r"\bPaciente\b\s*[:#]?\s*(.+)$", ln, re.I)
        if m:
            cand = m.group(1)
            cand = re.split(r"\b(Estudio\s*#?|Número\s+de\s+estudio|Numero\s+de\s+estudio|Fecha|Hora|Edad|Sexo|Peso|Altura|IMC|SC|Diagn[oó]stico|Medicaci[oó]n)\b", cand, flags=re.I)[0]
            cand = _clean_patient_name(cand)
            if cand and not _is_bad_patient_value(cand):
                out["paciente"] = cand
                break
            # Si el valor no está en la misma línea, mirar la línea siguiente.
            if i + 1 < len(clean_lines):
                nxt = _clean_patient_name(clean_lines[i+1])
                if nxt and not _is_bad_patient_value(nxt):
                    out["paciente"] = nxt
                    break

    # Fecha/hora
    for ln in clean_lines:
        m = re.search(r"\bFecha\b\s*[:#]?\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})", ln, re.I)
        if m:
            out["fecha"] = m.group(1); break
    for ln in clean_lines:
        m = re.search(r"\bHora\b\s*[:#]?\s*(\d{1,2}:\d{2}(?::\d{2})?)", ln, re.I)
        if m:
            out["hora"] = m.group(1); break

    take_num("Edad", "edad")
    take_num("Peso", "peso", r"(?:kg|kilos)?")
    take_num("Altura", "altura", r"(?:cm)?")
    take_num("IMC", "imc")
    take_num("SC", "sc", r"(?:m2|m²)?")

    for ln in clean_lines:
        m = re.search(r"\bSexo\b\s*[:#]?\s*([MF])\b", ln, re.I)
        if m:
            out["sexo"] = m.group(1).upper(); break

    return out

def _validate_and_repair_pac_data(data):
    """Repara valores mal cargados por columnas pegadas del PDF y limpia etiquetas contaminantes."""
    cand_patient = _patient_name_strict(data.get("paciente", ""))
    data["paciente"] = cand_patient if cand_patient else ""
    if _is_bad_patient_value(data.get("paciente", "")) or re.search(r"(?i)\b(en\s+posici[oó]n|posici[oó]n|realizado)\b", str(data.get("paciente", ""))):
        data["paciente"] = ""
    # Evitar que el número de estudio quede como sexo u otra etiqueta.
    if re.fullmatch(r"(?i)(M|F)", _collapse_spaces(data.get("estudio", ""))):
        data["estudio"] = ""
    data["sexo"] = _normalize_sex(data.get("sexo", ""))

    # Si falta PAD/PAM/PP central y hay periféricos, no inventar; solo completar PP si PAS/PAD central son reales.
    pas_c, pad_c = to_float(data.get("pas_central")), to_float(data.get("pad_central"))
    if (np.isnan(data.get("pp_central", np.nan)) or data.get("pp_central", 0) == 0) and not np.isnan(pas_c) and not np.isnan(pad_c) and pas_c > pad_c:
        data["pp_central"] = pas_c - pad_c
    pas_r, pad_r = to_float(data.get("pas_radial")), to_float(data.get("pad_radial"))
    if (np.isnan(data.get("pp_radial", np.nan)) or data.get("pp_radial", 0) == 0) and not np.isnan(pas_r) and not np.isnan(pad_r) and pas_r > pad_r:
        data["pp_radial"] = pas_r - pad_r

    # Rango fisiológico/administrativo: si falla, dejar editable como vacío en la interfaz.
    for k, lo, hi in [
        ("edad", 1, 120), ("peso", 20, 250), ("altura", 80, 230), ("imc", 10, 80), ("sc", 0.5, 3.5),
        ("pas_radial", 50, 260), ("pad_radial", 30, 160), ("pam_radial", 40, 200), ("pp_radial", 10, 120),
        ("pas_central", 50, 260), ("pad_central", 30, 160), ("pam_central", 40, 200), ("pp_central", 10, 120),
        ("fc", 25, 180), ("au", -80, 100), ("iau", -80, 100), ("rvse", 0, 300), ("pe", 10, 60), ("apc", 0.2, 3.5)
    ]:
        v = to_float(data.get(k))
        if np.isnan(v) or not (lo <= v <= hi):
            data[k] = np.nan
        else:
            data[k] = v
    return data


def parse_model_pac(text):
    """Parser robusto para PDF PAC/Exxer.

    Corrige errores frecuentes de pdfplumber: etiquetas pegadas al valor, paciente cargado como
    'Número de estudio', sexo dentro de estudio, altura/IMC mezclados y PAM/PP intercambiados.
    """
    lines = [_collapse_spaces(x) for x in text.splitlines() if _collapse_spaces(x)]
    joined = "\n".join(lines)
    flat = _collapse_spaces(" ".join(lines))

    data = {
        "paciente": _clean_patient_name(_value_after_label(flat, [r"Paciente", r"Nombre\s+del\s+paciente"], max_chars=90)),
        "estudio": _value_after_label(flat, [r"Estudio\s*#", r"Número\s+de\s+estudio"], max_chars=30),
        "fecha": _value_after_label(flat, [r"Fecha"], max_chars=25),
        "hora": _value_after_label(flat, [r"Hora"], max_chars=20),
        "hc": _value_after_label(flat, [r"H\.C\."], max_chars=30),
        "diagnostico_previo": _value_after_label(flat, [r"Diagnóstico", r"Diagnostico"], max_chars=80),
        "medicacion": _value_after_label(flat, [r"Medicación", r"Medicacion"], max_chars=80),
        "edad": _number_after_label(flat, [r"Edad"]),
        "sexo": _normalize_sex(_value_after_label(flat, [r"Sexo"], max_chars=20)),
        "peso": _number_after_label(flat, [r"Peso"]),
        "altura": _number_after_label(flat, [r"Altura"]),
        "imc": _number_after_label(flat, [r"IMC"]),
        "sc": _number_after_label(flat, [r"SC"]),
        "fc": _number_after_label(flat, [r"FC"]),
    }

    # Corrección prioritaria por líneas: evita que "Paciente" quede como "Número de estudio"
    # y que "Estudio" tome "M" u otra etiqueta de la cabecera.
    header_by_lines = _parse_header_fields_from_lines(lines)
    header_by_regex = _extract_header_by_regex(flat)
    for source in (header_by_lines, header_by_regex):
        for k, v in source.items():
            if k in ["paciente", "estudio", "fecha", "hora", "sexo"]:
                if v not in [None, ""]:
                    data[k] = v
            else:
                if not np.isnan(to_float(v)):
                    data[k] = v

    _parse_radial_central_table(flat, data)
    _parse_central_parameters(flat, data)

    # Respaldo para el patrón compacto antiguo: Radial Central 127 119 87 88 102 102 40 31.
    if any(np.isnan(to_float(data.get(k))) for k in ["pas_radial","pas_central","pad_radial","pad_central","pam_radial","pam_central","pp_radial","pp_central"]):
        m = re.search(r"Radial\s+Central\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)", flat, re.I)
        if m:
            vals = list(map(float, m.groups()))
            data.update({
                "pas_radial": vals[0], "pas_central": vals[1],
                "pad_radial": vals[2], "pad_central": vals[3],
                "pam_radial": vals[4], "pam_central": vals[5],
                "pp_radial": vals[6], "pp_central": vals[7],
            })

    for k in ["pas_radial","pad_radial","pam_radial","pp_radial","pas_central","pad_central","pam_central","pp_central","au","iau","rvse","pe","apc"]:
        data.setdefault(k, np.nan)

    return _validate_and_repair_pac_data(data)

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




def _curve_fingerprint(t, p):
    """Firma numérica corta de la morfología real para verificar que cada paciente use su propia curva."""
    try:
        t = np.asarray(t, dtype=float)
        p = np.asarray(p, dtype=float)
        ok = np.isfinite(t) & np.isfinite(p)
        t, p = t[ok], p[ok]
        if len(p) < 20:
            return "curva_insuficiente"
        x = np.linspace(0, 1, 48)
        tn = (t - np.nanmin(t)) / max(np.nanmax(t) - np.nanmin(t), 1e-6)
        pn = (p - np.nanmin(p)) / max(np.nanmax(p) - np.nanmin(p), 1e-6)
        sig = np.interp(x, tn, pn)
        sig = np.round(sig, 3)
        return str(abs(hash(tuple(sig))) % 10_000_000).zfill(7)
    except Exception:
        return "sin_firma"


def _first_true_index(mask, default=0):
    idx = np.where(mask)[0]
    return int(idx[0]) if len(idx) else int(default)


def _last_true_index(mask, default=0):
    idx = np.where(mask)[0]
    return int(idx[-1]) if len(idx) else int(default)


def _regularize_real_pressure_curve(wave_df, row, n=720):
    """Regulariza la curva real importada sin inventar morfología.

    La única calibración permitida es lineal contra PAS/PAD central reales del estudio.
    Si no hay suficientes puntos reales, la función falla y bloquea el informe.
    """
    df = wave_df.copy()
    t = pd.to_numeric(df.iloc[:, 0], errors="coerce").to_numpy(dtype=float)
    p = pd.to_numeric(df.iloc[:, 1], errors="coerce").to_numpy(dtype=float)
    ok = np.isfinite(t) & np.isfinite(p)
    t, p = t[ok], p[ok]
    if len(p) < 40:
        raise ValueError("Curva real insuficiente: se requieren al menos 40 puntos válidos del paciente.")

    order = np.argsort(t)
    t, p = t[order], p[order]
    keep = np.r_[True, np.diff(t) > 0]
    t, p = t[keep], p[keep]
    if len(p) < 40:
        raise ValueError("Curva real insuficiente luego de quitar tiempos duplicados.")

    t_norm = (t - np.nanmin(t)) / max(np.nanmax(t) - np.nanmin(t), 1e-6) * 1000.0
    t0 = np.linspace(0, 1000, n)
    p0 = np.interp(t0, t_norm, p)

    # Suavizado mínimo, proporcional a la cantidad de puntos, para no borrar morfología individual.
    win = max(5, min(17, (n // 90) * 2 + 1))
    p0 = pd.Series(p0).rolling(win, center=True, min_periods=1).median().to_numpy()
    p0 = pd.Series(p0).rolling(win, center=True, min_periods=1).mean().to_numpy()

    pas = to_float(row.get("pas_central"))
    pad = to_float(row.get("pad_central"))
    raw_min, raw_max = float(np.nanmin(p0)), float(np.nanmax(p0))
    if raw_max - raw_min < 5:
        raise ValueError("La curva real tiene amplitud insuficiente; se bloquea análisis para evitar curvas genéricas.")

    if not np.isnan(pas) and not np.isnan(pad) and pas > pad:
        p0 = pad + (p0 - raw_min) * (pas - pad) / max(raw_max - raw_min, 1e-6)
    else:
        pad, pas = raw_min, raw_max

    return t0, p0, _curve_fingerprint(t0, p0)


def estimate_wave_separation(wave_df, row):
    """Separación morfológica Pf/Pb derivada de la curva real del paciente.

    Modo estricto:
    - No usa curva sintética.
    - No usa plantilla fija entre pacientes.
    - Pf + Pb reconstruyen la onda real sobre PAD.
    - El flujo aórtico se deriva de la pendiente positiva real de la curva, no de una onda triangular.
    """
    t0, p0, curve_id = _regularize_real_pressure_curve(wave_df, row, n=720)

    pad = float(np.nanmin(p0))
    pas = float(np.nanmax(p0))
    pp = max(pas - pad, 1.0)
    excess = np.clip(p0 - pad, 0, None)

    peak_i = int(np.nanargmax(p0))
    t_peak = float(t0[peak_i])

    # Landmarks reales por derivadas: pie sistólico, punto de máxima pendiente y hombro/retorno reflejo.
    dt = float(np.nanmedian(np.diff(t0)))
    dp = np.gradient(p0, dt)
    d2 = np.gradient(dp, dt)

    pre = np.arange(0, max(peak_i, 1))
    foot_i = _first_true_index(excess[:max(peak_i, 1)] > 0.04 * pp, default=0)
    max_dp_i = int(pre[np.nanargmax(dp[pre])]) if len(pre) else max(0, peak_i // 2)

    post_start = min(len(t0)-1, peak_i + max(8, int(25/dt)))
    post_end = min(len(t0)-1, peak_i + max(45, int(380/dt)))
    post = np.arange(post_start, post_end)
    if len(post) > 10:
        # El retorno reflejo se ubica en la mayor convexidad/meseta pospico real, no en un tiempo fijo.
        curv_score = d2[post] - 0.20 * np.abs(dp[post])
        refl_i = int(post[np.nanargmax(curv_score)])
    else:
        refl_i = min(len(t0)-1, peak_i + int(140/dt))

    # Si la curva tiene hombro tardío visible, priorizar máximo local tardío.
    local_max = []
    for i in range(post_start + 1, post_end - 1):
        if p0[i-1] <= p0[i] >= p0[i+1] and t0[i] > t_peak + 45:
            local_max.append(i)
    if local_max:
        refl_i = int(local_max[0])

    t_ref = float(t0[refl_i])
    transition = float(np.clip(28 + 0.10 * max(t_ref - t_peak, 40), 28, 85))

    # Pesos temporales derivados de la curva real: antes de t_ref domina Pf; después domina Pb.
    late_weight = 1.0 / (1.0 + np.exp(-(t0 - t_ref) / transition))
    early_weight = 1.0 - late_weight

    # La reflexión no aparece antes del pie sistólico ni domina al inicio.
    early_weight[t0 < t0[foot_i]] = 1.0
    late_weight[t0 < t0[foot_i]] = 0.0

    # Reparto estricto del exceso real: no se inventa presión. Pf + Pb = presión real - PAD.
    pf = excess * early_weight
    pb = excess * late_weight

    # Refuerzo de hombro reflejado real: si hay convexidad/segunda meseta, se asigna a Pb con forma tomada del exceso real.
    shoulder_window = np.exp(-((t0 - t_ref) / max(transition * 2.2, 80)) ** 2)
    pb = np.maximum(pb, excess * shoulder_window * np.clip(0.35 + 0.45 * (excess / max(pp, 1)), 0.20, 0.80))
    # Rebalancear para que no exceda la presión real.
    total = pf + pb
    over = total > excess
    scale = np.ones_like(total)
    scale[over] = excess[over] / np.maximum(total[over], 1e-6)
    pf *= scale
    pb *= scale

    pf_abs = pad + pf
    pb_abs = pad + pb
    p_model = pad + pf + pb  # exactamente la curva real suavizada/calibrada

    pf_peak = float(np.nanmax(pf))
    pb_peak = float(np.nanmax(pb))
    tfor = float(t0[int(np.nanargmax(pf))]) if pf_peak > 0 else np.nan
    tref_m = float(t0[int(np.nanargmax(pb))]) if pb_peak > 0 else np.nan
    rm = pb_peak / pf_peak if pf_peak > 0 else np.nan
    ri = pb_peak / (pf_peak + pb_peak) if (pf_peak + pb_peak) > 0 else np.nan
    t_ratio = tfor / tref_m if not np.isnan(tref_m) and tref_m > 0 else np.nan

    # Flujo aórtico del paciente: derivado de la pendiente positiva real, redondeado, sin triángulo fijo.
    fc = to_float(row.get("fc"))
    pe_pct = to_float(row.get("pe"))
    cycle_ms = 60000.0 / fc if not np.isnan(fc) and fc > 20 else 1000.0
    if not np.isnan(pe_pct) and pe_pct > 10:
        ej_duration = float(np.clip(cycle_ms * pe_pct / 100.0, 190, 520))
    else:
        # Fin de eyección estimado por caída pospico al 35% de PP o por mínimo de dp/dt.
        post_peak = np.arange(peak_i, len(t0))
        fall_idx = post_peak[excess[post_peak] < 0.35 * pp]
        ej_duration = float(np.clip((t0[fall_idx[0]] - t0[foot_i]) if len(fall_idx) else 330, 210, 520))
    ej_start = float(t0[foot_i])
    ej_end = float(np.clip(ej_start + ej_duration, t_peak + 80, min(1000, ej_start + 540)))

    q = np.zeros_like(t0)
    eject = (t0 >= ej_start) & (t0 <= ej_end)
    positive_dp = np.clip(dp, 0, None)
    # Se agrega pequeña contribución por presión excedente durante eyección para producir pico romo fisiológico.
    q[eject] = positive_dp[eject] + 0.012 * excess[eject]
    q = pd.Series(q).rolling(21, center=True, min_periods=1).mean().to_numpy()
    q = pd.Series(q).rolling(21, center=True, min_periods=1).mean().to_numpy()
    if np.nanmax(q) > 0:
        # Escala orientativa individual por PP y FC; la forma sigue siendo de la curva real.
        qp = np.clip(190 + 3.0 * pp + (0 if np.isnan(fc) else 0.65 * fc), 160, 560)
        q = q / np.nanmax(q) * qp

    sep_df = pd.DataFrame({
        "tiempo_ms": t0,
        "presion_total_mmHg": p0,
        "presion_modelo_pf_pb_mmHg": p_model,
        "onda_anterograda_pf": pf,
        "onda_retrograda_pb": pb,
        "onda_anterograda_pf_abs": pf_abs,
        "onda_retrograda_pb_abs": pb_abs,
        "flujo_aortico_estimado_ml_s": q,
    })

    # Métricas de morfología real para auditar que no se repita la misma curva.
    syst_mask = (t0 >= ej_start) & (t0 <= ej_end)
    diast_mask = (t0 > ej_end) & (t0 <= 1000)

    # Áreas de presión-tiempo sobre PAD: útiles para morfología de la onda.
    systolic_area = float(safe_trapezoid(excess[syst_mask], t0[syst_mask]))
    total_area = float(safe_trapezoid(excess, t0))

    # RVSE/SEVR operativo: relación de área diastólica/sistólica de la curva central real.
    # Se calcula con la presión central absoluta porque representa una aproximación presión-tiempo
    # del balance oferta/demanda subendocárdica. No se calcula desde valores aislados.
    systolic_pti = float(safe_trapezoid(p0[syst_mask], t0[syst_mask]))
    diastolic_pti = float(safe_trapezoid(p0[diast_mask], t0[diast_mask]))
    rvse_calc = (diastolic_pti / systolic_pti * 100.0) if systolic_pti > 0 and np.sum(diast_mask) >= 2 else np.nan
    rvse_pdf = to_float(row.get("rvse"))
    rvse_delta = rvse_calc - rvse_pdf if not np.isnan(rvse_calc) and not np.isnan(rvse_pdf) else np.nan

    ai_morph = float((p0[refl_i] - p0[max_dp_i]) / pp * 100.0) if pp > 0 else np.nan

    metrics = {
        "pf_pico": pf_peak,
        "pb_pico": pb_peak,
        "tfor_ms": tfor,
        "tref_ms": tref_m,
        "rm": rm,
        "ri": ri,
        "tfor_tref": t_ratio,
        "qp_ml_s": float(np.nanmax(q)) if len(q) else np.nan,
        "pe_ms": float(ej_end - ej_start),
        "curve_id": curve_id,
        "t_pico_ms": t_peak,
        "t_pie_ms": float(t0[foot_i]),
        "t_max_dpdt_ms": float(t0[max_dp_i]),
        "area_sistolica": systolic_area,
        "area_total": total_area,
        "area_sistolica_pti": systolic_pti,
        "area_diastolica_pti": diastolic_pti,
        "rvse_calculado_%": rvse_calc,
        "rvse_importado_%": rvse_pdf,
        "rvse_delta_%": rvse_delta,
        "ai_morfologico_%": ai_morph,
        "t_ej_inicio_ms": ej_start,
        "t_ej_fin_ms": ej_end,
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


def plot_rvse_area(sep_df, sep_metrics):
    """Gráfico de áreas presión-tiempo para RVSE/SEVR calculado desde la curva real."""
    fig, ax = plt.subplots(figsize=(7.6, 4.2))
    t = pd.to_numeric(sep_df["tiempo_ms"], errors="coerce").to_numpy(dtype=float)
    p = pd.to_numeric(sep_df["presion_total_mmHg"], errors="coerce").to_numpy(dtype=float)
    ej_start = float(sep_metrics.get("t_ej_inicio_ms", np.nan))
    ej_end = float(sep_metrics.get("t_ej_fin_ms", np.nan))
    ok = np.isfinite(t) & np.isfinite(p)
    t, p = t[ok], p[ok]
    ax.plot(t, p, color="#111111", linewidth=2.2, label="Presión central real")
    if len(t) and not np.isnan(ej_start) and not np.isnan(ej_end):
        syst = (t >= ej_start) & (t <= ej_end)
        diast = t > ej_end
        ax.fill_between(t[syst], 0, p[syst], color="#C62828", alpha=0.13, label="Área sistólica")
        ax.fill_between(t[diast], 0, p[diast], color="#1565C0", alpha=0.12, label="Área diastólica")
        ax.axvline(ej_start, color="#78909C", linewidth=0.9, linestyle=":")
        ax.axvline(ej_end, color="#78909C", linewidth=0.9, linestyle="--")
    rvse_calc = sep_metrics.get("rvse_calculado_%", np.nan)
    title = "RVSE / SEVR por áreas presión-tiempo"
    if not np.isnan(rvse_calc):
        title += f" = {rvse_calc:.1f}%"
    _apply_professional_axes(ax, title, "Tiempo (ms)", "Presión central (mmHg)")
    ax.set_ylim(bottom=0)
    ax.legend(fontsize=8, loc="upper right", frameon=True, facecolor="white", edgecolor="#CFD8DC")
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


def interpret_rvse_profile(row, sep_metrics):
    """Interpreta RVSE/SEVR calculado desde la curva central real y lo compara con el valor importado."""
    rvse_calc = sep_metrics.get("rvse_calculado_%", np.nan)
    rvse_imp = to_float(row.get("rvse"))
    syst_pti = sep_metrics.get("area_sistolica_pti", np.nan)
    diast_pti = sep_metrics.get("area_diastolica_pti", np.nan)

    if np.isnan(rvse_calc):
        return "RVSE/SEVR: no calculable porque no se pudo definir en forma estable el área sistólica y diastólica de la curva central real."

    if rvse_calc < 120:
        grade = "reducido"
        meaning = "sugiere menor reserva subendocárdica relativa o mayor demanda sistólica respecto del tiempo diastólico disponible"
    elif rvse_calc < 150:
        grade = "limítrofe/intermedio"
        meaning = "sugiere balance subendocárdico intermedio, dependiente de la frecuencia cardíaca, presión diastólica central y poscarga sistólica"
    else:
        grade = "conservado"
        meaning = "sugiere balance presión-tiempo diastólico/sistólico favorable en la curva analizada"

    comparison = ""
    if not np.isnan(rvse_imp) and rvse_imp > 0:
        delta = rvse_calc - rvse_imp
        comparison = f" El RVSE informado por el equipo es {rvse_imp:.1f}% y el RVSE recalculado por la app es {rvse_calc:.1f}% (diferencia {delta:+.1f} puntos), útil como control de consistencia de la digitalización."

    return (
        f"RVSE/SEVR calculado desde la curva central real: {rvse_calc:.1f}%. "
        f"Área sistólica presión-tiempo: {format_optional(syst_pti, 0)} mmHg·ms; "
        f"área diastólica presión-tiempo: {format_optional(diast_pti, 0)} mmHg·ms. "
        f"Interpretación: RVSE {grade}, {meaning}.{comparison}"
    )


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

    rvse_interp = interpret_rvse_profile(row, sep_metrics)
    c1 = interpret_pressure_central_metrics(row, dx, cat, ref, amp_sbp, ppa, risk) + " " + rvse_interp

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
    rvse_calc = sep_metrics.get("rvse_calculado_%", np.nan)
    if not np.isnan(rvse_calc) and rvse_calc < 120:
        integrated_flags.append("RVSE reducido por análisis de área presión-tiempo")
    elif not np.isnan(rvse_calc) and rvse_calc >= 150:
        integrated_flags.append("RVSE conservado")
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
        ("4. Fenotipo final de presión central", c4),
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

    rvse_calc = sep_metrics.get("rvse_calculado_%", np.nan)
    if not np.isnan(rvse_calc):
        if rvse_calc < 120:
            reasons_pressure.append(f"RVSE calculado reducido ({rvse_calc:.1f}%)")
            pressure_score += 1
        elif rvse_calc >= 150:
            reasons_pressure.append(f"RVSE calculado conservado ({rvse_calc:.1f}%)")
        else:
            reasons_pressure.append(f"RVSE calculado intermedio ({rvse_calc:.1f}%)")
        total_score = pressure_score + wave_score + harmonic_score

    table = [
        ["Dominio", "Puntaje", "Elementos considerados"],
        ["Presión central, métricas y RVSE", str(pressure_score), "; ".join(reasons_pressure) if reasons_pressure else "sin datos suficientes"],
        ["Separación de ondas Pf/Pb", str(wave_score), "; ".join(reasons_wave) if reasons_wave else "sin datos suficientes"],
        ["Armónicos", str(harmonic_score), "; ".join(reasons_harm) if reasons_harm else "sin datos suficientes"],
        ["Puntaje integrado", str(total_score), phenotype],
    ]

    text = (
        f"Fenotipo final: {phenotype}. La integración de presión central, separación de ondas y análisis armónico indica que {clinical} "
        f"Valores integrados: PAS central {pas_c:.0f} mmHg si disponible, PAD central {pad_c:.0f} mmHg si disponible, "
        f"PP central {pp_c:.0f} mmHg si disponible, IAu {iau:.1f}% si disponible, PPA {ppa:.2f} si disponible, "
        f"RVSE calculado {rvse_calc:.1f}% si disponible, RM {rm:.2f}, RI {ri:.2f}, Tref {tref:.0f} ms, Pf pico {pf:.1f} mmHg y Pb pico {pb:.1f} mmHg. "
        "Este fenotipo es una clasificación operativa de apoyo y debe integrarse con edad, sexo, presión braquial, tratamiento, riesgo cardiovascular y lesión de órgano blanco."
    )
    text = text.replace("nan mmHg", "no disponible").replace("nan%", "no disponible").replace("nan", "no disponible")
    return phenotype, text, table

def build_pdf(row, wave_df, hdf, screenshot_png=None):
    """Construye un PDF compacto, con conclusiones primero y luego grilla gráfica profesional."""
    dx, cat, ref, amp_sbp, ppa, risk = central_diagnosis(row)
    conclusion_blocks, sep_df, sep_metrics, sep_interp = build_continuous_conclusions(row, wave_df, hdf)
    final_phenotype, final_phenotype_text, final_phenotype_table = classify_central_pressure_phenotype(row, sep_metrics, hdf)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        rightMargin=11*mm, leftMargin=11*mm,
        topMargin=21*mm, bottomMargin=13*mm
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="SmallPAC", parent=styles["BodyText"], fontName="Helvetica",
        fontSize=7.2, leading=8.5, textColor=colors.HexColor("#263238")
    ))
    styles.add(ParagraphStyle(
        name="BodyPAC", parent=styles["BodyText"], fontName="Helvetica",
        fontSize=8.3, leading=10.1, textColor=colors.HexColor("#1F2D3D"), spaceAfter=2
    ))
    styles.add(ParagraphStyle(
        name="ConclusionPAC", parent=styles["BodyText"], fontName="Helvetica",
        fontSize=8.15, leading=9.65, textColor=colors.HexColor("#1F2D3D"), spaceAfter=2
    ))
    styles.add(ParagraphStyle(
        name="TitlePAC", parent=styles["Title"], fontName="Helvetica-Bold",
        fontSize=14.7, leading=17.2, alignment=1, textColor=colors.HexColor("#12355B"), spaceAfter=2
    ))
    styles.add(ParagraphStyle(
        name="SectionPAC", parent=styles["Heading2"], fontName="Helvetica-Bold",
        fontSize=9.8, leading=11.5, textColor=colors.white, spaceAfter=0
    ))
    styles.add(ParagraphStyle(
        name="H3PAC", parent=styles["Heading3"], fontName="Helvetica-Bold",
        fontSize=8.15, leading=9.4, textColor=colors.HexColor("#17365D"), spaceAfter=1
    ))
    styles.add(ParagraphStyle(
        name="MiniTitlePAC", parent=styles["Heading3"], fontName="Helvetica-Bold",
        fontSize=8.25, leading=9.4, textColor=colors.HexColor("#12355B"), spaceBefore=1, spaceAfter=0
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
        return Table([[Paragraph(title, styles["SectionPAC"])]], colWidths=[188*mm], style=TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#17365D")),
            ("BOX", (0,0), (-1,-1), 0.35, colors.HexColor("#17365D")),
            ("LEFTPADDING", (0,0), (-1,-1), 5),
            ("RIGHTPADDING", (0,0), (-1,-1), 5),
            ("TOPPADDING", (0,0), (-1,-1), 3),
            ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ]))

    def _table_style(header_color="#D9EAF7", font_size=7.2):
        return TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor(header_color)),
            ("TEXTCOLOR", (0,0), (-1,0), colors.HexColor("#17365D")),
            ("FONT", (0,0), (-1,0), "Helvetica-Bold", font_size),
            ("FONT", (0,1), (-1,-1), "Helvetica", font_size),
            ("GRID", (0,0), (-1,-1), 0.22, colors.HexColor("#B0BEC5")),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F7FAFC")]),
            ("LEFTPADDING", (0,0), (-1,-1), 3),
            ("RIGHTPADDING", (0,0), (-1,-1), 3),
            ("TOPPADDING", (0,0), (-1,-1), 2),
            ("BOTTOMPADDING", (0,0), (-1,-1), 2),
        ])

    def _header_footer(canvas, doc_obj):
        canvas.saveState()
        width, height = A4
        canvas.setFillColor(colors.HexColor("#12355B"))
        canvas.rect(0, height-15*mm, width, 15*mm, stroke=0, fill=1)
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica-Bold", 8.8)
        canvas.drawString(11*mm, height-9*mm, "PAC IA | Presión Aórtica Central")
        canvas.setFont("Helvetica", 7.2)
        canvas.drawRightString(width-11*mm, height-9*mm, datetime.now().strftime("%d/%m/%Y"))
        canvas.setFillColor(colors.HexColor("#607D8B"))
        canvas.setFont("Helvetica", 6.8)
        canvas.drawString(11*mm, 7*mm, "Informe médico integrado - diseño compacto profesional")
        canvas.drawRightString(width-11*mm, 7*mm, f"Página {doc_obj.page}")
        canvas.restoreState()

    def _graph_cell(title, img, width=91*mm, height=50*mm):
        return [Paragraph(title, styles["H3PAC"]), Image(img, width=width, height=height)]

    story = []
    story.append(Paragraph("PRESIÓN AÓRTICA CENTRAL", styles["TitlePAC"]))
    story.append(Paragraph("Informe médico integrado con conclusiones clínicas continuas y panel gráfico compacto", styles["BodyPAC"]))
    story.append(Spacer(1, 1.7*mm))

    story.append(_section("1. Datos del paciente y valores principales"))
    datos = [
        ["Paciente", safe_text(row.get("paciente","")), "Estudio", safe_text(row.get("estudio",""))],
        ["Fecha", safe_text(row.get("fecha","")), "Hora", safe_text(row.get("hora",""))],
        ["Edad", _fmt(row.get("edad",""),0), "Sexo", safe_text(row.get("sexo",""))],
        ["Peso", _fmt(row.get("peso",""),1), "Altura", _fmt(row.get("altura",""),1)],
        ["IMC", _fmt(row.get("imc",""),1), "Medicación", safe_text(row.get("medicacion",""))],
    ]
    vals = [["Variable", "Radial/Braquial", "Central", "Unidad"],
            ["PAS", _fmt(row.get("pas_radial")), _fmt(row.get("pas_central")), "mmHg"],
            ["PAD", _fmt(row.get("pad_radial")), _fmt(row.get("pad_central")), "mmHg"],
            ["PAM", _fmt(row.get("pam_radial")), _fmt(row.get("pam_central")), "mmHg"],
            ["PP", _fmt(row.get("pp_radial")), _fmt(row.get("pp_central")), "mmHg"],
            ["FC", _fmt(row.get("fc"),0), "", "lpm"],
            ["Au", "", _fmt(row.get("au")), "mmHg"],
            ["IAu", "", _fmt(row.get("iau")), "%"],
            ["RVSE equipo", "", _fmt(row.get("rvse")), "%"],
            ["RVSE calculado", "", _fmt(sep_metrics.get("rvse_calculado_%")), "%"],
            ["PE", "", _fmt(row.get("pe")), "%"],
            ["APC", "", _fmt(row.get("apc")), "relación"]]
    patient_table = Table(datos, colWidths=[22*mm, 54*mm, 22*mm, 42*mm], style=_table_style("#EAF2F8", 6.9))
    values_table = Table(vals, colWidths=[20*mm, 24*mm, 23*mm, 16*mm], style=_table_style("#D9EAF7", 6.9))
    story.append(Table([[patient_table, values_table]], colWidths=[101*mm, 87*mm], style=TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 0),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
        ("TOPPADDING", (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ])))
    story.append(Spacer(1, 1.8*mm))

    story.append(_section("2. Conclusiones clínicas continuas"))
    conclusion_rows = []
    for title, body in conclusion_blocks:
        conclusion_rows.append([Paragraph(title, styles["MiniTitlePAC"])])
        conclusion_rows.append([Paragraph(body, styles["ConclusionPAC"])])
    story.append(Table(conclusion_rows, colWidths=[188*mm], style=TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#F7FAFC")),
        ("BOX", (0,0), (-1,-1), 0.4, colors.HexColor("#90A4AE")),
        ("INNERGRID", (0,0), (-1,-1), 0.12, colors.HexColor("#ECEFF1")),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 5),
        ("RIGHTPADDING", (0,0), (-1,-1), 5),
        ("TOPPADDING", (0,0), (-1,-1), 2.4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2.4),
    ])))

    # Sin salto forzado: ReportLab decide el pase de página y evita blancos grandes.
    story.append(Spacer(1, 2.5*mm))
    story.append(_section("3. Gráficos del informe"))
    story.append(Spacer(1, 1.5*mm))
    story.append(KeepTogether([
        Paragraph("Presión aórtica central con ondas Pf/Pb superpuestas", styles["H3PAC"]),
        Image(plot_wave_separation(sep_df), width=188*mm, height=88*mm)
    ]))
    story.append(Spacer(1, 1.5*mm))

    img_w = 91*mm
    img_h = 47*mm
    graph_table = Table([
        [_graph_cell("Presiones periféricas vs centrales", plot_pressure_comparison(row), img_w, img_h),
         _graph_cell("Onda de presión aórtica central", plot_waveform(wave_df), img_w, img_h)],
        [_graph_cell("Flujo aórtico estimado", plot_aortic_flow(sep_df), img_w, img_h),
         _graph_cell("Análisis armónico", plot_harmonics(hdf), img_w, img_h)],
        [_graph_cell("Semaforización clínica", plot_clinical_gauges(row, ppa), img_w, img_h),
         _graph_cell("RVSE / SEVR por áreas presión-tiempo", plot_rvse_area(sep_df, sep_metrics), img_w, img_h)],
    ], colWidths=[94*mm, 94*mm], rowHeights=[59*mm, 59*mm, 59*mm], style=TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("BOX", (0,0), (-1,-1), 0.25, colors.HexColor("#CFD8DC")),
        ("INNERGRID", (0,0), (-1,-1), 0.15, colors.HexColor("#ECEFF1")),
        ("LEFTPADDING", (0,0), (-1,-1), 3),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
        ("TOPPADDING", (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ]))
    story.append(graph_table)

    story.append(Spacer(1, 2.5*mm))
    story.append(_section("4. Tabla de análisis armónico y fenotipo final"))
    harm_table = [["Armónico", "Frecuencia (Hz)", "Amplitud", "Energía relativa (%)"]]
    for i, r in hdf.iterrows():
        harm_table.append([str(i+1), f"{r.get('frecuencia_hz',0):.2f}", f"{r.get('amplitud',0):.3f}", f"{r.get('energia_relativa_%',0):.1f}"])
    phenotype_rows = [[Paragraph(str(cell), styles["SmallPAC"]) for cell in row_cells] for row_cells in final_phenotype_table]
    story.append(Table([
        [Table(harm_table, colWidths=[18*mm, 25*mm, 24*mm, 28*mm], style=_table_style("#EAF2F8", 6.6)),
         Table(phenotype_rows, colWidths=[31*mm, 14*mm, 42*mm], style=_table_style("#D9EAF7", 6.4))]
    ], colWidths=[98*mm, 90*mm], style=TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 0),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
        ("TOPPADDING", (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ])))
    story.append(Spacer(1, 1.5*mm))
    story.append(Paragraph(
        "Nota metodológica: la separación Pf/Pb es una estimación clínica no invasiva. Pf y Pb se muestran sobre la línea diastólica basal para comparar directamente su contribución con la presión aórtica central completa. El fenotipo final integra presión/carga pulsátil central, magnitud y temporalidad de onda retrógrada, y complejidad armónica.",
        styles["SmallPAC"]
    ))

    if screenshot_png:
        story.append(PageBreak())
        story.append(_section("5. Captura pantalla de mediciones"))
        story.append(Spacer(1, 2*mm))
        story.append(Image(io.BytesIO(screenshot_png), width=182*mm, height=215*mm))

    story.append(Spacer(1, 3*mm))
    story.append(_section("6. Referencias bibliográficas"))
    refs = [
        ("Agabiti-Rosei E, et al. Central blood pressure measurements and "
         "antihypertensive therapy. Hypertension. 2007."),
        ("Zocalo Y, Bia D. Presion aortica central y parametros clinicos "
         "derivados de la onda del pulso. 2014."),
        ("SAHA. Manual de Mecanica Vascular. Grupo de Trabajo de Mecanica "
         "Vascular. 2024."),
        ("Westerhof BE, et al. Quantification of wave reflection in the human "
         "aorta from pressure alone. Hypertension. 2006."),
        ("Herbert A, et al. Establishing reference values for central blood "
         "pressure and amplification. Eur Heart J. 2014."),
        ("Huang QF, et al. Outcome-driven threshold for pulse pressure "
         "amplification. Hypertension Research. 2024."),
    ]
    ref_table = [[Paragraph(f"{i}. {ref_txt}", styles["SmallPAC"])] for i, ref_txt in enumerate(refs, 1)]
    story.append(Table(ref_table, colWidths=[188*mm], style=TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#FAFAFA")),
        ("BOX", (0,0), (-1,-1), 0.25, colors.HexColor("#CFD8DC")),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 1.6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 1.6),
    ])))

    doc.build(story, onFirstPage=_header_footer, onLaterPages=_header_footer)
    buf.seek(0)
    return buf.getvalue()

st.title(APP_TITLE)
st.caption("Importación tipo MODELO PAC, informe PDF integrado, captura de segunda hoja, historial Excel y análisis armónico.")

with st.sidebar:
    st.header("1) Importar estudio")
    pdf_file = st.file_uploader("PDF original PAC / Exxer", type=["pdf"])
    wave_file = st.file_uploader("Opcional: CSV/TXT curva central REAL del paciente (tiempo_ms, presion_mmHg)", type=["csv", "txt"])
    st.info("Modo datos reales: si no se carga CSV/TXT, la app digitaliza automáticamente la curva desde la imagen del PDF. No se aceptan curvas sintéticas ni genéricas.")

base = {}
screenshot = None
pdf_bytes = None
curve_debug_png = None
curve_source = ""
curve_meta = {}
if pdf_file:
    pdf_bytes = pdf_file.read()
    text = extract_pdf_text(pdf_bytes)
    base = parse_model_pac_from_pdf(pdf_bytes, text)
    screenshot = render_pdf_page_png(pdf_bytes, page_index=1)
else:
    base = parse_model_pac("")

st.subheader("Datos extraídos / edición manual")
cols = st.columns(4)
fields = ["paciente","estudio","fecha","hora","edad","sexo","peso","altura","imc","pas_radial","pad_radial","pam_radial","pp_radial","pas_central","pad_central","pam_central","pp_central","fc","au","iau","rvse","pe","apc","medicacion","diagnostico_previo"]
row = {}
for i, f in enumerate(fields):
    with cols[i%4]:
        val = base.get(f, "")
        if f in ["paciente","estudio","fecha","hora","sexo","medicacion","diagnostico_previo"]:
            row[f] = st.text_input(f, value="" if pd.isna(val) else str(val))
        else:
            row[f] = st.number_input(f, value=0.0 if pd.isna(val) else float(val), step=1.0, format="%.2f")

wave_df = None
curve_error = None
if wave_file:
    try:
        wave_df = read_curve_file_robust(wave_file, row)
        curve_source = "CSV/TXT real cargado por el usuario"
        curve_meta = {"metodo": curve_source, "puntos": int(len(wave_df))}
        st.success("Curva real importada y validada correctamente desde CSV/TXT. El análisis de ondas y armónicos usará únicamente estos puntos del estudio.")
    except Exception as e:
        curve_error = str(e)
        st.error("El archivo CSV/TXT importado no contiene una curva real válida. Se intentará digitalizar la curva desde el PDF, si está cargado.")
        st.caption(f"Detalle técnico CSV/TXT: {curve_error}")

if wave_df is None and pdf_bytes:
    try:
        with st.spinner("Digitalizando curva real desde la segunda hoja, sector superior izquierdo: panel de curva del PDF..."):
            wave_df, curve_debug_png, curve_meta = digitize_curve_from_pdf(pdf_bytes, row, max_pages=4, zoom=3.0)
        curve_source = f"PDF digitalizado automáticamente: página {curve_meta.get('pagina')} / sector {curve_meta.get('sector', 'izquierdo-superior')} / trazo {curve_meta.get('color_detectado')}"
        st.success("Curva real digitalizada desde la segunda hoja, sector superior izquierdo: panel de curva del PDF, y calibrada con PAS/PAD central del estudio. Cada paciente usará su propia morfología extraída del PDF.")
        st.caption(f"Fuente de curva: {curve_source}. Puntos generados: {curve_meta.get('puntos')}. BBox: {curve_meta.get('bbox_px')}.")
        if curve_debug_png:
            st.image(curve_debug_png, caption="Control visual: segunda hoja, sector superior izquierdo: panel de curva usado para digitalizar la curva", use_container_width=True)
    except Exception as e:
        curve_error = str(e)
        st.error("No se pudo obtener una curva real del paciente desde CSV/TXT ni desde la imagen del PDF. No se generarán curvas sintéticas.")
        st.caption(f"Detalle técnico digitalización PDF: {curve_error}")
elif wave_df is None:
    st.error("Para generar informe, separación de ondas y armónicos se debe cargar un PDF con curva visible o un CSV/TXT real del paciente. La app queda en modo estricto: no usa curva sintética ni genérica.")

dx, cat, ref, amp_sbp, ppa, risk = central_diagnosis(row)

st.subheader("Vista clínica previa")
summary_cols = st.columns(5)
summary_cols[0].metric("PAS central", f"{to_float(row.get('pas_central')):.0f} mmHg")
summary_cols[1].metric("PP central", f"{to_float(row.get('pp_central')):.0f} mmHg")
summary_cols[2].metric("PPA", f"{ppa:.2f}" if not np.isnan(ppa) else "No disponible")
summary_cols[3].metric("RVSE equipo", f"{to_float(row.get('rvse')):.0f}%" if not np.isnan(to_float(row.get('rvse'))) else "No disponible")
summary_cols[4].metric("Modo de curva", "REAL PDF/CSV" if wave_df is not None else "BLOQUEADO")

st.markdown("### Análisis de presión central y métricas")
st.write(dx)
st.write(f"Categoría braquial: {cat}. Amplificación PAS periférico-central: {amp_sbp:.1f} mmHg si disponible. Perfil agregado: {risk}.")

if wave_df is not None:
    hdf = harmonic_analysis(wave_df)
    sep_df_preview, sep_metrics_preview = estimate_wave_separation(wave_df, row)
    conclusion_blocks_preview, sep_df_preview, sep_metrics_preview, sep_interp_preview = build_continuous_conclusions(row, wave_df, hdf)

    summary_cols[3].metric("RM Pb/Pf", f"{sep_metrics_preview.get('rm', np.nan):.2f}")
    summary_cols[4].metric("RVSE calculado", f"{sep_metrics_preview.get('rvse_calculado_%', np.nan):.0f}%")
    st.caption(f"Fuente de curva real: {curve_source or curve_meta.get('metodo','no especificada')}")
    st.caption(f"Firma morfológica de curva real: {sep_metrics_preview.get('curve_id', 'sin_firma')} | Pico: {sep_metrics_preview.get('t_pico_ms', np.nan):.0f} ms | Retorno reflejo: {sep_metrics_preview.get('tref_ms', np.nan):.0f} ms")

    st.markdown("### Conclusiones clínicas continuas")
    for title, body in conclusion_blocks_preview:
        st.markdown(f"**{title}**")
        st.write(body)

    st.markdown("---")
    st.markdown("### Gráficos")
    st.image(plot_wave_separation(sep_df_preview), caption="Presión aórtica central real con onda anterógrada Pf y retrógrada Pb superpuestas", use_container_width=True)

    g1, g2 = st.columns(2)
    with g1:
        st.image(plot_waveform(wave_df), caption="Onda central real importada", use_container_width=True)
        st.image(plot_aortic_flow(sep_df_preview), caption="Flujo aórtico estimado desde curva real", use_container_width=True)
        st.image(plot_rvse_area(sep_df_preview, sep_metrics_preview), caption="RVSE / SEVR por áreas presión-tiempo", use_container_width=True)
    with g2:
        st.image(plot_pressure_comparison(row), caption="Presiones periféricas vs centrales", use_container_width=True)
        st.image(plot_harmonics(hdf), caption="Armónicos de la onda central real", use_container_width=True)

    st.image(plot_clinical_gauges(row, ppa), caption="Semaforización clínica", use_container_width=True)

    final_phenotype_preview, final_phenotype_text_preview, final_phenotype_table_preview = classify_central_pressure_phenotype(row, sep_metrics_preview, hdf)
    st.markdown("---")
    st.markdown("### Fenotipo final de presión central")
    st.success(final_phenotype_preview)
    st.write(final_phenotype_text_preview)
    st.dataframe(pd.DataFrame(final_phenotype_table_preview[1:], columns=final_phenotype_table_preview[0]), use_container_width=True)
else:
    st.warning("Carga pendiente: PDF con curva visible o archivo CSV/TXT de curva real con columnas tiempo_ms y presion_mmHg, o equivalentes reconocibles. Sin curva real no se habilita el PDF final.")
    st.image(plot_pressure_comparison(row), caption="Presiones periféricas vs centrales extraídas del estudio", use_container_width=True)

st.subheader("Historial y exportación")
if st.button("Guardar en historial"):
    hist = save_history(row)
    st.success(f"Registro guardado. Total: {len(hist)} estudios.")

if HISTORIAL_FILE.exists():
    hist = pd.read_excel(HISTORIAL_FILE)
    st.dataframe(hist, use_container_width=True)
    st.download_button("Descargar historial Excel", HISTORIAL_FILE.read_bytes(), file_name="historial_pac.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if wave_df is None:
    st.error("PDF médico integrado no habilitado: falta curva real válida del paciente desde CSV/TXT o digitalización del PDF. No se generará reporte con curvas simuladas.")
else:
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
