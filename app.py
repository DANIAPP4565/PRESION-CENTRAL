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



def read_curve_file_robust(uploaded_file):
    """Lee CSV/TXT de curva sin fallar por codificación ni por formato irregular.

    Soporta:
    - UTF-8, UTF-8-SIG, CP1252, Latin-1, ISO-8859-1.
    - CSV con coma, punto y coma, tabulador o espacios.
    - Archivos TXT sin encabezado.
    - Archivos con una sola columna de presión.
    - Archivos con texto del equipo mezclado con números.
    - Decimales con coma o punto.
    Devuelve DataFrame normalizado: tiempo_ms, presion_mmHg.
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

    # 1) Intento estructurado: CSV/TXT con separadores habituales.
    for sep in (None, ";", ",", "\t", r"\s+"):
        try:
            df = pd.read_csv(
                io.StringIO(text),
                sep=sep,
                engine="python",
                decimal=",",
                on_bad_lines="skip"
            )
            if df is not None and df.shape[0] >= 3:
                try:
                    return normalize_wave_dataframe(df)
                except Exception:
                    pass
        except Exception:
            pass

    # 2) Intento sin encabezado: todas las líneas con números.
    numeric_rows = []
    for line in text.splitlines():
        nums = re.findall(r"[-+]?\d+(?:[\.,]\d+)?", line)
        if nums:
            numeric_rows.append([to_float(n) for n in nums if not np.isnan(to_float(n))])

    # 2a) Pares tiempo-presión en líneas.
    pair_rows = []
    for row in numeric_rows:
        if len(row) >= 2:
            pair_rows.append([row[0], row[1]])
    if len(pair_rows) >= 3:
        try:
            return normalize_wave_dataframe(pd.DataFrame(pair_rows, columns=["tiempo_ms", "presion_mmHg"]))
        except Exception:
            pass

    # 2b) Secuencia larga de presiones sin tiempo: una presión por línea o vector.
    flat = []
    for row in numeric_rows:
        flat.extend(row)

    # Filtra valores fisiológicos compatibles con presión arterial de curva.
    pressure_like = [v for v in flat if 20 <= v <= 260]
    if len(pressure_like) >= 8:
        tiempo = np.linspace(0, 1000, len(pressure_like))
        return pd.DataFrame({"tiempo_ms": tiempo, "presion_mmHg": pressure_like})

    raise ValueError(
        "No se pudieron reconocer columnas de tiempo y presión en el archivo CSV/TXT. "
        "El lector acepta columnas como tiempo/time/ms/x y presión/pressure/PAC/mmHg/y, "
        "o una secuencia simple de valores de presión."
    )


def normalize_wave_dataframe(df):
    """Normaliza cualquier tabla de curva a tiempo_ms/presion_mmHg."""
    df = df.copy()

    # Si pandas tomó todo como una única columna con separadores internos, reintentar expandir.
    if df.shape[1] == 1:
        col = df.columns[0]
        joined = "\n".join(df[col].astype(str).tolist())
        for sep in (";", ",", "\t", r"\s+"):
            try:
                tmp = pd.read_csv(io.StringIO(joined), sep=sep, engine="python", header=None, on_bad_lines="skip")
                if tmp.shape[1] >= 2 and tmp.shape[0] >= 3:
                    df = tmp
                    break
            except Exception:
                pass

    original_cols = list(df.columns)
    df.columns = [str(c).strip().lower() for c in df.columns]

    # Convertir todo lo posible a numérico.
    num = pd.DataFrame()
    for c in df.columns:
        num[c] = df[c].map(to_float)

    valid_numeric_cols = [c for c in num.columns if num[c].notna().sum() >= 3]

    if len(valid_numeric_cols) >= 2:
        time_candidates = [
            c for c in valid_numeric_cols
            if any(k in str(c).lower() for k in ["tiempo", "time", "ms", "mseg", "miliseg", "seg", "sec", "x"])
        ]
        pressure_candidates = [
            c for c in valid_numeric_cols
            if any(k in str(c).lower() for k in ["pres", "pressure", "pao", "pac", "central", "mmhg", "aort", "y"])
        ]

        # Si no hay nombres claros, detectar por comportamiento:
        # tiempo = columna más monótonamente creciente; presión = rango fisiológico arterial.
        if time_candidates:
            tcol = time_candidates[0]
        else:
            monotonic_scores = {}
            for c in valid_numeric_cols:
                s = num[c].dropna().astype(float)
                if len(s) >= 3:
                    diffs = np.diff(s.values)
                    monotonic_scores[c] = np.mean(diffs >= 0)
            tcol = max(monotonic_scores, key=monotonic_scores.get)

        pcol = None
        for c in pressure_candidates:
            if c != tcol:
                pcol = c
                break

        if pcol is None:
            candidates = []
            for c in valid_numeric_cols:
                if c == tcol:
                    continue
                s = num[c].dropna().astype(float)
                if len(s) >= 3:
                    med = float(np.nanmedian(s))
                    amp = float(np.nanmax(s) - np.nanmin(s))
                    # presión arterial central/radial plausible
                    score = 0
                    if 40 <= med <= 180:
                        score += 2
                    if 10 <= amp <= 140:
                        score += 1
                    candidates.append((score, c))
            if candidates:
                pcol = sorted(candidates, reverse=True)[0][1]
            else:
                pcol = [c for c in valid_numeric_cols if c != tcol][0]

        out = pd.DataFrame({
            "tiempo_ms": num[tcol],
            "presion_mmHg": num[pcol],
        }).dropna()

    elif len(valid_numeric_cols) == 1:
        # Una sola columna numérica: interpretarla como presión y generar tiempo 0-1000 ms.
        pcol = valid_numeric_cols[0]
        pressure = num[pcol].dropna().astype(float)
        pressure = pressure[(pressure >= 20) & (pressure <= 260)]
        if len(pressure) < 3:
            raise ValueError("La columna única no contiene suficientes valores fisiológicos de presión.")
        out = pd.DataFrame({
            "tiempo_ms": np.linspace(0, 1000, len(pressure)),
            "presion_mmHg": pressure.values,
        })
    else:
        raise ValueError("No se detectaron columnas numéricas válidas en el archivo de curva.")

    out = out.sort_values("tiempo_ms").drop_duplicates("tiempo_ms")
    out = out.replace([np.inf, -np.inf], np.nan).dropna()

    if len(out) < 3:
        raise ValueError("La curva debe tener al menos 3 pares válidos de tiempo-presión.")

    # Si el tiempo parece estar en segundos, pasarlo a ms.
    if out["tiempo_ms"].max() <= 5:
        out["tiempo_ms"] = out["tiempo_ms"] * 1000.0

    # Si el tiempo no cubre un ciclo razonable, reescalar a 0-1000 ms preservando forma.
    if out["tiempo_ms"].max() - out["tiempo_ms"].min() <= 0:
        out["tiempo_ms"] = np.linspace(0, 1000, len(out))
    else:
        tmin = out["tiempo_ms"].min()
        tmax = out["tiempo_ms"].max()
        if tmax > 5000 or tmax < 100:
            out["tiempo_ms"] = (out["tiempo_ms"] - tmin) / (tmax - tmin) * 1000.0

    return out.reset_index(drop=True)

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

def make_waveform(row, n=512):
    cSBP = to_float(row.get("pas_central")); cDBP = to_float(row.get("pad_central"));
    if np.isnan(cSBP): cSBP = 120
    if np.isnan(cDBP): cDBP = 80
    t = np.linspace(0, 1, n)
    pp = cSBP - cDBP
    # Onda sintética didáctica, reemplazable por CSV real de presión central.
    primary = np.exp(-((t-0.22)/0.10)**2)
    reflected = 0.35*np.exp(-((t-0.42)/0.13)**2)
    diast = 0.20*np.exp(-((t-0.68)/0.23)**2)
    raw = primary + reflected + diast
    p = cDBP + pp*(raw - raw.min())/(raw.max()-raw.min())
    return pd.DataFrame({"tiempo_ms": t*1000, "presion_central_mmHg": p})

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
    fig, ax = plt.subplots(figsize=(7,4))
    ax.plot(wave_df.iloc[:,0], wave_df.iloc[:,1], linewidth=2)
    ax.set_xlabel("Tiempo (ms)"); ax.set_ylabel("Presión central (mmHg)"); ax.set_title("Onda de presión aórtica central")
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
    conclusion = f"Diagnóstico: {dx} Categoría braquial: {cat}. Referencia central operativa P50: {ref if not np.isnan(ref) else 'no disponible'} mmHg. Amplificación PAS periférico-central: {amp_sbp:.1f} mmHg. PPA: {ppa:.2f} si disponible. Perfil agregado: {risk}."
    story.append(Paragraph("Conclusión clínica", styles["Heading2"]))
    story.append(Paragraph(conclusion, styles["BodyText"]))
    story.append(Spacer(1, 4*mm))
    for title, png in [
        ("Gráfico comparativo de presiones", plot_pressure_comparison(row)),
        ("Onda de presión central", plot_waveform(wave_df)),
        ("Análisis armónico", plot_harmonics(hdf)),
        ("Semaforización clínica", plot_clinical_gauges(row, ppa)),
    ]:
        story.append(KeepTogether([Paragraph(title, styles["Heading3"]), Image(png, width=170*mm, height=90*mm)]))
    story.append(PageBreak())
    story.append(Paragraph("Análisis armónico de la onda de presión central", styles["Heading2"]))
    story.append(Paragraph(
        "Se calcula por transformada rápida de Fourier sobre la onda central importada o, "
        "si no se adjunta curva digitalizada, sobre una curva sintética calibrada con la presión "
        "sistólica central, presión diastólica central, presión de pulso, aumentación aórtica e "
        "índice de aumentación del estudio. El análisis armónico se interpreta como una estimación "
        "fisiológica de la distribución espectral de energía de la onda de presión central.",
        styles["Normal"]
    ))
    

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
        wave_df = read_curve_file_robust(wave_file)
        st.success("Curva importada correctamente con lector robusto CSV/TXT.")
    except Exception as e:
        st.error(f"No se pudo importar la curva. Se generará una curva sintética desde las métricas. Detalle: {e}")
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
    st.image(plot_waveform(wave_df), caption="Onda central")
    st.image(plot_clinical_gauges(row, ppa), caption="Semaforización clínica")

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
