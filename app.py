# app_pac.py
# App Streamlit para informe de Presión Aórtica Central (PAC)
# Importa PDF tipo MODELO PAC, extrae variables, genera historial Excel y PDF médico integrado.

import io, re, math, tempfile, json, textwrap, zipfile, gc, hashlib
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import matplotlib.pyplot as plt

try:
  import fitz # PyMuPDF
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
from reportlab.lib.utils import ImageReader
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


def read_uploaded_image_bytes(uploaded_file):
  """Devuelve bytes de imagen cargada, compatibles con ReportLab y Streamlit."""
  if uploaded_file is None:
    return None
  try:
    data = uploaded_file.getvalue()
  except Exception:
    try:
      data = uploaded_file.read()
    except Exception:
      data = None
  if not data:
    return None
  return bytes(data)


def image_aspect_ratio(image_bytes):
  """Calcula ancho/alto de una imagen preservando proporciones en el PDF."""
  if not image_bytes:
    return None
  try:
    if PILImage is not None:
      im = PILImage.open(io.BytesIO(image_bytes))
      w, h = im.size
      if w > 0 and h > 0:
        return float(w) / float(h)
    reader = ImageReader(io.BytesIO(image_bytes))
    w, h = reader.getSize()
    if w > 0 and h > 0:
      return float(w) / float(h)
  except Exception:
    return None
  return None


def fit_image_box(image_bytes, max_w, max_h):
  """Ajusta una imagen a una caja máxima sin deformar ni truncar."""
  ar = image_aspect_ratio(image_bytes)
  if not ar or ar <= 0:
    return max_w, max_h
  if max_w / max_h > ar:
    h = max_h
    w = h * ar
  else:
    w = max_w
    h = w / ar
  return w, h


st.set_page_config(page_title="INFORME MEDICION DE PRESION CENTRAL", layout="wide")

APP_TITLE = "INFORME MEDICION DE PRESION CENTRAL"
HISTORIAL_FILE = Path("historial_pac.xlsx")

BIBLIOGRAFIA = [
  "Agabiti-Rosei E, Mancia G, O'Rourke MF, et al. Central blood pressure measurements and antihypertensive therapy: a consensus document. Hypertension. 2007;50:154-160.",
  "Zócalo Y, Bia D. Presión aórtica central y parámetros clínicos derivados de la onda del pulso: evaluación no invasiva en la práctica clínica. Rev Urug Cardiol. 2014;29:215-230.",
  "Herbert A, Cruickshank JK, Laurent S, Boutouyrie P, et al. Establishing reference values for central blood pressure and its amplification. Eur Heart J. 2014;35:3122-3133.",
  "Westerhof BE, Guelen I, Westerhof N, Karemaker JM, Avolio A. Quantification of wave reflection in the human aorta from pressure alone. Hypertension. 2006;48:595-601.",
  "Norton GR, An DW, Aparicio LS, et al. Mortality and cardiovascular end points in relation to the aortic pulse wave components. Hypertension. 2024;81:1065-1075.",
  "Huang QF, An DW, Aparicio LS, et al. An outcome-driven threshold for pulse pressure amplification. Hypertension Research. 2024.",
  "SAHA. Manual de Mecánica Vascular. Grupo de Trabajo de Mecánica Vascular de la Sociedad Argentina de Hipertensión Arterial. 2024. Tabla de presión aórtica central por edad, sexo y calibración.",
  "Azizzadeh M, Karimi A, Breyer-Kohansal R, et al. Reference equations for pulse wave velocity, augmentation index, amplitude of forward and backward wave in a European general adult population. Scientific Reports. 2024;14:23151.",
  "Zócalo Y, Bia D. Central Pressure Waveform-Derived Indexes Obtained From Carotid and Radial Tonometry and Brachial Oscillometry in Healthy Subjects (2-84 Y): Age-, Height-, and Sex-Related Profiles and Analysis of Indexes Agreement. Front Physiol. 2022;12:774390. doi:10.3389/fphys.2021.774390."

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


# Referencias del Manual de Mecánica Vascular SAHA 2024, Tabla 2:
# Presión arterial sistólica aórtica central por sexo, grupo etario y esquema de calibración.
# Valores = percentil 50 (mediana) y DE. Se usa P90 = mediana + 1.282*DE como
# límite diagnóstico binario: CON HIPERTENSIÓN CENTRAL si PAS central >= P90 y
# SIN HIPERTENSIÓN CENTRAL si PAS central < P90.
SAHA_AOSBP_REF = {
  "M": {
    "<20": {"C_PAOC": (110.29, 16.41), "SD_PAOC": (102.41, 10.22)},
    "20-29":{"C_PAOC": (122.13, 14.22), "SD_PAOC": (112.48, 8.87)},
    "30-39":{"C_PAOC": (120.86, 12.56), "SD_PAOC": (111.97, 9.44)},
    "40-49":{"C_PAOC": (119.13, 11.95), "SD_PAOC": (113.83, 9.04)},
    "50-59":{"C_PAOC": (117.42, 9.82), "SD_PAOC": (112.65, 9.30)},
    "60-69":{"C_PAOC": (117.91, 10.23), "SD_PAOC": (114.73, 11.93)},
    ">=70": {"C_PAOC": (121.25, 4.86), "SD_PAOC": (108.25, 20.17)},
  },
  "F": {
    "<20": {"C_PAOC": (101.39, 13.61), "SD_PAOC": ( 99.14, 9.05)},
    "20-29":{"C_PAOC": (104.69, 10.06), "SD_PAOC": (104.26, 10.80)},
    "30-39":{"C_PAOC": (103.50, 10.21), "SD_PAOC": (105.63, 10.57)},
    "40-49":{"C_PAOC": (107.23, 8.47), "SD_PAOC": (111.38, 8.57)},
    "50-59":{"C_PAOC": (113.11, 10.80), "SD_PAOC": (114.33, 9.04)},
    "60-69":{"C_PAOC": (112.88, 11.01), "SD_PAOC": (113.81, 11.78)},
    ">=70": {"C_PAOC": (112.75, 11.17), "SD_PAOC": (111.92, 5.82)},
  },
}



# Referencias para IAu/AIx: LEAD 2024 / SphygmoCor por edad y sexo.
# Por indicación clínica de la app, la clasificación del índice de aumentación central
# se realiza con datos reales del paciente y referencia LEAD 2024/SphygmoCor.
# La publicación LEAD 2024 informa AIx por edad y sexo como media y rango media±2DE.
# Aquí se estima DE=(límite superior-límite inferior)/4 y se derivan percentiles
# normales aproximados para P10/P25/P50/P75/P90.
LEAD2024_AIX_MEAN_SD = {
  "M": {
    "18-30": (0.6, (23.4 - (-22.3))/4.0),
    "30-40": (8.7, (31.3 - (-14.0))/4.0),
    "40-50": (17.4, (40.2 - (-5.4))/4.0),
    "50-60": (22.9, (42.5 - 3.2)/4.0),
    "60-70": (28.3, (48.8 - 7.8)/4.0),
    ">=70": (30.1, (48.3 - 12.0)/4.0),
  },
  "F": {
    "18-30": (8.2, (31.6 - (-15.3))/4.0),
    "30-40": (18.2, (40.3 - (-3.9))/4.0),
    "40-50": (27.7, (47.8 - 7.6)/4.0),
    "50-60": (33.9, (55.2 - 12.6)/4.0),
    "60-70": (35.2, (53.2 - 17.2)/4.0),
    ">=70": (36.1, (55.5 - 16.7)/4.0),
  },
}

def saha_aix_age_group(age):
  """Grupo etario compatible con LEAD 2024 para AIx/IAu."""
  a = to_float(age)
  if np.isnan(a) or a < 18:
    return None
  if a < 30: return "18-30"
  if a < 40: return "30-40"
  if a < 50: return "40-50"
  if a < 60: return "50-60"
  if a < 70: return "60-70"
  return ">=70"

def _percentiles_from_mean_sd(mean, sd):
  return (
    mean - 1.282 * sd,
    mean - 0.674 * sd,
    mean,
    mean + 0.674 * sd,
    mean + 1.282 * sd,
  )

def _percentile_from_percentile_table(value, pcts=(10,25,50,75,90), vals=None):
  """Interpolación simple de percentil desde una tabla P10-P90."""
  try:
    y = float(value)
    xs = np.asarray(pcts, dtype=float)
    vs = np.asarray(vals, dtype=float)
    if len(vs) != len(xs) or np.any(~np.isfinite(vs)):
      return np.nan
    if y <= vs[0]:
      slope = (xs[1]-xs[0]) / max(vs[1]-vs[0], 1e-6)
      return float(np.clip(xs[0] + (y-vs[0])*slope, 0, 10))
    if y >= vs[-1]:
      slope = (xs[-1]-xs[-2]) / max(vs[-1]-vs[-2], 1e-6)
      return float(np.clip(xs[-1] + (y-vs[-1])*slope, 90, 99))
    return float(np.interp(y, vs, xs))
  except Exception:
    return np.nan

def get_saha_aix75_reference(row):
  """Clasifica IAu/AIx central con referencia LEAD 2024 / SphygmoCor.

  Mantiene el nombre de función por compatibilidad interna, pero la referencia usada
  es LEAD 2024/SphygmoCor por edad y sexo. La variable se calcula sobre el dato real
  del paciente y/o sobre la onda central real digitalizada del estudio original.
  """
  sex = safe_text(row.get("sexo", "M")).upper()[:1]
  if sex not in ("M", "F"):
    sex = "M"
  age_group = saha_aix_age_group(row.get("edad"))
  iau = to_float(row.get("iau"))
  if not age_group or sex not in LEAD2024_AIX_MEAN_SD or age_group not in LEAD2024_AIX_MEAN_SD[sex]:
    return {"ok": False, "motivo": "edad o sexo no disponibles para aplicar referencia LEAD 2024 de IAu/AIx"}
  if np.isnan(iau):
    return {"ok": False, "motivo": "IAu/AIx central no disponible"}
  mean, sd = LEAD2024_AIX_MEAN_SD[sex][age_group]
  p10, p25, p50, p75, p90 = _percentiles_from_mean_sd(mean, sd)
  pct = _percentile_from_percentile_table(iau, vals=(p10, p25, p50, p75, p90))
  z = (iau - mean) / sd if sd and not np.isnan(sd) else np.nan
  if iau >= p90:
    categoria = "CON AUMENTACIÓN CENTRAL AUMENTADA PARA EDAD Y SEXO"
    severidad = "con aumentación central aumentada"
    alterada = True
  else:
    categoria = "SIN AUMENTACIÓN CENTRAL AUMENTADA PARA EDAD Y SEXO"
    severidad = "sin aumentación central aumentada"
    alterada = False
  return {
    "ok": True, "sexo": sex, "edad_grupo": age_group, "iau": iau,
    "p10": p10, "p25": p25, "p50": p50, "p75": p75, "p90": p90,
    "percentil": pct, "z_aprox": z, "categoria": categoria,
    "severidad": severidad, "alterada": alterada,
    "metodo": "LEAD 2024 / SphygmoCor AIx por edad y sexo",
  }

def format_saha_aix75(row, compact=False):
  ref = get_saha_aix75_reference(row)
  if not ref.get("ok"):
    return "Clasificación IAu/AIx no disponible: " + ref.get("motivo", "datos insuficientes")
  txt = (
    f"Diagnóstico de aumentación central por IAu/AIx: {ref['categoria']}. "
    f"IAu {ref['iau']:.1f}%; referencia LEAD 2024 {ref['sexo']}, {ref['edad_grupo']} años: "
    f"P50 {ref['p50']:.1f}%, P75 {ref['p75']:.1f}% y P90 {ref['p90']:.1f}%. "
    f"Percentil estimado {ref['percentil']:.0f}; z aproximado {ref['z_aprox']:.2f}."
  )
  if not compact:
    txt += " La decisión del informe es binaria: CON AUMENTACIÓN CENTRAL AUMENTADA o SIN AUMENTACIÓN CENTRAL AUMENTADA; el punto de corte es IAu/AIx >= P90 de la referencia LEAD 2024/SphygmoCor."
  return txt



# ---------------------------------------------------------------------------
# REFERENCIAS RM / RIx POR EDAD Y MÉTODO
# Zócalo Y, Bia D. Front Physiol. 2022;12:774390. Tabla 4.
# Percentiles publicados para sujetos sanos: P50, P75, P90, P95, P97.5 y P99.
# Criterio operativo de esta app (acordado):
#   <P75       = esperado
#   P75-<P90   = relativamente elevado (sin criterio diagnóstico de aumento)
#   P90-<P95   = aumentado
#   >=P95      = marcadamente aumentado
# RM es la métrica primaria. RIx se informa como complemento proporcional y NO
# suma una segunda evidencia diagnóstica independiente.
# ---------------------------------------------------------------------------
RM_RI_METHOD_LABELS = {
  "MOG": "Oscilometría braquial / Mobil-O-Graph (Zócalo-Bia 2022)",
  "SCOR_RADIAL": "Tonometría radial / SphygmoCor (Zócalo-Bia 2022)",
  "SCOR_CAROTID": "Tonometría carotídea / SphygmoCor (Zócalo-Bia 2022)",
}

RM_RI_AGES = [3, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80, 84]
RM_RI_PERCENTILES = [50.0, 75.0, 90.0, 95.0, 97.5, 99.0]

RM_RI_AGE_REFERENCE = {
  "RM": {
    "MOG": [
      [0.54,0.59,0.63,0.66,0.68,0.70], [0.54,0.59,0.64,0.66,0.68,0.70],
      [0.56,0.61,0.65,0.67,0.69,0.72], [0.58,0.62,0.67,0.69,0.71,0.73],
      [0.59,0.64,0.68,0.70,0.72,0.74], [0.61,0.65,0.69,0.72,0.73,0.76],
      [0.62,0.67,0.71,0.73,0.75,0.77], [0.63,0.68,0.72,0.74,0.76,0.78],
      [0.64,0.69,0.73,0.75,0.77,0.79], [0.66,0.70,0.74,0.76,0.78,0.80],
      [0.67,0.71,0.75,0.77,0.79,0.81], [0.68,0.72,0.76,0.78,0.80,0.82],
      [0.69,0.73,0.77,0.79,0.80,0.82], [0.69,0.74,0.77,0.79,0.81,0.83],
      [0.70,0.74,0.78,0.80,0.82,0.84], [0.71,0.75,0.79,0.81,0.83,0.85],
      [0.72,0.76,0.80,0.82,0.83,0.85], [0.73,0.77,0.80,0.82,0.84,0.86],
    ],
    "SCOR_RADIAL": [
      [0.65,0.73,0.80,0.85,0.89,0.95], [0.53,0.59,0.67,0.71,0.76,0.81],
      [0.44,0.50,0.57,0.62,0.67,0.72], [0.42,0.49,0.56,0.61,0.66,0.72],
      [0.42,0.50,0.58,0.63,0.68,0.74], [0.44,0.52,0.60,0.66,0.71,0.78],
      [0.46,0.54,0.63,0.69,0.75,0.83], [0.49,0.57,0.67,0.74,0.80,0.88],
      [0.51,0.61,0.71,0.78,0.85,0.94], [0.55,0.65,0.76,0.84,0.91,1.01],
      [0.58,0.69,0.81,0.90,0.98,1.08], [0.62,0.74,0.87,0.96,1.05,1.16],
      [0.67,0.79,0.93,1.03,1.12,1.24], [0.71,0.85,1.00,1.10,1.20,1.34],
      [0.76,0.91,1.07,1.18,1.29,1.43], [0.82,0.98,1.15,1.27,1.39,1.54],
      [0.87,1.05,1.23,1.36,1.49,1.65], [0.92,1.11,1.30,1.44,1.58,1.75],
    ],
    "SCOR_CAROTID": [
      [0.42,0.49,0.56,0.61,0.65,0.70], [0.38,0.45,0.52,0.56,0.60,0.66],
      [0.35,0.42,0.50,0.54,0.59,0.64], [0.36,0.43,0.51,0.56,0.61,0.66],
      [0.37,0.45,0.53,0.59,0.63,0.70], [0.39,0.48,0.56,0.62,0.67,0.73],
      [0.42,0.51,0.59,0.65,0.71,0.78], [0.44,0.54,0.63,0.69,0.75,0.82],
      [0.47,0.57,0.67,0.73,0.79,0.86], [0.50,0.60,0.70,0.77,0.83,0.91],
      [0.53,0.64,0.74,0.81,0.88,0.96], [0.56,0.67,0.78,0.86,0.93,1.01],
      [0.59,0.71,0.83,0.90,0.97,1.06], [0.62,0.75,0.87,0.95,1.02,1.11],
      [0.66,0.78,0.91,0.99,1.07,1.16], [0.69,0.82,0.95,1.04,1.12,1.22],
      [0.73,0.86,1.00,1.09,1.17,1.27], [0.76,0.90,1.04,1.13,1.21,1.32],
    ],
  },
  "RI": {
    "MOG": [
      [0.35,0.37,0.39,0.40,0.40,0.41], [0.35,0.37,0.39,0.40,0.41,0.41],
      [0.36,0.38,0.39,0.40,0.41,0.42], [0.37,0.38,0.40,0.41,0.42,0.42],
      [0.37,0.39,0.40,0.41,0.42,0.43], [0.38,0.40,0.41,0.42,0.42,0.43],
      [0.38,0.40,0.41,0.42,0.43,0.44], [0.39,0.40,0.42,0.43,0.43,0.44],
      [0.39,0.41,0.42,0.43,0.44,0.44], [0.40,0.41,0.42,0.43,0.44,0.45],
      [0.40,0.42,0.43,0.44,0.44,0.45], [0.40,0.42,0.43,0.44,0.44,0.45],
      [0.41,0.42,0.43,0.44,0.45,0.45], [0.41,0.42,0.44,0.44,0.45,0.46],
      [0.41,0.43,0.44,0.45,0.45,0.46], [0.42,0.43,0.44,0.45,0.45,0.46],
      [0.42,0.43,0.44,0.45,0.46,0.46], [0.42,0.43,0.45,0.45,0.46,0.46],
    ],
    "SCOR_RADIAL": [
      [0.39,0.42,0.44,0.46,0.47,0.49], [0.34,0.37,0.40,0.42,0.43,0.45],
      [0.30,0.34,0.36,0.38,0.40,0.42], [0.30,0.33,0.36,0.38,0.40,0.42],
      [0.30,0.33,0.37,0.39,0.40,0.43], [0.31,0.34,0.37,0.40,0.42,0.44],
      [0.32,0.35,0.39,0.41,0.43,0.45], [0.33,0.36,0.40,0.42,0.44,0.47],
      [0.34,0.38,0.42,0.44,0.46,0.49], [0.35,0.39,0.43,0.46,0.48,0.50],
      [0.37,0.41,0.45,0.47,0.50,0.52], [0.38,0.43,0.47,0.49,0.51,0.54],
      [0.40,0.44,0.48,0.51,0.53,0.56], [0.42,0.46,0.50,0.53,0.55,0.58],
      [0.43,0.48,0.52,0.55,0.57,0.60], [0.45,0.50,0.54,0.57,0.59,0.62],
      [0.47,0.51,0.56,0.59,0.61,0.64], [0.48,0.53,0.58,0.60,0.63,0.66],
    ],
    "SCOR_CAROTID": [
      [0.30,0.33,0.36,0.38,0.39,0.41], [0.27,0.31,0.34,0.36,0.38,0.40],
      [0.26,0.30,0.33,0.35,0.37,0.39], [0.26,0.30,0.34,0.36,0.38,0.40],
      [0.27,0.31,0.35,0.37,0.39,0.41], [0.28,0.32,0.36,0.38,0.40,0.42],
      [0.29,0.34,0.37,0.40,0.42,0.44], [0.31,0.35,0.39,0.41,0.43,0.45],
      [0.32,0.36,0.40,0.42,0.44,0.47], [0.33,0.37,0.41,0.44,0.46,0.48],
      [0.34,0.39,0.43,0.45,0.47,0.50], [0.36,0.40,0.44,0.46,0.49,0.51],
      [0.37,0.41,0.45,0.48,0.50,0.52], [0.38,0.43,0.47,0.49,0.51,0.54],
      [0.40,0.44,0.48,0.50,0.53,0.55], [0.41,0.45,0.49,0.52,0.54,0.56],
      [0.42,0.47,0.51,0.53,0.55,0.58], [0.43,0.48,0.52,0.54,0.56,0.59],
    ],
  },
}


def _interp_rmri_percentiles(metric, method, age):
  """Interpola percentiles de Tabla 4 entre edades publicadas; limita a 3-84 años."""
  try:
    metric = str(metric).upper()
    method = str(method).upper()
    a = float(age)
    table = RM_RI_AGE_REFERENCE[metric][method]
    a_clip = float(np.clip(a, RM_RI_AGES[0], RM_RI_AGES[-1]))
    vals = []
    for j in range(len(RM_RI_PERCENTILES)):
      col = [float(row[j]) for row in table]
      vals.append(float(np.interp(a_clip, RM_RI_AGES, col)))
    return dict(zip(["p50","p75","p90","p95","p97_5","p99"], vals)), a_clip
  except Exception:
    return None, np.nan


def _percentile_from_rmri_reference(value, ref):
  """Percentil aproximado por interpolación/extrapolación local P50-P99."""
  try:
    v = float(value)
    pcts = np.asarray(RM_RI_PERCENTILES, dtype=float)
    vals = np.asarray([ref["p50"], ref["p75"], ref["p90"], ref["p95"], ref["p97_5"], ref["p99"]], dtype=float)
    if np.any(~np.isfinite(vals)):
      return np.nan
    if v <= vals[0]:
      slope = (pcts[1]-pcts[0]) / max(vals[1]-vals[0], 1e-6)
      return float(np.clip(pcts[0] + (v-vals[0])*slope, 0, 50))
    if v >= vals[-1]:
      slope = (pcts[-1]-pcts[-2]) / max(vals[-1]-vals[-2], 1e-6)
      return float(np.clip(pcts[-1] + (v-vals[-1])*slope, 99, 99.9))
    return float(np.interp(v, vals, pcts))
  except Exception:
    return np.nan


def _classify_rmri_value(value, ref):
  try:
    v = float(value)
    if np.isnan(v) or not ref:
      return {"ok": False, "categoria": "no clasificable", "alterada": False, "grado": "sin_datos"}
    if v < ref["p75"]:
      cat, altered, grade = "esperada para edad y método", False, "esperada"
    elif v < ref["p90"]:
      cat, altered, grade = "relativamente elevada (P75-P90), sin criterio diagnóstico de aumento", False, "intermedia"
    elif v < ref["p95"]:
      cat, altered, grade = "aumentada (P90-P95)", True, "aumentada"
    else:
      cat, altered, grade = "marcadamente aumentada (>=P95)", True, "marcada"
    return {"ok": True, "categoria": cat, "alterada": altered, "grado": grade,
            "percentil": _percentile_from_rmri_reference(v, ref)}
  except Exception:
    return {"ok": False, "categoria": "no clasificable", "alterada": False, "grado": "sin_datos"}


def get_rm_ri_reference(row, sep_metrics):
  """Referencia edad/método para RM y RIx. RM primaria; RIx complementaria no independiente."""
  age = to_float(row.get("edad"))
  method = safe_text(row.get("metodo_referencia_rmri", "SCOR_RADIAL")).upper()
  if method not in RM_RI_METHOD_LABELS:
    method = "SCOR_RADIAL"
  if np.isnan(age):
    return {"ok": False, "motivo": "edad no disponible", "metodo": method,
            "metodo_label": RM_RI_METHOD_LABELS[method]}
  rm_ref, age_used = _interp_rmri_percentiles("RM", method, age)
  ri_ref, _ = _interp_rmri_percentiles("RI", method, age)
  if not rm_ref or not ri_ref:
    return {"ok": False, "motivo": "referencia no disponible", "metodo": method,
            "metodo_label": RM_RI_METHOD_LABELS[method]}
  rm = to_float((sep_metrics or {}).get("rm"))
  ri = to_float((sep_metrics or {}).get("ri"))
  rm_cls = _classify_rmri_value(rm, rm_ref) if not np.isnan(rm) else {"ok": False, "alterada": False, "categoria": "no disponible", "grado": "sin_datos"}
  ri_cls = _classify_rmri_value(ri, ri_ref) if not np.isnan(ri) else {"ok": False, "alterada": False, "categoria": "no disponible", "grado": "sin_datos"}
  return {
    "ok": True, "edad": float(age), "edad_usada": age_used, "metodo": method,
    "metodo_label": RM_RI_METHOD_LABELS[method], "rm": rm, "ri": ri,
    "rm_ref": rm_ref, "ri_ref": ri_ref, "rm_clasif": rm_cls, "ri_clasif": ri_cls,
    "fuente": "Zócalo y Bia 2022, Front Physiol 12:774390, Tabla 4",
    "nota": "RM es primaria; RIx es complementaria y no se cuenta como evidencia diagnóstica independiente.",
  }


def format_rm_ri_reference(row, sep_metrics, compact=False):
  ref = get_rm_ri_reference(row, sep_metrics)
  if not ref.get("ok"):
    return "Clasificación RM/RI no disponible: " + ref.get("motivo", "datos insuficientes")
  rm_txt = "RM no disponible"
  if ref["rm_clasif"].get("ok"):
    rm_txt = (
      f"RM {ref['rm']:.2f}: {ref['rm_clasif']['categoria']}; "
      f"P75 {ref['rm_ref']['p75']:.2f}, P90 {ref['rm_ref']['p90']:.2f}, P95 {ref['rm_ref']['p95']:.2f}, "
      f"percentil estimado {ref['rm_clasif'].get('percentil', np.nan):.0f}"
    )
  ri_txt = "RI no disponible"
  if ref["ri_clasif"].get("ok"):
    ri_txt = (
      f"RI {ref['ri']:.2f}: {ref['ri_clasif']['categoria']}; "
      f"P90 {ref['ri_ref']['p90']:.2f} (complementaria, no independiente)"
    )
  txt = f"{rm_txt}. {ri_txt}. Referencia: {ref['metodo_label']}, edad {ref['edad']:.0f} años."
  if not compact:
    txt += " Criterio: <P75 esperado; P75-<P90 relativamente elevado; P90-<P95 aumentado; >=P95 marcadamente aumentado."
  return txt


SAHA_CALIBRATION_LABELS = {
  "SD_PAOC": "SD_PAOC: calibración sistólico-diastólica a PAS/PAD braquial",
  "C_PAOC": "C_PAOC: calibración a PAD braquial y PAM calculada",
}

def _norm_cdf(z):
  try:
    return 0.5 * (1.0 + math.erf(float(z) / math.sqrt(2.0)))
  except Exception:
    return np.nan

def saha_age_group(age):
  a = to_float(age)
  if np.isnan(a):
    return None
  if a < 20: return "<20"
  if a < 30: return "20-29"
  if a < 40: return "30-39"
  if a < 50: return "40-49"
  if a < 60: return "50-59"
  if a < 70: return "60-69"
  return ">=70"

def get_saha_central_sbp_reference(row):
  """Referencia SAHA 2024 para PAS aórtica central según edad, sexo y calibración.

  Devuelve mediana, DE, P90, P95, z-score y percentil estimado. La decisión
  diagnóstica es binaria: CON HIPERTENSIÓN CENTRAL si PAS central >= P90 y
  SIN HIPERTENSIÓN CENTRAL si PAS central < P90. No se generan categorías ambiguas.
  """
  sex = safe_text(row.get("sexo", "M")).upper()[:1]
  if sex not in ("M", "F"):
    sex = "M"
  age_group = saha_age_group(row.get("edad"))
  method = safe_text(row.get("metodo_calibracion_pac", "SD_PAOC")).upper()
  if method not in ("C_PAOC", "SD_PAOC"):
    method = "SD_PAOC"
  if not age_group or sex not in SAHA_AOSBP_REF or age_group not in SAHA_AOSBP_REF[sex]:
    return {"ok": False, "motivo": "edad o sexo no disponibles para aplicar tabla SAHA"}
  median, sd = SAHA_AOSBP_REF[sex][age_group][method]
  pas_c = to_float(row.get("pas_central"))
  if np.isnan(pas_c):
    return {"ok": False, "motivo": "PAS central no disponible"}
  z = (pas_c - median) / sd if sd > 0 else np.nan
  pct = _norm_cdf(z) * 100 if not np.isnan(z) else np.nan
  p90 = median + 1.282 * sd
  p95 = median + 1.645 * sd
  p97 = median + 1.960 * sd
  if pas_c >= p90:
    categoria = "CON HIPERTENSIÓN CENTRAL ajustada por edad, sexo y calibración"
    severidad = "con hipertensión central"
    alterada = True
  else:
    categoria = "SIN HIPERTENSIÓN CENTRAL ajustada por edad, sexo y calibración"
    severidad = "sin hipertensión central"
    alterada = False
  return {
    "ok": True, "sexo": sex, "edad_grupo": age_group, "metodo": method,
    "metodo_label": SAHA_CALIBRATION_LABELS[method], "pas_central": pas_c,
    "mediana": median, "de": sd, "p90": p90, "p95": p95, "p97": p97,
    "z": z, "percentil": pct, "categoria": categoria, "severidad": severidad,
    "alterada": alterada,
  }

def central_hypertension_status(row):
  """Diagnóstico binario de hipertensión central.

  Prioridad diagnóstica:
  1) Tabla SAHA 2024 por edad, sexo y método de calibración: CON HIPERTENSIÓN CENTRAL si PASc >= P90.
  2) Respaldo operativo si no hay datos para SAHA: CON HIPERTENSIÓN CENTRAL si PASc >= 130 mmHg.
  """
  pas_c = to_float(row.get("pas_central"))
  if np.isnan(pas_c):
    return {
      "ok": False,
      "tiene_hta_central": False,
      "diagnostico": "Hipertensión central no clasificable por falta de PAS central.",
      "diagnostico_breve": "Hipertensión central no clasificable.",
      "criterio": "PAS central no disponible",
      "umbral": np.nan,
      "pas_central": pas_c,
    }

  saha_ref = get_saha_central_sbp_reference(row)
  if saha_ref.get("ok"):
    umbral = saha_ref.get("p90", np.nan)
    hta = pas_c >= umbral
    criterio = (
      f"criterio SAHA 2024 ajustado por edad/sexo/calibración: PAS central >= P90 "
      f"({umbral:.1f} mmHg)"
    )
    detalle = (
      f"PAS central {pas_c:.0f} mmHg; P90 SAHA {umbral:.1f} mmHg; "
      f"{saha_ref['sexo']}, {saha_ref['edad_grupo']} años, {saha_ref['metodo']}; "
      f"z {saha_ref['z']:.2f}; percentil {saha_ref['percentil']:.0f}."
    )
  else:
    umbral = 130.0
    hta = pas_c >= umbral
    criterio = "criterio operativo de respaldo: PAS central >= 130 mmHg"
    detalle = f"PAS central {pas_c:.0f} mmHg; umbral operativo {umbral:.0f} mmHg."

  diagnostico_breve = "CON HIPERTENSIÓN CENTRAL." if hta else "SIN HIPERTENSIÓN CENTRAL."
  return {
    "ok": True,
    "tiene_hta_central": bool(hta),
    "diagnostico": f"{diagnostico_breve} {detalle}",
    "diagnostico_breve": diagnostico_breve,
    "criterio": criterio,
    "umbral": umbral,
    "pas_central": pas_c,
  }


def format_saha_central_htn(row, compact=False):
  status = central_hypertension_status(row)
  if not status.get("ok"):
    return status.get("diagnostico", "Hipertensión central: no clasificable.")
  txt = f"Diagnóstico de presión central: {status['diagnostico']} Criterio utilizado: {status['criterio']}."
  if not compact:
    txt += " La decisión del informe es binaria: CON HIPERTENSIÓN CENTRAL o SIN HIPERTENSIÓN CENTRAL."
  return txt


def _markdown_bold_conclusions(text):
  """Resalta conclusiones finales en la vista Streamlit."""
  txt = safe_text(text)
  patterns = ["Conclusión breve:", "Conclusión diagnóstica:", "Conclusión integrada:", "Conclusión final:"]
  for pat in patterns:
    txt = txt.replace(pat, f"**{pat}")
  # Cierra la negrita al final del párrafo si se abrió alguna etiqueta markdown.
  if any(f"**{pat}" in txt for pat in patterns) and not txt.endswith("**"):
    txt += "**"
  return txt


def _pdf_bold_conclusions(text):
  """Escapa texto dinámico y permite negrita solo en frases de conclusión controladas."""
  txt = pdf_text(text)
  keys = ["Conclusión breve:", "Conclusión diagnóstica:", "Conclusión integrada:", "Conclusión final:"]
  for key in keys:
    txt = txt.replace(pdf_text(key), f"<b>{pdf_text(key)}")
  if any(f"<b>{pdf_text(k)}" in txt for k in keys) and "</b>" not in txt:
    txt += "</b>"
  elif any(f"<b>{pdf_text(k)}" in txt for k in keys):
    # Asegurar cierre final aunque hubiese otro tag previo.
    open_tags = txt.count("<b>")
    close_tags = txt.count("</b>")
    if open_tags > close_tags:
      txt += "</b>"
  return txt

def safe_text(x):
  if x is None:
    return ""
  return str(x).replace("\x00", "").strip()

def pdf_text(x):
  """Texto seguro para ReportLab Paragraph.

  ReportLab interpreta <...> como etiquetas XML. Si una conclusión clínica
  contiene comparadores como <, >, >=, <= o símbolos & provenientes de
  percentiles/puntos de corte, el PDF puede fallar con paraparser.
  Esta función escapa todo texto dinámico antes de pasarlo a Paragraph.
  """
  txt = safe_text(x)
  return (txt.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;"))

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
  # Tolera signo menos Unicode y números entre paréntesis, habituales en PDFs PAC.
  txt = safe_text(txt).replace("−", "-").replace("–", "-").replace("—", "-")
  return [to_float(x) for x in re.findall(r"[-+]?\d+(?:[\.,]\d+)?", txt)]


def _normalize_central_metric_text(txt):
  """Normaliza etiquetas PAC partidas por el motor PDF/OCR.

  Ejemplos recuperados: A P C, A.P.C., A-P-C, I Au, I A u, A u, R V S E.
  La normalización de Au evita convertir IAu en Au.
  """
  t = safe_text(txt)
  t = t.replace("−", "-").replace("–", "-").replace("—", "-")
  # Etiquetas compuestas: primero las más largas para evitar colisiones.
  t = re.sub(r"(?i)(?<![A-Za-z])A\s*[.\-]?\s*P\s*[.\-]?\s*C\.?(?![A-Za-z])", "APC", t)
  t = re.sub(r"(?i)(?<![A-Za-z])I\s*[.\-]?\s*A\s*[.\-]?\s*u(?![A-Za-z])", "IAu", t)
  t = re.sub(r"(?i)(?<![A-Za-z])R\s*[.\-]?\s*V\s*[.\-]?\s*S\s*[.\-]?\s*E(?![A-Za-z])", "RVSE", t)
  t = re.sub(r"(?i)(?<![A-Za-z])P\s*[.\-]?\s*E(?![A-Za-z])", "PE", t)
  # Au aislado / partido. El lookbehind impide tomar la parte 'Au' de IAu.
  t = re.sub(r"(?i)(?<![A-Za-z])A\s*[.\-]?\s*u\.?(?![A-Za-z])", "Au", t)
  return _collapse_spaces(t)


def _extract_central_metric_value(txt, metric):
  """Extrae parámetros centrales con separación ESTRICTA entre Au e IAu.

  Reglas críticas:
  - Au = presión de aumentación en mmHg. Nunca se obtiene desde IAu/AIx.
  - IAu = índice de aumentación en %. Nunca se usa como respaldo de Au.
  - APC = amplificación periférico-central en mmHg. El valor principal es la
    diferencia periférico-central informada por el equipo (p. ej. 16 mmHg).
    La cifra entre paréntesis (p. ej. 1,11) es una relación secundaria y nunca
    debe reemplazar al APC en mmHg.
  - Se toleran etiquetas partidas por PDF/OCR: A u, I Au, A P C, A.P.C.
  """
  t = _normalize_central_metric_text(txt)
  metric = str(metric or "").lower()
  num = r"([-+]?\d+(?:[\.,]\d+)?)"

  # Blindaje absoluto: para buscar Au se enmascara IAu completo antes de cualquier regex.
  if metric == "au":
    t_search = re.sub(r"(?i)\bIAu\b", "__INDICE_AUMENTACION__", t)
  else:
    t_search = t

  patterns = {
    "apc": [
      # Prioridad máxima: APC principal seguido/precedido por unidad mmHg.
      rf"\bAPC\b\s*(?::|=)?\s*{num}\s*mmHg\b",
      rf"\bAPC\b\s*(?:mmHg)\s*(?::|=)?\s*{num}",
      rf"Amplificaci[oó]n\s+(?:de\s+)?(?:Presi[oó]n\s+)?Perif[eé]rico[- ]?Central\s*(?:\(mmHg\)|mmHg)?\s*(?::|=)?\s*{num}",
      rf"Amplificaci[oó]n\s+Perif[eé]rica[- ]?Central\s*(?:\(mmHg\)|mmHg)?\s*(?::|=)?\s*{num}",
      # Respaldo: APC + número; la consistencia final con PAS radial-central decide.
      rf"\bAPC\b\s*(?::|=)?\s*{num}(?!\s*[%])",
    ],
    "au": [
      # Prioridad máxima: Au independiente y unidad mmHg.
      rf"(?<![A-Za-z])Au(?![A-Za-z])\s*(?:\(?\s*)?mmHg\s*(?::|=)?\s*[\(\[]?\s*{num}",
      # Respaldo: Au independiente seguido directamente del valor.
      rf"(?<![A-Za-z])Au(?![A-Za-z])\s*(?::|=)\s*[\(\[]?\s*{num}",
      rf"(?:Presi[oó]n\s+de\s+Aumentaci[oó]n|Augmentation\s+Pressure)\s*(?:mmHg)?\s*(?::|=)?\s*[\(\[]?\s*{num}",
      rf"Aumentaci[oó]n\s+A[oó]rtica(?:\s+Central)?\s*(?:mmHg)?\s*(?::|=)?\s*[\(\[]?\s*{num}",
    ],
    "iau": [
      rf"\bIAu\b\s*(?:%|porcentaje)?\s*(?::|=)?\s*[\(\[]?\s*{num}",
      rf"(?:[ÍI]ndice\s+de\s+Aumentaci[oó]n|Augmentation\s+Index|AIx)\s*(?:%)?\s*(?::|=)?\s*[\(\[]?\s*{num}",
    ],
    "rvse": [rf"\bRVSE\b\s*(?:%)?\s*(?::|=)?\s*[\(\[]?\s*{num}"],
    "pe": [rf"\bPE\b\s*(?:%)?\s*(?::|=)?\s*[\(\[]?\s*{num}"],
    "pas_central": [rf"\bPAS\b\s*(?:mmHg)?\s*(?::|=)?\s*[\(\[]?\s*{num}"],
    "pp_central": [rf"\bPP\b\s*(?:mmHg)?\s*(?::|=)?\s*[\(\[]?\s*{num}"],
  }

  candidates = []
  for priority, pat in enumerate(patterns.get(metric, [])):
    for m in re.finditer(pat, t_search, flags=re.I):
      v = to_float(m.group(1))
      if np.isnan(v):
        continue
      candidates.append((priority, m.start(), v))

  if not candidates:
    return np.nan

  # Menor priority = patrón más específico. Entre iguales, la última aparición suele
  # corresponder a la tabla central real y no a una leyenda previa.
  best_priority = min(c[0] for c in candidates)
  selected = [c for c in candidates if c[0] == best_priority][-1]
  value = float(selected[2])

  if metric == "apc":
    # APC del informe Exxer = amplificación periférico-central en mmHg.
    # NO convertir a relación y NO usar la cifra parentética (p. ej. 1,11).
    if not (-100 <= value <= 150):
      return np.nan
  elif metric == "au":
    # Rango clínico amplio pero propio de presión de aumentación, no de porcentaje.
    if not (-80 <= value <= 100):
      return np.nan
  elif metric == "iau":
    if not (-100 <= value <= 150):
      return np.nan
  return value


def _metric_label_regex(metric):
  """Regex de etiqueta estricta para impedir cruces Au/IAu."""
  metric = str(metric or "").lower()
  if metric == "au":
    return r"(?<![A-Za-z])Au(?![A-Za-z])"
  if metric == "iau":
    return r"\bIAu\b"
  if metric == "apc":
    return r"\bAPC\b"
  return rf"\b{re.escape(metric)}\b"


def _normalize_apc_value(v):
  """Valida APC como amplificación periférico-central en mmHg.

  Importante: el valor entre paréntesis del equipo (p. ej. 1,11) es la razón
  PAS periférica/PAS central y no debe almacenarse en ``apc``.
  """
  x = to_float(v)
  if np.isnan(x):
    return np.nan
  return x if -100 <= x <= 150 else np.nan


def _derived_apc_mmhg(data):
  """APC esperado en mmHg = PAS radial - PAS central."""
  pr = to_float((data or {}).get("pas_radial"))
  pc = to_float((data or {}).get("pas_central"))
  if np.isnan(pr) or np.isnan(pc):
    return np.nan
  d = pr - pc
  return float(d) if -100 <= d <= 150 else np.nan


def _derived_apc_ratio(data):
  """Relación secundaria mostrada entre paréntesis por el equipo."""
  pr = to_float((data or {}).get("pas_radial"))
  pc = to_float((data or {}).get("pas_central"))
  if np.isnan(pr) or np.isnan(pc) or pc == 0:
    return np.nan
  r = pr / pc
  return float(r) if 0.5 <= r <= 2.5 else np.nan


def _repair_apc_semantics(data):
  """Impide confundir APC mmHg con la relación parentética.

  Cuando existen PAS radial y central reales, la diferencia es el control de
  consistencia del APC. Si el valor importado parece una razón (≈0,5-2,5) o no
  concuerda con la diferencia, se reemplaza por PAS radial - PAS central.
  """
  expected = _derived_apc_mmhg(data)
  raw = to_float((data or {}).get("apc"))
  ratio = _derived_apc_ratio(data)
  if not np.isnan(ratio):
    data["apc_ratio"] = ratio
  if not np.isnan(expected):
    # La captura real muestra APC 16 mmHg y, separado, (1,11).
    ratio_like = (not np.isnan(raw) and 0.5 <= abs(raw) <= 2.5 and abs(expected) >= 4)
    discordant = (not np.isnan(raw) and abs(raw - expected) > max(3.0, 0.20 * max(abs(expected), 1.0)))
    if np.isnan(raw) or ratio_like or discordant:
      data["apc"] = expected
    else:
      data["apc"] = raw
  elif not np.isnan(raw) and -100 <= raw <= 150:
    data["apc"] = raw
  else:
    data["apc"] = np.nan
  return data


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

  # 4) Parámetros hemodinámicos centrales / aumentaciones.
  # Normalización reforzada para PDFs que separan letras: A P C, A.P.C., A u, I Au.
  central_map = {
    "PAS": "pas_central", "PP": "pp_central", "Au": "au", "IAu": "iau",
    "RVSE": "rvse", "PE": "pe", "APC": "apc"
  }
  for ln in lines:
    txt = _normalize_central_metric_text(ln["text"])
    for lab, key in central_map.items():
      # La etiqueta puede no quedar al inicio visual por columnas o unidades.
      label_pat = r"(?<!I)\bAu\b" if lab == "Au" else rf"\b{re.escape(lab)}\b"
      if re.search(label_pat, txt, flags=re.I):
        v = _extract_central_metric_value(txt, key)
        if not np.isnan(v):
          out[key] = v
        else:
          # Respaldo por números de la misma fila: el primero es el valor principal,
          # los posteriores suelen ser tolerancia +/-.
          vals = _numeric_values_in_text(txt)
          if vals:
            if key == "apc":
              # APC primario es mmHg. Si PAS radial/central ya están disponibles,
              # escoger el candidato más cercano a su diferencia; no la razón 1,11.
              plaus = [x for x in vals if -100 <= x <= 150]
              expected = _derived_apc_mmhg(out)
              if plaus:
                if not np.isnan(expected):
                  out[key] = min(plaus, key=lambda x: abs(x - expected))
                else:
                  # Preferir magnitud en mmHg sobre candidatos tipo razón.
                  non_ratio = [x for x in plaus if abs(x) > 3.5]
                  out[key] = (non_ratio or plaus)[0]
            else:
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
    token = str(txt).strip().replace("−", "-").replace("–", "-").replace("—", "-")
    # El equipo puede exportar APC como '(1.08)' o '[1,08]'.
    token = token.strip("()[]{}:;=")
    if re.fullmatch(r"[-+]?\d+(?:[\.,]\d+)?", token):
      vals.append((x0, to_float(token)))
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
  """Extrae tablas PAC por coordenadas con Au/IAu estrictamente separados.

  Au se toma solo desde su propia etiqueta independiente y, preferentemente, su fila
  con unidad mmHg. IAu se toma desde IAu/AIx y %. APC se busca en la misma fila o en
  filas vecinas cercanas, porque algunos PDFs separan etiqueta y valor.
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
      param_y = row["y"]
      break

  def in_param_block(row):
    if param_y is not None:
      return param_y < row["y"] < param_y + 0.34*H
    return row["y"] > 0.50*H

  def row_value_after_label(row, metric):
    """Primer número a la derecha de la etiqueta exacta en esa fila."""
    text_norm = _normalize_central_metric_text(row["text"])
    label_pat = _metric_label_regex(metric)
    if not re.search(label_pat, text_norm, re.I):
      return np.nan

    # Si buscamos Au, una fila que solo contiene IAu no es candidata.
    if metric == "au":
      masked = re.sub(r"(?i)\bIAu\b", "", text_norm)
      if not re.search(_metric_label_regex("au"), masked, re.I):
        return np.nan

    # Intento textual exacto primero.
    v = _extract_central_metric_value(text_norm, metric)
    if not np.isnan(v):
      return v

    # Ubicar x final de la etiqueta usando secuencias de tokens.
    words = row["words"]
    label_x1 = None
    token_texts = [str(w[4]) for w in words]
    norm_tokens = [_normalize_central_metric_text(t) for t in token_texts]
    for i, wt in enumerate(norm_tokens):
      if metric == "au" and re.fullmatch(r"Au", wt, re.I):
        # No aceptar Au si está precedido inmediatamente por token I.
        prev = norm_tokens[i-1] if i > 0 else ""
        if re.fullmatch(r"I", prev, re.I):
          continue
        label_x1 = words[i][2]; break
      if metric == "iau" and re.fullmatch(r"IAu", wt, re.I):
        label_x1 = words[i][2]; break
      if metric == "apc" and re.fullmatch(r"APC", wt, re.I):
        label_x1 = words[i][2]; break
      if metric == "pe" and re.fullmatch(r"PE", wt, re.I):
        label_x1 = words[i][2]; break
      if metric == "rvse" and re.fullmatch(r"RVSE", wt, re.I):
        label_x1 = words[i][2]; break

    # Etiquetas partidas en varios tokens: A u / I A u / A P C.
    if label_x1 is None:
      compact = [re.sub(r"[^A-Za-z]", "", t) for t in token_texts]
      for i in range(len(compact)):
        if metric == "au" and i+1 < len(compact) and compact[i].upper() == "A" and compact[i+1].lower() == "u":
          if i > 0 and compact[i-1].upper() == "I":
            continue
          label_x1 = words[i+1][2]; break
        if metric == "iau" and i+2 < len(compact) and compact[i].upper() == "I" and compact[i+1].upper() == "A" and compact[i+2].lower() == "u":
          label_x1 = words[i+2][2]; break
        if metric == "apc" and i+2 < len(compact) and compact[i].upper() == "A" and compact[i+1].upper() == "P" and compact[i+2].upper() == "C":
          label_x1 = words[i+2][2]; break
        if metric == "pe" and i+1 < len(compact) and compact[i].upper() == "P" and compact[i+1].upper() == "E":
          label_x1 = words[i+1][2]; break
        if metric == "rvse" and i+3 < len(compact) and compact[i].upper() == "R" and compact[i+1].upper() == "V" and compact[i+2].upper() == "S" and compact[i+3].upper() == "E":
          label_x1 = words[i+3][2]; break

    if label_x1 is None:
      return np.nan
    # Limitar a números próximos a la derecha de la etiqueta para no capturar
    # ticks de gráficos/zonas de referencia de otras columnas.
    nums = _numbers_from_row_words(words, x_min=label_x1, x_max=label_x1 + 0.28*W)
    if not nums:
      return np.nan
    values = [v for _, v in nums]
    if metric == "apc":
      expected = _derived_apc_mmhg(out)
      plaus = [v for v in values if -100 <= v <= 150]
      if not plaus:
        return np.nan
      if not np.isnan(expected):
        return min(plaus, key=lambda x: abs(x - expected))
      non_ratio = [v for v in plaus if abs(v) > 3.5]
      return (non_ratio or plaus)[0]
    if metric == "pe":
      plaus = [v for v in values if 5 <= v <= 80]
      return plaus[0] if plaus else np.nan
    if metric == "rvse":
      plaus = [v for v in values if 0 <= v <= 300]
      return plaus[0] if plaus else np.nan
    return values[0]

  # Extraer Au e IAu en pasadas separadas: jamás una sirve de respaldo de la otra.
  for row in rows:
    if not in_param_block(row):
      continue
    v = row_value_after_label(row, "au")
    if not np.isnan(v):
      out["au"] = v
      break

  for row in rows:
    if not in_param_block(row):
      continue
    v = row_value_after_label(row, "iau")
    if not np.isnan(v):
      out["iau"] = v
      break

  # Otros parámetros centrales: extracción geométrica estricta por la fila propia.
  # Esto corrige PE 32,0% que antes podía quedar vacío/0 por mezcla con el gráfico.
  for metric, key in [("rvse","rvse"), ("pe","pe")]:
    for row in rows:
      if not in_param_block(row):
        continue
      v = row_value_after_label(row, metric)
      if not np.isnan(v):
        out[key] = v
        break

  # APC: misma fila y luego filas vecinas cercanas / misma columna.
  for i, row in enumerate(rows):
    rowtxt = _normalize_central_metric_text(row["text"])
    if not re.search(_metric_label_regex("apc"), rowtxt, re.I):
      continue
    v = row_value_after_label(row, "apc")
    if not np.isnan(v):
      out["apc"] = v
      break

    # x aproximada de la etiqueta APC.
    label_x = min((w[0] for w in row["words"]), default=0)
    for w in row["words"]:
      if re.search(r"(?i)APC|A\.?P\.?C", str(w[4])):
        label_x = w[2]
        break

    neighbor_candidates = []
    for j in range(max(0, i-2), min(len(rows), i+3)):
      if j == i: continue
      r2 = rows[j]
      if abs(r2["y"] - row["y"]) > 0.06*H:
        continue
      for x, val in _numbers_from_row_words(r2["words"]):
        if x >= label_x - 0.05*W:
          nv = _normalize_apc_value(val)
          if not np.isnan(nv):
            expected = _derived_apc_mmhg(out)
            # Penalizar candidatos de tipo relación (1,11) si se espera una diferencia mmHg.
            semantic_penalty = 0.0
            if not np.isnan(expected):
              semantic_penalty = abs(nv - expected) / max(abs(expected), 1.0)
            elif abs(nv) <= 2.5:
              semantic_penalty = 2.0
            neighbor_candidates.append((abs(r2["y"]-row["y"]) + abs(x-label_x)/max(W,1) + semantic_penalty, nv))
    if neighbor_candidates:
      neighbor_candidates.sort(key=lambda z: z[0])
      out["apc"] = neighbor_candidates[0][1]
      break

  return out


def _extract_central_params_global(flat_text):
  """Extracción global estricta de parámetros centrales.

  Au e IAu se procesan por canales independientes. El texto IAu se enmascara antes
  de cualquier búsqueda de Au para impedir que el índice sea importado como presión.
  APC se interpreta como amplificación periférico-central en mmHg; la relación parentética se mantiene separada.
  """
  out = {}
  txt = _normalize_central_metric_text(flat_text)

  # IAu primero, con etiqueta propia.
  v_iau = _extract_central_metric_value(txt, "iau")
  if not np.isnan(v_iau):
    out["iau"] = v_iau

  # Au sobre texto con IAu enmascarado de forma explícita.
  txt_au = re.sub(r"(?i)\bIAu\b", "__INDICE_AUMENTACION__", txt)
  v_au = _extract_central_metric_value(txt_au, "au")
  if not np.isnan(v_au):
    out["au"] = v_au

  for key in ("apc", "rvse", "pe", "pas_central", "pp_central"):
    v = _extract_central_metric_value(txt, key)
    if not np.isnan(v):
      out[key] = v

  # APC: ventana corta hasta la siguiente etiqueta conocida. El valor principal
  # es mmHg; la relación parentética se descarta/queda separada.
  if "apc" not in out:
    for m in re.finditer(r"\bAPC\b", txt, re.I):
      tail = txt[m.end():m.end()+120]
      tail = re.split(r"\b(?:PAS|PP|Au|IAu|RVSE|PE|FC|PAD|PAM)\b", tail, maxsplit=1, flags=re.I)[0]
      nums = _numeric_values_in_text(tail)
      for x in nums:
        nv = _normalize_apc_value(x)
        if not np.isnan(nv):
          out["apc"] = nv
          break
      if "apc" in out:
        break

  # Au: respaldo solo tras etiqueta Au independiente y antes de la próxima etiqueta.
  if "au" not in out:
    masked = re.sub(r"(?i)\bIAu\b", "__INDICE_AUMENTACION__", txt)
    for m in re.finditer(r"(?<![A-Za-z])Au(?![A-Za-z])", masked, re.I):
      tail = masked[m.end():m.end()+80]
      tail = re.split(r"\b(?:IAu|RVSE|PE|APC|PAS|PP|FC|PAD|PAM)\b", tail, maxsplit=1, flags=re.I)[0]
      nums = _numeric_values_in_text(tail)
      if nums:
        x = nums[0]
        if -80 <= x <= 100:
          out["au"] = x
          break
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

    # 3) Aumentaciones y parámetros centrales desde texto global.
    # Para Au/IAu/APC, las coordenadas tienen prioridad y el texto global solo completa faltantes.
    flat = _collapse_spaces(fallback_text or "")
    global_params = _extract_central_params_global(flat)
    for k, v in global_params.items():
      if np.isnan(to_float(v)):
        continue
      if k in ("au", "iau", "apc") and not np.isnan(to_float(coord_vars.get(k, np.nan))):
        continue
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
    data = _repair_apc_semantics(data)
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

  Incluye PAS, PP, Au, IAu, RVSE, PE y APC. Usa normalización de etiquetas
  partidas y un respaldo global para APC/Au fuera de la sección central.
  """
  text_norm = _normalize_central_metric_text(flat_text)
  msec = re.search(
    r"Par[aá]metros\s+hemodin[aá]micos\s+centrales(.*?)(?:Conclusiones|Conclusi[oó]n|PAS\s+Presi[oó]n|Pulso|$)",
    text_norm, re.I
  )
  sec = msec.group(1) if msec else text_norm
  sec = _normalize_central_metric_text(sec)

  mapping = {}
  for key in ("pas_central", "pp_central", "au", "iau", "rvse", "pe", "apc"):
    v = _extract_central_metric_value(sec, key)
    if np.isnan(v):
      # APC suele estar fuera de la tabla central; Au puede quedar mezclado por columnas.
      v = _extract_central_metric_value(text_norm, key)
    mapping[key] = v

  # Respaldo posicional de filas centrales, incluyendo etiquetas partidas ya normalizadas.
  positional = re.findall(
    r"(?<!I)\b(Au)\b\s*(?:mmHg)?\s*(?::|=)?\s*[\(\[]?\s*([-+]?\d+(?:[\.,]\d+)?)"
    r"|\b(IAu|RVSE|PE|APC|PAS|PP)\b\s*(?:mmHg|%)?\s*(?::|=)?\s*[\(\[]?\s*([-+]?\d+(?:[\.,]\d+)?)",
    sec, flags=re.I
  )
  for au_lab, au_num, lab2, num2 in positional:
    lab = au_lab or lab2
    num = au_num or num2
    key = {"pas":"pas_central", "pp":"pp_central", "au":"au", "iau":"iau", "rvse":"rvse", "pe":"pe", "apc":"apc"}.get(lab.lower())
    if key and np.isnan(to_float(mapping.get(key, np.nan))):
      mapping[key] = to_float(num)

  for k, v in mapping.items():
    if not np.isnan(to_float(v)):
      data[k] = to_float(v)


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

  # Separación final Au / IAu: nunca completar uno desde el otro.
  # Si solo existe IAu, Au queda vacío; si solo existe Au, IAu queda vacío.
  # No se copian ni se igualan valores entre ambas métricas.
  au_v = to_float(data.get("au"))
  iau_v = to_float(data.get("iau"))
  if not np.isnan(au_v):
    data["au"] = au_v
  if not np.isnan(iau_v):
    data["iau"] = iau_v

  # APC = amplificación periférico-central en mmHg. Controlar contra PAS radial-central
  # para no importar la relación parentética (1,11) ni números del gráfico.
  data = _repair_apc_semantics(data)

  # Rango fisiológico/administrativo: si falla, dejar editable como vacío en la interfaz.
  for k, lo, hi in [
    ("edad", 1, 120), ("peso", 20, 250), ("altura", 80, 230), ("imc", 10, 80), ("sc", 0.5, 3.5),
    ("pas_radial", 50, 260), ("pad_radial", 30, 160), ("pam_radial", 40, 200), ("pp_radial", 10, 120),
    ("pas_central", 50, 260), ("pad_central", 30, 160), ("pam_central", 40, 200), ("pp_central", 10, 120),
    ("fc", 25, 180), ("au", -80, 100), ("iau", -80, 100), ("rvse", 0, 300), ("pe", 5, 80), ("apc", -100, 150)
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
  """Resumen de presión central para pantalla/PDF basado en el estado canónico único."""
  cSBP = to_float(row.get("pas_central")); pSBP = to_float(row.get("pas_radial")); pDBP = to_float(row.get("pad_radial"))
  cat = brachial_bp_category(pSBP, pDBP)
  state = build_canonical_diagnostic_state(row, {}, None)
  d = state["domains"]
  hta = d["hta_central"]; pulse = d["carga_pulsatil"]; ppa_d = d["amplificacion"]; aug = d["aumentacion"]

  ref = hta.get("threshold", np.nan)
  amp_sbp = pSBP - cSBP if not np.isnan(pSBP) and not np.isnan(cSBP) else np.nan
  ppa = ppa_d.get("value", np.nan)

  if hta["status"] == "alterada":
    dx = "CON HIPERTENSIÓN CENTRAL. " + hta.get("detail", "")
  elif hta["status"] == "normal":
    dx = "SIN HIPERTENSIÓN CENTRAL. " + hta.get("detail", "")
  else:
    dx = "Hipertensión central no clasificable por datos insuficientes."

  risk = [
    hta["short"],
    pulse["short"],
    ppa_d["short"],
    aug["short"],
  ]
  return dx, cat, ref, amp_sbp, ppa, "; ".join(risk)

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
  amp_load = 0.5 if np.isnan(ppa) else float(np.clip((1.70 - ppa) / 0.70, 0, 1)) # menor PPA = más carga central

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

def harmonic_analysis(wave_df, max_display_harmonics=10):
  """Análisis armónico FFT de la onda central real.

  La tabla devuelta conserva H1-H10 (o menos si la señal no alcanza) para una
  visualización compacta. Las métricas globales HD y H4+ se calculan, en cambio,
  sobre TODO el espectro positivo disponible (excluido DC) y se guardan en
  ``DataFrame.attrs``. Así se evita llamar H4+ o HD a una suma truncada en H10.

  Además de frecuencia, amplitud y energía relativa, se conserva la fase de cada
  componente mostrada para reconstrucción/animación sincronizada con la curva real.
  """
  if wave_df is None or len(wave_df) < 4:
    return pd.DataFrame(columns=["armónico", "frecuencia_hz", "amplitud", "fase_rad", "energia_relativa_%"])

  y = pd.to_numeric(wave_df.iloc[:, 1], errors="coerce").to_numpy(dtype=float)
  t = pd.to_numeric(wave_df.iloc[:, 0], errors="coerce").to_numpy(dtype=float)
  ok = np.isfinite(y) & np.isfinite(t)
  y, t = y[ok], t[ok]
  if len(y) < 4:
    return pd.DataFrame(columns=["armónico", "frecuencia_hz", "amplitud", "fase_rad", "energia_relativa_%"])

  order = np.argsort(t)
  y, t = y[order], t[order]
  dt_ms = np.diff(t)
  dt_ms = dt_ms[np.isfinite(dt_ms) & (dt_ms > 0)]
  if len(dt_ms) == 0:
    return pd.DataFrame(columns=["armónico", "frecuencia_hz", "amplitud", "fase_rad", "energia_relativa_%"])
  dt_s = float(np.median(dt_ms)) / 1000.0
  if not np.isfinite(dt_s) or dt_s <= 0:
    return pd.DataFrame(columns=["armónico", "frecuencia_hz", "amplitud", "fase_rad", "energia_relativa_%"])

  y0 = y - np.mean(y)
  fft = np.fft.rfft(y0)
  amp = np.abs(fft) / len(y0) * 2.0
  freq = np.fft.rfftfreq(len(y0), d=dt_s)
  phase = np.angle(fft)

  # Espectro positivo sin DC: H1 corresponde al primer bin positivo del ciclo.
  pos_amp = amp[1:]
  pos_freq = freq[1:]
  pos_phase = phase[1:]
  if len(pos_amp) == 0:
    return pd.DataFrame(columns=["armónico", "frecuencia_hz", "amplitud", "fase_rad", "energia_relativa_%"])

  pos_energy = np.square(pos_amp)
  total_energy = float(np.nansum(pos_energy))
  if not np.isfinite(total_energy) or total_energy <= 0:
    total_energy = 1.0
  pos_energy_pct = pos_energy / total_energy * 100.0

  n_show = int(max(1, min(int(max_display_harmonics), len(pos_amp))))
  df = pd.DataFrame({
    "armónico": np.arange(1, n_show + 1, dtype=int),
    "frecuencia_hz": pos_freq[:n_show],
    "amplitud": pos_amp[:n_show],
    "fase_rad": pos_phase[:n_show],
    "energia_relativa_%": pos_energy_pct[:n_show],
  })

  fundamental_energy = float(pos_energy[0]) if len(pos_energy) > 0 else np.nan
  higher_energy = float(np.nansum(pos_energy[1:])) if len(pos_energy) > 1 else 0.0
  hd_ratio_full = higher_energy / fundamental_energy if np.isfinite(fundamental_energy) and fundamental_energy > 0 else np.nan
  h4plus_full = float(np.nansum(pos_energy_pct[3:])) if len(pos_energy_pct) > 3 else 0.0
  dom_i = int(np.nanargmax(pos_energy_pct)) if np.any(np.isfinite(pos_energy_pct)) else 0

  # Metadatos globales: no dependen del truncado visual H1-H10.
  df.attrs.update({
    "full_spectrum_harmonics": int(len(pos_amp)),
    "frequency_resolution_hz": float(pos_freq[0]) if len(pos_freq) else np.nan,
    "hd_ratio_full": float(hd_ratio_full) if np.isfinite(hd_ratio_full) else np.nan,
    "hd_percent_full": float(hd_ratio_full * 100.0) if np.isfinite(hd_ratio_full) else np.nan,
    "h4plus_energy_percent_full": h4plus_full,
    "dominant_frequency_hz_full": float(pos_freq[dom_i]) if len(pos_freq) > dom_i else np.nan,
  })
  return df



# ------------------------------------------------------------------
# BASE DE CONOCIMIENTO ARMÓNICA / LONGITUDINAL
# ------------------------------------------------------------------
# HD se implementa como relación de energía espectral por encima de la fundamental
# respecto de la energía de la fundamental:
#     HD = sum(A_n^2, n>=2) / A_1^2
# Se informa como razón y porcentaje. NO se aplica un punto de corte clínico universal.
HARMONIC_RESONANCE_REFERENCE_HZ = 3.5  # referencia mecanística/descriptiva, no umbral diagnóstico
HARMONIC_LONGITUDINAL_MDC_PCT = {
  "C1-C2": 5.0,
  "C3-C5": 10.0,
  "C6-C8": 20.0,
  "C9-C11": 40.0,
}


def harmonic_distortion_metrics(hdf):
  """Métricas armónicas continuas y HD energético.

  HD y H4+ usan el espectro positivo completo cuando ``harmonic_analysis`` dejó
  esos metadatos disponibles. La tabla/gráfico puede seguir mostrando solo H1-H10.
  No se convierten estas métricas en diagnóstico de rigidez ni en fenotipo mediante
  puntos de corte universales.
  """
  out = {
    "ok": False,
    "hd_ratio": np.nan,
    "hd_percent": np.nan,
    "h1_energy_percent": np.nan,
    "h2_energy_percent": np.nan,
    "h4plus_energy_percent": np.nan,
    "dominant_frequency_hz": np.nan,
    "distance_to_3_5_hz": np.nan,
    "spectrum_harmonics_analyzed": 0,
    "classification": "métricas continuas cuantificadas / sin corte clínico universal",
  }
  try:
    if hdf is None or len(hdf) == 0:
      return out
    amp = pd.to_numeric(hdf.get("amplitud"), errors="coerce").to_numpy(dtype=float)
    ene = pd.to_numeric(hdf.get("energia_relativa_%"), errors="coerce").to_numpy(dtype=float)
    fre = pd.to_numeric(hdf.get("frecuencia_hz"), errors="coerce").to_numpy(dtype=float)
    if len(amp) < 1 or not np.isfinite(amp[0]) or abs(amp[0]) <= 1e-12:
      return out

    attrs = getattr(hdf, "attrs", {}) or {}
    hd_ratio = to_float(attrs.get("hd_ratio_full"))
    if np.isnan(hd_ratio):
      fundamental_energy = float(amp[0] ** 2)
      higher_energy = float(np.nansum(np.square(amp[1:]))) if len(amp) > 1 else 0.0
      hd_ratio = higher_energy / fundamental_energy if fundamental_energy > 0 else np.nan

    h4plus = to_float(attrs.get("h4plus_energy_percent_full"))
    if np.isnan(h4plus):
      h4plus = float(np.nansum(ene[3:])) if len(ene) > 3 else 0.0

    dom_freq = to_float(attrs.get("dominant_frequency_hz_full"))
    if np.isnan(dom_freq):
      dom_i = int(np.nanargmax(ene)) if len(ene) and np.any(np.isfinite(ene)) else 0
      dom_freq = float(fre[dom_i]) if len(fre) > dom_i and np.isfinite(fre[dom_i]) else np.nan

    n_spec = attrs.get("full_spectrum_harmonics", len(amp))
    try:
      n_spec = int(n_spec)
    except Exception:
      n_spec = int(len(amp))

    out.update({
      "ok": True,
      "hd_ratio": hd_ratio,
      "hd_percent": hd_ratio * 100.0 if np.isfinite(hd_ratio) else np.nan,
      "h1_energy_percent": float(ene[0]) if len(ene) > 0 and np.isfinite(ene[0]) else np.nan,
      "h2_energy_percent": float(ene[1]) if len(ene) > 1 and np.isfinite(ene[1]) else np.nan,
      "h4plus_energy_percent": h4plus,
      "dominant_frequency_hz": dom_freq,
      "distance_to_3_5_hz": abs(dom_freq - HARMONIC_RESONANCE_REFERENCE_HZ) if np.isfinite(dom_freq) else np.nan,
      "spectrum_harmonics_analyzed": n_spec,
    })
    return out
  except Exception:
    return out


def longitudinal_harmonic_mdc_note():
  """Texto metodológico: umbrales propuestos solo para cambios longitudinales."""
  return (
    "Cambio mínimo detectable longitudinal propuesto (no diagnóstico transversal): "
    "C1-C2 >5%, C3-C5 >10%, C6-C8 >20% y C9-C11 >40%. "
    "Aplicable únicamente al comparar estudios repetidos con el mismo equipo, protocolo y calidad de señal."
  )


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
  p_model = pad + pf + pb # exactamente la curva real suavizada/calibrada

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
  # Auditoría de sincronización gráfico-animación sobre la misma serie Pf/Pb.
  sync = _wave_sync_landmarks(sep_df, metrics)
  metrics["t_pb_inicio_ms"] = sync.get("t_pb_onset_ms", np.nan)
  metrics["t_cruce_pf_pb_ms"] = sync.get("t_cross_ms", np.nan)
  return sep_df, metrics

def interpret_wave_separation(sep_metrics, row=None):
  """Interpreta separación de ondas con RM percentilar como criterio primario.

  Reglas de conocimiento:
  - RM: primaria, por percentiles de edad y método (P75/P90/P95).
  - RI: complementaria y matemáticamente derivada de Pf/Pb; no suma evidencia independiente.
  - Tref y Tfor/Tref: variables continuas de investigación, sin corte clínico universal fijo.
  """
  rm = sep_metrics.get("rm", np.nan)
  ri = sep_metrics.get("ri", np.nan)
  tref = sep_metrics.get("tref_ms", np.nan)
  ratio = sep_metrics.get("tfor_tref", np.nan)

  parts = []
  ref = get_rm_ri_reference(row or {}, sep_metrics)
  if ref.get("ok") and ref.get("rm_clasif", {}).get("ok"):
    c = ref["rm_clasif"]
    parts.append(
      f"La magnitud de reflexión (RM) es {c['categoria']} según edad y método "
      f"({ref['metodo_label']}; P75 {ref['rm_ref']['p75']:.2f}, P90 {ref['rm_ref']['p90']:.2f}, P95 {ref['rm_ref']['p95']:.2f})."
    )
  elif not np.isnan(rm):
    parts.append(
      f"RM {rm:.2f} disponible, pero no clasificable sin una referencia válida de edad y método comparable."
    )

  if ref.get("ok") and ref.get("ri_clasif", {}).get("ok"):
    c = ref["ri_clasif"]
    parts.append(
      f"RI {ri:.2f}: {c['categoria']} (P90 {ref['ri_ref']['p90']:.2f}); "
      "se informa como descriptor proporcional complementario y no como una segunda prueba diagnóstica independiente."
    )
  elif not np.isnan(ri):
    parts.append(
      f"RI {ri:.2f} disponible como descriptor complementario, sin corte fijo universal aplicado."
    )

  if not np.isnan(tref):
    parts.append(
      f"Tref {tref:.0f} ms se informa como variable temporal continua de investigación; "
      "no se aplica el antiguo corte fijo de 320 ms ni se genera diagnóstico automático de retorno precoz/tardío."
    )
  if not np.isnan(ratio):
    parts.append(
      f"La relación Tfor/Tref es {ratio:.2f} y se conserva como descriptor continuo de solapamiento temporal, "
      "sin umbral diagnóstico universal."
    )
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


def plot_wave_separation(sep_df, sep_metrics=None):
  """Gráfico clínico integrado: presión central + Pf/Pb y landmarks de sincronización."""
  t = pd.to_numeric(sep_df["tiempo_ms"], errors="coerce").to_numpy(dtype=float)
  p_total = pd.to_numeric(sep_df["presion_total_mmHg"], errors="coerce").to_numpy(dtype=float)
  pf_abs = pd.to_numeric(sep_df.get("onda_anterograda_pf_abs", sep_df["onda_anterograda_pf"]), errors="coerce").to_numpy(dtype=float)
  pb_abs = pd.to_numeric(sep_df.get("onda_retrograda_pb_abs", sep_df["onda_retrograda_pb"]), errors="coerce").to_numpy(dtype=float)
  ok = np.isfinite(t) & np.isfinite(p_total) & np.isfinite(pf_abs) & np.isfinite(pb_abs)
  t, p_total, pf_abs, pb_abs = t[ok], p_total[ok], pf_abs[ok], pb_abs[ok]
  pad_base = float(np.nanmin(p_total)) if len(p_total) else 0.0
  sync = _wave_sync_landmarks(sep_df, sep_metrics or {})

  fig, ax = plt.subplots(figsize=(8.8, 4.9))
  ax.plot(t, p_total, color="#111111", linewidth=3.0, label="Presión aórtica central completa", zorder=5)
  ax.plot(t, pf_abs, color="#168038", linewidth=2.35, label="Onda anterógrada Pf", zorder=4)
  ax.plot(t, pb_abs, color="#EF6C00", linestyle="--", linewidth=2.35, label="Onda retrógrada Pb", zorder=4)
  ax.fill_between(t, pad_base, pf_abs, alpha=0.07, color="#168038", zorder=2)
  ax.fill_between(t, pad_base, pb_abs, alpha=0.08, color="#EF6C00", zorder=2)
  ax.axhline(pad_base, color="#78909C", linewidth=0.9, alpha=0.8)
  # El cruce marcado es exactamente el usado por la animación.
  tc = sync.get("t_cross_ms", np.nan)
  ton = sync.get("t_pb_onset_ms", np.nan)
  if not np.isnan(ton):
    ax.axvline(ton, color="#EF6C00", linewidth=1.0, linestyle=":", alpha=0.75, label="Inicio material de Pb")
  if not np.isnan(tc):
    ax.axvline(tc, color="#6A1B9A", linewidth=1.25, linestyle="-.", alpha=0.85, label="Cruce Pf = Pb (sincroniza animación)")
  _apply_professional_axes(ax, "Separación de ondas superpuesta a la presión aórtica central", "Tiempo (ms)", "Presión / componentes sobre PAD (mmHg)")
  ax.legend(fontsize=7.4, loc="upper right", frameon=True, facecolor="white", edgecolor="#CFD8DC")
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



def _bezier_points(p0, p1, p2, p3, n=60):
  """Puntos de una curva cúbica de Bézier para dibujar una aorta anatómica didáctica."""
  p0 = np.asarray(p0, dtype=float); p1 = np.asarray(p1, dtype=float)
  p2 = np.asarray(p2, dtype=float); p3 = np.asarray(p3, dtype=float)
  u = np.linspace(0, 1, int(n))[:, None]
  pts = (1-u)**3*p0 + 3*(1-u)**2*u*p1 + 3*(1-u)*u**2*p2 + u**3*p3
  return pts[:, 0], pts[:, 1]


def _aorta_centerline_xy(n=220):
  """Centro anatómico simplificado: raíz, ascendente, arco y descendente."""
  segs = [
    ((0.23, 0.20), (0.25, 0.36), (0.27, 0.56), (0.35, 0.70)),
    ((0.35, 0.70), (0.43, 0.86), (0.62, 0.86), (0.71, 0.72)),
    ((0.71, 0.72), (0.78, 0.60), (0.75, 0.40), (0.69, 0.16)),
  ]
  xs, ys = [], []
  per = max(25, int(n/len(segs)))
  for j, seg in enumerate(segs):
    x, y = _bezier_points(*seg, n=per)
    if j > 0:
      x, y = x[1:], y[1:]
    xs.extend(x.tolist()); ys.extend(y.tolist())
  return np.asarray(xs), np.asarray(ys)


def _point_on_polyline(xs, ys, frac):
  """Devuelve punto interpolado sobre una polilínea según fracción de longitud."""
  frac = float(np.clip(frac, 0, 1))
  xs = np.asarray(xs, dtype=float); ys = np.asarray(ys, dtype=float)
  if len(xs) < 2:
    return float(xs[0]) if len(xs) else 0.5, float(ys[0]) if len(ys) else 0.5
  dist = np.r_[0, np.cumsum(np.sqrt(np.diff(xs)**2 + np.diff(ys)**2))]
  total = dist[-1] if dist[-1] > 0 else 1.0
  target = frac * total
  x = np.interp(target, dist, xs)
  y = np.interp(target, dist, ys)
  return float(x), float(y)


def _nearest_index_for_time(times, target_ms):
  t = np.asarray(times, dtype=float)
  ok = np.isfinite(t)
  if not np.any(ok):
    return 0
  t_ok = t[ok]
  idx_ok = np.where(ok)[0]
  return int(idx_ok[np.argmin(np.abs(t_ok - float(target_ms)))])


def _wave_sync_landmarks(sep_df, sep_metrics=None):
  """Landmarks temporales comunes para gráfico y animación Pf/Pb.

  La animación usa exactamente la misma serie temporal de ``sep_df`` que el gráfico.
  El cruce se define como el primer cambio Pf-Pb de positivo a no positivo dentro de
  la ventana sistólica útil; se evitan falsos cruces basales donde ambas ondas valen 0.
  La posición anatómica sigue siendo didáctica: el tiempo está medido/derivado de la
  curva, pero el punto espacial de encuentro no es una localización anatómica medida.
  """
  sep_metrics = sep_metrics or {}
  try:
    t = pd.to_numeric(sep_df["tiempo_ms"], errors="coerce").to_numpy(dtype=float)
    pf = pd.to_numeric(sep_df.get("onda_anterograda_pf", pd.Series(np.zeros(len(sep_df)))), errors="coerce").to_numpy(dtype=float)
    pb = pd.to_numeric(sep_df.get("onda_retrograda_pb", pd.Series(np.zeros(len(sep_df)))), errors="coerce").to_numpy(dtype=float)
    ok = np.isfinite(t) & np.isfinite(pf) & np.isfinite(pb)
    t, pf, pb = t[ok], pf[ok], pb[ok]
    if len(t) < 3:
      raise ValueError("serie insuficiente")
    order = np.argsort(t)
    t, pf, pb = t[order], pf[order], pb[order]

    total = np.clip(pf + pb, 0, None)
    max_total = max(float(np.nanmax(total)), 1e-6)
    max_pb = max(float(np.nanmax(pb)), 1e-6)

    def metric_time(key, default=np.nan):
      v = to_float(sep_metrics.get(key))
      return float(v) if not np.isnan(v) else float(default)

    t_min, t_max = float(t[0]), float(t[-1])
    t_foot = metric_time("t_pie_ms")
    if np.isnan(t_foot):
      cand = np.where(total >= 0.04 * max_total)[0]
      t_foot = float(t[cand[0]]) if len(cand) else t_min

    t_peak = metric_time("t_pico_ms")
    if np.isnan(t_peak):
      t_peak = float(t[int(np.nanargmax(total))])

    t_ej_end = metric_time("t_ej_fin_ms", t_max)
    t_ej_end = float(np.clip(t_ej_end, t_peak, t_max))

    # Inicio visible de Pb: primer aporte material, no el ruido basal.
    onset_mask = (
      (t >= t_foot)
      & (pb >= max(0.08 * max_pb, 0.025 * max_total))
      & (total >= 0.05 * max_total)
    )
    onset_idx = np.where(onset_mask)[0]
    t_pb_onset = float(t[onset_idx[0]]) if len(onset_idx) else metric_time("tref_ms", t_peak)

    # Cruce real Pf=Pb sobre la MISMA serie del gráfico. Se exige amplitud material.
    cross_mask = (
      (t >= max(t_foot, t_peak))
      & (t <= t_ej_end)
      & (total >= 0.08 * max_total)
    )
    idxs = np.where(cross_mask)[0]
    d = pf - pb
    cross_idx = None
    if len(idxs) >= 2:
      for a, b in zip(idxs[:-1], idxs[1:]):
        if d[a] > 0 and d[b] <= 0:
          cross_idx = b
          break
      if cross_idx is None:
        cross_idx = int(idxs[np.nanargmin(np.abs(d[idxs]))])
    else:
      cross_idx = int(np.nanargmin(np.abs(d)))
    t_cross = float(t[cross_idx])

    # Tref clínico/operativo ya calculado desde Pb; se mantiene en la misma línea temporal.
    t_ref = metric_time("tref_ms", t_cross)
    t_ref = float(np.clip(t_ref, t_min, t_max))

    # Coherencia temporal: Pb debe aparecer antes o, como máximo, en el cruce.
    if t_pb_onset >= t_cross:
      t_pb_onset = float(max(t_foot, t_cross - max(20.0, 0.12 * max(t_cross - t_foot, 1.0))))

    return {
      "t_min_ms": t_min,
      "t_max_ms": t_max,
      "t_foot_ms": float(t_foot),
      "t_peak_ms": float(t_peak),
      "t_pb_onset_ms": float(t_pb_onset),
      "t_cross_ms": float(t_cross),
      "t_ref_ms": float(t_ref),
      "t_ej_end_ms": float(t_ej_end),
      # Punto anatómico didáctico fijo; no representa sitio de reflexión medido.
      "meet_frac": 0.62,
    }
  except Exception:
    return {
      "t_min_ms": 0.0, "t_max_ms": 1000.0, "t_foot_ms": 0.0,
      "t_peak_ms": 250.0, "t_pb_onset_ms": 300.0, "t_cross_ms": 500.0,
      "t_ref_ms": 500.0, "t_ej_end_ms": 650.0, "meet_frac": 0.62,
    }


def _wave_position_fractions(ti, sync):
  """Fracciones espaciales Pf/Pb sincronizadas por el cruce real de sep_df."""
  ti = float(ti)
  t0 = float(sync.get("t_min_ms", 0.0))
  t1 = float(sync.get("t_max_ms", 1000.0))
  foot = float(sync.get("t_foot_ms", t0))
  pb_on = float(sync.get("t_pb_onset_ms", foot))
  cross = float(sync.get("t_cross_ms", (t0+t1)/2))
  meet = float(np.clip(sync.get("meet_frac", 0.62), 0.20, 0.85))

  def prog(x, a, b):
    if b <= a:
      return 1.0 if x >= b else 0.0
    return float(np.clip((x-a)/(b-a), 0.0, 1.0))

  # Pf nace en raíz y llega al punto de encuentro exactamente en t_cross.
  if ti <= cross:
    pf_frac = meet * prog(ti, foot, cross)
  else:
    pf_frac = meet + (1.0-meet) * prog(ti, cross, t1)

  # Pb no viaja antes de su inicio material; parte distal y llega al mismo punto en t_cross.
  pb_visible = ti >= pb_on
  if ti <= cross:
    pb_frac = 1.0 - (1.0-meet) * prog(ti, pb_on, cross)
  else:
    pb_frac = meet * (1.0 - prog(ti, cross, t1))

  pf_visible = ti >= foot
  return float(np.clip(pf_frac, 0, 1)), float(np.clip(pb_frac, 0, 1)), bool(pf_visible), bool(pb_visible)


def _harmonic_live_values(hdf, t_ms, max_harmonics=6):
  """Contribución instantánea de armónicos por amplitud y fase FFT."""
  vals = []
  try:
    if hdf is None or len(hdf) == 0:
      return vals
    for _, r in hdf.head(max_harmonics).iterrows():
      amp = float(r.get("amplitud", 0) or 0)
      freq = float(r.get("frecuencia_hz", 0) or 0)
      phase = float(r.get("fase_rad", 0) or 0)
      vals.append(amp * math.cos(2 * math.pi * freq * float(t_ms) / 1000.0 + phase))
  except Exception:
    return []
  return vals




def _wrap_lines_for_keyframe(text, width=54, max_lines=4):
  """Divide texto clínico para que entre dentro de una captura 2x2 sin recortarse."""
  txt = safe_text(text)
  if not txt:
    return []
  lines = []
  for part in re.split(r"\s*\|\s*|\.\s+", txt):
    part = safe_text(part)
    if not part:
      continue
    wrapped = textwrap.wrap(part, width=width, break_long_words=False, replace_whitespace=True)
    lines.extend(wrapped if wrapped else [part])
    if len(lines) >= max_lines:
      break
  if len(lines) > max_lines:
    lines = lines[:max_lines]
  if len(lines) == max_lines and len(" ".join(lines)) < len(txt) - 8:
    lines[-1] = lines[-1].rstrip(" .") + "..."
  return lines


def _frame_severity_color(severity):
  sev = safe_text(severity).lower()
  if "diagn" in sev:
    return "#B71C1C", "#FDECEA"
  if "alta" in sev:
    return "#D84315", "#FFF3E0"
  if "relev" in sev:
    return "#EF6C00", "#FFF8E1"
  return "#455A64", "#F5F7FA"


def _alerts_to_keyframes_for_report(row, sep_df, sep_metrics, hdf, max_frames=4):
  """Selecciona capturas del informe a partir de métricas alteradas reales.

  Usa el mismo motor de alertas de la pausa didáctica. Si hay más de cuatro
  alteraciones, prioriza diagnóstico y severidad. Si hay menos, completa el
  panel con una leyenda explícita de que no hay otra alerta automática para no
  reemplazar alteraciones por momentos genéricos.
  """
  try:
    t_vals = pd.to_numeric(sep_df.get("tiempo_ms", pd.Series([], dtype=float)), errors="coerce").dropna().to_numpy(dtype=float)
  except Exception:
    t_vals = np.asarray([], dtype=float)
  if len(t_vals) == 0:
    t_vals = np.linspace(0, 1000, 190)

  alerts = _build_animation_pause_alerts(row, sep_df, sep_metrics, hdf, t_vals)
  severity_rank = {"diagnóstica": 0, "diagnostica": 0, "alta": 1, "relevante": 2, "alterada": 3}
  alerts = sorted(alerts, key=lambda a: (severity_rank.get(safe_text(a.get("severity")).lower(), 9), int(a.get("pause_order", 99)), int(a.get("idx", 0))))

  frames = []
  for i, al in enumerate(alerts[:max_frames], start=1):
    value = safe_text(al.get("value", "ND"))
    threshold = safe_text(al.get("threshold", "criterio no disponible"))
    label = safe_text(al.get("label", "Métrica alterada"))
    frames.append({
      "title": f"{i}. Métrica alterada: {label}",
      "subtitle": f"Valor {value} | Criterio: {threshold}",
      "time_ms": float(al.get("t_ms", 0) or 0),
      "alert": al,
    })

  if not frames:
    base_frames = _animation_keyframe_definitions(sep_metrics)
    for i, fr in enumerate(base_frames[:max_frames], start=1):
      frames.append({
        **fr,
        "title": f"{i}. Sin métrica alterada automática",
        "subtitle": "No se detectó desviación relevante con los criterios disponibles; se muestra momento fisiológico de referencia.",
        "alert": {
          "label": "Sin alerta automática",
          "value": "sin alteración relevante",
          "threshold": "criterios operativos disponibles",
          "why": "Las métricas evaluadas no superan los umbrales definidos para pausa didáctica.",
          "mechanism": "La captura se mantiene como referencia anatómica y temporal de la onda central real.",
          "severity": "normal",
        }
      })
  elif len(frames) < max_frames:
    # Completar sin inventar alteraciones: se explicita que no hay más alertas.
    base_frames = _animation_keyframe_definitions(sep_metrics)
    k = 0
    while len(frames) < max_frames and k < len(base_frames):
      fr = base_frames[k]; k += 1
      frames.append({
        **fr,
        "title": f"{len(frames)+1}. Sin otra métrica alterada relevante",
        "subtitle": "Panel de cierre: no se agrega una alteración inexistente.",
        "alert": {
          "label": "Sin otra alerta automática",
          "value": "no corresponde",
          "threshold": "sin nuevo criterio superado",
          "why": "No se identifican más métricas alteradas con los puntos de corte disponibles.",
          "mechanism": "El informe prioriza solo valores reales alterados; este cuadro evita reemplazarlos por momentos genéricos.",
          "severity": "normal",
        }
      })
  return frames[:max_frames], alerts
def _plot_single_aorta_keyframe(ax, sep_df, sep_metrics, hdf, frame, row=None):
  """Dibuja una captura fija de la animación con anatomía, Pf/Pb, curva y armónicos."""
  t = pd.to_numeric(sep_df["tiempo_ms"], errors="coerce").to_numpy(dtype=float)
  p = pd.to_numeric(sep_df["presion_total_mmHg"], errors="coerce").to_numpy(dtype=float)
  pf = pd.to_numeric(sep_df.get("onda_anterograda_pf", pd.Series(np.zeros(len(sep_df)))), errors="coerce").to_numpy(dtype=float)
  pb = pd.to_numeric(sep_df.get("onda_retrograda_pb", pd.Series(np.zeros(len(sep_df)))), errors="coerce").to_numpy(dtype=float)
  q = pd.to_numeric(sep_df.get("flujo_aortico_estimado_ml_s", pd.Series(np.zeros(len(sep_df)))), errors="coerce").to_numpy(dtype=float)
  ok = np.isfinite(t) & np.isfinite(p)
  if not np.any(ok):
    ax.axis("off")
    ax.text(0.5, 0.5, "Captura no disponible", ha="center", va="center")
    return
  t, p = t[ok], p[ok]
  pf = pf[ok] if len(pf) == len(ok) else np.zeros_like(t)
  pb = pb[ok] if len(pb) == len(ok) else np.zeros_like(t)
  q = q[ok] if len(q) == len(ok) else np.zeros_like(t)

  idx = _nearest_index_for_time(t, frame["time_ms"])
  ti = float(t[idx]); pi = float(p[idx]); pfi = float(max(pf[idx], 0)); pbi = float(max(pb[idx], 0)); qi = float(max(q[idx], 0))
  min_p, max_p = float(np.nanmin(p)), float(np.nanmax(p))
  norm = (pi - min_p) / max(max_p - min_p, 1e-6)

  ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")
  ax.add_patch(plt.Rectangle((0.01, 0.01), 0.98, 0.98, facecolor="#FFFFFF", edgecolor="#CFD8DC", linewidth=0.9))
  ax.add_patch(plt.Rectangle((0.01, 0.86), 0.98, 0.13, facecolor="#F3F8FC", edgecolor="none"))
  ax.text(0.035, 0.945, frame["title"], fontsize=8.6, fontweight="bold", color="#12355B", ha="left", va="center")
  ax.text(0.035, 0.895, frame["subtitle"], fontsize=6.8, color="#455A64", ha="left", va="center")

  # Silueta cardíaca y raíz.
  heart = plt.Circle((0.18, 0.17), 0.075, color="#F8D7DA", alpha=0.80, ec="#C62828", lw=0.6)
  ax.add_patch(heart)
  ax.text(0.18, 0.17, "VI", ha="center", va="center", fontsize=6, color="#8E1B1B", fontweight="bold")

  xs, ys = _aorta_centerline_xy()
  outer_lw = 20 + 6.5 * norm
  inner_lw = 11 + 4.3 * norm
  ax.plot(xs, ys, color="#B23A48", linewidth=outer_lw, solid_capstyle="round", alpha=0.86, zorder=2)
  ax.plot(xs, ys, color="#FAD4D8", linewidth=inner_lw, solid_capstyle="round", alpha=0.97, zorder=3)
  ax.plot(xs, ys, color="#7F1D2D", linewidth=1.0, alpha=0.40, zorder=4)

  # Ramas supraaórticas.
  branches = [
    ((0.44,0.79), (0.41,0.93)),
    ((0.53,0.82), (0.53,0.96)),
    ((0.62,0.78), (0.69,0.93)),
  ]
  for (x0,y0),(x1,y1) in branches:
    ax.plot([x0,x1], [y0,y1], color="#B23A48", linewidth=8+2*norm, solid_capstyle="round", alpha=0.86, zorder=2)
    ax.plot([x0,x1], [y0,y1], color="#FAD4D8", linewidth=4+1.3*norm, solid_capstyle="round", alpha=0.97, zorder=3)
  ax.text(0.42, 0.965, "TBC", fontsize=5.4, color="#546E7A", ha="center")
  ax.text(0.535, 0.985, "CCI", fontsize=5.4, color="#546E7A", ha="center")
  ax.text(0.70, 0.945, "SCI", fontsize=5.4, color="#546E7A", ha="center")

  # Ondas Pf/Pb sincronizadas con la MISMA línea temporal del gráfico de descomposición.
  sync = _wave_sync_landmarks(sep_df, sep_metrics)
  pf_frac, pb_frac, pf_visible, pb_visible = _wave_position_fractions(ti, sync)
  xpf, ypf = _point_on_polyline(xs, ys, pf_frac)
  xpb, ypb = _point_on_polyline(xs, ys, pb_frac)
  max_pf = max(float(np.nanmax(pf)) if np.any(np.isfinite(pf)) else 1, 1e-6)
  max_pb = max(float(np.nanmax(pb)) if np.any(np.isfinite(pb)) else 1, 1e-6)
  if pf_visible:
    ax.add_patch(plt.Circle((xpf, ypf), 0.022 + 0.030*np.sqrt(pfi/max_pf), color="#1B8E3F", alpha=0.82, ec="white", lw=0.7, zorder=5))
    ax.text(xpf, min(0.98, ypf+0.058), "Pf", fontsize=6.5, color="#1B5E20", fontweight="bold", ha="center", zorder=6)
  if pb_visible:
    ax.add_patch(plt.Circle((xpb, ypb), 0.020 + 0.030*np.sqrt(pbi/max_pb), color="#EF6C00", alpha=0.82, ec="white", lw=0.7, zorder=5))
    ax.text(xpb, min(0.98, ypb+0.058), "Pb", fontsize=6.5, color="#E65100", fontweight="bold", ha="center", zorder=6)

  # Curva miniatura y cursor temporal.
  gx0, gy0, gw, gh = 0.06, 0.035, 0.45, 0.145
  ax.add_patch(plt.Rectangle((gx0, gy0), gw, gh, facecolor="#FBFCFE", edgecolor="#CFD8DC", linewidth=0.5, zorder=1))
  xx = gx0 + (t - float(np.nanmin(t))) / max(float(np.nanmax(t)-np.nanmin(t)), 1e-6) * gw
  yy = gy0 + (p - min_p) / max(max_p - min_p, 1e-6) * gh
  ax.plot(xx, yy, color="#111111", linewidth=1.1, zorder=5)
  frac = (ti - float(np.nanmin(t))) / max(float(np.nanmax(t) - np.nanmin(t)), 1e-6)
  xcur = gx0 + frac * gw
  ax.plot([xcur, xcur], [gy0, gy0+gh], color="#C62828", linewidth=1.1, zorder=6)
  ax.scatter([xcur], [gy0 + norm*gh], s=14, color="#C62828", zorder=7)
  ax.text(gx0, gy0+gh+0.012, "Curva central real", fontsize=5.6, color="#37474F", ha="left")

  # Barras armónicas sincronizadas.
  vals = _harmonic_live_values(hdf, ti, max_harmonics=6)
  hx0, hy0, hw, hh = 0.62, 0.05, 0.31, 0.22
  ax.add_patch(plt.Rectangle((hx0, hy0), hw, hh, facecolor="#FBFCFE", edgecolor="#CFD8DC", linewidth=0.5))
  ax.text(hx0, hy0+hh+0.014, "Armónicos FFT sincronizados", fontsize=5.7, color="#37474F", ha="left")
  if vals:
    vmax = max(max(abs(v) for v in vals), 1e-6)
    baseline = hy0 + hh*0.52
    ax.plot([hx0+0.015, hx0+hw-0.015], [baseline, baseline], color="#B0BEC5", linewidth=0.45)
    bar_w = hw / (len(vals)*1.45)
    for j, val in enumerate(vals):
      x = hx0 + 0.025 + j * (hw-0.06) / max(len(vals)-1, 1)
      h = abs(val) / vmax * hh*0.38
      y = baseline if val >= 0 else baseline - h
      c = "#1565C0" if val >= 0 else "#C62828"
      ax.add_patch(plt.Rectangle((x-bar_w/2, y), bar_w, h, facecolor=c, edgecolor="none", alpha=0.82))
      ax.text(x, hy0+0.006, f"H{j+1}", fontsize=5.1, color="#546E7A", ha="center")
  else:
    ax.text(hx0+hw/2, hy0+hh/2, "Sin datos FFT", fontsize=6, color="#78909C", ha="center", va="center")

  ax.text(0.57, 0.39, f"t {ti:.0f} ms", fontsize=7.1, color="#17365D", fontweight="bold")
  ax.text(0.57, 0.345, f"P {pi:.0f} mmHg | Q {qi:.0f} mL/s", fontsize=6.4, color="#263238")
  ax.text(0.57, 0.305, f"Pf {pfi:.1f} / Pb {pbi:.1f} mmHg", fontsize=6.4, color="#263238")
  if vals:
    ax.text(0.57, 0.265, f"H1 {vals[0]:+.2f}, H2 {vals[1] if len(vals)>1 else 0:+.2f}, H3 {vals[2] if len(vals)>2 else 0:+.2f}", fontsize=5.9, color="#263238")

  # Panel explicativo específico para capturas del informe integrado.
  alert = frame.get("alert") if isinstance(frame, dict) else None
  if alert:
    sev_color, sev_bg = _frame_severity_color(alert.get("severity", "alterada"))
    px, py, pw, ph = 0.535, 0.575, 0.425, 0.275
    ax.add_patch(plt.Rectangle((px, py), pw, ph, facecolor=sev_bg, edgecolor=sev_color, linewidth=0.9, zorder=20))
    ax.add_patch(plt.Rectangle((px, py+ph-0.045), pw, 0.045, facecolor=sev_color, edgecolor=sev_color, linewidth=0, zorder=21))
    ax.text(px+0.012, py+ph-0.023, "PAUSA DIDÁCTICA POR MÉTRICA ALTERADA", fontsize=5.7, color="white", fontweight="bold", ha="left", va="center", zorder=22)
    ax.text(px+pw-0.012, py+ph-0.023, safe_text(alert.get("severity", "ALERTA")).upper(), fontsize=5.4, color="white", fontweight="bold", ha="right", va="center", zorder=22)

    ytxt = py + ph - 0.070
    ax.text(px+0.012, ytxt, safe_text(alert.get("label", "Métrica alterada"))[:64], fontsize=6.3, color="#12355B", fontweight="bold", ha="left", va="top", zorder=22)
    ytxt -= 0.030
    ax.text(px+0.012, ytxt, f"Valor: {safe_text(alert.get('value', 'ND'))}", fontsize=5.9, color="#263238", fontweight="bold", ha="left", va="top", zorder=22)
    ytxt -= 0.027
    crit_lines = _wrap_lines_for_keyframe("Criterio: " + safe_text(alert.get("threshold", "")), width=52, max_lines=2)
    for line in crit_lines:
      ax.text(px+0.012, ytxt, line, fontsize=5.3, color="#37474F", ha="left", va="top", zorder=22)
      ytxt -= 0.023
    why_lines = _wrap_lines_for_keyframe("Por qué: " + safe_text(alert.get("why", "")), width=55, max_lines=2)
    for line in why_lines:
      ax.text(px+0.012, ytxt, line, fontsize=5.3, color="#37474F", ha="left", va="top", zorder=22)
      ytxt -= 0.023
    mech_lines = _wrap_lines_for_keyframe("Mecanismo: " + safe_text(alert.get("mechanism", "")), width=55, max_lines=3)
    for line in mech_lines:
      if ytxt < py + 0.018:
        break
      ax.text(px+0.012, ytxt, line, fontsize=5.15, color="#455A64", ha="left", va="top", zorder=22)
      ytxt -= 0.021


def _animation_keyframe_definitions(sep_metrics):
  """Elige cuatro momentos clínicamente relevantes para capturas fijas del ciclo."""
  t_pie = float(sep_metrics.get("t_pie_ms", 80) if not np.isnan(sep_metrics.get("t_pie_ms", np.nan)) else 80)
  t_peak = float(sep_metrics.get("t_pico_ms", 220) if not np.isnan(sep_metrics.get("t_pico_ms", np.nan)) else 220)
  t_ref = float(sep_metrics.get("tref_ms", 380) if not np.isnan(sep_metrics.get("tref_ms", np.nan)) else 380)
  t_ej_fin = float(sep_metrics.get("t_ej_fin_ms", 420) if not np.isnan(sep_metrics.get("t_ej_fin_ms", np.nan)) else 420)
  t_dia = float(np.clip(t_ej_fin + 0.38 * (1000 - t_ej_fin), t_ej_fin + 70, 850))
  return [
    {"title": "1. Inicio de eyección", "subtitle": "Apertura valvular y salida anterógrada inicial", "time_ms": t_pie},
    {"title": "2. Pico sistólico central", "subtitle": "Máxima distensión pulsátil de la aorta", "time_ms": t_peak},
    {"title": "3. Retorno de onda reflejada", "subtitle": "Contribución retrógrada Pb y carga pulsátil", "time_ms": t_ref},
    {"title": "4. Diástole / perfusión", "subtitle": "Descenso de presión y relación con área diastólica-SEVR", "time_ms": t_dia},
  ]


def plot_aortic_animation_keyframes(row, sep_df, sep_metrics, hdf):
  """Genera capturas del informe desde métricas alteradas reales.

  Las cuatro imágenes ya no son momentos genéricos fijos: se seleccionan desde
  las alertas de pausa didáctica, con valor real, criterio y mecanismo.
  """
  frames, alerts = _alerts_to_keyframes_for_report(row, sep_df, sep_metrics, hdf, max_frames=4)
  fig, axes = plt.subplots(2, 2, figsize=(12.4, 8.35))
  fig.patch.set_facecolor("white")
  for ax, frame in zip(axes.ravel(), frames):
    _plot_single_aorta_keyframe(ax, sep_df, sep_metrics, hdf, frame, row=row)
  n_alerts = len(alerts) if alerts is not None else 0
  title = "Capturas del informe: métricas alteradas de la animación hemodinámica"
  if n_alerts == 0:
    title = "Capturas del informe: sin métricas alteradas automáticas detectadas"
  fig.suptitle(title, fontsize=12.6, fontweight="bold", color="#12355B", y=0.997)
  fig.text(0.5, 0.014, "Cada cuadro usa la curva central real, Pf/Pb, flujo estimado y armónicos FFT sincronizados. Las tarjetas muestran valor real, criterio de alteración y explicación didáctica. La distensión aórtica es proporcional a la presión pulsátil y no mide diámetro anatómico.", ha="center", fontsize=7.35, color="#455A64")
  fig.tight_layout(rect=[0.006, 0.038, 0.994, 0.965], pad=1.05)
  buf = io.BytesIO()
  fig.savefig(buf, format="png", dpi=210, bbox_inches="tight", facecolor="white")
  plt.close(fig)
  buf.seek(0)
  return buf

def plot_harmonics(hdf):
  fig, ax = plt.subplots(figsize=(7.6, 4.2))
  ax.bar(range(1, len(hdf)+1), hdf["energia_relativa_%"], color="#455A64")
  ax.set_xticks(range(1, len(hdf)+1))
  _apply_professional_axes(ax, "Análisis armónico de la onda de presión central", "Armónico", "Energía relativa (%)")
  ax.grid(axis="y", alpha=.22)
  return fig_to_png(fig)



def _resample_for_animation(x, y, n=180):
  """Reduce una señal real a una longitud liviana para animación HTML."""
  try:
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)
    ok = np.isfinite(x) & np.isfinite(y)
    x, y = x[ok], y[ok]
    if len(x) < 3:
      return [], []
    order = np.argsort(x)
    x, y = x[order], y[order]
    keep = np.r_[True, np.diff(x) > 0]
    x, y = x[keep], y[keep]
    if len(x) < 3:
      return [], []
    grid = np.linspace(float(np.nanmin(x)), float(np.nanmax(x)), int(n))
    vals = np.interp(grid, x, y)
    return [round(float(v), 3) for v in grid], [round(float(v), 3) for v in vals]
  except Exception:
    return [], []


def _safe_metric_value(value, decimals=1, suffix=""):
  """Valor compacto para tarjetas de animación."""
  v = to_float(value)
  if np.isnan(v):
    return "ND"
  if decimals <= 0:
    return f"{v:.0f}{suffix}"
  return f"{v:.{decimals}f}{suffix}"


def _animation_vector_from_sep(sep_df, column, n=180, default_zero=False):
  """Obtiene una columna real de sep_df y la remuestrea para la animación."""
  try:
    t = pd.to_numeric(sep_df["tiempo_ms"], errors="coerce").to_numpy(dtype=float)
    if column in sep_df:
      y = pd.to_numeric(sep_df[column], errors="coerce").to_numpy(dtype=float)
    elif default_zero:
      y = np.zeros_like(t, dtype=float)
    else:
      return [], []
    return _resample_for_animation(t, y, n=n)
  except Exception:
    return [], []



def _nearest_index_for_time_ms(t_values, target_ms):
  """Índice de la serie de animación más cercano a un tiempo fisiológico dado."""
  try:
    arr = np.asarray(t_values, dtype=float)
    if len(arr) == 0:
      return 0
    target = float(target_ms)
    if np.isnan(target):
      target = float(arr[len(arr)//2])
    return int(np.nanargmin(np.abs(arr - target)))
  except Exception:
    return 0


def _build_animation_pause_alerts(row, sep_df, sep_metrics, hdf, t_values):
  """Construye alertas didácticas para pausar la animación en métricas alteradas.

  Cada alerta se ubica en un momento fisiológico del ciclo:
  pico sistólico para presión/carga pulsátil, Tref para reflexión/aumentación,
  fin de eyección/diástole para RVSE-SEVR y pico para complejidad armónica.
  """
  alerts = []

  def fmt(v, dec=1, unit=""):
    try:
      f = float(v)
      if np.isnan(f):
        return "ND"
      if dec == 0:
        return f"{f:.0f}{unit}"
      return f"{f:.{dec}f}{unit}"
    except Exception:
      return "ND"

  def add(label, value, threshold, why, mechanism, target_ms, severity="alterada"):
    # Permitir varias pausas didácticas en el mismo ciclo.
    # Antes se limitaba a pocas alertas y luego se fusionaban, lo que hacía
    # que se vieran solo una o dos pausas aunque hubiese más métricas alteradas.
    if len(alerts) >= 14:
      return
    idx = _nearest_index_for_time_ms(t_values, target_ms)
    try:
      t_ms = float(np.asarray(t_values, dtype=float)[idx])
    except Exception:
      t_ms = target_ms if target_ms is not None else 0
    alerts.append({
      "idx": int(idx),
      "t_ms": round(float(t_ms), 1) if t_ms is not None and not np.isnan(float(t_ms)) else 0,
      "label": safe_text(label),
      "value": safe_text(value),
      "threshold": safe_text(threshold),
      "why": safe_text(why),
      "mechanism": safe_text(mechanism),
      "severity": safe_text(severity),
    })

  try:
    pas_c = to_float(row.get("pas_central"))
    pad_c = to_float(row.get("pad_central"))
    pp_c = to_float(row.get("pp_central"))
    pas_r = to_float(row.get("pas_radial"))
    pp_r = to_float(row.get("pp_radial"))
    iau = to_float(row.get("iau"))
    au = to_float(row.get("au"))
    rm = to_float(sep_metrics.get("rm"))
    ri = to_float(sep_metrics.get("ri"))
    tref = to_float(sep_metrics.get("tref_ms"))
    rvse_calc = to_float(sep_metrics.get("rvse_calculado_%"))
    t_peak = to_float(sep_metrics.get("t_pico_ms"))
    t_ref = tref if not np.isnan(tref) else to_float(sep_metrics.get("t_ej_fin_ms"))
    t_ej_fin = to_float(sep_metrics.get("t_ej_fin_ms"))
    if np.isnan(t_peak):
      try:
        t_peak = float(sep_df.loc[pd.to_numeric(sep_df["presion_total_mmHg"], errors="coerce").idxmax(), "tiempo_ms"])
      except Exception:
        t_peak = 250.0
    if np.isnan(t_ref):
      t_ref = t_peak + 120.0
    if np.isnan(t_ej_fin):
      t_ej_fin = min(t_peak + 220.0, 650.0)

    canonical = build_canonical_diagnostic_state(row, sep_metrics, hdf)
    dom = canonical["domains"]
    saha_ref = get_saha_central_sbp_reference(row)
    if dom["hta_central"].get("altered"):
      add(
        "PAS central / hipertensión central",
        fmt(pas_c, 0, " mmHg"),
        f"límite diagnóstico P90 SAHA {saha_ref['p90']:.1f} mmHg",
        "La PAS central alcanza o supera el percentilo 90 ajustado por edad, sexo y calibración.",
        "Se expresa en el pico sistólico central: mayor presión transmitida a aorta y ventrículo izquierdo.",
        t_peak,
        "diagnóstica",
      )

    if dom["carga_pulsatil"].get("altered"):
      if dom["carga_pulsatil"].get("grade") == "alta":
        add(
          "Presión de pulso central elevada",
          fmt(pp_c, 0, " mmHg"),
          "desviación relevante si ≥60 mmHg",
          "La diferencia PAS-PAD central es amplia y aumenta la carga pulsátil central.",
          "Se produce por mayor amplitud de la onda central, menor amortiguación arterial y/o mayor retorno reflejo.",
          t_peak,
          "alta",
        )
      else:
        add(
          "Presión de pulso central aumentada",
          fmt(pp_c, 0, " mmHg"),
          "alterada si ≥50 mmHg",
          "La amplitud pulsátil central está aumentada.",
          "Refleja mayor carga pulsátil sobre aorta y ventrículo; debe integrarse con VOP, AIx y Pf/Pb.",
          t_peak,
          "relevante",
        )

    saha_aix = get_saha_aix75_reference(row)
    if dom["aumentacion"].get("altered"):
      add(
        "IAu/AIx central aumentado",
        fmt(iau, 1, "%"),
        f"aumentado si ≥P90 LEAD {saha_aix['p90']:.1f}%",
        "El índice de aumentación supera el límite esperado para edad y sexo.",
        "Suele producirse cuando la onda reflejada se suma precozmente a la sístole central y eleva el hombro sistólico.",
        t_ref,
        "relevante",
      )


    ppa = dom["amplificacion"].get("value", np.nan)
    if dom["amplificacion"].get("altered"):
      add(
        "Amplificación central-periférica reducida",
        fmt(ppa, 2, ""),
        "reducida si PPA <1.30; franca si <1.20",
        "La presión de pulso periférica no se amplifica lo esperable respecto de la central.",
        "Puede expresar menor amortiguación arterial y transmisión central más directa de la carga pulsátil.",
        t_peak,
        "relevante" if ppa >= 1.20 else "alta",
      )

    rmri_ref = get_rm_ri_reference(row, sep_metrics)
    if dom["reflexion"].get("altered"):
      rmcat = rmri_ref["rm_clasif"].get("categoria", "aumentada")
      add(
        "RM Pb/Pf aumentada por percentil",
        fmt(rm, 2, ""),
        f"P90 {rmri_ref['rm_ref']['p90']:.2f}; P95 {rmri_ref['rm_ref']['p95']:.2f}",
        f"La RM se clasifica como {rmcat} para la edad y el método seleccionado.",
        "Mayor magnitud retrógrada relativa; RI se informa como complemento y no genera una segunda alerta independiente.",
        t_ref,
        "alta" if rmri_ref["rm_clasif"].get("grado") == "marcada" else "relevante",
      )


    if dom["rvse"].get("altered"):
      add(
        "RVSE/SEVR calculado reducido",
        fmt(rvse_calc, 1, "%"),
        "reducido si <120%",
        "La relación área diastólica/área sistólica es baja.",
        "Se produce por menor predominio relativo del área diastólica de perfusión frente al área sistólica de demanda.",
        t_ej_fin,
        "alta",
      )

    e_high = np.nan
    e1 = np.nan
    try:
      if hdf is not None and len(hdf):
        e = pd.to_numeric(hdf.get("energia_relativa_%"), errors="coerce").to_numpy(dtype=float)
        if len(e) > 0:
          e1 = e[0]
        if len(e) > 4:
          e_high = float(np.nansum(e[3:]))
    except Exception:
      pass

    # Pausas secuenciales por métrica:
    # No se fusionan alertas cercanas. Si varias métricas pertenecen al mismo
    # momento fisiológico —por ejemplo pico sistólico o retorno reflejo— se
    # separan unos cuadros para que cada una tenga su propia pausa didáctica.
    n_frames = max(1, len(t_values) if t_values is not None else 1)
    t_arr = np.asarray(t_values, dtype=float) if t_values is not None else np.asarray([0.0])
    severity_rank = {"diagnóstica": 0, "alta": 1, "relevante": 2, "alterada": 3}
    alerts = sorted(alerts, key=lambda a: (a.get("idx", 0), severity_rank.get(a.get("severity", "alterada"), 9), a.get("label", "")))

    groups = []
    for al in alerts:
      if groups and abs(al.get("idx", 0) - groups[-1]["base_idx"]) <= 2:
        groups[-1]["items"].append(al)
      else:
        groups.append({"base_idx": int(al.get("idx", 0)), "items": [al]})

    sequential = []
    min_gap = max(4, n_frames // 40)  # separación visual breve dentro del mismo evento fisiológico
    for group in groups:
      base_idx = int(group["base_idx"])
      for j, al in enumerate(group["items"]):
        new_idx = min(n_frames - 1, max(0, base_idx + j * min_gap))
        al["idx_original"] = int(al.get("idx", new_idx))
        al["idx"] = int(new_idx)
        try:
          al["t_ms"] = round(float(t_arr[new_idx]), 1)
        except Exception:
          pass
        al["pause_order"] = len(sequential) + 1
        sequential.append(al)

    # Si el corrimiento de un grupo se acerca demasiado al siguiente, conservar orden
    # y dejar que el motor JS use pause_order para no saltear ninguna alerta.
    return sequential[:12]
  except Exception:
    return alerts[:12]

def render_aortic_real_metrics_animation(row, sep_df, sep_metrics, hdf, height=1080):
  """Animación HTML/SVG de aorta con presión real, separación Pf/Pb y armónicos.

  La animación no crea una curva nueva: usa la curva real regularizada del paciente,
  las ondas Pf/Pb estimadas desde esa curva y el espectro armónico calculado por FFT.
  El cambio de calibre aórtico es didáctico y proporcional a la presión pulsátil real;
  no representa medición anatómica directa del diámetro de la aorta.
  """
  try:
    t, p = _animation_vector_from_sep(sep_df, "presion_total_mmHg", n=190)
    _, pf = _animation_vector_from_sep(sep_df, "onda_anterograda_pf", n=190, default_zero=True)
    _, pb = _animation_vector_from_sep(sep_df, "onda_retrograda_pb", n=190, default_zero=True)
    _, q = _animation_vector_from_sep(sep_df, "flujo_aortico_estimado_ml_s", n=190, default_zero=True)
    if not t or not p:
      return ""

    energies = []
    freqs = []
    amps = []
    phases = []
    if hdf is not None and len(hdf):
      energies = [round(float(x), 3) for x in pd.to_numeric(hdf.get("energia_relativa_%", []), errors="coerce").fillna(0).tolist()[:10]]
      freqs = [round(float(x), 3) for x in pd.to_numeric(hdf.get("frecuencia_hz", []), errors="coerce").fillna(0).tolist()[:10]]
      amps = [round(float(x), 6) for x in pd.to_numeric(hdf.get("amplitud", []), errors="coerce").fillna(0).tolist()[:10]]
      if "fase_rad" in hdf:
        phases = [round(float(x), 6) for x in pd.to_numeric(hdf.get("fase_rad", []), errors="coerce").fillna(0).tolist()[:10]]
      else:
        phases = [0.0 for _ in amps]
    while len(energies) < 10:
      energies.append(0.0)
    while len(freqs) < 10:
      freqs.append(0.0)
    while len(amps) < 10:
      amps.append(0.0)
    while len(phases) < 10:
      phases.append(0.0)

    pas = to_float(row.get("pas_central"))
    pad = to_float(row.get("pad_central"))
    pp = to_float(row.get("pp_central"))
    iau = to_float(row.get("iau"))
    au = to_float(row.get("au"))
    rvse_eq = to_float(row.get("rvse"))
    rvse_calc = to_float(sep_metrics.get("rvse_calculado_%"))
    rm = to_float(sep_metrics.get("rm"))
    ri = to_float(sep_metrics.get("ri"))
    tref = to_float(sep_metrics.get("tref_ms"))
    tfor = to_float(sep_metrics.get("tfor_ms"))
    qp = to_float(sep_metrics.get("qp_ml_s"))
    pe_ms = to_float(sep_metrics.get("pe_ms"))
    ai_morph = to_float(sep_metrics.get("ai_morfologico_%"))
    curve_id = safe_text(sep_metrics.get("curve_id", "sin_firma"))
    hta_status = central_hypertension_status(row)
    saha_ref = get_saha_central_sbp_reference(row)
    p90 = saha_ref.get("p90", np.nan) if saha_ref.get("ok") else np.nan
    aix_ref = get_saha_aix75_reference(row)
    aix_p90 = aix_ref.get("p90", np.nan) if aix_ref.get("ok") else np.nan

    pause_alerts = _build_animation_pause_alerts(row, sep_df, sep_metrics, hdf, t)
    sync = _wave_sync_landmarks(sep_df, sep_metrics)

    payload = {
      "t": t,
      "p": p,
      "pf": pf,
      "pb": pb,
      "q": q,
      "energies": energies,
      "freqs": freqs,
      "amps": amps,
      "phases": phases,
      "alerts": pause_alerts,
      "sync": {k: round(float(v), 4) for k, v in sync.items()},
      "metrics": {
        "pas": _safe_metric_value(pas, 0, " mmHg"),
        "pad": _safe_metric_value(pad, 0, " mmHg"),
        "pp": _safe_metric_value(pp, 0, " mmHg"),
        "p90": _safe_metric_value(p90, 1, " mmHg"),
        "hta": "CON HTA central" if hta_status.get("tiene_hta_central") else "SIN HTA central",
        "iau": _safe_metric_value(iau, 1, "%"),
        "iau_p90": _safe_metric_value(aix_p90, 1, "%"),
        "au": _safe_metric_value(au, 1, " mmHg"),
        "rm": _safe_metric_value(rm, 2, ""),
        "ri": _safe_metric_value(ri, 2, ""),
        "tref": _safe_metric_value(tref, 0, " ms"),
        "tref_label": "Tref retorno reflejo",
        "tfor": _safe_metric_value(tfor, 0, " ms"),
        "qp": _safe_metric_value(qp, 0, " mL/s"),
        "pe": _safe_metric_value(pe_ms, 0, " ms"),
        "rvse_calc": _safe_metric_value(rvse_calc, 0, "%"),
        "rvse_eq": _safe_metric_value(rvse_eq, 0, "%"),
        "ai_morph": _safe_metric_value(ai_morph, 1, "%"),
        "curve_id": curve_id,
      }
    }
    data_json = json.dumps(payload, ensure_ascii=False)

    html = f"""
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<style>
  :root {{ --bg:#f7fbff; --ink:#183044; --muted:#617385; --card:#ffffff; --line:#d5e3ef; --aorta:#b71c1c; --pf:#168038; --pb:#ef6c00; --flow:#6a1b9a; }}
  body {{ margin:0; font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif; background:transparent; color:var(--ink); }}
  .wrap {{ border:1px solid var(--line); border-radius:22px; background:linear-gradient(180deg,#ffffff 0%,var(--bg) 100%); padding:18px; box-sizing:border-box; }}
  .top {{ display:flex; gap:14px; align-items:stretch; flex-wrap:wrap; }}
  .title {{ flex:1 1 380px; padding:8px 4px; }}
  h2 {{ margin:0 0 6px 0; font-size:22px; letter-spacing:-.02em; }}
  .sub {{ color:var(--muted); font-size:13px; line-height:1.35; max-width:980px; }}
  .cards {{ display:grid; grid-template-columns:repeat(4,minmax(115px,1fr)); gap:10px; flex:1 1 470px; }}
  .card {{ background:var(--card); border:1px solid var(--line); border-radius:16px; padding:10px 12px; box-shadow:0 6px 16px rgba(18,53,91,.05); }}
  .k {{ color:var(--muted); font-size:11px; text-transform:uppercase; letter-spacing:.08em; }}
  .v {{ font-size:18px; font-weight:800; margin-top:3px; }}
  .stage {{ display:grid; grid-template-columns:1.15fr .85fr; gap:16px; margin-top:16px; }}
  .panel {{ background:#fff; border:1px solid var(--line); border-radius:18px; padding:12px; box-shadow:0 8px 18px rgba(18,53,91,.05); }}
  .panel h3 {{ font-size:14px; margin:0 0 8px 0; }}
  svg {{ width:100%; height:auto; display:block; }}
  .legend {{ display:flex; gap:12px; flex-wrap:wrap; color:var(--muted); font-size:12px; margin-top:8px; }}
  .dot {{ width:10px; height:10px; border-radius:50%; display:inline-block; margin-right:5px; vertical-align:-1px; }}
  .controls {{ margin-top:12px; display:flex; align-items:center; gap:10px; flex-wrap:wrap; }}
  button {{ border:0; background:#12355b; color:#fff; border-radius:999px; padding:9px 14px; font-weight:700; cursor:pointer; }}
  .toggle {{ display:flex; align-items:center; gap:7px; color:var(--muted); font-size:12px; background:#f5f9fd; border:1px solid var(--line); border-radius:999px; padding:7px 10px; }}
  .toggle input {{ accent-color:#12355b; }}
  input[type=range] {{ flex:1 1 320px; accent-color:#12355b; }}
  .timebox {{ color:var(--muted); font-variant-numeric:tabular-nums; min-width:86px; text-align:right; }}
  .alertCard {{
    margin-top:12px;
    border:1px solid #f3c37a;
    background:#fff8eb;
    border-radius:16px;
    padding:14px 16px;
    display:none;
    box-shadow:0 8px 18px rgba(135,86,10,.08);
    width:100%;
    box-sizing:border-box;
    max-height:none;
    overflow:visible;
  }}
  .alertCard.show {{ display:block; }}
  .alertHead {{ display:flex; align-items:flex-start; justify-content:space-between; gap:10px; margin-bottom:8px; flex-wrap:wrap; }}
  .alertTitle {{ color:#7a3d00; font-weight:900; font-size:15px; line-height:1.25; overflow-wrap:anywhere; }}
  .alertPill {{ background:#7a3d00; color:#fff; border-radius:999px; padding:4px 9px; font-size:10px; font-weight:800; text-transform:uppercase; letter-spacing:.07em; white-space:nowrap; }}
  .alertLine {{
    font-size:13px;
    color:#48311a;
    line-height:1.55;
    margin-top:6px;
    white-space:normal;
    overflow-wrap:anywhere;
    word-break:normal;
  }}
  .alertLine b {{ color:#2a1d0f; }}
  .grid2 {{ display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:10px; }}
  .small {{ font-size:12px; color:var(--muted); line-height:1.35; }}
  .barbg {{ fill:#e8f0f7; }}
  .barEnergy {{ fill:#455a64; opacity:.35; }}
  .barLive {{ fill:#12355b; opacity:.88; transition:height .08s, y .08s, opacity .08s; }}
  .barLive.positive {{ fill:#1565c0; }}
  .barLive.negative {{ fill:#ad1457; }}
  .barEnergy.active {{ opacity:.78; }}
  .phaseDot {{ fill:#12355b; opacity:.85; transition:cy .08s; }}
  .waveLine {{ fill:none; stroke:#111; stroke-width:2.5; }}
  .pfLine {{ fill:none; stroke:var(--pf); stroke-width:2; }}
  .pbLine {{ fill:none; stroke:var(--pb); stroke-width:2; stroke-dasharray:5 4; }}
  .marker {{ stroke:#fff; stroke-width:2; }}
  .metricOverlayBox {{ fill:#ffffff; stroke:#d5e3ef; stroke-width:1.4; opacity:.96; }}
  .metricOverlayTitle {{ fill:#12355b; font-size:11px; font-weight:900; letter-spacing:.05em; text-transform:uppercase; }}
  .metricOverlayLabel {{ fill:#617385; font-size:10.5px; }}
  .metricOverlayValue {{ fill:#183044; font-size:16px; font-weight:900; }}
  .metricOverlayMini {{ fill:#617385; font-size:10.5px; }}
  .metricOverlayLine {{ stroke:#d5e3ef; stroke-width:1; }}
  @media (max-width:850px) {{ .stage {{ grid-template-columns:1fr; }} .cards {{ grid-template-columns:repeat(2,minmax(115px,1fr)); }} .alertLine {{ font-size:12.5px; line-height:1.5; }} }}
</style>
</head>
<body>
<div class="wrap">
  <div class="top">
    <div class="title">
      <h2>Animación hemodinámica central con datos reales</h2>
      <div class="sub">Integra la curva central real digitalizada/importada, separación de ondas Pf/Pb, flujo aórtico estimado y armónicos por FFT. Pf/Pb se mueven con la misma línea temporal de la descomposición: inicio material de Pb y cruce Pf=Pb están sincronizados. La posición anatómica del encuentro es didáctica y no un sitio de reflexión medido.</div>
    </div>
    <div class="cards">
      <div class="card"><div class="k">Diagnóstico</div><div class="v" id="mHTA"></div></div>
      <div class="card"><div class="k">PAS/PAD central</div><div class="v" id="mPASPAD"></div></div>
      <div class="card"><div class="k">RM Pb/Pf</div><div class="v" id="mRM"></div></div>
      <div class="card"><div class="k">Tref retorno reflejo</div><div class="v" id="mTref"></div></div>
    </div>
  </div>

  <div class="stage">
    <div class="panel">
      <h3>Aorta anatómica funcional: raíz, ascendente, arco, ramas supraaórticas y descendente</h3>
      <svg viewBox="0 0 820 430" role="img" aria-label="Animación anatómica de aorta con curva de presión central">
        <defs>
          <linearGradient id="aortaGrad" x1="0" x2="1" y1="0" y2="0">
            <stop offset="0%" stop-color="#6f0808"/><stop offset="45%" stop-color="#d93d31"/><stop offset="100%" stop-color="#7d1010"/>
          </linearGradient>
          <linearGradient id="lumenGrad" x1="0" x2="1" y1="0" y2="0">
            <stop offset="0%" stop-color="#3f0505"/><stop offset="48%" stop-color="#8c1212"/><stop offset="100%" stop-color="#4a0606"/>
          </linearGradient>
          <radialGradient id="rootGlow" cx="50%" cy="50%" r="70%">
            <stop offset="0%" stop-color="#ffb4a9" stop-opacity=".85"/><stop offset="78%" stop-color="#b71c1c" stop-opacity=".20"/><stop offset="100%" stop-color="#b71c1c" stop-opacity="0"/>
          </radialGradient>
          <filter id="softShadow"><feDropShadow dx="0" dy="6" stdDeviation="7" flood-color="#12355b" flood-opacity=".18"/></filter>
          <filter id="waveGlow"><feDropShadow dx="0" dy="0" stdDeviation="4" flood-color="#ffffff" flood-opacity=".80"/></filter>
        </defs>
        <rect x="18" y="18" width="784" height="394" rx="20" fill="#fbfdff" stroke="#d5e3ef"/>
        <g id="anatomyLayer" filter="url(#softShadow)">
          <!-- Sombra de silueta cardiaca para orientar la raíz aórtica -->
          <path d="M165 215 C122 182 124 129 169 107 C198 92 227 105 244 130 C263 101 303 92 331 116 C377 157 347 226 247 286 C211 263 184 238 165 215 Z" fill="#f2b5ae" opacity=".28"/>
          <path id="rootPulse" d="M246 198 C232 177 234 144 256 126 C278 108 312 111 329 134 C344 154 338 184 317 198 C296 212 263 217 246 198 Z" fill="url(#rootGlow)" opacity=".90"/>

          <!-- Centro anatómico de la aorta; también se usa para mover Pf/Pb -->
          <path id="aortaCenter" d="M284 199 C258 158 266 103 309 83 C357 61 431 63 481 91 C525 115 541 159 522 193 C506 222 476 239 451 248 C425 258 415 275 420 304 C424 329 423 354 414 380" fill="none" stroke="transparent" stroke-width="2"/>

          <!-- Pared y lumen de la aorta anatómica -->
          <path id="aortaOuter" d="M284 199 C258 158 266 103 309 83 C357 61 431 63 481 91 C525 115 541 159 522 193 C506 222 476 239 451 248 C425 258 415 275 420 304 C424 329 423 354 414 380" fill="none" stroke="url(#aortaGrad)" stroke-width="60" stroke-linecap="round" stroke-linejoin="round" opacity=".98"/>
          <path id="aortaInner" d="M284 199 C258 158 266 103 309 83 C357 61 431 63 481 91 C525 115 541 159 522 193 C506 222 476 239 451 248 C425 258 415 275 420 304 C424 329 423 354 414 380" fill="none" stroke="url(#lumenGrad)" stroke-width="35" stroke-linecap="round" stroke-linejoin="round" opacity=".82"/>
          <path id="aortaHighlight" d="M292 175 C285 135 292 104 324 90 C368 72 430 78 470 101 C503 120 516 151 506 177" fill="none" stroke="#ffd4cc" stroke-width="7" stroke-linecap="round" opacity=".55"/>

          <!-- Ramas supraaórticas: braquiocefálica, carótida común izquierda y subclavia izquierda -->
          <path id="branch1Outer" d="M358 72 C352 44 337 27 316 13" fill="none" stroke="url(#aortaGrad)" stroke-width="28" stroke-linecap="round"/>
          <path id="branch1Inner" d="M358 72 C352 44 337 27 316 13" fill="none" stroke="url(#lumenGrad)" stroke-width="15" stroke-linecap="round" opacity=".83"/>
          <path id="branch1bOuter" d="M344 42 C324 37 306 44 290 59" fill="none" stroke="url(#aortaGrad)" stroke-width="20" stroke-linecap="round"/>
          <path id="branch1bInner" d="M344 42 C324 37 306 44 290 59" fill="none" stroke="url(#lumenGrad)" stroke-width="10" stroke-linecap="round" opacity=".78"/>
          <path id="branch2Outer" d="M416 68 C418 42 414 24 405 8" fill="none" stroke="url(#aortaGrad)" stroke-width="24" stroke-linecap="round"/>
          <path id="branch2Inner" d="M416 68 C418 42 414 24 405 8" fill="none" stroke="url(#lumenGrad)" stroke-width="13" stroke-linecap="round" opacity=".82"/>
          <path id="branch3Outer" d="M472 90 C498 63 523 48 552 36" fill="none" stroke="url(#aortaGrad)" stroke-width="24" stroke-linecap="round"/>
          <path id="branch3Inner" d="M472 90 C498 63 523 48 552 36" fill="none" stroke="url(#lumenGrad)" stroke-width="13" stroke-linecap="round" opacity=".82"/>

          <!-- Válvula aórtica y referencias anatómicas -->
          <ellipse id="valveRing" cx="279" cy="199" rx="31" ry="18" fill="#5a0707" opacity=".92" transform="rotate(28 279 199)"/>
          <path d="M263 196 Q278 183 293 198 Q278 205 263 196" fill="#ffd1cb" opacity=".74"/>
          <path d="M275 210 Q286 194 304 205 Q288 216 275 210" fill="#ffd1cb" opacity=".62"/>
          <text x="202" y="323" fill="#617385" font-size="12">Raíz aórtica</text>
          <text x="292" y="112" fill="#617385" font-size="12">Ascendente</text>
          <text x="430" y="132" fill="#617385" font-size="12">Arco</text>
          <text x="456" y="354" fill="#617385" font-size="12">Descendente</text>
          <text x="282" y="35" fill="#617385" font-size="10">Tronco braquiocefálico</text>
          <text x="390" y="28" fill="#617385" font-size="10">Carótida izq.</text>
          <text x="529" y="61" fill="#617385" font-size="10">Subclavia izq.</text>
        </g>

        <g id="pfWave" filter="url(#waveGlow)">
          <circle r="14" fill="#168038" opacity=".92" class="marker"/>
          <circle r="26" fill="#168038" opacity=".12"/>
        </g>
        <g id="pbWave" filter="url(#waveGlow)">
          <circle r="12" fill="#ef6c00" opacity=".92" class="marker"/>
          <circle r="23" fill="#ef6c00" opacity=".13"/>
        </g>

        <!-- Tablero dinámico en primer plano: evita que las métricas queden detrás del arco aórtico -->
        <g id="liveMetricOverlay" transform="translate(34 34)">
          <rect class="metricOverlayBox" x="0" y="0" width="218" height="124" rx="16"/>
          <text class="metricOverlayTitle" x="14" y="22">Métricas dinámicas</text>
          <line class="metricOverlayLine" x1="14" x2="204" y1="32" y2="32"/>
          <text class="metricOverlayLabel" x="14" y="51">Presión central instantánea</text>
          <text id="liveP" class="metricOverlayValue" x="14" y="72">-- mmHg</text>
          <text class="metricOverlayLabel" x="116" y="51">Flujo</text>
          <text id="liveQOverlay" class="metricOverlayValue" x="116" y="72">--</text>
          <text class="metricOverlayLabel" x="14" y="94">Pf / Pb</text>
          <text id="livePfPbOverlay" class="metricOverlayMini" x="14" y="112">-- / --</text>
          <text class="metricOverlayLabel" x="116" y="94">Tiempo ciclo</text>
          <text id="liveTimeOverlay" class="metricOverlayMini" x="116" y="112">0 ms</text>
        </g>

        <text x="92" y="192" fill="#168038" font-size="13" font-weight="700">Pf anterógrada</text>
        <text x="562" y="222" fill="#ef6c00" font-size="13" font-weight="700">Pb retrógrada</text>

        <g transform="translate(70 242)">
          <rect x="0" y="0" width="630" height="140" rx="14" fill="#ffffff" stroke="#d5e3ef"/>
          <path id="pressurePath" class="waveLine" d=""/>
          <path id="pfPath" class="pfLine" d=""/>
          <path id="pbPath" class="pbLine" d=""/>
          <line id="cursor" x1="0" x2="0" y1="10" y2="130" stroke="#12355b" stroke-width="1.4" opacity=".75"/>
          <circle id="pDot" r="5" fill="#111" class="marker"/>
          <circle id="pfDot" r="4" fill="#168038" class="marker"/>
          <circle id="pbDot" r="4" fill="#ef6c00" class="marker"/>
          <text x="12" y="22" fill="#617385" font-size="12">Presión total + Pf/Pb sobre la línea basal</text>
        </g>
      </svg>
      <div class="legend"><span><i class="dot" style="background:#111"></i>Presión central</span><span><i class="dot" style="background:#168038"></i>Pf</span><span><i class="dot" style="background:#ef6c00"></i>Pb</span><span><i class="dot" style="background:#6a1b9a"></i>Flujo estimado</span></div>
      <div class="controls"><button id="playBtn">Pausar</button><label class="toggle"><input id="smartPause" type="checkbox" checked> Pausa didáctica secuencial</label><input id="slider" type="range" min="0" max="0" value="0" step="1"><span class="timebox" id="timeBox">0 ms</span></div>
      <div class="alertCard" id="alertBox">
        <div class="alertHead"><span class="alertTitle" id="alertTitle">Métrica alterada</span><span class="alertPill" id="alertSeverity">ALERTA</span></div>
        <div class="alertLine"><b>Valor:</b> <span id="alertValue"></span></div>
        <div class="alertLine"><b>Criterio:</b> <span id="alertThreshold"></span></div>
        <div class="alertLine"><b>Por qué se marca:</b> <span id="alertWhy"></span></div>
        <div class="alertLine"><b>Mecanismo didáctico:</b> <span id="alertMechanism"></span></div>
      </div>
    </div>

    <div class="panel">
      <h3>Métricas reales integradas</h3>
      <div class="grid2">
        <div class="card"><div class="k">Presión instantánea</div><div class="v" id="instP">--</div></div>
        <div class="card"><div class="k">Flujo estimado</div><div class="v" id="instQ">--</div></div>
        <div class="card"><div class="k">Pf / Pb inst.</div><div class="v" id="instPfPb">--</div></div>
        <div class="card"><div class="k">RVSE calculado</div><div class="v" id="mRVSE"></div></div>
      </div>
      <svg viewBox="0 0 460 245" aria-label="Energía armónica relativa">
        <rect x="16" y="20" width="428" height="190" rx="15" fill="#fff" stroke="#d5e3ef"/>
        <text x="30" y="44" fill="#12355b" font-size="14" font-weight="800">Armónicos por FFT de la onda central real</text>
        <g id="bars" transform="translate(34 62)"></g>
        <text id="domH" x="30" y="232" fill="#617385" font-size="12"></text>
      </svg>
      <div class="small" id="detailBox"></div>
    </div>
  </div>
</div>
<script>
const data = {data_json};
const N = data.t.length;
let i = 0, playing = true;
const minP = Math.min(...data.p), maxP = Math.max(...data.p);
const maxPf = Math.max(...data.pf, 1e-6), maxPb = Math.max(...data.pb, 1e-6), maxQ = Math.max(...data.q, 1e-6);
const slider = document.getElementById('slider'); slider.max = Math.max(N-1,0);
const smartPause = document.getElementById('smartPause');
const alerts = Array.isArray(data.alerts) ? data.alerts.slice().sort((a,b)=>(a.pause_order||0)-(b.pause_order||0)) : [];
let shownAlerts = new Set();
let lastIdx = 0;
let currentCycle = 1;
function fmt(x,d=0) {{ if(!Number.isFinite(x)) return 'ND'; return x.toFixed(d); }}
function alertKey(a) {{ return 'cycle' + currentCycle + '|' + String(a.pause_order || 0) + '|' + String(a.label || ''); }}
function resetSmartPauseCycleIfNeeded(idx) {{
  // Al completar un nuevo ciclo cardíaco, las pausas didácticas vuelven a quedar disponibles.
  // Esto evita que la animación corra libremente después de haber mostrado una vez las alertas.
  if(idx < lastIdx) {{
    currentCycle += 1;
    shownAlerts = new Set();
    const box = document.getElementById('alertBox');
    if(box) box.classList.remove('show');
  }}
  lastIdx = idx;
}}
function showMetricAlert(a) {{
  const box = document.getElementById('alertBox');
  if(!box || !a) return;
  document.getElementById('alertTitle').textContent = a.label || 'Métrica alterada';
  document.getElementById('alertSeverity').textContent = a.severity || 'ALERTA';
  document.getElementById('alertValue').textContent = a.value || 'ND';
  document.getElementById('alertThreshold').textContent = a.threshold || 'criterio no disponible';
  document.getElementById('alertWhy').textContent = a.why || 'La métrica supera el límite definido.';
  document.getElementById('alertMechanism').textContent = a.mechanism || 'Interpretar de forma integrada con la curva central y el resto de métricas.';
  box.classList.add('show');
  try {{ box.scrollIntoView({{behavior:'smooth', block:'nearest'}}); }} catch(e) {{}}
}}
function checkSmartPause(idx) {{
  if(!smartPause || !smartPause.checked || !alerts.length) return;
  resetSmartPauseCycleIfNeeded(idx);
  // Buscar la próxima alerta no mostrada de este ciclo. El margen permite no saltearla
  // si el navegador pierde un cuadro. Como las alertas se separan por pause_order,
  // cada métrica alterada se muestra en una pausa propia.
  const hit = alerts.find(a => Math.abs((a.idx || 0) - idx) <= 2 && !shownAlerts.has(alertKey(a)));
  if(hit) {{
    shownAlerts.add(alertKey(hit));
    playing = false;
    const btn = document.getElementById('playBtn');
    if(btn) btn.textContent = 'Continuar';
    showMetricAlert(hit);
  }}
}}
if(smartPause && !alerts.length) {{
  smartPause.checked = false;
  smartPause.disabled = true;
  smartPause.parentElement.title = 'No se detectaron métricas alteradas o desviaciones relevantes para pausar automáticamente.';
}} else if(smartPause && alerts.length) {{
  smartPause.parentElement.title = 'Modo didáctico secuencial: pausa una vez por cada métrica alterada y reinicia las pausas en cada nuevo ciclo cardíaco.';
}}
function sx(idx) {{ return (data.t[idx] - data.t[0]) / Math.max(data.t[N-1]-data.t[0], 1e-6) * 630; }}
function sy(v) {{ return 130 - (v - minP) / Math.max(maxP-minP, 1e-6) * 110; }}
function syComp(v) {{ return 130 - v / Math.max(maxP-minP, 1e-6) * 95; }}
function linePath(vals, mode) {{
  let d='';
  for(let k=0;k<N;k++) {{ let x=sx(k), y=(mode==='comp'? syComp(vals[k]):sy(vals[k])); d += (k===0?'M':'L') + x.toFixed(1)+' '+y.toFixed(1)+' '; }}
  return d;
}}
document.getElementById('pressurePath').setAttribute('d', linePath(data.p, 'p'));
document.getElementById('pfPath').setAttribute('d', linePath(data.pf, 'comp'));
document.getElementById('pbPath').setAttribute('d', linePath(data.pb, 'comp'));
// Landmarks temporales idénticos a los usados para mover Pf/Pb en la aorta.
(function drawSyncLandmarks() {{
  const graph = document.getElementById('pressurePath').parentNode;
  const t0 = data.t[0], t1 = data.t[N-1];
  const marks = [
    {{id:'pbOnsetMark', t:data.sync.t_pb_onset_ms, color:'#ef6c00', dash:'2 4'}},
    {{id:'crossMark', t:data.sync.t_cross_ms, color:'#6a1b9a', dash:'6 4'}}
  ];
  marks.forEach(m => {{
    if(!Number.isFinite(m.t)) return;
    const x = (m.t-t0)/Math.max(t1-t0,1e-6)*630;
    const ln = document.createElementNS('http://www.w3.org/2000/svg','line');
    ln.setAttribute('x1',x); ln.setAttribute('x2',x); ln.setAttribute('y1',8); ln.setAttribute('y2',132);
    ln.setAttribute('stroke',m.color); ln.setAttribute('stroke-width','1.1'); ln.setAttribute('stroke-dasharray',m.dash); ln.setAttribute('opacity','.78');
    graph.insertBefore(ln, graph.firstChild);
  }});
}})();
document.getElementById('mHTA').textContent = data.metrics.hta;
document.getElementById('mPASPAD').textContent = data.metrics.pas.replace(' mmHg','') + '/' + data.metrics.pad.replace(' mmHg','');
document.getElementById('mRM').textContent = data.metrics.rm;
document.getElementById('mTref').textContent = data.metrics.tref;
document.getElementById('mRVSE').textContent = data.metrics.rvse_calc;
document.getElementById('detailBox').innerHTML = 'PAS central '+data.metrics.pas+'; PAD central '+data.metrics.pad+'; PP '+data.metrics.pp+'; P90 SAHA '+data.metrics.p90+'. IAu/AIx '+data.metrics.iau+'; P90 IAu/AIx '+data.metrics.iau_p90+'; Au '+data.metrics.au+'. Tfor '+data.metrics.tfor+'; Tref '+data.metrics.tref+'; inicio material Pb '+fmt(data.sync.t_pb_onset_ms,0)+' ms; cruce Pf=Pb '+fmt(data.sync.t_cross_ms,0)+' ms; PE estimado '+data.metrics.pe+'; Qp '+data.metrics.qp+'. Firma morfológica: <b>'+data.metrics.curve_id+'</b>.';
const bars = document.getElementById('bars');
const maxE = Math.max(...data.energies, 1e-6);
const maxAmp = Math.max(...data.amps.map(v => Math.abs(v)), 1e-6);
const harmonicBaseline = 86;
let domIdx = 0; data.energies.forEach((e, idx)=>{{ if(e>data.energies[domIdx]) domIdx=idx; }});
for(let b=0;b<10;b++) {{
  const x = b*39;
  const h = Math.max(2, data.energies[b]/maxE*54);
  const bg = document.createElementNS('http://www.w3.org/2000/svg','rect'); bg.setAttribute('x',x); bg.setAttribute('y',18); bg.setAttribute('width',25); bg.setAttribute('height',136); bg.setAttribute('rx',5); bg.setAttribute('class','barbg'); bars.appendChild(bg);
  const mid = document.createElementNS('http://www.w3.org/2000/svg','line'); mid.setAttribute('x1',x+1); mid.setAttribute('x2',x+24); mid.setAttribute('y1',harmonicBaseline); mid.setAttribute('y2',harmonicBaseline); mid.setAttribute('stroke','#9eb2c5'); mid.setAttribute('stroke-width','1'); mid.setAttribute('opacity','.80'); bars.appendChild(mid);
  const energy = document.createElementNS('http://www.w3.org/2000/svg','rect'); energy.setAttribute('id','energy'+b); energy.setAttribute('x',x+3); energy.setAttribute('y',harmonicBaseline-h); energy.setAttribute('width',19); energy.setAttribute('height',h); energy.setAttribute('rx',4); energy.setAttribute('class','barEnergy'); bars.appendChild(energy);
  const live = document.createElementNS('http://www.w3.org/2000/svg','rect'); live.setAttribute('id','liveBar'+b); live.setAttribute('x',x+8); live.setAttribute('y',harmonicBaseline-1); live.setAttribute('width',9); live.setAttribute('height',2); live.setAttribute('rx',4); live.setAttribute('class','barLive positive'); bars.appendChild(live);
  const dot = document.createElementNS('http://www.w3.org/2000/svg','circle'); dot.setAttribute('id','phaseDot'+b); dot.setAttribute('cx',x+12.5); dot.setAttribute('cy',harmonicBaseline); dot.setAttribute('r',3.2); dot.setAttribute('class','phaseDot'); bars.appendChild(dot);
  const tx = document.createElementNS('http://www.w3.org/2000/svg','text'); tx.setAttribute('x',x+12.5); tx.setAttribute('y',166); tx.setAttribute('text-anchor','middle'); tx.setAttribute('fill','#617385'); tx.setAttribute('font-size','11'); tx.textContent = 'H'+(b+1); bars.appendChild(tx);
  const ty = document.createElementNS('http://www.w3.org/2000/svg','text'); ty.setAttribute('x',x+12.5); ty.setAttribute('y',14); ty.setAttribute('text-anchor','middle'); ty.setAttribute('fill','#617385'); ty.setAttribute('font-size','10'); ty.textContent = fmt(data.energies[b],0)+'%'; bars.appendChild(ty);
}}
document.getElementById('domH').textContent = 'Barras sincronizadas: altura viva = contribución instantánea por amplitud y fase FFT. Dominante H'+(domIdx+1)+' | '+fmt(data.freqs[domIdx],2)+' Hz | energía '+fmt(data.energies[domIdx],1)+'%';
function harmonicValue(b, idx) {{
  const timeS = data.t[idx] / 1000.0;
  return data.amps[b] * Math.cos(2*Math.PI*data.freqs[b]*timeS + data.phases[b]);
}}
function updateHarmonicBars(idx) {{
  let liveVals = [];
  for(let b=0;b<10;b++) liveVals.push(harmonicValue(b, idx));
  for(let b=0;b<10;b++) {{
    const v = liveVals[b];
    const mag = Math.max(2, Math.abs(v)/maxAmp*62);
    const live = document.getElementById('liveBar'+b);
    const dot = document.getElementById('phaseDot'+b);
    const energy = document.getElementById('energy'+b);
    if(live) {{
      live.setAttribute('height', mag);
      live.setAttribute('y', v >= 0 ? harmonicBaseline-mag : harmonicBaseline);
      live.setAttribute('class', v >= 0 ? 'barLive positive' : 'barLive negative');
      live.setAttribute('opacity', b===domIdx ? '.98' : '.72');
    }}
    if(dot) {{ dot.setAttribute('cy', harmonicBaseline - Math.max(-58, Math.min(58, v/maxAmp*58))); }}
    if(energy) energy.setAttribute('class', b===domIdx ? 'barEnergy active' : 'barEnergy');
  }}
  const h1 = liveVals[0] || 0, h2 = liveVals[1] || 0, h3 = liveVals[2] || 0;
  document.getElementById('domH').textContent = 't '+fmt(data.t[idx],0)+' ms | H1 '+fmt(h1,2)+' mmHg, H2 '+fmt(h2,2)+' mmHg, H3 '+fmt(h3,2)+' mmHg | dominante energético H'+(domIdx+1)+' ('+fmt(data.energies[domIdx],1)+'%)';
}}
function update(idx) {{
  i = Math.max(0, Math.min(N-1, idx)); slider.value = i;
  const p = data.p[i], pf = data.pf[i], pb = data.pb[i], q = data.q[i];
  const norm = (p-minP)/Math.max(maxP-minP, 1e-6);
  const pulse = 1 + norm*0.115;
  const outerW = 56 + norm*18;
  const innerW = 32 + norm*11;
  const branchOuterW = 24 + norm*8;
  const branchInnerW = 12 + norm*5;
  ['aortaOuter'].forEach(id => {{ const el=document.getElementById(id); if(el) el.setAttribute('stroke-width', outerW); }});
  ['aortaInner'].forEach(id => {{ const el=document.getElementById(id); if(el) el.setAttribute('stroke-width', innerW); }});
  ['branch1Outer','branch2Outer','branch3Outer'].forEach(id => {{ const el=document.getElementById(id); if(el) el.setAttribute('stroke-width', branchOuterW); }});
  ['branch1Inner','branch2Inner','branch3Inner'].forEach(id => {{ const el=document.getElementById(id); if(el) el.setAttribute('stroke-width', branchInnerW); }});
  ['branch1bOuter'].forEach(id => {{ const el=document.getElementById(id); if(el) el.setAttribute('stroke-width', 19 + norm*6); }});
  ['branch1bInner'].forEach(id => {{ const el=document.getElementById(id); if(el) el.setAttribute('stroke-width', 9 + norm*4); }});
  const root = document.getElementById('rootPulse');
  if(root) root.setAttribute('transform', 'translate('+(-279*(pulse-1)).toFixed(2)+' '+(-178*(pulse-1)).toFixed(2)+') scale('+pulse.toFixed(3)+')');
  const valve = document.getElementById('valveRing');
  if(valve) valve.setAttribute('transform', 'rotate(28 279 199) scale('+pulse.toFixed(3)+')');
  const centerPath = document.getElementById('aortaCenter');
  const totalLen = centerPath.getTotalLength();
  const ti = data.t[i];
  const sync = data.sync || {{}};
  const foot = Number.isFinite(sync.t_foot_ms) ? sync.t_foot_ms : data.t[0];
  const pbOn = Number.isFinite(sync.t_pb_onset_ms) ? sync.t_pb_onset_ms : foot;
  const cross = Number.isFinite(sync.t_cross_ms) ? sync.t_cross_ms : (data.t[0]+data.t[N-1])/2;
  const meet = Number.isFinite(sync.meet_frac) ? Math.max(.20, Math.min(.85, sync.meet_frac)) : .62;
  const tEnd = data.t[N-1];
  function prog(x,a,b) {{ if(b<=a) return x>=b?1:0; return Math.max(0,Math.min(1,(x-a)/(b-a))); }}
  // Misma línea temporal que el gráfico: ambas ondas ocupan el mismo punto en t_cross.
  let pfFrac = ti<=cross ? meet*prog(ti,foot,cross) : meet+(1-meet)*prog(ti,cross,tEnd);
  let pbFrac = ti<=cross ? 1-(1-meet)*prog(ti,pbOn,cross) : meet*(1-prog(ti,cross,tEnd));
  const pfPt = centerPath.getPointAtLength(Math.max(0, Math.min(totalLen, pfFrac*totalLen)));
  const pbPt = centerPath.getPointAtLength(Math.max(0, Math.min(totalLen, pbFrac*totalLen)));
  const pfG = document.getElementById('pfWave');
  const pbG = document.getElementById('pbWave');
  if(pfG) {{
    pfG.setAttribute('transform', 'translate('+pfPt.x.toFixed(1)+' '+pfPt.y.toFixed(1)+')');
    pfG.style.opacity = ti>=foot ? '1' : '0';
    pfG.querySelector('circle').setAttribute('r', 9 + Math.sqrt(Math.max(pf,0)/maxPf)*13);
  }}
  if(pbG) {{
    pbG.setAttribute('transform', 'translate('+pbPt.x.toFixed(1)+' '+pbPt.y.toFixed(1)+')');
    pbG.style.opacity = ti>=pbOn ? '1' : '0';
    pbG.querySelector('circle').setAttribute('r', 8 + Math.sqrt(Math.max(pb,0)/maxPb)*12);
  }}
  document.getElementById('liveP').textContent = fmt(p,0)+' mmHg';
  const qOverlay = document.getElementById('liveQOverlay');
  if(qOverlay) qOverlay.textContent = fmt(q,0)+' mL/s';
  const pfPbOverlay = document.getElementById('livePfPbOverlay');
  if(pfPbOverlay) pfPbOverlay.textContent = fmt(pf,1)+' / '+fmt(pb,1)+' mmHg';
  const timeOverlay = document.getElementById('liveTimeOverlay');
  if(timeOverlay) timeOverlay.textContent = fmt(data.t[i],0)+' ms';
  const x = sx(i), y = sy(p);
  document.getElementById('cursor').setAttribute('x1', x); document.getElementById('cursor').setAttribute('x2', x);
  document.getElementById('pDot').setAttribute('cx', x); document.getElementById('pDot').setAttribute('cy', y);
  document.getElementById('pfDot').setAttribute('cx', x); document.getElementById('pfDot').setAttribute('cy', syComp(pf));
  document.getElementById('pbDot').setAttribute('cx', x); document.getElementById('pbDot').setAttribute('cy', syComp(pb));
  document.getElementById('timeBox').textContent = fmt(data.t[i],0)+' ms';
  document.getElementById('instP').textContent = fmt(p,0)+' mmHg';
  document.getElementById('instQ').textContent = fmt(q,0)+' mL/s';
  document.getElementById('instPfPb').textContent = fmt(pf,1)+' / '+fmt(pb,1);
  updateHarmonicBars(i);
  checkSmartPause(i);
}}
function loop() {{ if(playing) update((i+1)%N); window.setTimeout(loop, 45); }}
document.getElementById('playBtn').onclick = () => {{
  playing = !playing;
  document.getElementById('playBtn').textContent = playing ? 'Pausar' : 'Reproducir';
  if(playing && alerts.length && smartPause && smartPause.checked) {{
    const box = document.getElementById('alertBox');
    if(box) box.classList.remove('show');
    update(Math.min(N-1, i+1));
  }}
}};
slider.oninput = (e) => {{ playing=false; document.getElementById('playBtn').textContent='Reproducir'; update(parseInt(e.target.value)); }};
update(0); loop();
</script>
</body>
</html>
"""
    return html
  except Exception as e:
    return f"<div style='padding:12px;border:1px solid #e0e0e0;border-radius:12px;color:#7a1f1f;background:#fff5f5'>No se pudo generar la animación hemodinámica real: {pdf_text(e)}</div>"


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


def plot_clinical_gauges(row, ppa=None):
  """Semaforización de pantalla derivada del mismo estado canónico del informe."""
  state = build_canonical_diagnostic_state(row, {}, None)
  d = state["domains"]

  pas = to_float(d["hta_central"].get("value")); pp = to_float(d["carga_pulsatil"].get("value"))
  iau = to_float(d["aumentacion"].get("value")); ppa_val = to_float(d["amplificacion"].get("value"))
  pas_ref = to_float(d["hta_central"].get("threshold"))
  if np.isnan(pas_ref): pas_ref = 130.0
  iau_ref = to_float(d["aumentacion"].get("p90"))

  # Índice >1 significa mayor desviación adversa respecto del límite en todos los casos.
  metrics = [
    ("PAS central", pas / pas_ref if not np.isnan(pas) and pas_ref > 0 else np.nan),
    ("PP central", pp / 50.0 if not np.isnan(pp) else np.nan),
    ("IAu/AIx", iau / iau_ref if not np.isnan(iau) and not np.isnan(iau_ref) and iau_ref > 0 else np.nan),
    ("PPA", 1.30 / ppa_val if not np.isnan(ppa_val) and ppa_val > 0 else np.nan),
  ]
  names = [x[0] for x in metrics]
  score = [x[1] for x in metrics]
  fig, ax = plt.subplots(figsize=(7.6,3.9))
  ax.barh(names, score, color="#546E7A")
  ax.axvline(1, linestyle="--", linewidth=1.2, color="#B71C1C")
  _apply_professional_axes(ax, "Semaforización clínica de parámetros centrales", "Índice de desviación respecto del límite", "")
  ax.grid(axis="x", alpha=.22)
  return fig_to_png(fig)

def interpret_pressure_central_metrics(row, dx, cat, ref, amp_sbp, ppa, risk):
  """Conclusión textual de pantalla desde el estado canónico único."""
  pas_c = to_float(row.get("pas_central")); pad_c = to_float(row.get("pad_central")); pp_c = to_float(row.get("pp_central"))
  pam_c = to_float(row.get("pam_central")); iau = to_float(row.get("iau")); au = to_float(row.get("au")); fc = to_float(row.get("fc"))
  state = build_canonical_diagnostic_state(row, {}, None)
  d = state["domains"]

  def fmt(v, dec=1):
    try:
      f = float(v)
      if np.isnan(f): return "no disponible"
      return f"{f:.{dec}f}"
    except Exception:
      return "no disponible"

  synthesis = "; ".join([
    d["hta_central"]["short"],
    d["carga_pulsatil"]["short"],
    d["amplificacion"]["short"],
    d["aumentacion"]["short"],
  ])

  return (
    f"El análisis de presión central muestra PAS central {fmt(pas_c,0)} mmHg, PAD central {fmt(pad_c,0)} mmHg, "
    f"PAM central {fmt(pam_c,0)} mmHg y PP central {fmt(pp_c,0)} mmHg. La categoría tensional periférica/braquial es {cat}. "
    f"La amplificación PAS periférico-central es {fmt(amp_sbp,1)} mmHg y la PPA es {fmt(d['amplificacion'].get('value'),2)}. "
    f"Au: {fmt(au,1)} mmHg (descriptivo), IAu: {fmt(iau,1)}%, FC: {fmt(fc,0)} lpm. "
    f"Conclusión diagnóstica: {dx} "
    f"Síntesis canónica: {synthesis}."
  )

def interpret_harmonic_profile(hdf):
  """Interpretación descriptiva del análisis armónico sin cortes clínicos universales."""
  try:
    if hdf is None or len(hdf) == 0:
      return "No fue posible calcular un perfil armónico interpretable por falta de datos válidos de la onda central."
    m = harmonic_distortion_metrics(hdf)
    if not m.get("ok"):
      return "No fue posible estimar en forma estable HD ni la distribución espectral de la onda central."
    hd = m.get("hd_percent", np.nan)
    h1 = m.get("h1_energy_percent", np.nan)
    h2 = m.get("h2_energy_percent", np.nan)
    h4p = m.get("h4plus_energy_percent", np.nan)
    domf = m.get("dominant_frequency_hz", np.nan)
    return (
      f"El análisis armónico muestra H1 {h1:.1f}%, H2 {h2:.1f}% y energía acumulada desde H4 {h4p:.1f}%. "
      f"La distorsión armónica energética (HD = energía por encima de la fundamental / energía fundamental) es {hd:.1f}%. "
      f"La frecuencia dominante es {domf:.2f} Hz. La referencia cercana a 3,5 Hz se conserva solo como contexto mecanístico, "
      "no como límite normal/patológico. HD, H1, H2, H4+ y frecuencia dominante se informan como métricas continuas "
      "de caracterización espectral, sin aplicar los antiguos cortes H1 <35% o H4+ >=25% ni un umbral clínico universal de HD."
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
    grade = "conservado"
    meaning = "no cumple criterio de reducción franca; interpretar junto con frecuencia cardíaca, presión diastólica central y poscarga sistólica"
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



def build_canonical_diagnostic_state(row, sep_metrics=None, hdf=None):
  """Única fuente de verdad diagnóstica para toda la app PAC.

  Este diccionario canónico alimenta pantalla, PDF, conclusiones y fenotipo.
  No usa puntajes acumulativos ni doble conteo entre dominios.

  Dominios clínicos independientes:
  - HTA central: PAS central vs referencia SAHA P90 / respaldo operativo.
  - Carga pulsátil: PP central, independiente de la presencia de HTA central.
  - Amplificación: PPA = PP radial / PP central.
  - Aumentación: IAu/AIx vs P90 por edad/sexo.
  - Reflexión: RM por percentiles de edad/método; RI complementaria.
  - Reserva subendocárdica: RVSE/SEVR calculado.

  Dominio espectral cuantificado:
  - Armónicos/HD: métricas continuas de caracterización espectral; nunca definen por sí solas un fenotipo clínico.
  """
  sep_metrics = sep_metrics or {}

  def _domain(status, altered, label, short, detail, value=np.nan, criterion="", grade=None, classificable=True, extra=None):
    d = {
      "status": status,
      "altered": bool(altered),
      "label": label,
      "short": short,
      "detail": detail,
      "value": value,
      "criterion": criterion,
      "grade": grade or status,
      "classifiable": bool(classificable),
    }
    if extra:
      d.update(extra)
    return d

  # 1) HTA central: completamente separada de carga pulsátil.
  hta = central_hypertension_status(row)
  if hta.get("ok"):
    if hta.get("tiene_hta_central"):
      hta_d = _domain(
        "alterada", True, "Hipertensión central", "con hipertensión central",
        "PAS central por encima del límite diagnóstico aplicable.",
        hta.get("pas_central", np.nan), hta.get("criterio", ""), "alterada",
        extra={"threshold": hta.get("umbral", np.nan)}
      )
    else:
      hta_d = _domain(
        "normal", False, "Hipertensión central", "sin hipertensión central",
        "PAS central por debajo del límite diagnóstico aplicable.",
        hta.get("pas_central", np.nan), hta.get("criterio", ""), "normal",
        extra={"threshold": hta.get("umbral", np.nan)}
      )
  else:
    hta_d = _domain(
      "sin_datos", False, "Hipertensión central", "hipertensión central no clasificable",
      "No hay PAS central válida o referencia aplicable suficiente.",
      np.nan, hta.get("criterio", "PAS central no disponible"), "sin_datos", False
    )

  # 2) Carga pulsátil: SOLO PP central; nunca se altera por tener HTA central.
  pp_c = to_float(row.get("pp_central"))
  if np.isnan(pp_c):
    pulse_d = _domain(
      "sin_datos", False, "Carga pulsátil central", "carga pulsátil no clasificable",
      "PP central no disponible.", np.nan, "requiere PP central", "sin_datos", False
    )
  elif pp_c >= 60:
    pulse_d = _domain(
      "alterada", True, "Carga pulsátil central", "carga pulsátil central alta",
      "PP central alta.", pp_c, ">=60 mmHg", "alta"
    )
  elif pp_c >= 50:
    pulse_d = _domain(
      "alterada", True, "Carga pulsátil central", "carga pulsátil central aumentada",
      "PP central aumentada.", pp_c, ">=50 y <60 mmHg", "aumentada"
    )
  else:
    pulse_d = _domain(
      "normal", False, "Carga pulsátil central", "carga pulsátil central no aumentada",
      "PP central por debajo del umbral operativo de aumento.", pp_c, "<50 mmHg", "normal"
    )

  # 3) Amplificación periférico-central: dominio independiente.
  pp_r = to_float(row.get("pp_radial"))
  ppa = pp_r / pp_c if not np.isnan(pp_r) and not np.isnan(pp_c) and pp_c > 0 else np.nan
  if np.isnan(ppa):
    ppa_d = _domain(
      "sin_datos", False, "Amplificación periférico-central", "amplificación no clasificable",
      "No se dispone de PP radial y PP central válidas.", np.nan,
      "PPA = PP radial / PP central", "sin_datos", False
    )
  elif ppa < 1.20:
    ppa_d = _domain(
      "alterada", True, "Amplificación periférico-central", "amplificación marcadamente reducida",
      "Pérdida marcada de amplificación periférico-central.", ppa, "<1,20", "marcada"
    )
  elif ppa < 1.30:
    ppa_d = _domain(
      "alterada", True, "Amplificación periférico-central", "amplificación reducida",
      "Amplificación periférico-central reducida.", ppa, "1,20 a <1,30", "reducida"
    )
  else:
    ppa_d = _domain(
      "normal", False, "Amplificación periférico-central", "amplificación conservada",
      "Amplificación periférico-central conservada.", ppa, ">=1,30", "normal"
    )

  # 4) Aumentación: solo referencia P90 edad/sexo; Au aislado no diagnostica.
  iau = to_float(row.get("iau")); au = to_float(row.get("au"))
  aix = get_saha_aix75_reference(row)
  if aix.get("ok"):
    if aix.get("alterada"):
      aug_d = _domain(
        "alterada", True, "Aumentación central", "aumentación central aumentada",
        "IAu/AIx alcanza o supera P90 para edad y sexo.", iau,
        f">=P90 ({aix.get('p90', np.nan):.1f}%)", "alterada",
        extra={"percentile": aix.get("percentil", np.nan), "p90": aix.get("p90", np.nan), "au": au}
      )
    else:
      aug_d = _domain(
        "normal", False, "Aumentación central", "aumentación central no aumentada",
        "IAu/AIx permanece por debajo de P90 para edad y sexo.", iau,
        f"<P90 ({aix.get('p90', np.nan):.1f}%)", "normal",
        extra={"percentile": aix.get("percentil", np.nan), "p90": aix.get("p90", np.nan), "au": au}
      )
  elif not np.isnan(iau):
    aug_d = _domain(
      "sin_datos", False, "Aumentación central", "aumentación central no clasificable",
      "IAu/AIx disponible, pero sin referencia válida de edad/sexo; no se usan cortes fijos 25/35%.",
      iau, "requiere referencia de edad/sexo", "sin_datos", False, extra={"au": au}
    )
  elif not np.isnan(au):
    aug_d = _domain(
      "sin_datos", False, "Aumentación central", "aumentación central no clasificable",
      "Au se informa en mmHg como descriptor; no define diagnóstico aislado.",
      au, "Au descriptivo", "sin_datos", False
    )
  else:
    aug_d = _domain(
      "sin_datos", False, "Aumentación central", "aumentación central no clasificable",
      "IAu/AIx no disponible.", np.nan, "requiere IAu/AIx", "sin_datos", False
    )

  # 5) Reflexión: RM primaria percentilar; RI complementaria; Tref continuo.
  rm = sep_metrics.get("rm", np.nan); ri = sep_metrics.get("ri", np.nan)
  tref = sep_metrics.get("tref_ms", np.nan); ratio = sep_metrics.get("tfor_tref", np.nan)
  rmri = get_rm_ri_reference(row, sep_metrics)
  if rmri.get("ok") and rmri.get("rm_clasif", {}).get("ok"):
    c = rmri["rm_clasif"]
    grade = c.get("grado", "normal")
    altered = bool(c.get("alterada", False))
    status = "alterada" if altered else ("intermedia" if grade == "intermedia" else "normal")
    short = (
      "reflexión de onda marcadamente aumentada" if grade == "marcada" else
      "reflexión de onda aumentada" if altered else
      "reflexión relativamente elevada sin criterio diagnóstico" if grade == "intermedia" else
      "reflexión de onda no aumentada"
    )
    reflection_d = _domain(
      status, altered, "Reflexión de onda", short,
      f"RM clasificada por percentiles específicos de edad y método: {c.get('categoria', grade)}.",
      rm,
      f"P75 {rmri['rm_ref']['p75']:.2f}; P90 {rmri['rm_ref']['p90']:.2f}; P95 {rmri['rm_ref']['p95']:.2f}",
      grade, True,
      extra={
        "percentile": c.get("percentil", np.nan), "ri": ri, "tref_ms": tref,
        "tfor_tref": ratio, "rm_reference": rmri.get("rm_ref"), "ri_reference": rmri.get("ri_ref")
      }
    )
  elif not np.isnan(rm):
    reflection_d = _domain(
      "sin_datos", False, "Reflexión de onda", "reflexión de onda no clasificable",
      "RM disponible, pero sin referencia válida de edad/método.", rm,
      "requiere edad y método de referencia", "sin_datos", False,
      extra={"ri": ri, "tref_ms": tref, "tfor_tref": ratio}
    )
  else:
    reflection_d = _domain(
      "sin_datos", False, "Reflexión de onda", "reflexión de onda no clasificable",
      "No hay separación Pf/Pb suficiente para clasificar RM.", np.nan,
      "requiere RM válida", "sin_datos", False,
      extra={"ri": ri, "tref_ms": tref, "tfor_tref": ratio}
    )

  # 6) Reserva subendocárdica: independiente; nunca suma a otro dominio.
  rvse = sep_metrics.get("rvse_calculado_%", np.nan)
  if np.isnan(rvse):
    rvse_d = _domain(
      "sin_datos", False, "Reserva subendocárdica", "reserva subendocárdica no clasificable",
      "RVSE/SEVR calculado no disponible.", np.nan, "requiere RVSE calculado", "sin_datos", False
    )
  elif rvse < 120:
    rvse_d = _domain(
      "alterada", True, "Reserva subendocárdica", "reserva subendocárdica reducida",
      "RVSE/SEVR calculado reducido.", rvse, "<120%", "reducida"
    )
  else:
    rvse_d = _domain(
      "normal", False, "Reserva subendocárdica", "reserva subendocárdica conservada",
      "RVSE/SEVR calculado conservado.", rvse, ">=120%", "normal"
    )

  # 7) Armónicos: dominio espectral cuantificado; nunca alteración clínica transversal.
  hm = harmonic_distortion_metrics(hdf)
  if hm.get("ok"):
    harm_d = _domain(
      "cuantificado", False, "Análisis armónico", "perfil armónico cuantificado",
      "HD, H1, H2, H4+ y frecuencia dominante se informan como métricas continuas de caracterización espectral; HD y H4+ se calculan sobre el espectro positivo completo disponible.",
      hm.get("hd_percent", np.nan), "métricas continuas; sin umbral clínico universal", "cuantificado", False,
      extra={
        "metrics": hm,
        "quantified": True,
        "diagnostic_use": False,
        "defines_phenotype": False,
      }
    )
  else:
    harm_d = _domain(
      "sin_datos", False, "Análisis armónico", "análisis armónico no clasificable",
      "FFT/HD no disponibles o inestables.", np.nan, "requiere análisis espectral válido", "sin_datos", False
    )

  domains = {
    "hta_central": hta_d,
    "carga_pulsatil": pulse_d,
    "amplificacion": ppa_d,
    "aumentacion": aug_d,
    "reflexion": reflection_d,
    "rvse": rvse_d,
    "armonicos": harm_d,
  }
  clinical_keys = ["hta_central", "carga_pulsatil", "amplificacion", "aumentacion", "reflexion", "rvse"]
  altered_keys = [k for k in clinical_keys if domains[k].get("altered")]
  unclassifiable_keys = [k for k in clinical_keys if not domains[k].get("classifiable", True)]

  return {
    "domains": domains,
    "clinical_keys": clinical_keys,
    "altered_keys": altered_keys,
    "unclassifiable_keys": unclassifiable_keys,
    "count_altered": len(altered_keys),
    "partially_classifiable": bool(unclassifiable_keys),
  }


def _phenotype_from_canonical_state(state):
  """Construye un fenotipo determinista exclusivamente desde los estados canónicos."""
  d = state.get("domains", {})
  altered = state.get("altered_keys", [])

  if not altered:
    if state.get("partially_classifiable"):
      return (
        "Fenotipo vascular central sin alteraciones en los dominios clasificables",
        "no se identifican dominios clínicos alterados entre los evaluables; existen dominios no clasificables por datos insuficientes"
      )
    return (
      "Fenotipo vascular central conservado en los dominios evaluados",
      "sin hipertensión central, sin aumento de carga pulsátil, con amplificación conservada, sin aumentación ni reflexión aumentadas y con reserva subendocárdica conservada"
    )

  isolated_map = {
    "hta_central": ("Fenotipo de hipertensión central aislada", "hipertensión central sin otras alteraciones clínicas dominantes"),
    "carga_pulsatil": ("Fenotipo de carga pulsátil central aumentada aislada", "PP central aumentada sin otras alteraciones clínicas dominantes"),
    "amplificacion": ("Fenotipo de amplificación periférico-central reducida aislada", "PPA reducida sin otras alteraciones clínicas dominantes"),
    "aumentacion": ("Fenotipo de aumentación central aumentada aislada", "IAu/AIx >= P90 sin otras alteraciones clínicas dominantes"),
    "reflexion": ("Fenotipo reflectivo predominante aislado", "RM >= P90 para edad/método sin otras alteraciones clínicas dominantes"),
    "rvse": ("Fenotipo con reserva subendocárdica relativa reducida aislada", "RVSE/SEVR <120% sin otras alteraciones clínicas dominantes"),
  }
  if len(altered) == 1:
    return isolated_map[altered[0]]

  names = {
    "hta_central": "hipertensión central",
    "carga_pulsatil": "carga pulsátil central aumentada",
    "amplificacion": "amplificación periférico-central reducida",
    "aumentacion": "aumentación central aumentada",
    "reflexion": "reflexión de onda aumentada",
    "rvse": "reserva subendocárdica reducida",
  }
  ordered = [k for k in state.get("clinical_keys", []) if k in altered]
  components = [names[k] for k in ordered]

  # Nombre principal explícito, sin inferir rigidez no medida.
  if "reflexion" in altered and ("hta_central" in altered or "carga_pulsatil" in altered):
    base = "Fenotipo de presión/carga central alterada con reflexión aumentada"
  elif "hta_central" in altered and "carga_pulsatil" in altered:
    base = "Fenotipo de hipertensión central con carga pulsátil aumentada"
  elif "reflexion" in altered:
    base = "Fenotipo reflectivo combinado"
  elif "hta_central" in altered or "carga_pulsatil" in altered:
    base = "Fenotipo de presión/carga central alterada combinado"
  else:
    base = "Fenotipo vascular central combinado"

  modifiers = []
  if "amplificacion" in altered: modifiers.append("amplificación reducida")
  if "aumentacion" in altered: modifiers.append("aumentación aumentada")
  if "rvse" in altered: modifiers.append("reserva subendocárdica reducida")
  phenotype = base + (" con " + ", ".join(modifiers) if modifiers else "")
  mechanism = "; ".join(components)
  return phenotype, mechanism

def _didactic_grade_pressure(row):
  """Conclusiones de HTA central y carga pulsátil desde el estado canónico único."""
  state = build_canonical_diagnostic_state(row, {}, None)
  d = state["domains"]
  hta = d["hta_central"]
  pulse = d["carga_pulsatil"]
  ppa = d["amplificacion"]

  if hta["status"] == "alterada":
    pressure = "El paciente se clasifica CON HIPERTENSIÓN CENTRAL según el criterio diagnóstico aplicable."
    pressure_short, pressure_level = "con hipertensión central.", "alterada"
  elif hta["status"] == "normal":
    pressure = "El paciente se clasifica SIN HIPERTENSIÓN CENTRAL según el criterio diagnóstico aplicable."
    pressure_short, pressure_level = "sin hipertensión central.", "normal"
  else:
    pressure = "La hipertensión central no pudo clasificarse por datos insuficientes."
    pressure_short, pressure_level = "hipertensión central no clasificable.", "sin_datos"

  if pulse["status"] == "sin_datos" and ppa["status"] == "sin_datos":
    pulse_text = "La carga pulsátil central y la amplificación periférico-central no pudieron clasificarse por datos insuficientes."
    pulse_short, pulse_level = "carga pulsátil y amplificación no clasificables.", "sin_datos"
  elif pulse["altered"] and ppa["altered"]:
    pulse_text = "El estudio evidencia aumento de la carga pulsátil central junto con amplificación periférico-central reducida."
    pulse_short, pulse_level = "carga pulsátil central aumentada con amplificación reducida.", "alterada"
  elif pulse["altered"]:
    pulse_text = "El estudio evidencia aumento de la carga pulsátil central; la amplificación periférico-central no está reducida cuando es clasificable."
    pulse_short, pulse_level = "carga pulsátil central aumentada.", "alterada"
  elif ppa["altered"]:
    pulse_text = "El estudio evidencia amplificación periférico-central reducida sin aumento de la carga pulsátil central."
    pulse_short, pulse_level = "amplificación reducida sin aumento de carga pulsátil central.", "alterada"
  elif pulse["status"] == "normal" and ppa["status"] == "normal":
    pulse_text = "El estudio no evidencia aumento de la carga pulsátil central y mantiene amplificación periférico-central conservada."
    pulse_short, pulse_level = "carga pulsátil no aumentada con amplificación conservada.", "normal"
  else:
    # Uno de los dos dominios es clasificable y normal; el otro carece de datos.
    pulse_text = "No se evidencia alteración en el componente clasificable de carga pulsátil/amplificación, aunque la evaluación es parcial por datos insuficientes."
    pulse_short, pulse_level = "carga pulsátil/amplificación parcialmente clasificable sin alteración demostrada.", "sin_datos"

  return pressure, pressure_short, pressure_level, pulse_text, pulse_short, pulse_level

def _didactic_grade_augmentation(row):
  """Aumentación central desde el estado canónico único."""
  aug = build_canonical_diagnostic_state(row, {}, None)["domains"]["aumentacion"]
  if aug["status"] == "alterada":
    return (
      "El estudio se clasifica CON AUMENTACIÓN CENTRAL AUMENTADA para edad y sexo (IAu/AIx >= P90).",
      "con aumentación central aumentada para edad y sexo.",
      "alterada",
    )
  if aug["status"] == "normal":
    return (
      "El estudio se clasifica SIN AUMENTACIÓN CENTRAL AUMENTADA para edad y sexo (IAu/AIx < P90).",
      "sin aumentación central aumentada para edad y sexo.",
      "normal",
    )
  return (
    aug.get("detail", "La aumentación central no pudo clasificarse."),
    "aumentación central no clasificable.",
    "sin_datos",
  )

def _didactic_grade_rvse(sep_metrics):
  rvse = build_canonical_diagnostic_state({}, sep_metrics, None)["domains"]["rvse"]
  if rvse["status"] == "alterada":
    return (
      "La reserva subendocárdica se interpreta como reducida, sugiriendo menor balance relativo entre perfusión diastólica y demanda sistólica.",
      "reserva subendocárdica reducida.",
      "alterada",
    )
  if rvse["status"] == "normal":
    return (
      "La reserva subendocárdica se interpreta como conservada, con balance presión-tiempo globalmente favorable en la curva analizada.",
      "reserva subendocárdica conservada.",
      "normal",
    )
  return (
    "La reserva subendocárdica no pudo estimarse de forma estable por datos insuficientes.",
    "reserva subendocárdica no clasificable.",
    "sin_datos",
  )

def _didactic_grade_wave(sep_metrics, row=None):
  """Separación de ondas desde el estado canónico único."""
  ref = build_canonical_diagnostic_state(row or {}, sep_metrics, None)["domains"]["reflexion"]
  tref = ref.get("tref_ms", np.nan)
  extra = f" Tref {tref:.0f} ms se informa como variable continua, sin corte fijo universal." if not np.isnan(tref) else ""

  if ref["status"] == "alterada":
    return (
      f"La magnitud de reflexión se clasifica como {ref.get('grade','aumentada')} por percentiles específicos de edad y método; RI es complementario y no suma evidencia independiente.{extra}",
      "reflexión de onda aumentada por RM percentilar.",
      "alterada",
    )
  if ref["status"] == "intermedia":
    return (
      "La RM se ubica entre P75 y P90: relativamente elevada para edad/método, sin alcanzar el criterio diagnóstico P90. RI es complementario." + extra,
      "reflexión relativamente elevada, sin criterio de aumento.",
      "intermedia",
    )
  if ref["status"] == "normal":
    return (
      "La magnitud de reflexión permanece por debajo de P75 para edad y método, sin evidencia percentilar de reflexión aumentada; RI es complementario." + extra,
      "reflexión de onda no aumentada.",
      "normal",
    )
  return (
    "La separación Pf/Pb está disponible de forma insuficiente para una clasificación percentilar válida de RM. RI y Tref no generan diagnóstico independiente.",
    "reflexión de onda no clasificable por referencia.",
    "sin_datos",
  )

def _didactic_grade_harmonics(hdf):
  """Conclusión didáctica armónica cuantificada, sin binarizar normal/anormal."""
  try:
    m = harmonic_distortion_metrics(hdf)
    if not m.get("ok"):
      raise ValueError("sin datos armónicos")
    hd = format_optional(m.get("hd_percent"), 1)
    h1 = format_optional(m.get("h1_energy_percent"), 1)
    h2 = format_optional(m.get("h2_energy_percent"), 1)
    h4p = format_optional(m.get("h4plus_energy_percent"), 1)
    domf = format_optional(m.get("dominant_frequency_hz"), 2)
    return (
      f"El análisis armónico cuantificado sobre la curva central real muestra HD energético {hd}%, "
      f"H1 {h1}%, H2 {h2}%, H4+ {h4p}% y frecuencia dominante {domf} Hz. "
      "Estas métricas se interpretan como variables continuas de caracterización espectral; no se utilizan aisladamente "
      "para clasificar normalidad/anormalidad, diagnosticar rigidez arterial ni definir el fenotipo vascular integrado.",
      f"perfil armónico cuantificado: HD {hd}%, H1 {h1}%, H2 {h2}%, H4+ {h4p}%, frecuencia dominante {domf} Hz.",
      "cuantificado",
    )
  except Exception:
    return (
      "El análisis armónico no pudo caracterizarse de forma estable; se recomienda revisar el gráfico espectral y la calidad de la curva digitalizada.",
      "análisis armónico no clasificable.",
      "sin_datos",
    )

def _didactic_phenotype_text(phenotype):
  """Texto didáctico compatible con los fenotipos canónicos; sin fenotipos armónicos ni 'rígido-reflectivo'."""
  p = safe_text(phenotype).lower()
  if "conservado en los dominios evaluados" in p:
    return (
      "La integración global no identifica alteraciones clínicas en los dominios evaluados.",
      "fenotipo vascular central conservado en los dominios evaluados."
    )
  if "sin alteraciones en los dominios clasificables" in p:
    return (
      "La integración global no identifica alteraciones en los dominios clasificables, aunque existen componentes no evaluables por datos insuficientes.",
      "fenotipo sin alteraciones en los dominios clasificables."
    )
  if "reserva subendocárdica relativa reducida aislada" in p:
    return (
      "La integración global identifica reducción aislada de la reserva subendocárdica relativa, sin otras alteraciones clínicas dominantes.",
      "fenotipo con reserva subendocárdica relativa reducida aislada."
    )
  if "hipertensión central aislada" in p:
    return ("La integración global identifica hipertensión central aislada.", "fenotipo de hipertensión central aislada.")
  if "carga pulsátil central aumentada aislada" in p:
    return ("La integración global identifica aumento aislado de la carga pulsátil central.", "fenotipo de carga pulsátil central aumentada aislada.")
  if "amplificación periférico-central reducida aislada" in p:
    return ("La integración global identifica reducción aislada de la amplificación periférico-central.", "fenotipo de amplificación periférico-central reducida aislada.")
  if "aumentación central aumentada aislada" in p:
    return ("La integración global identifica aumentación central aumentada aislada.", "fenotipo de aumentación central aumentada aislada.")
  if "reflectivo predominante aislado" in p:
    return ("La integración global identifica reflexión de onda aumentada aislada por RM percentilar.", "fenotipo reflectivo predominante aislado.")
  # Todo fenotipo combinado se describe con su denominación canónica exacta.
  if "fenotipo" in p:
    return (
      "La integración global identifica una combinación coherente de los dominios clínicos alterados consignados en la conclusión integrada.",
      safe_text(phenotype).rstrip(".") + "."
    )
  return (
    "La integración global no pudo caracterizarse de forma completa.",
    "fenotipo vascular central no clasificable."
  )



def build_automatic_integrated_conclusion(row, sep_metrics=None, hdf=None):
  """Conclusión integrada generada exclusivamente desde el estado canónico único."""
  sep_metrics = sep_metrics or {}
  state = build_canonical_diagnostic_state(row, sep_metrics, hdf)
  d = state["domains"]
  phenotype, mechanism = _phenotype_from_canonical_state(state)

  def _fmt(v, dec=1, unit=""):
    try:
      f = float(v)
      if np.isnan(f): return "no disponible"
      return f"{f:.{dec}f}{unit}"
    except Exception:
      return "no disponible"

  # Seis dominios clínicos independientes. PP/PPA/RVSE están siempre incluidos.
  hta = d["hta_central"]
  pulse = d["carga_pulsatil"]
  ppa = d["amplificacion"]
  aug = d["aumentacion"]
  refl = d["reflexion"]
  rvse = d["rvse"]
  harm = d["armonicos"]

  domain_summary = (
    f"HTA central: {hta['short']}; "
    f"carga pulsátil: {pulse['short']}; "
    f"amplificación: {ppa['short']}; "
    f"aumentación: {aug['short']}; "
    f"reflexión: {refl['short']}; "
    f"reserva subendocárdica: {rvse['short']}"
  )

  technical = []
  if not np.isnan(to_float(hta.get("value"))): technical.append(f"PASc {_fmt(hta.get('value'),0,' mmHg')}")
  if not np.isnan(to_float(pulse.get("value"))): technical.append(f"PPc {_fmt(pulse.get('value'),0,' mmHg')}")
  if not np.isnan(to_float(ppa.get("value"))): technical.append(f"PPA {_fmt(ppa.get('value'),2)}")
  if not np.isnan(to_float(aug.get("value"))): technical.append(f"IAu/AIx {_fmt(aug.get('value'),1,'%')}")
  if not np.isnan(to_float(refl.get("value"))): technical.append(f"RM {_fmt(refl.get('value'),2)}")
  ri = refl.get("ri", np.nan)
  if not np.isnan(to_float(ri)): technical.append(f"RI {_fmt(ri,2)} complementario")
  tref = refl.get("tref_ms", np.nan)
  if not np.isnan(to_float(tref)): technical.append(f"Tref {_fmt(tref,0,' ms')} continuo")
  if not np.isnan(to_float(rvse.get("value"))): technical.append(f"RVSE {_fmt(rvse.get('value'),1,'%')}")

  if harm["status"] == "cuantificado":
    hm = harm.get("metrics", {})
    harmonic_phrase = (
      f"Análisis armónico: HD energético {_fmt(hm.get('hd_percent'),1,'%')}, "
      f"H1 {_fmt(hm.get('h1_energy_percent'),1,'%')}, "
      f"H2 {_fmt(hm.get('h2_energy_percent'),1,'%')}, "
      f"H4+ {_fmt(hm.get('h4plus_energy_percent'),1,'%')}, "
      f"frecuencia dominante {_fmt(hm.get('dominant_frequency_hz'),2,' Hz')}. "
      "Métricas continuas de caracterización espectral, sin clasificación patológica transversal por un corte universal."
    )
  else:
    harmonic_phrase = "Análisis armónico no clasificable por datos insuficientes."

  if state["count_altered"] == 0:
    global_phrase = (
      "sin alteraciones clínicas en los dominios evaluados" if not state["partially_classifiable"]
      else "sin alteraciones en los dominios clasificables, con evaluación parcial por datos insuficientes"
    )
  elif state["count_altered"] == 1:
    global_phrase = "con una alteración clínica dominante aislada"
  else:
    global_phrase = f"con {state['count_altered']} dominios clínicos alterados"

  conclusion = (
    f"Conclusión integrada automática: {domain_summary}. "
    f"En conjunto, configura {phenotype}; mecanismo: {mechanism}. "
    f"El perfil global queda {global_phrase}. "
    + ("Valores de control: " + "; ".join(technical) + ". " if technical else "")
    + harmonic_phrase + " "
    "RI es complementario de RM; Tref y Tfor/Tref son continuos; los armónicos no suman evidencia diagnóstica ni definen rigidez por sí solos."
  )

  return conclusion, {
    "canonical_state": state,
    "hta_central_alterada": d["hta_central"]["altered"],
    "carga_pulsatil_alterada": d["carga_pulsatil"]["altered"],
    "amplificacion_alterada": d["amplificacion"]["altered"],
    "aumentacion_alterada": d["aumentacion"]["altered"],
    "reflexion_alterada": d["reflexion"]["altered"],
    "rvse_alterado": d["rvse"]["altered"],
    "armonicos_alterados": False,
    "armonicos_cuantificados": d["armonicos"]["status"] == "cuantificado",
    "cantidad_dominios_alterados": state["count_altered"],
    "frase_global": global_phrase,
    "fenotipo": phenotype,
  }

def build_continuous_conclusions(row, wave_df, hdf):
  """Conclusiones clínicas por dominios desde la misma fuente canónica."""
  sep_df, sep_metrics = estimate_wave_separation(wave_df, row)
  sep_interp = interpret_wave_separation(sep_metrics, row)
  phenotype, _, _ = classify_central_pressure_phenotype(row, sep_metrics, hdf)

  pressure_txt, pressure_short, pressure_level, pulse_txt, pulse_short, pulse_level = _didactic_grade_pressure(row)
  aug_txt, aug_short, aug_level = _didactic_grade_augmentation(row)
  rvse_txt, rvse_short, rvse_level = _didactic_grade_rvse(sep_metrics)
  wave_txt, wave_short, wave_level = _didactic_grade_wave(sep_metrics, row)
  harm_txt, harm_short, harm_level = _didactic_grade_harmonics(hdf)
  phenotype_txt, phenotype_short = _didactic_phenotype_text(phenotype)

  # Sección 1: solo HTA central. No repite carga pulsátil.
  c1 = f"{pressure_txt} Conclusión diagnóstica: {pressure_short.capitalize()}"

  # Sección 2: PP central + PPA, independientes de HTA central.
  c2 = (
    f"{pulse_txt} "
    "Este dominio integra la carga pulsátil central y la amplificación periférico-central sin inferirlas desde la presencia de hipertensión central. "
    f"Conclusión breve: {pulse_short.capitalize()}"
  )

  c3 = (
    "El análisis de aumentación central evalúa la contribución reflejada al componente sistólico central. "
    f"{aug_txt} Conclusión breve: {aug_short.capitalize()}"
  )
  c4 = f"{wave_txt} Conclusión breve: {wave_short.capitalize()}"
  c5 = f"{rvse_txt} Conclusión breve: {rvse_short.capitalize()}"
  c6 = f"{harm_txt} Conclusión breve: {harm_short.capitalize()}"

  automatic_integrated_text, _ = build_automatic_integrated_conclusion(row, sep_metrics, hdf)
  c7 = (
    f"{phenotype_txt} "
    "Esta conclusión utiliza exactamente los mismos estados diagnósticos que las secciones previas. "
    f"Conclusión integrada: {phenotype_short.capitalize()} "
    f"{automatic_integrated_text}"
  )

  return [
    ("1. Presión aórtica central", c1),
    ("2. Carga pulsátil central", c2),
    ("3. Aumentación central", c3),
    ("4. Separación de ondas", c4),
    ("5. Reserva subendocárdica", c5),
    ("6. Análisis armónico", c6),
    ("7. Fenotipo vascular central integrado", c7),
  ], sep_df, sep_metrics, sep_interp

def classify_central_pressure_phenotype(row, sep_metrics, hdf):
  """Fenotipo integrado sin puntajes: usa exclusivamente el estado canónico único."""
  state = build_canonical_diagnostic_state(row, sep_metrics, hdf)
  d = state["domains"]
  phenotype, mechanism = _phenotype_from_canonical_state(state)

  def _fmt(v, dec=1, unit=""):
    try:
      f = float(v)
      if np.isnan(f): return "no disponible"
      return f"{f:.{dec}f}{unit}"
    except Exception:
      return "no disponible"

  def _row(key):
    x = d[key]
    value = x.get("value", np.nan)
    if key in ("hta_central", "carga_pulsatil"):
      valtxt = _fmt(value, 0, " mmHg")
    elif key == "amplificacion":
      valtxt = _fmt(value, 2)
    elif key == "aumentacion":
      valtxt = _fmt(value, 1, "%")
    elif key == "reflexion":
      valtxt = _fmt(value, 2)
    elif key == "rvse":
      valtxt = _fmt(value, 1, "%")
    else:
      valtxt = _fmt(value, 1)
    return [x["label"], x["status"], valtxt, x.get("criterion", ""), x.get("short", "")]

  harm_domain = d["armonicos"]
  harm_metrics = harm_domain.get("metrics", {}) if harm_domain.get("status") == "cuantificado" else {}
  if harm_metrics:
    harm_value = (
      f"HD {_fmt(harm_metrics.get('hd_percent'),1,'%')}; "
      f"H1 {_fmt(harm_metrics.get('h1_energy_percent'),1,'%')}; "
      f"H2 {_fmt(harm_metrics.get('h2_energy_percent'),1,'%')}; "
      f"H4+ {_fmt(harm_metrics.get('h4plus_energy_percent'),1,'%')}; "
      f"fdom {_fmt(harm_metrics.get('dominant_frequency_hz'),2,' Hz')}"
    )
  else:
    harm_value = "no disponible"

  table = [
    ["Dominio", "Estado", "Valor", "Criterio", "Conclusión"],
    _row("hta_central"),
    _row("carga_pulsatil"),
    _row("amplificacion"),
    _row("aumentacion"),
    _row("reflexion"),
    _row("rvse"),
    ["Análisis armónico", harm_domain["status"], harm_value, harm_domain.get("criterion", ""), harm_domain.get("short", "")],
  ]

  altered_text = "; ".join(d[k]["short"] for k in state["altered_keys"])
  if not altered_text:
    altered_text = "no se identifican dominios clínicos alterados entre los clasificables"

  text = (
    f"Fenotipo final: {phenotype}. Mecanismo predominante: {mechanism}. "
    f"Dominios clínicos alterados: {altered_text}. "
    "La HTA central y la carga pulsátil se clasifican por separado; PPA, IAu/AIx, RM y RVSE conservan estados independientes. "
    "RI es complementario de RM y no se cuenta como evidencia adicional; Tref es continuo. "
    "HD, H1, H2, H4+ y frecuencia dominante son métricas continuas de caracterización espectral y no generan fenotipos armónicos patológicos ni diagnóstico de rigidez. "
    f"{longitudinal_harmonic_mdc_note()}"
  )
  return phenotype, text, table

def build_pdf(row, wave_df, hdf, screenshot_png=None, firma_png=None, sello_png=None, logo_png=None):
  """Construye un PDF compacto, con conclusiones primero y luego grilla gráfica profesional."""
  dx, cat, ref, amp_sbp, ppa, risk = central_diagnosis(row)
  conclusion_blocks, sep_df, sep_metrics, sep_interp = build_continuous_conclusions(row, wave_df, hdf)
  final_phenotype, final_phenotype_text, final_phenotype_table = classify_central_pressure_phenotype(row, sep_metrics, hdf)

  buf = io.BytesIO()
  doc = SimpleDocTemplate(
    buf, pagesize=A4,
    rightMargin=11*mm, leftMargin=11*mm,
    topMargin=21*mm, bottomMargin=38*mm
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

  def _draw_image_contained(canvas, image_bytes, x, y, max_w, max_h):
    """Dibuja imagen dentro de caja máxima, centrada y sin distorsión."""
    if not image_bytes:
      return False
    try:
      w, h = fit_image_box(image_bytes, max_w, max_h)
      dx = x + (max_w - w) / 2.0
      dy = y + (max_h - h) / 2.0
      canvas.drawImage(ImageReader(io.BytesIO(image_bytes)), dx, dy, width=w, height=h, mask='auto')
      return True
    except Exception:
      return False

  def _draw_first_page_branding(canvas):
    """Logo institucional, firma y sello al pie de la primera hoja.

    Se dibujan en una banda reservada por el margen inferior del documento,
    por lo que no se superponen con tablas, conclusiones ni pie de página.
    Cada imagen conserva su relación de aspecto y se ajusta a su caja máxima.
    """
    if not any([logo_png, firma_png, sello_png]):
      return
    width, _ = A4
    left = 11*mm
    right = width - 11*mm
    band_y = 12.5*mm
    band_h = 23.5*mm
    gap = 3*mm

    canvas.saveState()
    canvas.setStrokeColor(colors.HexColor('#CFD8DC'))
    canvas.setLineWidth(0.25)
    canvas.line(left, band_y + band_h + 1.2*mm, right, band_y + band_h + 1.2*mm)

    col_w = (right - left - 2*gap) / 3.0
    logo_box = (left, band_y, col_w, band_h)
    firma_box = (left + col_w + gap, band_y, col_w, band_h)
    sello_box = (left + 2*(col_w + gap), band_y, col_w, band_h)

    _draw_image_contained(canvas, logo_png, *logo_box)
    _draw_image_contained(canvas, firma_png, *firma_box)
    _draw_image_contained(canvas, sello_png, *sello_box)

    canvas.setFillColor(colors.HexColor('#607D8B'))
    canvas.setFont('Helvetica', 5.9)
    if logo_png:
      canvas.drawCentredString(logo_box[0] + logo_box[2]/2.0, band_y - 1.7*mm, 'Logo institucional')
    if firma_png:
      canvas.drawCentredString(firma_box[0] + firma_box[2]/2.0, band_y - 1.7*mm, 'Firma digital')
    if sello_png:
      canvas.drawCentredString(sello_box[0] + sello_box[2]/2.0, band_y - 1.7*mm, 'Sello digital')
    canvas.restoreState()

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
    if doc_obj.page == 1:
      _draw_first_page_branding(canvas)
    canvas.restoreState()

  def _graph_cell(title, img, width=91*mm, height=50*mm):
    return [Paragraph(title, styles["H3PAC"]), Image(img, width=width, height=height)]

  story = []
  story.append(Paragraph("PRESIÓN AÓRTICA CENTRAL", styles["TitlePAC"]))
  story.append(Paragraph("Informe médico integrado con conclusiones clínicas resumidas y didácticas", styles["BodyPAC"]))
  story.append(Spacer(1, 1.7*mm))

  story.append(_section("1. Datos del paciente y valores principales"))
  datos = [
    ["Paciente", safe_text(row.get("paciente","")), "Estudio", safe_text(row.get("estudio",""))],
    ["Fecha", safe_text(row.get("fecha","")), "Hora", safe_text(row.get("hora",""))],
    ["Edad", _fmt(row.get("edad",""),0), "Sexo", safe_text(row.get("sexo",""))],
    ["Peso", _fmt(row.get("peso",""),1), "Altura", _fmt(row.get("altura",""),1)],
    ["IMC", _fmt(row.get("imc",""),1), "Medicación", safe_text(row.get("medicacion",""))],
  ]
  saha_pdf_ref = get_saha_central_sbp_reference(row)
  saha_p90 = _fmt(saha_pdf_ref.get("p90"), 1) if saha_pdf_ref.get("ok") else ""
  saha_zpct = (f"z {saha_pdf_ref.get('z', np.nan):.2f} / P{saha_pdf_ref.get('percentil', np.nan):.0f}" if saha_pdf_ref.get("ok") else "")
  saha_aix_pdf_ref = get_saha_aix75_reference(row)
  saha_aix_p90 = _fmt(saha_aix_pdf_ref.get("p90"), 1) if saha_aix_pdf_ref.get("ok") else ""
  saha_aix_pct = (f"P{saha_aix_pdf_ref.get('percentil', np.nan):.0f} / P75 {saha_aix_pdf_ref.get('p75', np.nan):.1f}" if saha_aix_pdf_ref.get("ok") else "")
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
      ["APC", "", _fmt(row.get("apc")), "mmHg"],
      ["SAHA P90 edad/sexo", "", saha_p90, "mmHg"],
      ["SAHA z / percentil", "", saha_zpct, ""],
      ["LEAD IAu P90 edad/sexo", "", saha_aix_p90, "%"],
      ["LEAD IAu percentil", "", saha_aix_pct, ""]]
  # Tablas apiladas en ancho completo para evitar superposición de columnas.
  # La versión anterior colocaba patient_table y values_table lado a lado:
  # 140 mm + 83 mm dentro de un contenedor de 188 mm, produciendo solapamiento.
  def _wrap_table_cells(table_rows, style_name="SmallPAC"):
    wrapped = []
    for row_cells in table_rows:
      wrapped.append([Paragraph(pdf_text(c), styles[style_name]) for c in row_cells])
    return wrapped

  patient_table = Table(
    _wrap_table_cells(datos),
    colWidths=[25*mm, 70*mm, 25*mm, 68*mm],
    style=_table_style("#EAF2F8", 7.0)
  )
  values_table = Table(
    _wrap_table_cells(vals),
    colWidths=[36*mm, 50*mm, 50*mm, 52*mm],
    style=_table_style("#D9EAF7", 7.0)
  )
  story.append(patient_table)
  story.append(Spacer(1, 1.1*mm))
  story.append(values_table)
  story.append(Spacer(1, 1.8*mm))

  story.append(_section("2. Conclusiones clínicas resumidas y didácticas"))

  # Los puntos 1-6 permanecen en el primer bloque de conclusiones.
  # El punto 7 se fuerza a comenzar en la página siguiente para evitar que su
  # subtítulo quede huérfano al pie de la primera página. Título y cuerpo se
  # mantienen juntos mediante KeepTogether.
  conclusion_rows_page1 = []
  for title, body in conclusion_blocks[:6]:
    conclusion_rows_page1.append([Paragraph(pdf_text(title), styles["MiniTitlePAC"])])
    conclusion_rows_page1.append([Paragraph(_pdf_bold_conclusions(body), styles["ConclusionPAC"])])

  conclusion_table_style = TableStyle([
    ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#F7FAFC")),
    ("BOX", (0,0), (-1,-1), 0.4, colors.HexColor("#90A4AE")),
    ("INNERGRID", (0,0), (-1,-1), 0.12, colors.HexColor("#ECEFF1")),
    ("VALIGN", (0,0), (-1,-1), "TOP"),
    ("LEFTPADDING", (0,0), (-1,-1), 5),
    ("RIGHTPADDING", (0,0), (-1,-1), 5),
    ("TOPPADDING", (0,0), (-1,-1), 2.4),
    ("BOTTOMPADDING", (0,0), (-1,-1), 2.4),
  ])

  if conclusion_rows_page1:
    story.append(Table(
      conclusion_rows_page1,
      colWidths=[188*mm],
      style=conclusion_table_style
    ))

  # Punto 7: comienza en página 2 y conserva subtítulo + conclusión unidos.
  if len(conclusion_blocks) >= 7:
    story.append(PageBreak())
    title7, body7 = conclusion_blocks[6]
    point7_table = Table([
      [Paragraph(pdf_text(title7), styles["MiniTitlePAC"])],
      [Paragraph(_pdf_bold_conclusions(body7), styles["ConclusionPAC"])],
    ], colWidths=[188*mm], style=TableStyle([
      ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#F7FAFC")),
      ("BOX", (0,0), (-1,-1), 0.4, colors.HexColor("#90A4AE")),
      ("INNERGRID", (0,0), (-1,-1), 0.12, colors.HexColor("#ECEFF1")),
      ("VALIGN", (0,0), (-1,-1), "TOP"),
      ("LEFTPADDING", (0,0), (-1,-1), 5),
      ("RIGHTPADDING", (0,0), (-1,-1), 5),
      ("TOPPADDING", (0,0), (-1,-1), 2.4),
      ("BOTTOMPADDING", (0,0), (-1,-1), 2.4),
    ]))
    story.append(KeepTogether([point7_table]))

  # Cualquier bloque futuro adicional (8+) se conserva después del punto 7.
  if len(conclusion_blocks) > 7:
    extra_rows = []
    for title, body in conclusion_blocks[7:]:
      extra_rows.append([Paragraph(pdf_text(title), styles["MiniTitlePAC"])])
      extra_rows.append([Paragraph(_pdf_bold_conclusions(body), styles["ConclusionPAC"])])
    story.append(Spacer(1, 1.2*mm))
    story.append(Table(extra_rows, colWidths=[188*mm], style=TableStyle([
      ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#F7FAFC")),
      ("BOX", (0,0), (-1,-1), 0.4, colors.HexColor("#90A4AE")),
      ("INNERGRID", (0,0), (-1,-1), 0.12, colors.HexColor("#ECEFF1")),
      ("VALIGN", (0,0), (-1,-1), "TOP"),
      ("LEFTPADDING", (0,0), (-1,-1), 5),
      ("RIGHTPADDING", (0,0), (-1,-1), 5),
      ("TOPPADDING", (0,0), (-1,-1), 2.4),
      ("BOTTOMPADDING", (0,0), (-1,-1), 2.4),
    ])))

  story.append(Spacer(1, 2.5*mm))
  story.append(_section("3. Gráficos del informe"))
  story.append(Spacer(1, 1.5*mm))
  story.append(KeepTogether([
    Paragraph("Presión aórtica central con ondas Pf/Pb superpuestas", styles["H3PAC"]),
    Image(plot_wave_separation(sep_df, sep_metrics), width=188*mm, height=88*mm)
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
  story.append(_section("4. Análisis armónico cuantificado y fenotipo final"))

  hm_pdf = harmonic_distortion_metrics(hdf)
  if hm_pdf.get("ok"):
    harm_summary = [
      ["HD energético", "H1", "H2", "H4+", "Frecuencia dominante"],
      [
        f"{hm_pdf.get('hd_percent', np.nan):.1f}%",
        f"{hm_pdf.get('h1_energy_percent', np.nan):.1f}%",
        f"{hm_pdf.get('h2_energy_percent', np.nan):.1f}%",
        f"{hm_pdf.get('h4plus_energy_percent', np.nan):.1f}%",
        f"{hm_pdf.get('dominant_frequency_hz', np.nan):.2f} Hz",
      ],
    ]
    story.append(Table(
      [[Paragraph(pdf_text(c), styles["SmallPAC"]) for c in rr] for rr in harm_summary],
      colWidths=[38*mm, 30*mm, 30*mm, 30*mm, 60*mm],
      style=_table_style("#EAF2F8", 6.8)
    ))
    story.append(Spacer(1, 1.2*mm))

  harm_table = [["Armónico", "Frecuencia (Hz)", "Amplitud", "Energía relativa (%)"]]
  for _, r in hdf.iterrows():
    harm_table.append([
      str(int(r.get("armónico", 0))),
      f"{r.get('frecuencia_hz',0):.2f}",
      f"{r.get('amplitud',0):.3f}",
      f"{r.get('energia_relativa_%',0):.1f}",
    ])
  story.append(Table(
    [[Paragraph(pdf_text(c), styles["SmallPAC"]) for c in rr] for rr in harm_table],
    colWidths=[28*mm, 48*mm, 48*mm, 64*mm],
    style=_table_style("#EAF2F8", 6.6)
  ))
  story.append(Spacer(1, 1.5*mm))

  # La tabla fenotípica tiene 5 columnas: se renderiza con 5 anchos explícitos.
  # Esto corrige el desajuste previo de 5 columnas de datos contra solo 3 colWidths.
  phenotype_rows = [[Paragraph(pdf_text(cell), styles["SmallPAC"]) for cell in row_cells] for row_cells in final_phenotype_table]
  story.append(Table(
    phenotype_rows,
    colWidths=[34*mm, 20*mm, 42*mm, 48*mm, 44*mm],
    style=_table_style("#D9EAF7", 6.2)
  ))
  story.append(Spacer(1, 1.5*mm))
  story.append(Paragraph(
    "Nota metodológica: la separación Pf/Pb es una estimación clínica no invasiva. RM se clasifica por percentiles de edad y método (P90 como límite diagnóstico; P95 como aumento marcado); RI es complemento proporcional sin doble conteo. Tref y Tfor/Tref se informan como variables continuas sin corte universal. El análisis armónico cuantifica HD energético y la distribución H1-Hn como métricas continuas de caracterización espectral; HD y H4+ se calculan sobre el espectro positivo completo disponible, aunque la tabla visual muestre solo los primeros armónicos; no se aplican H1 <35%, H4+ >=25%, Eh 1158 N/m ni Zc 675/1103 como umbrales clínicos. Estas métricas no se utilizan aisladamente para diagnosticar rigidez arterial ni definir el fenotipo vascular integrado. La aplicabilidad es aproximada cuando el algoritmo o método difiere de la referencia publicada.",
    styles["SmallPAC"]
  ))

  # Capturas fijas de la animación para el informe médico integrado.
  try:
    keyframes_png = plot_aortic_animation_keyframes(row, sep_df, sep_metrics, hdf)
    story.append(PageBreak())
    story.append(_section("5. Capturas didácticas de métricas alteradas"))
    story.append(Spacer(1, 2*mm))
    story.append(Image(keyframes_png, width=188*mm, height=126*mm))
    story.append(Spacer(1, 1.4*mm))
    story.append(Paragraph(
      "Las capturas corresponden a las métricas alteradas detectadas por el mismo motor de pausa didáctica de la animación. Cada cuadro muestra valor real del paciente, criterio/umbral aplicado y explicación del mecanismo fisiológico. Si hay menos de cuatro alteraciones, los cuadros restantes se identifican explícitamente como ausencia de otra alerta automática.",
      styles["SmallPAC"]
    ))
  except Exception as e:
    story.append(Spacer(1, 1.2*mm))
    story.append(Paragraph(pdf_text(f"No se pudieron generar las capturas didácticas de la animación: {e}"), styles["SmallPAC"]))

  if screenshot_png:
    story.append(PageBreak())
    story.append(_section("6. Captura pantalla de mediciones"))
    story.append(Spacer(1, 2*mm))
    story.append(Image(io.BytesIO(screenshot_png), width=182*mm, height=215*mm))

  story.append(Spacer(1, 3*mm))
  story.append(_section("7. Referencias bibliográficas"))
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
  ref_table = [[Paragraph(pdf_text(f"{i}. {ref_txt}"), styles["SmallPAC"])] for i, ref_txt in enumerate(refs, 1)]
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




def _componente_seguro_nombre_archivo(valor, fallback="SD"):
  """Limpia un componente del nombre de archivo para Windows/macOS/Linux."""
  s = safe_text(valor)
  if not s:
    s = fallback
  s = re.sub(r'[<>:"/\\|?*\x00-\x1F]', '-', s)
  s = re.sub(r'\s+', ' ', s).strip(' .,-_')
  return s or fallback


def _fecha_estudio_para_archivo(valor):
  """Normaliza la fecha del estudio a DD-MM-AAAA cuando es reconocible."""
  s = safe_text(valor)
  if not s:
    return "SD"
  formatos = (
    "%d/%m/%Y", "%d-%m-%Y",
    "%d/%m/%y", "%d-%m-%y",
    "%Y/%m/%d", "%Y-%m-%d",
  )
  for formato in formatos:
    try:
      return datetime.strptime(s, formato).strftime("%d-%m-%Y")
    except Exception:
      pass
  return _componente_seguro_nombre_archivo(s, "SD")


def _apellido_nombre_para_archivo(valor):
  """Devuelve (apellido, nombre) para el nombre del PDF PAC."""
  s = re.sub(r'\s+', ' ', safe_text(valor)).strip(' ,')
  if not s:
    return "SD", "SD"

  if "," in s:
    apellido, nombres = s.split(",", 1)
    apellido = _componente_seguro_nombre_archivo(apellido, "SD")
    nombres = _componente_seguro_nombre_archivo(nombres, "SD")
    return apellido, nombres

  partes = s.split()
  if len(partes) == 1:
    return _componente_seguro_nombre_archivo(partes[0], "SD"), "SD"

  apellido = partes[-1]
  nombres = " ".join(partes[:-1])
  return (
    _componente_seguro_nombre_archivo(apellido, "SD"),
    _componente_seguro_nombre_archivo(nombres, "SD"),
  )


def nombre_archivo_informe_pac(row):
  """APELLIDO, NOMBRE, FECHA DEL ESTUDIO, PAC, OBRA SOCIAL.pdf"""
  apellido, nombres = _apellido_nombre_para_archivo(row.get("paciente", ""))
  fecha = _fecha_estudio_para_archivo(row.get("fecha", ""))
  obra_social = _componente_seguro_nombre_archivo(row.get("obra_social", ""), "SD")
  return f"{apellido}, {nombres}, {fecha}, PAC, {obra_social}.pdf"


def save_history(row):
  """Guarda el registro actual en historial Excel de forma robusta.

  Esta función se mantiene separada del generador de PDF para evitar
  errores NameError cuando el usuario presiona "Guardar en historial".
  """
  try:
    clean_row = {}
    for k, v in dict(row).items():
      if isinstance(v, (np.generic,)):
        v = v.item()
      if isinstance(v, float) and np.isnan(v):
        v = ""
      clean_row[k] = v

    new = pd.DataFrame([clean_row])
    if HISTORIAL_FILE.exists():
      try:
        old = pd.read_excel(HISTORIAL_FILE)
        out = pd.concat([old, new], ignore_index=True)
      except Exception:
        out = new
    else:
      out = new

    out.to_excel(HISTORIAL_FILE, index=False)
    return out
  except Exception as e:
    st.error(f"No se pudo guardar el historial Excel: {e}")
    return pd.DataFrame([row])


# -----------------------------------------------------------------------------
# PROCESAMIENTO MASIVO DE PDF PAC
# -----------------------------------------------------------------------------
BATCH_MAX_FILES = 100


def _batch_excel_safe_value(value):
  """Normaliza valores para Excel/CSV sin perder datos clínicos."""
  if isinstance(value, np.generic):
    value = value.item()
  if value is None:
    return ""
  if isinstance(value, float) and np.isnan(value):
    return ""
  if isinstance(value, (dict, list, tuple, set)):
    try:
      value = json.dumps(value, ensure_ascii=False, default=str)
    except Exception:
      value = str(value)
  if isinstance(value, str):
    # Excel no admite determinados caracteres de control XML.
    value = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", " ", value)
  return value


def _batch_clean_dataframe(df):
  out = df.copy()
  for col in out.columns:
    out[col] = out[col].map(_batch_excel_safe_value)
  return out


def _batch_summary_excel_bytes(summary_df):
  """Genera un libro Excel con resumen general, exitosos y errores."""
  summary_df = _batch_clean_dataframe(summary_df)
  ok_df = summary_df[summary_df.get("estado", pd.Series(dtype=str)) == "GENERADO"].copy()
  err_df = summary_df[summary_df.get("estado", pd.Series(dtype=str)) != "GENERADO"].copy()
  buf = io.BytesIO()
  with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    summary_df.to_excel(writer, sheet_name="Resumen_lote", index=False)
    ok_df.to_excel(writer, sheet_name="Informes_generados", index=False)
    err_df.to_excel(writer, sheet_name="Errores", index=False)
  return buf.getvalue()


def _batch_unique_filename(filename, used_names):
  """Evita sobrescribir informes de pacientes con nombres coincidentes."""
  filename = safe_text(filename) or "INFORME_PAC.pdf"
  key = filename.casefold()
  if key not in used_names:
    used_names[key] = 1
    return filename
  used_names[key] += 1
  p = Path(filename)
  return f"{p.stem} ({used_names[key]}){p.suffix or '.pdf'}"


def _batch_files_signature(files):
  """Firma estable para advertir si cambió la selección después de procesar."""
  h = hashlib.sha256()
  for item in list(files or []):
    name = safe_text(getattr(item, "name", "archivo.pdf"))
    data = item.getvalue() if hasattr(item, "getvalue") else b""
    h.update(name.encode("utf-8", errors="ignore"))
    h.update(str(len(data or b"")).encode("ascii"))
    h.update(bytes(data or b""))
  return h.hexdigest()


def append_batch_history(rows):
  """Agrega estudios exitosos al historial sin borrar registros existentes.

  La escritura se realiza una sola vez al finalizar el lote. Se eliminan únicamente
  duplicados exactos por identidad del estudio, conservando la versión más reciente.
  """
  if not rows:
    if HISTORIAL_FILE.exists():
      try:
        return pd.read_excel(HISTORIAL_FILE)
      except Exception:
        pass
    return pd.DataFrame()

  clean_rows = []
  for row in rows:
    clean_rows.append({k: _batch_excel_safe_value(v) for k, v in dict(row).items()})
  new_df = pd.DataFrame(clean_rows)

  if HISTORIAL_FILE.exists():
    try:
      old_df = pd.read_excel(HISTORIAL_FILE)
      out = pd.concat([old_df, new_df], ignore_index=True, sort=False)
    except Exception:
      out = new_df
  else:
    out = new_df

  dedupe_candidates = [
    "paciente", "estudio", "fecha", "hora", "pas_central", "pad_central"
  ]
  dedupe_cols = [c for c in dedupe_candidates if c in out.columns]
  if dedupe_cols:
    out = out.drop_duplicates(subset=dedupe_cols, keep="last", ignore_index=True)
  out = _batch_clean_dataframe(out)
  out.to_excel(HISTORIAL_FILE, index=False)
  return out


def process_pac_pdf_batch(
  uploaded_files,
  metodo_calibracion="SD_PAOC",
  metodo_rmri="SCOR_RADIAL",
  logo_png=None,
  firma_png=None,
  sello_png=None,
  include_originals=False,
  save_to_history=False,
  progress_callback=None,
):
  """Procesa múltiples PDF PAC y devuelve informes individuales dentro de un ZIP.

  Cada archivo usa exactamente el mismo flujo clínico del modo individual:
  extracción del PDF, digitalización de la curva real, separación de ondas,
  armónicos, fenotipo integrado y construcción del PDF médico.
  Un error individual queda registrado y no interrumpe el resto del lote.
  """
  files = list(uploaded_files or [])
  if not files:
    raise ValueError("No se seleccionaron archivos PDF para procesar.")
  if len(files) > BATCH_MAX_FILES:
    raise ValueError(
      f"El lote contiene {len(files)} PDF. El máximo por ejecución es {BATCH_MAX_FILES} "
      "para proteger la memoria de Streamlit Cloud. Divida la carga en varios lotes."
    )

  summary_rows = []
  generated_reports = []
  successful_history_rows = []
  used_names = {}
  total = len(files)

  for index, uploaded in enumerate(files, start=1):
    source_name = safe_text(getattr(uploaded, "name", f"estudio_{index}.pdf"))
    if progress_callback:
      progress_callback(index - 1, total, source_name, "PROCESANDO")

    base_summary = {
      "orden": index,
      "archivo_origen": source_name,
      "estado": "ERROR",
      "paciente": "",
      "estudio": "",
      "fecha": "",
      "hora": "",
      "obra_social": "",
      "edad": "",
      "sexo": "",
      "pas_central_mmHg": "",
      "pad_central_mmHg": "",
      "pp_central_mmHg": "",
      "ppa": "",
      "diagnostico_hta_central": "",
      "rm_pb_pf": "",
      "ri": "",
      "rvse_calculado_pct": "",
      "fenotipo_final": "",
      "informe_generado": "",
      "pagina_curva": "",
      "color_curva": "",
      "puntos_curva": "",
      "conclusion_integrada": "",
      "error": "",
    }

    try:
      raw = uploaded.getvalue() if hasattr(uploaded, "getvalue") else uploaded.read()
      pdf_in = bytes(raw or b"")
      if not pdf_in:
        raise ValueError("El PDF está vacío.")
      if not pdf_in.lstrip().startswith(b"%PDF"):
        raise ValueError("El archivo no tiene una cabecera PDF válida.")

      text = extract_pdf_text(pdf_in)
      row = parse_model_pac_from_pdf(pdf_in, text)
      if not isinstance(row, dict):
        raise ValueError("No se pudieron estructurar los datos del estudio.")
      row = dict(row)
      row["metodo_calibracion_pac"] = metodo_calibracion
      row["metodo_referencia_rmri"] = metodo_rmri

      screenshot_png = render_pdf_page_png(pdf_in, page_index=1)
      wave_df, curve_debug_png, curve_meta = digitize_curve_from_pdf(
        pdf_in, row, max_pages=4, zoom=3.0
      )
      hdf = harmonic_analysis(wave_df)
      sep_df, sep_metrics = estimate_wave_separation(wave_df, row)
      final_phenotype, _, _ = classify_central_pressure_phenotype(row, sep_metrics, hdf)
      integrated_text, _ = build_automatic_integrated_conclusion(row, sep_metrics, hdf)
      hta_status = central_hypertension_status(row)
      _, _, _, _, ppa, _ = central_diagnosis(row)

      report_bytes = ensure_download_bytes(
        build_pdf_one_page(
          row,
          wave_df,
          hdf,
          screenshot_png,
          firma_png=firma_png,
          sello_png=sello_png,
          logo_png=logo_png,
        )
      )
      if not report_bytes or not report_bytes.lstrip().startswith(b"%PDF"):
        raise ValueError("El generador no devolvió un PDF médico válido.")

      report_name = _batch_unique_filename(nombre_archivo_informe_pac(row), used_names)
      generated_reports.append({
        "filename": report_name,
        "data": report_bytes,
        "patient": safe_text(row.get("paciente")) or "Sin identificar",
        "source_name": source_name,
        "original": pdf_in if include_originals else None,
      })
      successful_history_rows.append(row)

      base_summary.update({
        "estado": "GENERADO",
        "paciente": safe_text(row.get("paciente")),
        "estudio": safe_text(row.get("estudio")),
        "fecha": safe_text(row.get("fecha")),
        "hora": safe_text(row.get("hora")),
        "obra_social": safe_text(row.get("obra_social")),
        "edad": to_float(row.get("edad")),
        "sexo": safe_text(row.get("sexo")),
        "pas_central_mmHg": to_float(row.get("pas_central")),
        "pad_central_mmHg": to_float(row.get("pad_central")),
        "pp_central_mmHg": to_float(row.get("pp_central")),
        "ppa": ppa,
        "diagnostico_hta_central": hta_status.get("diagnostico_breve", ""),
        "rm_pb_pf": to_float(sep_metrics.get("rm")),
        "ri": to_float(sep_metrics.get("ri")),
        "rvse_calculado_pct": to_float(sep_metrics.get("rvse_calculado_%")),
        "fenotipo_final": final_phenotype,
        "informe_generado": report_name,
        "pagina_curva": curve_meta.get("pagina", ""),
        "color_curva": curve_meta.get("color_detectado", ""),
        "puntos_curva": curve_meta.get("puntos", len(wave_df)),
        "conclusion_integrada": integrated_text,
        "error": "",
      })
      if progress_callback:
        progress_callback(index, total, source_name, "GENERADO")

    except Exception as exc:
      base_summary["error"] = safe_text(exc)[:1500]
      if progress_callback:
        progress_callback(index, total, source_name, "ERROR")
    finally:
      summary_rows.append(base_summary)
      # Evita acumulación de figuras y matrices de imagen entre pacientes.
      try:
        plt.close("all")
      except Exception:
        pass
      gc.collect()

  summary_df = pd.DataFrame(summary_rows)
  excel_bytes = _batch_summary_excel_bytes(summary_df)

  history_total = None
  if save_to_history and successful_history_rows:
    history_total = len(append_batch_history(successful_history_rows))

  zip_buffer = io.BytesIO()
  used_original_names = {}
  with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
    for original_index, item in enumerate(generated_reports, start=1):
      zf.writestr(f"informes_pdf/{item['filename']}", item["data"])
      if include_originals and item.get("original"):
        original_name = _componente_seguro_nombre_archivo(
          item["source_name"], f"original_{original_index}.pdf"
        )
        if not original_name.lower().endswith(".pdf"):
          original_name += ".pdf"
        original_name = _batch_unique_filename(original_name, used_original_names)
        zf.writestr(f"pdf_originales/{original_name}", item["original"])
    zf.writestr("resumen_lote_pac.xlsx", excel_bytes)
    error_df = summary_df[summary_df["estado"] != "GENERADO"]
    if not error_df.empty:
      zf.writestr(
        "errores_lote.csv",
        _batch_clean_dataframe(error_df).to_csv(index=False).encode("utf-8-sig"),
      )
    manifest = (
      "PROCESAMIENTO MASIVO PAC\n"
      f"Fecha de generación: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
      f"PDF recibidos: {total}\n"
      f"Informes generados: {len(generated_reports)}\n"
      f"Errores: {total - len(generated_reports)}\n"
      f"Método calibración PAC: {metodo_calibracion}\n"
      f"Método referencia RM/RI: {metodo_rmri}\n"
      "Cada informe fue generado desde la curva real digitalizada de su PDF.\n"
    )
    zf.writestr("LEEME.txt", manifest.encode("utf-8"))

  return {
    "zip_bytes": zip_buffer.getvalue(),
    "excel_bytes": excel_bytes,
    "summary_df": summary_df,
    "reports": generated_reports,
    "total": total,
    "success_count": len(generated_reports),
    "error_count": total - len(generated_reports),
    "history_total": history_total,
  }




# -----------------------------------------------------------------------------
# PDF DE 1 HOJA: INFORME MEDICION DE PRESION CENTRAL
# -----------------------------------------------------------------------------
def _pc_short_text(v, dec=1, unit=""):
  try:
    f = float(v)
    if np.isnan(f):
      return "no disponible"
    if dec == 0:
      return f"{f:.0f}{unit}"
    return f"{f:.{dec}f}{unit}"
  except Exception:
    return "no disponible"


def _pc_extract_bold_only(texto):
  txt = safe_text(texto)
  markers = ["Conclusión diagnóstica:", "Conclusión breve:", "Conclusión integrada:", "Conclusión final:"]
  pos = -1
  marker_used = None
  for m in markers:
    i = txt.lower().find(m.lower())
    if i != -1 and i > pos:
      pos = i
      marker_used = m
  if pos == -1 or marker_used is None:
    return txt.strip()
  out = txt[pos + len(marker_used):].strip()
  out = re.split(r"(?<=[\.!?])\s+", out)[0].strip()
  out = out.strip(" -:;")
  return out or txt.strip()


def _pc_make_conclusion_items(row, wave_df, hdf):
  blocks, sep_df, sep_metrics, sep_interp = build_continuous_conclusions(row, wave_df, hdf)
  items = []
  for title, body in blocks:
    num_title = safe_text(title)
    if not any(num_title.startswith(f"{n}.") for n in ["1", "2", "3", "4", "5"]):
      continue
    items.append((num_title, _pc_extract_bold_only(body)))
  return items, sep_df, sep_metrics


def _pc_build_didactic_sheet(row, wave_df, sep_df, sep_metrics, hdf, screenshot_png=None):
  """Compone una lámina didáctica compacta 2x2 en PNG para el PDF de 1 hoja."""
  try:
    sources = []
    try:
      sources.append(("Presiones periféricas vs centrales", plot_pressure_comparison(row)))
    except Exception:
      pass
    try:
      sources.append(("Pulso aórtico central y separación de ondas", plot_wave_separation(sep_df, sep_metrics)))
    except Exception:
      pass
    try:
      sources.append(("Reserva subendocárdica", plot_rvse_area(sep_df, sep_metrics)))
    except Exception:
      pass
    try:
      sources.append(("Análisis armónico", plot_harmonics(hdf)))
    except Exception:
      pass
    if not sources:
      try:
        return plot_waveform(wave_df)
      except Exception:
        return screenshot_png
    if PILImage is None or ImageDraw is None:
      return sources[0][1]

    tile_w, tile_h = 860, 390
    canvas = PILImage.new("RGB", (tile_w * 2, tile_h * 2), "white")
    draw = ImageDraw.Draw(canvas)
    positions = [(0, 0), (tile_w, 0), (0, tile_h), (tile_w, tile_h)]
    for i, (title, img_bytes) in enumerate(sources[:4]):
      try:
        img_bytes = ensure_download_bytes(img_bytes)
        if not img_bytes:
          continue
        img = PILImage.open(io.BytesIO(img_bytes)).convert("RGB")
        img.thumbnail((tile_w - 36, tile_h - 52))
        x0, y0 = positions[i]
        draw.rectangle([x0 + 6, y0 + 6, x0 + tile_w - 6, y0 + tile_h - 6], outline=(210, 218, 226), width=2)
        draw.text((x0 + 18, y0 + 14), title, fill=(31, 45, 61))
        px = x0 + (tile_w - img.width) // 2
        py = y0 + 40 + (tile_h - 46 - img.height) // 2
        canvas.paste(img, (px, py))
      except Exception:
        continue
    buf = io.BytesIO()
    canvas.save(buf, format="PNG")
    return buf.getvalue()
  except Exception:
    try:
      return plot_waveform(wave_df)
    except Exception:
      return screenshot_png


def _pc_paragraph(text, style):
  return Paragraph(pdf_text(text), style)


def _pc_maybe_image(image_bytes, max_w, max_h):
  image_bytes = ensure_download_bytes(image_bytes)
  if not image_bytes:
    return None
  try:
    w, h = fit_image_box(image_bytes, max_w, max_h)
    return Image(io.BytesIO(image_bytes), width=w, height=h)
  except Exception:
    return None


def _pc_diag_palette(text):
  """Color institucional de apoyo según el sentido de la conclusión, sin alterar el diagnóstico."""
  t = safe_text(text).lower()
  alert_terms = ["con hipertensión", "aumentada", "aumentado", "alterada", "alterado", "reducida", "reducido", "patológica", "patologico", "marcadamente"]
  ok_terms = ["sin hipertensión", "sin aumentación", "no aumentada", "no aumentado", "normal", "conservada", "conservado", "esperada", "esperado", "adecuada", "adecuado"]
  mid_terms = ["intermedia", "intermedio", "relativamente", "limítrofe", "limitrofe", "subóptima", "suboptima"]
  # Las expresiones negativas (p. ej. "sin aumentación aumentada" / "no aumentada")
  # tienen prioridad para no colorearlas erróneamente como patológicas por contener la palabra "aumentada".
  if any(k in t for k in ok_terms):
    return colors.HexColor("#18794E"), colors.HexColor("#EAF7F0")
  if any(k in t for k in mid_terms):
    return colors.HexColor("#A15C00"), colors.HexColor("#FFF4E5")
  if any(k in t for k in alert_terms):
    return colors.HexColor("#B42318"), colors.HexColor("#FDECEC")
  return colors.HexColor("#315B7D"), colors.HexColor("#EEF4F8")


def _pc_kpi_cell(label, value, styles, accent="#1F4E79"):
  return Table(
    [[Paragraph(f"<font color='{accent}'><b>{pdf_text(label)}</b></font>", styles["PC_KPI_Label"])],
     [Paragraph(f"<b>{pdf_text(value)}</b>", styles["PC_KPI_Value"])]],
    colWidths=[56*mm],
    rowHeights=[5.2*mm, 7.2*mm],
    style=TableStyle([
      ("BACKGROUND", (0,0), (-1,-1), colors.white),
      ("BOX", (0,0), (-1,-1), 0.6, colors.HexColor("#D7E1EA")),
      ("LINEBEFORE", (0,0), (0,-1), 3.2, colors.HexColor(accent)),
      ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
      ("LEFTPADDING", (0,0), (-1,-1), 5),
      ("RIGHTPADDING", (0,0), (-1,-1), 4),
      ("TOPPADDING", (0,0), (-1,-1), 1),
      ("BOTTOMPADDING", (0,0), (-1,-1), 1),
    ])
  )


def build_pdf_one_page(row, wave_df, hdf, screenshot_png=None, firma_png=None, sello_png=None, logo_png=None):
  """Versión institucional optimizada: informe PDF A4 de una sola hoja para Presión Central."""
  conclusion_items, sep_df, sep_metrics = _pc_make_conclusion_items(row, wave_df, hdf)
  didactic_png = _pc_build_didactic_sheet(row, wave_df, sep_df, sep_metrics, hdf, screenshot_png=screenshot_png)
  dx, cat, ref, amp_sbp, ppa, risk = central_diagnosis(row)

  buf = io.BytesIO()
  doc = SimpleDocTemplate(
    buf, pagesize=A4,
    rightMargin=6*mm, leftMargin=6*mm,
    topMargin=6*mm, bottomMargin=6*mm,
    allowSplitting=0,
  )
  styles = getSampleStyleSheet()
  styles.add(ParagraphStyle(name="PC_Title", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=12.2, leading=13.0, textColor=colors.white, alignment=0, spaceAfter=0))
  styles.add(ParagraphStyle(name="PC_Sub", parent=styles["BodyText"], fontName="Helvetica", fontSize=6.3, leading=7.0, textColor=colors.HexColor("#D8E7F2"), alignment=0, spaceAfter=0))
  styles.add(ParagraphStyle(name="PC_Section", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=7.4, leading=8.0, textColor=colors.HexColor("#153B5B"), spaceBefore=0.5, spaceAfter=1.0))
  styles.add(ParagraphStyle(name="PC_Body", parent=styles["BodyText"], fontName="Helvetica", fontSize=6.0, leading=6.8, textColor=colors.HexColor("#263746"), spaceAfter=0.4))
  styles.add(ParagraphStyle(name="PC_Small", parent=styles["BodyText"], fontName="Helvetica", fontSize=5.85, leading=6.5, textColor=colors.HexColor("#31424F"), spaceAfter=0))
  styles.add(ParagraphStyle(name="PC_Bold", parent=styles["BodyText"], fontName="Helvetica-Bold", fontSize=6.05, leading=6.8, textColor=colors.HexColor("#12263A"), spaceAfter=0))
  styles.add(ParagraphStyle(name="PC_KPI_Label", parent=styles["BodyText"], fontName="Helvetica-Bold", fontSize=5.6, leading=6.0, textColor=colors.HexColor("#486577"), spaceAfter=0))
  styles.add(ParagraphStyle(name="PC_KPI_Value", parent=styles["BodyText"], fontName="Helvetica-Bold", fontSize=8.5, leading=9.0, textColor=colors.HexColor("#17324A"), spaceAfter=0))
  styles.add(ParagraphStyle(name="PC_ConcTitle", parent=styles["BodyText"], fontName="Helvetica-Bold", fontSize=5.85, leading=6.4, textColor=colors.HexColor("#203746"), spaceAfter=0))
  styles.add(ParagraphStyle(name="PC_Conc", parent=styles["BodyText"], fontName="Helvetica-Bold", fontSize=5.95, leading=6.55, textColor=colors.HexColor("#152536"), spaceAfter=0))

  story = []

  logo = _pc_maybe_image(logo_png, 22*mm, 12*mm)
  title_block = [
    Paragraph("INFORME MEDICION DE PRESION CENTRAL", styles["PC_Title"]),
    Paragraph("Evaluación no invasiva de presión aórtica central y mecánica de la onda de pulso", styles["PC_Sub"]),
  ]
  header = Table(
    [[logo if logo is not None else "", title_block]],
    colWidths=[26*mm, 165*mm],
    rowHeights=[15*mm],
    style=TableStyle([
      ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#153B5B")),
      ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
      ("LEFTPADDING", (0,0), (0,0), 4),
      ("RIGHTPADDING", (0,0), (0,0), 2),
      ("LEFTPADDING", (1,0), (1,0), 7),
      ("RIGHTPADDING", (1,0), (1,0), 4),
      ("TOPPADDING", (0,0), (-1,-1), 2),
      ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ])
  )
  story.append(header)
  story.append(Spacer(1, 1.3))

  story.append(Paragraph("DATOS DEL PACIENTE", styles["PC_Section"]))
  patient_data = [
    ["Paciente", safe_text(row.get("paciente")) or "SD", "Edad / Sexo", f"{safe_text(row.get('edad')) or 'SD'} / {safe_text(row.get('sexo')) or 'SD'}", "Fecha", safe_text(row.get("fecha")) or "SD"],
    ["Documento", safe_text(row.get("documento")) or "SD", "Obra social", safe_text(row.get("obra_social")) or "SD", "Estudio", safe_text(row.get("estudio")) or "Presión Central"],
  ]
  t_patient = Table(patient_data, colWidths=[18*mm, 48*mm, 22*mm, 40*mm, 18*mm, 45*mm], rowHeights=[6.2*mm, 6.2*mm])
  t_patient.setStyle(TableStyle([
    ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#F7FAFC")),
    ("GRID", (0,0), (-1,-1), 0.35, colors.HexColor("#D8E2EA")),
    ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
    ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
    ("FONTNAME", (2,0), (2,-1), "Helvetica-Bold"),
    ("FONTNAME", (4,0), (4,-1), "Helvetica-Bold"),
    ("FONTSIZE", (0,0), (-1,-1), 6.0),
    ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ("LEFTPADDING", (0,0), (-1,-1), 2.5),
    ("RIGHTPADDING", (0,0), (-1,-1), 2.5),
    ("TOPPADDING", (0,0), (-1,-1), 1),
    ("BOTTOMPADDING", (0,0), (-1,-1), 1),
  ]))
  story.append(t_patient)
  story.append(Spacer(1, 1.1))

  story.append(Paragraph("METODOLOGIA", styles["PC_Section"]))
  metodologia = (
    "Se realizo medicion de presion central y metricas derivadas con sensores piezo electricos "
    "con dispositivo Aortic (Exxer), previo reposo de 5 minutos con paciente en decubito supino, "
    "se registraron 3 mediciones promediadas."
  )
  meth = Table([[Paragraph(pdf_text(metodologia), styles["PC_Body"])]], colWidths=[191*mm])
  meth.setStyle(TableStyle([
    ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#EEF5FA")),
    ("BOX", (0,0), (-1,-1), 0.45, colors.HexColor("#C9DAE6")),
    ("LEFTPADDING", (0,0), (-1,-1), 4),
    ("RIGHTPADDING", (0,0), (-1,-1), 4),
    ("TOPPADDING", (0,0), (-1,-1), 2.5),
    ("BOTTOMPADDING", (0,0), (-1,-1), 2.5),
  ]))
  story.append(meth)
  story.append(Spacer(1, 1.1))

  story.append(Paragraph("RESULTADOS", styles["PC_Section"]))
  kpis = [
    _pc_kpi_cell("PAS CENTRAL", _pc_short_text(row.get("pas_central"), 0, " mmHg"), styles, "#1F4E79"),
    _pc_kpi_cell("PP CENTRAL", _pc_short_text(row.get("pp_central"), 0, " mmHg"), styles, "#315B7D"),
    _pc_kpi_cell("IAu / AIx", _pc_short_text(row.get("iau"), 1, "%"), styles, "#6F4E7C"),
    _pc_kpi_cell("PPA", _pc_short_text(ppa, 2, ""), styles, "#8A5A00"),
    _pc_kpi_cell("RM", _pc_short_text(sep_metrics.get("rm"), 2, ""), styles, "#3C6E71"),
    _pc_kpi_cell("RVSE", _pc_short_text(sep_metrics.get("rvse_calculado_%"), 1, "%"), styles, "#18794E"),
  ]
  kpi_table = Table([kpis[:3], kpis[3:]], colWidths=[63.3*mm]*3, rowHeights=[11.8*mm, 11.8*mm])
  kpi_table.setStyle(TableStyle([
    ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ("LEFTPADDING", (0,0), (-1,-1), 1.5),
    ("RIGHTPADDING", (0,0), (-1,-1), 1.5),
    ("TOPPADDING", (0,0), (-1,-1), 1),
    ("BOTTOMPADDING", (0,0), (-1,-1), 1),
  ]))
  story.append(kpi_table)

  dx_short = _pc_extract_bold_only(dx)
  status_color, status_bg = _pc_diag_palette(dx_short)
  status_box = Table(
    [[Paragraph("<b>Presión aórtica central</b>", styles["PC_Bold"]), Paragraph(f"<b>{pdf_text(dx_short)}</b>", styles["PC_Bold"])]],
    colWidths=[45*mm, 146*mm],
    rowHeights=[7.2*mm],
    style=TableStyle([
      ("BACKGROUND", (0,0), (-1,-1), status_bg),
      ("BOX", (0,0), (-1,-1), 0.5, status_color),
      ("LINEBEFORE", (0,0), (0,0), 3.0, status_color),
      ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
      ("LEFTPADDING", (0,0), (-1,-1), 4),
      ("RIGHTPADDING", (0,0), (-1,-1), 4),
      ("TOPPADDING", (0,0), (-1,-1), 1),
      ("BOTTOMPADDING", (0,0), (-1,-1), 1),
    ])
  )
  story.append(status_box)
  story.append(Spacer(1, 1.1))

  story.append(Paragraph("CONCLUSIONES", styles["PC_Section"]))
  conc_rows = []
  for title, conc in conclusion_items:
    conc_rows.append([
      Paragraph(f"<b>{pdf_text(title)}</b>", styles["PC_ConcTitle"]),
      Paragraph(f"<b>{pdf_text(conc)}</b>", styles["PC_Conc"]),
      ""
    ])
  if not conc_rows:
    conc_rows = [[Paragraph("<b>1. Presión aórtica central</b>", styles["PC_ConcTitle"]), Paragraph(f"<b>{pdf_text(dx_short)}</b>", styles["PC_Conc"]), ""]]
  t_conc = Table(conc_rows, colWidths=[52*mm, 136*mm, 3*mm])
  ts = [
    ("GRID", (0,0), (-2,-1), 0.35, colors.HexColor("#D9E2EC")),
    ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ("LEFTPADDING", (0,0), (-1,-1), 2.5),
    ("RIGHTPADDING", (0,0), (-1,-1), 2.5),
    ("TOPPADDING", (0,0), (-1,-1), 1.2),
    ("BOTTOMPADDING", (0,0), (-1,-1), 1.2),
  ]
  for i, (_, conc) in enumerate(conclusion_items[:len(conc_rows)]):
    fg, bg = _pc_diag_palette(conc)
    ts += [
      ("BACKGROUND", (0,i), (1,i), bg),
      ("BACKGROUND", (2,i), (2,i), fg),
    ]
  t_conc.setStyle(TableStyle(ts))
  story.append(t_conc)
  story.append(Spacer(1, 1.1))

  story.append(Paragraph("LAMINAS GRAFICAS", styles["PC_Section"]))
  did_img = _pc_maybe_image(didactic_png, 191*mm, 67*mm)
  if did_img is not None:
    img_table = Table([[did_img]], colWidths=[191*mm])
    img_table.setStyle(TableStyle([
      ("BACKGROUND", (0,0), (-1,-1), colors.white),
      ("BOX", (0,0), (-1,-1), 0.5, colors.HexColor("#D4DEE7")),
      ("LEFTPADDING", (0,0), (-1,-1), 1.5),
      ("RIGHTPADDING", (0,0), (-1,-1), 1.5),
      ("TOPPADDING", (0,0), (-1,-1), 1.5),
      ("BOTTOMPADDING", (0,0), (-1,-1), 1.5),
      ("ALIGN", (0,0), (-1,-1), "CENTER"),
    ]))
    story.append(img_table)
  story.append(Spacer(1, 0.9))

  firma = _pc_maybe_image(firma_png, 29*mm, 9*mm) if firma_png else None
  sello = _pc_maybe_image(sello_png, 21*mm, 9*mm) if sello_png else None
  footer_text = "Informe resumido de una hoja. La interpretación debe integrarse con el contexto clínico del paciente."
  footer = Table(
    [[firma if firma is not None else "", sello if sello is not None else "", Paragraph(pdf_text(footer_text), styles["PC_Small"]) ]],
    colWidths=[33*mm, 24*mm, 134*mm],
    rowHeights=[9.5*mm],
    style=TableStyle([
      ("LINEABOVE", (0,0), (-1,0), 0.45, colors.HexColor("#BFCBD5")),
      ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
      ("LEFTPADDING", (0,0), (-1,-1), 0),
      ("RIGHTPADDING", (0,0), (-1,-1), 1.5),
      ("TOPPADDING", (0,0), (-1,-1), 0.8),
      ("BOTTOMPADDING", (0,0), (-1,-1), 0),
    ])
  )
  story.append(footer)

  doc.build(story)
  return buf.getvalue()


st.markdown("""
<style>
.block-container {padding-top: 1.2rem;}
.pac-institutional-header {background: linear-gradient(90deg,#153B5B,#1F4E79); padding:18px 22px; border-radius:12px; margin-bottom:10px; box-shadow:0 5px 18px rgba(21,59,91,.15);}
.pac-institutional-header h1 {color:white; margin:0; font-size:2rem; letter-spacing:.01em;}
.pac-institutional-header p {color:#D8E7F2; margin:5px 0 0 0; font-size:.95rem;}
</style>
<div class="pac-institutional-header">
<h1>INFORME MEDICION DE PRESION CENTRAL</h1>
<p>Evaluación no invasiva de presión aórtica central y mecánica de la onda de pulso</p>
</div>
""", unsafe_allow_html=True)
st.caption("Importación tipo MODELO PAC, digitalización real de curva del estudio original, informe PDF institucional de 1 hoja, historial Excel y análisis armónico.")

with st.sidebar:
  st.header("1) Informe individual")
  pdf_file = st.file_uploader(
    "PDF original PAC / Exxer",
    type=["pdf"],
    key="pac_pdf_individual",
  )
  wave_file = st.file_uploader(
    "Opcional: CSV/TXT curva central REAL del paciente (tiempo_ms, presion_mmHg)",
    type=["csv", "txt"],
    key="pac_curva_individual",
  )
  st.markdown("---")
  st.header("2) Informes por lotes")
  batch_pdf_files = st.file_uploader(
    "Importar múltiples PDF PAC",
    type=["pdf"],
    accept_multiple_files=True,
    key="pac_pdf_lote",
    help=f"Seleccione hasta {BATCH_MAX_FILES} PDF. Se generará un informe individual por paciente y un ZIP para descargar todo el lote.",
  )
  if batch_pdf_files:
    st.caption(f"PDF seleccionados: {len(batch_pdf_files)}")
  st.markdown("---")
  st.info("Modo datos reales: si no se carga CSV/TXT, la app digitaliza automáticamente la curva desde la imagen del PDF. No se aceptan curvas sintéticas ni genéricas.")
  st.caption("El logo, la firma y el sello cargados en la pantalla principal se aplican también a todos los informes del lote.")

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

# -----------------------------------------------------------------------------
# IDENTIDAD VISUAL DEL INFORME
# -----------------------------------------------------------------------------
st.markdown("## Identidad visual del informe")
st.info("Cargue aquí los archivos que se insertarán automáticamente al final de la primera hoja del PDF. La app reserva una banda inferior para evitar superposiciones, deformaciones o truncamientos.")
with st.expander("Cargar logo institucional, firma digital y sello digital", expanded=True):
  c_logo, c_firma, c_sello = st.columns(3)
  with c_logo:
    logo_file = st.file_uploader(
      "Logo institucional",
      type=["png", "jpg", "jpeg"],
      key="logo_institucional_main",
      help="Imagen opcional. Se ubicará al pie de la primera hoja, sector izquierdo."
    )
    if logo_file is not None:
      st.image(logo_file, caption="Logo cargado", use_container_width=True)
  with c_firma:
    firma_file = st.file_uploader(
      "Firma digital",
      type=["png", "jpg", "jpeg"],
      key="firma_digital_main",
      help="Imagen opcional. Se ubicará al pie de la primera hoja, sector central."
    )
    if firma_file is not None:
      st.image(firma_file, caption="Firma cargada", use_container_width=True)
  with c_sello:
    sello_file = st.file_uploader(
      "Sello digital",
      type=["png", "jpg", "jpeg"],
      key="sello_digital_main",
      help="Imagen opcional. Se ubicará al pie de la primera hoja, sector derecho."
    )
    if sello_file is not None:
      st.image(sello_file, caption="Sello cargado", use_container_width=True)
  st.caption("Formatos aceptados: PNG, JPG y JPEG. Cada imagen conserva su proporción dentro de su caja asignada.")

logo_png = read_uploaded_image_bytes(logo_file)
firma_png = read_uploaded_image_bytes(firma_file)
sello_png = read_uploaded_image_bytes(sello_file)

# -----------------------------------------------------------------------------
# INTERFAZ DE PROCESAMIENTO POR LOTES
# -----------------------------------------------------------------------------
st.markdown("## Generación de informes PAC por lotes")
st.caption(
  "Importe varios PDF del equipo. La app procesa cada paciente de forma independiente, "
  "genera un PDF médico por estudio y prepara un archivo ZIP para descargar el lote completo."
)

with st.expander(
  "Procesar múltiples PDF y descargar informes individuales",
  expanded=bool(batch_pdf_files),
):
  batch_count = len(batch_pdf_files or [])
  if batch_count:
    total_mb = sum(len(f.getvalue()) for f in batch_pdf_files) / (1024 * 1024)
    st.info(
      f"Lote preparado: {batch_count} PDF, {total_mb:.1f} MB. "
      f"Máximo permitido por ejecución: {BATCH_MAX_FILES} PDF."
    )
    with st.expander("Ver archivos seleccionados", expanded=False):
      selected_df = pd.DataFrame([
        {
          "archivo": safe_text(getattr(f, "name", "archivo.pdf")),
          "tamaño_MB": round(len(f.getvalue()) / (1024 * 1024), 2),
        }
        for f in batch_pdf_files
      ])
      st.dataframe(selected_df, use_container_width=True, hide_index=True)
  else:
    st.info("Seleccione los PDF en la barra lateral, en “2) Informes por lotes”.")

  bc1, bc2 = st.columns(2)
  with bc1:
    batch_calibration = st.selectbox(
      "Calibración PAC común para el lote",
      options=["SD_PAOC", "C_PAOC"],
      index=0,
      key="batch_metodo_calibracion",
      format_func=lambda x: SAHA_CALIBRATION_LABELS.get(x, x),
    )
  with bc2:
    batch_rmri = st.selectbox(
      "Referencia RM/RI común para el lote",
      options=["SCOR_RADIAL", "SCOR_CAROTID", "MOG"],
      index=0,
      key="batch_metodo_rmri",
      format_func=lambda x: RM_RI_METHOD_LABELS.get(x, x),
    )

  bo1, bo2 = st.columns(2)
  with bo1:
    batch_save_history = st.checkbox(
      "Agregar estudios exitosos al historial Excel",
      value=False,
      key="batch_save_history",
      help="No borra el historial existente. Agrega los estudios exitosos y evita duplicados exactos.",
    )
  with bo2:
    batch_include_originals = st.checkbox(
      "Incluir también los PDF originales dentro del ZIP",
      value=False,
      key="batch_include_originals",
    )

  action_col, clear_col = st.columns([3, 1])
  with action_col:
    process_batch_clicked = st.button(
      "Procesar lote y generar informes",
      type="primary",
      use_container_width=True,
      disabled=(batch_count == 0 or batch_count > BATCH_MAX_FILES),
      key="process_pac_batch",
    )
  with clear_col:
    clear_batch_clicked = st.button(
      "Limpiar resultados",
      use_container_width=True,
      key="clear_pac_batch_results",
    )

  if clear_batch_clicked:
    st.session_state.pop("pac_batch_result", None)
    st.success("Resultados anteriores eliminados de la sesión. Los PDF seleccionados permanecen disponibles.")

  if batch_count > BATCH_MAX_FILES:
    st.error(
      f"Se seleccionaron {batch_count} PDF. Divida la carga en lotes de hasta "
      f"{BATCH_MAX_FILES} archivos para evitar fallos de memoria."
    )

  if process_batch_clicked:
    progress = st.progress(0.0, text="Preparando procesamiento por lotes...")
    status_placeholder = st.empty()

    def _update_batch_progress(done, total, filename, state):
      fraction = 0.0 if total <= 0 else min(max(done / total, 0.0), 1.0)
      progress.progress(
        fraction,
        text=f"{state}: {filename} ({min(done, total)}/{total})",
      )
      status_placeholder.caption(
        "Cada error queda aislado; el procesamiento continúa con el siguiente paciente."
      )

    try:
      result = process_pac_pdf_batch(
        batch_pdf_files,
        metodo_calibracion=batch_calibration,
        metodo_rmri=batch_rmri,
        logo_png=logo_png,
        firma_png=firma_png,
        sello_png=sello_png,
        include_originals=batch_include_originals,
        save_to_history=batch_save_history,
        progress_callback=_update_batch_progress,
      )
      result["input_signature"] = _batch_files_signature(batch_pdf_files)
      result["generated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
      st.session_state["pac_batch_result"] = result
      progress.progress(1.0, text="Lote finalizado")
      status_placeholder.empty()
    except Exception as exc:
      st.error(f"No se pudo iniciar o completar el procesamiento por lotes: {exc}")

  batch_result = st.session_state.get("pac_batch_result")
  if batch_result:
    current_signature = _batch_files_signature(batch_pdf_files) if batch_pdf_files else ""
    if current_signature and current_signature != batch_result.get("input_signature"):
      st.warning(
        "La selección actual de PDF cambió después del último procesamiento. "
        "Los archivos descargables corresponden al lote previamente generado."
      )

    st.markdown("### Resultado del lote")
    mr1, mr2, mr3 = st.columns(3)
    mr1.metric("PDF recibidos", batch_result.get("total", 0))
    mr2.metric("Informes generados", batch_result.get("success_count", 0))
    mr3.metric("Errores", batch_result.get("error_count", 0))

    if batch_result.get("history_total") is not None:
      st.success(
        f"Historial actualizado sin borrar registros previos. Total actual: "
        f"{batch_result['history_total']} estudios."
      )

    summary_batch_df = batch_result.get("summary_df", pd.DataFrame())
    if not summary_batch_df.empty:
      display_cols = [
        c for c in [
          "orden", "archivo_origen", "estado", "paciente", "fecha",
          "pas_central_mmHg", "pp_central_mmHg", "rm_pb_pf",
          "rvse_calculado_pct", "fenotipo_final", "informe_generado", "error"
        ] if c in summary_batch_df.columns
      ]
      st.dataframe(
        summary_batch_df[display_cols],
        use_container_width=True,
        hide_index=True,
      )

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dl1, dl2 = st.columns(2)
    with dl1:
      st.download_button(
        "Descargar lote completo en ZIP",
        data=batch_result.get("zip_bytes", b""),
        file_name=f"INFORMES_PAC_LOTE_{stamp}.zip",
        mime="application/zip",
        use_container_width=True,
        key="download_pac_batch_zip",
      )
    with dl2:
      st.download_button(
        "Descargar resumen del lote en Excel",
        data=batch_result.get("excel_bytes", b""),
        file_name=f"RESUMEN_PAC_LOTE_{stamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="download_pac_batch_excel",
      )
    st.caption(
      "Al pulsar Descargar, el navegador guarda el archivo en la carpeta Descargas configurada en su equipo. "
      "El ZIP contiene una carpeta informes_pdf con un PDF independiente por paciente."
    )

    reports = batch_result.get("reports", [])
    if reports:
      with st.expander("Descargar informes individuales del lote", expanded=False):
        for ridx, item in enumerate(reports):
          label = f"{item.get('patient', 'Paciente')} — {item.get('filename', 'informe.pdf')}"
          st.download_button(
            label,
            data=item.get("data", b""),
            file_name=item.get("filename", f"informe_{ridx+1}.pdf"),
            mime="application/pdf",
            key=f"download_batch_report_{ridx}_{hashlib.md5(item.get('filename','').encode('utf-8')).hexdigest()[:8]}",
          )

    if batch_result.get("error_count", 0):
      with st.expander("Detalle de PDF que no pudieron generar informe", expanded=True):
        err_df = summary_batch_df[summary_batch_df["estado"] != "GENERADO"]
        st.dataframe(
          err_df[[c for c in ["archivo_origen", "error"] if c in err_df.columns]],
          use_container_width=True,
          hide_index=True,
        )

# Cuando solo se cargó un lote, se evita mostrar debajo el formulario individual vacío.
if batch_pdf_files and pdf_file is None:
  st.info("Modo por lotes activo. Los resultados y descargas se administran en el panel anterior.")
  st.stop()

st.markdown("---")
st.subheader("Datos extraídos / edición manual")
cols = st.columns(4)
fields = ["paciente","estudio","fecha","hora","obra_social","edad","sexo","peso","altura","imc","sc","pas_radial","pad_radial","pam_radial","pp_radial","pas_central","pad_central","pam_central","pp_central","fc","au","iau","rvse","pe","apc","medicacion","diagnostico_previo"]
row = {}
for i, f in enumerate(fields):
  with cols[i%4]:
    val = base.get(f, "")
    if f in ["paciente","estudio","fecha","hora","obra_social","sexo","medicacion","diagnostico_previo"]:
      row[f] = st.text_input(f, value="" if pd.isna(val) else str(val))
    else:
      vnum = to_float(val)
      row[f] = st.number_input(f, value=0.0 if np.isnan(vnum) else float(vnum), step=1.0, format="%.2f")

st.markdown("#### Criterio SAHA para hipertensión central")
row["metodo_calibracion_pac"] = st.selectbox(
  "Método de calibración para aplicar tabla SAHA de PAS aórtica central",
  options=["SD_PAOC", "C_PAOC"],
  index=0,
  format_func=lambda x: SAHA_CALIBRATION_LABELS.get(x, x),
  help="Para PAS aórtica central se habilitan solo C_PAOC y SD_PAOC del Manual SAHA; para equipos calibrados con PAS/PAD braquial, use SD_PAOC como criterio operativo."
)
st.info(format_saha_central_htn(row))
st.info(format_saha_aix75(row))

st.markdown("#### Referencia percentilar para RM / RI")
row["metodo_referencia_rmri"] = st.selectbox(
  "Método de referencia para clasificar RM y RI por edad",
  options=["SCOR_RADIAL", "SCOR_CAROTID", "MOG"],
  index=0,
  format_func=lambda x: RM_RI_METHOD_LABELS.get(x, x),
  help=(
    "Los intervalos de referencia cambian según técnica y sitio. Seleccione el método más comparable al estudio. "
    "Para informes con onda central derivada de registro radial, use por defecto Tonometría radial/SphygmoCor. "
    "La referencia es externa y la clasificación debe interpretarse con cautela si el algoritmo de separación no es idéntico al publicado."
  )
)
st.caption("Criterio RM: <P75 esperado; P75-<P90 relativamente elevado; P90-<P95 aumentado; ≥P95 marcadamente aumentado. RI es complementario y no suma doble evidencia diagnóstica.")
st.caption("Tref/Tfor-Tref: variables continuas sin corte fijo universal. Armónicos: HD/H1/H2/H4+ y frecuencia dominante se muestran como métricas continuas de caracterización espectral, sin diagnóstico transversal por cortes fijos.")

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
st.markdown(f"**{dx}**")
st.write(f"Categoría braquial: {cat}. Amplificación PAS periférico-central: {amp_sbp:.1f} mmHg si disponible. Perfil agregado: {risk}.")

if wave_df is not None:
  hdf = harmonic_analysis(wave_df)
  harmonic_metrics_preview = harmonic_distortion_metrics(hdf)
  sep_df_preview, sep_metrics_preview = estimate_wave_separation(wave_df, row)
  conclusion_blocks_preview, sep_df_preview, sep_metrics_preview, sep_interp_preview = build_continuous_conclusions(row, wave_df, hdf)

  summary_cols[3].metric("RM Pb/Pf", f"{sep_metrics_preview.get('rm', np.nan):.2f}")
  summary_cols[4].metric("RVSE calculado", f"{sep_metrics_preview.get('rvse_calculado_%', np.nan):.0f}%")
  st.caption(f"Fuente de curva real: {curve_source or curve_meta.get('metodo','no especificada')}")
  st.caption(f"Firma morfológica de curva real: {sep_metrics_preview.get('curve_id', 'sin_firma')} | Pico: {sep_metrics_preview.get('t_pico_ms', np.nan):.0f} ms | Tref retorno reflejo: {sep_metrics_preview.get('tref_ms', np.nan):.0f} ms")
  if harmonic_metrics_preview.get("ok"):
    st.info(
      f"Análisis armónico: HD energético {harmonic_metrics_preview.get('hd_percent', np.nan):.1f}%; "
      f"H1 {harmonic_metrics_preview.get('h1_energy_percent', np.nan):.1f}%; "
      f"H2 {harmonic_metrics_preview.get('h2_energy_percent', np.nan):.1f}%; "
      f"H4+ {harmonic_metrics_preview.get('h4plus_energy_percent', np.nan):.1f}%; "
      f"frecuencia dominante {harmonic_metrics_preview.get('dominant_frequency_hz', np.nan):.2f} Hz. "
      "Métricas continuas de caracterización espectral; no se usan aisladamente para diagnosticar rigidez ni definir el fenotipo vascular."
    )
    st.caption(longitudinal_harmonic_mdc_note())

  st.markdown("### Animación hemodinámica real de la aorta")
  st.caption("Animación didáctica basada en la curva real del paciente, separación Pf/Pb, flujo aórtico estimado y armónicos. No usa curvas sintéticas ni plantillas fijas.")
  animation_html = render_aortic_real_metrics_animation(row, sep_df_preview, sep_metrics_preview, hdf, height=1080)
  if animation_html:
    components.html(animation_html, height=1120, scrolling=True)
  else:
    st.warning("No fue posible construir la animación con los datos disponibles.")

  st.markdown("### Capturas de métricas alteradas para el informe médico integrado")
  st.caption("Estas imágenes se incorporan al PDF y se seleccionan desde las métricas alteradas reales: cada cuadro incluye valor, criterio y explicación didáctica del mecanismo.")
  try:
    st.image(plot_aortic_animation_keyframes(row, sep_df_preview, sep_metrics_preview, hdf), caption="Capturas de métricas alteradas con valor, criterio y explicación didáctica", use_container_width=True)
  except Exception as e:
    st.warning(f"No se pudieron generar las capturas didácticas: {e}")

  st.markdown("### Conclusiones clínicas diagnósticas")
  for title, body in conclusion_blocks_preview:
    st.markdown(f"**{title}**")
    st.markdown(_markdown_bold_conclusions(body))

  automatic_integrated_preview, automatic_integrated_state_preview = build_automatic_integrated_conclusion(row, sep_metrics_preview, hdf)
  st.markdown("### Conclusión integrada automática")
  st.success(_markdown_bold_conclusions(automatic_integrated_preview))

  st.markdown("---")
  st.markdown("### Gráficos")
  st.image(plot_wave_separation(sep_df_preview, sep_metrics_preview), caption="Presión aórtica central real con onda anterógrada Pf y retrógrada Pb superpuestas y sincronizadas", use_container_width=True)

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
  st.markdown(_markdown_bold_conclusions(final_phenotype_text_preview))
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
  st.error("PDF informe de 1 hoja no habilitado: falta curva real válida del paciente desde CSV/TXT o digitalización del PDF. No se generará reporte con curvas simuladas.")
else:
  pdf_bytes_out = build_pdf_one_page(row, wave_df, hdf, screenshot, firma_png=firma_png, sello_png=sello_png, logo_png=logo_png)
  pdf_download_bytes = ensure_download_bytes(pdf_bytes_out)
  if not pdf_download_bytes:
    st.error("No se pudo generar el PDF informe de 1 hoja.")
  else:
    st.download_button(
      "Generar y descargar PDF informe de 1 hoja",
      data=pdf_download_bytes,
      file_name=nombre_archivo_informe_pac(row),
      mime="application/pdf"
    )
