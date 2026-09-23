# -*- coding: utf-8 -*-
"""
==============================================================================
 TP N°1 - GESTIÓN Y MANTENIMIENTO ELECTROMECÁNICO (UTN FRSR)
 Línea Principal de Elevación y Transporte de Áridos de Alta Densidad
==============================================================================
 Aplicación interactiva (Streamlit + Plotly) para el análisis de confiabilidad,
 taxonomía ISO 14224, matriz de criticidad dinámica, análisis de causa raíz
 (Ishikawa 6M y 5 Porqués), simulación económica de motor Standby 55 kW,
 visualizador de patrones de Nowlan & Heap, dashboard y exportación profesional.

 Uso:
     streamlit run app.py
==============================================================================
"""

import io
import sys
from pathlib import Path

# Permitir importación de módulos auxiliares desde scripts/
scripts_path = Path(__file__).resolve().parent / "scripts"
if str(scripts_path) not in sys.path:
    sys.path.insert(0, str(scripts_path))

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from report_builder import build_technical_report

# ==============================================================================
# 0) CONFIGURACIÓN DE PÁGINA E IDENTIDAD VISUAL (temática ingeniería industrial)
# ==============================================================================
st.set_page_config(page_title="TP N°1 - Gestión y Mantenimiento (UTN FRSR)",
                   layout="wide", initial_sidebar_state="expanded")

AZUL   = "#1F4E79"
AZUL2  = "#2E75B6"
NARANJ = "#E67E22"
ROJO   = "#C0392B"
VERDE  = "#27AE60"

st.markdown("""
<style>
    .stApp { background: #F4F6F8; }
    .baner { background: linear-gradient(90deg,#1F4E79,#2E75B6);
             padding: 0.9rem 1.2rem; border-radius: 8px; color:#fff; }
    .baner h1 { color:#fff; margin:0; font-size:1.35rem; }
    .baner p  { color:#D9E7F2; margin:0; font-size:0.85rem; }
    div[data-testid="stMetric"] {
        background:#fff; border:1px solid #DCE1E6; border-radius:8px;
        padding:0.6rem 0.8rem; box-shadow:0 1px 2px rgba(31,78,121,.08); }
    div[data-testid="stMetric"] label { color:#1F4E79; }
    .tarjeta { background:#fff; border:1px solid #DCE1E6; border-radius:8px;
               padding:0.7rem 0.9rem; margin-bottom:0.5rem; }
    .rca-card { background:#fff; border:1px solid #DCE1E6; border-left:4px solid #1F4E79;
                border-radius:6px; padding:0.8rem 1rem; margin-bottom:0.8rem; }
    .why-step { background:#F8FAFC; border-left:3px solid #2E75B6; border-radius:4px;
                padding:0.6rem 0.9rem; margin-bottom:0.5rem; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="baner">
  <h1>TP N°1 - GESTIÓN Y MANTENIMIENTO ELECTROMECÁNICO</h1>
  <p>UTN FRSR · Ingeniería Electromecánica · Línea de Elevación y Transporte de Áridos de Alta Densidad
  · Confiabilidad, Taxonomía ISO&nbsp;14224, Criticidad, Ishikawa 6M y Decisión de Reemplazo</p>
</div>
""", unsafe_allow_html=True)

st.caption("Herramienta de Ingeniería de Confiabilidad y Mantenimiento - análisis del ciclo anual de 7.200 h")

# ==============================================================================
# 1) BLOQUE DE DATOS EDITABLES (registros y parámetros del enunciado del TP)
# ==============================================================================
DICT_PARAM = {
    "TOP": 7200.0,            # Tiempo Operativo Programado [h/año]
    "costo_parada": 2500.0,   # US$/h de detención (lucro cesante + costo fijo)
    "costo_hh": 25.0,         # US$/hora-hombre de mantenimiento
    "umbral_A": 15,           # Clase A (Crítico)   : C >= 15
    "umbral_B": 10,           # Clase B (Semicrítico): 10 <= C < 15
}

DATOS_FALLAS = [
    {"evento": "F1", "subsistema": "Sistema Motriz",
     "componente": "Rodamiento oscilante 22220 EK (polea motriz)",
     "descripcion": "Picado/spalling en pista exterior por contaminación abrasiva y "
                    "degradación de grasa; requirió desarme completo de la polea motriz.",
     "tarea": "Cambio de rodamiento con desarme total de polea, recalado, alineación y re-tensionado",
     "ttp": 28.0, "repuestos": 1850.0, "hh": 48.0, "t0": 12.0, "cuadrilla": 2.0},
    {"evento": "F2", "subsistema": "Sistema Motriz",
     "componente": "Motor asincrónico trifásico 55 kW (4 polos, 380/660 V)",
     "descripcion": "Cortocircuito entre espiras del estator por bloqueo del ventilador y "
                    "acumulación de polvo en carcasa; se requirió rebobinado completo.",
     "tarea": "Rebobinado completo del estator en taller tercero + desmontaje/remontaje y reconexión",
     "ttp": 42.0, "repuestos": 3400.0, "hh": 24.0, "t0": 40.0, "cuadrilla": 1.0},
    {"evento": "F3", "subsistema": "Sistema de Tensión",
     "componente": "Latiguillo de alta presión (circuito oleohidráulico 20 bar)",
     "descripcion": "Fuga masiva por fisura del latiguillo por fatiga de pulsación; caída de "
                    "presión en el tensionado y disparo del interbloqueo.",
     "tarea": "Reemplazo de latiguillo HP + cebado, purga de aire y prueba de presostatos/interbloqueo",
     "ttp": 6.0, "repuestos": 320.0, "hh": 12.0, "t0": 3.0, "cuadrilla": 2.0},
    {"evento": "F4", "subsistema": "Banda y Rodillos",
     "componente": "Cinta transportadora (rasgado longitudinal de 15 m de goma superior)",
     "descripcion": "Desalineación severa por acumulación de material en tambor de reenvío, con "
                    "rasgado longitudinal; vulcanizado en caliente por terceros.",
     "tarea": "Vulcanizado en caliente (terceros) de 15 m de goma + realineación y limpieza de reenvío",
     "ttp": 36.0, "repuestos": 4200.0, "hh": 36.0, "t0": 18.0, "cuadrilla": 2.0},
    {"evento": "F5", "subsistema": "Banda y Rodillos",
     "componente": "Rascador primario de uretano",
     "descripcion": "Desgaste prematuro y atrancamiento por fricción contra la banda; sobrecarga "
                    "térmica del reductor y variación de corriente del motor.",
     "tarea": "Destrabe / reemplazo de hoja de uretano + limpieza de la zona de retorno",
     "ttp": 8.0, "repuestos": 450.0, "hh": 16.0, "t0": 4.0, "cuadrilla": 2.0},
]

ESCALAS = {
    "F": "Frecuencia anual: 1=Rara (1), 2=Ocasional (2), 3=Frecuente (3-4), 4=Muy Frecuente (>=5)",
    "S": "Seguridad de las personas: 1=Sin riesgo, 2=Accidente menor, 3=Accidente grave/fatal",
    "A": "Medio ambiente: 1=Sin impacto, 2=Impacto local, 3=Impacto externo severo",
    "O": "Impacto operacional: 1=Sin parada de planta, 3=Parada parcial, 10=Parada total",
    "D": "Costo directo: 1=<1.000 USD, 2=1.000-3.000 USD, 4=>3.000 USD",
}

CRITICIDAD_DEF = [
    {"Subsistema": "Sistema Motriz",             "F": 3, "S": 2, "A": 1, "O": 10, "D": 4},
    {"Subsistema": "Banda y Rodillos",           "F": 3, "S": 2, "A": 1, "O": 10, "D": 4},
    {"Subsistema": "Sistema de Tensión",         "F": 2, "S": 2, "A": 2, "O":  3, "D": 1},
    {"Subsistema": "Estructura / Tolva",         "F": 2, "S": 1, "A": 1, "O":  3, "D": 1},
    {"Subsistema": "Sistema Eléctrico / Seguridad","F": 1, "S": 2, "A": 1, "O":  3, "D": 1},
]

NIVELES_ISO = {
    1: "Industria",
    2: "Instalación / Planta",
    3: "Área de proceso",
    4: "Unidad de equipo",
    5: "Sistema del equipo",
    6: "Subsistema / Ítem de mantenimiento",
    7: "Parte componente",
    8: "Elemento de desgaste / material",
}


def taxonomia_df():
    """Árbol canónico ISO 14224 del TP N°1."""
    datos = [
        (1, "IND-MIN", "Industria Minera / Procesamiento de Minerales", "Industria", None),
        (2, "PLT-SR",  "Planta de Molienda y Beneficio San Rafael",    "Instalación / Planta", "IND-MIN"),
        (3, "ARE-TRN", "Área de Elevación y Transporte de Áridos",     "Área de proceso", "PLT-SR"),
        (4, "CV-01",   "Cinta Transportadora Principal de Áridos",     "Unidad de equipo", "ARE-TRN"),
        (5, "SYS-MOT", "Sistema Motriz y Reducción",                   "Sistema", "CV-01"),
        (6, "MOT-01",  "Motor asincrónico 55 kW (4 polos, 380/660V)",  "Componente", "SYS-MOT"),
        (6, "RED-01",  "Reductor ortogonal de tres etapas (i=28:1)",   "Componente", "SYS-MOT"),
        (6, "POL-MOT", "Polea motriz Ø500 mm revestida y eje 4140",     "Componente", "SYS-MOT"),
        (6, "ROD-01",  "Rodamiento oscilante SKF 22220 EK (lado acopl)","Componente", "POL-MOT"),
        (6, "ROD-02",  "Rodamiento oscilante SKF 22220 EK (lado libre)","Componente", "POL-MOT"),
        (5, "SYS-BAN", "Banda y Elementos de Rodadura",                "Sistema", "CV-01"),
        (6, "BAN-01",  "Banda transportadora continua multicapa",      "Componente", "SYS-BAN"),
        (6, "TAM-REE", "Tambor de reenvío / retorno Ø400 mm",          "Componente", "SYS-BAN"),
        (6, "ROD-CAR", "Estaciones de rodillos de carga (triple art)", "Componente", "SYS-BAN"),
        (6, "ROD-RET", "Rodillos de retorno planos",                   "Componente", "SYS-BAN"),
        (6, "RAS-01",  "Rascador primario frontal de uretano",         "Componente", "SYS-BAN"),
        (6, "RAS-02",  "Rascador secundario de hojas de carburo",      "Componente", "SYS-BAN"),
        (5, "SYS-TEN", "Sistema Oleohidráulico de Tensionado",         "Sistema", "CV-01"),
        (6, "CEN-HID", "Centralita hidráulica (bomba 20 bar, tanque)", "Componente", "SYS-TEN"),
        (6, "CIL-HID", "Cilindro hidráulico de doble efecto",          "Componente", "SYS-TEN"),
        (6, "LAT-HP",  "Latiguillos flexibles de alta presión",        "Componente", "SYS-TEN"),
        (6, "VAL-PRE", "Válvulas reguladoras y presostatos",           "Componente", "SYS-TEN"),
        (5, "SYS-EST", "Estructura, Bastidor y Guías",                 "Sistema", "CV-01"),
        (6, "BAS-MET", "Bastidor metálico estructural y largueros",    "Componente", "SYS-EST"),
        (6, "TOL-CAR", "Tolva de carga y recepción de áridos",         "Componente", "SYS-EST"),
        (6, "GUI-LAT", "Guías laterales y baberos de confinamiento",   "Componente", "SYS-EST"),
        (5, "SYS-ELE", "Sistema Eléctrico, Mando y Seguridad",         "Sistema", "CV-01"),
        (6, "TAB-POT", "Tablero de potencia y arrancador suave",       "Componente", "SYS-ELE"),
        (6, "SEN-DES", "Sensores inductivos de desalineación",         "Componente", "SYS-ELE"),
        (6, "PUL-EME", "Interruptores de tiro por cable (emergencia)", "Componente", "SYS-ELE"),
    ]
    return pd.DataFrame(datos, columns=["Nivel", "Código", "Denominación", "Tipo", "Código padre"])


def _build_tree(df):
    hijos = {}
    raices = []
    for _, r in df.iterrows():
        p = r["Código padre"]
        c = r["Código"]
        if pd.isna(p) or p is None or p == "":
            raices.append(c)
        else:
            hijos.setdefault(p, []).append(c)
    return hijos, raices


def dataframe_directo():
    jer = taxonomia_df()
    denom = jer.set_index("Código")["Denominación"].to_dict()

    def desc_padre(cod):
        return denom.get(cod, "—")

    df = jer.copy()
    df["Descripción del nivel"] = df["Nivel"].map(NIVELES_ISO)
    df["Descripción del equipo padre"] = df["Código padre"].fillna("").map(desc_padre)
    df = df.rename(columns={
        "Código": "Código del equipo",
        "Denominación": "Descripción del equipo",
        "Tipo": "Tipo de equipo",
        "Código padre": "Equipo padre",
    })
    return df[["Nivel", "Descripción del nivel", "Código del equipo",
               "Descripción del equipo", "Tipo de equipo", "Equipo padre",
               "Descripción del equipo padre"]]


def dataframe_escalonado():
    df = dataframe_directo()
    jer = taxonomia_df()
    denom = jer.set_index("Código")["Denominación"].to_dict()
    hijos, raices = _build_tree(jer)

    ruta = {}
    def recorrer(cod, prefijo):
        if prefijo == "":
            ruta[cod] = denom[cod]
        else:
            ruta[cod] = prefijo + " > " + denom[cod]
        for h in hijos.get(cod, []):
            recorrer(h, ruta[cod])

    for r in raices:
        recorrer(r, "")

    df["Ruta completa"] = df["Código del equipo"].map(ruta)
    return df


FICHAS = {
    "Sistema Motriz": [
        ("Motor asincrónico trifásico 55 kW",
         "Imparte la potencia motriz al accionamiento de la cinta.",
         "4 polos, 380/660 V, 50 Hz, aislamiento Clase F, IP55, arranque por arrancador suave.",
         "F2 - Cortocircuito entre espiras del estator (rebobinado)"),
        ("Reductor ortogonal de ejes paralelos (3 etapas)",
         "Reduce la velocidad del motor y aumenta el par sobre la polea motriz.",
         "Relación i=28:1, lubricación por barboteo con aceite sintético ISO VG 220.",
         "Sobrecarga térmica por arrastre del rascador (encadenado de F5)"),
        ("Polea motriz Ø500 mm + eje SAE 4140 + rodamientos 22220 EK",
         "Transmite el movimiento por fricción a la banda.",
         "Revestimiento cerámico, soportes monobloc SNH, rodamiento oscilante de doble hilera con manguito de apriete.",
         "F1 - Picado/spalling de pista exterior"),
    ],
    "Banda y Rodillos": [
        ("Banda transportadora de goma",
         "Elemento de transporte continuo del árido de alta densidad.",
         "Goma superior e inferior con refuerzo textil/antimetal; vulcanizable en caliente.",
         "F4 - Rasgado longitudinal de 15 m de goma superior"),
        ("Tambor de reenvío y rodillos de apoyo",
         "Guían y soportan el ramal de retorno y carga de la banda.",
         "Tambor de reenvío con acumulación de material (foco de desalineación).",
         "F4 - Desalineación por material acumulado"),
        ("Rascadores primario (uretano) y secundario (metal duro)",
         "Limpieza de la cara portante de la banda para evitar arrastre y desgaste.",
         "Hoja de uretano de desgaste prematuro por fricción.",
         "F5 - Atrancamiento del rascador de uretano"),
    ],
    "Estructura / Tolva": [
        ("Estructura metálica y caballetes",
         "Sustenta la cinta y los órganos móviles.", "Acero estructural con protecciones, sin fallas registradas en el ciclo.", "-"),
        ("Tolva de recepción de árido",
         "Recibe y orienta el flujo de material hacia la banda.", "Diseño de alta densidad, revestimiento antidesgaste.", "-"),
    ],
    "Sistema de Tensión": [
        ("Centralita oleohidráulica (bomba de engranajes)",
         "Genera la presión para el tensionado automático de la banda.",
         "Circuito de 20 bar con tanque, filtros y válvulas de seguridad.",
         "F3 - Caída de presión y disparo del interbloqueo"),
        ("Cilindro hidráulico de doble efecto y latiguillos HP",
         "Aplica la tensión directa al tambor de cola.",
         "Latiguillo de alta presión fisurado por fatiga de pulsación.",
         "F3 - Fuga masiva de fluido por fisura"),
        ("Presostatos de seguridad",
         "Interbloquean el circuito hidráulico con el motor de accionamiento.",
         "Interbloqueo de presión mínima / máxima.", "F3 - Disparo del interbloqueo"),
    ],
    "Sistema Eléctrico / Seguridad": [
        ("Tablero eléctrico y arrancador suave",
         "Alimenta, protege y arranca gradualmente el motor de 55 kW.",
         "Protecciones térmicas y de cortocircuito, maniobra local/remota.", "-"),
        ("Sensores inductivos de desalineación de banda",
         "Detectan el desplazamiento lateral excesivo de la cinta.",
         "Salidas al sistema de control y disparo por seguridad.",
         "Preventivo del modo de falla F4"),
        ("Tirador de emergencia y circuitos de seguridad",
         "Permite detener la línea en forma inmediata desde cualquier punto.",
         "Cable de tiro con retorno por resortes, corte de mando registrado.", "-"),
    ],
}

PLAN_INSPECCIONES = [
    (1, "Centralita oleohidráulica y latiguillos", "Sensorial (TPM)", "Fugas de fluido, nivel, fijaciones", "Inspección visual", "Diaria", "Operador (TPM)", "Sin goteo ni fisuras"),
    (2, "Rodamiento polea motriz (ambos soportes)", "Sensorial (TPM)", "Ruido anómalo y temperatura al tacto", "Auditivo + táctil", "Diaria", "Operador (TPM)", "T < 60 °C al tacto"),
    (3, "Banda y tambor de reenvío", "Sensorial (TPM)", "Centrado, desalineación, material acumulado", "Inspección visual", "Diaria", "Operador (TPM)", "Banda centrada / retorno limpio"),
    (4, "Motor eléctrico 55 kW", "Sensorial (TPM)", "Limpieza de carcasa, ventilador, rejillas", "Inspección visual", "Semanal", "Operador (TPM)", "Carcasa y ventilador libres de polvo"),
    (5, "Rascador primario de uretano", "Sensorial (TPM)", "Desgaste y atrancamiento de la hoja", "Visual + táctil", "Semanal", "Operador (TPM)", "Espesor de hoja sobre límite"),
    (6, "Motor eléctrico 55 kW", "Instrumental (PdM)", "Temperatura de carcasa", "Pirómetro infrarrojo", "Mensual", "Mantenimiento (PdM)", "Alerta si carcasa > 80 °C"),
    (7, "Motor eléctrico 55 kW", "Instrumental (PdM)", "Corriente por fase / desbalance", "Pinza amperimétrica", "Mensual", "Mantenimiento (PdM)", "Desbalance < 10 % / I < In"),
    (8, "Reductor (aceite ISO VG 220)", "Instrumental (PdM)", "Temperatura del aceite en cárter", "Pirómetro infrarrojo", "Mensual", "Mantenimiento (PdM)", "Alerta si aceite > 80 °C"),
    (9, "Rodamientos polea (22220 EK)", "Instrumental (PdM)", "Vibración global y bandas de falla", "Vibrómetro (ISO 10816)", "Trimestral", "Mantenimiento (PdM)", "Sin picos crecientes en banda de rodamiento"),
    (10, "Reductor y rodamientos", "Instrumental (PdM)", "Fricción / descargas eléctricas", "Análisis por ultrasonido", "Trimestral", "Mantenimiento (PdM)", "Nivel dB estable"),
    (11, "Centralita oleohidráulica", "Instrumental (PdM)", "Presión de trabajo (20 bar) y presostatos", "Manómetro / prueba de interbloqueo", "Mensual", "Mantenimiento (PdM)", "Desviación < ±10 %"),
    (12, "Cárter del reductor", "Instrumental (PdM)", "Estado del lubricante VG 220", "Muestreo y análisis de aceite", "Trimestral", "Mantenimiento (PdM)", "Contaminación < límites"),
]

# ==============================================================================
# 2) FUNCIONES DE CÁLCULO
# ==============================================================================
def factor_frecuencia(n):
    if n >= 5:
        return 4
    if n >= 3:
        return 3
    if n == 2:
        return 2
    return 1


def clasificar(c, umbral_A, umbral_B):
    if c >= umbral_A:
        return ("Clase A - CRÍTICO",
                "Alta Disponibilidad: monitoreo continuo + overhaul programado")
    if c >= umbral_B:
        return ("Clase B - SEMICRÍTICO",
                "Sistemático / Predictivo: re-lubricación y mediciones por calendario/condición")
    return ("Clase C - NO CRÍTICO",
            "Condicional / TPM: inspección CIL por operador y atención a la falla")


def estimar_repeticiones(df_ed):
    df = df_ed.copy()
    df["HH_oc"] = df["t0"] * df["cuadrilla"]
    df["N_ttp"] = np.floor(df["ttp"] / df["t0"]).clip(lower=1).astype(int)
    df["N_hh"]  = np.floor(df["hh"] / df["HH_oc"]).clip(lower=1).astype(int)
    df["N_est"] = [max(a, b) for a, b in zip(df["N_ttp"], df["N_hh"])]
    df["Revisar"] = df["N_ttp"] != df["N_hh"]
    return df


LETRAS_OCURRENCIA = "abcdefghijklmnopqrstuvwxyz"


def _aplicar_ediciones_estim(df_estim):
    st_key = "editor_estim"
    if st_key in st.session_state:
        edited = st.session_state[st_key].get("edited_rows", {})
        for ridx, cambios in edited.items():
            i = int(ridx)
            if i >= df_estim.shape[0]:
                continue
            for col, val in cambios.items():
                if col in df_estim.columns:
                    df_estim.loc[i, col] = val
    return df_estim


def generar_fallas_supuesto(con_rep, df_estim, standby=False):
    """Base canónica de fallas según el supuesto vigente y si el motor Standby está activo."""
    base = pd.DataFrame(DATOS_FALLAS)
    if standby:
        base.loc[base["evento"] == "F2", "ttp"] = 4.0
    cols = ["evento", "subsistema", "componente", "descripcion", "tarea",
            "ttp", "repuestos", "hh", "t0", "cuadrilla"]
    if not con_rep:
        return base[cols].copy()
    n_por_evento = df_estim.set_index("evento")["N_est"].to_dict()
    filas = []
    for _, r in base.iterrows():
        n = max(1, int(n_por_evento.get(r["evento"], 1)))
        for k in range(1, n + 1):
            fila = r.to_dict()
            fila["evento"] = f"{r['evento']}{LETRAS_OCURRENCIA[k - 1]}"
            fila["ttp"] = r["ttp"] / n
            fila["repuestos"] = r["repuestos"] / n
            fila["hh"] = r["hh"] / n
            filas.append(fila)
    return pd.DataFrame(filas)[cols]


if st.session_state.pop("reiniciar_registro", False):
    for k in ("editor_fallas", "editor_estim", "usar_n_est", "sup_aplicado__fallas", "standby_aplicado__fallas"):
        st.session_state.pop(k, None)
    st.session_state["fallas_custom"] = []


# ==============================================================================
# 3) BARRA LATERAL - PARÁMETROS, NÓMINA DINÁMICA Y SIMULADOR STANDBY
# ==============================================================================
with st.sidebar:
    st.header("⚙️ Parámetros dinámicos")
    st.caption("Recálculo instantáneo en vivo de todos los KPIs.")

    TOP_anio   = st.number_input("Tiempo Operativo Programado (TOP)", min_value=1000.0,
                                 max_value=10000.0, value=DICT_PARAM["TOP"],
                                 step=100.0, format="%.0f", help="h/año de diseño")
    costo_par  = st.number_input("Costo de parada de planta (USD/h)", min_value=100.0,
                                 max_value=10000.0, value=DICT_PARAM["costo_parada"],
                                 step=50.0, format="%.0f")
    costo_hh   = st.number_input("Costo de Mano de Obra (USD/HH)", min_value=1.0,
                                 max_value=500.0, value=DICT_PARAM["costo_hh"],
                                 step=1.0, format="%.0f")

    st.divider()
    st.subheader("📊 Umbrales de Criticidad Total")
    umbral_A = st.slider("Clase A (Crítico):  C ≥", min_value=8, max_value=30,
                         value=DICT_PARAM["umbral_A"], step=1, key="umbral_A")
    b_max = max(2, umbral_A - 1)
    if st.session_state.get("umbral_B", -1) > b_max:
        st.session_state["umbral_B"] = b_max
    umbral_B = st.slider("Clase B (Semicrítico):  C ≥", min_value=2, max_value=b_max,
                         value=min(DICT_PARAM["umbral_B"], b_max), step=1, key="umbral_B")

    st.divider()
    usar_n_est = st.toggle(
        "Supuesto con repetición (N_est ≈ 9 fallas/año)",
        value=False, key="usar_n_est",
        help="Desglosa cada falla registrada en sus ocurrencias repetidas según estimación por tareas.",
    )

    st.divider()
    st.subheader("⚡ Motor Standby 55 kW (R3)")
    standby_motor = st.toggle(
        "Activar Motor Standby en Pañol",
        value=False,
        key="standby_motor",
        help="Simula la compra y almacenamiento de un motor de 55 kW en pañol. "
             "Reduce el MTTR de F2 a 4.0 h (ahorro directo de $95,000 USD de lucro cesante)."
    )
    if standby_motor:
        st.success("🟢 **Motor Standby ACTIVO**\n\n"
                   "• MTTR F2: **4.0 h** (antes 42.0 h)\n"
                   "• Ahorro Neto Parada: **+$95,000 USD**\n"
                   "• MTTR Global: **16.40 h** (antes 24.0 h)\n"
                   "• Disponibilidad: **98.86%** (+0.53%)\n"
                   "• ROI: **2,275%** | Payback: **15.3 días**")

    st.divider()
    with st.expander("👥 Nómina de Alumnos / Grupo (R2)", expanded=True):
        st.caption("Cargue o modifique los integrantes del equipo:")
        if "students" not in st.session_state:
            st.session_state["students"] = [
                {"nombre": "Martín", "apellido": "Méndez", "legajo": "11335"}
            ]
        
        df_stud_ui = pd.DataFrame(st.session_state["students"])
        edited_stud = st.data_editor(
            df_stud_ui,
            num_rows="dynamic",
            column_config={
                "nombre": st.column_config.TextColumn("Nombre", required=True),
                "apellido": st.column_config.TextColumn("Apellido", required=True),
                "legajo": st.column_config.TextColumn("Legajo", required=True),
            },
            key="editor_students",
            width="stretch",
            hide_index=True
        )
        st.session_state["students"] = edited_stud.to_dict("records")
        ciclo_lectivo = st.text_input("Ciclo Lectivo", value="2026", key="ciclo_lectivo")

    st.divider()
    with st.expander("Escalas del TP (referencia)"):
        for k, v in ESCALAS.items():
            st.markdown(f"**{k}.** {v}")

# ==============================================================================
# 4) DATAFRAMES BASE Y ESTIMACIÓN DE REPETICIONES
# ==============================================================================
df_fallas_base = pd.DataFrame(DATOS_FALLAS)
df_estim = df_fallas_base[["evento", "subsistema", "tarea", "ttp", "hh", "repuestos",
                           "t0", "cuadrilla"]].copy()
df_estim = _aplicar_ediciones_estim(df_estim)
df_estim = estimar_repeticiones(df_estim)
N_est_total = int(df_estim["N_est"].sum())

N_por_sub = df_estim.groupby("subsistema")["N_est"].sum()
F_sugerido = N_por_sub.apply(factor_frecuencia)

# ==============================================================================
# 5) PESTAÑAS PRINCIPALES DE LA APLICACIÓN
# ==============================================================================
tab_taxo, tab_kpi, tab_matriz, tab_rca, tab_con_decision, tab_dash, tab_excel, tab_informe = st.tabs([
    "1 · Taxonomía (ISO 14224)",
    "2 · Registro de Fallas y KPIs",
    "3 · Matriz de Criticidad Dinámica",
    "4 · Análisis Causa Raíz (Ishikawa & 5 Porqués)",
    "5 · Confiabilidad y Decisión (Nowlan & Heap / Standby)",
    "6 · Dashboard (Plotly)",
    "7 · Exportación a Excel",
    "8 · Informe Técnico Oficial (.docx)"
])

# -------------------------------------------------------------------------------
# PESTAÑA 1 - TAXONOMÍA DEL ACTIVO (NIVELES ISO 14224)
# -------------------------------------------------------------------------------
with tab_taxo:
    st.subheader("Jerarquía del activo - Taxonomía ISO 14224")
    st.caption("Industria → Planta → Área → Equipo → Sistema → Componente. "
               "Establece la trazabilidad de decisiones de Reemplazo (MTTF) vs. Reparación (MTBF).")

    modo = st.radio(
        "Modo de visualización",
        ["Escalonado (árbol)", "Listado directo (tabla)"],
        horizontal=True,
    )

    def to_csv_bytes(df):
        return df.to_csv(index=False).encode("utf-8-sig")

    if modo == "Escalonado (árbol)":
        jer = taxonomia_df()
        hijos, raices = _build_tree(jer)
        etiqueta = {
            r["Código"]: (int(r["Nivel"]), r["Tipo"], r["Código"], r["Denominación"])
            for _, r in jer.iterrows()
        }
        colores = {1: "#2E75B6", 2: "#2E8B57", 3: "#E67E22", 4: "#8E44AD", 5: "#1F4E79", 6: "#7F8C8D"}

        def mostrar_nodo(codigo):
            nivel, tipo, cod, denom = etiqueta[codigo]
            color = colores.get(nivel, "#1F4E79")
            hijos_nodo = hijos.get(codigo, [])
            etiqueta_linea = f"(L{nivel}) {tipo} — **[ {cod} ]** {denom}"
            if hijos_nodo:
                with st.expander(etiqueta_linea + f"  ·  {len(hijos_nodo)} hijo(s)",
                                 expanded=(nivel <= 1)):
                    st.markdown(
                        f"<div style='border-left:3px solid {color};padding-left:8px;"
                        f"color:#555;font-size:0.85rem'>{denom}</div>",
                        unsafe_allow_html=True)
                    for h in hijos_nodo:
                        mostrar_nodo(h)
            else:
                st.markdown(
                    f"<div style='border-left:3px solid {color};padding:2px 8px;"
                    f"margin:2px 0'><b>[{cod}]</b> <span"
                    f"style='color:{color}'>(L{nivel}) {tipo}</span> · {denom}</div>",
                    unsafe_allow_html=True)

        for r in raices:
            mostrar_nodo(r)

        col_csv, col_n = st.columns([1, 3])
        with col_csv:
            st.download_button(
                "⬇  Descargar taxonomía escalonada (CSV)",
                data=to_csv_bytes(dataframe_escalonado()),
                file_name="taxonomia_escalonada.csv",
                mime="text/csv",
            )
    else:
        df_dir = dataframe_directo()
        st.dataframe(df_dir, width="stretch", hide_index=True)
        st.download_button(
            "⬇  Descargar listado directo (CSV)",
            data=to_csv_bytes(df_dir),
            file_name="taxonomia_directa.csv",
            mime="text/csv",
        )

# -------------------------------------------------------------------------------
# PESTAÑA 2 - REGISTRO DE FALLAS Y KPIS
# -------------------------------------------------------------------------------
with tab_kpi:
    est = st.session_state
    if "sup_aplicado__fallas" not in est:
        est["sup_aplicado__fallas"] = usar_n_est
    if "standby_aplicado__fallas" not in est:
        est["standby_aplicado__fallas"] = standby_motor
    if "fallas_custom" not in est:
        est["fallas_custom"] = []

    if (est["sup_aplicado__fallas"] != usar_n_est) or (est["standby_aplicado__fallas"] != standby_motor):
        est["sup_aplicado__fallas"] = usar_n_est
        est["standby_aplicado__fallas"] = standby_motor
        if "editor_fallas" in est:
            del est["editor_fallas"]

    st.subheader("Registro editable de fallas del ciclo anual")
    col_tit, col_reset = st.columns([3, 1])
    with col_tit:
        st.caption("Los costos y KPIs se recalculan en tiempo real según el costo de parada ($2,500 USD/h), "
                   "mano de obra ($25 USD/HH) y el estado del motor Standby.")
    with col_reset:
        if st.button("Restaurar datos originales del TP",
                     help="Vuelve a los 5 registros oficiales del TP."):
            st.session_state["reiniciar_registro"] = True
            st.rerun()

    base_fallas = generar_fallas_supuesto(usar_n_est, df_estim, standby=standby_motor)
    etiquetas_canon = set(base_fallas["evento"])
    df_pasado = pd.concat([base_fallas, pd.DataFrame(est["fallas_custom"])],
                          ignore_index=True)

    col_editor = {
        "evento": st.column_config.TextColumn("Evento", width="small", help="Código de falla"),
        "subsistema": st.column_config.TextColumn("Subsistema"),
        "componente": st.column_config.TextColumn("Componente", width="large"),
        "ttp": st.column_config.NumberColumn("TTP (h)", min_value=0.0, step=0.5, format="%.1f"),
        "repuestos": st.column_config.NumberColumn("Repuestos (USD)", min_value=0.0, step=50.0, format="%.0f"),
        "hh": st.column_config.NumberColumn("HH", min_value=0.0, step=1.0, format="%.0f"),
    }
    df_edit = st.data_editor(
        df_pasado,
        column_config=col_editor,
        num_rows="dynamic",
        column_order=["evento", "subsistema", "componente", "ttp", "repuestos", "hh"],
        key="editor_fallas", width="stretch", hide_index=True)

    df_custom = df_edit[~df_edit["evento"].isin(etiquetas_canon)]
    est["fallas_custom"] = df_custom.to_dict("records")

    df_fallas = df_edit.copy()
    for c in ("ttp", "repuestos", "hh"):
        df_fallas[c] = pd.to_numeric(df_fallas[c], errors="coerce").fillna(0.0)
    df_fallas["Costo_Parada_USD"] = df_fallas["ttp"] * costo_par
    df_fallas["Costo_MO_USD"]     = df_fallas["hh"]  * costo_hh
    df_fallas["Costo_Total_USD"]  = (df_fallas["Costo_Parada_USD"]
                                     + df_fallas["repuestos"]
                                     + df_fallas["Costo_MO_USD"])

    N_activo = len(df_fallas)
    TTP_total = float(df_fallas["ttp"].sum())
    TFR = TOP_anio - TTP_total
    if N_activo > 0:
        MTBF = TFR / N_activo
        MTTR = TTP_total / N_activo
    else:
        MTBF, MTTR = float("nan"), float("nan")
    Ai = MTBF / (MTBF + MTTR) * 100 if (MTBF + MTTR) > 0 else 0
    Ao = TFR / TOP_anio * 100

    cols_tabla = ["evento", "subsistema", "ttp", "Costo_Parada_USD", "repuestos",
                  "hh", "Costo_MO_USD", "Costo_Total_USD"]
    tabla = df_fallas[cols_tabla].copy()
    tabla.columns = ["Evento", "Subsistema", "TTP (h)", "Costo Parada (USD)",
                     "Repuestos (USD)", "HH", "Costo MO (USD)", "Costo Total (USD)"]
    total_row = pd.DataFrame([["TOTAL", "—", TTP_total,
                               df_fallas["Costo_Parada_USD"].sum(),
                               df_fallas["repuestos"].sum(),
                               df_fallas["hh"].sum(),
                               df_fallas["Costo_MO_USD"].sum(),
                               df_fallas["Costo_Total_USD"].sum()]],
                             columns=tabla.columns)
    tabla_completa = pd.concat([tabla, total_row], ignore_index=True)

    st.dataframe(
        tabla_completa.style.format({
            "TTP (h)": "{:.1f}", "Costo Parada (USD)": "${:,.0f}",
            "Repuestos (USD)": "${:,.0f}", "HH": "{:.0f}",
            "Costo MO (USD)": "${:,.0f}", "Costo Total (USD)": "${:,.0f}",
        }),
        width="stretch", hide_index=True)

    st.markdown("### Indicadores Clave de Desempeño (KPIs)")
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("TOP (Diseño)", f"{TOP_anio:,.0f} h")
    k2.metric("TTP (Paradas)", f"{TTP_total:,.1f} h",
              delta="-38.0 h" if standby_motor else None)
    k3.metric("TFR (Operación)", f"{TFR:,.1f} h",
              delta="+38.0 h" if standby_motor else None)
    k4.metric("MTBF", f"{MTBF:,.1f} h")
    k5.metric("MTTR Global", f"{MTTR:,.2f} h",
              delta="-7.60 h" if standby_motor else None, delta_color="inverse")
    k6.metric("Disponibilidad (Ao)", f"{Ao:.2f}%",
              delta="+0.53%" if standby_motor else None)

# -------------------------------------------------------------------------------
# PESTAÑA 3 - MATRIZ DE CRITICIDAD DINÁMICA
# -------------------------------------------------------------------------------
with tab_matriz:
    st.subheader("Matriz de Criticidad Dinámica (C = F × Ci)")
    st.caption("Ajuste los selectores discretos oficiales (F, S, A, O, D). "
               "Ci, C, Clase y Modelo se recalculan en vivo.")

    matriz0 = pd.DataFrame(CRITICIDAD_DEF).copy()
    matriz0["F"] = matriz0["Subsistema"].map(F_sugerido).fillna(1).astype(int)
    matriz0["N_est/año"] = matriz0["Subsistema"].map(N_por_sub).fillna(0).astype(int)

    col_mat = {
        "Subsistema": st.column_config.TextColumn("Subsistema", disabled=True),
        "F": st.column_config.SelectboxColumn("F", options=[1, 2, 3, 4], help=ESCALAS["F"]),
        "N_est/año": st.column_config.NumberColumn("N_est (fallas/año)", disabled=True),
        "S": st.column_config.SelectboxColumn("S", options=[1, 2, 3], help=ESCALAS["S"]),
        "A": st.column_config.SelectboxColumn("A", options=[1, 2, 3], help=ESCALAS["A"]),
        "O": st.column_config.SelectboxColumn("O", options=[1, 3, 10], help=ESCALAS["O"]),
        "D": st.column_config.SelectboxColumn("D", options=[1, 2, 4], help=ESCALAS["D"]),
    }
    df_mat_edit = st.data_editor(
        matriz0, column_config=col_mat,
        disabled=["Subsistema", "N_est/año"],
        width="stretch", hide_index=True, key="editor_matriz")

    df_matriz = df_mat_edit.copy()
    df_matriz["Ci"] = df_matriz["S"] + df_matriz["A"] + df_matriz["O"] + df_matriz["D"]
    df_matriz["C"] = df_matriz["F"] * df_matriz["Ci"]
    df_matriz[["Clase", "Modelo"]] = df_matriz["C"].apply(
        lambda c: pd.Series(clasificar(c, umbral_A, umbral_B)))

    mapa_clase = {"Clase A - CRÍTICO": ROJO, "Clase B - SEMICRÍTICO": NARANJ,
                  "Clase C - NO CRÍTICO": VERDE}
    st.dataframe(
        df_matriz.style.map(lambda v: f"background-color:{mapa_clase.get(v, '')}",
                            subset=["Clase"]).format({"Ci": "{:.0f}", "C": "{:.0f}"}),
        width="stretch", hide_index=True,
        column_config={"Modelo": st.column_config.TextColumn("Modelo asignado", width="large")})

# -------------------------------------------------------------------------------
# PESTAÑA 4 - ANÁLISIS DE CAUSA RAÍZ (ISHIKAWA 6M Y 5 PORQUÉS) (R1)
# -------------------------------------------------------------------------------
with tab_rca:
    st.subheader("Análisis de Causa Raíz de Contingencias (Ishikawa 6M & 5 Porqués)")
    st.caption("Desglose estructurado de causas primarias, secundarias y organizacionales (RCM / TPM).")

    st.markdown("#### 1. Diagrama de Causa y Efecto — Ishikawa 6M")
    m_col1, m_col2 = st.columns(2)

    with m_col1:
        st.markdown("""
        <div class="rca-card">
            <h4 style="color:#1F4E79;margin-top:0;">⚙️ Maquinaria (Machine)</h4>
            <ul>
                <li><b>F1:</b> Retén de labio sin deflector exterior en soporte monobloc SNH; holgura residual incorrecta en manguito cónico.</li>
                <li><b>F2:</b> Malla de tobera de ventilador con paso demasiado cerrado para ambientes con polvo; aletas con poca separación.</li>
                <li><b>F3:</b> Circuito de 20 bar sin acumulador hidroneumático ni amortiguador de pulsaciones; latiguillo con radio inferior al mínimo.</li>
                <li><b>F4:</b> Tambor de reenvío liso sin tambor autolimpiante en jaula de ardilla ni rascador interior en V en retorno.</li>
                <li><b>F5:</b> Soporte basculante rígido con resorte sin mecanismo de compensación de desgaste de hoja.</li>
            </ul>
        </div>
        <div class="rca-card">
            <h4 style="color:#1F4E79;margin-top:0;">👷 Mano de Obra (Manpower)</h4>
            <ul>
                <li><b>F1:</b> Sobredosis manual de grasa con grasera sin retirar tapón de purga (sobrepresión interna que reventó el labio).</li>
                <li><b>F2:</b> Omisión de inspección táctil/visual de descarga de aire en rondas de turno; desatención a zumbidos térmicos.</li>
                <li><b>F3:</b> Montaje de latiguillo aplicando torsión axial indebida sobre el caucho al ajustar racores.</li>
                <li><b>F4:</b> Falta de verificación visual de centrado de banda en cola durante el arranque diario.</li>
                <li><b>F5:</b> Falta de ajuste periódico de tensión y verificación de escuadra del filo de la hoja de uretano.</li>
            </ul>
        </div>
        <div class="rca-card">
            <h4 style="color:#1F4E79;margin-top:0;">📋 Método (Method)</h4>
            <ul>
                <li><b>F1:</b> Lubricación rígida por días de calendario en lugar de lubricación asistida por ultrasonido acústico.</li>
                <li><b>F2:</b> Inexistencia de un estándar CIL (Mantenimiento Autónomo TPM) para soplado semanal de carcasas de motores.</li>
                <li><b>F3:</b> Política de 'run-to-failure' en latiguillos hidráulicos en vez de sustitución preventiva a 24 meses.</li>
                <li><b>F4:</b> Procedimiento que no contempla parada inmediata ante acumulación en tolva de transferencia.</li>
                <li><b>F5:</b> Falta de un checklist de 3 minutos al inicio de turno para limpieza y calibración de rascadores.</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    with m_col2:
        st.markdown("""
        <div class="rca-card">
            <h4 style="color:#1F4E79;margin-top:0;">🧱 Materiales (Materials)</h4>
            <ul>
                <li><b>F1:</b> Grasa mineral básica con viscosidad insuficiente y sin aditivos sellantes para polvo de cuarzo abrasivo.</li>
                <li><b>F2:</b> Costra mineral compactada en carcasa con baja conductividad térmica (k < 0.2 W/m·K) bloqueando el enfriamiento.</li>
                <li><b>F3:</b> Manguera de una sola malla de acero (1SN) sometida a fatiga por pulsación cíclica a 20 bar.</li>
                <li><b>F4:</b> Cubierta de caucho estándar sin cordones transversales antidesgarro (rip-stop breakers de aramida).</li>
                <li><b>F5:</b> Poliuretano de baja resiliencia que sufre desgaste abrasivo asimétrico y desprendimiento.</li>
            </ul>
        </div>
        <div class="rca-card">
            <h4 style="color:#1F4E79;margin-top:0;">🌡️ Medio Ambiente (Milieu / Entorno)</h4>
            <ul>
                <li><b>F1:</b> Atmósfera altamente pulvígena en cabezal motriz por caída de finos abrasivos desde tolva.</li>
                <li><b>F2:</b> Ambiente confinado y pulvígeno con temperatura ambiente estival elevada (> 38 °C) en recinto del motor.</li>
                <li><b>F3:</b> Exposición solar UV y salpicaduras de mineral que degradan y resecan la cubierta externa de caucho.</li>
                <li><b>F4:</b> Caída accidental de terrones de áridos sobre el ramal inferior por tolva sin baberos estancos.</li>
                <li><b>F5:</b> Finos húmedos que forman pasta arcillosa de alta adherencia sobre la cara motriz de la cinta.</li>
            </ul>
        </div>
        <div class="rca-card">
            <h4 style="color:#1F4E79;margin-top:0;">📏 Medición (Measurement)</h4>
            <ul>
                <li><b>F1:</b> Inexistencia de monitoreo periódico de vibraciones con demodulación de envolvente (gE / BPFO).</li>
                <li><b>F2:</b> Falta de sondas termométricas PT-100 en estator con corte en arrancador suave; sin termografía de rutina.</li>
                <li><b>F3:</b> Falta de manómetro testigo con glicerina para registrar picos de sobrepresión y golpes de ariete.</li>
                <li><b>F4:</b> Sensores inductivos de desalineación descalibrados o bloqueados por barro mineral.</li>
                <li><b>F5:</b> Falta de sensor térmico en reductor ortogonal conectado al SCADA y sin monitoreo de corriente por fase.</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    st.divider()
    st.markdown("#### 2. Árbol de los 5 Porqués (Cadena Causal Profunda)")

    sel_why = st.radio("Seleccione el Evento Crítico a Analizar:",
                       ["⚡ Falla F2: Quemado de Estator en Motor 55 kW por Bloqueo de Ventilador",
                        "⚙️ Falla F5: Atrancamiento de Rascador de Uretano y Efecto Dominó en Reductor"],
                       horizontal=True)

    if "F2" in sel_why:
        st.markdown("""
        **Problema:** Motor de 55 kW sufre sobrecalentamiento y cortocircuito entre espiras en estator, deteniendo la planta por 42 h ($109,000 USD).
        """)
        st.markdown("""
        <div class="why-step"><b>1. ¿Por qué ocurrió el cortocircuito entre espiras?</b><br>
        <i>Nivel Físico Inmediato:</i> La película dieléctrica de esmalte (Clase F) de los conductores se degradó y quemó al superar los 160 °C.</div>
        <div class="why-step"><b>2. ¿Por qué se superaron los 160 °C en el motor?</b><br>
        <i>Nivel Físico / Mecánico:</i> Colapsó la disipación térmica: el ventilador forzado se bloqueó por suciedad y la carcasa quedó cubierta por polvo aislante (k < 0.2 W/m·K).</div>
        <div class="why-step"><b>3. ¿Por qué se bloquearon el ventilador y las aletas sin ser advertidos?</b><br>
        <i>Nivel Humano / Operacional:</i> Los operadores y mantenedores no verificaron la descarga de aire ni palparon la temperatura de la carcasa durante sus rondas.</div>
        <div class="why-step"><b>4. ¿Por qué el personal de planta no detectó la acumulación de polvo?</b><br>
        <i>Nivel de Método / Control:</i> El motor no figuraba en las hojas de control del Mantenimiento Autónomo (TPM CIL semanal) y no se aplicaba termografía infrarroja periódica.</div>
        <div class="why-step" style="border-left-color:#C0392B;background:#FBE5E5;"><b>5. ¿Por qué no estaban programadas las rutinas de limpieza ni la protección térmica?</b><br>
        <b style="color:#C0392B;">Causa Raíz Latente / Organizacional:</b> La dirección mantenía una política de mantenimiento reactivo ('a la rotura') en accionamientos, sin haber aplicado RCM que evidenciara que la falla de un motor de $4,000 USD induciría $105,000 USD de parada y riesgo de petrificación en el reactor.</div>
        """, unsafe_allow_html=True)

        st.info("""
        🛡️ **Acciones Preventivas de Bloqueo Formuladas para F2:**
        1. **Poka-Yoke / Físico:** Instalar 3 sondas PT-100 embebidas en devanados conectadas a disparo del arrancador suave a 145 °C (alarma a 130 °C).
        2. **Operativo / TPM:** Implementar estándar semanal CIL con soplado y aspirado de carcasa y tobera con planilla firmada.
        3. **Estratégico:** Adquirir 1 motor de reserva Standby de 55 kW en pañol de planta para limitar el MTTR a 4 h ante cualquier contingencia.
        """)
    else:
        st.markdown("""
        **Problema:** Reductor ortogonal trabaja en sobretemperatura (>95 °C) y sobrecorriente en motor por atrancamiento de rascador contra la banda (8 h, $20,850 USD).
        """)
        st.markdown("""
        <div class="why-step"><b>1. ¿Por qué sobrecalentó el aceite del reductor y subió la corriente del motor?</b><br>
        <i>Nivel Físico Inmediato:</i> La cinta demandó una cupla resistente excesiva ($T_{res} \\uparrow$), forzando a los engranajes a disipar calor por fricción severa y al motor a absorber corriente de sobrecarga ($P = 3 I^2 R$).</div>
        <div class="why-step"><b>2. ¿Por qué se disparó la cupla resistente en la cinta?</b><br>
        <i>Nivel Físico / Mecánico:</i> La hoja de uretano se inclinó y trabó en ángulo agudo contra la goma de la cinta, transformándose en una zapata de freno continuo.</div>
        <div class="why-step"><b>3. ¿Por qué se trabó la hoja del rascador?</b><br>
        <i>Nivel Físico / Componente:</i> La lámina sufrió desgaste asimétrico y la acumulación de finos en el soporte basculante anuló la carrera del resorte compensador.</div>
        <div class="why-step"><b>4. ¿Por qué no se limpió la costra ni se reguló el resorte?</b><br>
        <i>Nivel Humano / Operacional:</i> Los operadores no realizaron la inspección sensorial diaria (ronda CIL de 3 minutos) para verificar el estado de los filos y la articulación.</div>
        <div class="why-step" style="border-left-color:#C0392B;background:#FBE5E5;"><b>5. ¿Por qué no existía la ronda diaria ni sensores en el reductor?</b><br>
        <b style="color:#C0392B;">Causa Raíz Latente / Organizacional:</b> La supervisión consideraba a los rascadores como 'accesorios de limpieza sin impacto operacional', subestimando el acoplamiento físico y efecto dominó sobre motor y reductor; ausencia de un programa estructurado de TPM.</div>
        """, unsafe_allow_html=True)

        st.info("""
        🛡️ **Acciones Preventivas de Bloqueo Formuladas para F5:**
        1. **Metodológica / TPM:** Rutina diaria obligatoria CIL (3 minutos por cambio de turno) para limpieza de costras en el eje porta-rascador y regulación del resorte.
        2. **Poka-Yoke / Mecánico:** Soporte basculante autolimpiante con fusible mecánico desacoplable ante sobreesfuerzo de frenado y deflector de goma.
        3. **Predictivo:** Termocupla PT-100 en tapón de drenaje del reductor conectada al PLC con parada de emergencia a 85 °C para proteger el lubricante ISO VG 220.
        """)

# -------------------------------------------------------------------------------
# PESTAÑA 5 - CONFIABILIDAD AVANZADA Y DECISIÓN (NOWLAN & HEAP / STANDBY) (R3)
# -------------------------------------------------------------------------------
with tab_con_decision:
    st.subheader("Confiabilidad Avanzada y Herramientas de Decisión (R3)")

    st.markdown("#### 1. Análisis de los 6 Patrones de Falla de Nowlan & Heap (A a F)")
    st.caption("Comparativa interactiva de la tasa condicional de falla λ(t) demostrando la prevalencia moderna de los Patrones D, E y F.")

    t = np.linspace(0, 100, 200)
    # Curvas de Nowlan & Heap
    # A: Bañera (mortalidad infantil + meseta + desgaste terminal)
    curva_A = 0.5 * np.exp(-t / 15) + 0.1 + 0.0008 * (t ** 2)
    # B: Desgaste terminal tradicional
    curva_B = 0.1 + 0.0005 * (np.maximum(0, t - 50) ** 2)
    # C: Aumento gradual monótono
    curva_C = 0.1 + 0.005 * t
    # D: Asentamiento inicial y constante
    curva_D = 0.05 + 0.25 * (1 - np.exp(-t / 10))
    # E: Tasa aleatoria perfectamente constante
    curva_E = np.full_like(t, 0.22)
    # F: Mortalidad infantil pronunciada + constante
    curva_F = 0.6 * np.exp(-t / 8) + 0.15

    fig_nh = go.Figure()
    fig_nh.add_trace(go.Scatter(x=t, y=curva_A, mode="lines", name="Patrón A: Curva Bañera (4%)", line=dict(color="#C0392B", width=2.5)))
    fig_nh.add_trace(go.Scatter(x=t, y=curva_B, mode="lines", name="Patrón B: Desgaste Tradicional (2%)", line=dict(color="#E67E22", width=2)))
    fig_nh.add_trace(go.Scatter(x=t, y=curva_C, mode="lines", name="Patrón C: Aumento Gradual (5%)", line=dict(color="#F39C12", width=2)))
    fig_nh.add_trace(go.Scatter(x=t, y=curva_D, mode="lines", name="Patrón D: Asentamiento Inicial (7%)", line=dict(color="#2980B9", width=2.5)))
    fig_nh.add_trace(go.Scatter(x=t, y=curva_E, mode="lines", name="Patrón E: Falla Aleatoria Constante (14%)", line=dict(color="#27AE60", width=3)))
    fig_nh.add_trace(go.Scatter(x=t, y=curva_F, mode="lines", name="Patrón F: Mortalidad Infantil (68%)", line=dict(color="#8E44AD", width=3.5)))

    fig_nh.update_layout(
        title="Los 6 Patrones de Falla de Nowlan & Heap (Aeronáutica / Industria Compleja)",
        xaxis_title="Tiempo Operativo o Edad del Activo (t)",
        yaxis_title="Tasa Condicional de Falla λ(t)",
        template="plotly_white",
        height=450,
        legend=dict(orientation="h", y=-0.2)
    )
    st.plotly_chart(fig_nh, width="stretch")

    nh_c1, nh_c2 = st.columns(2)
    with nh_c1:
        st.markdown("""
        **Distribución de Prevalencia Industrial:**
        - **Patrones Dependientes de la Edad (A, B, C):** **11%** (¡Solo el 4% responde a la bañera A!).
        - **Patrones Aleatorios e Incondicionales (D, E, F):** **89%** (Patrón F dominante con 68%).
        """)
    with nh_c2:
        st.markdown("""
        **Justificación Técnica para la Cinta de Áridos ($F_1$ a $F_5$):**
        - El 100% de las fallas obedecen a contingencias aleatorias por polvo abrasivo, barro o fatiga por pulsación.
        - **Peligro de Overhaul por Calendario:** Desarmar periódicamente el equipo introduce desalineación o daños en retenes, induciendo fallas infantiles (**Patrón F**). La estrategia correcta es **CBM + TPM**.
        """)

    st.divider()
    st.markdown("#### 2. Simulador Económico Interactivo: Motor Standby de 55 kW")
    st.caption("Evaluación financiera de la adquisición y almacenamiento de 1 motor idéntico en pañol de planta.")

    sb_c1, sb_c2, sb_c3 = st.columns(3)
    inv_motor = sb_c1.number_input("Costo de Inversión del Motor Standby (USD)", min_value=1000.0, max_value=15000.0, value=4000.0, step=250.0)
    mttr_standby = sb_c2.number_input("Nuevo MTTR in situ con Standby (horas)", min_value=1.0, max_value=20.0, value=4.0, step=0.5)
    costo_posesion = sb_c3.number_input("Costo Anual de Almacenamiento / Posesión (USD)", min_value=0.0, max_value=2000.0, value=400.0, step=50.0)

    # Cálculos económicos en vivo
    ahorro_horas_f2 = 42.0 - mttr_standby
    ahorro_downtime_f2 = ahorro_horas_f2 * costo_par
    nuevo_ttp_global = 120.0 - ahorro_horas_f2
    nuevo_mttr_global = nuevo_ttp_global / 5.0
    nueva_disp_global = ((TOP_anio - nuevo_ttp_global) / TOP_anio) * 100
    beneficio_neto_y1 = ahorro_downtime_f2 - inv_motor - costo_posesion
    roi_standby = ((ahorro_downtime_f2 - inv_motor) / inv_motor) * 100 if inv_motor > 0 else 0
    payback_dias = (inv_motor / (ahorro_downtime_f2 / 365.0)) if ahorro_downtime_f2 > 0 else 0
    horas_prod_recup = inv_motor / costo_par

    sm1, sm2, sm3, sm4, sm5 = st.columns(5)
    sm1.metric("Ahorro Neto Parada", f"${ahorro_downtime_f2:,.0f} USD", help="Lucro cesante ahorrado en F2")
    sm2.metric("MTTR Global Nuevo", f"{nuevo_mttr_global:.2f} h", delta=f"-{(24.0 - nuevo_mttr_global):.2f} h", delta_color="inverse")
    sm3.metric("Disponibilidad Nueva", f"{nueva_disp_global:.2f}%", delta=f"+{(nueva_disp_global - 98.33):.2f}%")
    sm4.metric("ROI de la Inversión", f"{roi_standby:,.0f}%", help="Retorno sobre el costo del motor")
    sm5.metric("Período de Payback", f"{payback_dias:.1f} días", help=f"Equivale a {horas_prod_recup:.1f} h de producción de planta")

    st.markdown("##### Comparativa Lado a Lado: Caso Base vs. Escenario con Motor Standby")
    df_comp = pd.DataFrame([
        {"Parámetro": "MTTR de Falla F2 (Motor 55 kW)", "Caso Base (Histórico)": "42.0 horas", "Con Motor Standby": f"{mttr_standby:.1f} horas", "Variación": f"-{ahorro_horas_f2:.1f} h"},
        {"Parámetro": "Costo de Parada en F2", "Caso Base (Histórico)": "$105,000 USD", "Con Motor Standby": f"${(mttr_standby * costo_par):,.0f} USD", "Variación": f"-${ahorro_downtime_f2:,.0f} USD"},
        {"Parámetro": "Costo Total del Evento F2", "Caso Base (Histórico)": "$109,000 USD", "Con Motor Standby": f"${(mttr_standby * costo_par + 3800):,.0f} USD", "Variación": f"-${(ahorro_downtime_f2 + 200):,.0f} USD"},
        {"Parámetro": "Tiempo Total de Paradas (TTP Global)", "Caso Base (Histórico)": "120.0 horas", "Con Motor Standby": f"{nuevo_ttp_global:.1f} horas", "Variación": f"-{ahorro_horas_f2:.1f} h"},
        {"Parámetro": "MTTR Global del Sistema", "Caso Base (Histórico)": "24.00 horas", "Con Motor Standby": f"{nuevo_mttr_global:.2f} horas", "Variación": f"-{(24.0 - nuevo_mttr_global):.2f} h"},
        {"Parámetro": "Disponibilidad Operacional (Ao)", "Caso Base (Histórico)": "98.33%", "Con Motor Standby": f"{nueva_disp_global:.2f}%", "Variación": f"+{(nueva_disp_global - 98.33):.2f}%"},
        {"Parámetro": "Inversión en Motor Nuevo", "Caso Base (Histórico)": "$0 USD", "Con Motor Standby": f"${inv_motor:,.0f} USD", "Variación": f"+${inv_motor:,.0f} USD"},
        {"Parámetro": "Beneficio Neto Año 1", "Caso Base (Histórico)": "$0 USD", "Con Motor Standby": f"${beneficio_neto_y1:,.0f} USD", "Variación": f"+${beneficio_neto_y1:,.0f} USD"},
    ])
    st.table(df_comp)

# -------------------------------------------------------------------------------
# PESTAÑA 6 - DASHBOARD (PLOTLY)
# -------------------------------------------------------------------------------
with tab_dash:
    dfp = df_fallas.sort_values("Costo_Total_USD", ascending=False).reset_index(drop=True)
    acum_pct = dfp["Costo_Total_USD"].cumsum() / dfp["Costo_Total_USD"].sum() * 100
    fig1 = make_subplots(specs=[[{"secondary_y": True}]])
    fig1.add_trace(go.Bar(x=dfp["evento"], y=dfp["Costo_Total_USD"], name="Costo total",
                          marker_color=[f"rgba(46,117,182,{0.55 + 0.45 * (1 - i / max(1, len(dfp) - 1))})"
                                        for i in range(len(dfp))],
                          text=[f"${v:,.0f}" for v in dfp["Costo_Total_USD"]],
                          textposition="outside", cliponaxis=False,
                          hovertemplate="%{x}<br>$%{y:,.0f}<extra>Costos</extra>"),
                   secondary_y=False)
    fig1.add_trace(go.Scatter(x=dfp["evento"], y=acum_pct, name="% acumulado",
                              mode="lines+markers", line=dict(color=ROJO, width=2.4),
                              hovertemplate="%{x}<br>%{y:.0f}%<extra></extra>"),
                   secondary_y=True)
    fig1.add_hline(y=80, line_dash="dash", line_color="gray", secondary_y=True,
                   annotation_text="Regla 80/20", annotation_position="top right")
    fig1.update_yaxes(title="Costo total [USD]", secondary_y=False, tickprefix="$",
                      tickformat=",.0f")
    fig1.update_yaxes(title="% acumulado", secondary_y=True, range=[0, 105])
    fig1.update_layout(title="G1 · Diagrama de Pareto - Costo Total por Evento",
                       template="plotly_white", height=420, showlegend=False)

    dfm = df_fallas.sort_values("ttp")
    fig2 = go.Figure(go.Bar(
        x=dfm["evento"], y=dfm["ttp"],
        marker_color=[NARANJ if i % 2 else AZUL for i in range(len(dfm))],
        text=dfm["ttp"].apply(lambda v: f"{v:.1f} h"), textposition="outside",
        hovertemplate="%{x} · MTTR %{y:.1f} h<extra></extra>"))
    fig2.update_layout(title="G2 · Tiempo de Reparación (MTTR) por evento",
                       xaxis_title="Evento de falla", yaxis_title="Tiempo [h]",
                       template="plotly_white", height=420)

    df_matriz_fig = df_matriz.copy()
    fig3 = go.Figure()
    for cls in df_matriz_fig["Clase"].unique():
        sub = df_matriz_fig[df_matriz_fig["Clase"] == cls]
        fig3.add_trace(go.Scatter(
            x=sub["Ci"], y=sub["F"], mode="markers+text", name=cls,
            text=sub["Subsistema"], textposition="top center",
            marker=dict(size=sub["C"].clip(lower=8) * 2.2, color=mapa_clase[cls],
                        opacity=0.85, line=dict(color="white", width=1.5)),
            hovertemplate="%{text}<br>Ci=%{x:.0f} · F=%{y}<br>C=%{customdata:.0f}<extra></extra>",
            customdata=sub["C"]))
    fig3.update_layout(title="G3 · Matriz de Criticidad - Frecuencia vs Consecuencia (Ci)",
                       xaxis_title="Consecuencia Ci = S+A+O+D", yaxis_title="Frecuencia F",
                       template="plotly_white", height=420,
                       legend=dict(orientation="h", y=-0.2))

    labels_costos = ["Costo de Parada", "Repuestos / Materiales", "Mano de Obra"]
    valores_costos = [df_fallas["Costo_Parada_USD"].sum(), df_fallas["repuestos"].sum(),
                      df_fallas["Costo_MO_USD"].sum()]
    fig4 = go.Figure(go.Pie(
        labels=labels_costos, values=valores_costos, hole=0.58,
        marker=dict(colors=[ROJO, AZUL2, VERDE], line=dict(color="white", width=2)),
        textinfo="label+percent", insidetextorientation="horizontal",
        hovertemplate="%{label}<br>$%{value:,.0f} (%{percent})<extra></extra>"))
    fig4.update_layout(title=(f"G4 · Estructura de Costos Global "
                              f"(Total ${sum(valores_costos):,.0f} USD)"),
                       template="plotly_white", height=420, showlegend=True)

    g1, g2 = st.columns(2)
    g1.plotly_chart(fig1, width="stretch")
    g2.plotly_chart(fig2, width="stretch")
    g3, g4 = st.columns(2)
    g3.plotly_chart(fig3, width="stretch")
    g4.plotly_chart(fig4, width="stretch")

# -------------------------------------------------------------------------------
# PESTAÑA 7 - EXPORTACIÓN A EXCEL (openpyxl)
# -------------------------------------------------------------------------------
def generar_excel(TOP, cost_par, cost_hh, uA, uB, usar_est,
                  df_fallas_, df_matriz_, df_estim_, n_est, mbf, mtr, ai, ao):
    AZULX, GRISX, AMBARX, VERDEX = "1F4E79", "F2F2F2", "FFF2CC", "E2EFDA"
    _t = Side(style="thin", color="B7B7B7")
    BORD = Border(left=_t, right=_t, top=_t, bottom=_t)
    F_TIT = Font(name="Calibri", size=13, bold=True, color="1F4E79")
    F_SUB = Font(name="Calibri", size=10, italic=True, color="7F7F7F")
    F_HDR = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    F_TXT = Font(name="Calibri", size=10)
    F_BLD = Font(name="Calibri", size=10, bold=True)
    USD = '"$"#,##0'

    def titulo(ws, ncols, texto, sub=None):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
        ws.cell(1, 1, texto).font = F_TIT
        ws.row_dimensions[1].height = 24
        fila = 1
        if sub:
            fila = 2
            ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
            ws.cell(2, 1, sub).font = F_SUB
            ws.row_dimensions[2].height = 16
        return fila + 1

    def cabecera(ws, fila, headers):
        for j, h in enumerate(headers, start=1):
            c = ws.cell(fila, j, h)
            c.font = F_HDR
            c.fill = PatternFill("solid", fgColor=AZULX)
            c.border = BORD
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        return fila + 1

    def anchos(ws, valores):
        for j, w in enumerate(valores, start=1):
            ws.column_dimensions[get_column_letter(j)].width = w

    wb = Workbook()

    # Hoja 1: KPIs
    ws = wb.active
    ws.title = "KPIs_Confiabilidad"
    fila = titulo(ws, 4, "KPIs de Confiabilidad y Disponibilidad",
                  "Línea de Transporte de Áridos - ciclo anual - UTN FRSR")
    fila = cabecera(ws, fila, ["Métrica / Indicador", "Valor", "Unidad", "Fórmula y observación"])
    kpis = [
        ("Tiempo Operativo Programado (TOP)", f"{TOP:,.0f}", "h", "24 h/día x 300 días/año (diseño)"),
        ("Tiempo Total de Paradas (TTP)", f"{df_fallas_['ttp'].sum():,.0f}", "h", "Σ de los TTP de F1-F5"),
        ("Tiempo de Funcionamiento Real (TFR)", f"{TOP - df_fallas_['ttp'].sum():,.0f}", "h", "TFR = TOP - TTP"),
        ("Cantidad de Fallas - oficial (N)", f"{len(df_fallas_)}", "fallas", "Registradas en el enunciado"),
        ("Cantidad de Fallas - estimada (N_est)", f"{n_est}", "fallas", "Por repetición según tareas (t0 y HH)"),
        ("MTBF - Tiempo Medio Entre Fallas", f"{mbf:,.2f}", "h", "MTBF = TFR / N"),
        ("MTTR - Tiempo Medio Para Reparar", f"{mtr:,.2f}", "h", "MTTR = TTP / N"),
        ("Disponibilidad Inherente (Ai)", f"{ai:.2f}", "%", "Ai = MTBF / (MTBF + MTTR)"),
        ("Disponibilidad Operacional (Ao)", f"{ao:.2f}", "%", "Ao = TFR / TOP"),
    ]
    for i, (m, v, u, fo) in enumerate(kpis, start=fila):
        ws.cell(i, 1, m).font = F_TXT
        cv = ws.cell(i, 2, v); cv.font = F_BLD; cv.alignment = Alignment(horizontal="right")
        ws.cell(i, 3, u).alignment = Alignment(horizontal="center")
        ws.cell(i, 4, fo).font = F_TXT
        for j in range(1, 5):
            ws.cell(i, j).border = BORD
        if i % 2 == 0:
            for j in range(1, 5):
                ws.cell(i, j).fill = PatternFill("solid", fgColor=GRISX)
    anchos(ws, [42, 16, 10, 70])

    # Hoja 2: Registro de fallas / finanzas
    ws2 = wb.create_sheet("Registro_Fallas_Finanzas")
    fila = titulo(ws2, 9, "Registro de Fallas y Evaluación Económica",
                  f"Parada = TTP x ${cost_par:,.0f}/h | MO = HH x ${cost_hh:,.0f}/HH")
    fila = cabecera(ws2, fila, ["Evento", "Subsistema", "Componente", "TTP (h)", "Costo Parada (USD)",
                                "Repuestos (USD)", "HH", "Costo MO (USD)", "Costo Total (USD)"])
    fc = df_fallas_.copy()
    fc["Costo_Parada_USD"] = fc["ttp"] * cost_par
    fc["Costo_MO_USD"] = fc["hh"] * cost_hh
    fc["Costo_Total_USD"] = fc["Costo_Parada_USD"] + fc["repuestos"] + fc["Costo_MO_USD"]
    for i, r in fc.iterrows():
        fila_r = fila + i
        for j, v in enumerate([r["evento"], r["subsistema"], r["componente"], r["ttp"],
                               r["Costo_Parada_USD"], r["repuestos"], r["hh"],
                               r["Costo_MO_USD"], r["Costo_Total_USD"]], start=1):
            c = ws2.cell(fila_r, j, v)
            c.border = BORD; c.font = F_TXT
    tot = ["TOTAL", "", "", fc["ttp"].sum(), fc["Costo_Parada_USD"].sum(), fc["repuestos"].sum(),
           fc["hh"].sum(), fc["Costo_MO_USD"].sum(), fc["Costo_Total_USD"].sum()]
    fila_t = fila + len(fc)
    for j, v in enumerate(tot, start=1):
        c = ws2.cell(fila_t, j, v)
        c.border = BORD; c.font = F_BLD; c.fill = PatternFill("solid", fgColor=AMBARX)
    for j, nf in {4: "0.0", 5: USD, 6: USD, 7: "0", 8: USD, 9: USD}.items():
        ws2.cell(fila_t, j).number_format = nf
    for i in range(len(fc)):
        for j, nf in {4: "0.0", 5: USD, 6: USD, 7: "0", 8: USD, 9: USD}.items():
            ws2.cell(fila + i, j).number_format = nf
    anchos(ws2, [9, 22, 44, 12, 17, 16, 9, 14, 17])

    # Hoja 3: Matriz de criticidad
    ws3 = wb.create_sheet("Matriz_Criticidad")
    fila = titulo(ws3, 11, "Matriz de Criticidad - Frecuencia x Consecuencia",
                  f"Ci = S+A+O+D · C = F x Ci · A≥{uA} · B: {uB}≤C<{uA} · C<{uB}")
    fila = cabecera(ws3, fila, ["Subsistema", "F", "N_est/año", "S", "A", "O", "D", "Ci", "C",
                                "Clasificación", "Modelo asignado"])
    fill_clase = {"Clase A - CRÍTICO": "C0392B", "Clase B - SEMICRÍTICO": "E67E22",
                  "Clase C - NO CRÍTICO": "27AE60"}
    for i, r in df_matriz_.iterrows():
        fila_r = fila + i
        for j, v in enumerate([r["Subsistema"], r["F"], r["N_est/año"], r["S"], r["A"],
                               r["O"], r["D"], r["Ci"], r["C"], r["Clase"], r["Modelo"]], start=1):
            c = ws3.cell(fila_r, j, v)
            c.border = BORD; c.font = F_TXT
        for j in range(2, 10):
            ws3.cell(fila_r, j).alignment = Alignment(horizontal="center")
        cc = ws3.cell(fila_r, 10)
        cc.fill = PatternFill("solid", fgColor=fill_clase[r["Clase"]])
        cc.font = Font(bold=True, color="FFFFFF", size=10)
    anchos(ws3, [28, 6, 14, 6, 6, 6, 6, 6, 8, 18, 52])

    # Hoja 4: Plan de inspecciones
    ws4 = wb.create_sheet("Plan_Inspecciones")
    fila = titulo(ws4, 8, "Plan de Ruta de Inspecciones",
                  "Sensoriales (TPM / operador) e Instrumentales (Mantenimiento / PdM)")
    fila = cabecera(ws4, fila, ["N°", "Punto de Inspección", "Tipo", "Parámetro Controlado",
                                "Técnica / Herramienta", "Frecuencia", "Responsable", "Criterio de Alerta"])
    for i, row in enumerate(PLAN_INSPECCIONES, start=fila):
        for j, v in enumerate(row, start=1):
            c = ws4.cell(i, j, v)
            c.border = BORD; c.font = F_TXT
            c.alignment = Alignment(vertical="top", wrap_text=True)
        if row[2].startswith("Instrumental"):
            for j in range(1, 9):
                ws4.cell(i, j).fill = PatternFill("solid", fgColor=VERDEX)
    anchos(ws4, [5, 34, 20, 34, 30, 12, 22, 38])

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


with tab_excel:
    st.subheader("Exportación del informe profesional a Excel")
    st.markdown("Se genera **`TP1_Gestion_Mantenimiento_UTN.xlsx`** con 4 hojas formateadas "
                "usando los datos y parámetros vigentes de la aplicación:")
    for hoja, desc in [
        ("KPIs_Confiabilidad", "Indicadores TOP, TTP, TFR, N / N_est, MTBF, MTTR, Ai y Ao"),
        ("Registro_Fallas_Finanzas", "Tabla económica F1-F5 con Parada, Repuestos, MO y TOTAL"),
        ("Matriz_Criticidad", "F, S, A, O, D, Ci, C, clasificación y modelo asignado"),
        ("Plan_Inspecciones", "Ruta sensorial (TPM) e instrumental (PdM) con frecuencias")]:
        st.markdown(f"- **{hoja}** · {desc}")

    datos_excel = generar_excel(
        TOP=TOP_anio, cost_par=costo_par, cost_hh=costo_hh,
        uA=umbral_A, uB=umbral_B, usar_est=usar_n_est,
        df_fallas_=df_fallas, df_matriz_=df_matriz, df_estim_=df_estim,
        n_est=N_est_total, mbf=MTBF, mtr=MTTR, ai=Ai, ao=Ao)

    st.download_button(
        label="⬇  Descargar TP1_Gestion_Mantenimiento_UTN.xlsx",
        data=datos_excel, file_name="TP1_Gestion_Mantenimiento_UTN.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -------------------------------------------------------------------------------
# PESTAÑA 8 - INFORME TÉCNICO OFICIAL DE CÁTEDRA (.DOCX)
# -------------------------------------------------------------------------------
with tab_informe:
    st.subheader("Informe Técnico Oficial de Cátedra — Justificación Integral")
    st.caption("Documento formal que responde y justifica las consignas del enunciado oficial de UTN FRSR "
               "con rigor de ingeniería electromecánica, sincronizado en vivo con la nómina de alumnos y supuestos.")

    col_inf1, col_inf2, col_inf3 = st.columns(3)
    n_alumnos = len(st.session_state.get("students", []))
    col_inf1.metric("Alumnos del Grupo (R2)", f"{n_alumnos} integrante(s)",
                    help=", ".join([f"{s.get('apellido','')} {s.get('nombre','')} ({s.get('legajo','')})" for s in st.session_state.get("students", [])]))
    col_inf2.metric("Ciclo / Régimen", f"{ciclo_lectivo} · {TOP_anio:,.0f} h/año",
                    help=f"Costo Parada: ${costo_par:,.0f} USD/h | MO: ${costo_hh:,.0f} USD/HH")
    estado_motor_str = "Standby 55 kW ACTIVO" if standby_motor else "Caso Base Nominal"
    col_inf3.metric("Configuración Motriz", estado_motor_str,
                    help="Determina el MTTR de F2 y los KPIs de parada en el informe.")

    # Generación en memoria del archivo DOCX con los alumnos dinámicos y estado de standby
    docx_buffer = build_technical_report(
        students=st.session_state.get("students", []),
        standby_active=standby_motor,
        anio=ciclo_lectivo,
        TOP=TOP_anio,
        cost_parada=costo_par,
        cost_hh=costo_hh,
        df_fallas=df_fallas,
        kpis_nominal={"ttp": TTP_total, "tfr": TOP_anio - TTP_total,
                      "mtbf": (TOP_anio - TTP_total) / max(1, N_activo),
                      "mttr": TTP_total / max(1, N_activo),
                      "ai": ((TOP_anio - TTP_total) / TOP_anio) * 100,
                      "ao": ((TOP_anio - TTP_total) / TOP_anio) * 100,
                      "n": N_activo},
        kpis_rep={"ttp": TTP_total, "tfr": TOP_anio - TTP_total,
                  "mtbf": (TOP_anio - TTP_total) / max(1, N_est_total),
                  "mttr": TTP_total / max(1, N_est_total),
                  "ai": ((TOP_anio - TTP_total) / TOP_anio) * 100,
                  "ao": ((TOP_anio - TTP_total) / TOP_anio) * 100,
                  "n": N_est_total},
        df_matriz=df_matriz
    )

    st.download_button(
        label="📥 Descargar Informe Técnico Oficial (.docx)",
        data=docx_buffer.getvalue(),
        file_name="TP1_Informe_Tecnico_Electromecanico_UTN.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        help="Descarga el documento Word (.docx) formateado con estilos institucionales UTN FRSR y todas las respuestas completas."
    )

st.markdown("<div style='text-align:center;color:#7F8C8F;font-size:0.8rem;margin-top:1rem'>"
            "UTN FRSR · Gestión y Mantenimiento Electromecánico · TP N°1 · "
            "Herramienta de análisis interactivo</div>", unsafe_allow_html=True)