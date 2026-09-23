# -*- coding: utf-8 -*-
"""
==============================================================================
 TP N°1 - GESTIÓN Y MANTENIMIENTO ELECTROMECÁNICO (UTN FRSR)
 Línea Principal de Elevación y Transporte de Áridos de Alta Densidad
==============================================================================
 Aplicación interactiva (Streamlit + Plotly) para el análisis de confiabilidad,
 taxonomía ISO 14224, matriz de criticidad dinámica, dashboard y exportación.

 Uso:
     streamlit run app.py

 El bloque DATOS_* al inicio es EDITABLE: permite cargar/modificar los registros
 sin tocar la lógica del resto de la aplicación.
==============================================================================
"""

import io

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
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="baner">
  <h1>TP N°1 - GESTIÓN Y MANTENIMIENTO ELECTROMECÁNICO</h1>
  <p>UTN FRSR · Ingeniería Electromecánica · Línea de Elevación y Transporte de Áridos de Alta Densidad
  · Confiabilidad, Taxonomía ISO&nbsp;14224, Criticidad y Mantenibilidad</p>
</div>
""", unsafe_allow_html=True)

st.caption("Herramienta de Ingeniería de Confiabilidad y Mantenimiento - análisis del ciclo anual de 7.200 h")

# ==============================================================================
# 1) BLOQUE DE DATOS EDITABLES (registros y parámetros del enunciado del TP)
# ==============================================================================
# Parámetros por defecto (se editan en la barra lateral)
DICT_PARAM = {
    "TOP": 7200.0,            # Tiempo Operativo Programado [h/año]
    "costo_parada": 2500.0,   # US$/h de detención (lucro cesante + costo fijo)
    "costo_hh": 25.0,         # US$/hora-hombre de mantenimiento
    "umbral_A": 15,           # Clase A (Crítico)   : C >= 15
    "umbral_B": 10,           # Clase B (Semicrítico): 10 <= C < 15
}

# Registro de los 5 eventos de falla. Campos:
#   ttp, repuestos, hh  -> datos oficiales del enunciado
#   t0                  -> tiempo estimado de UNA ejecución de la tarea (h)
#   cuadrilla           -> n° medio de operarios en paralelo
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

# Escalas TF de criticidad (editable) - ver consigna 1b del enunciado.
# Ci = S + A + O + D ;  C = F x Ci
ESCALAS = {
    "F": "Frecuencia: 1 Rara (1/año) | 2 Ocasional (2/año) | 3 Frecuente (3-4/año) | 4 Muy frecuente (>4/año)",
    "S": "Seguridad: 1 Sin riesgo | 2 Lesión menor | 3 Accidente grave",
    "A": "Ambiente: 1 Sin impacto | 2 Impacto menor/contenido | 3 Contaminación severa de proceso",
    "O": "Operación: 1 Sin parada | 3 Parada parcial | 10 Parada total (petrificación de reactor)",
    "D": "Costo Directo de reparación: 1 < $1.000 | 2 $1.000-$3.000 | 4 > $3.000",
}

# Puntajes de consecuencia por defecto por subsistema (editables en Tab 3)
CRITICIDAD_DEF = [
    {"Subsistema": "Sistema Motriz",                    "S": 2, "A": 2, "O": 10, "D": 4},
    {"Subsistema": "Banda y Rodillos",                  "S": 2, "A": 2, "O": 10, "D": 4},
    {"Subsistema": "Estructura / Tolva",                "S": 2, "A": 1, "O": 3,  "D": 1},
    {"Subsistema": "Sistema de Tensión",                "S": 2, "A": 3, "O": 10, "D": 1},
    {"Subsistema": "Sistema Eléctrico / Seguridad",     "S": 3, "A": 1, "O": 3,  "D": 1},
]

# Taxonomía ISO 14224 (6 niveles: Industria→Planta→Área→Equipo→Sistema→Componente)
# con códigos y nivel padre. Los niveles intermedios normativos (Categoría de negocio
# e Instalación) se condensan para reflejar el arbolado pedido por la consigna.
TAXONOMIA = [
    # nivel, tipo, código, denominación, padre
    (1, "INDUSTRIA",   "IND-01",   "Industria Minera y de Procesamiento de Minerales", None),
    (2, "PLANTA",      "PLT-01",   "Planta de Alimentación Continua y Reacción", "IND-01"),
    (3, "ÁREA",        "ARE-01",   "Área de Elevación y Carga de Áridos al Reactor", "PLT-01"),
    (4, "EQUIPO",      "EQ-01",    "Línea de Transporte de Áridos (Cinta CT-01)", "ARE-01"),
    (5, "SISTEMA",     "SUB-01",   "Sistema Motriz (accionamiento y reducción)", "EQ-01"),
    (5, "SISTEMA",     "SUB-02",   "Banda y Rodillos (banda, tambores y rascadores)", "EQ-01"),
    (5, "SISTEMA",     "SUB-03",   "Estructura / Tolva (soporte y alimentación)", "EQ-01"),
    (5, "SISTEMA",     "SUB-04",   "Sistema de Tensión (oleohidráulico 20 bar)", "EQ-01"),
    (5, "SISTEMA",     "SUB-05",   "Sistema Eléctrico / Seguridad (sensores y protecciones)", "EQ-01"),
    (6, "COMPONENTE",  "COM-011",  "Motor asincrónico trifásico 55 kW", "SUB-01"),
    (6, "COMPONENTE",  "COM-012",  "Reductor ortogonal de ejes paralelos (3 etapas)", "SUB-01"),
    (6, "COMPONENTE",  "COM-013",  "Polea motriz Ø500 mm + eje SAE 4140 + rodamientos 22220 EK", "SUB-01"),
    (6, "COMPONENTE",  "COM-021",  "Banda transportadora de goma", "SUB-02"),
    (6, "COMPONENTE",  "COM-022",  "Tambor de reenvío y rodillos de apoyo", "SUB-02"),
    (6, "COMPONENTE",  "COM-023",  "Rascadores primario (uretano) y secundario (metal duro)", "SUB-02"),
    (6, "COMPONENTE",  "COM-031",  "Estructura metálica y caballetes", "SUB-03"),
    (6, "COMPONENTE",  "COM-032",  "Tolva de recepción de árido", "SUB-03"),
    (6, "COMPONENTE",  "COM-041",  "Centralita oleohidráulica (bomba de engranajes)", "SUB-04"),
    (6, "COMPONENTE",  "COM-042",  "Cilindro hidráulico de doble efecto y latiguillos HP", "SUB-04"),
    (6, "COMPONENTE",  "COM-043",  "Presostatos de seguridad", "SUB-04"),
    (6, "COMPONENTE",  "COM-051",  "Tablero eléctrico y arrancador suave", "SUB-05"),
    (6, "COMPONENTE",  "COM-052",  "Sensores inductivos de desalineación de banda", "SUB-05"),
    (6, "COMPONENTE",  "COM-053",  "Tirador de emergencia y circuitos de seguridad", "SUB-05"),
]

# Descripción de cada nivel de la taxonomía según norma ISO 14224 / consigna del TP.
NIVELES_ISO = {
    1: "Industria",
    2: "Planta",
    3: "Área",
    4: "Equipo",
    5: "Sistema",
    6: "Componente",
}

# ==============================================================================
# 1b) HELPERS DE TAXONOMÍA (árbol, niveles, listado directo y escalonado)
# ==============================================================================
def taxonomia_df():
    """DataFrame canónico de la taxonomía ISO 14224."""
    return pd.DataFrame(TAXONOMIA, columns=["Nivel", "Tipo", "Código", "Denominación", "Código padre"])


def _build_tree(jer):
    """Construye un árbol dirigido {código: [hijos]} a partir del campo padre."""
    hijos = {r["Código"]: [] for _, r in jer.iterrows()}
    raices = []
    for _, r in jer.iterrows():
        padre = r["Código padre"]
        if padre is None or pd.isna(padre) or padre == "":
            raices.append(r["Código"])
        elif padre in hijos:
            hijos[padre].append(r["Código"])
    return hijos, raices


def dataframe_directo():
    """Listado directo (tabla) con el orden lógico de campos solicitado."""
    jer = taxonomia_df()
    denom = jer.set_index("Código")["Denominación"].to_dict()

    def desc_padre(cod):
        return denom.get(cod, "—")

    df = jer.copy()
    df["Descripción del nivel"] = df["Nivel"].map(NIVELES_ISO)
    df["Descripción del equipo padre"] = df["Código padre"].fillna("").map(desc_padre)
    # Orden lógico: nivel -> descripción del nivel -> código -> descripción ->
    # tipo -> código padre -> descripción del padre
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
    """Listado escalonado (plano con columna de ruta completa para CSV)."""
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

# Fichas técnicas de componentes y función del equipo (Tab 1)
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

# Plan de ruta de inspecciones TPM (sensorial) y PdM (instrumental) - Tab 5 / Excel
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
# 2) FUNCIONES DE CÁLCULO (KPIs, estimación de repeticiones, criticidad)
# ==============================================================================
def factor_frecuencia(n):
    """Traduce el número estimado de fallas/año al factor F (escala del TP)."""
    if n >= 5:
        return 4
    if n >= 3:
        return 3
    if n == 2:
        return 2
    return 1


def clasificar(c, umbral_A, umbral_B):
    """Clasificación por Criticidad Total C = F x Ci con umbrales dinámicos."""
    if c >= umbral_A:
        return ("Clase A - CRÍTICO",
                "Alta Disponibilidad: monitoreo continuo + overhaul programado")
    if c >= umbral_B:
        return ("Clase B - SEMICRÍTICO",
                "Sistemático / Predictivo: re-lubricación y mediciones por calendario/condición")
    return ("Clase C - NO CRÍTICO",
            "Condicional / TPM: inspección CIL por operador y atención a la falla")


def estimar_repeticiones(df_ed):
    """Estima cuántas veces se repitió cada falla en el ciclo anual.

    N_ttp = redondeo(TTP / t0)      -> ambos caminos (temporal)
    N_hh  = redondeo(HH / (t0*c))   -> consumo de horas-hombre
    N_est = max(N_ttp, N_hh)        -> conciliación conservadora;
                                       si difieren se marca para revisión.
    """
    df = df_ed.copy()
    df["HH_oc"] = df["t0"] * df["cuadrilla"]              # HH por ocurrencia
    df["N_ttp"] = np.floor(df["ttp"] / df["t0"]).clip(lower=1).astype(int)
    df["N_hh"]  = np.floor(df["hh"] / df["HH_oc"]).clip(lower=1).astype(int)
    df["N_est"] = [max(a, b) for a, b in zip(df["N_ttp"], df["N_hh"])]
    df["Revisar"] = df["N_ttp"] != df["N_hh"]
    return df


LETRAS_OCURRENCIA = "abcdefghijklmnopqrstuvwxyz"


def _aplicar_ediciones_estim(df_estim):
    """Aplica las ediciones de t0/cuadrilla hechas en el editor de la Pestaña 3.

    El valor vigente de un st.data_editor vive en st.session_state[key] como
    {"edited_rows": {fila: {columna: valor}}}; se usa para que el desglose por
    repeticiones de la Pestaña 2 respete lo que el usuario ajustó en t0/cuadrilla.
    """
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


def generar_fallas_supuesto(con_rep, df_estim):
    """Base canónica de fallas según el supuesto vigente.

    sin repetición -> las 5 filas originales del TP (F1..F5).
    con repetición -> cada F_i se desglosa en N_est[i] ocurrencias (F1a, F1b, …)
                      repartiendo TTP / Repuestos / HH en partes iguales.
    Los totales anuales se conservan: TTP = 120 h, Repuestos = $10.220, HH = 136.
    """
    base = pd.DataFrame(DATOS_FALLAS)
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


# Si el usuario pidió restaurar el registro original, se limpian acá los estados
# de los widgets ANTES de instanciarlos (setear/borrar un widget ya instanciado
# genera StreamlitAPIException). El flag lo levanta el botón de la Pestaña 2.
if st.session_state.pop("reiniciar_registro", False):
    for k in ("editor_fallas", "editor_estim", "usar_n_est", "sup_aplicado__fallas"):
        st.session_state.pop(k, None)
    st.session_state["fallas_custom"] = []


# ==============================================================================
# 3) BARRA LATERAL - CRITERIOS Y PARÁMETROS DINÁMICOS
# ==============================================================================
with st.sidebar:
    st.header("Parámetros dinámicos")
    st.caption("Todos los resultados se recalculan en vivo.")

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
    st.subheader("Umbrales de Criticidad Total")
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
        help="Desglosa cada falla registrada en sus ocurrencias repetidas según la "
             "estimación por tareas: F1→2, F2→1, F3→2, F4→2 y F5→2 (9 filas). "
             "La tabla se regenera conservando las filas que hayas agregado; con "
             "'Restaurar datos originales' vuelve a los 5 registros del TP.",
    )

    st.divider()
    with st.expander("Escalas del TP (referencia)"):
        for k, v in ESCALAS.items():
            st.markdown(f"**{k}.** {v}")

    st.divider()
    with st.expander("👤 Datos del Alumno (Portada del Informe)", expanded=False):
        alumno_nombre = st.text_input("Nombre y Apellido", value="Martín Méndez", key="alumno_nombre")
        alumno_legajo = st.text_input("Legajo / Matrícula", value="UTN FRSR", key="alumno_legajo")
        ciclo_lectivo = st.text_input("Ciclo Lectivo", value="2026", key="ciclo_lectivo")

# ==============================================================================
# 4) DATAFRAMES BASE Y ESTIMACIÓN DE REPETICIONES
# ==============================================================================
# Base canónica del TP (F1..F5) con tarea/descripción/t0/cuadrilla. La tabla
# operativa real (con ediciones del usuario) se define dentro de la Pestaña 2.
df_fallas_base = pd.DataFrame(DATOS_FALLAS)

# Estimación de repeticiones sobre eventos canónicos: usa t0/cuadrilla del TP
# más las ediciones vigentes (t0/cuadrilla) del editor de la Pestaña 3.
df_estim = df_fallas_base[["evento", "subsistema", "tarea", "ttp", "hh", "repuestos",
                           "t0", "cuadrilla"]].copy()
df_estim = _aplicar_ediciones_estim(df_estim)
df_estim = estimar_repeticiones(df_estim)
N_est_total = int(df_estim["N_est"].sum())

# Factor F sugerido por subsistema (a partir de las repeticiones estimadas)
N_por_sub = df_estim.groupby("subsistema")["N_est"].sum()
F_sugerido = N_por_sub.apply(factor_frecuencia)

# KPIs: se calculan dentro de la Pestaña 2, en vivo, sobre la tabla editada
# del usuario (df_fallas) con N = cantidad de filas vigentes.

# ==============================================================================
# 5) PESTAÑAS
# ==============================================================================
tab_taxo, tab_kpi, tab_matriz, tab_dash, tab_excel, tab_informe = st.tabs(
    ["1 · Taxonomía del Activo (ISO 14224)",
     "2 · Registro de Fallas y KPIs",
     "3 · Matriz de Criticidad Dinámica",
     "4 · Dashboard (Plotly)",
     "5 · Exportación a Excel",
     "6 · Informe Técnico Oficial (.docx)"])

# -------------------------------------------------------------------------------
# PESTAÑA 1 - TAXONOMÍA DEL ACTIVO (NIVELES ISO 14224)
# -------------------------------------------------------------------------------
with tab_taxo:
    st.subheader("Jerarquía del activo - Taxonomía ISO 14224")
    st.caption("Industria → Planta → Área → Equipo → Sistema → Componente. Cada "
               "nivel explícita su nivel padre y genera la trazabilidad de las "
               "decisiones (reemplazo vs. reparación) según la consigna del TP.")

    # Selector de modo de visualización
    modo = st.radio(
        "Modo de visualización",
        ["Escalonado (árbol)", "Listado directo (tabla)"],
        horizontal=True,
        help="Escalonado: jerarquía desplegable con sus hijos anidados. "
             "Listado directo: tabla plana con el detalle de cada equipo.",
    )

    # --- Utilidades para descargas CSV -------------------------------
    def to_csv_bytes(df):
        return df.to_csv(index=False).encode("utf-8-sig")

    if modo == "Escalonado (árbol)":
        # ------------------------------------------------------------------
        # MODO ESCALONADO: árbol jerárquico con st.expander anidados
        # ------------------------------------------------------------------
        st.markdown("Despliegue cada jerarquía para ver sus hijos. Puede ocultar "
                    "(colapsar) cualquier nivel de forma independiente.")

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
        with col_n:
            st.caption("El CSV incluye la columna **Ruta completa** (Industria > "
                       "Planta > …) para reconstruir la jerarquía en Excel u otras "
                       "herramientas.")

    else:
        # ------------------------------------------------------------------
        # MODO LISTADO DIRECTO: tabla plana con detalle de cada equipo
        # ------------------------------------------------------------------
        df_directo = dataframe_directo()
        st.dataframe(df_directo, width="stretch", hide_index=True)

        col_csv, col_n = st.columns([1, 3])
        with col_csv:
            st.download_button(
                "⬇  Descargar listado directo (CSV)",
                data=to_csv_bytes(df_directo),
                file_name="taxonomia_listado_directo.csv",
                mime="text/csv",
            )
        with col_n:
            st.caption("Campos ordenados: nivel, descripción del nivel, código del "
                       "equipo, descripción, tipo, equipo padre y descripción del padre.")

    st.divider()

    # Fichas técnicas de los subsistemas
    st.markdown("### Fichas técnicas de los subsistemas")
    for subsistema, componentes in FICHAS.items():
        with st.expander(f"**{subsistema}** - {len(componentes)} componente(s)") as exp:
            for nombre, funcion, espec, falla in componentes:
                c1, c2, c3 = st.columns([1, 1, 1])
                c1.markdown(f"**{nombre}**")
                c2.markdown(f"**Función:** {funcion}")
                c3.markdown(f"**Especificaciones:** {espec}")
                st.caption(f"Falla asociada / observación: {falla}")
                st.divider()

# -------------------------------------------------------------------------------
# PESTAÑA 2 - REGISTRO EDITABLE DE FALLAS Y KPIs DE CONFIABILIDAD
# -------------------------------------------------------------------------------
with tab_kpi:
    est = st.session_state
    if "sup_aplicado__fallas" not in est:
        est["sup_aplicado__fallas"] = usar_n_est
    if "fallas_custom" not in est:
        est["fallas_custom"] = []

    # Al cambiar el supuesto se regenera la base canónica, conservando las filas
    # que el usuario hubiera agregado (se vuelcan al editar la tabla).
    if est["sup_aplicado__fallas"] != usar_n_est:
        est["sup_aplicado__fallas"] = usar_n_est
        if "editor_fallas" in est:
            del est["editor_fallas"]

    st.subheader("Registro editable de fallas del ciclo anual")
    col_tit, col_reset = st.columns([3, 1])
    with col_tit:
        st.caption("Agregue, borre o edite filas (Evento, Subsistema, Componente, "
                   "TTP, Repuestos, HH). Los costos derivados y los KPIs se "
                   "recalculan en vivo; la fila TOTAL queda bajo la tabla.")
    with col_reset:
        if st.button("Restaurar datos originales del TP (estado inicial)",
                     help="Vuelve a los 5 registros oficiales (F1-F5), sin repetición "
                          "y sin filas agregadas."):
            st.session_state["reiniciar_registro"] = True
            st.rerun()

    base_fallas = generar_fallas_supuesto(usar_n_est, df_estim)
    etiquetas_canon = set(base_fallas["evento"])
    df_pasado = pd.concat([base_fallas, pd.DataFrame(est["fallas_custom"])],
                          ignore_index=True)

    col_editor = {
        "evento": st.column_config.TextColumn("Evento", width="small",
                                              help="Código de la falla (F1..F5 o nueva)"),
        "subsistema": st.column_config.TextColumn("Subsistema"),
        "componente": st.column_config.TextColumn("Componente", width="large"),
        "ttp": st.column_config.NumberColumn("TTP (h)", min_value=0.0, step=0.5,
                                             format="%.1f"),
        "repuestos": st.column_config.NumberColumn("Repuestos (USD)", min_value=0.0,
                                                   step=50.0, format="%.0f"),
        "hh": st.column_config.NumberColumn("HH", min_value=0.0, step=1.0,
                                            format="%.0f"),
    }
    df_edit = st.data_editor(
        df_pasado,
        column_config=col_editor,
        num_rows="dynamic",
        column_order=["evento", "subsistema", "componente", "ttp", "repuestos", "hh"],
        key="editor_fallas", width="stretch", hide_index=True)

    # Custodiar las filas agregadas por el usuario para conservarlas al cambiar
    # de supuesto o regenerar la base.
    df_custom = df_edit[~df_edit["evento"].isin(etiquetas_canon)]
    est["fallas_custom"] = df_custom.to_dict("records")

    # --- Recalcular costos derivados y KPIs sobre la tabla editada -----------
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
    Ai = MTBF / (MTBF + MTTR) * 100
    Ao = TFR / TOP_anio * 100

    cols_tabla = ["evento", "subsistema", "ttp", "Costo_Parada_USD", "repuestos",
                  "hh", "Costo_MO_USD", "Costo_Total_USD"]
    tabla = df_fallas[cols_tabla].copy()
    tabla.columns = ["Evento", "Subsistema", "TTP (h)", "Costo Parada (USD)",
                     "Repuestos (USD)", "HH", "Costo MO (USD)", "Costo Total (USD)"]
    total_row = pd.DataFrame([["TOTAL", "—", TTP_total,
                               df_fallas["Costo_Parada_USD"].sum(),
                               df_fallas["repuestos"].sum(), df_fallas["hh"].sum(),
                               df_fallas["Costo_MO_USD"].sum(),
                               df_fallas["Costo_Total_USD"].sum()]],
                             columns=tabla.columns)
    tabla = pd.concat([tabla, total_row], ignore_index=True)
    st.dataframe(tabla.style.format({
        "TTP (h)": "{:,.1f}",
        "Costo Parada (USD)": lambda x: f"${x:,.0f}",
        "Repuestos (USD)": lambda x: f"${x:,.0f}",
        "HH": "{:,.0f}",
        "Costo MO (USD)": lambda x: f"${x:,.0f}",
        "Costo Total (USD)": lambda x: f"${x:,.0f}"}),
        width="stretch", hide_index=True)

    if usar_n_est:
        st.info(f"Supuesto **con repetición**: cada falla se desglosó según N_est "
                f"(estimación por tareas, {N_est_total} fallas/año). TTP total y TFR "
                "se mantienen iguales (120 h y TFR derivado).")
    else:
        st.caption("Supuesto **sin repetición**: registros oficiales del TP (5 fallas).")

    st.markdown("### Indicadores de Confiabilidad y Disponibilidad")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Tiempo Operativo Programado (TOP)", f"{TOP_anio:,.0f} h",
              help="TOP = tiempo programado de operación anual")
    k2.metric("Tiempo Total de Paradas (TTP)", f"{TTP_total:,.0f} h",
              help=f"Σ TTP de las {N_activo} filas vigentes")
    k3.metric("Funcionamiento Real (TFR)", f"{TFR:,.0f} h",
              help="TFR = TOP − TTP")
    k4.metric("N° de Fallas", f"{N_activo}",
              help="N = cantidad de filas vigentes de la tabla"
                   + (" · hipótesis de repetición por tareas" if usar_n_est else
                      " · registradas oficialmente"))

    k5, k6, k7, k8 = st.columns(4)
    k5.metric("MTBF", f"{MTBF:,.2f} h",
              help=f"MTBF = TFR / N = {TFR:,.0f} / {N_activo}")
    k6.metric("MTTR", f"{MTTR:,.2f} h",
              help=f"MTTR = TTP / N = {TTP_total:,.0f} / {N_activo}")
    k7.metric("Disponibilidad Inherente (Ai)", f"{Ai:.2f} %",
              help="Ai = MTBF / (MTBF + MTTR) · condiciones logísticas ideales")
    k8.metric("Disponibilidad Operacional (Ao)", f"{Ao:.2f} %",
              help="Ao = TFR / TOP · gestión real del activo")

    st.caption("Nota: Ai = Ao cuando no hubo paradas por mantenimiento preventivo "
               "(todo el tiempo fuera de servicio correspondió a reparación).")
    if usar_n_est:
        st.caption(f"Referencia del supuesto: MTBF ≈ {TFR/N_est_total:,.1f} h "
                   f"y MTTR ≈ {TTP_total/N_est_total:,.1f} h/falla "
                   "(la estimación no altera Ai/Ao porque ambos colapsan a TFR/TOP).")

# -------------------------------------------------------------------------------
# PESTAÑA 3 - MATRIZ DE CRITICIDAD DINÁMICA
# -------------------------------------------------------------------------------
with tab_matriz:
    st.subheader("Paso 1 · Estimación de repeticiones por tipo de tarea")
    st.caption("Se supone que el registro anual (TTP y HH) agrupa la REPETICIÓN de "
               "fallas similares. Edite t0 (horas de una ejecución de la tarea) y la "
               "cuadrilla para ver cómo cambia la frecuencia estimada.")

    columnas_edit = {
        "evento": st.column_config.TextColumn("Evento", disabled=True, width="small"),
        "subsistema": st.column_config.TextColumn("Subsistema", disabled=True),
        "tarea": st.column_config.TextColumn("Tarea de reparación (tipo)", disabled=True, width="large"),
        "ttp": st.column_config.NumberColumn("TTP total (h)", disabled=True),
        "hh": st.column_config.NumberColumn("HH total", disabled=True),
        "repuestos": st.column_config.NumberColumn("Repuestos (USD)", disabled=True, format="%.0f"),
        "t0": st.column_config.NumberColumn("t0 (h/tarea)", min_value=0.5, max_value=200.0, step=0.5, format="%.1f"),
        "cuadrilla": st.column_config.NumberColumn("Cuadrilla", min_value=1, max_value=8, step=1, format="%d"),
        "N_ttp": st.column_config.NumberColumn("N por TTP", disabled=True),
        "N_hh": st.column_config.NumberColumn("N por HH", disabled=True),
        "N_est": st.column_config.NumberColumn("N est. anual", disabled=True),
        "Revisar": st.column_config.CheckboxColumn("Revisar", disabled=True, help="Diferencias entre ambas lecturas"),
    }
    df_estim_edit = st.data_editor(
        df_estim,
        column_config=columnas_edit,
        disabled=["evento", "subsistema", "tarea", "ttp", "hh", "repuestos",
                  "N_ttp", "N_hh", "N_est", "Revisar"],
        width="stretch", hide_index=True, key="editor_estim")

    df_estim = estimar_repeticiones(df_estim_edit)          # recalcular con lo editado
    N_est_total = int(df_estim["N_est"].sum())
    N_por_sub = df_estim.groupby("subsistema")["N_est"].sum()
    F_sugerido = N_por_sub.apply(factor_frecuencia)

    c1, c2 = st.columns(2)
    c1.metric("Total de repeticiones estimadas por año (N_est)", f"{N_est_total} fallas",
              help="F1+F2+F3+F4+F5 estimadas según t0 y cuadrilla")
    c2.caption("Reasignación automática de F en la matriz de abajo.")

    st.divider()
    st.subheader("Paso 2 · Matriz de Criticidad dinámica (C = F × Ci)")
    st.caption("Edite libremente F, S, A, O y D. Ci, C, Clase y Modelo se recalculan "
               "en vivo con los umbrales de la barra lateral.")

    matriz0 = pd.DataFrame(CRITICIDAD_DEF).copy()
    matriz0["F"] = matriz0["Subsistema"].map(F_sugerido).fillna(1).astype(int)
    matriz0["N_est/año"] = matriz0["Subsistema"].map(N_por_sub).fillna(0).astype(int)

    col_mat = {
        "Subsistema": st.column_config.TextColumn("Subsistema", disabled=True),
        "F": st.column_config.SelectboxColumn("F", options=[1, 2, 3, 4],
                                              help=ESCALAS["F"]),
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

    st.markdown("### Gráfico de críticidad: Consecuencia (Ci) vs Frecuencia (F)")
    fig_crit = go.Figure()
    for cls in df_matriz["Clase"].unique():
        sub = df_matriz[df_matriz["Clase"] == cls]
        fig_crit.add_trace(go.Scatter(
            x=sub["Ci"], y=sub["F"], mode="markers+text",
            name=cls, text=sub["Subsistema"], textposition="top center",
            marker=dict(size=sub["C"].clip(lower=8) * 2.2, color=mapa_clase[cls],
                        opacity=0.85, line=dict(color="white", width=1.5)),
            hovertemplate=("<b>%{text}</b><br>Ci=%{x:.0f}<br>F=%{y}<br>"
                           "C=%{customdata:.0f}<extra></extra>"),
            customdata=sub["C"]))
    ci_range = np.linspace(1, 26, 120)
    for umb, color in ((umbral_A, ROJO), (umbral_B, NARANJ)):
        fig_crit.add_trace(go.Scatter(
            x=ci_range, y=np.clip(umb / ci_range, 0, 17), mode="lines",
            line=dict(color=color, dash="dash", width=1.2),
            name=f"C = {umb}", showlegend=True, hovertemplate=f"isotónica C={umb}<br>Ci=%{{x:.0f}}"))
    fig_crit.update_layout(
        title="Matriz de Criticidad - Frecuencia (F) vs Consecuencia (Ci = S+A+O+D)",
        xaxis=dict(title="Consecuencia Ci", range=[0, df_matriz["Ci"].max() * 1.15]),
        yaxis=dict(title="Frecuencia de falla (F)", dtick=1, range=[0, max(4.5, df_matriz["F"].max() + 0.5)]),
        height=480, template="plotly_white",
        legend=dict(orientation="h", y=-0.18))
    st.plotly_chart(fig_crit, width="stretch")

# -------------------------------------------------------------------------------
# PESTAÑA 4 - DASHBOARD DE GRÁFICOS (PLOTLY)
# -------------------------------------------------------------------------------
with tab_dash:
    # G1 - Pareto de costos por evento
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

    # G2 - MTTR por evento
    dfm = df_fallas.sort_values("ttp")
    fig2 = go.Figure(go.Bar(
        x=dfm["evento"], y=dfm["ttp"],
        marker_color=[NARANJ if i % 2 else AZUL for i in range(len(dfm))],
        text=dfm["ttp"].apply(lambda v: f"{v:.1f} h"), textposition="outside",
        hovertemplate="%{x} · MTTR %{y:.1f} h<extra></extra>"))
    fig2.update_layout(title="G2 · Tiempo de Reparación (MTTR) por evento",
                       xaxis_title="Evento de falla", yaxis_title="Tiempo [h]",
                       template="plotly_white", height=420)

    # G3 - Matriz de criticidad (reutiliza la figura de Pestaña 3)
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

    # G4 - Donut de estructura de costos
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

    st.divider()
    st.markdown("### G5 · Distribución Temporal de Fallas y Patrones de Nowlan & Heap (Consigna 4.b)")
    st.caption("Cronología anual de los eventos de falla registrados sobre las 7.200 h programadas. "
               "Permite contrastar si el activo responde al patrón clásico de la bañera (Patrón A) o a fallas aleatorias (Patrones D, E o F).")

    df_temp = pd.DataFrame({
        "evento": ["F1", "F2", "F3", "F4", "F5"],
        "hora_anual": [1100, 2750, 4100, 5400, 6600],
        "componente": ["Rodamiento oscilante SKF 22220 EK", "Motor eléctrico 55 kW",
                       "Latiguillo alta presión", "Banda de transporte (goma)", "Rascador primario uretano"],
        "ttp": [28, 42, 6, 36, 8],
        "costo_total": [73050, 109000, 15620, 95100, 20850],
        "patron": ["Patrón E (Falla aleatoria por polvo abrasivo)",
                   "Patrón E (Bloqueo térmico accidental por polvo)",
                   "Patrón D (Fatiga por pulsación hidráulica)",
                   "Patrón F (Desalineación y atasco mecánico)",
                   "Patrón E (Desgaste prematuro abrasivo)"]
    })

    fig5 = go.Figure()
    fig5.add_trace(go.Scatter(
        x=df_temp["hora_anual"], y=df_temp["ttp"],
        mode="markers+text",
        text=df_temp["evento"],
        textposition="top center",
        marker=dict(size=[24, 32, 16, 28, 18], color=[ROJO, ROJO, NARANJ, ROJO, VERDE],
                    line=dict(color="white", width=2), opacity=0.9),
        customdata=np.stack((df_temp["componente"], df_temp["costo_total"], df_temp["patron"]), axis=-1),
        hovertemplate="<b>%{text} · %{customdata[0]}</b><br>"
                      "Hora Operativa: %{x:,.0f} h<br>TTP: %{y:.0f} h<br>"
                      "Costo Total: $%{customdata[1]:,.0f} USD<br>"
                      "Clasificación: %{customdata[2]}<extra></extra>"
    ))

    fig5.add_hline(y=df_temp["ttp"].mean(), line_dash="dash", line_color=AZUL2,
                   annotation_text=f"TTP Promedio: {df_temp['ttp'].mean():.1f} h (Tasa aleatoria ~ constante)",
                   annotation_position="bottom right")

    fig5.update_layout(
        title="Distribución Temporal de Contingencias en el Ciclo Operativo (TOP = 7.200 h)",
        xaxis=dict(title="Horas de Operación Acumuladas en el Año [h]", range=[0, 7500], dtick=1000),
        yaxis=dict(title="Tiempo de Parada TTP [h]", range=[0, 50]),
        template="plotly_white", height=380
    )
    st.plotly_chart(fig5, width="stretch")

    st.info(
        "💡 **Conclusión Técnica (Consigna 4.b):** Este activo **NO** sigue la tradicional 'curva de la bañera' (Patrón A), "
        "la cual solo representa ~4% de los activos industriales complejos. En cambio, responde a los **Patrones D, E y F "
        "(fallas aleatorias y prematuras condicionales al entorno)** debido a la contaminación abrasiva, polvo y vibración. "
        "Por lo tanto, la sustitución rígida por horas de servicio o calendario es ineficaz y contraproducente, debiéndose "
        "priorizar el Mantenimiento Basado en la Condición (CBM) y el Mantenimiento Autónomo (TPM)."
    )

# -------------------------------------------------------------------------------
# PESTAÑA 5 - EXPORTACIÓN PROFESIONAL A EXCEL (openpyxl)
# -------------------------------------------------------------------------------
def generar_excel(TOP, cost_par, cost_hh, uA, uB, usar_est,
                  df_fallas_, df_matriz_, df_estim_, n_est, mbf, mtr, ai, ao):
    """Genera el informe Excel TP1_Gestion_Mantenimiento_UTN.xlsx en memoria."""
    AZULX, GRISX, AMBARX, VERDEX, ROJOX = "1F4E79", "F2F2F2", "FFF2CC", "E2EFDA", "FBE5E5"
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

    # --- Hoja 1: KPIs ---
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
    if usar_est:
        ws.cell(fila + len(kpis), 1, "KPIs calculados con N_est (hipótesis de repetición de fallas).").font = \
            Font(italic=True, size=9, color="7F7F7F")

    # --- Hoja 2: Registro de fallas / finanzas ---
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

    # --- Hoja 3: Matriz de criticidad ---
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

    # --- Hoja 4: Plan de inspecciones ---
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

    st.caption("Los valores del Excel corresponden a la configuración actual de la "
               "barra lateral (parámetros, umbrales y modo N_est).")

# -------------------------------------------------------------------------------
# PESTAÑA 6 - INFORME TÉCNICO OFICIAL DE CÁTEDRA (.DOCX)
# -------------------------------------------------------------------------------
with tab_informe:
    st.subheader("Informe Técnico Oficial de Cátedra — Justificación Integral")
    st.caption("Documento formal que responde y justifica las 6 consignas del enunciado oficial de UTN FRSR "
               "con rigor de ingeniería electromecánica, sincronizado en vivo con los parámetros y supuestos de la aplicación.")

    col_inf1, col_inf2, col_inf3 = st.columns(3)
    col_inf1.metric("Alumno / Responsable", alumno_nombre, help=f"Legajo: {alumno_legajo}")
    col_inf2.metric("Ciclo / Régimen", f"{ciclo_lectivo} · {TOP_anio:,.0f} h/año",
                    help=f"Costo Parada: ${costo_par:,.0f} USD/h | MO: ${costo_hh:,.0f} USD/HH")
    supuesto_str = f"Con Repetición (N_est ≈ {N_est_total})" if usar_n_est else f"Nominal (N = {N_activo})"
    col_inf3.metric("Supuesto Activo", supuesto_str,
                    help="Determina el desglose de fallas y MTBF en el informe.")

    # Generación en memoria del archivo DOCX
    docx_buffer = build_technical_report(
        alumno=alumno_nombre,
        legajo=alumno_legajo,
        anio=ciclo_lectivo,
        TOP=TOP_anio,
        cost_parada=costo_par,
        cost_hh=costo_hh,
        df_fallas=df_fallas,
        kpis_nominal={"ttp": 120.0, "tfr": TOP_anio - 120.0, "mtbf": (TOP_anio - 120.0) / 5.0,
                      "mttr": 120.0 / 5.0, "ai": ((TOP_anio - 120.0) / TOP_anio) * 100,
                      "ao": ((TOP_anio - 120.0) / TOP_anio) * 100, "n": 5},
        kpis_rep={"ttp": 120.0, "tfr": TOP_anio - 120.0, "mtbf": (TOP_anio - 120.0) / N_est_total,
                  "mttr": 120.0 / N_est_total, "ai": ((TOP_anio - 120.0) / TOP_anio) * 100,
                  "ao": ((TOP_anio - 120.0) / TOP_anio) * 100, "n": N_est_total},
        df_matriz=df_matriz
    )

    st.download_button(
        label="📥 Descargar Informe Técnico Oficial (.docx)",
        data=docx_buffer.getvalue(),
        file_name="TP1_Informe_Tecnico_Electromecanico_UTN.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        help="Descarga el documento Word (.docx) formateado con estilos institucionales UTN FRSR y todas las respuestas completas."
    )

    st.markdown("---")
    st.markdown("### 📖 Vista previa interactiva de las consignas respondidas")

    with st.expander("Consigna 1: Jerarquía de Activos (ISO 14224) y Modelos de Mantenimiento", expanded=True):
        st.markdown("""
        **a) Arbolado Jerárquico y Decisiones de Reemplazo (MTTF) vs. Reparación (MTBF):**
        - **Nivel 2 (Complejo Industrial):** Planta de Procesamiento de Minerales / Áridos.
        - **Nivel 3 (Área Operativa):** Sección de Elevación, Transporte y Molienda Primaria.
        - **Nivel 4 (Equipo):** Cinta Transportadora Principal (CV-01).
        - **Nivel 5 (Sistemas):** Unidad Motriz, Banda y Rodillos, Sistema de Tensión, Estructura/Chasis, Tablero y Control.
        - **Nivel 6 (Componentes):** Rodamientos SKF 22220 EK, motor 55 kW, reductor i=28:1, latiguillos hidráulicos, rascador de uretano.
        - **Nivel 8 (Elementos):** Pistas de rodadura, retenes de grasa, sellos O-ring, espiras de cobre.

        *Criterio Técnico:*
        - **Reemplazo (Nivel 6 y 8):** Componentes unitarios indivisibles descartables (rodamientos, sellos, mangueras). Métrica: **MTTF**.
        - **Reparación (Nivel 4 y 5):** Sistemas y equipos restaurables (motor eléctrico, reductor ortogonal, tambores). Métrica: **MTBF**.

        **b) y c) Matriz de Criticidad y Modelos Asignados:**
        - **Unidad Motriz (Clase A — Crítico, C=51):** Modelo de Alta Disponibilidad (monitoreo continuo de vibraciones y corriente, overhaul programado).
        - **Banda y Rodillos (Clase A — Crítico, C=51):** Modelo de Alta Disponibilidad (inspección de empalmes, termografía en rodillos y alineación continua).
        - **Sistema de Tensión (Clase B — Semicrítico, C=16):** Modelo Sistemático / Predictivo (análisis de aceite hidráulico, control de presión e inspección de latiguillos).
        - **Estructura y Chasis (Clase B/C):** Modelo Condicional / TPM (rutina CIL por operador y remoción de áridos).
        - **Tablero y Control (Clase C — No Crítico, C=7):** Modelo Condicional / TPM (termografía semestral y limpieza de filtros).
        """)

    with st.expander("Consigna 2: Plan de Tareas de Inspección (Sensoriales e Instrumentales)"):
        st.markdown("""
        Se establece la Ruta de Inspección Rutinaria del activo:
        1. **Inspecciones Sensoriales (Operador / TPM en marcha):**
           - *Rascadores de uretano:* Control visual de filo, presión y acumulación de finos (Frecuencia: Diaria por turno).
           - *Tambores y poleas:* Detección auditiva de rozamiento y control visual de material pegado (Frecuencia: Diaria).
           - *Centralita oleohidráulica:* Control visual de nivel, fugas en racores y pulsaciones anormales (Frecuencia: Diaria).
        2. **Inspecciones Instrumentales (Mantenimiento / Predictivo):**
           - *Rodamientos SKF 22220 EK:* Análisis de vibraciones (demodulación de aceleración gE) y ultrasonido (Frecuencia: Quincenal).
           - *Motor 55 kW:* Termografía infrarroja de carcasa/bornera y medición de corriente con pinza amperimétrica (Frecuencia: Mensual).
           - *Reductor ortogonal:* Análisis de laboratorio de lubricante sintético ISO VG 220 (viscosidad, humedad ppm, conteo de partículas) (Frecuencia: Trimestral).
        """)

    with st.expander("Consigna 3: Análisis de Tipos de Mantenimiento Proactivo (Tabla 3 del TP)"):
        st.markdown("""
        Cuadro de acciones proactivas para prevenir las 5 contingencias históricas:
        - **F1 (Rodamiento oscilante SKF 22220 EK):** Mantenimiento Predictivo (CBM) mediante análisis de vibraciones espectral y lubricación asistida por ultrasonido (Quincenal).
        - **F2 (Motor eléctrico 55 kW):** Mantenimiento Autónomo (TPM) de limpieza de rejilla de ventilador + Termografía y ensayo de resistencia de aislación (Megado mensual).
        - **F3 (Latiguillo de alta presión):** Mantenimiento Preventivo Sistemático con sustitución por horas de servicio / calendario (Vida útil máxima 2 años).
        - **F4 (Banda transportadora):** Mantenimiento Autónomo (TPM) con rascador interno en tambor de reenvío + Interbloqueo seguro por sensores inductivos de desalineación (Diario).
        - **F5 (Rascador de uretano):** Mantenimiento Autónomo (TPM) diario con verificación visual y ajuste de tensión en contrapeso (Diario por turno).
        """)

    with st.expander("Consigna 4: Determinación de KPIs de Confiabilidad y Patrones de Nowlan & Heap"):
        st.markdown(f"""
        **a) Memoria de Cálculo de KPIs:**
        - **TFR:** {TOP_anio:,.0f} h - {TTP_total:.1f} h = **{TFR:,.1f} h/año**.
        - **Caso Base Nominal (N = {N_activo}):**
          - $MTBF = {TFR:,.1f} / {N_activo} =$ **{MTBF:,.2f} h**.
          - $MTTR = {TTP_total:.1f} / {N_activo} =$ **{MTTR:,.2f} h**.
          - $A_i = MTBF / (MTBF + MTTR) =$ **{Ai:.2f}%**.
          - $A_o = TFR / TOP =$ **{Ao:.2f}%**.
        - **Caso Refinado con Repetición (N_est ≈ {N_est_total}):**
          - $MTBF = {TFR:,.1f} / {N_est_total} =$ **{(TFR/N_est_total):,.2f} h**.
          - $MTTR = {TTP_total:.1f} / {N_est_total} =$ **{(TTP_total/N_est_total):,.2f} h**.
          - Disponibilidades idénticas ($A_i = A_o = {Ao:.2f}\\%$).

        **b) Patrones de Falla de Nowlan & Heap:**
        - El activo **no responde al Patrón A (curva de la bañera)**, ya que sus fallas no obedecen a un desgaste uniforme por edad.
        - Responde a los **Patrones D, E y F (aleatoriedad condicional al entorno)**: la abrasión por polvo, el bloqueo por suciedad y la fatiga por pulsación ocurren de manera aleatoria. Las políticas de overhaul calendario son ineficientes frente al monitoreo de condición (PdM/TPM).
        """)

    with st.expander("Consigna 5: Evaluación de Impacto Financiero y Motor Standby (Tabla 4 del TP)"):
        st.markdown(f"""
        **a) Balance Económico Consolidado:**
        - **Costo de Parada:** ${df_fallas["Costo_Parada_USD"].sum():,.0f} USD (95,66% del total).
        - **Repuestos e Insumos:** ${df_fallas["repuestos"].sum():,.0f} USD (3,26%).
        - **Mano de Obra Propia:** ${df_fallas["Costo_MO_USD"].sum():,.0f} USD (1,08%).
        - **Costo Total Directo:** **${df_fallas["Costo_Total_USD"].sum():,.0f} USD**.

        **b) Falla Crítica (F2) y Propuesta de Motor Standby:**
        - **Falla F2 (Motor 55 kW):** Causó el mayor impacto económico (**$109.000 USD**, 34,75% de las pérdidas).
        - **Estrategia de Mitigación:** Almacenar un **motor Standby (55 kW, IP55, 4 polos)** en el pañol.
        - **Impacto:** Reduce el $MTTR$ de **42 h** (retraso por rebobinado en taller externo) a solo **4 h** (reemplazo mecánico directo y alineación).
        - **Ahorro Neto:** **$95.000 USD** en un solo evento, amortizando el costo del motor (~$3.500 USD) más de 25 veces.
        """)

    with st.expander("Consigna 6: Hipótesis Diagnóstica de Falla Combinada (Efecto Dominó en F5)"):
        st.markdown("""
        **Cadena Causa-Efecto Electromecánica:**
        1. Desgaste no uniforme y fisura del rascador de uretano por contacto abrasivo continuo.
        2. Pérdida de perfil y trabamiento mecánico de la lámina contra la banda.
        3. Incremento brusco de la fuerza de fricción ($F_r = \\mu \\cdot F_n \\uparrow$), aumentando la cupla resistente ($T_{res} \\uparrow$) sobre la polea motriz.
        4. El motor asincrónico incrementa su deslizamiento ($s \\uparrow$) y, en consecuencia, eleva su corriente estatórica ($I_1 \\propto T_{res}$), generando calentamiento excesivo por efecto Joule ($P = 3 I^2 R$).
        5. El reductor ortogonal trabaja en sobrecarga mecánica continua, transmitiendo calor al baño de aceite sintético ISO VG 220 (> 95 °C). A esa temperatura, la viscosidad cae drásticamente, rompiendo la película lubricante EHL y acelerando el desgaste por micropitting.

        **Interrupción por Mantenimiento Autónomo (TPM):**
        Una inspección visual diaria de 3 minutos al inicio de turno permite al operador detectar la holgura o desgaste inicial del labio del rascador y reajustar el contrapeso, interrumpiendo el efecto dominó antes de que sobrecargue el reductor y el motor.
        """)

st.markdown("<div style='text-align:center;color:#7F8C8F;font-size:0.8rem;margin-top:1rem'>"
            "UTN FRSR · Gestión y Mantenimiento Electromecánico · TP N°1 · "
            "Herramienta de análisis interactivo</div>", unsafe_allow_html=True)