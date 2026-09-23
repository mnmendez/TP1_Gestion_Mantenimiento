# -*- coding: utf-8 -*-
"""
report_builder.py — Generador de Informe Técnico Oficial en Word (.docx)
Cátedra: Gestión y Mantenimiento Electromecánico — UTN FRSR (Año 2026)
Línea Principal de Elevación y Transporte de Áridos de Alta Densidad

Genera un documento Word (.docx) profesional con diseño institucional UTN FRSR,
respondiendo y justificando íntegramente las 6 consignas del enunciado oficial
del TP N°1, integrando Análisis de Causa Raíz (Ishikawa 6M y 5 Porqués),
gestión dinámica multigrupal de alumnos y estudio económico de motor Standby de 55 kW.
"""

import io
import pathlib
from typing import List, Dict, Any, Optional
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

# --- Paleta y estilos institucionales UTN FRSR ---
COLOR_UTN_NAVY = RGBColor(0x1F, 0x4E, 0x79)   # #1F4E79 Azul Acero UTN
COLOR_UTN_DARK = RGBColor(0x0E, 0x28, 0x41)   # #0E2841 Azul Noche
COLOR_GRAY_TEXT = RGBColor(0x59, 0x59, 0x59)  # #595959 Gris Secundario
COLOR_RED_ALERT = RGBColor(0xC0, 0x00, 0x00)  # Rojo Alerta Crítico
COLOR_GREEN_OK = RGBColor(0x1E, 0x7E, 0x34)   # Verde Aprobado / Ahorro
HEX_NAVY = "1F4E79"
HEX_LIGHT_BG = "F2F5F9"
HEX_BORDER = "D9D9D9"
HEX_ALT_ROW = "F9FBFD"
HEX_GREEN_LIGHT = "E8F5E9"


def _set_cell_background(cell, hex_color: str):
    """Aplica color de fondo hexadecimal a una celda de tabla."""
    shading_xml = f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))


def _set_cell_margins(cell, top: int = 100, bottom: int = 100, left: int = 150, right: int = 150):
    """Configura márgenes internos (padding) en dxa (1 pt = 20 dxa)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def _format_table_header(row, titles: List[str], hex_bg: str = HEX_NAVY, font_size: float = 9.0):
    """Estiliza la fila de encabezado de una tabla."""
    for idx, cell in enumerate(row.cells):
        _set_cell_background(cell, hex_bg)
        _set_cell_margins(cell, top=120, bottom=120, left=130, right=130)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(titles[idx])
        run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(font_size)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)


def _add_section_heading(doc: Document, text: str, level: int = 1):
    """Inserta encabezado numerado con tipografía y color corporativo UTN."""
    h = doc.add_heading(level=level)
    h.paragraph_format.space_before = Pt(14)
    h.paragraph_format.space_after = Pt(6)
    h.paragraph_format.keep_with_next = True
    run = h.add_run(text)
    run.font.name = "Calibri"
    run.bold = True
    if level == 1:
        run.font.size = Pt(14)
        run.font.color.rgb = COLOR_UTN_NAVY
    elif level == 2:
        run.font.size = Pt(12)
        run.font.color.rgb = COLOR_UTN_DARK
    else:
        run.font.size = Pt(11)
        run.font.color.rgb = COLOR_GRAY_TEXT
    return h


def _add_callout_box(doc: Document, title: str, text: str, hex_bg: str = HEX_LIGHT_BG, hex_border: str = HEX_NAVY):
    """Agrega una caja destacada (callout) técnica con borde izquierdo."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    _set_cell_background(cell, hex_bg)
    _set_cell_margins(cell, top=120, bottom=120, left=180, right=180)

    # Borde izquierdo grueso, otros nulos
    tcPr = cell._tc.get_or_add_tcPr()
    borders_xml = f'''
    <w:tcBorders {nsdecls("w")}>
        <w:top w:val="none"/>
        <w:left w:val="single" w:sz="24" w:space="0" w:color="{hex_border}"/>
        <w:bottom w:val="none"/>
        <w:right w:val="none"/>
    </w:tcBorders>'''
    tcPr.append(parse_xml(borders_xml))

    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    r_title = p.add_run(title + "\n")
    r_title.bold = True
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(10)
    r_title.font.color.rgb = COLOR_UTN_NAVY

    r_text = p.add_run(text)
    r_text.font.name = "Calibri"
    r_text.font.size = Pt(9.5)
    r_text.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def _normalize_students(students: Optional[List[Dict[str, Any]]], fallback_alumno: str, fallback_legajo: str) -> List[Dict[str, str]]:
    """Normaliza y valida el listado de alumnos para la portada institucional."""
    result = []
    if students:
        for s in students:
            if isinstance(s, dict):
                nom = str(s.get("nombre", "")).strip()
                ape = str(s.get("apellido", "")).strip()
                leg = str(s.get("legajo", "")).strip()
                if nom or ape or leg:
                    result.append({
                        "nombre": nom,
                        "apellido": ape,
                        "legajo": leg or "S/L"
                    })
    if not result:
        # Fallback por defecto: Martín Méndez / Legajo 11335
        result = [{
            "nombre": fallback_alumno or "Martín",
            "apellido": "Méndez" if "Méndez" not in fallback_alumno else "",
            "legajo": fallback_legajo or "11335"
        }]
    return result


def build_technical_report(
    students: Optional[List[Dict[str, Any]]] = None,
    standby_active: bool = False,
    alumno: str = "Martín Méndez",
    legajo: str = "11335",
    anio: str = "2026",
    TOP: float = 7200.0,
    cost_parada: float = 2500.0,
    cost_hh: float = 25.0,
    df_fallas=None,
    kpis_nominal=None,
    kpis_rep=None,
    df_matriz=None,
) -> io.BytesIO:
    """
    Construye el documento Word oficial del TP N°1 respetando íntegramente
    el enunciado, tablas y justificaciones técnicas de ingeniería electromecánica.
    
    Parámetros:
    - students: Lista de diccionarios [{'nombre': '...', 'apellido': '...', 'legajo': '...'}]
    - standby_active: Booleano que activa el escenario con motor Standby de 55 kW en pañol.
    - alumno / legajo: Compatibilidad hacia atrás si no se especifica students.
    """
    doc = Document()
    student_list = _normalize_students(students, alumno, legajo)

    # --- Configuración de página (Márgenes estándar de 2.2 cm / 0.85 pulg) ---
    for section in doc.sections:
        section.top_margin = Inches(0.85)
        section.bottom_margin = Inches(0.85)
        section.left_margin = Inches(0.85)
        section.right_margin = Inches(0.85)

        # Encabezado institucional
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run("UTN FRSR · Cátedra de Gestión y Mantenimiento Electromecánico · TP N°1")
        hrun.font.name = "Calibri"
        hrun.font.size = Pt(8.5)
        hrun.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

        # Pie de página institucional
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        frun = fp.add_run("Línea Principal de Elevación y Transporte de Áridos — Confiabilidad, Ishikawa 6M y KPIs")
        frun.font.name = "Calibri"
        frun.font.size = Pt(8.5)
        frun.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    # =========================================================================
    # PORTADA INSTITUCIONAL UTN FRSR
    # =========================================================================
    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst.paragraph_format.space_before = Pt(8)
    p_inst.paragraph_format.space_after = Pt(2)

    r1 = p_inst.add_run("UNIVERSIDAD TECNOLÓGICA NACIONAL\n")
    r1.bold = True
    r1.font.name = "Calibri"
    r1.font.size = Pt(17)
    r1.font.color.rgb = COLOR_UTN_NAVY

    r2 = p_inst.add_run("FACULTAD REGIONAL SAN RAFAEL\n")
    r2.bold = True
    r2.font.name = "Calibri"
    r2.font.size = Pt(13.5)
    r2.font.color.rgb = COLOR_GRAY_TEXT

    r3 = p_inst.add_run("DEPARTAMENTO DE INGENIERÍA ELECTROMECÁNICA\nCátedra: Gestión y Mantenimiento Electromecánico")
    r3.font.name = "Calibri"
    r3.font.size = Pt(11)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Título Principal del Trabajo Práctico
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(4)
    r_tp = p_title.add_run("TRABAJO PRÁCTICO N° 1\n")
    r_tp.bold = True
    r_tp.font.size = Pt(18)
    r_tp.font.color.rgb = COLOR_UTN_NAVY

    r_tema = p_title.add_run(
        "INFORME TÉCNICO DE INGENIERÍA DE MANTENIMIENTO Y CONFIABILIDAD\n"
        "Línea de Elevación y Transporte de Áridos de Alta Densidad\n"
        "Diagnóstico Operativo, Matriz de Criticidad, Ishikawa 6M, 5 Porqués y Plan de Mantenimiento"
    )
    r_tema.bold = True
    r_tema.font.size = Pt(11)
    r_tema.font.color.rgb = COLOR_UTN_DARK

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Tabla de Nómina Oficial de Integrantes del Grupo (R2)
    p_roster_tit = doc.add_paragraph()
    p_roster_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_roster_tit.paragraph_format.space_after = Pt(4)
    r_rtit = p_roster_tit.add_run("NÓMINA DE ALUMNOS / INTEGRANTES DEL EQUIPO DE TRABAJO")
    r_rtit.bold = True
    r_rtit.font.name = "Calibri"
    r_rtit.font.size = Pt(10.5)
    r_rtit.font.color.rgb = COLOR_UTN_NAVY

    tbl_roster = doc.add_table(rows=len(student_list) + 1, cols=3)
    tbl_roster.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_roster.autofit = False

    _format_table_header(tbl_roster.rows[0], ["N°", "Nombre y Apellido", "Legajo Oficial"], font_size=9.5)
    col_w_roster = [Inches(0.6), Inches(3.6), Inches(1.8)]

    for idx, s in enumerate(student_list):
        row = tbl_roster.rows[idx + 1]
        bg = HEX_ALT_ROW if idx % 2 == 1 else "FFFFFF"
        full_name = f"{s.get('apellido', '')} {s.get('nombre', '')}".strip() or s.get('nombre', '')
        vals = [str(idx + 1), full_name, str(s.get("legajo", ""))]
        for c_idx, val in enumerate(vals):
            cell = row.cells[c_idx]
            cell.width = col_w_roster[c_idx]
            _set_cell_background(cell, bg)
            _set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx != 1 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.name = "Calibri"
            run.font.size = Pt(9.5)
            if c_idx == 0:
                run.bold = True

    p_ciclo = doc.add_paragraph()
    p_ciclo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ciclo.paragraph_format.space_before = Pt(14)
    p_ciclo.paragraph_format.space_after = Pt(4)
    rciclo = p_ciclo.add_run(f"Ciclo Lectivo: {anio} · San Rafael, Mendoza, Argentina")
    rciclo.font.size = Pt(10)
    rciclo.font.color.rgb = COLOR_GRAY_TEXT

    doc.add_page_break()

    # =========================================================================
    # SECCIÓN 1: CONTEXTO INDUSTRIAL Y PARÁMETROS OPERATIVOS
    # =========================================================================
    _add_section_heading(doc, "1. Contexto Industrial y Especificaciones del Activo", level=1)
    p_ctx = doc.add_paragraph()
    p_ctx.add_run(
        f"El presente trabajo de ingeniería se desarrolla sobre la Línea Principal de Elevación y Transporte de "
        f"Áridos de Alta Densidad de un complejo industrial de procesamiento continuo. La línea opera nominalmente "
        f"en régimen continuo de 24 horas por día durante 300 días al año ({TOP:,.0f} horas anuales programadas). "
        f"El flujo continuo de material abastece de forma directa e ininterrumpida a un reactor químico principal, "
        f"en el cual cualquier detención imprevista en el suministro de sólidos genera la petrificación prematura de la masa "
        f"en proceso, provocando pérdidas catastróficas por lucro cesante, daño de equipos y extensas paradas no programadas."
    )

    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    p_comp = doc.add_paragraph()
    p_comp.add_run("Componentes Principales del Sistema Electromecánico:\n").bold = True
    bullet_points = [
        ("Unidad Motriz (SYS-MOT):", " Motor asincrónico trifásico de jaula de ardilla, 55 kW, 4 polos, 1475 RPM, 380/660 V, "
                                     "50 Hz, aislamiento dieléctrico Clase F (155 °C), grado de protección IP55, acoplado mediante arrancador suave electrónico."),
        ("Reducción de Velocidad:", " Reductor ortogonal de tres etapas de engranajes helicoidales cónicos rectificados, "
                                    "lubricado por barboteo con aceite sintético ISO VG 220, relación de reducción i = 28:1."),
        ("Polea Motriz y Apoyos:", " Polea de accionamiento de Ø500 mm con revestimiento elastomérico cerámico vulcanizado, montada "
                                   "sobre eje macizo de acero bonificado SAE 4140, soportado por dos cajas de rodamientos de pie monobloc "
                                   "(SNH) con rodamientos oscilantes de rodillos esféricos SKF 22220 EK y fijación por manguito cónico."),
        ("Sistema de Tensionado y Control (SYS-TEN):", " Centralita oleohidráulica compacta con bomba de engranajes (20 bar de presión nominal), "
                                                       "cilindro hidráulico de doble efecto para tensionado automático de banda, válvulas check pilotadas y presostatos de enclavamiento."),
        ("Limpieza y Protección (SYS-EST / SYS-ELE):", " Rascador primario frontal de poliuretano de alta densidad y rascador secundario en V, "
                                                       "acompañados de sensores inductivos de desalineación lateral de banda e interruptores de parada por cable tirador.")
    ]
    for title, desc in bullet_points:
        bp = doc.add_paragraph(style='List Bullet')
        bp.paragraph_format.space_before = Pt(1)
        bp.paragraph_format.space_after = Pt(2)
        r_t = bp.add_run(title)
        r_t.bold = True
        r_t.font.color.rgb = COLOR_UTN_NAVY
        bp.add_run(desc)

    # =========================================================================
    # SECCIÓN 2: REGISTRO HISTÓRICO DE CONTINGENCIAS Y PARÁMETROS ECONÓMICOS
    # =========================================================================
    _add_section_heading(doc, "2. Registro Histórico de Contingencias y Parámetros Económicos", level=1)
    p_reg = doc.add_paragraph()
    p_reg.add_run(
        f"Durante el último ciclo operativo anual ({TOP:,.0f} h nominales), el activo registró 5 eventos de falla "
        f"intempestivos que ocasionaron tiempos improductivos, erogaciones en repuestos y consumo de mano de obra técnica:"
    )

    tbl_fallas = doc.add_table(rows=6, cols=5)
    tbl_fallas.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_fallas.autofit = False

    t1_headers = ["Evento", "Descripción Técnica de la Falla", "TTP (h)", "Repuestos (USD)", "Mano de Obra (HH)"]
    _format_table_header(tbl_fallas.rows[0], t1_headers)

    t1_data = [
        ("F1", "Picado severo (spalling) en pista exterior de rodamiento SKF 22220 EK por contaminación con polvo abrasivo y degradación térmica de grasa. Desmontaje completo de polea motriz.", "28.0", "$1,850", "48"),
        ("F2", "Sobrecalentamiento y cortocircuito entre espiras en estator de motor eléctrico de 55 kW por bloqueo de ventilador y capa de polvo sobre carcasa. Rebobinado integral en taller externo.", "42.0", "$3,400", "24"),
        ("F3", "Fuga masiva de fluido hidráulico por fisura en latiguillo de alta presión debido a fatiga por pulsación cíclica en centralita de 20 bar. Caída de presión y disparo de interbloqueo.", "6.0", "$320", "12"),
        ("F4", "Desalineación severa de banda por acumulación de material arcilloso en tambor de reenvío. Rasgado longitudinal de 15 m de cubierta superior. Vulcanizado en caliente por terceros.", "36.0", "$4,200", "36"),
        ("F5", "Desgaste prematuro y atrancamiento de rascador de uretano contra la banda, induciendo sobrecarga mecánica y sobretemperatura severa en reductor ortogonal.", "8.0", "$450", "16"),
    ]

    col_w_t1 = [Inches(0.6), Inches(3.4), Inches(0.8), Inches(1.0), Inches(0.7)]
    for r_idx, row_data in enumerate(t1_data):
        row = tbl_fallas.rows[r_idx + 1]
        bg = HEX_ALT_ROW if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.width = col_w_t1[c_idx]
            _set_cell_background(cell, bg)
            _set_cell_margins(cell, top=80, bottom=80, left=90, right=90)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in [0, 2, 4] else (WD_ALIGN_PARAGRAPH.RIGHT if c_idx == 3 else WD_ALIGN_PARAGRAPH.LEFT)
            run = p.add_run(val)
            run.font.name = "Calibri"
            run.font.size = Pt(8.5)
            if c_idx == 0:
                run.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    p_eco = doc.add_paragraph()
    p_eco.add_run("Parámetros Económicos y Tarifas Oficiales de Planta:\n").bold = True
    p_eco.add_run(f"• Costo de Parada de Planta (Lucro cesante + Costos fijos): ${cost_parada:,.0f} USD/hora.\n")
    p_eco.add_run(f"• Tarifa Horaria de Mano de Obra Propia de Mantenimiento: ${cost_hh:,.0f} USD/HH.")

    # =========================================================================
    # SECCIÓN 3: RESOLUCIÓN DE CONSIGNAS DEL ENUNCIADO OFICIAL
    # =========================================================================
    _add_section_heading(doc, "3. Resolución Técnica de Consignas Oficiales (C1 a C6)", level=1)

    # Consigna 1
    _add_section_heading(doc, "Consigna 1: Taxonomía de Activos (ISO 14224) y Matriz de Criticidad", level=2)
    p_1a = doc.add_paragraph()
    p_1a.add_run("a) Arbolado Jerárquico del Sistema y Criterio de Decisión Reparación vs. Reemplazo:\n").bold = True
    p_1a.add_run(
        "De acuerdo con los lineamientos de la norma internacional ISO 14224, la estructura jerárquica del activo "
        "se clasifica en niveles sucesivos para asegurar la integridad de la gestión técnica y contable:\n"
        "• Nivel 2: Complejo Industrial de Molienda y Procesamiento de Minerales.\n"
        "• Nivel 3: Sección Operativa de Elevación y Transporte de Áridos.\n"
        "• Nivel 4 (Unidad de Equipo Principal): Cinta Transportadora Principal (CV-01).\n"
        "• Nivel 5 (Sistemas Funcionales): SYS-MOT (Unidad Motriz), SYS-BAN (Banda y Rodillos), SYS-TEN (Tensionado Oleohidráulico), SYS-EST (Estructura y Limpieza), SYS-ELE (Control y Seguridad Eléctrica).\n"
        "• Nivel 6 (Componente / Ítem Mantenible): Rodamiento SKF 22220 EK, Bobinado de estator, Latiguillo hidráulico flexible, Hoja rascadora de uretano.\n"
        "• Nivel 8 (Parte Indivisible): Pistas de rodadura de rodamiento, retenes de labio, aislación de esmalte de cobre."
    )

    _add_callout_box(
        doc,
        "Límite Metodológico: Reparación (MTBF) vs. Reemplazo (MTTF)",
        "• Nivel de Reparación (Niveles 4 y 5): Activos mecánicos complejos reparables (motor 55 kW, reductor de velocidad, polea motriz). "
        "Se gestionan bajo el parámetro MTBF (Mean Time Between Failures), buscando extender su confiabilidad mediante intervenciones programadas.\n"
        "• Nivel de Reemplazo (Niveles 6 y 8): Componentes unitarios no reparables económicamente (rodamientos, mangueras hidráulicas, rascadores). "
        "Se gestionan bajo el parámetro MTTF (Mean Time To Failure), realizándose su descarte definitivo y sustitución ante detección de degradación."
    )

    p_1b = doc.add_paragraph()
    p_1b.add_run("b) Matriz de Criticidad de Subsistemas (Escalas Discretas Cátedra):\n").bold = True
    p_1b.add_run(
        "Se evalúa el riesgo relativo mediante la formulación oficial Ci = S + A + O + D y Criticidad Total C = F × Ci, "
        "respetando estrictamente las escalas discretas de la cátedra:\n"
        "• Frecuencia (F): 1 (Rara), 2 (Ocasional), 3 (Frecuente), 4 (Muy Frecuente).\n"
        "• Seguridad (S): 1 (Sin daño), 2 (Lesión leve), 3 (Grave / Incapacitante).\n"
        "• Medio Ambiente (A): 1 (Sin impacto), 2 (Contenido local), 3 (Severo / Derrame externo).\n"
        "• Impacto Operacional (O): 1 (Sin parada), 3 (Parada parcial), 10 (Parada total con riesgo de petrificación en reactor).\n"
        "• Costo Directo (D): 1 (< $1,000 USD), 2 ($1,000 - $3,000 USD), 4 (> $3,000 USD).\n"
        "Límites de Clasificación: Clase A - CRÍTICO (C ≥ 15), Clase B - SEMICRÍTICO (10 ≤ C < 15), Clase C - NO CRÍTICO (C < 10)."
    )

    tbl_crit = doc.add_table(rows=6, cols=9)
    tbl_crit.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_crit.autofit = False
    crit_headers = ["Subsistema", "F", "S", "A", "O", "D", "Ci", "C", "Clase Asignada"]
    _format_table_header(tbl_crit.rows[0], crit_headers)

    crit_data = [
        ("Unidad Motriz (SYS-MOT)", "3", "2", "1", "10", "4", "17", "51", "Clase A - Crítico"),
        ("Banda y Rodillos (SYS-BAN)", "3", "2", "1", "10", "4", "17", "51", "Clase A - Crítico"),
        ("Sistema de Tensión (SYS-TEN)", "2", "2", "2", "3", "1", "8", "16", "Clase A / B"),
        ("Estructura y Chasis (SYS-EST)", "2", "1", "1", "3", "1", "6", "12", "Clase B - Semicrítico"),
        ("Tablero y Control (SYS-ELE)", "1", "2", "1", "3", "1", "7", "7", "Clase C - No Crítico"),
    ]
    if df_matriz is not None and not df_matriz.empty:
        crit_data = []
        for _, row in df_matriz.iterrows():
            crit_data.append((
                str(row.get("subsistema", "")),
                str(row.get("F", "")),
                str(row.get("S", "")),
                str(row.get("A", "")),
                str(row.get("O", "")),
                str(row.get("D", "")),
                str(row.get("Ci", "")),
                str(row.get("C", "")),
                str(row.get("clase", ""))
            ))

    col_w_crit = [Inches(1.8), Inches(0.4), Inches(0.4), Inches(0.4), Inches(0.4), Inches(0.4), Inches(0.6), Inches(0.6), Inches(1.5)]
    for r_idx, row_vals in enumerate(crit_data):
        row = tbl_crit.rows[r_idx + 1]
        bg = HEX_ALT_ROW if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            cell.width = col_w_crit[c_idx]
            _set_cell_background(cell, bg)
            _set_cell_margins(cell, top=60, bottom=60, left=70, right=70)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(str(val))
            run.font.name = "Calibri"
            run.font.size = Pt(8.5)
            if c_idx == 8:
                run.bold = True
                if "A" in str(val):
                    run.font.color.rgb = COLOR_RED_ALERT
                elif "B" in str(val):
                    run.font.color.rgb = COLOR_UTN_NAVY

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    p_1c = doc.add_paragraph()
    p_1c.add_run("c) Modelos de Mantenimiento Justificados:\n").bold = True
    p_1c.add_run(
        "• Clase A — Modelo de Alta Disponibilidad: Aplicado a Unidad Motriz y Banda. Combina monitoreo continuo de vibraciones "
        "(envolvente y aceleración), termografía infrarroja, análisis de lubricante y overhaul mayor programado en paradas de planta.\n"
        "• Clase B — Modelo Sistemático / Predictivo: Aplicado a Sistema de Tensión y Estructura. Sustitución a vida límite (mangueras "
        "hidráulicas cada 24 meses) e inspección periódica de presiones.\n"
        "• Clase C — Modelo Condicional y Mantenimiento Autónomo (TPM): Rutinas diarias CIL realizadas por operadores para limpieza de "
        "rascadores y carcasas, permitiéndose el correctivo a la rotura solo si no induce daños secundarios."
    )

    # Consigna 2
    _add_section_heading(doc, "Consigna 2: Plan de Inspecciones Sensoriales (TPM) e Instrumentales (PdM)", level=2)
    tbl_insp = doc.add_table(rows=7, cols=5)
    tbl_insp.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_insp.autofit = False
    insp_headers = ["Tipo Inspección", "Componente Objetivo", "Técnica / Sentido", "Parámetro de Control", "Frecuencia"]
    _format_table_header(tbl_insp.rows[0], insp_headers)

    insp_data = [
        ("Sensorial (TPM)", "Rascadores de uretano", "Visual / Palpado", "Desgaste de filos, tensión contra banda, atascamientos", "Diaria (por turno)"),
        ("Sensorial (TPM)", "Tambor de reenvío y polea", "Visual / Auditivo", "Acumulación de áridos en cilindro, chirridos de rozamiento", "Diaria"),
        ("Sensorial (TPM)", "Central hidráulica y mangueras", "Visual / Olfativo", "Fugas superficiales, nivel de tanque, pulsación excesiva", "Diaria"),
        ("Instrumental (PdM)", "Rodamientos SKF 22220 EK", "Vibrómetro / Ultrasonido", "Velocidad RMS (ISO 10816-3) y aceleración de envolvente", "Quincenal"),
        ("Instrumental (PdM)", "Motor eléctrico 55 kW", "Termografía infrarroja", "Gradiente térmico carcasa/bornera (< 75 °C) y consumo trifásico", "Mensual"),
        ("Instrumental (PdM)", "Reductor ortogonal 28:1", "Análisis físico-químico aceite", "Viscosidad a 40 °C, ppm agua, conteo de partículas ISO 4406", "Trimestral"),
    ]
    col_w_insp = [Inches(1.2), Inches(1.3), Inches(1.3), Inches(1.7), Inches(1.0)]
    for r_idx, row_vals in enumerate(insp_data):
        row = tbl_insp.rows[r_idx + 1]
        bg = HEX_ALT_ROW if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            cell.width = col_w_insp[c_idx]
            _set_cell_background(cell, bg)
            _set_cell_margins(cell, top=60, bottom=60, left=70, right=70)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in [0, 4] else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(val)
            run.font.name = "Calibri"
            run.font.size = Pt(8.5)
            if c_idx == 0:
                run.bold = True

    # Consigna 3
    _add_section_heading(doc, "Consigna 3: Asignación de Estrategias de Mantenimiento Proactivo", level=2)
    tbl3 = doc.add_table(rows=6, cols=5)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl3.autofit = False
    t3_headers = ["Evento", "Componente", "Tipo Mantenimiento", "Tarea de Inspección Técnica", "Frecuencia"]
    _format_table_header(tbl3.rows[0], t3_headers)

    t3_data = [
        ("F1", "Rodamiento SKF 22220 EK", "Predictivo (CBM)", "Análisis de vibraciones por demodulación y relubricación asistida por ultrasonido acústico.", "Quincenal"),
        ("F2", "Motor asincrónico 55 kW", "Preventivo / TPM", "Limpieza de tobera de ventilador y carcasa (TPM) + Termografía y megado de aislación.", "Semanal / Mensual"),
        ("F3", "Latiguillo de 20 bar", "Preventivo Sistemático", "Inspección de fatiga por pulsación y sustitución calendarizada preventiva cada 24 meses.", "Mensual / 2 años"),
        ("F4", "Banda transportadora 15 m", "TPM y Predictivo", "Limpieza de tambor de cola por rascador interior + Chequeo funcional de sensores inductivos.", "Diaria"),
        ("F5", "Rascador de uretano", "Mantenimiento Autónomo (TPM)", "Inspección de labio de contacto, ajuste de resorte tensor y remoción de costra de mineral.", "Diaria")
    ]
    col_w_t3 = [Inches(0.6), Inches(1.4), Inches(1.3), Inches(2.2), Inches(1.0)]
    for r_idx, row_vals in enumerate(t3_data):
        row = tbl3.rows[r_idx + 1]
        bg = HEX_ALT_ROW if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            cell.width = col_w_t3[c_idx]
            _set_cell_background(cell, bg)
            _set_cell_margins(cell, top=60, bottom=60, left=70, right=70)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in [0, 4] else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(val)
            run.font.name = "Calibri"
            run.font.size = Pt(8.5)
            if c_idx == 0:
                run.bold = True

    # Consigna 4
    _add_section_heading(doc, "Consigna 4: Determinación de KPIs y Análisis de Patrones de Nowlan & Heap", level=2)
    nom = kpis_nominal or {"ttp": 120.0, "tfr": 7080.0, "mtbf": 1416.0, "mttr": 24.0, "ai": 98.33, "ao": 98.33, "n": 5}
    rep = kpis_rep or {"ttp": 120.0, "tfr": 7080.0, "mtbf": 786.67, "mttr": 13.33, "ai": 98.33, "ao": 98.33, "n": 9}

    p_kpis = doc.add_paragraph()
    p_kpis.add_run("a) Memoria de Cálculo de Indicadores de Confiabilidad y Mantenibilidad:\n").bold = True
    p_kpis.add_run(
        f"• Tiempo Operativo Programado (TOP): {TOP:,.0f} h/año (24 h × 300 días).\n"
        f"• Tiempo Total de Paradas por Falla (TTP): {nom['ttp']:.1f} h (sumatoria F1 a F5).\n"
        f"• Tiempo de Funcionamiento Real (TFR): TOP - TTP = {nom['tfr']:,.1f} h.\n"
        f"• Caso Base Nominal (N = {nom['n']}):\n"
        f"   - MTBF = TFR / N = {nom['mtbf']:,.2f} h | MTTR = TTP / N = {nom['mttr']:,.2f} h.\n"
        f"   - Disponibilidad Inherente (Ai) = {nom['ai']:.2f}% | Disponibilidad Operacional (Ao) = {nom['ao']:.2f}%.\n"
        f"• Caso Refinado con Repetición Estimada de Fallas (N_est = {rep['n']}):\n"
        f"   - MTBF = {rep['mtbf']:,.2f} h | MTTR = {rep['mttr']:,.2f} h | Ai = Ao = {rep['ai']:.2f}%."
    )

    p_nh = doc.add_paragraph()
    p_nh.add_run("b) Análisis de Patrones de Falla de Nowlan & Heap:\n").bold = True
    p_nh.add_run(
        "¿Sigue la línea de transporte la 'curva de la bañera' (Patrón A)?\n"
        "RESPUESTA: NO. La investigación clásica de Nowlan & Heap (1978) demostró que la curva de la bañera (Patrón A) "
        "solo representa el 4% de las fallas en la industria moderna, mientras que los patrones dependientes de la edad (A, B, C) "
        "suman apenas el 11%. En marcado contraste, los Patrones Aleatorios e Incondicionales a la Edad (D, E, F) representan el 89%:\n"
        "• Patrón D (7%): Falla aleatoria tras período inicial de asentamiento.\n"
        "• Patrón E (14%): Tasa de falla estrictamente constante y aleatoria durante toda la vida útil.\n"
        "• Patrón F (68%): Altísima mortalidad infantil inicial seguida de tasa de falla aleatoria constante.\n\n"
        "En nuestra cinta transportadora, el 100% de las fallas obedecen a los Patrones D, E y F:\n"
        "• F1 (Rodamiento): Contaminación accidental con polvo de sílice y daño en retén (Patrón E/F), no fatiga L10.\n"
        "• F2 (Motor): Bloqueo aerodinámico de ventilador por barro mineral y calentamiento extremo (Patrón E).\n"
        "• F3 (Latiguillo): Fatiga cíclica por picos de pulsación hidráulica (Patrón D/E).\n"
        "• F4 (Cinta): Rasgado aleatorio por atascamiento de áridos en cola (Patrón E).\n"
        "• F5 (Rascador): Trabamiento por acumulación de barro (Patrón D/E).\n\n"
        "Conclusión Estratégica: Desarmar equipos por horas de calendario no previene estas contingencias y aumenta "
        "drásticamente el riesgo de inducir fallas de mortalidad infantil (Patrón F). La estrategia óptima es el CBM y TPM."
    )

    # Consigna 5
    _add_section_heading(doc, "Consigna 5: Balance Económico y Caso de Negocio de Motor Standby", level=2)
    tbl4 = doc.add_table(rows=7, cols=7)
    tbl4.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl4.autofit = False
    t4_headers = ["Evento", "TTP (h)", "Costo Parada (USD)", "Repuestos (USD)", "HH", "Costo MO (USD)", "Costo Total (USD)"]
    _format_table_header(tbl4.rows[0], t4_headers)

    t4_data = [
        ("F1", "28", "$70,000", "$1,850", "48", "$1,200", "$73,050"),
        ("F2", "42", "$105,000", "$3,400", "24", "$600", "$109,000"),
        ("F3", "6", "$15,000", "$320", "12", "$300", "$15,620"),
        ("F4", "36", "$90,000", "$4,200", "36", "$900", "$95,100"),
        ("F5", "8", "$20,000", "$450", "16", "$400", "$20,850"),
        ("TOTAL", "120", "$300,000", "$10,220", "136", "$3,400", "$313,620")
    ]
    col_w_t4 = [Inches(0.6), Inches(0.7), Inches(1.3), Inches(1.0), Inches(0.6), Inches(1.0), Inches(1.3)]
    for r_idx, row_vals in enumerate(t4_data):
        row = tbl4.rows[r_idx + 1]
        is_tot = (r_idx == len(t4_data) - 1)
        bg = HEX_LIGHT_BG if is_tot else (HEX_ALT_ROW if r_idx % 2 == 1 else "FFFFFF")
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            cell.width = col_w_t4[c_idx]
            _set_cell_background(cell, bg)
            _set_cell_margins(cell, top=60, bottom=60, left=70, right=70)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in [0, 1, 4] else WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(val)
            run.font.name = "Calibri"
            run.font.size = Pt(8.5)
            if is_tot or c_idx == 0:
                run.bold = True
            if is_tot:
                run.font.color.rgb = COLOR_UTN_NAVY

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    p_standby = doc.add_paragraph()
    p_standby.add_run("Estudio de Caso: Incorporación de Motor Standby de 55 kW en Pañol (R3):\n").bold = True
    p_standby.add_run(
        "1. Diagnóstico de Cuello de Botella: Las 42 horas históricas de parada del evento F2 correspondieron a esperas logísticas "
        "externas (transporte a taller, rebobinado manual, barnizado e impregnación al vacío, horneado y ensayos dieléctricos).\n"
        "2. Propuesta Técnica: Adquirir y mantener en pañol 1 motor trifásico idéntico (55 kW, 4 polos, carcasa B3 normalizada). "
        "Ante falla, se realiza la sustitución electromecánica directa in situ con alineación láser de ejes.\n"
        "3. Demostración Financiera y Operativa:\n"
        "   • Reducción de MTTR de F2: De 42.0 h a 4.0 h (ahorro directo de 38 horas de parada de planta).\n"
        "   • Ahorro Neto Anual por Lucro Cesante: $105,000 USD - (4 h × $2,500 USD/h = $10,000 USD) = $95,000.00 USD de ahorro neto.\n"
        "   • Impacto en KPIs Globales:\n"
        "     - Tiempo Total de Paradas (TTP): Se reduce de 120.0 h a 82.0 h.\n"
        "     - MTTR Global de Planta: Se reduce de 24.00 h a 16.40 h (82.0 h / 5 fallas).\n"
        "     - Disponibilidad Global de Planta: Aumenta de 98.33% a 98.86% (+0.53% de disponibilidad productiva neta).\n"
        "   • Retorno de la Inversión (ROI) y Período de Recuperación (Payback):\n"
        "     - Costo de Inversión Inicial (I0): $4,000 USD (motor 55 kW IE3).\n"
        "     - Costo de Posesión Anual: ~$400 USD/año (mantenimiento en pañol y rotación trimestral de rotor).\n"
        "     - Beneficio Neto en el Año 1: $95,000 USD - $4,000 USD - $400 USD = $90,600 USD.\n"
        "     - ROI: (($95,000 - $4,000) / $4,000) × 100% = 2,275%.\n"
        "     - Payback: $4,000 / ($95,000 / 365 días) = 15.3 días de operación equivalente (o 1.6 horas de producción recuperada)."
    )

    if standby_active:
        _add_callout_box(
            doc,
            "ESTADO ACTIVO: CONFIGURACIÓN CON MOTOR STANDBY IMPLEMENTADA",
            "El escenario evaluado incorpora formalmente la disponibilidad del motor Standby de 55 kW en almacén:\n"
            "• TTP de Falla F2: 4.0 horas (antes 42.0 h).\n"
            "• Tiempo Total de Paradas Anuales: 82.0 horas (antes 120.0 h).\n"
            "• Costo de Parada Anual: $205,000 USD (antes $300,000 USD -> Ahorro: $95,000 USD).\n"
            "• MTTR Global de Línea: 16.40 horas (antes 24.00 h).\n"
            "• Disponibilidad Operacional: 98.86% (antes 98.33%).",
            hex_bg=HEX_GREEN_LIGHT,
            hex_border="1E7E34"
        )

    # Consigna 6
    _add_section_heading(doc, "Consigna 6: Modelo de Efecto Dominó en Falla F5 y Barrera TPM", level=2)
    p_6 = doc.add_paragraph()
    p_6.add_run(
        "Se modela la cadena electromecánica desencadenada por la falla F5 (atrancamiento de rascador):\n"
        "1. Desgaste y Bloqueo Mecánico: La lámina de uretano sufre desgaste asimétrico y acumulación de costra en la articulación, trabándose en ángulo agudo contra la banda.\n"
        "2. Aumento de Rozamiento: Se incrementa drásticamente la fuerza de fricción tangencial (Fr = μ · Fn), convirtiendo el rascador en un freno continuo.\n"
        "3. Reacción Electromecánica en el Motor: Para vencer la cupla resistente (Tres), el motor de inducción incrementa su deslizamiento s = (ns - n)/ns, "
        "lo que eleva la corriente estatórica demandada, disparando el calentamiento por pérdidas Joule (P = 3 · I² · R).\n"
        "4. Daño Térmico en Reductor: La transmisión de sobretorque calienta el aceite sintético ISO VG 220 por encima de 95 °C, degradando la viscosidad "
        "y colapsando la película elastohidrodinámica (EHL), induciendo desgaste prematuro por micropitting.\n"
        "5. Bloqueo TPM: Una rutina diaria de 3 minutos al inicio de turno realizada por el operador (limpieza de costras y calibración de resorte) "
        "anula el Eslabón 1 y protege el tren de potencia completo."
    )

    # =========================================================================
    # SECCIÓN 4: ANÁLISIS DE CAUSA RAÍZ (ISHIKAWA 6M Y 5 PORQUÉS) (R1)
    # =========================================================================
    _add_section_heading(doc, "4. Metodología de Análisis de Causa Raíz de Contingencias (Ishikawa 6M y 5 Porqués)", level=1)
    p_rca_intro = doc.add_paragraph()
    p_rca_intro.add_run(
        "Para erradicar la recurrencia de las contingencias operativas, se aplica un análisis riguroso de resolución "
        "de problemas de ingeniería combinando el Diagrama de Causa y Efecto (Ishikawa 6M) para la visión sistémica "
        "del activo y el Árbol de los 5 Porqués para investigar la profundidad física, humana y latente de los eventos críticos."
    )

    # 4.1 Ishikawa 6M
    _add_section_heading(doc, "4.1 Diagrama de Ishikawa de las 6M Aplicado a la Cinta y Eventos F1 a F5", level=2)
    tbl_ishi = doc.add_table(rows=7, cols=3)
    tbl_ishi.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_ishi.autofit = False
    _format_table_header(tbl_ishi.rows[0], ["Dimensión (6M)", "Factores Contribuyentes Identificados en el Activo", "Eventos Asociados"])

    ishi_data = [
        ("Maquinaria\n(Machine)",
         "• F1: Retén laberíntico defectuoso en soporte monobloc SNH con ingreso de sílice.\n"
         "• F2: Tobera del ventilador con paso de rejilla inadecuado para polvo mineral.\n"
         "• F3: Circuito hidráulico de 20 bar sin amortiguador de pulsaciones; radio de curvatura crítico.\n"
         "• F4: Tambor de reenvío cilíndrico liso en vez de tambor autolimpiante en jaula de ardilla.\n"
         "• F5: Soporte de rascador rígido sin resorte autocompensador de desgaste de hoja.",
         "F1, F2, F3, F4, F5"),
        ("Mano de Obra\n(Manpower)",
         "• F1: Sobredosis manual de grasa con grasera sin retirar tapón de purga (reventón de retén).\n"
         "• F2: Omisión de verificación táctil/visual del flujo de aire forzado en rondas de turno.\n"
         "• F3: Montaje de latiguillo con torsión axial indebida sobre el cuerpo elastomérico.\n"
         "• F4: Omisión de control visual de alineación de banda en cola en arranque de turno.\n"
         "• F5: Falta de regulación periódica de tensión y perpendicularidad del filo rascador.",
         "F1, F2, F3, F4, F5"),
        ("Método\n(Method)",
         "• F1: Frecuencia de relubricación fija por calendario en lugar de lubricación por ultrasonido.\n"
         "• F2: Inexistencia de un estándar CIL (Mantenimiento Autónomo TPM) para soplado semanal de carcasa.\n"
         "• F3: Política 'run-to-failure' en componentes flexibles en vez de cambio preventivo a 24 meses.\n"
         "• F4: Procedimiento que no contempla parada inmediata ante acumulación en tolva de carga.\n"
         "• F5: Falta de checklist operativo de 3 minutos al inicio de cada turno de trabajo.",
         "F1, F2, F3, F4, F5"),
        ("Materiales\n(Materials)",
         "• F1: Grasa mineral básica con viscosidad insuficiente y sin aditivos sellantes para polvo de cuarzo.\n"
         "• F2: Costra de polvo mineral compactado con bajísima conductividad térmica (k < 0.2 W/m·K).\n"
         "• F3: Manguera de 1 sola malla de acero (1SN) sometida a fatiga severa en 20 bar.\n"
         "• F4: Cubierta de caucho estándar sin cordones transversales antidesgarro (rip-stop breakers).\n"
         "• F5: Poliuretano de baja resiliencia que sufre desgaste abrasivo asimétrico y desgarro.",
         "F1, F2, F3, F4, F5"),
        ("Medio Ambiente\n(Milieu / Entorno)",
         "• F1: Atmósfera altamente pulvígena en cabezal de accionamiento por caída de finos.\n"
         "• F2: Temperatura ambiente estival elevada (> 38 °C) en recinto confinado del motor.\n"
         "• F3: Radiación solar directa y salpicaduras de hidrocarburos sobre cubierta de mangueras.\n"
         "• F4: Caída accidental de terrones de áridos sobre el ramal inferior de retorno de la banda.\n"
         "• F5: Mezcla de finos con humedad ambiental que produce pasta arcillosa de alta adherencia.",
         "F1, F2, F3, F4, F5"),
        ("Medición\n(Measurement)",
         "• F1: Inexistencia de monitoreo periódico de vibraciones con demodulación de envolvente.\n"
         "• F2: Falta de sondas térmicas PT-100 en bobinas conectadas a disparo de arrancador suave.\n"
         "• F3: Ausencia de manómetro con glicerina para registrar transitorios y golpes de ariete.\n"
         "• F4: Sensores inductivos de desalineación descalibrados o bloqueados por barro mineral.\n"
         "• F5: Falta de monitoreo de corriente por fase en motor y sin sensor térmico en reductor.",
         "F1, F2, F3, F4, F5"),
    ]

    col_w_ishi = [Inches(1.3), Inches(4.3), Inches(0.9)]
    for r_idx, row_vals in enumerate(ishi_data):
        row = tbl_ishi.rows[r_idx + 1]
        bg = HEX_ALT_ROW if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            cell.width = col_w_ishi[c_idx]
            _set_cell_background(cell, bg)
            _set_cell_margins(cell, top=70, bottom=70, left=80, right=80)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in [0, 2] else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(val)
            run.font.name = "Calibri"
            run.font.size = Pt(8.5)
            if c_idx == 0:
                run.bold = True
                run.font.color.rgb = COLOR_UTN_NAVY

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 4.2 5 Whys para F2
    _add_section_heading(doc, "4.2 Árbol Causal Profundo de los 5 Porqués para Falla Crítica F2 (Motor 55 kW)", level=2)
    p_f2_whys = doc.add_paragraph()
    p_f2_whys.add_run(
        "Falla Crítica Evaluada: Cortocircuito y quemado de devanado estatórico en motor trifásico de 55 kW (42 h de parada, $109,000 USD).\n"
    ).bold = True

    f2_steps = [
        ("1. ¿Por qué ocurrió el cortocircuito entre espiras del estator?",
         "Porque la película dieléctrica de esmalte (Clase F) de los conductores de cobre se quemó al superar los 160 °C de temperatura interna.",
         "Causa Física Inmediata"),
        ("2. ¿Por qué se superaron los 160 °C en el interior del motor?",
         "Porque colapsó la disipación térmica: el ventilador forzado se trabó por suciedad y la carcasa quedó recubierta de una costra mineral aislante (k < 0.2 W/m·K).",
         "Causa Física / Mecánica"),
        ("3. ¿Por qué se bloquearon el ventilador y las aletas sin advertirse?",
         "Porque los operadores y mantenedores no verificaron la descarga de aire ni palparon la temperatura de la carcasa durante sus rondas.",
         "Causa Humana / Operativa"),
        ("4. ¿Por qué el personal de planta no detectó la acumulación de polvo?",
         "Porque no existía un estándar formal de Mantenimiento Autónomo (TPM CIL semanal) ni inspecciones de termografía infrarroja programadas en el sector.",
         "Causa de Método / Control"),
        ("5. ¿Por qué no estaban implementados el estándar TPM ni la protección térmica?",
         "Porque la gerencia mantenía una política puramente reactiva ('run-to-failure') en accionamientos, sin haber aplicado RCM que evidenciara que la falla de un motor de $4,000 USD costaría $105,000 USD de parada y riesgo de petrificación en el reactor.",
         "Causa Raíz Latente / Organizacional")
    ]

    tbl_f2 = doc.add_table(rows=6, cols=3)
    tbl_f2.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_f2.autofit = False
    _format_table_header(tbl_f2.rows[0], ["Nivel del Porqué", "Deducción Causal Técnica y Evidencia", "Nivel de Causa"])
    col_w_whys = [Inches(1.8), Inches(3.6), Inches(1.1)]

    for r_idx, (q, a, lvl) in enumerate(f2_steps):
        row = tbl_f2.rows[r_idx + 1]
        bg = HEX_ALT_ROW if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate([q, a, lvl]):
            cell = row.cells[c_idx]
            cell.width = col_w_whys[c_idx]
            _set_cell_background(cell, bg)
            _set_cell_margins(cell, top=60, bottom=60, left=70, right=70)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 2 else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(val)
            run.font.name = "Calibri"
            run.font.size = Pt(8.5)
            if c_idx == 0:
                run.bold = True
            elif c_idx == 2:
                run.bold = True
                run.font.color.rgb = COLOR_RED_ALERT if "Latente" in val else COLOR_UTN_NAVY

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    _add_callout_box(
        doc,
        "Acciones Preventivas de Bloqueo para Falla F2 (Poka-Yoke + TPM + Standby)",
        "1. Poka-Yoke / Físico: Instalar 3 sondas termométricas PT-100 embebidas en las cabezas de bobina del estator, vinculadas a alarma a 130 °C y disparo de corte a 145 °C en el arrancador suave.\n"
        "2. Procedimental / TPM: Implementar rutina semanal obligatoria de limpieza y soplado con aire comprimido seco de carcasa y tobera, con planilla firmada.\n"
        "3. Estratégico de Pañol: Almacenar 1 motor de reserva Standby de 55 kW para garantizar reemplazo en 4 h en caso de contingencia."
    )

    # 4.3 5 Whys para F5
    _add_section_heading(doc, "4.3 Árbol Causal Profundo de los 5 Porqués para Falla F5 (Rascador y Efecto Dominó)", level=2)
    p_f5_whys = doc.add_paragraph()
    p_f5_whys.add_run(
        "Falla Evaluada: Desgaste prematuro y atrancamiento de rascador primario con sobrecarga térmica en reductor (8 h, $20,850 USD).\n"
    ).bold = True

    f5_steps = [
        ("1. ¿Por qué sobrecalentó el reductor y aumentó la corriente del motor?",
         "Porque la cinta demandó una cupla resistente excesiva, forzando a los engranajes a disipar calor por fricción y al motor a absorber sobrecorriente.",
         "Causa Física Inmediata"),
        ("2. ¿Por qué se disparó la cupla resistente en la cinta?",
         "Porque la hoja de uretano del rascador primario se inclinó y trabó en ángulo agudo contra la goma, actuando como zapata de freno continua.",
         "Causa Física / Mecánica"),
        ("3. ¿Por qué se trabó la hoja del rascador?",
         "Porque sufrió desgaste asimétrico y la acumulación de costra de mineral en el soporte basculante bloqueó el resorte compensador de tensión.",
         "Causa Física / Componente"),
        ("4. ¿Por qué no se limpió la costra ni se ajustó el resorte?",
         "Porque los operadores de turno no realizaron la inspección sensorial diaria (ronda CIL de 3 minutos) para verificar el filo y la movilidad del eje.",
         "Causa Humana / Operativa"),
        ("5. ¿Por qué no existía la ronda de rascadores ni sensores térmicos?",
         "Porque la supervisión consideraba al rascador como un elemento secundario de limpieza sin impacto productivo, desconociendo el efecto dominó sobre reductor y motor; ausencia de Mantenimiento Autónomo estructurado.",
         "Causa Raíz Latente / Organizacional")
    ]

    tbl_f5 = doc.add_table(rows=6, cols=3)
    tbl_f5.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_f5.autofit = False
    _format_table_header(tbl_f5.rows[0], ["Nivel del Porqué", "Deducción Causal Técnica y Evidencia", "Nivel de Causa"])

    for r_idx, (q, a, lvl) in enumerate(f5_steps):
        row = tbl_f5.rows[r_idx + 1]
        bg = HEX_ALT_ROW if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate([q, a, lvl]):
            cell = row.cells[c_idx]
            cell.width = col_w_whys[c_idx]
            _set_cell_background(cell, bg)
            _set_cell_margins(cell, top=60, bottom=60, left=70, right=70)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 2 else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(val)
            run.font.name = "Calibri"
            run.font.size = Pt(8.5)
            if c_idx == 0:
                run.bold = True
            elif c_idx == 2:
                run.bold = True
                run.font.color.rgb = COLOR_RED_ALERT if "Latente" in val else COLOR_UTN_NAVY

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    _add_callout_box(
        doc,
        "Acciones Preventivas de Bloqueo para Falla F5 (TPM + Poka-Yoke Mecánico + PdM)",
        "1. Metodológica / TPM: Ronda obligatoria de Mantenimiento Autónomo (3 minutos al inicio de cada turno) para remoción de costra en eje porta-rascador y palpado de filo.\n"
        "2. Poka-Yoke / Mecánico: Instalar soporte basculante autolimpiante con fusible mecánico desacoplable ante sobreesfuerzo de frenado y deflector de goma anti-incrustaciones.\n"
        "3. Monitoreo Predictivo: Instalar sensor de temperatura PT-100 en el tapón de vaciado del reductor ortogonal conectado al PLC de planta con parada de seguridad a 85 °C."
    )

    # =========================================================================
    # SECCIÓN 5: CONCLUSIONES EJECUTIVAS Y HOJA DE RUTA
    # =========================================================================
    _add_section_heading(doc, "5. Conclusiones Ejecutivas y Hoja de Ruta de Mantenimiento", level=1)
    p_conc = doc.add_paragraph()
    p_conc.add_run(
        "1. La Línea de Transporte de Áridos de Alta Densidad presenta un comportamiento de fallas gobernado en un 100% "
        "por factores aleatorios e incondicionales a la edad (Patrones D, E y F de Nowlan & Heap). En consecuencia, "
        "las intervenciones intrusivas por horas de calendario deben ser sustituidas definitivamente por Mantenimiento Basado en la Condición (CBM).\n"
        "2. El costo dominante de las fallas no radica en los repuestos o mano de obra, sino en el lucro cesante por lucro cesante ($300,000 USD de $313,620 USD totales, 95.66%).\n"
        "3. La adquisición de un motor Standby de 55 kW en pañol ($4,000 USD de inversión) genera un ahorro neto de $95,000 USD "
        "en la primera falla evitada, reduciendo el MTTR global a 16.40 h y elevando la disponibilidad de planta al 98.86%, con un ROI del 2,275%.\n"
        "4. El Análisis de Causa Raíz mediante Ishikawa 6M y los 5 Porqués demostró que las causas raíz profundas son de carácter latente y organizativo: "
        "la implementación de Mantenimiento Autónomo (TPM) mediante rondas CIL de 3 minutos y protecciones Poka-Yoke blindan a la planta contra fallas en cadena."
    )

    # Guardar en memoria y retornar buffer
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


if __name__ == "__main__":
    # Genera el documento oficial en la raíz del proyecto TP N°1
    current_file = pathlib.Path(__file__).resolve()
    # Si report_builder.py está en TP N°1/scripts/, la raíz es parent.parent
    tp1_root = current_file.parent.parent if current_file.parent.name == "scripts" else current_file.parent
    output_docx = tp1_root / "TP1_Informe_Tecnico_Electromecanico_UTN.docx"
    
    print(f"Generando informe Word oficial en: {output_docx}")
    buf = build_technical_report()
    with open(output_docx, "wb") as f:
        f.write(buf.read())
    size_kb = output_docx.stat().st_size / 1024
    print(f"Documento generado con éxito: {output_docx.name} ({size_kb:.1f} KB)")
