# -*- coding: utf-8 -*-
"""
report_builder.py — Generador de Informe Técnico Oficial en Word (.docx)
Cátedra: Gestión y Mantenimiento Electromecánico — UTN FRSR (Año 2026)
Línea Principal de Elevación y Transporte de Áridos de Alta Densidad

Genera un documento Word (.docx) profesional con diseño institucional,
respondiendo y justificando íntegramente las 6 consignas del enunciado oficial
del TP N°1 y completando las Tablas 3 y 4 con datos sincronizados de la app.
"""

import io
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
HEX_NAVY = "1F4E79"
HEX_LIGHT_BG = "F2F5F9"
HEX_BORDER = "D9D9D9"
HEX_ALT_ROW = "F9FBFD"


def _set_cell_background(cell, hex_color):
    """Aplica color de fondo hexadecimal a una celda de tabla."""
    shading_xml = f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))


def _set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Configura márgenes internos (padding) en dxa (1 pt = 20 dxa)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def _format_table_header(row, titles, hex_bg=HEX_NAVY, font_size=9.5):
    """Estiliza la fila de encabezado de una tabla."""
    for idx, cell in enumerate(row.cells):
        _set_cell_background(cell, hex_bg)
        _set_cell_margins(cell, top=120, bottom=120, left=140, right=140)
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


def _add_section_heading(doc, text, level=1):
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


def _add_callout_box(doc, title, text, hex_bg=HEX_LIGHT_BG, hex_border=HEX_NAVY):
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


def build_technical_report(
    alumno: str = "Martín Méndez",
    legajo: str = "Legajo N° - UTN FRSR",
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
    Construye el documento Word del TP N°1 respetando íntegramente
    el enunciado, tablas y justificaciones técnicas de ingeniería electromecánica.
    """
    doc = Document()

    # --- Configuración de página (Márgenes estándar de 2.5 cm / ~1 pulgada) ---
    for section in doc.sections:
        section.top_margin = Inches(0.9)
        section.bottom_margin = Inches(0.9)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)

        # Encabezado institucional
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run("UTN FRSR · Gestión y Mantenimiento Electromecánico · TP N°1")
        hrun.font.name = "Calibri"
        hrun.font.size = Pt(8.5)
        hrun.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

        # Pie de página
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        frun = fp.add_run("Línea Principal de Elevación y Transporte de Áridos — Confiabilidad y Disponibilidad")
        frun.font.name = "Calibri"
        frun.font.size = Pt(8.5)
        frun.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    # =========================================================================
    # PORTADA INSTITUCIONAL
    # =========================================================================
    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst.paragraph_format.space_before = Pt(10)
    p_inst.paragraph_format.space_after = Pt(2)
    r1 = p_inst.add_run("UNIVERSIDAD TECNOLÓGICA NACIONAL\n")
    r1.bold = True
    r1.font.name = "Calibri"
    r1.font.size = Pt(16)
    r1.font.color.rgb = COLOR_UTN_NAVY

    r2 = p_inst.add_run("FACULTAD REGIONAL SAN RAFAEL\n")
    r2.bold = True
    r2.font.name = "Calibri"
    r2.font.size = Pt(13)
    r2.font.color.rgb = COLOR_GRAY_TEXT

    r3 = p_inst.add_run("Cátedra: GESTIÓN Y MANTENIMIENTO ELECTROMECÁNICO\nEspecialidad: Ingeniería Electromecánica")
    r3.font.name = "Calibri"
    r3.font.size = Pt(11)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Título del Trabajo Práctico
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(4)
    r_tp = p_title.add_run("TRABAJO PRÁCTICO N° 1\n")
    r_tp.bold = True
    r_tp.font.size = Pt(18)
    r_tp.font.color.rgb = COLOR_UTN_NAVY

    r_tema = p_title.add_run(
        "TEMA: INGENIERÍA DE CONFIABILIDAD, DIAGNÓSTICO DE PARÁMETROS OPERATIVOS, "
        "TAREAS DE INSPECCIÓN Y ESTRATEGIAS DE MANTENIMIENTO"
    )
    r_tema.bold = True
    r_tema.font.size = Pt(11.5)
    r_tema.font.color.rgb = COLOR_UTN_DARK

    # Bloque de datos del alumno
    p_alumno = doc.add_paragraph()
    p_alumno.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_alumno.paragraph_format.space_before = Pt(14)
    p_alumno.paragraph_format.space_after = Pt(20)
    r_al = p_alumno.add_run(f"Alumno(s) y Legajo: {alumno} — {legajo}\nCiclo Lectivo: {anio}")
    r_al.font.size = Pt(11)
    r_al.bold = True
    r_al.font.color.rgb = COLOR_GRAY_TEXT

    doc.add_page_break()

    # =========================================================================
    # SECCIÓN 1: CONTEXTO INDUSTRIAL
    # =========================================================================
    _add_section_heading(doc, "1. Contexto Industrial", level=1)
    p = doc.add_paragraph()
    p.add_run(
        f"El trabajo se desarrolla sobre la Línea Principal de Elevación y Transporte de Áridos de Alta Densidad "
        f"de un complejo industrial de procesamiento continuo. El sistema opera nominalmente 24 horas/día durante "
        f"300 días al año ({TOP:,.0f} horas/año). El flujo continuo abastece a un reactor principal donde la "
        f"detención imprevista de la alimentación provoca la petrificación de la masa en proceso, lo que genera "
        f"pérdidas catastróficas adicionales asociadas a tareas críticas de limpieza manual y reinicio de planta."
    )

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    p_comp = doc.add_paragraph()
    p_comp.add_run("Componentes Principales del Sistema Electromecánico:\n").bold = True
    bullet_points = [
        ("Unidad Motriz:", " Motor asincrónico trifásico de alta eficiencia, 55 kW, 4 polos, 380/660 V, 50 Hz, "
                           "aislamiento Clase F, grado de protección IP55, acoplado mediante arrancador suave electrónico."),
        ("Reducción de Velocidad:", " Reductor ortogonal de ejes paralelos de tres etapas, lubricado por barboteo "
                                    "con aceite sintético ISO VG 220, relación i = 28:1."),
        ("Polea Motriz y Apoyos:", " Polea de Ø500 mm con revestimiento cerámico vulcanizado, montada sobre ejes de acero "
                                   "SAE 4140 apoyados en dos soportes monobloc de pie (SNH) equipados con rodamientos "
                                   "oscilantes de doble hilera de rodillos (SKF 22220 EK) con fijación por manguito de apriete."),
        ("Sistema de Tensionado y Control:", " Centralita oleohidráulica con bomba de engranajes (20 bar), cilindro hidráulico "
                                             "de doble efecto para tensionado automático de banda, presostatos de seguridad e interbloqueo con el circuito del motor."),
        ("Limpieza y Protección:", " Rascadores primarios de uretano y secundario de hojas de metal duro, sensores inductivos "
                                   "de desalineación de banda e interruptores de parada de emergencia por cable tirador.")
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
    # SECCIÓN 2: DESAFÍO PRÁCTICO Y PROBLEMA DE INGENIERÍA
    # =========================================================================
    _add_section_heading(doc, "2. Desafío Práctico y Problema de Ingeniería", level=1)
    p_des = doc.add_paragraph()
    p_des.add_run(
        f"Se analizó el comportamiento del activo durante el último ciclo anual ({TOP:,.0f} h programadas), "
        f"en el cual la planta sufrió paradas por contingencias operativas imprevistas registradas en los siguientes eventos:"
    )

    # Tabla 1 Oficial del TP
    tbl1 = doc.add_table(rows=6, cols=5)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl1.autofit = False

    t1_headers = ["Evento", "Descripción de la Falla Detallada", "Tiempo Parada (h)", "Costo Repuestos (USD)", "Mano de Obra (HH)"]
    _format_table_header(tbl1.rows[0], t1_headers)

    t1_data = [
        ("F1", "Falla por picado/spalling en pista exterior del rodamiento oscilante lado acoplamiento debido a contaminación con material abrasivo y degradación de grasa. Se requirió desmontaje completo de la polea motriz.", "28", "$1,850", "48"),
        ("F2", "Sobrecalentamiento y cortocircuito entre espiras en el estator del motor eléctrico de 55 kW provocado por bloqueo del ventilador y acumulación de capa de polvo en la carcasa. Se requirió rebobinado completo.", "42", "$3,400", "24"),
        ("F3", "Fuga masiva de fluido hidráulico por fisura en latiguillo de alta presión debido a fatiga por pulsación, provocando caída de presión en el sistema de tensionado y disparo del interbloqueo.", "6", "$320", "12"),
        ("F4", "Desalineación severa de la banda por acumulación de material en tambor de reenvío, derivando en rasgado longitudinal de 15 metros de goma superior de la cinta. Se requirió vulcanizado en caliente por terceros.", "36", "$4,200", "36"),
        ("F5", "Desgaste prematuro y atrancamiento del rascador de uretano, causando fricción contra la banda y sobretemperatura en reductor por trabajo en sobrecarga.", "8", "$450", "16"),
    ]

    col_widths_t1 = [Inches(0.7), Inches(3.2), Inches(0.9), Inches(1.0), Inches(0.8)]
    for r_idx, row_data in enumerate(t1_data):
        row = tbl1.rows[r_idx + 1]
        bg = HEX_ALT_ROW if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.width = col_widths_t1[c_idx]
            _set_cell_background(cell, bg)
            _set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
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

    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    p_eco = doc.add_paragraph()
    p_eco.add_run("Parámetros Económicos de Planta:\n").bold = True
    p_eco.add_run(f"• Costo de Parada de Planta (Lucro cesante + Costos fijos): ${cost_parada:,.0f} USD/hora.\n")
    p_eco.add_run(f"• Costo Horario Promedio de Mano de Obra Propia de Mantenimiento: ${cost_hh:,.0f} USD/HH.")

    # =========================================================================
    # SECCIÓN 3: CONSIGNAS Y RESOLUCIÓN TÉCNICA
    # =========================================================================
    _add_section_heading(doc, "3. Consignas y Desarrollo Técnico", level=1)

    # -------------------------------------------------------------------------
    # Consigna 1
    # -------------------------------------------------------------------------
    _add_section_heading(doc, "Consigna 1: Jerarquía de Activos (ISO 14224) y Modelos de Mantenimiento", level=2)

    p_1a = doc.add_paragraph()
    p_1a.add_run("a) Arbolado Jerárquico del Sistema y Criterios de Decisión (Reemplazo vs. Reparación):\n").bold = True
    p_1a.add_run(
        "Conforme a los lineamientos de la norma internacional ISO 14224 (Petroleum, petrochemical and natural gas "
        "industries — Collection and exchange of reliability and maintenance data for equipment), la estructura jerárquica "
        "se establece en 6 niveles sucesivos para garantizar la trazabilidad operacional y financiera:\n"
    )

    arbol_txt = (
        "• Nivel 2 (Complejo Industrial): Planta de Procesamiento de Minerales / Áridos.\n"
        "• Nivel 3 (Área Operativa): Sección de Elevación, Transporte y Molienda Primaria.\n"
        "• Nivel 4 (Unidad de Proceso / Equipo Principal): Cinta Transportadora Principal de Áridos de Alta Densidad (CV-01).\n"
        "• Nivel 5 (Sistemas / Subsistemas Funcionales):\n"
        "   - SYS-MOT: Unidad Motriz y Accionamiento (Motor 55 kW, Reductor ortogonal i=28:1, Arrancador suave).\n"
        "   - SYS-BAN: Banda y Rodillos (Banda continua multicapa, tambor motriz Ø500mm, tambor de reenvío, estaciones de rodillos).\n"
        "   - SYS-TEN: Sistema Oleohidráulico de Tensionado Automático (Centralita 20 bar, bomba, cilindro doble efecto, latiguillos).\n"
        "   - SYS-EST: Estructura, Bastidor y Sistemas de Limpieza (Chasis estructural, rascador primario uretano y secundario).\n"
        "   - SYS-ELE: Sistema Eléctrico, Instrumentación y Seguridad (Presostatos, sensores desalineación, cables tirador).\n"
        "• Nivel 6 (Componente / Ítem Mantenible): Rodamiento oscilante SKF 22220 EK, devanado estatórico, latiguillo hidráulico, lámina rascadora.\n"
        "• Nivel 8 (Parte / Elemento Indivisible): Pistas de rodadura, rodillos esféricos, retenes de grasa, espiras aisladas, sellos O-ring."
    )
    p_arb = doc.add_paragraph()
    p_arb.paragraph_format.left_indent = Inches(0.2)
    p_arb.add_run(arbol_txt)

    _add_callout_box(
        doc,
        "Criterio Técnico: Decisiones de Reparación (MTBF) vs. Reemplazo (MTTF)",
        "• Nivel de Reparación (Nivel 4 Equipo y Nivel 5 Sistemas): Aplica a activos reparables complejos "
        "(motor eléctrico, reductor de velocidad, centralita hidráulica, tambor motriz). Se gestionan bajo la métrica "
        "MTBF (Mean Time Between Failures) e implican intervenciones de reacondicionamiento, rebobinado, mecanizado o ajuste in situ.\n"
        "• Nivel de Reemplazo (Nivel 6 Componentes descartables y Nivel 8 Elementos): Aplica a componentes unitarios indivisibles "
        "no reparables económicamente (rodamientos, latiguillos hidráulicos, labios de rascadores de uretano, retenes). "
        "Se gestionan bajo la métrica MTTF (Mean Time To Failure) mediante descarte y sustitución por repuesto nuevo."
    )

    p_1b = doc.add_paragraph()
    p_1b.add_run("b) Determinación del Nivel de Criticidad de los Subsistemas:\n").bold = True
    p_1b.add_run(
        "Se aplica la matriz de riesgo del TP mediante el cálculo del Factor de Consecuencia "
        "Ci = S + A + O + D y la Criticidad Total C = F × Ci, categorizando en:\n"
        "• Clase A - CRÍTICO (C ≥ 15): Asignar Modelo de Alta Disponibilidad.\n"
        "• Clase B - SEMICRÍTICO (10 ≤ C < 15): Asignar Modelo Sistemático / Predictivo.\n"
        "• Clase C - NO CRÍTICO (C < 10): Asignar Modelo Condicional / TPM.\n"
    )

    # Tabla de Criticidad
    tbl_crit = doc.add_table(rows=6, cols=9)
    tbl_crit.alignment = WD_TABLE_ALIGNMENT.CENTER
    crit_headers = ["Subsistema", "F", "S", "A", "O", "D", "Ci", "C", "Clase Asignada"]
    _format_table_header(tbl_crit.rows[0], crit_headers)

    crit_data = [
        ("Unidad Motriz", "3", "2", "1", "10", "4", "17", "51", "Clase A - Crítico"),
        ("Banda y Rodillos", "3", "2", "1", "10", "4", "17", "51", "Clase A - Crítico"),
        ("Sistema de Tensión", "2", "2", "2", "3", "1", "8", "16", "Clase A / B"),
        ("Estructura y Chasis", "2", "1", "1", "3", "1", "6", "12", "Clase B - Semicrítico"),
        ("Tablero y Control", "1", "2", "1", "3", "1", "7", "7", "Clase C - No Crítico"),
    ]
    if df_matriz is not None and not df_matriz.empty:
        crit_data = []
        for _, row in df_matriz.iterrows():
            sub = str(row.get("subsistema", ""))
            f_val = str(row.get("F", ""))
            s_val = str(row.get("S", ""))
            a_val = str(row.get("A", ""))
            o_val = str(row.get("O", ""))
            d_val = str(row.get("D", ""))
            ci_val = str(row.get("Ci", ""))
            c_val = str(row.get("C", ""))
            clase_val = str(row.get("clase", ""))
            crit_data.append((sub, f_val, s_val, a_val, o_val, d_val, ci_val, c_val, clase_val))

    for r_idx, row_vals in enumerate(crit_data):
        row = tbl_crit.rows[r_idx + 1]
        bg = HEX_ALT_ROW if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            _set_cell_background(cell, bg)
            _set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
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
    p_1c.add_run("c) Justificación de los Modelos de Mantenimiento Asignados:\n").bold = True
    justificaciones = [
        ("Unidad Motriz (Clase A — Crítico):", " Debido al impacto operacional catastrófico (O=10, petrificación de reactor) "
         "y costos de reparación elevados (D=4), se adopta el Modelo de Alta Disponibilidad. Requiere monitoreo predictivo "
         "continuo de vibraciones (aceleración y envolvente) en motor y reductor, termografía en carcasa y arrancador suave, "
         "más un plan de overhaul programado de rodamientos y engranajes."),
        ("Banda y Rodillos (Clase A — Crítico):", " La rotura longitudinal o trabamiento detiene la planta al 100% (O=10). "
         "Se asigna Modelo de Alta Disponibilidad con inspección ultrasónica de empalmes vulcanizados, monitoreo termográfico "
         "de rodillos de apoyo y control permanente de alineación mediante sensores inductivos."),
        ("Sistema de Tensión (Clase B — Semicrítico):", " Presenta fallas con impacto de parada parcial o contenida por "
         "presostatos (O=3, D=1). Corresponde un Modelo Sistemático / Predictivo: análisis semestral de aceite hidráulico "
         "(conteo de partículas ISO 4406), verificación mensual de presión de timbrado e inspección programada de mangueras."),
        ("Estructura, Chasis y Limpieza (Clase B/C):", " Su deterioro es de evolución lenta pero puede inducir fallas mayores. "
         "Se gestiona mediante Modelo Condicional y Mantenimiento Autónomo (TPM): el operador realiza inspección sensorial "
         "(CIL: Limpieza, Inspección, Lubricación) diaria de rascadores y remoción de áridos apelmazados."),
        ("Tablero y Control (Clase C — No Crítico):", " Tasa de falla muy baja (F=1) y costos reducidos. Modelo Condicional / TPM: "
         "limpieza periódica de filtros de aire en tableros, termografía semestral en contactores y bornes, y atención ante eventos.")
    ]
    for tit, texto in justificaciones:
        p_j = doc.add_paragraph(style='List Bullet')
        p_j.paragraph_format.space_before = Pt(1)
        p_j.paragraph_format.space_after = Pt(2)
        r_tit = p_j.add_run(tit)
        r_tit.bold = True
        p_j.add_run(texto)

    # -------------------------------------------------------------------------
    # Consigna 2
    # -------------------------------------------------------------------------
    _add_section_heading(doc, "Consigna 2: Plan de Tareas de Inspección (Sensoriales e Instrumentales)", level=2)
    p_ins = doc.add_paragraph()
    p_ins.add_run(
        "Se diseña la Ruta de Inspección Rutinaria del sistema, separando claramente las inspecciones sensoriales en "
        "marcha (Mantenimiento Autónomo / Operador) de las instrumentales avanzadas (Mantenimiento Predictivo / Especialista):"
    )

    tbl_insp = doc.add_table(rows=7, cols=5)
    tbl_insp.alignment = WD_TABLE_ALIGNMENT.CENTER
    insp_headers = ["Tipo Inspección", "Componente / Punto", "Técnica / Sentido", "Parámetro de Control", "Frecuencia"]
    _format_table_header(tbl_insp.rows[0], insp_headers)

    insp_data = [
        ("Sensorial (TPM)", "Rascadores de uretano", "Visual / Palpado", "Desgaste de filos, tensión contra banda, atascos", "Diaria (por turno)"),
        ("Sensorial (TPM)", "Tambor de reenvío y polea", "Visual / Auditivo", "Acumulación de áridos, ruidos metálicos de raspado", "Diaria"),
        ("Sensorial (TPM)", "Central hidráulica y latiguillos", "Visual / Olfativo", "Fugas de fluido, nivel de tanque, pulsaciones", "Diaria"),
        ("Instrumental (PdM)", "Rodamientos SKF 22220 EK", "Vibrómetro / Ultrasonido", "Valores globales RMS (ISO 10816-3) y demodulación", "Quincenal"),
        ("Instrumental (PdM)", "Motor eléctrico 55 kW", "Termografía infrarroja", "Gradiente térmico en carcasa y bornera (< 75 °C)", "Mensual"),
        ("Instrumental (PdM)", "Reductor ortogonal i=28:1", "Análisis de aceite de laboratorio", "Viscosidad cSt a 40°C, agua ppm, desgaste Fe/Cu", "Trimestral"),
    ]
    for r_idx, row_vals in enumerate(insp_data):
        row = tbl_insp.rows[r_idx + 1]
        bg = HEX_ALT_ROW if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            _set_cell_background(cell, bg)
            _set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
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

    # -------------------------------------------------------------------------
    # Consigna 3
    # -------------------------------------------------------------------------
    _add_section_heading(doc, "Consigna 3: Análisis de Tipos de Mantenimiento Proactivo", level=2)
    p_3 = doc.add_paragraph()
    p_3.add_run(
        "Se evalúa la estrategia de mantenimiento proactivo (Preventivo, Predictivo o TPM) que debió haberse aplicado "
        "para evitar la ocurrencia de cada una de las 5 fallas registradas, completando la Tabla 3 oficial del TP:"
    )

    tbl3 = doc.add_table(rows=6, cols=5)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    t3_headers = ["Evento", "Componente Afectado", "Tipo de Mantenimiento Propuesto", "Tarea de Inspección o Técnica Asociada", "Frecuencia de Aplicación"]
    _format_table_header(tbl3.rows[0], t3_headers)

    t3_data = [
        ("F1", "Rodamiento oscilante SKF 22220 EK (Polea motriz)", "Mantenimiento Predictivo (CBM)",
         "Análisis de vibraciones (envolvente/demodulación) y lubricación por ultrasonido para evitar sobre-engrase y degradación de grasa.", "Quincenal"),
        ("F2", "Motor asincrónico 55 kW (Bobinado estatórico)", "Mantenimiento Preventivo / TPM",
         "Limpieza periódica de cubierta de ventilador y carcasa (TPM) + Termografía infrarroja y ensayo de resistencia de aislación (Megado).", "Semanal (limpieza) / Mensual (PdM)"),
        ("F3", "Latiguillo de alta presión (Central oleohidráulica)", "Mantenimiento Preventivo Sistemático",
         "Inspección visual de fatiga por flexión/pulsación y sustitución programada por horas de servicio (vida límite calendario de mangueras).", "Mensual / Reemplazo a 2 años"),
        ("F4", "Banda de transporte (Cubierta de caucho superior)", "Mantenimiento Autónomo (TPM) y Predictivo",
         "Limpieza de tambor de reenvío por rascador interior + Verificación funcional de sensores inductivos de desalineación e interbloqueo.", "Diaria (por turno)"),
        ("F5", "Rascador primario de uretano", "Mantenimiento Autónomo (TPM)",
         "Inspección diaria de desgaste de labio, ajuste de tensión mediante resorte/contrapeso y remoción de áridos incrustados.", "Diaria (cambio de turno)")
    ]
    for r_idx, row_vals in enumerate(t3_data):
        row = tbl3.rows[r_idx + 1]
        bg = HEX_ALT_ROW if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            _set_cell_background(cell, bg)
            _set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
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

    # -------------------------------------------------------------------------
    # Consigna 4
    # -------------------------------------------------------------------------
    _add_section_heading(doc, "Consigna 4: Determinación de KPIs y Análisis de Confiabilidad", level=2)

    p_4a = doc.add_paragraph()
    p_4a.add_run("a) Memoria de Cálculo de KPIs (MTBF, MTTR, Ai, Ao):\n").bold = True

    nom = kpis_nominal or {"ttp": 120.0, "tfr": 7080.0, "mtbf": 1416.0, "mttr": 24.0, "ai": 98.33, "ao": 98.33, "n": 5}
    rep = kpis_rep or {"ttp": 120.0, "tfr": 7080.0, "mtbf": 786.67, "mttr": 13.33, "ai": 98.33, "ao": 98.33, "n": 9}

    kpi_txt = (
        f"1. Tiempo Total de Paradas por Falla (TTP): Sumatoria de paradas = {nom['ttp']:.1f} h.\n"
        f"2. Tiempo de Funcionamiento Real (TFR): TOP - TTP = {TOP:,.0f} h - {nom['ttp']:.1f} h = {nom['tfr']:,.1f} h.\n"
        f"3. Análisis Comparativo de Escenarios Operativos:\n"
        f"   • Caso Base Nominal (N = {nom['n']} eventos consolidados):\n"
        f"     - MTBF = TFR / N = {nom['tfr']:,.1f} / {nom['n']} = {nom['mtbf']:,.2f} h.\n"
        f"     - MTTR = TTP / N = {nom['ttp']:.1f} / {nom['n']} = {nom['mttr']:,.2f} h.\n"
        f"     - Disponibilidad Inherente (Ai) = MTBF / (MTBF + MTTR) = {nom['ai']:.2f}%.\n"
        f"     - Disponibilidad Operacional (Ao) = TFR / TOP = {nom['ao']:.2f}%.\n"
        f"   • Caso Refinado con Repetición Estimada (N_est ≈ {rep['n']} fallas desagregadas):\n"
        f"     - MTBF = TFR / N_est = {rep['tfr']:,.1f} / {rep['n']} = {rep['mtbf']:,.2f} h.\n"
        f"     - MTTR = TTP / N_est = {rep['ttp']:.1f} / {rep['n']} = {rep['mttr']:,.2f} h.\n"
        f"     - Ai = {rep['ai']:.2f}% | Ao = {rep['ao']:.2f}% (Idénticas ya que TTP y TFR son constantes anuales)."
    )
    p_kpis = doc.add_paragraph()
    p_kpis.paragraph_format.left_indent = Inches(0.2)
    p_kpis.add_run(kpi_txt)

    _add_callout_box(
        doc,
        "Distinción Conceptual: MTBF (Sistemas Reparables) vs. MTTF (Componentes Descartables)",
        "El indicador MTBF (Mean Time Between Failures) mide la confiabilidad de sistemas e instalaciones reparables "
        "(cinta completa, motorreductor, unidad de accionamiento), donde el equipo vuelve a servicio tras la intervención. "
        "En contraste, el MTTF (Mean Time To Failure) aplica estrictamente a componentes unitarios no reparables "
        "(rodamientos oscilantes reemplazados, fusibles, sellos hidráulicos, tramos de manguera desechados), cuya vida "
        "útil concluye definitivamente al ocurrir la primera falla."
    )

    p_4b = doc.add_paragraph()
    p_4b.add_run("b) Análisis de Patrones de Falla de Nowlan & Heap:\n").bold = True
    p_4b.add_run(
        "¿Sigue este activo la clásica 'curva de la bañera'?\n"
        "RESPUESTA: NO. La investigación fundamental de Nowlan & Heap (1978) para la industria aeronáutica (DoD/FAA) "
        "demostró que la curva de la bañera tradicional (Patrón A: alta mortalidad infantil, tasa constante y fuerte "
        "zona de desgaste final por edad) solo representa aproximadamente el 4% de las fallas en activos industriales complejos.\n\n"
        "En la Línea de Transporte de Áridos, el 100% de las fallas registradas responden a los Patrones D, E o F (Fallas "
        "Aleatorias e Incondicionales a la Edad):\n"
        "• Falla F1 (Rodamiento): Ocurrió por ingreso prematuro de polvo abrasivo debido a daño en el retén y degradación de grasa. "
        "No fue por agotamiento de la vida nominal a la fatiga L10 (Patrón E / F).\n"
        "• Falla F2 (Motor eléctrico): El cortocircuito se desencadenó por ensuciamiento y bloqueo del flujo aerodinámico externo, "
        "no por degradación dieléctrica por envejecimiento natural de 20 años (Patrón E).\n"
        "• Falla F3 (Latiguillo hidráulico): Rotura por picos de fatiga por pulsación y roce superficial (Patrón D/E).\n"
        "• Falla F4 (Cinta desgarrada): Daño aleatorio externo por colmatación mecánica en el tambor de reenvío.\n\n"
        "Conclusión para la Gestión: Sustituir componentes basándose puramente en horas de calendario u 'overhaul por horas' "
        "no previene este tipo de fallas y puede incluso inducir mortalidad infantil por errores de montaje (Patrón F). "
        "La estrategia técnica correcta es el Mantenimiento Basado en la Condición (CBM) mediante monitoreo instrumental "
        "y el Mantenimiento Autónomo (TPM) enfocado en la limpieza y preservación básica."
    )

    # -------------------------------------------------------------------------
    # Consigna 5
    # -------------------------------------------------------------------------
    _add_section_heading(doc, "Consigna 5: Evaluación de Impacto Financiero y Estrategia de Mitigación", level=2)

    p_5a = doc.add_paragraph()
    p_5a.add_run("a) Balance Económico Anual de las Fallas:\n").bold = True
    p_5a.add_run(
        "Se cuantifica el impacto real total mediante la ecuación de costo directo:\n"
        "Ctotal = Cparada + Crepuestos + CMO\n"
        f"Donde Cparada = TTP (h) × ${cost_parada:,.0f} USD/h y CMO = HH × ${cost_hh:,.0f} USD/HH."
    )

    tbl4 = doc.add_table(rows=7, cols=7)
    tbl4.alignment = WD_TABLE_ALIGNMENT.CENTER
    t4_headers = ["Evento", "TTP (h)", "Costo Parada (USD)", "Repuestos (USD)", "HH", "Costo MO (USD)", "Costo Total Evento (USD)"]
    _format_table_header(tbl4.rows[0], t4_headers)

    t4_data = [
        ("F1", "28", "$70,000", "$1,850", "48", "$1,200", "$73,050"),
        ("F2", "42", "$105,000", "$3,400", "24", "$600", "$109,000"),
        ("F3", "6", "$15,000", "$320", "12", "$300", "$15,620"),
        ("F4", "36", "$90,000", "$4,200", "36", "$900", "$95,100"),
        ("F5", "8", "$20,000", "$450", "16", "$400", "$20,850"),
        ("TOTAL", "120", "$300,000", "$10,220", "136", "$3,400", "$313,620")
    ]

    if df_fallas is not None and not df_fallas.empty and "costo_total" in df_fallas.columns:
        t4_data = []
        tot_ttp = df_fallas["ttp"].sum()
        tot_cpar = df_fallas["costo_parada"].sum()
        tot_rep = df_fallas["repuestos"].sum()
        tot_hh = df_fallas["hh"].sum()
        tot_cmo = df_fallas["costo_mo"].sum()
        tot_ctot = df_fallas["costo_total"].sum()
        for _, r in df_fallas.iterrows():
            t4_data.append((
                str(r.get("evento", "")),
                f"{r.get('ttp', 0):.0f}",
                f"${r.get('costo_parada', 0):,.0f}",
                f"${r.get('repuestos', 0):,.0f}",
                f"{r.get('hh', 0):.0f}",
                f"${r.get('costo_mo', 0):,.0f}",
                f"${r.get('costo_total', 0):,.0f}"
            ))
        t4_data.append((
            "TOTAL",
            f"{tot_ttp:.0f}",
            f"${tot_cpar:,.0f}",
            f"${tot_rep:,.0f}",
            f"{tot_hh:.0f}",
            f"${tot_cmo:,.0f}",
            f"${tot_ctot:,.0f}"
        ))

    for r_idx, row_vals in enumerate(t4_data):
        row = tbl4.rows[r_idx + 1]
        is_total = (r_idx == len(t4_data) - 1)
        bg = HEX_LIGHT_BG if is_total else (HEX_ALT_ROW if r_idx % 2 == 1 else "FFFFFF")
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            _set_cell_background(cell, bg)
            _set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in [0, 1, 4] else WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(val)
            run.font.name = "Calibri"
            run.font.size = Pt(8.5)
            if is_total or c_idx == 0:
                run.bold = True
            if is_total:
                run.font.color.rgb = COLOR_UTN_NAVY

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    p_5b = doc.add_paragraph()
    p_5b.add_run("b) Demostración de Falla Crítica y Estrategia de Motor Standby en Almacén:\n").bold = True
    p_5b.add_run(
        "1. Demostración Numérica: La falla de mayor impacto económico fue el Evento F2 (Cortocircuito en motor eléctrico "
        "de 55 kW), con un costo total de $109,000 USD (34.75% de las pérdidas anuales de la planta), seguida muy de cerca "
        "por la F4 (Rotura de banda, $95,100 USD / 30.32%).\n\n"
        "2. Justificación de la Estrategia de Motor Standby (Reserva en Almacén):\n"
        "• Diagnóstico de la situación histórica: Las 42 horas de MTTR del motor se debieron principalmente a retrasos "
        "logísticos externos: desmontaje, transporte a taller especializado, rebobinado integral de bobinas estatóricas, "
        "impregnación por inmersión y secado en horno, reensamblaje y pruebas de aislamiento.\n"
        "• Propuesta técnica: Adquirir y almacenar en pañol de planta un motor idéntico en reserva Standby (55 kW, 4 polos, "
        "380/660 V, carcasa normalizada, patas estandarizadas).\n"
        "• Impacto en el MTTR: Ante una contingencia, el protocolo de mantenimiento no espera el rebobinado, sino que ejecuta "
        "la sustitución electromecánica directa (izaje, desacople mecánico, alineación láser y conexión eléctrica al arrancador "
        "suave), reduciendo el tiempo de parada de 42 horas a solo 4 horas.\n"
        "• Análisis Costo-Beneficio (ROI):\n"
        "   - Costo de parada con motor en reserva: 4 h × $2,500 USD/h = $10,000 USD.\n"
        "   - Ahorro inmediato por lucro cesante: $105,000 USD - $10,000 USD = $95,000 USD de ahorro neto en un solo evento.\n"
        "   - Costo de adquisición de motor Standby nuevo: ~$3,500 - $4,500 USD.\n"
        "   - Conclusión: La inversión en el motor de repuesto se amortiza más de 20 veces en la primera falla evitada."
    )

    # -------------------------------------------------------------------------
    # Consigna 6
    # -------------------------------------------------------------------------
    _add_section_heading(doc, "Consigna 6: Hipótesis Diagnóstica de Fallas Combinadas (Efecto Dominó en F5)", level=2)
    p_6 = doc.add_paragraph()
    p_6.add_run(
        "Se analiza la Falla F5 (Atrancamiento de rascador de uretano), modelando la secuencia electromecánica en cadena "
        "y cómo el Mantenimiento Autónomo (TPM) detiene este efecto dominó:"
    )

    efecto_txt = (
        "1. Eslabón Mecánico Primario: El rascador primario sufre desgaste acelerado no uniforme por contacto con mineral abrasivo. "
        "La pérdida de perfil y la acumulación de finos provoca que el soporte metálico pierda elasticidad y se trabe contra la banda.\n"
        "2. Aumento de Fuerza de Fricción: La fuerza tangencial de rozamiento contra la cara de transporte se dispara drásticamente "
        "(Fr = μ × Fn ↑), convirtiéndose en un freno continuo sobre la polea motriz.\n"
        "3. Reacción Electromecánica en el Motor Asincrónico:\n"
        "   - Para sostener la velocidad de avance, el motor debe suministrar una cupla mecánica resistente significativamente mayor (Tres ↑).\n"
        "   - En una máquina de inducción, la cupla es proporcional al deslizamiento: s = (ns - n)/ns ↑.\n"
        "   - Al aumentar el deslizamiento, la fuerza contraelectromotriz interna se reduce, exigiendo un mayor consumo de corriente "
        "estatórica de línea (I1 ≈ k · Tres ↑), lo que satura térmicamente el devanado por pérdidas Joule (P = 3 · I² · R).\n"
        "4. Consecuencia Térmica en el Reductor Ortogonal:\n"
        "   - El par torsional elevado se transmite a través del reductor de 3 etapas (i=28:1). La fricción en los engranajes "
        "incrementa la generación de calor en los dientes.\n"
        "   - La temperatura del baño de aceite sintético ISO VG 220 se eleva por encima de los límites de diseño (> 95 °C).\n"
        "   - A alta temperatura, la viscosidad cinemática del aceite disminuye sensiblemente, lo que reduce el espesor de la película "
        "elastohidrodinámica (EHL). Esto provoca contacto metal-metal entre flancos de dientes, acelerando el micropitting y "
        "la oxidación prematura del lubricante."
    )
    p_dom = doc.add_paragraph()
    p_dom.paragraph_format.left_indent = Inches(0.2)
    p_dom.add_run(efecto_txt)

    _add_callout_box(
        doc,
        "Interrupción del Efecto Dominó mediante Rutina Diaria de TPM (Mantenimiento Autónomo)",
        "Una Rutina de Inspección Diaria (inspección de 3 minutos al inicio de cada turno realizada por el operador del sector) "
        "detecta el labio desgastado, la acumulación de costra de mineral o la pérdida de alineación del rascador en el Eslabón 1. "
        "El operador limpia la acumulación y ajusta la tensión del resorte de compensación antes de que se produzca el trabamiento "
        "mecánico. De esta forma, se cancela la sobrecarga mecánica, evitando el pico de corriente en el motor y la sobretemperatura "
        "en el aceite del reductor."
    )

    # =========================================================================
    # ANEXO: FORMULACIÓN DE KPIs DE CONFIABILIDAD
    # =========================================================================
    _add_section_heading(doc, "Anexo: Formulación Teórica Oficial de KPIs (Cátedra UTN)", level=1)
    p_an = doc.add_paragraph()
    p_an.add_run(
        "A. MTBF (Mean Time Between Failures / Tiempo Medio Entre Fallas):\n"
        "Métrica de confiabilidad para sistemas reparables. Mide el tiempo promedio de operación continua sin paradas no programadas:\n"
        "MTBF = TFR / N = (TOP - TTP) / N\n\n"
        "B. MTTR (Mean Time To Repair / Tiempo Medio Para Reparar):\n"
        "Métrica de mantenibilidad y eficiencia técnica/logística. Mide el tiempo promedio consumido en diagnosticar, "
        "reparar, probar y habilitar el activo tras una falla:\n"
        "MTTR = TTP / N\n\n"
        "C. Disponibilidad Inherente (Ai) vs. Disponibilidad Operacional (Ao):\n"
        "• Disponibilidad Inherente (Ai): Considera exclusivamente el tiempo de mantenimiento correctivo bajo condiciones ideales "
        "de diseño y logística perfecta:\n"
        "Ai = MTBF / (MTBF + MTTR)\n"
        "• Disponibilidad Operacional (Ao): Métrica real de planta que refleja la fracción de tiempo que el sistema estuvo produciendo "
        "respecto al tiempo programado:\n"
        "Ao = TFR / TOP\n"
        "(Nota: Cuando no existen paradas programadas por preventivo en el cómputo del ciclo, Ai = Ao matemáticamente)."
    )

    # Guardar en memoria y retornar buffer
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


if __name__ == "__main__":
    buf = build_technical_report()
    with open("TP1_Informe_Tecnico_Oficial_UTN.docx", "wb") as f:
        f.write(buf.read())
    print("Documento de prueba generado con éxito: TP1_Informe_Tecnico_Oficial_UTN.docx")
