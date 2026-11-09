# -*- coding: utf-8 -*-
"""
Script para generar el documento oficial de entrega de la Actividad 3:
Examen de Conceptualización Básica - Lenguaje DronLang
Asignatura: Compiladores - Docente: Ing. Hasler Gamboa Berrío
Autores: Jainer Andrés Morales Cuesta & Sharly Andrés Mosquera Rodríguez
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    """Aplica color de fondo a una celda."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=120, bottom=120, left=180, right=180):
    """Aplica márgenes internos (padding) a una celda en dxa."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def set_cell_border(cell, **kwargs):
    """
    Configura bordes de celda.
    kwargs: top, bottom, left, right con dict(sz=12, val='single', color='1B365D')
    """
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
    for edge, border_args in kwargs.items():
        sz = border_args.get('sz', 4)
        val = border_args.get('val', 'single')
        color = border_args.get('color', 'CCCCCC')
        b = parse_xml(f'<w:{edge} {nsdecls("w")} w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>')
        tcBorders.append(b)
    tcPr.append(tcBorders)

def add_code_block(doc, code_text):
    """Crea una caja de código sombreada con fuente monoespaciada."""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    cell = table.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, "F8F9FA")
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)
    set_cell_border(cell, 
                    left={'sz': 16, 'val': 'single', 'color': '1B365D'},
                    top={'sz': 4, 'val': 'single', 'color': 'E2E8F0'},
                    right={'sz': 4, 'val': 'single', 'color': 'E2E8F0'},
                    bottom={'sz': 4, 'val': 'single', 'color': 'E2E8F0'})
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.15
    
    run = p.add_run(code_text.strip())
    run.font.name = 'Consolas'
    run.font.size = Pt(9.5)
    run.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)
    
    # Espacio después de la caja
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(4)

def build_document():
    doc = docx.Document()
    
    # Configurar márgenes estándar (2.5 cm = 0.984 in)
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
        # Header
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run("Compiladores | DronLang - Actividad 3: Conceptualización Básica")
        hrun.font.name = 'Calibri'
        hrun.font.size = Pt(8.5)
        hrun.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)
        
        # Footer
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        frun = fp.add_run("Página ")
        frun.font.name = 'Calibri'
        frun.font.size = Pt(9)
        frun.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)
        # Añadir campo de página en Word XML
        fldSimple = OxmlElement('w:fldSimple')
        fldSimple.set(qn('w:instr'), 'PAGE')
        fp._p.append(fldSimple)
        frun2 = fp.add_run("  •  Universidad - Programa de Ingeniería  •  2026")
        frun2.font.name = 'Calibri'
        frun2.font.size = Pt(9)
        frun2.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)

    # Colores corporativos del documento
    COLOR_PRIMARY = RGBColor(0x1B, 0x36, 0x5D)    # Azul marino institucional
    COLOR_SECONDARY = RGBColor(0x2E, 0x5B, 0x88)  # Azul medio
    COLOR_TEXT = RGBColor(0x2D, 0x37, 0x48)       # Gris carbón para texto
    COLOR_MUTED = RGBColor(0x64, 0x74, 0x8B)      # Gris secundario

    # =========================================================================
    # PORTADA / ENCABEZADO INSTITUCIONAL
    # =========================================================================
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(4)
    run_sub = p_title.add_run("ASIGNATURA: COMPILADORES\n")
    run_sub.font.name = 'Calibri'
    run_sub.font.size = Pt(12)
    run_sub.font.bold = True
    run_sub.font.color.rgb = COLOR_SECONDARY
    
    run_main = p_title.add_run("ACTIVIDAD 3: EXAMEN DE CONCEPTUALIZACIÓN BÁSICA\n")
    run_main.font.name = 'Calibri'
    run_main.font.size = Pt(18)
    run_main.font.bold = True
    run_main.font.color.rgb = COLOR_PRIMARY
    
    run_desc = p_title.add_run("Diseño y Especificación del Lenguaje Propio: DronLang (Dron de Reparto)")
    run_desc.font.name = 'Calibri'
    run_desc.font.size = Pt(13)
    run_desc.font.italic = True
    run_desc.font.color.rgb = COLOR_TEXT

    # Tabla de datos de entrega
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    meta_table = doc.add_table(rows=5, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_data = [
        ("Docente Evaluador:", "Ing. Hasler Gamboa Berrío"),
        ("Temática Asignada:", "Dron de Reparto"),
        ("Nombre del Lenguaje:", "DronLang (.dron)"),
        ("Integrantes del Equipo:", "Jainer Andrés Morales Cuesta\nSharly Andrés Mosquera Rodríguez"),
        ("Repositorio Oficial GitHub:", "https://github.com/andrescrack2024/Compilador")
    ]
    
    for i, (label, val) in enumerate(meta_data):
        row = meta_table.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.2)
        c1.width = Inches(4.3)
        set_cell_background(c0, "F0F4F8")
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c0, 80, 80, 120, 120)
        set_cell_margins(c1, 80, 80, 120, 120)
        set_cell_border(c0, bottom={'sz': 4, 'val': 'single', 'color': 'CBD5E1'})
        set_cell_border(c1, bottom={'sz': 4, 'val': 'single', 'color': 'CBD5E1'})
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(label)
        r0.font.name = 'Calibri'
        r0.font.bold = True
        r0.font.size = Pt(10)
        r0.font.color.rgb = COLOR_PRIMARY
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(val)
        r1.font.name = 'Calibri'
        r1.font.size = Pt(10)
        r1.font.color.rgb = COLOR_TEXT

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(12)
    p_div.paragraph_format.space_after = Pt(14)
    r_line = p_div.add_run("—" * 55)
    r_line.font.color.rgb = RGBColor(0xCB, 0xD5, 0xE1)
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # =========================================================================
    # PUNTO 1: QUÉ HACE CADA INSTRUCCIÓN Y CÓMO SE ESCRIBE
    # =========================================================================
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(6)
    r_h1 = h1.add_run("1. Catálogo de Instrucciones de DronLang (¿Qué hace y cómo se escribe?)")
    r_h1.font.name = 'Calibri'
    r_h1.font.size = Pt(14)
    r_h1.font.bold = True
    r_h1.font.color.rgb = COLOR_PRIMARY

    p_intro1 = doc.add_paragraph()
    p_intro1.paragraph_format.space_after = Pt(8)
    p_intro1.paragraph_format.line_spacing = 1.15
    r_intro1 = p_intro1.add_run(
        "El lenguaje DronLang fue diseñado para la planificación, telemetría y ejecución de misiones autónomas de un dron de reparto urbano. "
        "Adopta una sintaxis clara y minimalista inspirada en Python: no utiliza puntos y coma (;) para finalizar sentencias, "
        "cada instrucción se delimita por saltos de línea y las estructuras de control se cierran de forma determinista mediante la palabra clave "
    )
    r_intro1.font.name = 'Calibri'
    r_intro1.font.size = Pt(10.5)
    r_intro1.font.color.rgb = COLOR_TEXT
    r_fin = p_intro1.add_run("fin")
    r_fin.font.bold = True
    r_fin.font.name = 'Consolas'

    # Tabla de especificación de instrucciones
    inst_table = doc.add_table(rows=8, cols=4)
    inst_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["Instrucción", "Sintaxis Formal", "¿Qué hace en el Dron? (Semántica)", "Ejemplo Práctico"]
    widths = [Inches(1.2), Inches(1.8), Inches(2.2), Inches(1.3)]

    # Fila de cabecera
    for col_idx, (text, width) in enumerate(zip(headers, widths)):
        cell = inst_table.cell(0, col_idx)
        cell.width = width
        set_cell_background(cell, "1B365D")
        set_cell_margins(cell, 120, 120, 100, 100)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.bold = True
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    instructions_data = [
        ("despegar", "despegar", "Enciende motores y realiza ascenso vertical automático hasta alcanzar la altitud base de crucero. No requiere argumentos.", "despegar"),
        ("volar", "volar <expresión>", "Desplaza horizontalmente el dron en línea recta hacia el punto objetivo la cantidad de metros calculada (requiere valor >= 0).", "volar 15\nvolar distancia * 2"),
        ("entregar", "entregar <expresión>", "Acciona el servomecanismo de descarga para depositar en el destino el número de paquetes indicado (requiere valor >= 0).", "entregar 1\nentregar paquetes"),
        ("aterrizar", "aterrizar", "Inicia descenso vertical controlado en la zona de aterrizaje actual y corta la alimentación de los motores.", "aterrizar"),
        ("Asignación", "<id> = <expresión>", "Almacena en memoria un valor numérico o el resultado de evaluar una expresión aritmética para su posterior utilización.", "bateria = 100\nresto = bateria - 15"),
        ("repetir", "repetir <expr> veces\n    <bloque>\nfin", "Estructura iterativa determinista. Ejecuta el bloque interno la cantidad de veces resultante de la expresión (entero >= 0).", "repetir 3 veces\n    volar 10\n    entregar 1\nfin"),
        ("si ... sino", "si <cond> entonces\n    <bloque_si>\nsino\n    <bloque_sino>\nfin", "Estructura de bifurcación condicional. Evalúa una comparación binaria; ejecuta el primer bloque si es verdadera o la alternativa si no.", "si bateria > 20 entonces\n    volar 30\nsino\n    aterrizar\nfin")
    ]

    for row_idx, (name, syntax, desc, ex) in enumerate(instructions_data, start=1):
        row = inst_table.rows[row_idx]
        bg_col = "FFFFFF" if row_idx % 2 != 0 else "F8FAFC"
        for col_idx, text in enumerate([name, syntax, desc, ex]):
            cell = row.cells[col_idx]
            cell.width = widths[col_idx]
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, 100, 100, 100, 100)
            set_cell_border(cell, 
                            bottom={'sz': 4, 'val': 'single', 'color': 'E2E8F0'},
                            top={'sz': 4, 'val': 'single', 'color': 'E2E8F0'})
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.1
            run = p.add_run(text)
            run.font.name = 'Consolas' if col_idx in (0, 1, 3) else 'Calibri'
            run.font.size = Pt(8.5) if col_idx in (1, 3) else Pt(9)
            run.font.bold = (col_idx == 0)
            run.font.color.rgb = COLOR_PRIMARY if col_idx == 0 else COLOR_TEXT

    # Párrafo de operadores aceptados
    p_ops = doc.add_paragraph()
    p_ops.paragraph_format.space_before = Pt(10)
    p_ops.paragraph_format.space_after = Pt(12)
    p_ops.paragraph_format.line_spacing = 1.15
    r_op_title = p_ops.add_run("Conjunto de Operadores y Delimitadores Soportados:\n")
    r_op_title.font.name = 'Calibri'
    r_op_title.font.bold = True
    r_op_title.font.size = Pt(10.5)
    r_op_title.font.color.rgb = COLOR_PRIMARY
    
    ops_detail = (
        "• Aritméticos: Suma (+), Resta (-), Multiplicación (*), con precedencia matemática estándar.\n"
        "• Comparación Relacional: Mayor que (>), Menor que (<), Igualdad estricta (==).\n"
        "• Delimitadores y Agrupación: Paréntesis '(' y ')' para alterar precedencias de evaluación.\n"
        "• Comentarios en código: Estilo una línea con '#' (estilo Python) o '//' (estilo C/JavaScript)."
    )
    r_ops_desc = p_ops.add_run(ops_detail)
    r_ops_desc.font.name = 'Calibri'
    r_ops_desc.font.size = Pt(9.5)
    r_ops_desc.font.color.rgb = COLOR_TEXT

    # =========================================================================
    # PUNTO 2: DIBUJO DE LAS FASES DEL COMPILADOR Y UNA FRASE POR FASE
    # =========================================================================
    doc.add_page_break()
    
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(8)
    h2.paragraph_format.space_after = Pt(6)
    r_h2 = h2.add_run("2. Arquitectura del Compilador: Fases y Descripción")
    r_h2.font.name = 'Calibri'
    r_h2.font.size = Pt(14)
    r_h2.font.bold = True
    r_h2.font.color.rgb = COLOR_PRIMARY

    p_intro2 = doc.add_paragraph()
    p_intro2.paragraph_format.space_after = Pt(10)
    r_intro2 = p_intro2.add_run(
        "A continuación se presenta el flujo secuencial de procesamiento por el cual transita un archivo de código fuente escrito en DronLang "
        "hasta transformarse en la secuencia de micro-instrucciones o comandos binarios de vuelo que comandan el dron de reparto."
    )
    r_intro2.font.name = 'Calibri'
    r_intro2.font.size = Pt(10.5)
    r_intro2.font.color.rgb = COLOR_TEXT

    # Diagrama de Fases estilizado como Tabla
    diagram_table = doc.add_table(rows=9, cols=3)
    diagram_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    d_widths = [Inches(1.8), Inches(3.2), Inches(1.5)]
    
    flow_steps = [
        ("Entrada", "CÓDIGO FUENTE (.dron)", "Texto plano del programa de vuelo"),
        ("↓", "Flujo de caracteres individuales", "Lectura secuencial carácter a carácter"),
        ("FASE 1", "ANÁLISIS LÉXICO (Lexer)", "Agrupación en tokens reconocibles"),
        ("↓", "Secuencia de Tokens [ {tipo, valor, pos} ]", "Descarte de espacios y comentarios"),
        ("FASE 2", "ANÁLISIS SINTÁCTICO (Parser)", "Validación de la gramática formal"),
        ("↓", "Árbol de Sintaxis Abstracta (AST)", "Estructura jerárquica de nodos"),
        ("FASE 3", "ANÁLISIS SEMÁNTICO (Checker)", "Validación de tipos, rangos y variables"),
        ("↓", "AST Validado y Tabla de Símbolos", "Lógica coherente y verificada"),
        ("FASE 4", "GENERACIÓN DE CÓDIGO (Codegen)", "Emisión de comandos de vuelo ejecutables")
    ]
    
    for row_idx, (fase, main_box, note) in enumerate(flow_steps):
        row = diagram_table.rows[row_idx]
        c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
        c0.width, c1.width, c2.width = d_widths
        
        is_phase = fase.startswith("FASE")
        is_arrow = (fase == "↓")
        
        if is_phase:
            bg = "1B365D"
            txt_color = RGBColor(0xFF, 0xFF, 0xFF)
            border_color = "0F172A"
        elif is_arrow:
            bg = "FFFFFF"
            txt_color = RGBColor(0x64, 0x74, 0x8B)
            border_color = "FFFFFF"
        else:
            bg = "F0F4F8"
            txt_color = COLOR_PRIMARY
            border_color = "CBD5E1"
            
        for cell in (c0, c1, c2):
            set_cell_background(cell, bg)
            set_cell_margins(cell, 70, 70, 100, 100)
            if is_phase:
                set_cell_border(cell, 
                                top={'sz': 6, 'val': 'single', 'color': border_color},
                                bottom={'sz': 6, 'val': 'single', 'color': border_color})
            elif not is_arrow:
                set_cell_border(cell, 
                                top={'sz': 4, 'val': 'single', 'color': border_color},
                                bottom={'sz': 4, 'val': 'single', 'color': border_color})
            else:
                set_cell_border(cell, 
                                top={'sz': 0, 'val': 'none', 'color': 'FFFFFF'},
                                bottom={'sz': 0, 'val': 'none', 'color': 'FFFFFF'})

        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(fase)
        r0.font.bold = True
        r0.font.size = Pt(9.5) if is_phase else Pt(11 if is_arrow else 9)
        r0.font.color.rgb = txt_color
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(main_box)
        r1.font.bold = is_phase or (row_idx == 0)
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = txt_color
        
        p2 = c2.paragraphs[0]
        p2.paragraph_format.space_after = Pt(0)
        r2 = p2.add_run(note)
        r2.font.italic = True
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = RGBColor(0x94, 0xA3, 0xB8) if is_phase else COLOR_MUTED

    # Sección: Una frase por fase
    h2_sub = doc.add_paragraph()
    h2_sub.paragraph_format.space_before = Pt(14)
    h2_sub.paragraph_format.space_after = Pt(6)
    r_h2_sub = h2_sub.add_run("Una Frase por Fase en el Contexto de DronLang:")
    r_h2_sub.font.name = 'Calibri'
    r_h2_sub.font.size = Pt(11.5)
    r_h2_sub.font.bold = True
    r_h2_sub.font.color.rgb = COLOR_PRIMARY

    phrases = [
        ("Fase 1 — Análisis Léxico (Scanner/Lexer):", 
         "Lee secuencialmente el texto fuente del programa de vuelo carácter por carácter y lo transforma en una colección de fichas o tokens reconocibles (palabras clave como 'despegar', números, operadores e identificadores), descartando espacios y comentarios, y abortando con error si encuentra caracteres no admitidos (como '@' o '$')."),
        
        ("Fase 2 — Análisis Sintáctico (Parser):", 
         "Verifica que la secuencia de tokens cumpla rigurosamente con la estructura y las reglas de producción gramaticales de DronLang (por ejemplo, que toda estructura 'repetir' esté acompañada de 'veces' y cerrada con 'fin'), construyendo en memoria el Árbol de Sintaxis Abstracta (AST) representativo."),
        
        ("Fase 3 — Análisis Semántico (Type & Logic Checker):", 
         "Comprueba el significado y la coherencia lógica de las instrucciones del vuelo mediante la construcción de la Tabla de Símbolos, garantizando que ninguna variable sea empleada antes de ser inicializada y comprobando que valores críticos (como distancias o cantidades de paquetes) no adopten números negativos."),
        
        ("Fase 4 — Generación de Código (Code Generator):", 
         "Recorre el árbol de sintaxis validado y lo traduce a una secuencia final de comandos de telemetría de bajo nivel ('DESPEGAR', 'VOLAR', 'ENTREGAR', 'ATERRIZAR') o a código JavaScript ejecutable para simular la trayectoria del dron.")
    ]

    for title, phrase in phrases:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15
        
        r_t = p.add_run(f"• {title} ")
        r_t.font.name = 'Calibri'
        r_t.font.bold = True
        r_t.font.size = Pt(10)
        r_t.font.color.rgb = COLOR_SECONDARY
        
        r_p = p.add_run(phrase)
        r_p.font.name = 'Calibri'
        r_p.font.size = Pt(10)
        r_p.font.color.rgb = COLOR_TEXT

    # =========================================================================
    # PUNTO 3 & 4: SEIS PROGRAMAS DE EJEMPLO Y EXPLICACIÓN DE ERRORES
    # =========================================================================
    doc.add_page_break()
    
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(8)
    h3.paragraph_format.space_after = Pt(6)
    r_h3 = h3.add_run("3. Banco de Programas de Prueba: Tres Correctos y Tres con Error")
    r_h3.font.name = 'Calibri'
    r_h3.font.size = Pt(14)
    r_h3.font.bold = True
    r_h3.font.color.rgb = COLOR_PRIMARY

    p_intro3 = doc.add_paragraph()
    p_intro3.paragraph_format.space_after = Pt(8)
    r_intro3 = p_intro3.add_run(
        "Se diseñaron seis programas completos para validar las diferentes capacidades y límites del compilador. "
        "Los primeros tres programas demuestran el funcionamiento correcto y la interacción de variables, condiciones y bucles. "
        "Los tres programas posteriores introducen errores intencionales diseñados específicamente para ser capturados de forma unívoca "
        "por la Fase Léxica, la Fase Sintáctica y la Fase Semántica respectivamente."
    )
    r_intro3.font.name = 'Calibri'
    r_intro3.font.size = Pt(10)
    r_intro3.font.color.rgb = COLOR_TEXT

    # PROGRAMA 1
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(6)
    p_p1.paragraph_format.space_after = Pt(2)
    r_p1 = p_p1.add_run("Programa 1 (Correcto) — Misión de Reparto Simple de Punto a Punto")
    r_p1.font.name = 'Calibri'
    r_p1.font.bold = True
    r_p1.font.size = Pt(11)
    r_p1.font.color.rgb = COLOR_SECONDARY
    
    code_p1 = (
        "# Misión básica: variables de ruta y entrega\n"
        "distancia = 15\n"
        "paquetes = 2\n\n"
        "despegar\n"
        "volar distancia\n"
        "entregar paquetes\n"
        "volar distancia\n"
        "aterrizar"
    )
    add_code_block(doc, code_p1)
    
    p_exp1 = doc.add_paragraph()
    p_exp1.paragraph_format.space_after = Pt(10)
    p_exp1.paragraph_format.line_spacing = 1.15
    r_exp1 = p_exp1.add_run(
        "Descripción funcional: Define las variables 'distancia' (15 m) y 'paquetes' (2 unidades). "
        "El dron despega, vuela 15 metros hasta la coordenada destino, descarga 2 paquetes mediante el actuador, "
        "emprende el vuelo de retorno por otros 15 metros y ejecuta un aterrizaje seguro. "
        "Pasa exitosamente las 4 fases del compilador."
    )
    r_exp1.font.size = Pt(9.5)
    r_exp1.font.color.rgb = COLOR_TEXT

    # PROGRAMA 2
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(4)
    p_p2.paragraph_format.space_after = Pt(2)
    r_p2 = p_p2.add_run("Programa 2 (Correcto) — Condición de Autonomía de Batería con Operaciones")
    r_p2.font.name = 'Calibri'
    r_p2.font.bold = True
    r_p2.font.size = Pt(11)
    r_p2.font.color.rgb = COLOR_SECONDARY
    
    code_p2 = (
        "# Evaluación de autonomía antes de vuelo largo\n"
        "bateria = 75\n"
        "consumoRuta = 10 * 3\n"
        "bateriaRestante = bateria - consumoRuta\n\n"
        "despegar\n"
        "si bateriaRestante > 30 entonces\n"
        "    volar 20\n"
        "    entregar 1\n"
        "sino\n"
        "    volar 5\n"
        "    entregar 1\n"
        "fin\n"
        "aterrizar"
    )
    add_code_block(doc, code_p2)
    
    p_exp2 = doc.add_paragraph()
    p_exp2.paragraph_format.space_after = Pt(10)
    p_exp2.paragraph_format.line_spacing = 1.15
    r_exp2 = p_exp2.add_run(
        "Descripción funcional: Aplica multiplicaciones (10 * 3 = 30) y restas (75 - 30 = 45) en las expresiones. "
        "Evalúa la condición 'bateriaRestante > 30' (45 > 30 es verdadero), tomando la bifurcación principal de 20 metros "
        "para la entrega antes de aterrizar. Valida el uso correcto de operadores y estructura 'si ... entonces ... sino ... fin'."
    )
    r_exp2.font.size = Pt(9.5)
    r_exp2.font.color.rgb = COLOR_TEXT

    # PROGRAMA 3
    p_p3 = doc.add_paragraph()
    p_p3.paragraph_format.space_before = Pt(4)
    p_p3.paragraph_format.space_after = Pt(2)
    r_p3 = p_p3.add_run("Programa 3 (Correcto) — Reparto Multizona con Bucle Iterativo")
    r_p3.font.name = 'Calibri'
    r_p3.font.bold = True
    r_p3.font.size = Pt(11)
    r_p3.font.color.rgb = COLOR_SECONDARY
    
    code_p3 = (
        "# Entrega cíclica en 4 sectores residenciales\n"
        "zonas = 4\n"
        "distanciaSector = 5 + 3\n\n"
        "despegar\n"
        "repetir zonas veces\n"
        "    volar distanciaSector\n"
        "    entregar 1\n"
        "fin\n"
        "aterrizar"
    )
    add_code_block(doc, code_p3)
    
    p_exp3 = doc.add_paragraph()
    p_exp3.paragraph_format.space_after = Pt(12)
    p_exp3.paragraph_format.line_spacing = 1.15
    r_exp3 = p_exp3.add_run(
        "Descripción funcional: Calcula la distancia por sector (5 + 3 = 8 m) y programa un bucle repetitivo de 4 ciclos. "
        "En cada iteración del bucle, el dron vuela 8 metros y descarga 1 paquete. Al finalizar las 4 entregas (32 m recorridos), "
        "el control sale del ciclo y ejecuta 'aterrizar'. Demuestra el manejo de variables dentro del encabezado 'repetir ... veces'."
    )
    r_exp3.font.size = Pt(9.5)
    r_exp3.font.color.rgb = COLOR_TEXT

    # =========================================================================
    # EXPLICACIÓN DETALLADA DE ERRORES (PUNTO 4)
    # =========================================================================
    doc.add_page_break()
    
    h4 = doc.add_paragraph()
    h4.paragraph_format.space_before = Pt(8)
    h4.paragraph_format.space_after = Pt(6)
    r_h4 = h4.add_run("4. Demostración y Diagnóstico Riguroso de Errores Intencionales")
    r_h4.font.name = 'Calibri'
    r_h4.font.size = Pt(14)
    r_h4.font.bold = True
    r_h4.font.color.rgb = COLOR_PRIMARY

    p_intro4 = doc.add_paragraph()
    p_intro4.paragraph_format.space_after = Pt(8)
    r_intro4 = p_intro4.add_run(
        "A continuación se presenta el análisis de los tres programas erróneos, detallando en qué fase exacta se intercepta cada fallo, "
        "el fundamento teórico de por qué ocurre allí y por qué dicha anomalía no puede ser procesada por las fases siguientes."
    )
    r_intro4.font.size = Pt(10)
    r_intro4.font.color.rgb = COLOR_TEXT

    # PROGRAMA 4 (ERROR LÉXICO)
    p_p4 = doc.add_paragraph()
    p_p4.paragraph_format.space_before = Pt(4)
    p_p4.paragraph_format.space_after = Pt(2)
    r_p4 = p_p4.add_run("Programa 4 — Error Intencional: Error Léxico")
    r_p4.font.name = 'Calibri'
    r_p4.font.bold = True
    r_p4.font.size = Pt(11)
    r_p4.font.color.rgb = RGBColor(0x99, 0x1B, 0x1B) # Rojo oscuro
    
    code_p4 = (
        "bateria = 100 @\n"
        "despegar\n"
        "volar 10\n"
        "aterrizar"
    )
    add_code_block(doc, code_p4)

    # Diagnóstico del error léxico
    t_err1 = doc.add_table(rows=3, cols=2)
    t_err1.alignment = WD_TABLE_ALIGNMENT.CENTER
    diag_e1 = [
        ("Fase de Detección:", "Fase 1: Análisis Léxico (Lexer)"),
        ("Mensaje del Compilador:", "[Error Léxico] Línea 1, Columna 15: Símbolo desconocido o no permitido en DronLang: '@'"),
        ("Fundamento y Explicación Técnica:", 
         "El compilador procesa el código fuente carácter por carácter mediante autómatas de reconocimiento. "
         "El símbolo '@' no forma parte del alfabeto formal (∑) de DronLang: no es un dígito, no es letra, "
         "no es operador aritmético ni relacional, ni tampoco un delimitador válido. "
         "Al encontrarse en la línea 1, columna 15, el método tokenizar() de src/lexer.js no puede asignarlo a ningún token válido "
         "y arroja una excepción inmediata, abortando la compilación sin permitir que el código pase al Parser.")
    ]
    for idx, (lbl, val) in enumerate(diag_e1):
        row = t_err1.rows[idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.2)
        c1.width = Inches(4.3)
        set_cell_background(c0, "FEF2F2") # Fondo rojizo suave
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c0, 70, 70, 100, 100)
        set_cell_margins(c1, 70, 70, 100, 100)
        set_cell_border(c0, bottom={'sz': 4, 'val': 'single', 'color': 'FECACA'})
        set_cell_border(c1, bottom={'sz': 4, 'val': 'single', 'color': 'FECACA'})
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(lbl)
        r0.font.bold = True
        r0.font.size = Pt(9)
        r0.font.color.rgb = RGBColor(0x99, 0x1B, 0x1B)
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        p1.paragraph_format.line_spacing = 1.15
        r1 = p1.add_run(val)
        r1.font.size = Pt(9)
        r1.font.color.rgb = COLOR_TEXT

    # PROGRAMA 5 (ERROR SINTÁCTICO)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    p_p5 = doc.add_paragraph()
    p_p5.paragraph_format.space_before = Pt(6)
    p_p5.paragraph_format.space_after = Pt(2)
    r_p5 = p_p5.add_run("Programa 5 — Error Intencional: Error Sintáctico")
    r_p5.font.name = 'Calibri'
    r_p5.font.bold = True
    r_p5.font.size = Pt(11)
    r_p5.font.color.rgb = RGBColor(0xC2, 0x41, 0x0C) # Naranja oscuro
    
    code_p5 = (
        "entregas = 3\n"
        "despegar\n"
        "repetir entregas veces\n"
        "    volar 12\n"
        "    entregar 1\n"
        "aterrizar"
    )
    add_code_block(doc, code_p5)

    t_err2 = doc.add_table(rows=3, cols=2)
    t_err2.alignment = WD_TABLE_ALIGNMENT.CENTER
    diag_e2 = [
        ("Fase de Detección:", "Fase 2: Análisis Sintáctico (Parser)"),
        ("Mensaje del Compilador:", "[Error Sintáctico] Línea 6, Columna 1: Se esperaba 'fin' para cerrar la estructura 'repetir', pero se encontró 'aterrizar'"),
        ("Fundamento y Explicación Técnica:", 
         "Todos los caracteres y palabras individuales del programa existen y son válidos ('entregas', 'repetir', 'volar', etc.), "
         "por lo que la Fase 1 (Lexer) termina con éxito al 100% generando la lista de tokens sin reclamo alguno. "
         "Sin embargo, al llegar a la Fase 2, el Parser aplica las reglas de la gramática libre de contexto: "
         "la producción formal exige que todo bloque 'repetir' esté concluido de forma determinista por el token 'fin'. "
         "Al toparse con 'aterrizar' sin haberse cerrado el ciclo, el analizador sintáctico detecta una oración mal construida y lanza el error sintáctico.")
    ]
    for idx, (lbl, val) in enumerate(diag_e2):
        row = t_err2.rows[idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.2)
        c1.width = Inches(4.3)
        set_cell_background(c0, "FFF7ED") # Fondo anaranjado suave
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c0, 70, 70, 100, 100)
        set_cell_margins(c1, 70, 70, 100, 100)
        set_cell_border(c0, bottom={'sz': 4, 'val': 'single', 'color': 'FED7AA'})
        set_cell_border(c1, bottom={'sz': 4, 'val': 'single', 'color': 'FED7AA'})
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(lbl)
        r0.font.bold = True
        r0.font.size = Pt(9)
        r0.font.color.rgb = RGBColor(0xC2, 0x41, 0x0C)
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        p1.paragraph_format.line_spacing = 1.15
        r1 = p1.add_run(val)
        r1.font.size = Pt(9)
        r1.font.color.rgb = COLOR_TEXT

    # PROGRAMA 6 (ERROR SEMÁNTICO)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    p_p6 = doc.add_paragraph()
    p_p6.paragraph_format.space_before = Pt(6)
    p_p6.paragraph_format.space_after = Pt(2)
    r_p6 = p_p6.add_run("Programa 6 — Error Intencional: Error Semántico")
    r_p6.font.name = 'Calibri'
    r_p6.font.bold = True
    r_p6.font.size = Pt(11)
    r_p6.font.color.rgb = RGBColor(0x7C, 0x2D, 0x12) # Marrón/Ámbar oscuro
    
    code_p6 = (
        "paquetes = 2\n"
        "despegar\n"
        "volar distanciaDestino\n"
        "entregar paquetes\n"
        "aterrizar"
    )
    add_code_block(doc, code_p6)

    t_err3 = doc.add_table(rows=3, cols=2)
    t_err3.alignment = WD_TABLE_ALIGNMENT.CENTER
    diag_e3 = [
        ("Fase de Detección:", "Fase 3: Análisis Semántico (Semantic Analyzer)"),
        ("Mensaje del Compilador:", "[Error Semántico] Línea 3: La variable 'distanciaDestino' no ha sido declarada o inicializada"),
        ("Fundamento y Explicación Técnica:", 
         "Este caso ilustra a la perfección el propósito del análisis semántico. "
         "1. Para el Léxico: 'distanciaDestino' es un IDENTIFICADOR completamente válido según el patrón [a-zA-Z_]. Pasa sin error.\n"
         "2. Para el Sintáctico: la estructura 'volar <IDENTIFICADOR>' cumple la regla gramatical formal. El árbol AST se construye sin fallo.\n"
         "3. En el Semántico: el analizador consulta la Tabla de Símbolos. Al no existir un registro previo de asignación para 'distanciaDestino', "
         "determina que la instrucción carece de sentido operativo y lógico: el dron recibiría una orden de desplazarse una distancia nula o indefinida, "
         "lo cual compromete la misión y la integridad del dron en el mundo real.")
    ]
    for idx, (lbl, val) in enumerate(diag_e3):
        row = t_err3.rows[idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.2)
        c1.width = Inches(4.3)
        set_cell_background(c0, "FEF3C7") # Fondo amarillo suave
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c0, 70, 70, 100, 100)
        set_cell_margins(c1, 70, 70, 100, 100)
        set_cell_border(c0, bottom={'sz': 4, 'val': 'single', 'color': 'FDE68A'})
        set_cell_border(c1, bottom={'sz': 4, 'val': 'single', 'color': 'FDE68A'})
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(lbl)
        r0.font.bold = True
        r0.font.size = Pt(9)
        r0.font.color.rgb = RGBColor(0x92, 0x40, 0x0E)
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        p1.paragraph_format.line_spacing = 1.15
        r1 = p1.add_run(val)
        r1.font.size = Pt(9)
        r1.font.color.rgb = COLOR_TEXT

    # =========================================================================
    # TABLA COMPARATIVA DE CIERRE
    # =========================================================================
    doc.add_page_break()
    
    h_concl = doc.add_paragraph()
    h_concl.paragraph_format.space_before = Pt(8)
    h_concl.paragraph_format.space_after = Pt(6)
    r_hc = h_concl.add_run("5. Cuadro Comparativo de los Tres Niveles de Error en Compilación")
    r_hc.font.name = 'Calibri'
    r_hc.font.size = Pt(13)
    r_hc.font.bold = True
    r_hc.font.color.rgb = COLOR_PRIMARY

    summary_table = doc.add_table(rows=4, cols=4)
    summary_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    s_headers = ["Nivel de Error", "Fase que lo Captura", "Causa en el Código", "Impacto en el Dron de Reparto"]
    s_widths = [Inches(1.4), Inches(1.6), Inches(1.8), Inches(1.7)]

    for col_idx, (text, width) in enumerate(zip(s_headers, s_widths)):
        cell = summary_table.cell(0, col_idx)
        cell.width = width
        set_cell_background(cell, "1B365D")
        set_cell_margins(cell, 100, 100, 80, 80)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    summary_rows = [
        ("Léxico", "Fase 1 (Lexer)", "Carácter '@' no existente en el alfabeto del lenguaje.", "El programa ni siquiera puede ser leído en palabras; no genera tokens."),
        ("Sintáctico", "Fase 2 (Parser)", "Bloque 'repetir' sin palabra de cierre 'fin'.", "Palabras reconocidas, pero la oración no sigue la gramática; no genera AST."),
        ("Semántico", "Fase 3 (Semántico)", "Variable 'distanciaDestino' no declarada previa.", "Gramática impecable, pero carece de sentido lógico (distancia indefinida en metros).")
    ]

    for row_idx, data in enumerate(summary_rows, start=1):
        row = summary_table.rows[row_idx]
        bg = "FFFFFF" if row_idx % 2 != 0 else "F8FAFC"
        for col_idx, text in enumerate(data):
            cell = row.cells[col_idx]
            cell.width = s_widths[col_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, 80, 80, 80, 80)
            set_cell_border(cell, 
                            bottom={'sz': 4, 'val': 'single', 'color': 'E2E8F0'},
                            top={'sz': 4, 'val': 'single', 'color': 'E2E8F0'})
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.15
            run = p.add_run(text)
            run.font.name = 'Calibri'
            run.font.size = Pt(8.5)
            run.font.bold = (col_idx == 0)
            run.font.color.rgb = COLOR_PRIMARY if col_idx == 0 else COLOR_TEXT

    # Enlace y nota final
    p_final = doc.add_paragraph()
    p_final.paragraph_format.space_before = Pt(20)
    p_final.paragraph_format.space_after = Pt(4)
    p_final.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_f1 = p_final.add_run("Proyecto alojado en GitHub: ")
    r_f1.font.bold = True
    r_f1.font.size = Pt(10)
    r_f1.font.color.rgb = COLOR_PRIMARY
    
    r_f2 = p_final.add_run("https://github.com/andrescrack2024/Compilador\n")
    r_f2.font.underline = True
    r_f2.font.size = Pt(10)
    r_f2.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
    
    r_f3 = p_final.add_run("Desarrollado por Jainer Andrés Morales Cuesta & Sharly Andrés Mosquera Rodríguez")
    r_f3.font.italic = True
    r_f3.font.size = Pt(9.5)
    r_f3.font.color.rgb = COLOR_MUTED

    # Guardar documento
    output_path = os.path.join(os.getcwd(), "Entrega_Actividad_3_DronLang.docx")
    doc.save(output_path)
    print(f"Documento generado exitosamente en: {output_path}")

if __name__ == "__main__":
    build_document()
