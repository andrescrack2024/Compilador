# -*- coding: utf-8 -*-
"""
Generador de documento Word oficial para la Actividad 3
Cumplimiento estricto:
- Fuente: Times New Roman
- Sin colores (Monocromo / Escala de grises pura, texto negro #000000)
- Tamaño: 12 pt para cuerpo de texto
- Interlineado: 1.5 líneas
- Formato y Portada: Normas APA (7ma edición estudiantil)
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

BLACK = RGBColor(0, 0, 0)
GRAY_LINE = "000000"

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
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

def apply_apa_table_borders(table):
    """
    Aplica las normas APA para tablas:
    - Borde superior en la cabecera (fila 0)
    - Borde inferior en la cabecera (fila 0)
    - Borde inferior en la última fila
    - Sin líneas verticales
    - Sin líneas horizontales intermedias
    """
    num_rows = len(table.rows)
    for r_idx, row in enumerate(table.rows):
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            # Limpiar bordes previos
            tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
            
            # Borde superior solo en fila 0
            if r_idx == 0:
                top_b = parse_xml(f'<w:top {nsdecls("w")} w:val="single" w:sz="8" w:space="0" w:color="{GRAY_LINE}"/>')
                tcBorders.append(top_b)
                bottom_b = parse_xml(f'<w:bottom {nsdecls("w")} w:val="single" w:sz="6" w:space="0" w:color="{GRAY_LINE}"/>')
                tcBorders.append(bottom_b)
            elif r_idx == num_rows - 1:
                bottom_b = parse_xml(f'<w:bottom {nsdecls("w")} w:val="single" w:sz="8" w:space="0" w:color="{GRAY_LINE}"/>')
                tcBorders.append(bottom_b)
                
            # Desactivar explícitamente izquierda y derecha
            left_b = parse_xml(f'<w:left {nsdecls("w")} w:val="none"/>')
            right_b = parse_xml(f'<w:right {nsdecls("w")} w:val="none"/>')
            tcBorders.append(left_b)
            tcBorders.append(right_b)
            
            tcPr.append(tcBorders)

def add_apa_code_block(doc, code_text):
    """Bloque de código formateado estilo APA monocromático."""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    cell = table.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_margins(cell, top=100, bottom=100, left=150, right=150)
    
    # Borde sutil negro alrededor
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="555555"/>'
        f'<w:left w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="555555"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="555555"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(tcBorders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    
    run = p.add_run(code_text.strip())
    run.font.name = 'Courier New'
    run.font.size = Pt(10)
    run.font.color.rgb = BLACK
    
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(6)

def build_apa_document():
    doc = docx.Document()
    
    # -------------------------------------------------------------------------
    # CONFIGURACIÓN DE PÁGINA (NORMAS APA 7: 2.54 cm = 1 pulgada en todos lados)
    # -------------------------------------------------------------------------
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
        # Encabezado APA: Solo número de página alineado a la derecha
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        # Campo de página
        fldSimple = OxmlElement('w:fldSimple')
        fldSimple.set(qn('w:instr'), 'PAGE')
        hp._p.append(fldSimple)
        for r in hp.runs:
            r.font.name = 'Times New Roman'
            r.font.size = Pt(12)
            r.font.color.rgb = BLACK

    # =========================================================================
    # PORTADA ESTUDIANTIL SEGÚN NORMAS APA (7ma EDICIÓN)
    # =========================================================================
    # Espaciado inicial de 3 a 4 líneas en blanco
    for _ in range(3):
        p_blank = doc.add_paragraph()
        p_blank.paragraph_format.line_spacing = 1.5
        p_blank.paragraph_format.space_before = Pt(0)
        p_blank.paragraph_format.space_after = Pt(0)

    # Título del trabajo (Centrado, Negrita, 12 pt, interlineado 1.5)
    p_tit = doc.add_paragraph()
    p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tit.paragraph_format.line_spacing = 1.5
    p_tit.paragraph_format.space_after = Pt(18)
    
    r_tit = p_tit.add_run("Diseño y Especificación del Lenguaje DronLang para Drones de Reparto\n"
                          "Actividad 3: Examen de Conceptualización Básica")
    r_tit.font.name = 'Times New Roman'
    r_tit.font.size = Pt(12)
    r_tit.font.bold = True
    r_tit.font.color.rgb = BLACK

    # Línea en blanco de separación
    p_sep = doc.add_paragraph()
    p_sep.paragraph_format.line_spacing = 1.5
    p_sep.paragraph_format.space_after = Pt(18)

    # Autores (Centrado, 12 pt)
    p_aut = doc.add_paragraph()
    p_aut.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_aut.paragraph_format.line_spacing = 1.5
    p_aut.paragraph_format.space_after = Pt(24)
    
    r_aut = p_aut.add_run("Jainer Andrés Morales Cuesta\nSharly Andrés Mosquera Rodríguez")
    r_aut.font.name = 'Times New Roman'
    r_aut.font.size = Pt(12)
    r_aut.font.color.rgb = BLACK

    # Afiliación institucional, asignatura, docente y fecha
    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst.paragraph_format.line_spacing = 1.5
    p_inst.paragraph_format.space_after = Pt(0)
    
    info_text = (
        "Facultad de Ingeniería\n"
        "Programa de Ingeniería de Sistemas y Computación\n"
        "Asignatura: Compiladores\n"
        "Docente: Ing. Hasler Gamboa Berrío\n"
        "9 de noviembre de 2026\n"
        "Repositorio GitHub: https://github.com/andrescrack2024/Compilador"
    )
    r_inst = p_inst.add_run(info_text)
    r_inst.font.name = 'Times New Roman'
    r_inst.font.size = Pt(12)
    r_inst.font.color.rgb = BLACK

    # Salto de página para iniciar el contenido formal
    doc.add_page_break()

    # =========================================================================
    # INTRODUCCIÓN GENERAL DEL LENGUAJE
    # =========================================================================
    p_h1 = doc.add_paragraph()
    p_h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_h1.paragraph_format.space_before = Pt(6)
    p_h1.paragraph_format.space_after = Pt(12)
    p_h1.paragraph_format.line_spacing = 1.5
    r_h1 = p_h1.add_run("Diseño y Conceptualización de DronLang")
    r_h1.font.name = 'Times New Roman'
    r_h1.font.size = Pt(12)
    r_h1.font.bold = True
    r_h1.font.color.rgb = BLACK

    p_body1 = doc.add_paragraph()
    p_body1.paragraph_format.line_spacing = 1.5
    p_body1.paragraph_format.first_line_indent = Inches(0.5)
    p_body1.paragraph_format.space_after = Pt(12)
    r_b1 = p_body1.add_run(
        "El proyecto curricular de la asignatura Compiladores plantea el diseño y desarrollo de un compilador "
        "propio para un lenguaje de dominio específico. Para el presente grupo de trabajo, la temática asignada "
        "por la cátedra corresponde a Dron de reparto. Con base en dicho requerimiento, se diseñó DronLang, "
        "un lenguaje orientado a la planificación, navegación, entrega de paquetería y control seguro de misiones "
        "aéreas autónomas. Su sintaxis adopta una filosofía de claridad y legibilidad influenciada por Python: "
        "se suprimen caracteres de terminación obligatorios como el punto y coma (;), las instrucciones se delimitan "
        "por saltos de línea y las estructuras de control se cierran de forma determinista mediante la palabra clave fin."
    )
    r_b1.font.name = 'Times New Roman'
    r_b1.font.size = Pt(12)
    r_b1.font.color.rgb = BLACK

    # =========================================================================
    # PUNTO 1: CATÁLOGO DE INSTRUCCIONES (CÓMO SE ESCRIBE Y QUÉ HACE)
    # =========================================================================
    p_h2 = doc.add_paragraph()
    p_h2.paragraph_format.space_before = Pt(12)
    p_h2.paragraph_format.space_after = Pt(6)
    p_h2.paragraph_format.line_spacing = 1.5
    r_h2 = p_h2.add_run("Catálogo de Instrucciones del Lenguaje")
    r_h2.font.name = 'Times New Roman'
    r_h2.font.size = Pt(12)
    r_h2.font.bold = True
    r_h2.font.color.rgb = BLACK

    p_p1_desc = doc.add_paragraph()
    p_p1_desc.paragraph_format.line_spacing = 1.5
    p_p1_desc.paragraph_format.first_line_indent = Inches(0.5)
    p_p1_desc.paragraph_format.space_after = Pt(12)
    r_p1d = p_p1_desc.add_run(
        "DronLang contempla un conjunto de primitivas especializadas en maniobras de vuelo y gestión de carga, "
        "acompañadas de facilidades para el cálculo aritmético, el almacenamiento de variables y el control de flujo. "
        "En la Tabla 1 se presenta la especificación formal requerida para cada instrucción aceptada por el compilador."
    )
    r_p1d.font.name = 'Times New Roman'
    r_p1d.font.size = Pt(12)
    r_p1d.font.color.rgb = BLACK

    # Etiqueta de Tabla APA
    p_t1_num = doc.add_paragraph()
    p_t1_num.paragraph_format.space_before = Pt(6)
    p_t1_num.paragraph_format.space_after = Pt(2)
    p_t1_num.paragraph_format.line_spacing = 1.5
    r_t1n = p_t1_num.add_run("Tabla 1")
    r_t1n.font.name = 'Times New Roman'
    r_t1n.font.size = Pt(12)
    r_t1n.font.bold = True
    r_t1n.font.color.rgb = BLACK

    p_t1_tit = doc.add_paragraph()
    p_t1_tit.paragraph_format.space_before = Pt(0)
    p_t1_tit.paragraph_format.space_after = Pt(6)
    p_t1_tit.paragraph_format.line_spacing = 1.5
    r_t1t = p_t1_tit.add_run("Especificación de Instrucciones, Sintaxis y Semántica en DronLang")
    r_t1t.font.name = 'Times New Roman'
    r_t1t.font.size = Pt(12)
    r_t1t.font.italic = True
    r_t1t.font.color.rgb = BLACK

    # Tabla APA 1: Instrucciones
    table_inst = doc.add_table(rows=8, cols=4)
    table_inst.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths = [Inches(1.2), Inches(1.7), Inches(2.3), Inches(1.3)]
    
    headers_t1 = ["Instrucción", "Sintaxis Formal", "Semántica y Efecto en el Dron", "Ejemplo"]
    for c_idx, (h_text, w) in enumerate(zip(headers_t1, widths)):
        cell = table_inst.cell(0, c_idx)
        cell.width = w
        set_cell_margins(cell, 80, 80, 80, 80)
        p = cell.paragraphs[0]
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h_text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = BLACK

    t1_rows = [
        ("despegar", "despegar", "Activa los motores y eleva el dron verticalmente a la altura base de crucero. No recibe parámetros.", "despegar"),
        ("volar", "volar <expresion>", "Desplaza horizontalmente el dron hacia el destino la distancia indicada en metros. La distancia debe ser >= 0.", "volar 15\nvolar distancia"),
        ("entregar", "entregar <expresion>", "Acciona el mecanismo de descarga para depositar la cantidad de paquetes indicada en el destino (>= 0).", "entregar 1\nentregar paquetes"),
        ("aterrizar", "aterrizar", "Realiza el descenso controlado en la zona de aterrizaje actual y apaga los motores.", "aterrizar"),
        ("Asignación", "<id> = <expresion>", "Guarda un valor numérico o el resultado de un cálculo en memoria para su utilización posterior.", "bateria = 100\npaquetes = 2"),
        ("repetir", "repetir <expr> veces\n    <bloque>\nfin", "Estructura iterativa determinista. Repite el bloque de instrucciones la cantidad de veces especificada (entero >= 0).", "repetir 3 veces\n    volar 10\n    entregar 1\nfin"),
        ("si ... sino", "si <cond> entonces\n    <bloque1>\nsino\n    <bloque2>\nfin", "Estructura condicional. Evalúa una comparación relacional; si se cumple ejecuta el primer bloque, y si no ejecuta el bloque sino.", "si bateria > 20 entonces\n    volar 30\nsino\n    aterrizar\nfin")
    ]

    for r_idx, (c0, c1, c2, c3) in enumerate(t1_rows, start=1):
        row = table_inst.rows[r_idx]
        for col_idx, (text, w) in enumerate(zip([c0, c1, c2, c3], widths)):
            cell = row.cells[col_idx]
            cell.width = w
            set_cell_margins(cell, 60, 60, 80, 80)
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(text)
            run.font.name = 'Courier New' if col_idx in (1, 3) else 'Times New Roman'
            run.font.size = Pt(9.5) if col_idx in (1, 3) else Pt(10)
            run.font.bold = (col_idx == 0)
            run.font.color.rgb = BLACK

    apply_apa_table_borders(table_inst)

    # Nota de tabla APA
    p_t1_nota = doc.add_paragraph()
    p_t1_nota.paragraph_format.space_before = Pt(4)
    p_t1_nota.paragraph_format.space_after = Pt(12)
    p_t1_nota.paragraph_format.line_spacing = 1.15
    r_t1n_lbl = p_t1_nota.add_run("Nota. ")
    r_t1n_lbl.font.name = 'Times New Roman'
    r_t1n_lbl.font.size = Pt(10)
    r_t1n_lbl.font.italic = True
    r_t1n_lbl.font.color.rgb = BLACK
    
    r_t1n_txt = p_t1_nota.add_run(
        "Todas las instrucciones del dron requieren ser escritas en minúsculas. "
        "Las operaciones aritméticas admitidas son suma (+), resta (-) y multiplicación (*). "
        "Las operaciones de comparación admitidas son mayor que (>), menor que (<) e igualdad estricta (==). "
        "Se admiten comentarios de una sola línea mediante el prefijo # o //."
    )
    r_t1n_txt.font.name = 'Times New Roman'
    r_t1n_txt.font.size = Pt(10)
    r_t1n_txt.font.color.rgb = BLACK

    # =========================================================================
    # PUNTO 2: DIBUJO DE LAS FASES Y UNA FRASE POR FASE
    # =========================================================================
    doc.add_page_break()

    p_h2_fases = doc.add_paragraph()
    p_h2_fases.paragraph_format.space_before = Pt(6)
    p_h2_fases.paragraph_format.space_after = Pt(6)
    p_h2_fases.paragraph_format.line_spacing = 1.5
    r_h2f = p_h2_fases.add_run("Arquitectura y Fases del Compilador DronLang")
    r_h2f.font.name = 'Times New Roman'
    r_h2f.font.size = Pt(12)
    r_h2f.font.bold = True
    r_h2f.font.color.rgb = BLACK

    p_fases_desc = doc.add_paragraph()
    p_fases_desc.paragraph_format.line_spacing = 1.5
    p_fases_desc.paragraph_format.first_line_indent = Inches(0.5)
    p_fases_desc.paragraph_format.space_after = Pt(12)
    r_fd = p_fases_desc.add_run(
        "El compilador procesa el programa de vuelo a través de cuatro fases consecutivas, "
        "garantizando que el texto ingresado sea sintácticamente válido, semánticamente coherente "
        "y convertible en comandos físicos de vuelo. En la Figura 1 se ilustra el esquema secuencial de transformación."
    )
    r_fd.font.name = 'Times New Roman'
    r_fd.font.size = Pt(12)
    r_fd.font.color.rgb = BLACK

    # Figura 1: Diagrama de fases en texto/tabla APA
    p_f1_num = doc.add_paragraph()
    p_f1_num.paragraph_format.space_before = Pt(4)
    p_f1_num.paragraph_format.space_after = Pt(2)
    p_f1_num.paragraph_format.line_spacing = 1.5
    r_f1n = p_f1_num.add_run("Figura 1")
    r_f1n.font.name = 'Times New Roman'
    r_f1n.font.size = Pt(12)
    r_f1n.font.bold = True
    r_f1n.font.color.rgb = BLACK

    p_f1_tit = doc.add_paragraph()
    p_f1_tit.paragraph_format.space_before = Pt(0)
    p_f1_tit.paragraph_format.space_after = Pt(6)
    p_f1_tit.paragraph_format.line_spacing = 1.5
    r_f1t = p_f1_tit.add_run("Flujo de Procesamiento en las Cuatro Fases del Compilador DronLang")
    r_f1t.font.name = 'Times New Roman'
    r_f1t.font.size = Pt(12)
    r_f1t.font.italic = True
    r_f1t.font.color.rgb = BLACK

    # Diagrama en bloque de texto con bordes sobrios
    diagram_code = (
        "       [ Código Fuente en DronLang (.dron) ]\n"
        "                         │\n"
        "                         ▼\n"
        "       ┌─────────────────────────────────────┐\n"
        "       │       1. ANÁLISIS LÉXICO (Lexer)    │\n"
        "       └─────────────────────────────────────┘\n"
        "                         │ Tira de Tokens reconocidos\n"
        "                         ▼\n"
        "       ┌─────────────────────────────────────┐\n"
        "       │     2. ANÁLISIS SINTÁCTICO (Parser) │\n"
        "       └─────────────────────────────────────┘\n"
        "                         │ Árbol de Sintaxis Abstracta (AST)\n"
        "                         ▼\n"
        "       ┌─────────────────────────────────────┐\n"
        "       │    3. ANÁLISIS SEMÁNTICO (Checker)  │\n"
        "       └─────────────────────────────────────┘\n"
        "                         │ AST Validado + Tabla de Símbolos\n"
        "                         ▼\n"
        "       ┌─────────────────────────────────────┐\n"
        "       │   4. GENERACIÓN DE CÓDIGO (Codegen) │\n"
        "       └─────────────────────────────────────┘\n"
        "                         │\n"
        "                         ▼\n"
        "       [ Comandos de Vuelo Ejecutables del Dron ]"
    )
    add_apa_code_block(doc, diagram_code)

    p_f1_nota = doc.add_paragraph()
    p_f1_nota.paragraph_format.space_before = Pt(2)
    p_f1_nota.paragraph_format.space_after = Pt(12)
    p_f1_nota.paragraph_format.line_spacing = 1.15
    r_f1_notalbl = p_f1_nota.add_run("Nota. ")
    r_f1_notalbl.font.name = 'Times New Roman'
    r_f1_notalbl.font.size = Pt(10)
    r_f1_notalbl.font.italic = True
    r_f1_notalbl.font.color.rgb = BLACK
    
    r_f1_notatxt = p_f1_nota.add_run(
        "Esquema de etapas del compilador de DronLang adaptado a las directrices de la Actividad 3."
    )
    r_f1_notatxt.font.name = 'Times New Roman'
    r_f1_notatxt.font.size = Pt(10)
    r_f1_notatxt.font.color.rgb = BLACK

    # Una frase por fase
    p_sub_frases = doc.add_paragraph()
    p_sub_frases.paragraph_format.space_before = Pt(8)
    p_sub_frases.paragraph_format.space_after = Pt(6)
    p_sub_frases.paragraph_format.line_spacing = 1.5
    r_subf = p_sub_frases.add_run("Descripción de las Fases (Una Frase por Fase):")
    r_subf.font.name = 'Times New Roman'
    r_subf.font.size = Pt(12)
    r_subf.font.bold = True
    r_subf.font.color.rgb = BLACK

    phrases_apa = [
        ("Fase 1 (Análisis Léxico):", 
         "Lee secuencialmente el código fuente del dron y agrupa los caracteres en unidades mínimas denominadas tokens "
         "(tales como palabras clave despegar, números y operadores), ignorando espacios y comentarios, "
         "y rechazando de inmediato caracteres no reconocidos en el alfabeto del lenguaje."),
        
        ("Fase 2 (Análisis Sintáctico):", 
         "Valida que la sucesión de tokens se organice conforme a las reglas gramaticales formales de DronLang "
         "(como asegurar que a la instrucción repetir no le falte la palabra veces ni el cierre fin) "
         "y construye el Árbol de Sintaxis Abstracta (AST) jerárquico."),
        
        ("Fase 3 (Análisis Semántico):", 
         "Verifica la consistencia lógica del programa mediante la Tabla de Símbolos, garantizando que ninguna "
         "variable sea utilizada sin inicialización previa y asegurando que magnitudes físicas como distancias "
         "o cantidades de entrega no asuman valores negativos."),
        
        ("Fase 4 (Generación de Código):", 
         "Traduce el árbol sintáctico validado en una serie ordenada de comandos de bajo nivel "
         "(DESPEGAR, VOLAR, ENTREGAR, ATERRIZAR) para la ejecución física o la simulación interactiva de la trayectoria.")
    ]

    for title, phrase in phrases_apa:
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.first_line_indent = Inches(0.5)
        
        r_t = p.add_run(f"{title} ")
        r_t.font.name = 'Times New Roman'
        r_t.font.size = Pt(12)
        r_t.font.bold = True
        r_t.font.color.rgb = BLACK
        
        r_txt = p.add_run(phrase)
        r_txt.font.name = 'Times New Roman'
        r_txt.font.size = Pt(12)
        r_txt.font.color.rgb = BLACK

    # =========================================================================
    # PUNTO 3: SEIS PROGRAMAS DE PRUEBA
    # =========================================================================
    doc.add_page_break()

    p_h3 = doc.add_paragraph()
    p_h3.paragraph_format.space_before = Pt(6)
    p_h3.paragraph_format.space_after = Pt(6)
    p_h3.paragraph_format.line_spacing = 1.5
    r_h3 = p_h3.add_run("Banco de Programas de Prueba de DronLang")
    r_h3.font.name = 'Times New Roman'
    r_h3.font.size = Pt(12)
    r_h3.font.bold = True
    r_h3.font.color.rgb = BLACK

    p_p3_desc = doc.add_paragraph()
    p_p3_desc.paragraph_format.line_spacing = 1.5
    p_p3_desc.paragraph_format.first_line_indent = Inches(0.5)
    p_p3_desc.paragraph_format.space_after = Pt(12)
    r_p3d = p_p3_desc.add_run(
        "Para corroborar el comportamiento del compilador se establecieron seis programas de prueba. "
        "Los primeros tres programas ilustran secuencias operativas válidas que superan satisfactoriamente "
        "todas las fases. Los tres programas subsiguientes contienen un error intencional diseñado para "
        "disparar de forma exclusiva la detección en la fase léxica, sintáctica y semántica, respectivamente."
    )
    r_p3d.font.name = 'Times New Roman'
    r_p3d.font.size = Pt(12)
    r_p3d.font.color.rgb = BLACK

    # Programa 1
    p_prog1_tit = doc.add_paragraph()
    p_prog1_tit.paragraph_format.space_before = Pt(8)
    p_prog1_tit.paragraph_format.space_after = Pt(2)
    p_prog1_tit.paragraph_format.line_spacing = 1.5
    r_p1t = p_prog1_tit.add_run("Programa 1 (Correcto): Misión Básica de Reparto Punto a Punto")
    r_p1t.font.name = 'Times New Roman'
    r_p1t.font.size = Pt(12)
    r_p1t.font.bold = True
    r_p1t.font.color.rgb = BLACK

    p1_code = (
        "distancia = 15\n"
        "paquetes = 2\n\n"
        "despegar\n"
        "volar distancia\n"
        "entregar paquetes\n"
        "volar distancia\n"
        "aterrizar"
    )
    add_apa_code_block(doc, p1_code)

    p1_exp = doc.add_paragraph()
    p1_exp.paragraph_format.line_spacing = 1.5
    p1_exp.paragraph_format.first_line_indent = Inches(0.5)
    p1_exp.paragraph_format.space_after = Pt(12)
    r_p1e = p1_exp.add_run(
        "Explicación operativa: El programa asigna las variables distancia con 15 metros y paquetes con 2 unidades. "
        "El dron despega, vuela 15 metros en línea recta hacia el punto de entrega, deposita los 2 paquetes solicitados, "
        "retorna recorriendo nuevamente 15 metros hacia la base de operaciones y realiza un aterrizaje seguro. "
        "La compilación es exitosa a través de las cuatro fases sin generar advertencias."
    )
    r_p1e.font.name = 'Times New Roman'
    r_p1e.font.size = Pt(12)
    r_p1e.font.color.rgb = BLACK

    # Programa 2
    p_prog2_tit = doc.add_paragraph()
    p_prog2_tit.paragraph_format.space_before = Pt(8)
    p_prog2_tit.paragraph_format.space_after = Pt(2)
    p_prog2_tit.paragraph_format.line_spacing = 1.5
    r_p2t = p_prog2_tit.add_run("Programa 2 (Correcto): Condición de Autonomía de Batería con Expresiones")
    r_p2t.font.name = 'Times New Roman'
    r_p2t.font.size = Pt(12)
    r_p2t.font.bold = True
    r_p2t.font.color.rgb = BLACK

    p2_code = (
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
    add_apa_code_block(doc, p2_code)

    p2_exp = doc.add_paragraph()
    p2_exp.paragraph_format.line_spacing = 1.5
    p2_exp.paragraph_format.first_line_indent = Inches(0.5)
    p2_exp.paragraph_format.space_after = Pt(12)
    r_p2e = p2_exp.add_run(
        "Explicación operativa: Evalúa expresiones aritméticas de multiplicación y sustracción antes del despegue. "
        "Calcula un consumo estimado de 30 unidades y una batería restante de 45 unidades. Al verificarse "
        "que la condición 45 > 30 es verdadera, el compilador genera la bifurcación para el vuelo extenso de 20 metros "
        "y la entrega respectiva, concluyendo con el aterrizaje reglamentario."
    )
    r_p2e.font.name = 'Times New Roman'
    r_p2e.font.size = Pt(12)
    r_p2e.font.color.rgb = BLACK

    # Programa 3
    p_prog3_tit = doc.add_paragraph()
    p_prog3_tit.paragraph_format.space_before = Pt(8)
    p_prog3_tit.paragraph_format.space_after = Pt(2)
    p_prog3_tit.paragraph_format.line_spacing = 1.5
    r_p3t = p_prog3_tit.add_run("Programa 3 (Correcto): Reparto Cíclico en Múltiples Zonas con Repetición")
    r_p3t.font.name = 'Times New Roman'
    r_p3t.font.size = Pt(12)
    r_p3t.font.bold = True
    r_p3t.font.color.rgb = BLACK

    p3_code = (
        "zonas = 4\n"
        "distanciaSector = 5 + 3\n\n"
        "despegar\n"
        "repetir zonas veces\n"
        "    volar distanciaSector\n"
        "    entregar 1\n"
        "fin\n"
        "aterrizar"
    )
    add_apa_code_block(doc, p3_code)

    p3_exp = doc.add_paragraph()
    p3_exp.paragraph_format.line_spacing = 1.5
    p3_exp.paragraph_format.first_line_indent = Inches(0.5)
    p3_exp.paragraph_format.space_after = Pt(12)
    r_p3e = p3_exp.add_run(
        "Explicación operativa: Combina la suma aritmética con el bucle iterativo repetir ... veces ... fin. "
        "Se calculan 8 metros por sector y se comanda al dron a repetir 4 veces la secuencia de vuelo y descarga. "
        "Al completarse las iteraciones, el control prosigue hacia la instrucción aterrizar."
    )
    r_p3e.font.name = 'Times New Roman'
    r_p3e.font.size = Pt(12)
    r_p3e.font.color.rgb = BLACK

    # =========================================================================
    # PUNTO 4: PROGRAMAS CON ERRORES Y JUSTIFICACIÓN TÉCNICA
    # =========================================================================
    doc.add_page_break()

    p_h4 = doc.add_paragraph()
    p_h4.paragraph_format.space_before = Pt(6)
    p_h4.paragraph_format.space_after = Pt(6)
    p_h4.paragraph_format.line_spacing = 1.5
    r_h4 = p_h4.add_run("Análisis y Diagnóstico Riguroso de Errores Intencionales")
    r_h4.font.name = 'Times New Roman'
    r_h4.font.size = Pt(12)
    r_h4.font.bold = True
    r_h4.font.color.rgb = BLACK

    p_p4_desc = doc.add_paragraph()
    p_p4_desc.paragraph_format.line_spacing = 1.5
    p_p4_desc.paragraph_format.first_line_indent = Inches(0.5)
    p_p4_desc.paragraph_format.space_after = Pt(12)
    r_p4d = p_p4_desc.add_run(
        "De conformidad con las especificaciones de la Actividad 3, se construyeron tres programas erróneos "
        "para demostrar la delimitación funcional entre las fases del compilador. Para cada caso se describe "
        "el código, la fase que intercepta la anomalía, el reporte arrojado por el compilador y la justificación teórica."
    )
    r_p4d.font.name = 'Times New Roman'
    r_p4d.font.size = Pt(12)
    r_p4d.font.color.rgb = BLACK

    # Error Léxico
    p_prog4_tit = doc.add_paragraph()
    p_prog4_tit.paragraph_format.space_before = Pt(8)
    p_prog4_tit.paragraph_format.space_after = Pt(2)
    p_prog4_tit.paragraph_format.line_spacing = 1.5
    r_p4t = p_prog4_tit.add_run("Programa 4 (Error Intencional): Presencia de Error Léxico")
    r_p4t.font.name = 'Times New Roman'
    r_p4t.font.size = Pt(12)
    r_p4t.font.bold = True
    r_p4t.font.color.rgb = BLACK

    p4_code = (
        "bateria = 100 @\n"
        "despegar\n"
        "volar 10\n"
        "aterrizar"
    )
    add_apa_code_block(doc, p4_code)

    p_e1_det = doc.add_paragraph()
    p_e1_det.paragraph_format.line_spacing = 1.5
    p_e1_det.paragraph_format.first_line_indent = Inches(0.5)
    p_e1_det.paragraph_format.space_after = Pt(6)
    r_e1_txt = p_e1_det.add_run(
        "Fase de detección: Fase 1 (Análisis Léxico).\n"
        "Mensaje del compilador: [Error Léxico] Línea 1, Columna 15: Símbolo desconocido o no permitido en DronLang: '@'.\n"
        "Justificación teórica: El símbolo arroba (@) no está contemplado en el alfabeto formal de DronLang. "
        "Cuando el cursor del analizador léxico procesa la línea 1, evalúa el carácter contra los patrones de números, "
        "letras, operadores y delimitadores válidos. Al no encontrar correspondencia con ningún patrón formal, "
        "el método tokenizar() de src/lexer.js aborta inmediatamente la ejecución lanzando la excepción correspondiente. "
        "El programa no avanza hacia el análisis sintáctico porque la cadena de tokens no logra completarse."
    )
    r_e1_txt.font.name = 'Times New Roman'
    r_e1_txt.font.size = Pt(12)
    r_e1_txt.font.color.rgb = BLACK

    # Error Sintáctico
    p_prog5_tit = doc.add_paragraph()
    p_prog5_tit.paragraph_format.space_before = Pt(10)
    p_prog5_tit.paragraph_format.space_after = Pt(2)
    p_prog5_tit.paragraph_format.line_spacing = 1.5
    r_p5t = p_prog5_tit.add_run("Programa 5 (Error Intencional): Presencia de Error Sintáctico")
    r_p5t.font.name = 'Times New Roman'
    r_p5t.font.size = Pt(12)
    r_p5t.font.bold = True
    r_p5t.font.color.rgb = BLACK

    p5_code = (
        "entregas = 3\n"
        "despegar\n"
        "repetir entregas veces\n"
        "    volar 12\n"
        "    entregar 1\n"
        "aterrizar"
    )
    add_apa_code_block(doc, p5_code)

    p_e2_det = doc.add_paragraph()
    p_e2_det.paragraph_format.line_spacing = 1.5
    p_e2_det.paragraph_format.first_line_indent = Inches(0.5)
    p_e2_det.paragraph_format.space_after = Pt(6)
    r_e2_txt = p_e2_det.add_run(
        "Fase de detección: Fase 2 (Análisis Sintáctico).\n"
        "Mensaje del compilador: [Error Sintáctico] Línea 6, Columna 1: Se esperaba 'fin' para cerrar la estructura 'repetir', pero se encontró 'aterrizar'.\n"
        "Justificación teórica: Desde el punto de vista léxico, todas las palabras y números empleados son válidos "
        "y generan tokens sin objeción. Sin embargo, en la fase sintáctica, el analizador recursivo descendente aplica "
        "la gramática formal del lenguaje, la cual exige de manera determinista que la regla de repetición concluya con el token 'fin'. "
        "Al recibir el token 'aterrizar' mientras la estructura continuaba abierta, el Parser identifica una oración incompleta "
        "y mal estructurada, deteniendo el compilador antes de pasar a la etapa semántica."
    )
    r_e2_txt.font.name = 'Times New Roman'
    r_e2_txt.font.size = Pt(12)
    r_e2_txt.font.color.rgb = BLACK

    # Error Semántico
    p_prog6_tit = doc.add_paragraph()
    p_prog6_tit.paragraph_format.space_before = Pt(10)
    p_prog6_tit.paragraph_format.space_after = Pt(2)
    p_prog6_tit.paragraph_format.line_spacing = 1.5
    r_p6t = p_prog6_tit.add_run("Programa 6 (Error Intencional): Presencia de Error Semántico")
    r_p6t.font.name = 'Times New Roman'
    r_p6t.font.size = Pt(12)
    r_p6t.font.bold = True
    r_p6t.font.color.rgb = BLACK

    p6_code = (
        "paquetes = 2\n"
        "despegar\n"
        "volar distanciaDestino\n"
        "entregar paquetes\n"
        "aterrizar"
    )
    add_apa_code_block(doc, p6_code)

    p_e3_det = doc.add_paragraph()
    p_e3_det.paragraph_format.line_spacing = 1.5
    p_e3_det.paragraph_format.first_line_indent = Inches(0.5)
    p_e3_det.paragraph_format.space_after = Pt(12)
    r_e3_txt = p_e3_det.add_run(
        "Fase de detección: Fase 3 (Análisis Semántico).\n"
        "Mensaje del compilador: [Error Semántico] Línea 3: La variable 'distanciaDestino' no ha sido declarada o inicializada.\n"
        "Justificación teórica: Este programa supera con éxito la fase léxica (el identificador cumple el patrón de letras válidas) "
        "y la fase sintáctica (la estructura volar <identificador> se ajusta a la regla gramatical). "
        "No obstante, al intervenir la fase semántica, el verificador consulta la Tabla de Símbolos y constata que 'distanciaDestino' "
        "no fue registrada mediante asignación previa. Aunque la sentencia está ortográfica y gramaticalmente bien construida, "
        "carece de sentido operativo en el mundo físico: el dron no dispone de una cantidad válida de metros para efectuar su trayectoria."
    )
    r_e3_txt.font.name = 'Times New Roman'
    r_e3_txt.font.size = Pt(12)
    r_e3_txt.font.color.rgb = BLACK

    # =========================================================================
    # CUADRO COMPARATIVO DE ERRORES (TABLA APA)
    # =========================================================================
    doc.add_page_break()

    p_t2_num = doc.add_paragraph()
    p_t2_num.paragraph_format.space_before = Pt(6)
    p_t2_num.paragraph_format.space_after = Pt(2)
    p_t2_num.paragraph_format.line_spacing = 1.5
    r_t2n = p_t2_num.add_run("Tabla 2")
    r_t2n.font.name = 'Times New Roman'
    r_t2n.font.size = Pt(12)
    r_t2n.font.bold = True
    r_t2n.font.color.rgb = BLACK

    p_t2_tit = doc.add_paragraph()
    p_t2_tit.paragraph_format.space_before = Pt(0)
    p_t2_tit.paragraph_format.space_after = Pt(6)
    p_t2_tit.paragraph_format.line_spacing = 1.5
    r_t2t = p_t2_tit.add_run("Cuadro Comparativo de los Tres Niveles de Error en el Compilador de DronLang")
    r_t2t.font.name = 'Times New Roman'
    r_t2t.font.size = Pt(12)
    r_t2t.font.italic = True
    r_t2t.font.color.rgb = BLACK

    table_comp = doc.add_table(rows=4, cols=4)
    table_comp.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths_t2 = [Inches(1.2), Inches(1.5), Inches(1.8), Inches(2.0)]
    
    headers_t2 = ["Nivel de Error", "Fase de Captura", "Causa en Código Fuente", "Consecuencia en el Compilador"]
    for c_idx, (h_text, w) in enumerate(zip(headers_t2, widths_t2)):
        cell = table_comp.cell(0, c_idx)
        cell.width = w
        set_cell_margins(cell, 80, 80, 80, 80)
        p = cell.paragraphs[0]
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h_text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = BLACK

    t2_rows = [
        ("Léxico", "Fase 1: Lexer", "Carácter '@' no admitido en el alfabeto.", "Aborto inmediato. No genera la tira de tokens completa ni pasa al Parser."),
        ("Sintáctico", "Fase 2: Parser", "Estructura 'repetir' sin palabra de cierre 'fin'.", "Tokens válidos, pero la frase viola la gramática formal. No genera AST."),
        ("Semántico", "Fase 3: Semantic", "Variable 'distanciaDestino' sin declarar ni inicializar.", "Gramática correcta, pero carece de significado y coherencia lógica en la misión.")
    ]

    for r_idx, (c0, c1, c2, c3) in enumerate(t2_rows, start=1):
        row = table_comp.rows[r_idx]
        for col_idx, (text, w) in enumerate(zip([c0, c1, c2, c3], widths_t2)):
            cell = row.cells[col_idx]
            cell.width = w
            set_cell_margins(cell, 60, 60, 80, 80)
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(10)
            run.font.bold = (col_idx == 0)
            run.font.color.rgb = BLACK

    apply_apa_table_borders(table_comp)

    # Nota Tabla 2
    p_t2_nota = doc.add_paragraph()
    p_t2_nota.paragraph_format.space_before = Pt(4)
    p_t2_nota.paragraph_format.space_after = Pt(18)
    p_t2_nota.paragraph_format.line_spacing = 1.15
    r_t2n_lbl = p_t2_nota.add_run("Nota. ")
    r_t2n_lbl.font.name = 'Times New Roman'
    r_t2n_lbl.font.size = Pt(10)
    r_t2n_lbl.font.italic = True
    r_t2n_lbl.font.color.rgb = BLACK
    
    r_t2n_txt = p_t2_nota.add_run(
        "Clasificación elaborada según las definiciones de las fases del compilador en el proyecto académico de Compiladores."
    )
    r_t2n_txt.font.name = 'Times New Roman'
    r_t2n_txt.font.size = Pt(10)
    r_t2n_txt.font.color.rgb = BLACK

    # =========================================================================
    # CONCLUSIONES Y REFERENCIAS
    # =========================================================================
    p_concl_h = doc.add_paragraph()
    p_concl_h.paragraph_format.space_before = Pt(12)
    p_concl_h.paragraph_format.space_after = Pt(6)
    p_concl_h.paragraph_format.line_spacing = 1.5
    r_ch = p_concl_h.add_run("Conclusiones")
    r_ch.font.name = 'Times New Roman'
    r_ch.font.size = Pt(12)
    r_ch.font.bold = True
    r_ch.font.color.rgb = BLACK

    p_concl_b = doc.add_paragraph()
    p_concl_b.paragraph_format.line_spacing = 1.5
    p_concl_b.paragraph_format.first_line_indent = Inches(0.5)
    p_concl_b.paragraph_format.space_after = Pt(18)
    r_cb = p_concl_b.add_run(
        "El diseño de DronLang responde con rigurosidad a los objetivos pedagógicos y técnicos de la Actividad 3. "
        "La separación modular de las fases permite desacoplar la validación léxica de la sintaxis y de la coherencia "
        "semántica del sistema de navegación aérea. Con este fundamento teórico y práctico concluido, el compilador "
        "se encuentra preparado para las fases de análisis sintáctico formal (Actividad 5) y la posterior generación "
        "definitiva de código de telemetría de vuelo (Actividad 7)."
    )
    r_cb.font.name = 'Times New Roman'
    r_cb.font.size = Pt(12)
    r_cb.font.color.rgb = BLACK

    # Referencias APA
    p_ref_h = doc.add_paragraph()
    p_ref_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ref_h.paragraph_format.space_before = Pt(12)
    p_ref_h.paragraph_format.space_after = Pt(12)
    p_ref_h.paragraph_format.line_spacing = 1.5
    r_rh = p_ref_h.add_run("Referencias")
    r_rh.font.name = 'Times New Roman'
    r_rh.font.size = Pt(12)
    r_rh.font.bold = True
    r_rh.font.color.rgb = BLACK

    p_ref1 = doc.add_paragraph()
    p_ref1.paragraph_format.line_spacing = 1.5
    p_ref1.paragraph_format.left_indent = Inches(0.5)
    p_ref1.paragraph_format.first_line_indent = Inches(-0.5) # Sangría francesa APA
    p_ref1.paragraph_format.space_after = Pt(6)
    r_r1 = p_ref1.add_run(
        "Aho, A. V., Lam, M. S., Sethi, R., & Ullman, J. D. (2008). "
        "Compiladores: Principios, técnicas y herramientas (2.ª ed.). Pearson Educación."
    )
    r_r1.font.name = 'Times New Roman'
    r_r1.font.size = Pt(12)
    r_r1.font.color.rgb = BLACK

    p_ref2 = doc.add_paragraph()
    p_ref2.paragraph_format.line_spacing = 1.5
    p_ref2.paragraph_format.left_indent = Inches(0.5)
    p_ref2.paragraph_format.first_line_indent = Inches(-0.5)
    p_ref2.paragraph_format.space_after = Pt(6)
    r_r2 = p_ref2.add_run(
        "Gamboa Berrío, H. (2026). Guía de proyecto: Mi lenguaje, mi compilador. "
        "Cátedra de Compiladores, Facultad de Ingeniería."
    )
    r_r2.font.name = 'Times New Roman'
    r_r2.font.size = Pt(12)
    r_r2.font.color.rgb = BLACK

    p_ref3 = doc.add_paragraph()
    p_ref3.paragraph_format.line_spacing = 1.5
    p_ref3.paragraph_format.left_indent = Inches(0.5)
    p_ref3.paragraph_format.first_line_indent = Inches(-0.5)
    p_ref3.paragraph_format.space_after = Pt(6)
    r_r3 = p_ref3.add_run(
        "Morales Cuesta, J. A., & Mosquera Rodríguez, S. A. (2026). "
        "Compilador DronLang: Repositorio oficial de código abierto [Software]. GitHub. "
        "https://github.com/andrescrack2024/Compilador"
    )
    r_r3.font.name = 'Times New Roman'
    r_r3.font.size = Pt(12)
    r_r3.font.color.rgb = BLACK

    # Guardar sobre el documento oficial
    output_path = os.path.join(os.getcwd(), "Entrega_Actividad_3_DronLang.docx")
    doc.save(output_path)
    print(f"Documento APA generado exitosamente en: {output_path}")

if __name__ == "__main__":
    build_apa_document()
