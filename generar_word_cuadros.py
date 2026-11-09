# -*- coding: utf-8 -*-
"""
Generador de documento Word oficial para la Actividad 3
Requerimientos explícitos del usuario:
- Todo en Times New Roman, tamaño 12 puntos
- Interlineado 1.5
- Sin colores (100% monocromático / blanco y negro)
- Cuadros / tablas de Word con todos los bordes (cuadrícula completa como la imagen de muestra)
- Títulos centrados
- Conclusiones en su propia página
- Referencias en su propia página
- Portada formal completa
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

BLACK = RGBColor(0, 0, 0)

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

def apply_full_grid_borders(table, border_sz="6"):
    """
    Aplica bordes negros a absolutamente todas las celdas de la tabla
    (estilo cuadrícula completa de Word como en la imagen del usuario).
    """
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:top w:val="single" w:sz="{border_sz}" w:space="0" w:color="000000"/>'
                f'<w:left w:val="single" w:sz="{border_sz}" w:space="0" w:color="000000"/>'
                f'<w:bottom w:val="single" w:sz="{border_sz}" w:space="0" w:color="000000"/>'
                f'<w:right w:val="single" w:sz="{border_sz}" w:space="0" w:color="000000"/>'
                f'</w:tcBorders>'
            )
            tcPr.append(tcBorders)

def add_centered_title(doc, text, space_before=12, space_after=12):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = BLACK
    return p

def add_body_paragraph(doc, text, space_after=10, indent=0.5):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    if indent > 0:
        p.paragraph_format.first_line_indent = Inches(indent)
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)
    run.font.color.rgb = BLACK
    return p

def add_word_code_box(doc, code_text):
    """Crea una caja de código como cuadro de Word con bordes completos."""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    cell = table.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_margins(cell, top=100, bottom=100, left=150, right=150)
    
    # Cuadrícula completa
    apply_full_grid_borders(table, border_sz="8")
    
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.15
    
    run = p.add_run(code_text.strip())
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)
    run.font.color.rgb = BLACK
    
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(6)

def build_document():
    doc = docx.Document()
    
    # Márgenes estándar APA (2.54 cm / 1 pulgada)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
        # Encabezado: número de página a la derecha en 12 pt Times New Roman
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        fldSimple = OxmlElement('w:fldSimple')
        fldSimple.set(qn('w:instr'), 'PAGE')
        hp._p.append(fldSimple)
        for r in hp.runs:
            r.font.name = 'Times New Roman'
            r.font.size = Pt(12)
            r.font.color.rgb = BLACK

    # =========================================================================
    # 1. PORTADA FORMAL (PÁGINA 1)
    # =========================================================================
    for _ in range(3):
        p_b = doc.add_paragraph()
        p_b.paragraph_format.line_spacing = 1.5
        p_b.paragraph_format.space_before = Pt(0)
        p_b.paragraph_format.space_after = Pt(0)

    # Título institucional
    p_inst_top = doc.add_paragraph()
    p_inst_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst_top.paragraph_format.line_spacing = 1.5
    p_inst_top.paragraph_format.space_after = Pt(18)
    r_it = p_inst_top.add_run("UNIVERSIDAD - FACULTAD DE INGENIERÍA\n"
                             "PROGRAMA DE INGENIERÍA DE SISTEMAS Y COMPUTACIÓN")
    r_it.font.name = 'Times New Roman'
    r_it.font.size = Pt(12)
    r_it.font.bold = True
    r_it.font.color.rgb = BLACK

    # Título del trabajo centrado y en negrita
    p_tit = doc.add_paragraph()
    p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tit.paragraph_format.line_spacing = 1.5
    p_tit.paragraph_format.space_before = Pt(18)
    p_tit.paragraph_format.space_after = Pt(24)
    r_t = p_tit.add_run("DISEÑO Y CONCEPTUALIZACIÓN DEL LENGUAJE DRONLANG\n"
                        "ACTIVIDAD 3: EXAMEN DE CONCEPTUALIZACIÓN BÁSICA\n"
                        "TEMÁTICA ASIGNADA: DRON DE REPARTO")
    r_t.font.name = 'Times New Roman'
    r_t.font.size = Pt(12)
    r_t.font.bold = True
    r_t.font.color.rgb = BLACK

    # Autores
    p_aut = doc.add_paragraph()
    p_aut.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_aut.paragraph_format.line_spacing = 1.5
    p_aut.paragraph_format.space_before = Pt(18)
    p_aut.paragraph_format.space_after = Pt(24)
    r_a = p_aut.add_run("AUTORES:\n"
                        "JAINER ANDRÉS MORALES CUESTA\n"
                        "SHARLY ANDRÉS MOSQUERA RODRÍGUEZ")
    r_a.font.name = 'Times New Roman'
    r_a.font.size = Pt(12)
    r_a.font.bold = True
    r_a.font.color.rgb = BLACK

    # Asignatura, Docente y Fecha
    p_bot = doc.add_paragraph()
    p_bot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_bot.paragraph_format.line_spacing = 1.5
    p_bot.paragraph_format.space_before = Pt(24)
    p_bot.paragraph_format.space_after = Pt(0)
    r_b = p_bot.add_run("ASIGNATURA: COMPILADORES\n"
                        "DOCENTE: ING. HASLER GAMBOA BERRÍO\n"
                        "9 DE NOVIEMBRE DE 2026\n\n"
                        "REPOSITORIO GITHUB: https://github.com/andrescrack2024/Compilador")
    r_b.font.name = 'Times New Roman'
    r_b.font.size = Pt(12)
    r_b.font.color.rgb = BLACK

    # =========================================================================
    # 2. INTRODUCCIÓN Y CATÁLOGO DE INSTRUCCIONES (PUNTO 1) - PÁGINA 2
    # =========================================================================
    doc.add_page_break()

    add_centered_title(doc, "1. INTRODUCCIÓN Y CATÁLOGO DE INSTRUCCIONES")

    add_body_paragraph(
        doc,
        "El presente informe expone el diseño formal, la especificación conceptual y la arquitectura "
        "del compilador para el lenguaje DronLang, correspondiente a la Actividad 3 del curso de Compiladores. "
        "La temática obligatoria asignada por el docente es Dron de reparto. DronLang ha sido estructurado "
        "para permitir la programación declarativa y secuencial de rutas de reparto, administración de paquetes "
        "y verificación del estado de vuelo y batería de vehículos aéreos no tripulados. "
        "Siguiendo lineamientos modernos de simplicidad inspirados en Python, el lenguaje no exige el uso de "
        "punto y coma (;) al final de las oraciones, delimita las instrucciones mediante saltos de línea y emplea "
        "la palabra clave fin para concluir de forma inequívoca las estructuras de control repetitivas y condicionales."
    )

    add_centered_title(doc, "Tabla 1. Catálogo Completo de Instrucciones de DronLang", space_before=6, space_after=6)

    # Tabla con cuadrícula completa estilo Word
    t_inst = doc.add_table(rows=8, cols=4)
    t_inst.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths_t1 = [Inches(1.2), Inches(1.8), Inches(2.2), Inches(1.3)]

    headers_t1 = ["INSTRUCCIÓN", "SINTAXIS", "¿QUÉ HACE EN EL DRON?", "EJEMPLO"]
    for idx, (h, w) in enumerate(zip(headers_t1, widths_t1)):
        cell = t_inst.cell(0, idx)
        cell.width = w
        set_cell_margins(cell, 80, 80, 80, 80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12)
        r.font.bold = True
        r.font.color.rgb = BLACK

    rows_data_t1 = [
        ("despegar", "despegar", "Enciende motores y realiza ascenso vertical automático a la altura base de crucero.", "despegar"),
        ("volar", "volar <expresion>", "Desplaza horizontalmente el dron hacia el destino la distancia indicada en metros (requiere valor >= 0).", "volar 15\nvolar distancia"),
        ("entregar", "entregar <expresion>", "Acciona el mecanismo de descarga para depositar la cantidad de paquetes indicada (requiere valor >= 0).", "entregar 1\nentregar paquetes"),
        ("aterrizar", "aterrizar", "Inicia descenso vertical controlado en la zona de aterrizaje y apaga los motores.", "aterrizar"),
        ("Asignación", "<id> = <expresion>", "Guarda en memoria una variable numérica o el cálculo de una expresión aritmética para su posterior utilización.", "bateria = 100\npaquetes = 2"),
        ("repetir", "repetir <expr> veces\n    <bloque>\nfin", "Estructura iterativa determinista. Ejecuta el bloque interno la cantidad de veces resultante de la expresión (entero >= 0).", "repetir 3 veces\n    volar 10\n    entregar 1\nfin"),
        ("si ... sino", "si <cond> entonces\n    <bloque1>\nsino\n    <bloque2>\nfin", "Estructura de bifurcación condicional. Evalúa una comparación relacional; si es verdadera ejecuta el primer bloque, si no el bloque alternativo.", "si bateria > 20 entonces\n    volar 30\nsino\n    aterrizar\nfin")
    ]

    for r_idx, row_values in enumerate(rows_data_t1, start=1):
        row = t_inst.rows[r_idx]
        for c_idx, (val, w) in enumerate(zip(row_values, widths_t1)):
            cell = row.cells[c_idx]
            cell.width = w
            set_cell_margins(cell, 60, 60, 80, 80)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(12)
            r.font.bold = (c_idx == 0)
            r.font.color.rgb = BLACK

    apply_full_grid_borders(t_inst, border_sz="6")

    # Operadores y reglas
    add_body_paragraph(
        doc,
        "Reglas complementarias de la gramática de DronLang:\n"
        "1. Operadores aritméticos: Suma (+), Resta (-), Multiplicación (*).\n"
        "2. Operadores relacionales de comparación: Mayor que (>), Menor que (<), Igualdad estricta (==).\n"
        "3. Delimitadores: Paréntesis '(' y ')' para agrupación de expresiones aritméticas.\n"
        "4. Comentarios: Se admiten comentarios de línea única antecedidos por '#' o '//'.\n"
        "5. Tipos de datos: Números enteros y con punto flotante.",
        space_after=12
    )

    # =========================================================================
    # 3. ARQUITECTURA Y FASES DEL COMPILADOR (PUNTO 2) - PÁGINA 3
    # =========================================================================
    doc.add_page_break()

    add_centered_title(doc, "2. ARQUITECTURA Y FASES DEL COMPILADOR")

    add_body_paragraph(
        doc,
        "La compilación de un programa en DronLang comprende un cauce estructurado en cuatro fases sucesivas. "
        "Cada etapa asume una responsabilidad específica de verificación y transformación, convirtiendo "
        "las cadenas de caracteres iniciales en comandos de vuelo precisos y seguros. "
        "En la Tabla 2 se ilustra el esquema de bloques y flujo de datos entre las fases."
    )

    add_centered_title(doc, "Tabla 2. Flujo de Datos y Fases del Compilador DronLang", space_before=4, space_after=6)

    t_flow = doc.add_table(rows=9, cols=3)
    t_flow.alignment = WD_TABLE_ALIGNMENT.CENTER
    w_flow = [Inches(1.8), Inches(3.2), Inches(1.5)]

    flow_data = [
        ("ENTRADA", "CÓDIGO FUENTE (.dron)", "Archivo de texto plano"),
        ("Flujo de Caracteres", "Lectura secuencial carácter a carácter", "Cursor pos/línea/col"),
        ("FASE 1", "ANÁLISIS LÉXICO (LEXER)", "Generación de Tokens"),
        ("Tira de Tokens", "Secuencia: {tipo, valor, línea, col}", "Sin espacios ni notas"),
        ("FASE 2", "ANÁLISIS SINTÁCTICO (PARSER)", "Gramática y estructura"),
        ("Árbol Jerárquico", "Árbol de Sintaxis Abstracta (AST)", "Nodos de operaciones"),
        ("FASE 3", "ANÁLISIS SEMÁNTICO (CHECKER)", "Validación lógica"),
        ("AST Validado", "Tabla de Símbolos y Tipos Verificados", "Variables y rangos >= 0"),
        ("FASE 4", "GENERACIÓN DE CÓDIGO (CODEGEN)", "Comandos de vuelo finales")
    ]

    for r_idx, (c0, c1, c2) in enumerate(flow_data):
        row = t_flow.rows[r_idx]
        for col_idx, (text, w) in enumerate(zip([c0, c1, c2], w_flow)):
            cell = row.cells[col_idx]
            cell.width = w
            set_cell_margins(cell, 60, 60, 80, 80)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(text)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(12)
            r.font.bold = (col_idx == 0 or r_idx in (2, 4, 6, 8))
            r.font.color.rgb = BLACK

    apply_full_grid_borders(t_flow, border_sz="6")

    # Una frase por fase
    add_centered_title(doc, "Descripción Conceptual: Una Frase por Fase en DronLang", space_before=12, space_after=6)

    phrases = [
        ("Fase 1 (Análisis Léxico):", 
         "Lee secuencialmente el texto del código fuente y agrupa los caracteres en unidades atómicas denominadas tokens "
         "(como la palabra clave despegar, identificadores, números u operadores), descartando comentarios y espacios en blanco, "
         "y deteniendo el proceso de inmediato si detecta un carácter no reconocido en el alfabeto del lenguaje."),
        
        ("Fase 2 (Análisis Sintáctico):", 
         "Verifica que la sucesión de tokens se ajuste fielmente a las reglas de la gramática libre de contexto de DronLang "
         "(asegurando, por ejemplo, que a toda instrucción repetir le acompañe la palabra veces y concluya con fin), "
         "construyendo a partir de ella el Árbol de Sintaxis Abstracta (AST)."),
        
        ("Fase 3 (Análisis Semántico):", 
         "Valida la coherencia lógica de las instrucciones mediante la Tabla de Símbolos, certificando que ninguna variable "
         "sea empleada antes de haberse asignado y verificando que los argumentos de distancia, repeticiones y paquetes "
         "no tomen valores numéricos negativos."),
        
        ("Fase 4 (Generación de Código):", 
         "Recorre el árbol de sintaxis abstracta validado y lo traduce en una lista de comandos de telemetría de bajo nivel "
         "(DESPEGAR, VOLAR, ENTREGAR, ATERRIZAR) para la ejecución en el hardware o la simulación interactiva.")
    ]

    for title, phrase in phrases:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.first_line_indent = Inches(0.5)
        
        r_t = p.add_run(f"{title} ")
        r_t.font.name = 'Times New Roman'
        r_t.font.size = Pt(12)
        r_t.font.bold = True
        r_t.font.color.rgb = BLACK
        
        r_p = p.add_run(phrase)
        r_p.font.name = 'Times New Roman'
        r_p.font.size = Pt(12)
        r_p.font.color.rgb = BLACK

    # =========================================================================
    # 4. BANCO DE PROGRAMAS DE PRUEBA (PUNTO 3) - PÁGINA 4
    # =========================================================================
    doc.add_page_break()

    add_centered_title(doc, "3. BANCO DE SEIS PROGRAMAS DE PRUEBA")

    add_body_paragraph(
        doc,
        "A continuación se documenta el conjunto de seis programas desarrollados para comprobar el comportamiento "
        "integral del compilador. Los primeros tres programas corresponden a misiones legítimas que se ejecutan "
        "satisfactoriamente en todas las fases. Los tres programas subsiguientes contienen fallas intencionales "
        "concebidas para demostrar la capacidad de diagnóstico en cada nivel de análisis."
    )

    # Programa 1
    add_centered_title(doc, "Programa 1 (Correcto): Misión Básica de Reparto Punto a Punto", space_before=6, space_after=4)
    code_1 = (
        "distancia = 15\n"
        "paquetes = 2\n\n"
        "despegar\n"
        "volar distancia\n"
        "entregar paquetes\n"
        "volar distancia\n"
        "aterrizar"
    )
    add_word_code_box(doc, code_1)
    add_body_paragraph(
        doc,
        "Explicación funcional: Asigna en memoria las variables distancia (15 m) y paquetes (2 unidades). "
        "El dron despega, efectúa el vuelo horizontal de 15 metros hasta las coordenadas destino, libera los 2 paquetes "
        "mediante el actuador de entrega, regresa a la base recorriendo la misma distancia y aterriza sin incidencias. "
        "Supera de manera óptima las cuatro fases del compilador."
    )

    # Programa 2
    add_centered_title(doc, "Programa 2 (Correcto): Condición de Autonomía de Batería con Operaciones", space_before=8, space_after=4)
    code_2 = (
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
    add_word_code_box(doc, code_2)
    add_body_paragraph(
        doc,
        "Explicación funcional: Emplea multiplicación (10 * 3 = 30) y resta (75 - 30 = 45) en expresiones. "
        "Evalúa la estructura condicional si ... entonces ... sino ... fin. Como la condición 45 > 30 resulta "
        "verdadera, el compilador genera la rama de vuelo largo (20 m) y descarga de 1 paquete antes de aterrizar."
    )

    # Programa 3
    add_centered_title(doc, "Programa 3 (Correcto): Reparto Cíclico en Múltiples Zonas con Repetición", space_before=8, space_after=4)
    code_3 = (
        "zonas = 4\n"
        "distanciaSector = 5 + 3\n\n"
        "despegar\n"
        "repetir zonas veces\n"
        "    volar distanciaSector\n"
        "    entregar 1\n"
        "fin\n"
        "aterrizar"
    )
    add_word_code_box(doc, code_3)
    add_body_paragraph(
        doc,
        "Explicación funcional: Integra cálculo de suma (5 + 3 = 8 m) con el ciclo repetitivo repetir ... veces ... fin. "
        "El dron itera 4 veces sobre el bloque interior realizando el vuelo de 8 metros y descarga de 1 paquete por zona. "
        "Una vez satisfecho el conteo de repeticiones, concluye con aterrizaje seguro."
    )

    # =========================================================================
    # 5. DIAGNÓSTICO DE ERRORES Y TABLA COMPARATIVA (PUNTO 4) - PÁGINA 5
    # =========================================================================
    doc.add_page_break()

    add_centered_title(doc, "4. DIAGNÓSTICO Y JUSTIFICACIÓN TÉCNICA DE ERRORES")

    add_body_paragraph(
        doc,
        "Para responder con precisión a la evaluación de la Actividad 3, se detallan a continuación "
        "los tres programas con errores intencionales. Cada caso puntualiza el código erróneo, "
        "la fase exacta donde se intercepta la anomalía, el reporte técnico generado y la fundamentación "
        "de por qué dicho fallo no puede ser admitido por las fases posteriores."
    )

    # Error Léxico
    add_centered_title(doc, "Programa 4 (Error Intencional): Presencia de Error Léxico", space_before=6, space_after=4)
    code_4 = (
        "bateria = 100 @\n"
        "despegar\n"
        "volar 10\n"
        "aterrizar"
    )
    add_word_code_box(doc, code_4)

    # Cuadro de diagnóstico 1
    t_diag1 = doc.add_table(rows=3, cols=2)
    t_diag1.alignment = WD_TABLE_ALIGNMENT.CENTER
    diag1_data = [
        ("FASE DE DETECCIÓN", "Fase 1: Análisis Léxico (Lexer)"),
        ("MENSAJE DEL COMPILADOR", "[Error Léxico] Línea 1, Columna 15: Símbolo desconocido o no permitido en DronLang: '@'"),
        ("JUSTIFICACIÓN TÉCNICA", 
         "El símbolo '@' no forma parte del alfabeto formal de DronLang. Cuando el analizador léxico procesa el "
         "texto en src/lexer.js, evalúa el carácter contra los patrones de números, letras, operadores y delimitadores. "
         "Al no encajar en ninguna regla válida, la función tokenizar() arroja la excepción inmediatamente, impidiendo "
         "que se complete la tira de tokens. El código jamás llega a la etapa sintáctica.")
    ]
    for r_i, (k, v) in enumerate(diag1_data):
        row = t_diag1.rows[r_i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = Inches(2.2), Inches(4.3)
        set_cell_margins(c0, 60, 60, 80, 80)
        set_cell_margins(c1, 60, 60, 80, 80)
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.line_spacing = 1.15
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(k)
        r0.font.name = 'Times New Roman'
        r0.font.size = Pt(12)
        r0.font.bold = True
        r0.font.color.rgb = BLACK
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.line_spacing = 1.15
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(v)
        r1.font.name = 'Times New Roman'
        r1.font.size = Pt(12)
        r1.font.color.rgb = BLACK
    apply_full_grid_borders(t_diag1, border_sz="6")

    # Error Sintáctico
    add_centered_title(doc, "Programa 5 (Error Intencional): Presencia de Error Sintáctico", space_before=10, space_after=4)
    code_5 = (
        "entregas = 3\n"
        "despegar\n"
        "repetir entregas veces\n"
        "    volar 12\n"
        "    entregar 1\n"
        "aterrizar"
    )
    add_word_code_box(doc, code_5)

    t_diag2 = doc.add_table(rows=3, cols=2)
    t_diag2.alignment = WD_TABLE_ALIGNMENT.CENTER
    diag2_data = [
        ("FASE DE DETECCIÓN", "Fase 2: Análisis Sintáctico (Parser)"),
        ("MENSAJE DEL COMPILADOR", "[Error Sintáctico] Línea 6, Columna 1: Se esperaba 'fin' para cerrar la estructura 'repetir', pero se encontró 'aterrizar'"),
        ("JUSTIFICACIÓN TÉCNICA", 
         "Todos los caracteres y palabras del programa pertenecen al lenguaje, por lo que la Fase Léxica genera los tokens "
         "sin problemas. Sin embargo, en la Fase Sintáctica, el analizador comprueba la gramática del programa: la regla "
         "de producción del bucle exige determinísticamente el token de cierre 'fin'. Al encontrar 'aterrizar' sin haberse "
         "cerrado el bloque, el Parser detecta una estructura mal armada y aborta la construcción del árbol AST.")
    ]
    for r_i, (k, v) in enumerate(diag2_data):
        row = t_diag2.rows[r_i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = Inches(2.2), Inches(4.3)
        set_cell_margins(c0, 60, 60, 80, 80)
        set_cell_margins(c1, 60, 60, 80, 80)
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.line_spacing = 1.15
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(k)
        r0.font.name = 'Times New Roman'
        r0.font.size = Pt(12)
        r0.font.bold = True
        r0.font.color.rgb = BLACK
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.line_spacing = 1.15
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(v)
        r1.font.name = 'Times New Roman'
        r1.font.size = Pt(12)
        r1.font.color.rgb = BLACK
    apply_full_grid_borders(t_diag2, border_sz="6")

    # Error Semántico
    add_centered_title(doc, "Programa 6 (Error Intencional): Presencia de Error Semántico", space_before=10, space_after=4)
    code_6 = (
        "paquetes = 2\n"
        "despegar\n"
        "volar distanciaDestino\n"
        "entregar paquetes\n"
        "aterrizar"
    )
    add_word_code_box(doc, code_6)

    t_diag3 = doc.add_table(rows=3, cols=2)
    t_diag3.alignment = WD_TABLE_ALIGNMENT.CENTER
    diag3_data = [
        ("FASE DE DETECCIÓN", "Fase 3: Análisis Semántico (Semantic Analyzer)"),
        ("MENSAJE DEL COMPILADOR", "[Error Semántico] Línea 3: La variable 'distanciaDestino' no ha sido declarada o inicializada"),
        ("JUSTIFICACIÓN TÉCNICA", 
         "El código es léxicamente intachable ('distanciaDestino' es un identificador legal) y sintácticamente válido "
         "(la regla volar <id> es correcta). No obstante, en la Fase Semántica, el verificador consulta la Tabla de Símbolos "
         "y comprueba que 'distanciaDestino' no fue asignada previamente con un valor numérico. La orden carece de sentido "
         "físico: el dron recibiría una orden de desplazamiento con distancia indeterminada.")
    ]
    for r_i, (k, v) in enumerate(diag3_data):
        row = t_diag3.rows[r_i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = Inches(2.2), Inches(4.3)
        set_cell_margins(c0, 60, 60, 80, 80)
        set_cell_margins(c1, 60, 60, 80, 80)
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.line_spacing = 1.15
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(k)
        r0.font.name = 'Times New Roman'
        r0.font.size = Pt(12)
        r0.font.bold = True
        r0.font.color.rgb = BLACK
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.line_spacing = 1.15
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(v)
        r1.font.name = 'Times New Roman'
        r1.font.size = Pt(12)
        r1.font.color.rgb = BLACK
    apply_full_grid_borders(t_diag3, border_sz="6")

    # Tabla Comparativa de Cierre
    add_centered_title(doc, "Tabla 3. Cuadro Comparativo de los Tres Niveles de Error en Compilación", space_before=12, space_after=6)

    t_comp = doc.add_table(rows=4, cols=4)
    t_comp.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths_t3 = [Inches(1.2), Inches(1.5), Inches(1.8), Inches(2.0)]
    headers_t3 = ["NIVEL DE ERROR", "FASE RESPONSABLE", "ORIGEN DEL FALLO", "IMPACTO EN EL COMPILADOR"]

    for idx, (h, w) in enumerate(zip(headers_t3, widths_t3)):
        cell = t_comp.cell(0, idx)
        cell.width = w
        set_cell_margins(cell, 80, 80, 80, 80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12)
        r.font.bold = True
        r.font.color.rgb = BLACK

    rows_comp_data = [
        ("Léxico", "Fase 1 (Lexer)", "Carácter '@' ajeno al alfabeto del lenguaje.", "Aborto inmediato en lectura. No genera la lista de tokens ni pasa al Parser."),
        ("Sintáctico", "Fase 2 (Parser)", "Omisión de la palabra de cierre 'fin' en 'repetir'.", "Tokens válidos, pero la estructura infringe la gramática formal. No genera AST."),
        ("Semántico", "Fase 3 (Semántico)", "Uso de 'distanciaDestino' no registrada en memoria.", "Estructura correcta, pero carece de significado lógico y operativo en el dron.")
    ]

    for r_idx, row_values in enumerate(rows_comp_data, start=1):
        row = t_comp.rows[r_idx]
        for c_idx, (val, w) in enumerate(zip(row_values, widths_t3)):
            cell = row.cells[c_idx]
            cell.width = w
            set_cell_margins(cell, 60, 60, 80, 80)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(12)
            r.font.bold = (c_idx == 0)
            r.font.color.rgb = BLACK

    apply_full_grid_borders(t_comp, border_sz="6")

    # =========================================================================
    # 6. CONCLUSIONES (EN SU PROPIA PÁGINA)
    # =========================================================================
    doc.add_page_break()

    add_centered_title(doc, "CONCLUSIONES", space_before=18, space_after=18)

    add_body_paragraph(
        doc,
        "Primera. El diseño y desarrollo de DronLang demuestra la importancia de concebir una sintaxis limpia "
        "y un catálogo riguroso de instrucciones antes de abordar la codificación del compilador. La temática "
        "asignada de dron de reparto permitió contextualizar de forma concreta primitivas físicas de control aéreo "
        "(despegar, volar, entregar, aterrizar) combinadas con estructuras estándar de computación como variables, "
        "iteraciones y bifurcaciones condicionales.",
        space_after=14
    )

    add_body_paragraph(
        doc,
        "Segunda. La separación modular de las fases de un compilador (léxica, sintáctica, semántica y generación "
        "de código) es indispensable para el diagnóstico unívoco de fallas. Cada fase cumple una frontera "
        "específica: el analizador léxico valida símbolos y caracteres individuales, el analizador sintáctico "
        "garantiza el orden oracional de las producciones gramaticales, y el analizador semántico certifica "
        "que el programa tenga sentido físico y lógico mediante la Tabla de Símbolos.",
        space_after=14
    )

    add_body_paragraph(
        doc,
        "Tercera. La experimentación con los tres programas de error intencional corroboró que ningún fallo puede ser "
        "atendido prematuramente ni postergado indebidamente: un error léxico detiene el proceso antes del parser; "
        "un error sintáctico impide la formación del árbol AST; y un error semántico salvaguarda la seguridad física "
        "del dron, impidiendo que despegue con parámetros indeterminados o negativos.",
        space_after=14
    )

    add_body_paragraph(
        doc,
        "Cuarta. El proyecto cumple a cabalidad con la totalidad de los requisitos pedagógicos fijados para "
        "la Actividad 3: catálogo exhaustivo de instrucciones, diagrama de arquitectura de fases, formulación "
        "de una frase por fase, seis programas de prueba verificados y diagnóstico técnico detallado, "
        "dejando una base sólida para las entregas posteriores de la asignatura.",
        space_after=14
    )

    # =========================================================================
    # 7. REFERENCIAS (EN SU PROPIA PÁGINA)
    # =========================================================================
    doc.add_page_break()

    add_centered_title(doc, "REFERENCIAS", space_before=18, space_after=18)

    references = [
        "Aho, A. V., Lam, M. S., Sethi, R., & Ullman, J. D. (2008). Compiladores: Principios, técnicas y herramientas (2.ª ed.). Pearson Educación.",
        "Cooper, K. D., & Torczon, L. (2012). Engineering a compiler (2nd ed.). Morgan Kaufmann.",
        "Gamboa Berrío, H. (2026). Guía de proyecto: Mi lenguaje, mi compilador. Cátedra de Compiladores, Facultad de Ingeniería.",
        "Louden, K. C. (2004). Construcción de compiladores: Principios y práctica. Thomson Learning.",
        "Morales Cuesta, J. A., & Mosquera Rodríguez, S. A. (2026). DronLang: Compilador para drones de reparto autónomos [Código abierto]. GitHub. https://github.com/andrescrack2024/Compilador"
    ]

    for ref in references:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.left_indent = Inches(0.5)
        p.paragraph_format.first_line_indent = Inches(-0.5) # Sangría francesa
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(12)
        
        r = p.add_run(ref)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12)
        r.font.color.rgb = BLACK

    output_path = os.path.join(os.getcwd(), "Entrega_Actividad_3_DronLang.docx")
    try:
        doc.save(output_path)
        print(f"Documento con cuadros Word y páginas dedicadas guardado en: {output_path}")
    except PermissionError:
        output_alt = os.path.join(os.getcwd(), "Entrega_Actividad_3_DronLang_Final.docx")
        doc.save(output_alt)
        print(f"El archivo principal estaba abierto en Word. Guardado exitosamente en: {output_alt}")

if __name__ == "__main__":
    build_document()
