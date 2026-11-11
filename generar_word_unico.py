# -*- coding: utf-8 -*-
"""
Generador del Documento Oficial ÚNICO para la Actividad 3
Archivo de salida: Entrega_Actividad_3_DronLang.docx

Requisitos estrictos solicitados:
1. UN SOLO documento Word.
2. Todo en Times New Roman, tamaño 12 puntos estricto (texto, tablas, código, títulos).
3. Interlineado de 1.5 líneas en todos los párrafos.
4. Sin colores: 100% monocromático / blanco y negro (texto negro #000000).
5. Cuadros / tablas con cuadrícula completa estilo Word (todos los bordes activos).
6. Títulos y subtítulos centrados en negrita.
7. Conclusiones en su propia página dedicada.
8. Referencias en su propia página dedicada.
9. Estructura altamente organizada y profesional.
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

BLACK = RGBColor(0, 0, 0)

def set_cell_padding(cell, top=100, bottom=100, left=150, right=150):
    """Establece márgenes internos (padding) para que el texto no toque los bordes."""
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
    Aplica cuadrícula completa de Word (todos los bordes negros:
    superior, inferior, izquierdo, derecho en cada celda).
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

def add_centered_heading(doc, text, space_before=14, space_after=10):
    """Agrega un título o subtítulo centrado en negrita, Times New Roman 12."""
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
    """Párrafo de cuerpo de texto: Times New Roman 12, interlineado 1.5."""
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

def add_code_box(doc, code_text):
    """Caja de código con bordes completos estilo cuadro de Word."""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    cell = table.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_padding(cell, top=120, bottom=120, left=160, right=160)
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

def build_single_official_document():
    doc = docx.Document()
    
    # Configuración de márgenes estándar APA (2.54 cm / 1 pulgada)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
        # Encabezado: número de página alineado a la derecha en Times New Roman 12
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
    # PÁGINA 1: PORTADA FORMAL Y BIEN ELABORADA
    # =========================================================================
    for _ in range(2):
        pb = doc.add_paragraph()
        pb.paragraph_format.line_spacing = 1.5
        pb.paragraph_format.space_before = Pt(0)
        pb.paragraph_format.space_after = Pt(0)

    p_uni = doc.add_paragraph()
    p_uni.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_uni.paragraph_format.line_spacing = 1.5
    p_uni.paragraph_format.space_after = Pt(20)
    r_uni = p_uni.add_run(
        "UNIVERSIDAD TECNOLÓGICA DEL CHOCÓ\n"
        "FACULTAD DE INGENIERÍA\n"
        "PROGRAMA DE INGENIERÍA DE SISTEMAS Y COMPUTACIÓN"
    )
    r_uni.font.name = 'Times New Roman'
    r_uni.font.size = Pt(12)
    r_uni.font.bold = True
    r_uni.font.color.rgb = BLACK

    # Título principal del trabajo
    p_tit = doc.add_paragraph()
    p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tit.paragraph_format.line_spacing = 1.5
    p_tit.paragraph_format.space_before = Pt(18)
    p_tit.paragraph_format.space_after = Pt(24)
    r_t = p_tit.add_run(
        "DISEÑO Y CONCEPTUALIZACIÓN DEL LENGUAJE DRONLANG\n"
        "ACTIVIDAD 3: EXAMEN DE CONCEPTUALIZACIÓN BÁSICA\n"
        "TEMÁTICA ASIGNADA: DRON DE REPARTO"
    )
    r_t.font.name = 'Times New Roman'
    r_t.font.size = Pt(12)
    r_t.font.bold = True
    r_t.font.color.rgb = BLACK

    # Autores del equipo de trabajo
    p_aut = doc.add_paragraph()
    p_aut.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_aut.paragraph_format.line_spacing = 1.5
    p_aut.paragraph_format.space_before = Pt(18)
    p_aut.paragraph_format.space_after = Pt(24)
    r_a = p_aut.add_run(
        "INTEGRANTES DEL EQUIPO:\n"
        "JAINER ANDRÉS MORALES CUESTA\n"
        "SHARLY ANDRÉS MOSQUERA RODRÍGUEZ"
    )
    r_a.font.name = 'Times New Roman'
    r_a.font.size = Pt(12)
    r_a.font.bold = True
    r_a.font.color.rgb = BLACK

    # Docente, asignatura, fecha y enlace del repositorio
    p_pie = doc.add_paragraph()
    p_pie.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_pie.paragraph_format.line_spacing = 1.5
    p_pie.paragraph_format.space_before = Pt(24)
    p_pie.paragraph_format.space_after = Pt(0)
    r_p = p_pie.add_run(
        "ASIGNATURA: COMPILADORES\n"
        "DOCENTE: ING. HASLER GAMBOA BERRÍO\n"
        "QUIBDÓ, CHOCÓ - NOVIEMBRE DE 2026\n\n"
        "REPOSITORIO OFICIAL GITHUB:\n"
        "https://github.com/andrescrack2024/Compilador"
    )
    r_p.font.name = 'Times New Roman'
    r_p.font.size = Pt(12)
    r_p.font.color.rgb = BLACK

    # =========================================================================
    # PÁGINA 2: INTRODUCCIÓN Y CATÁLOGO DE INSTRUCCIONES (PUNTO 1)
    # =========================================================================
    doc.add_page_break()

    add_centered_heading(doc, "1. INTRODUCCIÓN Y CATÁLOGO DE INSTRUCCIONES")

    add_body_paragraph(
        doc,
        "El presente documento recoge la entrega formal de la Actividad 3: Examen de Conceptualización Básica "
        "para la asignatura Compiladores. Siguiendo las directrices del docente evaluador, a este grupo de trabajo "
        "le fue asignada de manera inmodificable la temática Dron de reparto. A partir de ella se concibió DronLang, "
        "un lenguaje de dominio específico diseñado para definir misiones de entrega autónoma de paquetes, "
        "gestión de energía y trayectorias aéreas seguras."
    )

    add_body_paragraph(
        doc,
        "La sintaxis de DronLang adopta un modelo limpio e intuitivo similar a Python: las instrucciones se separan "
        "mediante saltos de línea (sin requerir punto y coma al final), los bloques de control se cierran de forma determinista "
        "con la palabra reservada fin, y se aceptan comentarios explicativos mediante los prefijos # o //."
    )

    add_centered_heading(doc, "Tabla 1. Catálogo Completo de Instrucciones de DronLang", space_before=8, space_after=6)

    # Tabla 1 con cuadrícula completa
    t1 = doc.add_table(rows=8, cols=4)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths_t1 = [Inches(1.2), Inches(1.8), Inches(2.2), Inches(1.3)]
    headers_t1 = ["INSTRUCCIÓN", "SINTAXIS", "¿QUÉ HACE EN EL DRON?", "EJEMPLO"]

    for idx, (h, w) in enumerate(zip(headers_t1, widths_t1)):
        cell = t1.cell(0, idx)
        cell.width = w
        set_cell_padding(cell, 80, 80, 80, 80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = BLACK

    data_t1 = [
        ("despegar", "despegar", "Enciende motores y realiza el ascenso vertical a la altitud base de crucero.", "despegar"),
        ("volar", "volar <expresion>", "Desplaza horizontalmente el dron en línea recta la cantidad de metros indicada (>= 0).", "volar 15\nvolar distancia"),
        ("entregar", "entregar <expresion>", "Acciona el actuador para depositar en el destino la cantidad de paquetes indicada (>= 0).", "entregar 1\nentregar paquetes"),
        ("aterrizar", "aterrizar", "Ejecuta el descenso vertical controlado en la zona de aterrizaje y apaga los motores.", "aterrizar"),
        ("Asignación", "<id> = <expresion>", "Almacena en memoria una variable numérica o el resultado de evaluar operaciones aritméticas.", "bateria = 100\npaquetes = 2"),
        ("repetir", "repetir <expr> veces\n    <bloque>\nfin", "Estructura iterativa determinista. Repite el bloque de instrucciones la cantidad de veces especificada (entero >= 0).", "repetir 3 veces\n    volar 10\n    entregar 1\nfin"),
        ("si ... sino", "si <cond> entonces\n    <bloque1>\nsino\n    <bloque2>\nfin", "Estructura condicional. Evalúa una comparación binaria; ejecuta el primer bloque si es verdadera o la alternativa si no.", "si bateria > 20 entonces\n    volar 30\nsino\n    aterrizar\nfin")
    ]

    for r_idx, row_vals in enumerate(data_t1, start=1):
        row = t1.rows[r_idx]
        for c_idx, (val, w) in enumerate(zip(row_vals, widths_t1)):
            cell = row.cells[c_idx]
            cell.width = w
            set_cell_padding(cell, 60, 60, 80, 80)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            run.font.bold = (c_idx == 0)
            run.font.color.rgb = BLACK

    apply_full_grid_borders(t1, border_sz="6")

    add_body_paragraph(
        doc,
        "Operaciones y elementos complementarios admitidos en DronLang:\n"
        "• Operaciones aritméticas: Suma (+), Resta (-) y Multiplicación (*).\n"
        "• Operaciones de comparación relacional: Mayor que (>), Menor que (<) e Igualdad estricta (==).\n"
        "• Delimitadores: Paréntesis '(' y ')' para alterar el orden de evaluación en las expresiones.\n"
        "• Tipos de datos: Literales numéricos enteros y decimales (punto flotante).\n"
        "• Comentarios: Líneas antecedidas por # o // que el compilador ignora.",
        space_after=12
    )

    # =========================================================================
    # PÁGINA 3: ARQUITECTURA Y FASES DEL COMPILADOR (PUNTO 2)
    # =========================================================================
    doc.add_page_break()

    add_centered_heading(doc, "2. ARQUITECTURA Y FASES DEL COMPILADOR")

    add_body_paragraph(
        doc,
        "El compilador de DronLang opera mediante un cauce secuencial de cuatro fases claramente diferenciadas. "
        "Cada etapa cumple una función específica de verificación y traducción técnica, recibiendo como entrada "
        "el resultado de la fase anterior y entregando una estructura validada para la siguiente. "
        "En la Tabla 2 se ilustra el esquema de procesamiento completo."
    )

    add_centered_heading(doc, "Tabla 2. Diagrama de Procesamiento y Fases del Compilador", space_before=4, space_after=6)

    t2 = doc.add_table(rows=9, cols=3)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths_t2 = [Inches(1.8), Inches(3.2), Inches(1.5)]
    data_t2 = [
        ("ENTRADA", "CÓDIGO FUENTE (.dron)", "Texto plano del programa"),
        ("Flujo de Caracteres", "Lectura secuencial carácter a carácter", "Control de posición"),
        ("FASE 1", "ANÁLISIS LÉXICO (LEXER)", "Generación de Tokens"),
        ("Tira de Tokens", "Colección: {tipo, valor, línea, col}", "Sin espacios ni notas"),
        ("FASE 2", "ANÁLISIS SINTÁCTICO (PARSER)", "Gramática formal"),
        ("Árbol Sintáctico", "Árbol de Sintaxis Abstracta (AST)", "Estructura jerárquica"),
        ("FASE 3", "ANÁLISIS SEMÁNTICO (CHECKER)", "Validación lógica"),
        ("AST Validado", "Tabla de Símbolos y Reglas de Vuelo", "Variables y rangos >= 0"),
        ("FASE 4", "GENERACIÓN DE CÓDIGO (CODEGEN)", "Comandos de vuelo")
    ]

    for r_idx, row_vals in enumerate(data_t2):
        row = t2.rows[r_idx]
        for c_idx, (text, w) in enumerate(zip(row_vals, widths_t2)):
            cell = row.cells[c_idx]
            cell.width = w
            set_cell_padding(cell, 60, 60, 80, 80)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            run.font.bold = (c_idx == 0 or r_idx in (2, 4, 6, 8))
            run.font.color.rgb = BLACK

    apply_full_grid_borders(t2, border_sz="6")

    add_centered_heading(doc, "Descripción Conceptual: Una Frase por Fase en DronLang", space_before=14, space_after=8)

    phrases = [
        ("Fase 1 (Análisis Léxico):", 
         "Lee secuencialmente los caracteres del código fuente del dron y los agrupa en unidades atómicas denominadas tokens "
         "(como la palabra clave despegar, números y operadores), descartando comentarios y espacios en blanco, "
         "y deteniendo el proceso de inmediato si encuentra un símbolo no reconocido en el alfabeto del lenguaje."),
        
        ("Fase 2 (Análisis Sintáctico):", 
         "Verifica que la sucesión de tokens se organice cumpliendo estrictamente la gramática formal de DronLang "
         "(garantizando, por ejemplo, que a la instrucción repetir le acompañe la palabra veces y termine obligatoriamente en fin), "
         "construyendo a partir de ella el Árbol de Sintaxis Abstracta (AST)."),
        
        ("Fase 3 (Análisis Semántico):", 
         "Comprueba el significado lógico de las instrucciones mediante la Tabla de Símbolos, verificando que ninguna "
         "variable sea empleada sin previa inicialización y certificando que valores críticos como distancias, repeticiones "
         "y paquetes no adopten cantidades negativas."),
        
        ("Fase 4 (Generación de Código):", 
         "Traduce el árbol sintáctico validado en una lista de instrucciones de telemetría de bajo nivel "
         "(DESPEGAR, VOLAR, ENTREGAR, ATERRIZAR) para comandar el dron físico o alimentar el simulador interactivo.")
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
    # PÁGINA 4: BANCO DE SEIS PROGRAMAS DE PRUEBA (PUNTO 3)
    # =========================================================================
    doc.add_page_break()

    add_centered_heading(doc, "3. BANCO DE SEIS PROGRAMAS DE PRUEBA")

    add_body_paragraph(
        doc,
        "A continuación se presentan los seis programas desarrollados para comprobar el funcionamiento "
        "integral del compilador. Los primeros tres programas corresponden a misiones válidas que ilustran "
        "el uso correcto de variables, cálculos, estructuras repetitivas y bifurcaciones condicionales. "
        "Los programas 4, 5 y 6 contienen anomalías deliberadas diseñadas para demostrar el diagnóstico en cada fase."
    )

    # Programa 1
    add_centered_heading(doc, "Programa 1 (Correcto): Misión Básica de Reparto Punto a Punto", space_before=6, space_after=4)
    code_1 = (
        "distancia = 15\n"
        "paquetes = 2\n\n"
        "despegar\n"
        "volar distancia\n"
        "entregar paquetes\n"
        "volar distancia\n"
        "aterrizar"
    )
    add_code_box(doc, code_1)
    add_body_paragraph(
        doc,
        "Explicación operativa: El programa define las variables distancia con 15 metros y paquetes con 2 unidades. "
        "El dron despega, vuela 15 metros en línea recta hacia el punto objetivo, descarga los 2 paquetes solicitados "
        "mediante el servomecanismo, emprende el retorno recorriendo 15 metros y ejecuta un aterrizaje controlado. "
        "Supera exitosamente las cuatro fases del compilador."
    )

    # Programa 2
    add_centered_heading(doc, "Programa 2 (Correcto): Condición de Autonomía de Batería con Operaciones", space_before=8, space_after=4)
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
    add_code_box(doc, code_2)
    add_body_paragraph(
        doc,
        "Explicación operativa: Aplica multiplicación (10 * 3 = 30) y resta (75 - 30 = 45) en expresiones. "
        "Evalúa la estructura condicional si ... entonces ... sino ... fin. Como la condición 45 > 30 es verdadera, "
        "el dron toma la ruta principal de 20 metros y realiza la entrega de 1 paquete antes de aterrizar."
    )

    # Programa 3
    add_centered_heading(doc, "Programa 3 (Correcto): Reparto Cíclico en Múltiples Zonas con Repetición", space_before=8, space_after=4)
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
    add_code_box(doc, code_3)
    add_body_paragraph(
        doc,
        "Explicación operativa: Combina suma (5 + 3 = 8 m) con el ciclo repetitivo repetir ... veces ... fin. "
        "El dron repite 4 veces la secuencia de vuelo de 8 metros y descarga de 1 paquete. Al completarse "
        "las 4 zonas de entrega, el control continúa hacia la instrucción aterrizar."
    )

    # =========================================================================
    # PÁGINA 5: DIAGNÓSTICO DE ERRORES Y TABLA COMPARATIVA (PUNTO 4)
    # =========================================================================
    doc.add_page_break()

    add_centered_heading(doc, "4. DIAGNÓSTICO Y JUSTIFICACIÓN TÉCNICA DE ERRORES")

    add_body_paragraph(
        doc,
        "Para responder con exactitud a los requerimientos de la cátedra, a continuación se documentan los tres programas "
        "con fallas intencionales. Cada caso detalla el código erróneo, la fase que captura el fallo, el mensaje exacto "
        "arrojado por el compilador con su línea y columna, y la justificación teórica de por qué se detecta en dicha etapa."
    )

    # Error Léxico
    add_centered_heading(doc, "Programa 4 (Error Intencional): Presencia de Error Léxico", space_before=6, space_after=4)
    code_4 = (
        "bateria = 100 @\n"
        "despegar\n"
        "volar 10\n"
        "aterrizar"
    )
    add_code_box(doc, code_4)

    t_err1 = doc.add_table(rows=3, cols=2)
    t_err1.alignment = WD_TABLE_ALIGNMENT.CENTER
    diag1 = [
        ("FASE DE DETECCIÓN", "Fase 1: Análisis Léxico (Lexer)"),
        ("MENSAJE DEL COMPILADOR", "[Error Léxico] Línea 1, Columna 15: Símbolo desconocido o no permitido en DronLang: '@'"),
        ("JUSTIFICACIÓN TÉCNICA", 
         "El símbolo '@' no forma parte del alfabeto formal de DronLang. Al leer el texto carácter por carácter, "
         "la función tokenizar() en src/lexer.js no encuentra correspondencia con ningún patrón válido (palabras clave, "
         "números, operadores o identificadores) y dispara inmediatamente una excepción, deteniendo el compilador "
         "antes de que el código pueda llegar a la fase sintáctica.")
    ]
    for r_i, (k, v) in enumerate(diag1):
        row = t_err1.rows[r_i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = Inches(2.2), Inches(4.3)
        set_cell_padding(c0, 60, 60, 80, 80)
        set_cell_padding(c1, 60, 60, 80, 80)
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
    apply_full_grid_borders(t_err1, border_sz="6")

    # Error Sintáctico
    add_centered_heading(doc, "Programa 5 (Error Intencional): Presencia de Error Sintáctico", space_before=10, space_after=4)
    code_5 = (
        "entregas = 3\n"
        "despegar\n"
        "repetir entregas veces\n"
        "    volar 12\n"
        "    entregar 1\n"
        "aterrizar"
    )
    add_code_box(doc, code_5)

    t_err2 = doc.add_table(rows=3, cols=2)
    t_err2.alignment = WD_TABLE_ALIGNMENT.CENTER
    diag2 = [
        ("FASE DE DETECCIÓN", "Fase 2: Análisis Sintáctico (Parser)"),
        ("MENSAJE DEL COMPILADOR", "[Error Sintáctico] Línea 6, Columna 1: Se esperaba 'fin' para cerrar la estructura 'repetir', pero se encontró 'aterrizar'"),
        ("JUSTIFICACIÓN TÉCNICA", 
         "Todos los caracteres y palabras individuales pertenecen al lenguaje, por lo que la fase léxica genera tokens sin objeción. "
         "No obstante, en la fase sintáctica, el analizador recursivo descendente exige que toda estructura 'repetir' concluya "
         "con la palabra de cierre 'fin'. Al encontrar 'aterrizar' sin haberse cerrado el ciclo, el Parser identifica una "
         "oración incompleta y mal estructurada, abortando la construcción del árbol AST.")
    ]
    for r_i, (k, v) in enumerate(diag2):
        row = t_err2.rows[r_i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = Inches(2.2), Inches(4.3)
        set_cell_padding(c0, 60, 60, 80, 80)
        set_cell_padding(c1, 60, 60, 80, 80)
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
    apply_full_grid_borders(t_err2, border_sz="6")

    # Error Semántico
    add_centered_heading(doc, "Programa 6 (Error Intencional): Presencia de Error Semántico", space_before=10, space_after=4)
    code_6 = (
        "paquetes = 2\n"
        "despegar\n"
        "volar distanciaDestino\n"
        "entregar paquetes\n"
        "aterrizar"
    )
    add_code_box(doc, code_6)

    t_err3 = doc.add_table(rows=3, cols=2)
    t_err3.alignment = WD_TABLE_ALIGNMENT.CENTER
    diag3 = [
        ("FASE DE DETECCIÓN", "Fase 3: Análisis Semántico (Semantic Analyzer)"),
        ("MENSAJE DEL COMPILADOR", "[Error Semántico] Línea 3: La variable 'distanciaDestino' no ha sido declarada o inicializada"),
        ("JUSTIFICACIÓN TÉCNICA", 
         "El código es léxicamente impecable ('distanciaDestino' cumple el patrón de identificadores) y sintácticamente válido "
         "(la regla volar <id> es correcta). Sin embargo, al intervenir la fase semántica, el verificador consulta la "
         "Tabla de Símbolos y constata que 'distanciaDestino' jamás fue asignada previamente. La instrucción carece de significado "
         "lógico y operativo: el dron no dispone de una cantidad válida de metros para efectuar su trayectoria física.")
    ]
    for r_i, (k, v) in enumerate(diag3):
        row = t_err3.rows[r_i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = Inches(2.2), Inches(4.3)
        set_cell_padding(c0, 60, 60, 80, 80)
        set_cell_padding(c1, 60, 60, 80, 80)
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
    apply_full_grid_borders(t_err3, border_sz="6")

    # Tabla Comparativa de Cierre
    add_centered_heading(doc, "Tabla 3. Cuadro Comparativo de los Tres Niveles de Error en Compilación", space_before=14, space_after=6)

    t3 = doc.add_table(rows=4, cols=4)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths_t3 = [Inches(1.2), Inches(1.5), Inches(1.8), Inches(2.0)]
    headers_t3 = ["NIVEL DE ERROR", "FASE RESPONSABLE", "ORIGEN DEL FALLO", "IMPACTO EN EL COMPILADOR"]

    for idx, (h, w) in enumerate(zip(headers_t3, widths_t3)):
        cell = t3.cell(0, idx)
        cell.width = w
        set_cell_padding(cell, 80, 80, 80, 80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = BLACK

    data_t3 = [
        ("Léxico", "Fase 1 (Lexer)", "Carácter '@' no contemplado en el alfabeto.", "Aborto inmediato en lectura. No genera la tira de tokens ni pasa al Parser."),
        ("Sintáctico", "Fase 2 (Parser)", "Omisión del token de cierre 'fin' en 'repetir'.", "Tokens reconocidos, pero la estructura infringe la gramática. No genera AST."),
        ("Semántico", "Fase 3 (Semántico)", "Uso de variable 'distanciaDestino' no declarada.", "Gramática correcta, pero carece de sentido lógico y operativo para el vuelo.")
    ]

    for r_idx, row_vals in enumerate(data_t3, start=1):
        row = t3.rows[r_idx]
        for c_idx, (val, w) in enumerate(zip(row_vals, widths_t3)):
            cell = row.cells[c_idx]
            cell.width = w
            set_cell_padding(cell, 60, 60, 80, 80)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            run.font.bold = (c_idx == 0)
            run.font.color.rgb = BLACK

    apply_full_grid_borders(t3, border_sz="6")

    # =========================================================================
    # PÁGINA 6: CONCLUSIONES (EN SU PROPIA PÁGINA)
    # =========================================================================
    doc.add_page_break()

    add_centered_heading(doc, "CONCLUSIONES", space_before=18, space_after=18)

    add_body_paragraph(
        doc,
        "Primera. El diseño conceptual y la especificación sintáctica de DronLang resuelven exitosamente "
        "el encargo pedagógico asignado por la cátedra para la temática Dron de reparto. La integración de primitivas "
        "aéreas especializadas (despegar, volar, entregar, aterrizar) junto a facilidades de cálculo de batería "
        "y control de flujo proporciona un entorno de programación expresivo, seguro y libre de ambigüedades.",
        space_after=14
    )

    add_body_paragraph(
        doc,
        "Segunda. La arquitectura en cuatro fases garantiza una separación estricta de responsabilidades en el compilador. "
        "El análisis léxico asegura la pureza del alfabeto; el análisis sintáctico valida la concordancia gramatical "
        "según la gramática libre de contexto; el análisis semántico previene fallos operativos en el mundo físico "
        "mediante la Tabla de Símbolos; y la generación de código produce secuencias de bajo nivel directamente ejecutables.",
        space_after=14
    )

    add_body_paragraph(
        doc,
        "Tercera. El experimento sistemático con los tres programas erróneos confirma la inviolabilidad de los límites "
        "entre etapas: un error léxico jamás alcanza al analizador sintáctico; una falta de delimitador sintáctico detiene "
        "la formación del árbol AST; y un error semántico salvaguarda al dron impidiendo que intente navegar con distancias "
        "indeterminadas o valores negativos.",
        space_after=14
    )

    add_body_paragraph(
        doc,
        "Cuarta. El trabajo desarrollado por el equipo satisface con máxima rigurosidad todos los entregables de la Actividad 3: "
        "catálogo de instrucciones detallado, arquitectura de fases con definición textual de cada una, banco de seis programas "
        "probados y fundamentación teórica de los fallos, consolidando una base robusta para la implementación técnica de las fases subsiguientes.",
        space_after=14
    )

    # =========================================================================
    # PÁGINA 7: REFERENCIAS (EN SU PROPIA PÁGINA)
    # =========================================================================
    doc.add_page_break()

    add_centered_heading(doc, "REFERENCIAS", space_before=18, space_after=18)

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

    # Guardar en el archivo ÚNICO
    output_path = os.path.join(os.getcwd(), "Entrega_Actividad_3_DronLang.docx")
    doc.save(output_path)
    print(f"Documento UNICO generado exitosamente en: {output_path}")

if __name__ == "__main__":
    build_single_official_document()
