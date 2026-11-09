# Compilador DronLang - Dron de Reparto
> Proyecto de Implementación de Compilador y Lenguaje de Dominio Específico para la asignatura de **Compiladores**.  
> **Docente:** Ing. Hasler Gamboa Berrío  
> **Temática:** Dron de Reparto  
> **Equipo de Trabajo:** Jainer Andrés Morales Cuesta, Sharly Andrés Mosquera Rodríguez

---

## 1. Identificación del Proyecto
* **Nombre del Lenguaje:** **DronLang** (archivos `.dron`)
* **Propósito:** Control autónomo de secuencias de vuelo, entregas y autonomía de un dron de reparto.
* **Paradigma:** Imperativo, estructurado, en español, con tipado numérico e inferencia.
* **Motor:** JavaScript (Node.js) con arquitectura modular de 4 fases.

---

## 2. Estructura de Archivos del Proyecto
```text
Compilador actividad/
├── ENTREGA_ACTIVIDAD_3.md    (Documento oficial de entrega para la Actividad 3)
├── bitacora.md               (Bitácora de reuniones y decisiones del equipo)
├── interfaz_visual.html      (Entorno gráfico de demostración y pruebas)
├── index.js                  (Compilador por línea de comandos para Node.js)
├── package.json              (Configuración del proyecto)
├── README.md                 (Documentación técnica general)
├── src/
│   ├── lexer.js              (Fase 1: Analizador Léxico)
│   ├── parser.js             (Fase 2: Analizador Sintáctico y AST)
│   ├── semantic.js           (Fase 3: Analizador Semántico y Tabla de Símbolos)
│   └── codegen.js            (Fase 4: Generador de Comandos y Código Objeto)
└── ejemplos/
    ├── 01_mision_simple.dron        (Correcto 1: Despegue, vuelo, entrega, aterrizaje)
    ├── 02_condicional_bateria.dron  (Correcto 2: Matemáticas y condicional si/sino)
    ├── 03_reparto_multiple.dron     (Correcto 3: Bucle repetir...veces...fin)
    ├── 04_error_lexico.dron         (Error 1: Carácter no reconocido '@')
    ├── 05_error_sintactico.dron     (Error 2: Estructura repetir sin 'fin')
    └── 06_error_semantico.dron      (Error 3: Variable no definida)
```

---

## 3. Instrucciones Propias de DronLang

| Instrucción | Sintaxis | Descripción |
|---|---|---|
| **Despegar** | `despegar` | Enciende motores y asciende verticalmente. |
| **Volar** | `volar <distancia>` | Avanza los metros indicados (admite números o variables). |
| **Entregar** | `entregar <paquetes>` | Descarga los paquetes en las coordenadas actuales. |
| **Aterrizar** | `aterrizar` | Desciende y apaga motores de forma segura. |
| **Variables** | `nombre = expresion` | Asignación numérica en la tabla de símbolos. |
| **Repetición** | `repetir n veces ... fin` | Bucle de iteración finita de instrucciones. |
| **Condición** | `si cond entonces ... sino ... fin` | Bifurcación lógica de control de misión. |
| **Operaciones** | `+`, `-`, `*`, `>`, `<`, `==` | 6 operaciones matemáticas y lógicas exigidas. |

---

## 4. Ejecución por Consola de Comandos (CLI)

```bash
# Compilar y simular el Programa 1 (Misión básica):
node index.js ejemplos/01_mision_simple.dron

# Compilar viendo la tabla de tokens, AST y tabla de símbolos:
node index.js ejemplos/02_condicional_bateria.dron --fases

# Ejecutar el bucle de reparto múltiple:
node index.js ejemplos/03_reparto_multiple.dron

# Probar la detección del Error Léxico (Fase 1):
node index.js ejemplos/04_error_lexico.dron

# Probar la detección del Error Sintáctico (Fase 2):
node index.js ejemplos/05_error_sintactico.dron

# Probar la detección del Error Semántico (Fase 3):
node index.js ejemplos/06_error_semantico.dron
```

---

## 5. Demostración Gráfica (interfaz_visual.html)
Abra con doble clic el archivo `interfaz_visual.html` en el navegador para la sustentación:
- Editor con números de línea y atajo `Ctrl + Enter`.
- Selector superior con los 6 programas precargados (3 correctos y 3 con error).
- Pestañas con los resultados de las 4 fases en tiempo real.
- Panel de comandos de bajo nivel (`DESPEGAR`, `VOLAR`, `ENTREGAR`, `ATERRIZAR`) y simulación de vuelo.
