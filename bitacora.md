# Bitácora de Desarrollo - Proyecto Compilador DronLang
**Asignatura:** Compiladores  
**Docente:** Ing. Hasler Gamboa Berrío  
**Temática Asignada:** Dron de Reparto  
**Integrantes:**
- Jainer Andrés Morales Cuesta
- Sharly Andrés Mosquera Rodríguez

---

## Sesión 1: Selección de Temática y Definición de Objetivos
- **Fecha:** 25 de Octubre de 2026
- **Participantes:** Jainer Andrés Morales Cuesta y Sharly Andrés Mosquera Rodríguez.
- **Decisiones tomadas:**
  1. Se seleccionó la temática **"Dron de reparto"** de entre las 8 propuestas por el docente.
  2. **Justificación:** El control de un dron se presta naturalmente para operaciones numéricas puras (metros de vuelo, cantidad de paquetes, porcentaje de batería), cumpliendo estrictamente con la indicación del profesor de usar "variables con números" y evitando la complejidad innecesaria de cadenas o bases de datos de otros temas como biblioteca o tienda.
  3. Se definieron las 3 instrucciones propias del tema: `despegar`, `volar <distancia>`, `entregar <paquetes>` y se añadió la instrucción complementaria `aterrizar`.

---

## Sesión 2: Diseño de la Gramática y Sintaxis del Lenguaje
- **Fecha:** 28 de Octubre de 2026
- **Participantes:** Jainer Andrés Morales Cuesta y Sharly Andrés Mosquera Rodríguez.
- **Decisiones tomadas:**
  1. Se acordó adoptar una sintaxis limpia y directa sin punto y coma (`;`), alineada al ejemplo mostrado por el profesor en clase (`cajas = 4`, `repetir cajas veces ... fin`).
  2. Para las estructuras de control se definieron:
     - Repetición: `repetir <expr> veces ... fin`.
     - Bifurcación: `si <cond> entonces ... sino ... fin`.
  3. Se limitaron los operadores a los 6 exigidos: aritméticos (`+`, `-`, `*`) y relacionales (`>`, `<`, `==`).
  4. Se redactó la especificación formal para la **Actividad 3**.

---

## Sesión 3: Implementación de la Fase 1 (Léxico) y Fase 2 (Sintáctico)
- **Fecha:** 2 de Noviembre de 2026
- **Participantes:** Jainer Andrés Morales Cuesta y Sharly Andrés Mosquera Rodríguez.
- **Responsables de módulo:**
  - Jainer Andrés Morales Cuesta: Construcción de expresiones regulares y reconocedor de tokens en `src/lexer.js`.
  - Sharly Andrés Mosquera Rodríguez: Construcción del analizador sintáctico descendente recursivo y generación del AST en `src/parser.js`.
- **Decisiones técnicas:**
  - El escáner léxico ignora comentarios de una línea (`#` o `//`) y espacios, pero rastrea línea y columna exacta para reportar errores amigables al programador.
  - El parser valida de forma estricta el emparejamiento de `repetir...veces...fin` y `si...entonces...fin`.

---

## Sesión 4: Fase 3 (Semántica), Fase 4 (Generación de Código) y Casos de Prueba
- **Fecha:** 6 de Noviembre de 2026
- **Participantes:** Jainer Andrés Morales Cuesta y Sharly Andrés Mosquera Rodríguez.
- **Responsables de módulo:**
  - Sharly Andrés Mosquera Rodríguez: Implementación de la Tabla de Símbolos y validaciones de ámbito en `src/semantic.js`.
  - Jainer Andrés Morales Cuesta: Implementación de `src/codegen.js` para generar los comandos de bajo nivel (`DESPEGAR`, `VOLAR`, `ENTREGAR`, `ATERRIZAR`) y el simulador de vuelo en consola.
- **Decisiones técnicas:**
  - Se programaron 3 reglas semánticas propias del dron:
    1. Prohibido usar variables no definidas.
    2. La cantidad de repeticiones no puede ser negativa.
    3. Las distancias de vuelo y paquetes no pueden ser negativos.
  - Se construyeron y validaron los 6 programas de prueba requeridos por la Actividad 3 (3 correctos y 3 con errores a propósito: léxico, sintáctico y semántico).

---

## Sesión 5: Preparación de la Defensa Oral y Ensayo de Modificación en Vivo
- **Fecha:** 8 de Noviembre de 2026
- **Participantes:** Jainer Andrés Morales Cuesta y Sharly Andrés Mosquera Rodríguez.
- **Acuerdos para la sustentación:**
  1. **Reparto de roles para la videollamada:**
     - Jainer Andrés Morales Cuesta: Presentación del lenguaje DronLang, Fase Léxica y Fase de Generación de Código.
     - Sharly Andrés Mosquera Rodríguez: Fase Sintáctica (AST), Fase Semántica (Tabla de Símbolos) y Demostración de los 3 errores.
  2. **Plan de contingencia para la modificación en vivo del profesor (Actividad 7):**
     - Si el profesor pide agregar un comando nuevo (ejemplo: `recargar` o `girar <grados>`):
       - Se añade la palabra clave en el `Set` de `src/lexer.js`.
       - Se añade el método en `src/parser.js`.
       - Se añade la emisión de comando en `src/codegen.js`.
       - Tiempo estimado de cambio: menos de 3 minutos.
