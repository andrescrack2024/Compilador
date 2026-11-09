# Actividad 3: Examen de Conceptualización Básica
**Asignatura:** Compiladores  
**Docente:** Ing. Hasler Gamboa Berrío  
**Temática:** Dron de Reparto  
**Lenguaje:** DronLang  
**Integrantes:**
- Jainer Andrés Morales Cuesta
- Sharly Andrés Mosquera Rodríguez

---

## Punto 1: Qué hace cada instrucción y cómo se escribe

### Instrucciones Propias del Tema (Dron de Reparto)
1. **`despegar`**
   * **Cómo se escribe:** `despegar`
   * **Qué hace:** Enciende los motores del dron y realiza el ascenso vertical a la altura de crucero.
2. **`volar`**
   * **Cómo se escribe:** `volar <distancia>` (ejemplo: `volar 15` o `volar distancia`)
   * **Qué hace:** Desplaza el dron horizontalmente hacia el destino la cantidad de metros indicada.
3. **`entregar`**
   * **Cómo se escribe:** `entregar <paquetes>` (ejemplo: `entregar 2` o `entregar paquetes`)
   * **Qué hace:** Descarga en el punto de destino la cantidad de paquetes indicada.
4. **`aterrizar`**
   * **Cómo se escribe:** `aterrizar`
   * **Qué hace:** Realiza el descenso vertical en zona segura y apaga los motores.

### Variables y Estructuras de Control
5. **Variables con números:**
   * **Cómo se escribe:** `<nombre> = <numero>` o `<nombre> = <expresion>` (ejemplo: `bateria = 100`, `paquetes = 2`).
   * **Qué hace:** Guarda un valor numérico en memoria para usarlo en otras instrucciones.
6. **Repetición:**
   * **Cómo se escribe:**
     ```text
     repetir <veces> veces
         <instrucciones>
     fin
     ```
   * **Qué hace:** Repite un bloque de instrucciones la cantidad de veces especificada por un número o variable.
7. **Condición:**
   * **Cómo se escribe:**
     ```text
     si <condicion> entonces
         <instrucciones>
     sino
         <instrucciones>
     fin
     ```
   * **Qué hace:** Evalúa una comparación; si se cumple ejecuta el primer bloque, y si no se cumple ejecuta el bloque `sino`.

### Operaciones Aceptadas
* **Operaciones aritméticas:** Suma (`+`), Resta (`-`), Multiplicación (`*`).
* **Operaciones de comparación:** Mayor que (`>`), Menor que (`<`), Igualdad (`==`).

---

## Punto 2: Dibujo de las fases del compilador y una frase por fase

```text
       [ Código fuente en DronLang ]
                     │
                     ▼
       ┌───────────────────────────┐
       │     1. Análisis Léxico    │
       └───────────────────────────┘
                     │ Tokens
                     ▼
       ┌───────────────────────────┐
       │   2. Análisis Sintáctico  │
       └───────────────────────────┘
                     │ Árbol AST
                     ▼
       ┌───────────────────────────┐
       │   3. Análisis Semántico   │
       └───────────────────────────┘
                     │ Árbol Validado
                     ▼
       ┌───────────────────────────┐
       │ 4. Generación de Código   │
       └───────────────────────────┘
                     │
                     ▼
       [ Comandos de vuelo del Dron ]
```

### Una frase por fase en DronLang:
* **Fase 1 (Análisis Léxico):** Lee el texto del programa y lo convierte en fichas llamadas tokens (palabras como `despegar`, números y operadores), detectando si hay caracteres no permitidos.
* **Fase 2 (Análisis Sintáctico):** Revisa que las instrucciones sigan la estructura correcta del lenguaje (como que a `repetir` no le falte `veces` ni `fin`) y crea el Árbol de Sintaxis Abstracta (AST).
* **Fase 3 (Análisis Semántico):** Comprueba la lógica del programa, revisando que no se usen variables sin definir y que las distancias o repeticiones no sean números negativos.
* **Fase 4 (Generación de Código):** Traduce el árbol de instrucciones en una lista de comandos de vuelo (`DESPEGAR`, `VOLAR`, `ENTREGAR`, `ATERRIZAR`) para que el dron ejecute la misión.

---

## Punto 3 y 4: Los seis programas de ejemplo y explicación de errores

### Programa 1 (Correcto - Misión Básica de Entrega)
```text
distancia = 15
paquetes = 2

despegar
volar distancia
entregar paquetes
volar distancia
aterrizar
```
* **Qué hace:** El dron despega, vuela 15 metros, entrega 2 paquetes, regresa 15 metros y aterriza sin errores.

---

### Programa 2 (Correcto - Condición de Batería con Operaciones)
```text
bateria = 75
consumoRuta = 10 * 3
bateriaRestante = bateria - consumoRuta

despegar
si bateriaRestante > 30 entonces
    volar 20
    entregar 1
sino
    volar 5
    entregar 1
fin
aterrizar
```
* **Qué hace:** Calcula el gasto de energía con multiplicación y resta, evalúa la condición si la batería alcanza para la ruta larga, realiza la entrega y aterriza.

---

### Programa 3 (Correcto - Reparto Múltiple con Repetición)
```text
zonas = 4
distanciaSector = 5 + 3

despegar
repetir zonas veces
    volar distanciaSector
    entregar 1
fin
aterrizar
```
* **Qué hace:** Usa la estructura `repetir ... veces ... fin` para realizar 4 vuelos y entregas consecutivas.

---

### Programa 4 (Error a propósito: Error Léxico)
```text
bateria = 100 @
despegar
volar 10
aterrizar
```
* **¿En qué fase se detecta?** En la **Fase 1 (Análisis Léxico)**.
* **¿Por qué?** Porque el símbolo `@` no existe en el alfabeto de DronLang. El analizador léxico no lo reconoce como palabra clave, número, variable ni operador, y detiene el proceso de inmediato.

---

### Programa 5 (Error a propósito: Error Sintáctico)
```text
entregas = 3
despegar
repetir entregas veces
    volar 12
    entregar 1
aterrizar
```
* **¿En qué fase se detecta?** En la **Fase 2 (Análisis Sintáctico)**.
* **¿Por qué?** Porque a la instrucción `repetir` le falta su palabra de cierre obligatoria `fin`. Aunque todas las palabras existen individualmente, la estructura de la oración está mal armada según las reglas gramaticales del lenguaje.

---

### Programa 6 (Error a propósito: Error Semántico)
```text
paquetes = 2
despegar
volar distanciaDestino
entregar paquetes
aterrizar
```
* **¿En qué fase se detecta?** En la **Fase 3 (Análisis Semántico)**.
* **¿Por qué?** Porque la variable `distanciaDestino` nunca fue definida antes de usarla. Todas las palabras son válidas (pasa el léxico) y la estructura de la frase está bien escrita (pasa el sintáctico), pero el programa carece de sentido lógico porque el compilador no sabe cuántos metros debe volar el dron.
