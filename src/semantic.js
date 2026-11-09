// ====================================================================
// FASE 3: ANALIZADOR SEMÁNTICO - LENGUAJE DRONLANG
// Asignatura: Compiladores - Tema: Dron de Reparto
// ====================================================================

class AnalizadorSemantico {
  constructor() {
    this.tablaSimbolos = new Map(); // nombre -> { linea, valorConstante }
    this.errores = [];
    this.advertencias = [];
    this.dronEnElAire = false;
  }

  analizar(ast) {
    this.tablaSimbolos.clear();
    this.errores = [];
    this.advertencias = [];
    this.dronEnElAire = false;

    if (!ast || ast.tipo !== 'Programa') {
      throw new Error('[Error Semántico] Árbol AST no válido o vacío.');
    }

    for (const nodo of ast.cuerpo) {
      this.visitarInstruccion(nodo);
    }

    return {
      tablaSimbolos: Array.from(this.tablaSimbolos.entries()).map(([nombre, datos]) => ({
        nombre,
        lineaDeclaracion: datos.linea
      })),
      errores: this.errores,
      advertencias: this.advertencias,
      esValido: this.errores.length === 0
    };
  }

  visitarInstruccion(nodo) {
    if (!nodo) return;

    switch (nodo.tipo) {
      case 'Asignacion': {
        this.evaluarExpresion(nodo.valor);
        this.tablaSimbolos.set(nodo.nombre, {
          linea: nodo.linea
        });
        break;
      }

      case 'Despegar': {
        this.dronEnElAire = true;
        break;
      }

      case 'Aterrizar': {
        this.dronEnElAire = false;
        break;
      }

      case 'Volar': {
        // Regla Semántica 1: Distancia no puede ser negativa
        if (nodo.distancia.tipo === 'Numero' && nodo.distancia.valor < 0) {
          this.errores.push(
            `[Error Semántico] Línea ${nodo.linea}: Regla de vuelo violada. La distancia de 'volar' no puede ser negativa (${nodo.distancia.valor}).`
          );
        }
        this.evaluarExpresion(nodo.distancia);
        break;
      }

      case 'Entregar': {
        // Regla Semántica 2: Cantidad de paquetes no puede ser negativa
        if (nodo.cantidad.tipo === 'Numero' && nodo.cantidad.valor <= 0) {
          this.errores.push(
            `[Error Semántico] Línea ${nodo.linea}: Regla de entrega violada. La cantidad a entregar debe ser mayor a cero (${nodo.cantidad.valor}).`
          );
        }
        this.evaluarExpresion(nodo.cantidad);
        break;
      }

      case 'Repetir': {
        // Regla Semántica 3: La repetición no acepta números negativos
        if (nodo.cantidad.tipo === 'Numero' && nodo.cantidad.valor < 0) {
          this.errores.push(
            `[Error Semántico] Línea ${nodo.linea}: Regla de repetición violada. 'repetir' no acepta valores negativos (${nodo.cantidad.valor}).`
          );
        }
        this.evaluarExpresion(nodo.cantidad);
        for (const subNodo of nodo.cuerpo) {
          this.visitarInstruccion(subNodo);
        }
        break;
      }

      case 'CondicionalSi': {
        this.evaluarExpresion(nodo.condicion);
        for (const subNodo of nodo.cuerpoSi) {
          this.visitarInstruccion(subNodo);
        }
        if (nodo.cuerpoSino) {
          for (const subNodo of nodo.cuerpoSino) {
            this.visitarInstruccion(subNodo);
          }
        }
        break;
      }

      default:
        break;
    }
  }

  evaluarExpresion(expr) {
    if (!expr) return;

    if (expr.tipo === 'Identificador') {
      // Regla Semántica Fundamental: Variable debe existir en la tabla de símbolos
      if (!this.tablaSimbolos.has(expr.nombre)) {
        this.errores.push(
          `[Error Semántico] Línea ${expr.linea}: La variable '${expr.nombre}' no ha sido definida previamente.`
        );
      }
      return;
    }

    if (expr.tipo === 'ExpresionUnaria') {
      this.evaluarExpresion(expr.argumento);
      return;
    }

    if (expr.tipo === 'ExpresionBinaria') {
      this.evaluarExpresion(expr.izquierda);
      this.evaluarExpresion(expr.derecha);
      return;
    }
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { AnalizadorSemantico };
}
