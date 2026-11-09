// ====================================================================
// FASE 2: ANALIZADOR SINTÁCTICO (PARSER) - LENGUAJE DRONLANG
// Asignatura: Compiladores - Tema: Dron de Reparto
// ====================================================================

const { TIPOS } = typeof module !== 'undefined' && module.exports
  ? require('./lexer')
  : window;

class Parser {
  constructor(tokens) {
    this.tokens = tokens;
    this.pos = 0;
  }

  tokenActual() {
    return this.tokens[this.pos] || null;
  }

  tokenAnterior() {
    return this.tokens[this.pos - 1] || null;
  }

  estaAlFinal() {
    return this.pos >= this.tokens.length;
  }

  avanzar() {
    if (!this.estaAlFinal()) {
      this.pos++;
    }
    return this.tokenAnterior();
  }

  coincide(tipo, valor = null) {
    if (this.estaAlFinal()) return false;
    const t = this.tokenActual();
    if (t.tipo !== tipo) return false;
    if (valor !== null && t.valor !== valor) return false;
    this.avanzar();
    return true;
  }

  consumir(tipo, valorEsperado, mensajeError) {
    const t = this.tokenActual();
    if (t && t.tipo === tipo && (valorEsperado === null || t.valor === valorEsperado)) {
      return this.avanzar();
    }
    const lin = t ? t.linea : (this.tokenAnterior() ? this.tokenAnterior().linea : 1);
    const col = t ? t.columna : (this.tokenAnterior() ? this.tokenAnterior().columna : 1);
    const encontrado = t ? `'${t.valor}'` : 'fin de archivo';
    throw new Error(`[Error Sintáctico] Línea ${lin}, Columna ${col}: ${mensajeError}. Se encontró ${encontrado}.`);
  }

  parsear() {
    const instrucciones = [];
    while (!this.estaAlFinal()) {
      instrucciones.push(this.instruccion());
    }
    return {
      tipo: 'Programa',
      cuerpo: instrucciones
    };
  }

  instruccion() {
    const t = this.tokenActual();
    if (!t) return null;

    // 1. Instrucciones con palabras clave
    if (t.tipo === TIPOS.PALABRA_CLAVE) {
      if (t.valor === 'despegar') {
        this.avanzar();
        return { tipo: 'Despegar', linea: t.linea };
      }
      if (t.valor === 'aterrizar') {
        this.avanzar();
        return { tipo: 'Aterrizar', linea: t.linea };
      }
      if (t.valor === 'volar') {
        this.avanzar();
        const distancia = this.expresion();
        return { tipo: 'Volar', distancia, linea: t.linea };
      }
      if (t.valor === 'entregar') {
        this.avanzar();
        const cantidad = this.expresion();
        return { tipo: 'Entregar', cantidad, linea: t.linea };
      }
      if (t.valor === 'repetir') {
        return this.instruccionRepetir();
      }
      if (t.valor === 'si') {
        return this.instruccionSi();
      }
    }

    // 2. Asignación de variables: id = expresion
    if (t.tipo === TIPOS.IDENTIFICADOR) {
      const idToken = this.avanzar();
      this.consumir(TIPOS.OPERADOR, '=', "Se esperaba '=' en la asignación");
      const valor = this.expresion();
      return {
        tipo: 'Asignacion',
        nombre: idToken.valor,
        valor: valor,
        linea: idToken.linea
      };
    }

    throw new Error(`[Error Sintáctico] Línea ${t.linea}, Columna ${t.columna}: Instrucción no reconocida comenzando con '${t.valor}'.`);
  }

  // Estructura: repetir <expresion> veces ... fin
  instruccionRepetir() {
    const tRepetir = this.consumir(TIPOS.PALABRA_CLAVE, 'repetir', "Se esperaba 'repetir'");
    const cantidad = this.expresion();
    this.consumir(TIPOS.PALABRA_CLAVE, 'veces', "Se esperaba la palabra 'veces' después del número o variable a repetir");

    const cuerpo = [];
    while (!this.estaAlFinal() && !(this.tokenActual().tipo === TIPOS.PALABRA_CLAVE && this.tokenActual().valor === 'fin')) {
      cuerpo.push(this.instruccion());
    }

    this.consumir(TIPOS.PALABRA_CLAVE, 'fin', "Se esperaba 'fin' para cerrar la estructura 'repetir'");

    return {
      tipo: 'Repetir',
      cantidad: cantidad,
      cuerpo: cuerpo,
      linea: tRepetir.linea
    };
  }

  // Estructura: si <condicion> entonces ... (sino ...)? fin
  instruccionSi() {
    const tSi = this.consumir(TIPOS.PALABRA_CLAVE, 'si', "Se esperaba 'si'");
    const condicion = this.expresion();
    this.consumir(TIPOS.PALABRA_CLAVE, 'entonces', "Se esperaba la palabra 'entonces' después de la condición de 'si'");

    const cuerpoSi = [];
    while (
      !this.estaAlFinal() &&
      !(this.tokenActual().tipo === TIPOS.PALABRA_CLAVE && (this.tokenActual().valor === 'sino' || this.tokenActual().valor === 'fin'))
    ) {
      cuerpoSi.push(this.instruccion());
    }

    let cuerpoSino = null;
    if (this.coincide(TIPOS.PALABRA_CLAVE, 'sino')) {
      cuerpoSino = [];
      while (!this.estaAlFinal() && !(this.tokenActual().tipo === TIPOS.PALABRA_CLAVE && this.tokenActual().valor === 'fin')) {
        cuerpoSino.push(this.instruccion());
      }
    }

    this.consumir(TIPOS.PALABRA_CLAVE, 'fin', "Se esperaba 'fin' para cerrar la estructura 'si'");

    return {
      tipo: 'CondicionalSi',
      condicion: condicion,
      cuerpoSi: cuerpoSi,
      cuerpoSino: cuerpoSino,
      linea: tSi.linea
    };
  }

  // ==================== EXPRESIONES ====================

  expresion() {
    return this.comparacion();
  }

  comparacion() {
    let expr = this.adicion();

    while (
      !this.estaAlFinal() &&
      this.tokenActual().tipo === TIPOS.OPERADOR &&
      ['>', '<', '=='].includes(this.tokenActual().valor)
    ) {
      const operador = this.avanzar().valor;
      const derecha = this.adicion();
      expr = {
        tipo: 'ExpresionBinaria',
        operador: operador,
        izquierda: expr,
        derecha: derecha
      };
    }

    return expr;
  }

  adicion() {
    let expr = this.multiplicacion();

    while (
      !this.estaAlFinal() &&
      this.tokenActual().tipo === TIPOS.OPERADOR &&
      ['+', '-'].includes(this.tokenActual().valor)
    ) {
      const operador = this.avanzar().valor;
      const derecha = this.multiplicacion();
      expr = {
        tipo: 'ExpresionBinaria',
        operador: operador,
        izquierda: expr,
        derecha: derecha
      };
    }

    return expr;
  }

  multiplicacion() {
    let expr = this.unario();

    while (
      !this.estaAlFinal() &&
      this.tokenActual().tipo === TIPOS.OPERADOR &&
      this.tokenActual().valor === '*'
    ) {
      const operador = this.avanzar().valor;
      const derecha = this.unario();
      expr = {
        tipo: 'ExpresionBinaria',
        operador: operador,
        izquierda: expr,
        derecha: derecha
      };
    }

    return expr;
  }

  unario() {
    if (this.tokenActual() && this.tokenActual().tipo === TIPOS.OPERADOR && this.tokenActual().valor === '-') {
      const operador = this.avanzar().valor;
      const arg = this.unario();
      return {
        tipo: 'ExpresionUnaria',
        operador: operador,
        argumento: arg
      };
    }
    return this.primario();
  }

  primario() {
    const t = this.tokenActual();
    if (!t) {
      throw new Error('[Error Sintáctico] Expresión incompleta antes de fin de archivo.');
    }

    if (t.tipo === TIPOS.NUMERO) {
      this.avanzar();
      return { tipo: 'Numero', valor: t.valor, linea: t.linea };
    }

    if (t.tipo === TIPOS.IDENTIFICADOR) {
      this.avanzar();
      return { tipo: 'Identificador', nombre: t.valor, linea: t.linea };
    }

    if (t.tipo === TIPOS.DELIMITADOR && t.valor === '(') {
      this.avanzar();
      const expr = this.expresion();
      this.consumir(TIPOS.DELIMITADOR, ')', "Se esperaba ')' para cerrar la expresión");
      return expr;
    }

    throw new Error(`[Error Sintáctico] Línea ${t.linea}, Columna ${t.columna}: Se esperaba un número o identificador pero se encontró '${t.valor}'.`);
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { Parser };
}
