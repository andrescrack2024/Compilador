// ====================================================================
// FASE 1: ANALIZADOR LÉXICO (LEXER) - LENGUAJE DRONLANG
// Asignatura: Compiladores - Tema: Dron de Reparto
// ====================================================================

const TIPOS = {
  PALABRA_CLAVE: 'PALABRA_CLAVE',
  IDENTIFICADOR: 'IDENTIFICADOR',
  NUMERO: 'NUMERO',
  OPERADOR: 'OPERADOR',
  DELIMITADOR: 'DELIMITADOR'
};

// Palabras reservadas del lenguaje del Dron
const PALABRAS_CLAVE = new Set([
  'despegar',
  'volar',
  'entregar',
  'aterrizar',
  'repetir',
  'veces',
  'si',
  'entonces',
  'sino',
  'fin'
]);

class Lexer {
  constructor(codigo) {
    this.codigo = codigo;
    this.pos = 0;
    this.linea = 1;
    this.columna = 1;
    this.tokens = [];
  }

  error(mensaje) {
    throw new Error(`[Error Léxico] Línea ${this.linea}, Columna ${this.columna}: ${mensaje}`);
  }

  caracterActual() {
    return this.codigo[this.pos] || null;
  }

  mirarSiguiente() {
    return this.codigo[this.pos + 1] || null;
  }

  avanzar() {
    const c = this.caracterActual();
    if (c === '\n') {
      this.linea++;
      this.columna = 1;
    } else {
      this.columna++;
    }
    this.pos++;
    return this.caracterActual();
  }

  tokenizar() {
    while (this.pos < this.codigo.length) {
      let c = this.caracterActual();

      // 1. Ignorar espacios en blanco y tabulaciones
      if (c === ' ' || c === '\t' || c === '\r' || c === '\n') {
        this.avanzar();
        continue;
      }

      // 2. Ignorar comentarios // o #
      if ((c === '/' && this.mirarSiguiente() === '/') || c === '#') {
        while (this.caracterActual() && this.caracterActual() !== '\n') {
          this.avanzar();
        }
        continue;
      }

      const inicioCol = this.columna;
      const inicioLin = this.linea;

      // 3. Números (enteros o decimales)
      if (/[0-9]/.test(c)) {
        let numStr = '';
        while (this.caracterActual() && /[0-9]/.test(this.caracterActual())) {
          numStr += this.caracterActual();
          this.avanzar();
        }
        if (this.caracterActual() === '.' && /[0-9]/.test(this.mirarSiguiente())) {
          numStr += '.';
          this.avanzar();
          while (this.caracterActual() && /[0-9]/.test(this.caracterActual())) {
            numStr += this.caracterActual();
            this.avanzar();
          }
        }
        this.tokens.push({
          tipo: TIPOS.NUMERO,
          valor: parseFloat(numStr),
          linea: inicioLin,
          columna: inicioCol
        });
        continue;
      }

      // 4. Identificadores y Palabras Reservadas
      if (/[a-zA-ZáéíóúÁÉÍÓÚñÑ_]/.test(c)) {
        let texto = '';
        while (this.caracterActual() && /[a-zA-ZáéíóúÁÉÍÓÚñÑ0-9_]/.test(this.caracterActual())) {
          texto += this.caracterActual();
          this.avanzar();
        }

        const tipo = PALABRAS_CLAVE.has(texto)
          ? TIPOS.PALABRA_CLAVE
          : TIPOS.IDENTIFICADOR;

        this.tokens.push({
          tipo: tipo,
          valor: texto,
          linea: inicioLin,
          columna: inicioCol
        });
        continue;
      }

      // 5. Operador de igualdad '=='
      if (c === '=' && this.mirarSiguiente() === '=') {
        this.tokens.push({
          tipo: TIPOS.OPERADOR,
          valor: '==',
          linea: inicioLin,
          columna: inicioCol
        });
        this.avanzar();
        this.avanzar();
        continue;
      }

      // 6. Operadores simples: =, +, -, *, >, <
      if (['=', '+', '-', '*', '>', '<'].includes(c)) {
        this.tokens.push({
          tipo: TIPOS.OPERADOR,
          valor: c,
          linea: inicioLin,
          columna: inicioCol
        });
        this.avanzar();
        continue;
      }

      // 7. Paréntesis para agrupamiento en expresiones
      if (c === '(' || c === ')') {
        this.tokens.push({
          tipo: TIPOS.DELIMITADOR,
          valor: c,
          linea: inicioLin,
          columna: inicioCol
        });
        this.avanzar();
        continue;
      }

      // 8. Símbolo desconocido -> Error Léxico
      this.error(`Símbolo desconocido o no permitido en DronLang: '${c}'`);
    }

    return this.tokens;
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { Lexer, TIPOS, PALABRAS_CLAVE };
}
