// ====================================================================
// FASE 4: GENERACIÓN DE CÓDIGO - LENGUAJE DRONLANG
// Asignatura: Compiladores - Tema: Dron de Reparto
// ====================================================================

class GeneradorCodigo {
  constructor() {
    this.indentacion = 0;
  }

  sangria() {
    return '  '.repeat(this.indentacion);
  }

  generar(ast) {
    if (!ast || ast.tipo !== 'Programa') {
      throw new Error('[Error Generación] Árbol AST no válido.');
    }

    const lineas = [];
    lineas.push('// ======================================================');
    lineas.push('// CODIGO OBJETO GENERADO - CONTROL DEL DRON DE REPARTO');
    lineas.push('// ======================================================');
    lineas.push('(() => {');
    this.indentacion++;
    lineas.push(`${this.sangria()}console.log("[SISTEMA] Iniciando mision de vuelo del Dron...");`);

    for (const nodo of ast.cuerpo) {
      lineas.push(this.generarNodo(nodo));
    }

    lineas.push(`${this.sangria()}console.log("[SISTEMA] Mision del Dron finalizada con exito.");`);
    this.indentacion--;
    lineas.push('})();');

    return lineas.join('\n');
  }

  generarNodo(nodo) {
    switch (nodo.tipo) {
      case 'Asignacion': {
        const valor = this.generarExpresion(nodo.valor);
        return `${this.sangria()}let ${nodo.nombre} = ${valor};`;
      }

      case 'Despegar': {
        return `${this.sangria()}console.log("[DRON: COMANDO] -> DESPEGAR: Motores activos a 500 RPM");`;
      }

      case 'Aterrizar': {
        return `${this.sangria()}console.log("[DRON: COMANDO] -> ATERRIZAR: Descenso seguro completado");`;
      }

      case 'Volar': {
        const dist = this.generarExpresion(nodo.distancia);
        return `${this.sangria()}console.log("[DRON: COMANDO] -> VOLAR: Desplazamiento de " + (${dist}) + " metros");`;
      }

      case 'Entregar': {
        const cant = this.generarExpresion(nodo.cantidad);
        return `${this.sangria()}console.log("[DRON: COMANDO] -> ENTREGAR: Descargando " + (${cant}) + " paquete(s) en coordenadas");`;
      }

      case 'Repetir': {
        const limite = this.generarExpresion(nodo.cantidad);
        let codigo = `${this.sangria()}for (let __i = 0; __i < (${limite}); __i++) {\n`;
        this.indentacion++;
        for (const subNodo of nodo.cuerpo) {
          codigo += this.generarNodo(subNodo) + '\n';
        }
        this.indentacion--;
        codigo += `${this.sangria()}}`;
        return codigo;
      }

      case 'CondicionalSi': {
        const cond = this.generarExpresion(nodo.condicion);
        let codigo = `${this.sangria()}if (${cond}) {\n`;
        this.indentacion++;
        for (const subNodo of nodo.cuerpoSi) {
          codigo += this.generarNodo(subNodo) + '\n';
        }
        this.indentacion--;
        codigo += `${this.sangria()}}`;

        if (nodo.cuerpoSino && nodo.cuerpoSino.length > 0) {
          codigo += ` else {\n`;
          this.indentacion++;
          for (const subNodo of nodo.cuerpoSino) {
            codigo += this.generarNodo(subNodo) + '\n';
          }
          this.indentacion--;
          codigo += `${this.sangria()}}`;
        }
        return codigo;
      }

      default:
        throw new Error(`[Error Generación] Tipo de nodo desconocido: ${nodo.tipo}`);
    }
  }

  generarExpresion(expr) {
    if (!expr) return '0';

    switch (expr.tipo) {
      case 'Numero':
        return String(expr.valor);

      case 'Identificador':
        return expr.nombre;

      case 'ExpresionUnaria':
        return `(${expr.operador}${this.generarExpresion(expr.argumento)})`;

      case 'ExpresionBinaria': {
        const izq = this.generarExpresion(expr.izquierda);
        const der = this.generarExpresion(expr.derecha);
        return `(${izq} ${expr.operador} ${der})`;
      }

      default:
        return '0';
    }
  }

  // Genera lista de comandos de bajo nivel (como exige la Actividad 7)
  generarComandosPlanos(ast) {
    const comandos = [];
    const recorrer = (nodos) => {
      for (const nodo of nodos) {
        if (nodo.tipo === 'Despegar') comandos.push('DESPEGAR');
        else if (nodo.tipo === 'Aterrizar') comandos.push('ATERRIZAR');
        else if (nodo.tipo === 'Volar') comandos.push(`VOLAR ${this.generarExpresion(nodo.distancia)}`);
        else if (nodo.tipo === 'Entregar') comandos.push(`ENTREGAR ${this.generarExpresion(nodo.cantidad)}`);
        else if (nodo.tipo === 'Repetir') recorrer(nodo.cuerpo);
        else if (nodo.tipo === 'CondicionalSi') {
          recorrer(nodo.cuerpoSi);
          if (nodo.cuerpoSino) recorrer(nodo.cuerpoSino);
        }
      }
    };
    recorrer(ast.cuerpo);
    return comandos;
  }
}

function ejecutarCodigoGenerado(codigoJs) {
  const salidas = [];
  const fakeConsole = {
    log: (...args) => {
      salidas.push(args.join(' '));
    },
    error: (...args) => {
      salidas.push('[ERROR] ' + args.join(' '));
    }
  };

  try {
    const fn = new Function('console', codigoJs);
    fn(fakeConsole);
    return {
      exito: true,
      salidas: salidas
    };
  } catch (error) {
    return {
      exito: false,
      error: error.message,
      salidas: salidas
    };
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { GeneradorCodigo, ejecutarCodigoGenerado };
}
