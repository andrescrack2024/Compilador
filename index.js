#!/usr/bin/env node

// ====================================================================
// COMPILADOR DRONLANG (CLI)
// Asignatura: Compiladores
// Docente: Ing. Hasler Gamboa Berrio
// Temática: Dron de Reparto
// Integrantes: Jainer Andres Morales Cuesta, Sharly Andres Mosquera Rodriguez
// ====================================================================

const fs = require('fs');
const path = require('path');
const { Lexer } = require('./src/lexer');
const { Parser } = require('./src/parser');
const { AnalizadorSemantico } = require('./src/semantic');
const { GeneradorCodigo, ejecutarCodigoGenerado } = require('./src/codegen');

function compilar(codigoFuente, opciones = { mostrarFases: false }) {
  console.log('======================================================');
  console.log('   COMPILADOR DRONLANG v1.0 - DRON DE REPARTO         ');
  console.log('======================================================\n');

  try {
    // ----------------------------------------------------
    // FASE 1: ANÁLISIS LÉXICO
    // ----------------------------------------------------
    console.log('[FASE 1: ANÁLISIS LÉXICO]');
    const lexer = new Lexer(codigoFuente);
    const tokens = lexer.tokenizar();
    console.log(` -> Exito: ${tokens.length} tokens identificados.`);
    if (opciones.mostrarFases) {
      console.log(' -> Tabla de Tokens:');
      console.table(tokens.map(t => ({ Tipo: t.tipo, Lexema: t.valor, Linea: t.linea, Columna: t.columna })));
    }

    // ----------------------------------------------------
    // FASE 2: ANÁLISIS SINTÁCTICO
    // ----------------------------------------------------
    console.log('\n[FASE 2: ANÁLISIS SINTÁCTICO]');
    const parser = new Parser(tokens);
    const ast = parser.parsear();
    console.log(` -> Exito: Arbol AST generado (${ast.cuerpo.length} instrucciones raiz).`);
    if (opciones.mostrarFases) {
      console.log(' -> Estructura del Arbol AST:');
      console.log(JSON.stringify(ast, null, 2));
    }

    // ----------------------------------------------------
    // FASE 3: ANÁLISIS SEMÁNTICO
    // ----------------------------------------------------
    console.log('\n[FASE 3: ANÁLISIS SEMÁNTICO]');
    const semantico = new AnalizadorSemantico();
    const resultadoSemantico = semantico.analizar(ast);

    if (resultadoSemantico.advertencias.length > 0) {
      console.log(' -> Advertencias detectadas:');
      resultadoSemantico.advertencias.forEach(adv => console.log('    [AVISO] ' + adv));
    }

    if (!resultadoSemantico.esValido) {
      console.error('\n[ERROR SEMÁNTICO DETECTADO]:');
      resultadoSemantico.errores.forEach(err => console.error('    ' + err));
      process.exit(1);
    }
    console.log(' -> Exito: Variables validadas y reglas operativas del Dron cumplidas.');
    if (opciones.mostrarFases) {
      console.log(' -> Tabla de Simbolos:');
      console.table(resultadoSemantico.tablaSimbolos);
    }

    // ----------------------------------------------------
    // FASE 4: GENERACIÓN DE CÓDIGO
    // ----------------------------------------------------
    console.log('\n[FASE 4: GENERACIÓN DE CÓDIGO]');
    const generador = new GeneradorCodigo();
    const codigoJs = generador.generar(ast);
    const comandosPlanos = generador.generarComandosPlanos(ast);
    console.log(' -> Exito: Codigo objeto y secuencia de control de vuelo generados.');
    console.log(' -> Comandos de bajo nivel para el Dron:');
    comandosPlanos.forEach((cmd, idx) => console.log(`    ${idx + 1}. ${cmd}`));
    console.log('\n---------------- CODIGO OBJETO GENERADO -------------');
    console.log(codigoJs);
    console.log('------------------------------------------------------');

    // ----------------------------------------------------
    // EJECUCIÓN DE LA MISIÓN
    // ----------------------------------------------------
    console.log('\n---------------- SIMULACION DE VUELO ----------------');
    const resultadoEjecucion = ejecutarCodigoGenerado(codigoJs);
    if (resultadoEjecucion.exito) {
      resultadoEjecucion.salidas.forEach(linea => console.log(linea));
    } else {
      console.error('[ERROR DE TIEMPO DE EJECUCIÓN]:', resultadoEjecucion.error);
    }
    console.log('------------------------------------------------------\n');

  } catch (error) {
    console.error('\n[FALLO EN EL PROCESO DE COMPILACIÓN]:');
    console.error(' -> ' + error.message);
    process.exit(1);
  }
}

// Lectura de argumentos de consola
const args = process.argv.slice(2);
const mostrarFases = args.includes('--fases') || args.includes('-f');
const archivoIndex = args.findIndex(a => !a.startsWith('-'));

if (archivoIndex !== -1) {
  const rutaArchivo = path.resolve(args[archivoIndex]);
  if (!fs.existsSync(rutaArchivo)) {
    console.error(`[ERROR] Archivo no encontrado: ${rutaArchivo}`);
    process.exit(1);
  }
  const codigo = fs.readFileSync(rutaArchivo, 'utf-8');
  compilar(codigo, { mostrarFases });
} else {
  console.log('Compilando Programa 1 por defecto (01_mision_simple.dron)...\n');
  const demoRuta = path.join(__dirname, 'ejemplos', '01_mision_simple.dron');
  const demoCodigo = fs.readFileSync(demoRuta, 'utf-8');
  compilar(demoCodigo, { mostrarFases: true });
}
