/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * TALLER II - ACTIVIDAD COLABORATIVA - FISICA II
 * Universidad de Manizales
 * Autor: Juan Pablo Gonzalez Rodriguez
 *
 * Contenido:
 *   EJ1: Onda Sonora - Desplazamiento maximo
 *   EJ2: Efecto Doppler - Frecuencia percibida
 *   EJ3: Espejo Concavo - Formacion de imagenes
 *   EJ4: Espejo Plano - Reflexion y simetria
 *   EJ5: Doble Rendija - Interferencia de Young
 *   EJ6: Lentes Delgadas - Ecuacion del fabricante
 * ═══════════════════════════════════════════════════════════════════════════════
 */

const SPREADSHEET_ID = "1KoBQ6zXErqFpJb-avnHIdztPrsdO0ifQBXD_fvox3IY";

const COLORES = {
  titulo: "#1a73e8",
  seccion: "#34a853",
  subseccion: "#1565c0",
  datos: "#fff9c4",
  resultado: "#c8e6c9",
  formula: "#e3f2fd",
  header: "#e8eaed",
  borde: "#9e9e9e",
  blanco: "#ffffff"
};

// ═══════════════════════════════════════════════════════════════════════════════
// MENU Y FUNCIONES DE ENTRADA
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Crea el menu personalizado al abrir el Spreadsheet.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Fisica II - Taller 2")
    .addItem("Desplegar Taller", "DesplegarTaller")
    .addSeparator()
    .addItem("Ejercicio 1 - Onda Sonora", "crearEJ1")
    .addItem("Ejercicio 2 - Efecto Doppler", "crearEJ2")
    .addItem("Ejercicio 3 - Espejo Concavo", "crearEJ3")
    .addItem("Ejercicio 4 - Espejo Plano", "crearEJ4")
    .addItem("Ejercicio 5 - Doble Rendija", "crearEJ5")
    .addItem("Ejercicio 6 - Lentes Delgadas", "crearEJ6")
    .addSeparator()
    .addItem("Crear Portada", "crearPortada")
    .addToUi();
}

/**
 * Genera todas las hojas del taller en secuencia.
 */
function DesplegarTaller() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  SpreadsheetApp.getActiveSpreadsheet().toast("Generando Taller Completo...", "Fisica II", -1);

  crearPortada();
  crearEJ1();
  crearEJ2();
  crearEJ3();
  crearEJ4();
  crearEJ5();
  crearEJ6();

  // Activar la portada al finalizar
  const portada = ss.getSheetByName("PORTADA");
  if (portada) ss.setActiveSheet(portada);

  SpreadsheetApp.getActiveSpreadsheet().toast("Taller generado exitosamente.", "Fisica II", 5);
}

/**
 * Crea la hoja de portada del taller.
 */
function crearPortada() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let hoja = ss.getSheetByName("PORTADA");
  if (!hoja) {
    hoja = ss.insertSheet("PORTADA", 0);
  } else {
    ss.setActiveSheet(hoja);
    ss.moveActiveSheet(1);
  }

  hoja.clear();
  hoja.clearFormats();
  hoja.setColumnWidth(1, 600);

  let fila = 5;

  hoja.getRange(fila, 1).setValue("UNIVERSIDAD DE MANIZALES")
    .setFontSize(20)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  fila += 2;

  hoja.getRange(fila, 1).setValue("Facultad de Ciencias e Ingenieria")
    .setFontSize(14)
    .setHorizontalAlignment("center");
  fila += 4;

  hoja.getRange(fila, 1).setValue("FISICA II")
    .setFontSize(24)
    .setFontWeight("bold")
    .setFontColor(COLORES.titulo)
    .setHorizontalAlignment("center");
  fila += 2;

  hoja.getRange(fila, 1).setValue("TALLER II - ACTIVIDAD COLABORATIVA")
    .setFontSize(18)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  fila += 4;

  hoja.getRange(fila, 1).setValue("Temas:")
    .setFontSize(12)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  fila++;

  const temas = [
    "Ondas Sonoras y Efecto Doppler",
    "Optica Geometrica: Espejos y Lentes",
    "Interferencia y Difraccion"
  ];

  for (const tema of temas) {
    hoja.getRange(fila, 1).setValue(tema)
      .setFontSize(11)
      .setHorizontalAlignment("center");
    fila++;
  }
  fila += 4;

  hoja.getRange(fila, 1).setValue("Presentado por:")
    .setFontSize(12)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  fila++;

  hoja.getRange(fila, 1).setValue("Juan Pablo Gonzalez Rodriguez - CC: 1151970526")
    .setFontSize(14)
    .setHorizontalAlignment("center");
  fila += 3;

  hoja.getRange(fila, 1).setValue("Profesor:")
    .setFontSize(12)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  fila++;

  hoja.getRange(fila, 1).setValue("HELVER AUGUSTO GIRALDO DAZA")
    .setFontSize(14)
    .setHorizontalAlignment("center");
  fila += 3;

  hoja.getRange(fila, 1).setValue("2025")
    .setFontSize(12)
    .setHorizontalAlignment("center");
}


// ═══════════════════════════════════════════════════════════════════════════════
// EJERCICIO 1: ONDA SONORA
// ═══════════════════════════════════════════════════════════════════════════════

function crearEJ1() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let hoja = ss.getSheetByName("EJ1");
  if (!hoja) {
    hoja = ss.insertSheet("EJ1");
  }

  hoja.clear();
  hoja.clearFormats();

  crearEJ1_(hoja);

  ss.setActiveSheet(hoja);
}

function crearEJ1_(hoja) {
  hoja.setColumnWidths(1, 6, 160);
  hoja.setColumnWidth(1, 190);
  hoja.setColumnWidth(2, 260);

  let fila = 1;

  // ═══════════════════════════════════════════════════════════════════════════
  // FILA 1: Titulo corto en A1
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange("A1").setValue("EJ1: Onda sonora")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.titulo);
  fila = 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // FILA 2: Titulo largo academico
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange("A2:F2").merge()
    .setValue("EJERCICIO 1: ONDA SONORA - CALCULO DEL DESPLAZAMIENTO MAXIMO (s_max)")
    .setFontSize(13)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.titulo)
    .setHorizontalAlignment("center");
  fila = 4;

  // ═══════════════════════════════════════════════════════════════════════════
  // A. ENUNCIADO
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("A. ENUNCIADO")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  const enunciadoLineas = [
    "En una onda sonora senoidal de moderada intensidad, las variaciones maximas de presion son del orden",
    "de 3.0 x 10^-2 Pa por arriba y por debajo de la presion atmosferica pa (nominalmente 1.013 x 10^5 Pa",
    "al nivel del mar). Calcule el desplazamiento maximo correspondiente, si la frecuencia es de 1000 Hz.",
    "En aire a presion atmosferica y densidad normales, la rapidez del sonido es de 344 m/s y el modulo",
    "de volumen es de 1.42 x 10^5 Pa."
  ];

  for (let i = 0; i < enunciadoLineas.length; i++) {
    hoja.getRange(fila, 1, 1, 6).merge()
      .setValue(enunciadoLineas[i])
      .setFontSize(10)
      .setWrap(true);
    fila++;
  }
  fila++;

  // ═══════════════════════════════════════════════════════════════════════════
  // B. IDENTIFICACION DEL MODELO FISICO
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("B. IDENTIFICACION DEL MODELO FISICO")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  hoja.getRange(fila, 1).setValue("Tipo de fenomeno:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Onda sonora longitudinal en un medio elastico (aire)");
  fila++;

  hoja.getRange(fila, 1).setValue("Modelo aplicado:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Propagacion de ondas mecanicas en medios continuos");
  fila++;

  hoja.getRange(fila, 1).setValue("Fundamento teorico:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Las ondas sonoras son perturbaciones mecanicas longitudinales donde las particulas del medio oscilan paralelamente a la direccion de propagacion. La variacion de presion Delta_P mantiene un desfase de pi/2 radianes (90 grados) respecto al desplazamiento s de las particulas.")
    .setWrap(true);
  hoja.setRowHeight(fila, 40);
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // C. DATOS DEL PROBLEMA
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("C. DATOS DEL PROBLEMA")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  const headerDatos = ["Simbolo", "Magnitud fisica", "Valor dado", "Unidad", "Valor SI", "Unidad SI"];
  hoja.getRange(fila, 1, 1, 6).setValues([headerDatos])
    .setFontWeight("bold")
    .setBackground(COLORES.header)
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  const filaInicioData = fila;
  const datosProblema = [
    ["Delta_P_max", "Variacion maxima de presion", "3.0 x 10^-2", "Pa", 0.03, "Pa"],
    ["f", "Frecuencia de la onda sonora", "1000", "Hz", 1000, "Hz"],
    ["v", "Velocidad del sonido en aire", "344", "m/s", 344, "m/s"],
    ["B", "Modulo de volumen del aire", "1.42 x 10^5", "Pa", 142000, "Pa"],
    ["p_a", "Presion atmosferica (referencia)", "1.013 x 10^5", "Pa", 101300, "Pa"]
  ];

  hoja.getRange(fila, 1, datosProblema.length, 6).setValues(datosProblema)
    .setBackground(COLORES.datos)
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  hoja.getRange(fila, 5, datosProblema.length, 1).setNumberFormat("0.00E+00");
  fila += datosProblema.length + 1;

  const celdaDPmax = "E" + filaInicioData;
  const celdaF = "E" + (filaInicioData + 1);
  const celdaV = "E" + (filaInicioData + 2);
  const celdaB = "E" + (filaInicioData + 3);

  // ═══════════════════════════════════════════════════════════════════════════
  // D. CONSTANTES Y RELACIONES
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("D. CONSTANTES Y RELACIONES")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  hoja.getRange(fila, 1).setValue("Densidad del aire (rho):").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Calculada mediante la relacion rho = B / v^2");
  fila++;

  hoja.getRange(fila, 1).setValue("Formula:");
  hoja.getRange(fila, 2).setValue("rho = B / v^2").setBackground(COLORES.formula);
  hoja.getRange(fila, 3).setValue("Valor:");
  hoja.getRange(fila, 4).setFormula("=" + celdaB + "/POWER(" + celdaV + ",2)")
    .setNumberFormat("0.000")
    .setBackground(COLORES.formula);
  hoja.getRange(fila, 5).setValue("kg/m^3");
  const filaDensidad = fila;
  fila++;

  hoja.getRange(fila, 1).setValue("Frecuencia angular (omega):").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("omega = 2 * pi * f");
  fila++;

  hoja.getRange(fila, 1).setValue("Formula:");
  hoja.getRange(fila, 2).setValue("omega = 2 * PI() * f").setBackground(COLORES.formula);
  hoja.getRange(fila, 3).setValue("Valor:");
  hoja.getRange(fila, 4).setFormula("=2*PI()*" + celdaF)
    .setNumberFormat("0.00")
    .setBackground(COLORES.formula);
  hoja.getRange(fila, 5).setValue("rad/s");
  const filaOmega = fila;
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // E. DIAGRAMA / ESQUEMA
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("E. DIAGRAMA / ESQUEMA")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  const diagrama = [
    ["REPRESENTACION DE LA ONDA SONORA LONGITUDINAL", "", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["Direccion de propagacion:  ---------------------->  x", "", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["Desplazamiento s(x,t):      ~~~~~~~~~~~~~~~~~~~~~~", "", "", "", "", ""],
    ["Variacion de presion P(x,t): /\\/\\/\\/\\/\\/\\/\\/\\/\\", "(desfase de 90 grados)", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["Relacion de fase:", "", "", "", "", ""],
    ["  - Cuando s = 0         -->  |Delta_P| = maximo", "", "", "", "", ""],
    ["  - Cuando s = s_max     -->  Delta_P = 0", "", "", "", "", ""]
  ];

  hoja.getRange(fila, 1, diagrama.length, 6).setValues(diagrama)
    .setFontFamily("Courier New")
    .setBackground("#f8f9fa")
    .setBorder(true, true, true, true, false, false, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  hoja.getRange(fila, 1).setFontWeight("bold");
  fila += diagrama.length + 1;

  // ═══════════════════════════════════════════════════════════════════════════
  // F. DESARROLLO PASO A PASO
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("F. DESARROLLO PASO A PASO")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  // --- PASO 1 ---
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 1: Ecuaciones fundamentales")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion)
    .setFontSize(10);
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("La relacion entre la amplitud de presion y la amplitud de desplazamiento en una onda sonora esta dada por:");
  fila++;

  hoja.getRange(fila, 1, 1, 3).merge()
    .setValue("Delta_P_max = rho * v * omega * s_max")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  hoja.getRange(fila, 4, 1, 3).merge()
    .setValue("Ecuacion de amplitud de presion");
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Donde rho es la densidad del medio, v la velocidad de propagacion, omega la frecuencia angular y s_max el desplazamiento maximo.");
  fila += 2;

  // --- PASO 2 ---
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 2: Despeje simbolico de la incognita s_max")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion)
    .setFontSize(10);
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Despejando s_max de la ecuacion fundamental:");
  fila++;

  hoja.getRange(fila, 1, 1, 3).merge()
    .setValue("s_max = Delta_P_max / (rho * v * omega)")
    .setFontWeight("bold")
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Sustituyendo las relaciones rho = B/v^2 y omega = 2*pi*f:");
  fila++;

  hoja.getRange(fila, 1, 1, 3).merge()
    .setValue("s_max = Delta_P_max / ((B/v^2) * v * 2*pi*f)")
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Simplificando algebraicamente:");
  fila++;

  hoja.getRange(fila, 1, 1, 3).merge()
    .setValue("s_max = (Delta_P_max * v) / (B * 2 * pi * f)")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  hoja.getRange(fila, 4, 1, 3).merge()
    .setValue("Expresion simplificada final");
  fila += 2;

  // --- PASO 3 ---
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 3: Sustitucion numerica")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion)
    .setFontSize(10);
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Sustituyendo los valores numericos en la expresion simplificada:");
  fila++;

  hoja.getRange(fila, 1).setValue("Delta_P_max =");
  hoja.getRange(fila, 2).setFormula("=" + celdaDPmax).setNumberFormat("0.00E+00");
  hoja.getRange(fila, 3).setValue("Pa");
  fila++;

  hoja.getRange(fila, 1).setValue("v =");
  hoja.getRange(fila, 2).setFormula("=" + celdaV);
  hoja.getRange(fila, 3).setValue("m/s");
  fila++;

  hoja.getRange(fila, 1).setValue("B =");
  hoja.getRange(fila, 2).setFormula("=" + celdaB).setNumberFormat("0.00E+00");
  hoja.getRange(fila, 3).setValue("Pa");
  fila++;

  hoja.getRange(fila, 1).setValue("f =");
  hoja.getRange(fila, 2).setFormula("=" + celdaF);
  hoja.getRange(fila, 3).setValue("Hz");
  fila += 2;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Calculo numerico:");
  fila++;

  hoja.getRange(fila, 1).setValue("s_max =");
  hoja.getRange(fila, 2, 1, 4).merge()
    .setValue("(Delta_P_max * v) / (B * 2 * PI() * f)");
  fila++;

  hoja.getRange(fila, 1).setValue("s_max =");
  const formulaCalculo = "=(" + celdaDPmax + "*" + celdaV + ")/(" + celdaB + "*2*PI()*" + celdaF + ")";
  hoja.getRange(fila, 2).setFormula(formulaCalculo)
    .setNumberFormat("0.000E+00")
    .setBackground(COLORES.formula)
    .setFontWeight("bold");
  hoja.getRange(fila, 3).setValue("m");
  const filaCalculo = fila;
  fila += 2;

  // --- PASO 4 ---
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 4: Resultado final")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion)
    .setFontSize(10);
  fila++;

  hoja.getRange(fila, 1, 1, 4).merge()
    .setValue("DESPLAZAMIENTO MAXIMO DE LAS PARTICULAS DEL AIRE")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORES.resultado);
  fila++;

  hoja.getRange(fila, 1).setValue("s_max =")
    .setFontWeight("bold")
    .setFontSize(13)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 2).setFormula("=B" + filaCalculo)
    .setNumberFormat("0.000E+00")
    .setFontWeight("bold")
    .setFontSize(13)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 3).setValue("m")
    .setFontWeight("bold")
    .setFontSize(13)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 4).setValue("")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 1, 1, 4)
    .setBorder(true, true, true, true, false, false, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  const filaResultado = fila;
  fila++;

  hoja.getRange(fila, 1).setValue("Equivalente:")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 2).setFormula("=B" + filaResultado + "*1E9")
    .setNumberFormat("0.00")
    .setFontWeight("bold")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 3).setValue("nm (nanometros)")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 4).setValue("")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 1, 1, 4)
    .setBorder(true, true, true, true, false, false, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID);
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // G. COMPROBACION FISICA
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("G. COMPROBACION FISICA")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  // Orden de magnitud
  hoja.getRange(fila, 1).setValue("Orden de magnitud:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("El resultado obtenido (aproximadamente 10^-8 m = 10 nm) corresponde al orden de magnitud esperado para desplazamientos moleculares en ondas sonoras de intensidad moderada.")
    .setWrap(true);
  hoja.setRowHeight(fila, 35);
  fila++;

  // Analisis dimensional
  hoja.getRange(fila, 1).setValue("Analisis dimensional:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("[s_max] = [Pa * (m/s)] / [Pa * Hz] = [Pa * m * s^-1] / [Pa * s^-1] = [m]")
    .setFontFamily("Courier New")
    .setBackground(COLORES.formula);
  fila++;

  hoja.getRange(fila, 1).setValue("Verificacion:");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Las unidades del resultado son metros [m], confirmando la consistencia dimensional de la expresion utilizada.");
  fila++;

  // Coherencia fisica
  hoja.getRange(fila, 1).setValue("Coherencia fisica:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("El desfase de 90 grados entre presion y desplazamiento tiene origen fisico: la maxima compresion (Delta_P maximo) ocurre cuando las particulas atraviesan la posicion de equilibrio (s = 0) con velocidad maxima. Cuando las particulas alcanzan su desplazamiento maximo (s = s_max), la presion local iguala la presion de equilibrio (Delta_P = 0).")
    .setWrap(true);
  hoja.setRowHeight(fila, 55);
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // CUADRO RESUMEN FINAL
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 4).merge()
    .setValue("RESUMEN DE RESULTADOS - EJERCICIO 1")
    .setFontWeight("bold")
    .setFontSize(10)
    .setBackground(COLORES.header)
    .setHorizontalAlignment("center");
  fila++;

  const resumenHeader = ["Magnitud", "Simbolo", "Valor", "Unidad"];
  hoja.getRange(fila, 1, 1, 4).setValues([resumenHeader])
    .setFontWeight("bold")
    .setBackground(COLORES.header)
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  hoja.getRange(fila, 1).setValue("Densidad del aire");
  hoja.getRange(fila, 2).setValue("rho");
  hoja.getRange(fila, 3).setFormula("=D" + filaDensidad).setNumberFormat("0.000");
  hoja.getRange(fila, 4).setValue("kg/m^3");
  hoja.getRange(fila, 1, 1, 4).setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  hoja.getRange(fila, 1).setValue("Frecuencia angular");
  hoja.getRange(fila, 2).setValue("omega");
  hoja.getRange(fila, 3).setFormula("=D" + filaOmega).setNumberFormat("0.00");
  hoja.getRange(fila, 4).setValue("rad/s");
  hoja.getRange(fila, 1, 1, 4).setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  hoja.getRange(fila, 1).setValue("Desplazamiento maximo")
    .setFontWeight("bold");
  hoja.getRange(fila, 2).setValue("s_max")
    .setFontWeight("bold");
  hoja.getRange(fila, 3).setFormula("=B" + filaResultado)
    .setNumberFormat("0.000E+00")
    .setFontWeight("bold");
  hoja.getRange(fila, 4).setValue("m")
    .setFontWeight("bold");
  hoja.getRange(fila, 1, 1, 4)
    .setBackground(COLORES.resultado)
    .setBorder(true, true, true, true, true, true, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  fila++;

  hoja.getRange(fila, 1).setValue("Desplazamiento maximo")
    .setFontWeight("bold");
  hoja.getRange(fila, 2).setValue("s_max")
    .setFontWeight("bold");
  hoja.getRange(fila, 3).setFormula("=B" + filaResultado + "*1E9")
    .setNumberFormat("0.00")
    .setFontWeight("bold");
  hoja.getRange(fila, 4).setValue("nm")
    .setFontWeight("bold");
  hoja.getRange(fila, 1, 1, 4)
    .setBackground(COLORES.resultado)
    .setBorder(true, true, true, true, true, true, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

// =============================================================================
// EJERCICIO 2: EFECTO DOPPLER - FUENTE EN MOVIMIENTO
// =============================================================================

function crearEJ2() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let hoja = ss.getSheetByName("EJ2");
  if (!hoja) {
    hoja = ss.insertSheet("EJ2");
  }

  hoja.clear();
  hoja.clearFormats();

  crearEJ2_(hoja);

  ss.setActiveSheet(hoja);
}

function crearEJ2_(hoja) {
  hoja.setColumnWidths(1, 6, 160);
  hoja.setColumnWidth(1, 190);
  hoja.setColumnWidth(2, 260);

  let fila = 1;

  // ═══════════════════════════════════════════════════════════════════════════
  // FILA 1: Titulo corto en A1
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange("A1").setValue("EJ2: Efecto Doppler")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.titulo);
  fila = 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // FILA 2: Titulo largo academico
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange("A2:F2").merge()
    .setValue("EJERCICIO 2: EFECTO DOPPLER - LONGITUD DE ONDA CON FUENTE EN MOVIMIENTO")
    .setFontSize(13)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.titulo)
    .setHorizontalAlignment("center");
  fila = 4;

  // ═══════════════════════════════════════════════════════════════════════════
  // A. ENUNCIADO
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("A. ENUNCIADO")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  const enunciadoLineas = [
    "Una sirena de policia emite una onda senoidal con frecuencia f_s = 300 Hz.",
    "La rapidez del sonido es v = 340 m/s.",
    "",
    "a) Calcule la longitud de onda del sonido si la sirena esta en reposo en el aire.",
    "b) Si la sirena se mueve a 30 m/s, calcule las longitudes de onda para las ondas",
    "   adelante y atras de la fuente."
  ];

  for (let i = 0; i < enunciadoLineas.length; i++) {
    hoja.getRange(fila, 1, 1, 6).merge()
      .setValue(enunciadoLineas[i])
      .setFontSize(10)
      .setWrap(true);
    fila++;
  }
  fila++;

  // ═══════════════════════════════════════════════════════════════════════════
  // B. IDENTIFICACION DEL MODELO FISICO
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("B. IDENTIFICACION DEL MODELO FISICO")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  hoja.getRange(fila, 1).setValue("Tipo de fenomeno:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Onda sonora longitudinal en un medio elastico (aire)");
  fila++;

  hoja.getRange(fila, 1).setValue("Modelo aplicado:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Efecto Doppler para fuente en movimiento y observador en reposo");
  fila++;

  hoja.getRange(fila, 1).setValue("Fundamento teorico:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Cuando una fuente sonora se mueve respecto al medio de propagacion, la longitud de onda medida por un observador estacionario cambia. La frecuencia de emision f_s permanece constante (propiedad intrinseca de la fuente), pero el movimiento comprime los frentes de onda en la direccion del movimiento y los expande en la direccion opuesta.")
    .setWrap(true);
  hoja.setRowHeight(fila, 50);
  fila++;

  hoja.getRange(fila, 1).setValue("Nota importante:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("La frecuencia de emision NO cambia. Lo que se modifica es la distribucion espacial de los frentes de onda (longitud de onda).")
    .setWrap(true)
    .setFontStyle("italic");
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // C. DATOS DEL PROBLEMA
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("C. DATOS DEL PROBLEMA")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  const headerDatos = ["Simbolo", "Magnitud fisica", "Valor dado", "Unidad", "Valor SI", "Unidad SI"];
  hoja.getRange(fila, 1, 1, 6).setValues([headerDatos])
    .setFontWeight("bold")
    .setBackground(COLORES.header)
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  const filaInicioData = fila;
  const datosProblema = [
    ["f_s", "Frecuencia de emision de la sirena", "300", "Hz", 300, "Hz (s^-1)"],
    ["v", "Velocidad del sonido en aire", "340", "m/s", 340, "m/s"],
    ["v_s", "Velocidad de la sirena (fuente)", "30", "m/s", 30, "m/s"]
  ];

  hoja.getRange(fila, 1, datosProblema.length, 6).setValues(datosProblema)
    .setBackground(COLORES.datos)
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila += datosProblema.length;

  // Incognitas
  fila++;
  hoja.getRange(fila, 1).setValue("Incognitas:").setFontWeight("bold");
  fila++;
  const incognitas = [
    ["lambda_0", "Longitud de onda con sirena en reposo", "?", "m", "", ""],
    ["lambda_adelante", "Longitud de onda delante de la sirena", "?", "m", "", ""],
    ["lambda_atras", "Longitud de onda detras de la sirena", "?", "m", "", ""]
  ];
  hoja.getRange(fila, 1, incognitas.length, 6).setValues(incognitas)
    .setBackground("#e8f5e9")
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila += incognitas.length + 1;

  const celdaFs = "E" + filaInicioData;
  const celdaV = "E" + (filaInicioData + 1);
  const celdaVs = "E" + (filaInicioData + 2);

  // ═══════════════════════════════════════════════════════════════════════════
  // D. CONSTANTES Y RELACIONES
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("D. CONSTANTES Y RELACIONES FUNDAMENTALES")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  hoja.getRange(fila, 1).setValue("Relacion fundamental:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("v = lambda * f  (velocidad = longitud de onda x frecuencia)");
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("")
    .setBackground(COLORES.blanco);
  fila++;

  hoja.getRange(fila, 1).setValue("Fuente en reposo:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 2).merge()
    .setValue("lambda_0 = v / f_s")
    .setBackground(COLORES.formula)
    .setFontWeight("bold");
  hoja.getRange(fila, 4, 1, 3).merge()
    .setValue("Longitud de onda sin efecto Doppler");
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("")
    .setBackground(COLORES.blanco);
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Fuente en movimiento (Efecto Doppler):")
    .setFontWeight("bold");
  fila++;

  hoja.getRange(fila, 1).setValue("Ondas adelante:");
  hoja.getRange(fila, 2, 1, 2).merge()
    .setValue("lambda_adelante = (v - v_s) / f_s")
    .setBackground(COLORES.formula)
    .setFontWeight("bold");
  hoja.getRange(fila, 4, 1, 3).merge()
    .setValue("Frentes comprimidos (fuente acercandose)");
  fila++;

  hoja.getRange(fila, 1).setValue("Ondas atras:");
  hoja.getRange(fila, 2, 1, 2).merge()
    .setValue("lambda_atras = (v + v_s) / f_s")
    .setBackground(COLORES.formula)
    .setFontWeight("bold");
  hoja.getRange(fila, 4, 1, 3).merge()
    .setValue("Frentes expandidos (fuente alejandose)");
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("")
    .setBackground(COLORES.blanco);
  fila++;

  hoja.getRange(fila, 1).setValue("Interpretacion:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("En un periodo T=1/f_s, la fuente emite un frente de onda. Durante ese tiempo, el frente viaja v*T pero la fuente se mueve v_s*T. Adelante los frentes se comprimen; atras se expanden.")
    .setWrap(true);
  hoja.setRowHeight(fila, 40);
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // E. DIAGRAMA / ESQUEMA
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("E. DIAGRAMA / ESQUEMA")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  const diagrama = [
    ["EFECTO DOPPLER - FUENTE EN MOVIMIENTO", "", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["CASO A: SIRENA EN REPOSO", "", "", "", "", ""],
    ["    |     |     |  S  |     |     |", "", "", "", "", ""],
    ["    <--- lambda_0 --->", "", "", "", "", ""],
    ["    Frentes de onda uniformemente espaciados", "", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["CASO B: SIRENA EN MOVIMIENTO (---> v_s)", "", "", "", "", ""],
    ["    |      |      |   | S |  |  |  |", "", "", "", "", ""],
    ["    <-- lambda_atras      lambda_adelante -->", "", "", "", "", ""],
    ["         (expandida)        (comprimida)", "", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["RELACION:  lambda_adelante < lambda_0 < lambda_atras", "", "", "", "", ""]
  ];

  hoja.getRange(fila, 1, diagrama.length, 6).setValues(diagrama)
    .setFontFamily("Courier New")
    .setBackground("#f8f9fa")
    .setBorder(true, true, true, true, false, false, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  hoja.getRange(fila, 1).setFontWeight("bold");
  hoja.getRange(fila + 2, 1).setFontWeight("bold");
  hoja.getRange(fila + 7, 1).setFontWeight("bold");
  hoja.getRange(fila + 12, 1).setFontWeight("bold").setFontColor(COLORES.subseccion);
  fila += diagrama.length + 1;

  // ═══════════════════════════════════════════════════════════════════════════
  // F. DESARROLLO PASO A PASO
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("F. DESARROLLO PASO A PASO")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  // --- INCISO A ---
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("INCISO a) SIRENA EN REPOSO")
    .setFontWeight("bold")
    .setFontColor(COLORES.titulo)
    .setFontSize(11)
    .setBackground("#e3f2fd");
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 1: Identificar la ecuacion aplicable")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Cuando la fuente esta en reposo, la longitud de onda se calcula con la relacion fundamental:");
  fila++;

  hoja.getRange(fila, 1, 1, 3).merge()
    .setValue("lambda_0 = v / f_s")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 2: Sustitucion numerica")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1).setValue("lambda_0 =");
  hoja.getRange(fila, 2).setValue("v / f_s");
  fila++;

  hoja.getRange(fila, 1).setValue("lambda_0 =");
  hoja.getRange(fila, 2).setValue("(340 m/s) / (300 Hz)");
  fila++;

  hoja.getRange(fila, 1).setValue("lambda_0 =");
  hoja.getRange(fila, 2).setFormula("=" + celdaV + "/" + celdaFs)
    .setNumberFormat("0.0000")
    .setBackground(COLORES.formula)
    .setFontWeight("bold");
  hoja.getRange(fila, 3).setValue("m");
  const filaLambda0 = fila;
  fila += 2;

  // --- INCISO B.1 ---
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("INCISO b.1) LONGITUD DE ONDA ADELANTE DE LA FUENTE")
    .setFontWeight("bold")
    .setFontColor(COLORES.titulo)
    .setFontSize(11)
    .setBackground("#e3f2fd");
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 1: Identificar la ecuacion aplicable")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Delante de la fuente en movimiento, los frentes de onda se comprimen:");
  fila++;

  hoja.getRange(fila, 1, 1, 3).merge()
    .setValue("lambda_adelante = (v - v_s) / f_s")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 2: Sustitucion numerica")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1).setValue("lambda_adelante =");
  hoja.getRange(fila, 2).setValue("(v - v_s) / f_s");
  fila++;

  hoja.getRange(fila, 1).setValue("lambda_adelante =");
  hoja.getRange(fila, 2).setValue("(340 - 30) m/s / (300 Hz)");
  fila++;

  hoja.getRange(fila, 1).setValue("lambda_adelante =");
  hoja.getRange(fila, 2).setValue("(310 m/s) / (300 Hz)");
  fila++;

  hoja.getRange(fila, 1).setValue("lambda_adelante =");
  hoja.getRange(fila, 2).setFormula("=(" + celdaV + "-" + celdaVs + ")/" + celdaFs)
    .setNumberFormat("0.0000")
    .setBackground(COLORES.formula)
    .setFontWeight("bold");
  hoja.getRange(fila, 3).setValue("m");
  const filaLambdaAdelante = fila;
  fila += 2;

  // --- INCISO B.2 ---
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("INCISO b.2) LONGITUD DE ONDA ATRAS DE LA FUENTE")
    .setFontWeight("bold")
    .setFontColor(COLORES.titulo)
    .setFontSize(11)
    .setBackground("#e3f2fd");
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 1: Identificar la ecuacion aplicable")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Detras de la fuente en movimiento, los frentes de onda se expanden:");
  fila++;

  hoja.getRange(fila, 1, 1, 3).merge()
    .setValue("lambda_atras = (v + v_s) / f_s")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 2: Sustitucion numerica")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1).setValue("lambda_atras =");
  hoja.getRange(fila, 2).setValue("(v + v_s) / f_s");
  fila++;

  hoja.getRange(fila, 1).setValue("lambda_atras =");
  hoja.getRange(fila, 2).setValue("(340 + 30) m/s / (300 Hz)");
  fila++;

  hoja.getRange(fila, 1).setValue("lambda_atras =");
  hoja.getRange(fila, 2).setValue("(370 m/s) / (300 Hz)");
  fila++;

  hoja.getRange(fila, 1).setValue("lambda_atras =");
  hoja.getRange(fila, 2).setFormula("=(" + celdaV + "+" + celdaVs + ")/" + celdaFs)
    .setNumberFormat("0.0000")
    .setBackground(COLORES.formula)
    .setFontWeight("bold");
  hoja.getRange(fila, 3).setValue("m");
  const filaLambdaAtras = fila;
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // G. RESULTADOS FINALES
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("G. RESULTADOS FINALES")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  const headerRes = ["Magnitud", "Simbolo", "Formula", "Valor", "Unidad"];
  hoja.getRange(fila, 1, 1, 5).setValues([headerRes])
    .setFontWeight("bold")
    .setBackground(COLORES.header)
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  hoja.getRange(fila, 1).setValue("Longitud de onda en reposo");
  hoja.getRange(fila, 2).setValue("lambda_0");
  hoja.getRange(fila, 3).setValue("v / f_s");
  hoja.getRange(fila, 4).setFormula("=B" + filaLambda0).setNumberFormat("0.000");
  hoja.getRange(fila, 5).setValue("m");
  hoja.getRange(fila, 1, 1, 5)
    .setBackground(COLORES.resultado)
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  hoja.getRange(fila, 1).setValue("Longitud de onda adelante");
  hoja.getRange(fila, 2).setValue("lambda_adelante");
  hoja.getRange(fila, 3).setValue("(v - v_s) / f_s");
  hoja.getRange(fila, 4).setFormula("=B" + filaLambdaAdelante).setNumberFormat("0.000");
  hoja.getRange(fila, 5).setValue("m");
  hoja.getRange(fila, 1, 1, 5)
    .setBackground(COLORES.resultado)
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  hoja.getRange(fila, 1).setValue("Longitud de onda atras");
  hoja.getRange(fila, 2).setValue("lambda_atras");
  hoja.getRange(fila, 3).setValue("(v + v_s) / f_s");
  hoja.getRange(fila, 4).setFormula("=B" + filaLambdaAtras).setNumberFormat("0.000");
  hoja.getRange(fila, 5).setValue("m");
  hoja.getRange(fila, 1, 1, 5)
    .setBackground(COLORES.resultado)
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  // Verificacion de relacion
  fila++;
  hoja.getRange(fila, 1).setValue("Verificacion:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 4).merge()
    .setValue("lambda_adelante < lambda_0 < lambda_atras")
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center")
    .setFontWeight("bold");
  fila++;

  hoja.getRange(fila, 1).setValue("Comprobacion:");
  hoja.getRange(fila, 2).setFormula("=B" + filaLambdaAdelante).setNumberFormat("0.000");
  hoja.getRange(fila, 3).setValue("<");
  hoja.getRange(fila, 4).setFormula("=B" + filaLambda0).setNumberFormat("0.000");
  hoja.getRange(fila, 5).setValue("<");
  hoja.getRange(fila, 6).setFormula("=B" + filaLambdaAtras).setNumberFormat("0.000");
  fila++;

  hoja.getRange(fila, 2, 1, 5).merge()
    .setFormula("=IF(AND(B" + filaLambdaAdelante + "<B" + filaLambda0 + ", B" + filaLambda0 + "<B" + filaLambdaAtras + "), \"CORRECTO\", \"ERROR\")")
    .setFontWeight("bold")
    .setFontColor("#2e7d32")
    .setHorizontalAlignment("center");
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // H. COMPROBACION FISICA
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("H. COMPROBACION FISICA Y DIMENSIONAL")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  // Analisis dimensional
  hoja.getRange(fila, 1).setValue("Analisis dimensional:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("[lambda] = [v] / [f] = (m/s) / (s^-1) = (m/s) * s = m")
    .setFontFamily("Courier New")
    .setBackground(COLORES.formula);
  fila++;

  hoja.getRange(fila, 1).setValue("Verificacion:");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Las unidades del resultado son metros [m], confirmando la consistencia dimensional.");
  fila++;

  // Coherencia fisica
  hoja.getRange(fila, 1).setValue("Coherencia fisica:").setFontWeight("bold");
  fila++;

  hoja.getRange(fila, 1).setValue("1. Compresion adelante:");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("La sirena se mueve en la misma direccion que las ondas emitidas hacia adelante. Esto reduce la distancia entre frentes de onda sucesivos, resultando en lambda_adelante < lambda_0.")
    .setWrap(true);
  hoja.setRowHeight(fila, 35);
  fila++;

  hoja.getRange(fila, 1).setValue("2. Expansion atras:");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("La sirena se aleja de las ondas emitidas hacia atras. Esto aumenta la distancia entre frentes de onda sucesivos, resultando en lambda_atras > lambda_0.")
    .setWrap(true);
  hoja.setRowHeight(fila, 35);
  fila++;

  hoja.getRange(fila, 1).setValue("3. Limite fisico:");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Si v_s = v (velocidad sonica), lambda_adelante = 0 (boom sonico). Aqui v_s = 30 m/s << v = 340 m/s, regimen subsonico.")
    .setWrap(true);
  fila++;

  hoja.getRange(fila, 1).setValue("4. Frecuencia percibida:");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Un observador adelante percibe f' = v/lambda_adelante (mayor). Un observador atras percibe f' = v/lambda_atras (menor).")
    .setWrap(true);
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // CUADRO RESUMEN FINAL
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 5).merge()
    .setValue("RESUMEN DE RESULTADOS - EJERCICIO 2")
    .setFontWeight("bold")
    .setFontSize(10)
    .setBackground(COLORES.header)
    .setHorizontalAlignment("center");
  fila++;

  const resumenHeader = ["Magnitud", "Simbolo", "Valor", "Unidad", "Observacion"];
  hoja.getRange(fila, 1, 1, 5).setValues([resumenHeader])
    .setFontWeight("bold")
    .setBackground(COLORES.header)
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  hoja.getRange(fila, 1).setValue("Longitud de onda reposo");
  hoja.getRange(fila, 2).setValue("lambda_0").setFontWeight("bold");
  hoja.getRange(fila, 3).setFormula("=B" + filaLambda0).setNumberFormat("0.000").setFontWeight("bold");
  hoja.getRange(fila, 4).setValue("m").setFontWeight("bold");
  hoja.getRange(fila, 5).setValue("Referencia");
  hoja.getRange(fila, 1, 1, 5)
    .setBackground(COLORES.resultado)
    .setBorder(true, true, true, true, true, true, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  fila++;

  hoja.getRange(fila, 1).setValue("Longitud de onda adelante");
  hoja.getRange(fila, 2).setValue("lambda_adelante").setFontWeight("bold");
  hoja.getRange(fila, 3).setFormula("=B" + filaLambdaAdelante).setNumberFormat("0.000").setFontWeight("bold");
  hoja.getRange(fila, 4).setValue("m").setFontWeight("bold");
  hoja.getRange(fila, 5).setValue("Comprimida (-8.8%)");
  hoja.getRange(fila, 1, 1, 5)
    .setBackground(COLORES.resultado)
    .setBorder(true, true, true, true, true, true, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  fila++;

  hoja.getRange(fila, 1).setValue("Longitud de onda atras");
  hoja.getRange(fila, 2).setValue("lambda_atras").setFontWeight("bold");
  hoja.getRange(fila, 3).setFormula("=B" + filaLambdaAtras).setNumberFormat("0.000").setFontWeight("bold");
  hoja.getRange(fila, 4).setValue("m").setFontWeight("bold");
  hoja.getRange(fila, 5).setValue("Expandida (+8.8%)");
  hoja.getRange(fila, 1, 1, 5)
    .setBackground(COLORES.resultado)
    .setBorder(true, true, true, true, true, true, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

// =============================================================================
// EJERCICIO 3: ESPEJO CONCAVO - OPTICA GEOMETRICA
// =============================================================================

function crearEJ3() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let hoja = ss.getSheetByName("EJ3");
  if (!hoja) {
    hoja = ss.insertSheet("EJ3");
  }

  hoja.clear();
  hoja.clearFormats();

  crearEJ3_(hoja);

  ss.setActiveSheet(hoja);
}

function crearEJ3_(hoja) {
  hoja.setColumnWidths(1, 6, 160);
  hoja.setColumnWidth(1, 190);
  hoja.setColumnWidth(2, 260);

  let fila = 1;

  // ═══════════════════════════════════════════════════════════════════════════
  // FILA 1: Titulo corto en A1
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange("A1").setValue("EJ3: Espejo concavo")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.titulo);
  fila = 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // FILA 2: Titulo largo academico
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange("A2:F2").merge()
    .setValue("EJERCICIO 3: ESPEJO CONCAVO - RADIO DE CURVATURA, DISTANCIA FOCAL Y ALTURA DE IMAGEN")
    .setFontSize(13)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.titulo)
    .setHorizontalAlignment("center");
  fila = 4;

  // ═══════════════════════════════════════════════════════════════════════════
  // A. ENUNCIADO
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("A. ENUNCIADO")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  const enunciadoLineas = [
    "Un espejo concavo forma una imagen, sobre una pared situada a 3.00 m del espejo,",
    "del filamento de una lampara de reflector que esta a 10.0 cm delante del espejo.",
    "",
    "a) Cuales son el radio de curvatura y la distancia focal del espejo?",
    "b) Cual es la altura de la imagen, si la altura del objeto es de 5.00 mm?"
  ];

  for (let i = 0; i < enunciadoLineas.length; i++) {
    hoja.getRange(fila, 1, 1, 6).merge()
      .setValue(enunciadoLineas[i])
      .setFontSize(10)
      .setWrap(true);
    fila++;
  }
  fila++;

  // ═══════════════════════════════════════════════════════════════════════════
  // B. IDENTIFICACION DEL MODELO FISICO
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("B. IDENTIFICACION DEL MODELO FISICO")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  hoja.getRange(fila, 1).setValue("Tipo de fenomeno:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Optica geometrica - Reflexion en espejo esferico concavo");
  fila++;

  hoja.getRange(fila, 1).setValue("Modelo aplicado:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Formacion de imagenes por reflexion en superficies esfericas");
  fila++;

  hoja.getRange(fila, 1).setValue("Fundamento teorico:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("El espejo concavo enfoca rayos paralelos hacia el foco. La ecuacion de Gauss para espejos relaciona las posiciones del objeto (p), la imagen (q) y la distancia focal (f). El aumento lateral M relaciona los tamanos de objeto e imagen.")
    .setWrap(true);
  hoja.setRowHeight(fila, 45);
  fila++;

  hoja.getRange(fila, 1).setValue("Convencion de signos:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("p > 0 (objeto real, frente al espejo), q > 0 (imagen real, frente al espejo), f > 0 (espejo concavo), M < 0 (imagen invertida)")
    .setWrap(true)
    .setFontStyle("italic");
  hoja.setRowHeight(fila, 35);
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // C. DATOS DEL PROBLEMA
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("C. DATOS DEL PROBLEMA")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  const headerDatos = ["Simbolo", "Magnitud fisica", "Valor dado", "Unidad", "Valor SI", "Unidad SI"];
  hoja.getRange(fila, 1, 1, 6).setValues([headerDatos])
    .setFontWeight("bold")
    .setBackground(COLORES.header)
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  const filaInicioData = fila;
  const datosProblema = [
    ["q", "Distancia de la imagen (pared)", "3.00", "m", 3.00, "m"],
    ["p", "Distancia del objeto (filamento)", "10.0", "cm", 0.10, "m"],
    ["h", "Altura del objeto (filamento)", "5.00", "mm", 0.005, "m"]
  ];

  hoja.getRange(fila, 1, datosProblema.length, 6).setValues(datosProblema)
    .setBackground(COLORES.datos)
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila += datosProblema.length;

  // Incognitas
  fila++;
  hoja.getRange(fila, 1).setValue("Incognitas:").setFontWeight("bold");
  fila++;
  const incognitas = [
    ["f", "Distancia focal del espejo", "?", "m", "", ""],
    ["R", "Radio de curvatura del espejo", "?", "m", "", ""],
    ["h'", "Altura de la imagen", "?", "m", "", ""]
  ];
  hoja.getRange(fila, 1, incognitas.length, 6).setValues(incognitas)
    .setBackground("#e8f5e9")
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila += incognitas.length + 1;

  const celdaQ = "E" + filaInicioData;
  const celdaP = "E" + (filaInicioData + 1);
  const celdaH = "E" + (filaInicioData + 2);

  // ═══════════════════════════════════════════════════════════════════════════
  // D. CONSTANTES Y RELACIONES
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("D. ECUACIONES FUNDAMENTALES")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  hoja.getRange(fila, 1).setValue("Ecuacion de Gauss:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 2).merge()
    .setValue("1/p + 1/q = 1/f")
    .setBackground(COLORES.formula)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  hoja.getRange(fila, 4, 1, 3).merge()
    .setValue("Relaciona posiciones de objeto, imagen y foco");
  fila++;

  hoja.getRange(fila, 1).setValue("Relacion radio-foco:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 2).merge()
    .setValue("R = 2f")
    .setBackground(COLORES.formula)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  hoja.getRange(fila, 4, 1, 3).merge()
    .setValue("El centro de curvatura esta al doble del foco");
  fila++;

  hoja.getRange(fila, 1).setValue("Aumento lateral:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 2).merge()
    .setValue("M = h'/h = -q/p")
    .setBackground(COLORES.formula)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  hoja.getRange(fila, 4, 1, 3).merge()
    .setValue("Signo negativo indica imagen invertida");
  fila++;

  hoja.getRange(fila, 1).setValue("Despeje de f:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 2).merge()
    .setValue("f = (p * q) / (p + q)")
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  hoja.getRange(fila, 4, 1, 3).merge()
    .setValue("Forma explicita para calcular f");
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // E. DIAGRAMA / ESQUEMA
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("E. DIAGRAMA / ESQUEMA")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  const diagrama = [
    ["ESPEJO CONCAVO - FORMACION DE IMAGEN REAL", "", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["                    Espejo", "", "", "", "", ""],
    ["     Objeto           concavo                  Imagen", "", "", "", "", ""],
    ["        |               )                         |", "", "", "", "", ""],
    ["        | h             )                         | h' (invertida)", "", "", "", "", ""],
    ["   -----+-----     -----)----- F ----- C     -----+-----", "", "", "", "", ""],
    ["        |               )      |       |          |", "", "", "", "", ""],
    ["       \\|/              )      f       R         \\|/", "", "", "", "", ""],
    ["        v               )                         v", "", "", "", "", ""],
    ["    <-- p -->           )      <------ q ------>", "", "", "", "", ""],
    ["     (10 cm)                        (3.00 m)", "", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["Nota: F = foco, C = centro de curvatura, R = 2f", "", "", "", "", ""]
  ];

  hoja.getRange(fila, 1, diagrama.length, 6).setValues(diagrama)
    .setFontFamily("Courier New")
    .setBackground("#f8f9fa")
    .setBorder(true, true, true, true, false, false, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  hoja.getRange(fila, 1).setFontWeight("bold");
  fila += diagrama.length + 1;

  // ═══════════════════════════════════════════════════════════════════════════
  // F. DESARROLLO PASO A PASO
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("F. DESARROLLO PASO A PASO")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  // === PARTE A ===
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("PARTE (a): RADIO DE CURVATURA Y DISTANCIA FOCAL")
    .setFontWeight("bold")
    .setFontColor(COLORES.titulo)
    .setFontSize(11)
    .setBackground("#e3f2fd");
  fila++;

  // Paso 1
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 1: Identificar la ecuacion aplicable")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Usamos la ecuacion de Gauss para espejos:");
  fila++;

  hoja.getRange(fila, 1, 1, 3).merge()
    .setValue("1/p + 1/q = 1/f")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  fila++;

  // Paso 2
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 2: Despeje de la distancia focal f")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("De 1/f = 1/p + 1/q, despejamos f:");
  fila++;

  hoja.getRange(fila, 1, 1, 3).merge()
    .setValue("f = (p * q) / (p + q)")
    .setFontWeight("bold")
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  fila++;

  // Paso 3
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 3: Sustitucion numerica para f")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1).setValue("p =");
  hoja.getRange(fila, 2).setFormula("=" + celdaP).setNumberFormat("0.00");
  hoja.getRange(fila, 3).setValue("m");
  hoja.getRange(fila, 4).setValue("(10.0 cm convertido a m)");
  fila++;

  hoja.getRange(fila, 1).setValue("q =");
  hoja.getRange(fila, 2).setFormula("=" + celdaQ).setNumberFormat("0.00");
  hoja.getRange(fila, 3).setValue("m");
  fila++;

  hoja.getRange(fila, 1).setValue("f =");
  hoja.getRange(fila, 2).setValue("(p * q) / (p + q)");
  fila++;

  hoja.getRange(fila, 1).setValue("f =");
  hoja.getRange(fila, 2).setValue("(0.10 * 3.00) / (0.10 + 3.00)");
  fila++;

  hoja.getRange(fila, 1).setValue("f =");
  hoja.getRange(fila, 2).setFormula("=(" + celdaP + "*" + celdaQ + ")/(" + celdaP + "+" + celdaQ + ")")
    .setNumberFormat("0.0000")
    .setBackground(COLORES.formula)
    .setFontWeight("bold");
  hoja.getRange(fila, 3).setValue("m");
  const filaF = fila;
  fila++;

  // Paso 4: Radio de curvatura
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 4: Calculo del radio de curvatura R")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("El radio de curvatura es el doble de la distancia focal:");
  fila++;

  hoja.getRange(fila, 1).setValue("R = 2f =");
  hoja.getRange(fila, 2).setFormula("=2*B" + filaF)
    .setNumberFormat("0.0000")
    .setBackground(COLORES.formula)
    .setFontWeight("bold");
  hoja.getRange(fila, 3).setValue("m");
  const filaR = fila;
  fila++;

  // Resultados parte a
  fila++;
  hoja.getRange(fila, 1, 1, 4).merge()
    .setValue("RESULTADOS PARTE (a)")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORES.resultado);
  fila++;

  hoja.getRange(fila, 1).setValue("f =")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 2).setFormula("=B" + filaF + "*100")
    .setNumberFormat("0.00")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 3).setValue("cm")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 4).setValue("")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 1, 1, 4)
    .setBorder(true, true, true, true, false, false, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  fila++;

  hoja.getRange(fila, 1).setValue("R =")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 2).setFormula("=B" + filaR + "*100")
    .setNumberFormat("0.00")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 3).setValue("cm")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 4).setValue("")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 1, 1, 4)
    .setBorder(true, true, true, true, false, false, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  fila += 2;

  // === PARTE B ===
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("PARTE (b): ALTURA DE LA IMAGEN")
    .setFontWeight("bold")
    .setFontColor(COLORES.titulo)
    .setFontSize(11)
    .setBackground("#e3f2fd");
  fila++;

  // Paso 1
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 1: Ecuacion del aumento lateral")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("El aumento lateral M relaciona los tamanos de objeto e imagen:");
  fila++;

  hoja.getRange(fila, 1, 1, 3).merge()
    .setValue("M = h'/h = -q/p")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  fila++;

  // Paso 2
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 2: Despeje de la altura de imagen h'")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1, 1, 3).merge()
    .setValue("h' = M * h = -(q/p) * h")
    .setFontWeight("bold")
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  fila++;

  // Paso 3
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 3: Calculo del aumento lateral M")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1).setValue("M =");
  hoja.getRange(fila, 2).setValue("-q / p");
  fila++;

  hoja.getRange(fila, 1).setValue("M =");
  hoja.getRange(fila, 2).setValue("-3.00 / 0.10");
  fila++;

  hoja.getRange(fila, 1).setValue("M =");
  hoja.getRange(fila, 2).setFormula("=-" + celdaQ + "/" + celdaP)
    .setNumberFormat("0.00")
    .setBackground(COLORES.formula)
    .setFontWeight("bold");
  hoja.getRange(fila, 3).setValue("(adimensional)");
  const filaM = fila;
  fila++;

  // Paso 4
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 4: Calculo de la altura de imagen h'")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1).setValue("h' =");
  hoja.getRange(fila, 2).setValue("M * h");
  fila++;

  hoja.getRange(fila, 1).setValue("h' =");
  hoja.getRange(fila, 2).setValue("(-30) * (0.005 m)");
  fila++;

  hoja.getRange(fila, 1).setValue("h' =");
  hoja.getRange(fila, 2).setFormula("=B" + filaM + "*" + celdaH)
    .setNumberFormat("0.000")
    .setBackground(COLORES.formula)
    .setFontWeight("bold");
  hoja.getRange(fila, 3).setValue("m");
  const filaHprima = fila;
  fila++;

  // Resultado parte b
  fila++;
  hoja.getRange(fila, 1, 1, 4).merge()
    .setValue("RESULTADO PARTE (b)")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORES.resultado);
  fila++;

  hoja.getRange(fila, 1).setValue("|h'| =")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 2).setFormula("=ABS(B" + filaHprima + ")*100")
    .setNumberFormat("0.00")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 3).setValue("cm")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 4).setValue("")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 1, 1, 4)
    .setBorder(true, true, true, true, false, false, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  fila++;

  hoja.getRange(fila, 1).setValue("Equivalente:")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 2).setFormula("=ABS(B" + filaHprima + ")*1000")
    .setNumberFormat("0.0")
    .setFontWeight("bold")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 3).setValue("mm")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 4).setValue("")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 1, 1, 4)
    .setBorder(true, true, true, true, false, false, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  // Interpretacion
  hoja.getRange(fila, 1).setValue("Interpretacion:");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("M < 0 indica imagen INVERTIDA y REAL");
  fila++;

  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("|M| = 30 indica imagen AUMENTADA (30 veces mas grande que el objeto)");
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // G. COMPROBACION FISICA
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("G. COMPROBACION FISICA Y DIMENSIONAL")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  // Analisis dimensional
  hoja.getRange(fila, 1).setValue("Analisis dimensional:").setFontWeight("bold");
  fila++;

  hoja.getRange(fila, 1).setValue("Para f:");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("[f] = [p*q]/[p+q] = (m*m)/(m) = m")
    .setFontFamily("Courier New")
    .setBackground(COLORES.formula);
  fila++;

  hoja.getRange(fila, 1).setValue("Para M:");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("[M] = [q]/[p] = m/m = adimensional")
    .setFontFamily("Courier New")
    .setBackground(COLORES.formula);
  fila++;

  hoja.getRange(fila, 1).setValue("Para h':");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("[h'] = [M]*[h] = (adim)*(m) = m")
    .setFontFamily("Courier New")
    .setBackground(COLORES.formula);
  fila++;

  // Coherencia fisica
  fila++;
  hoja.getRange(fila, 1).setValue("Coherencia fisica:").setFontWeight("bold");
  fila++;

  hoja.getRange(fila, 1).setValue("1. Tipo de imagen:");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("p > f (0.10 m > 0.097 m), por tanto el objeto esta fuera del foco y se forma imagen REAL.")
    .setWrap(true);
  fila++;

  hoja.getRange(fila, 1).setValue("2. Posicion imagen:");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("q > 0 confirma que la imagen es real y se forma frente al espejo (en la pared).")
    .setWrap(true);
  fila++;

  hoja.getRange(fila, 1).setValue("3. Aumento:");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("q >> p implica |M| >> 1, consistente con lamparas reflectoras que proyectan imagenes muy aumentadas.")
    .setWrap(true);
  fila++;

  hoja.getRange(fila, 1).setValue("4. Orientacion:");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("M < 0 confirma imagen invertida, caracteristico de imagenes reales en espejos concavos.")
    .setWrap(true);
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // CUADRO RESUMEN FINAL
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 5).merge()
    .setValue("RESUMEN DE RESULTADOS - EJERCICIO 3")
    .setFontWeight("bold")
    .setFontSize(10)
    .setBackground(COLORES.header)
    .setHorizontalAlignment("center");
  fila++;

  const resumenHeader = ["Magnitud", "Simbolo", "Valor", "Unidad", "Observacion"];
  hoja.getRange(fila, 1, 1, 5).setValues([resumenHeader])
    .setFontWeight("bold")
    .setBackground(COLORES.header)
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  hoja.getRange(fila, 1).setValue("Distancia focal");
  hoja.getRange(fila, 2).setValue("f").setFontWeight("bold");
  hoja.getRange(fila, 3).setFormula("=B" + filaF + "*100").setNumberFormat("0.00").setFontWeight("bold");
  hoja.getRange(fila, 4).setValue("cm").setFontWeight("bold");
  hoja.getRange(fila, 5).setValue("f > 0: espejo concavo");
  hoja.getRange(fila, 1, 1, 5)
    .setBackground(COLORES.resultado)
    .setBorder(true, true, true, true, true, true, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  fila++;

  hoja.getRange(fila, 1).setValue("Radio de curvatura");
  hoja.getRange(fila, 2).setValue("R").setFontWeight("bold");
  hoja.getRange(fila, 3).setFormula("=B" + filaR + "*100").setNumberFormat("0.00").setFontWeight("bold");
  hoja.getRange(fila, 4).setValue("cm").setFontWeight("bold");
  hoja.getRange(fila, 5).setValue("R = 2f");
  hoja.getRange(fila, 1, 1, 5)
    .setBackground(COLORES.resultado)
    .setBorder(true, true, true, true, true, true, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  fila++;

  hoja.getRange(fila, 1).setValue("Aumento lateral");
  hoja.getRange(fila, 2).setValue("M").setFontWeight("bold");
  hoja.getRange(fila, 3).setFormula("=B" + filaM).setNumberFormat("0.0").setFontWeight("bold");
  hoja.getRange(fila, 4).setValue("(adim)").setFontWeight("bold");
  hoja.getRange(fila, 5).setValue("M<0: invertida, |M|>1: aumentada");
  hoja.getRange(fila, 1, 1, 5)
    .setBackground(COLORES.resultado)
    .setBorder(true, true, true, true, true, true, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  fila++;

  hoja.getRange(fila, 1).setValue("Altura de imagen");
  hoja.getRange(fila, 2).setValue("|h'|").setFontWeight("bold");
  hoja.getRange(fila, 3).setFormula("=ABS(B" + filaHprima + ")*100").setNumberFormat("0.0").setFontWeight("bold");
  hoja.getRange(fila, 4).setValue("cm").setFontWeight("bold");
  hoja.getRange(fila, 5).setValue("Imagen real e invertida");
  hoja.getRange(fila, 1, 1, 5)
    .setBackground(COLORES.resultado)
    .setBorder(true, true, true, true, true, true, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

// =============================================================================
// EJERCICIO 4: ESPEJO PLANO - SEMEJANZA DE TRIANGULOS
// =============================================================================

function crearEJ4() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let hoja = ss.getSheetByName("EJ4");
  if (!hoja) {
    hoja = ss.insertSheet("EJ4");
  }

  hoja.clear();
  hoja.clearFormats();

  crearEJ4_(hoja);

  ss.setActiveSheet(hoja);
}

function crearEJ4_(hoja) {
  hoja.setColumnWidths(1, 6, 160);
  hoja.setColumnWidth(1, 190);
  hoja.setColumnWidth(2, 260);

  let fila = 1;

  // ═══════════════════════════════════════════════════════════════════════════
  // FILA 1: Titulo corto en A1
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange("A1").setValue("EJ4: Espejo plano")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.titulo);
  fila = 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // FILA 2: Titulo largo academico
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange("A2:F2").merge()
    .setValue("EJERCICIO 4: ESPEJO PLANO - ALTURA DEL ARBOL POR SEMEJANZA DE TRIANGULOS")
    .setFontSize(13)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.titulo)
    .setHorizontalAlignment("center");
  fila = 4;

  // ═══════════════════════════════════════════════════════════════════════════
  // A. ENUNCIADO
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("A. ENUNCIADO")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  const enunciadoLineas = [
    "La imagen de un arbol cubre exactamente la longitud de un espejo plano de 4.00 cm de alto,",
    "cuando el espejo se sostiene a 35.0 cm del ojo. El arbol esta a 28.0 m del espejo.",
    "",
    "Cual es la altura del arbol?"
  ];

  for (let i = 0; i < enunciadoLineas.length; i++) {
    hoja.getRange(fila, 1, 1, 6).merge()
      .setValue(enunciadoLineas[i])
      .setFontSize(10)
      .setWrap(true);
    fila++;
  }
  fila++;

  // ═══════════════════════════════════════════════════════════════════════════
  // B. IDENTIFICACION DEL MODELO FISICO
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("B. IDENTIFICACION DEL MODELO FISICO")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  hoja.getRange(fila, 1).setValue("Tipo de fenomeno:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Optica geometrica - Reflexion en espejo plano con semejanza de triangulos");
  fila++;

  hoja.getRange(fila, 1).setValue("Modelo aplicado:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Formacion de imagen virtual en espejo plano y geometria de rayos");
  fila++;

  hoja.getRange(fila, 1).setValue("Fundamento teorico:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("En un espejo plano, la imagen virtual se forma a la misma distancia detras del espejo que el objeto delante. Los rayos reflejados definen triangulos semejantes con el espejo.")
    .setWrap(true);
  hoja.setRowHeight(fila, 40);
  fila++;

  hoja.getRange(fila, 1).setValue("Principio clave:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("El angulo subtendido por el espejo desde el ojo es igual al angulo subtendido por la imagen del arbol.")
    .setWrap(true)
    .setFontStyle("italic");
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // C. DATOS DEL PROBLEMA
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("C. DATOS DEL PROBLEMA")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  const headerDatos = ["Simbolo", "Magnitud fisica", "Valor dado", "Unidad", "Valor SI", "Unidad SI"];
  hoja.getRange(fila, 1, 1, 6).setValues([headerDatos])
    .setFontWeight("bold")
    .setBackground(COLORES.header)
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  const filaInicioData = fila;
  const datosProblema = [
    ["h_espejo", "Altura del espejo", "4.00", "cm", 0.04, "m"],
    ["d_ojo", "Distancia ojo-espejo", "35.0", "cm", 0.35, "m"],
    ["d_arbol", "Distancia arbol-espejo", "28.0", "m", 28.0, "m"]
  ];

  hoja.getRange(fila, 1, datosProblema.length, 6).setValues(datosProblema)
    .setBackground(COLORES.datos)
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila += datosProblema.length;

  fila++;
  hoja.getRange(fila, 1).setValue("Incognita:").setFontWeight("bold");
  fila++;
  hoja.getRange(fila, 1, 1, 6).setValues([["H", "Altura del arbol", "?", "m", "", ""]])
    .setBackground("#e8f5e9")
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila += 2;

  const celdaHespejo = "E" + filaInicioData;
  const celdaDojo = "E" + (filaInicioData + 1);
  const celdaDarbol = "E" + (filaInicioData + 2);

  // ═══════════════════════════════════════════════════════════════════════════
  // D. ECUACIONES FUNDAMENTALES
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("D. ECUACIONES FUNDAMENTALES")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  hoja.getRange(fila, 1).setValue("Distancia imagen-ojo:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 2).merge()
    .setValue("d_total = d_arbol + d_ojo")
    .setBackground(COLORES.formula)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  fila++;

  hoja.getRange(fila, 1).setValue("Semejanza triangulos:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 3).merge()
    .setValue("H / (d_arbol + d_ojo) = h_espejo / d_ojo")
    .setBackground(COLORES.formula)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  fila++;

  hoja.getRange(fila, 1).setValue("Despeje de H:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 3).merge()
    .setValue("H = h_espejo x (d_arbol + d_ojo) / d_ojo")
    .setBackground(COLORES.formula)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // E. DIAGRAMA / ESQUEMA
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("E. DIAGRAMA / ESQUEMA")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  const diagrama = [
    ["ESPEJO PLANO - SEMEJANZA DE TRIANGULOS", "", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["  ARBOL              ESPEJO              OJO", "", "", "", "", ""],
    ["    |                  |                  *", "", "", "", "", ""],
    ["    |                  |h_e             / ", "", "", "", "", ""],
    ["    | H                |              /   ", "", "", "", "", ""],
    ["    |                  |            /     ", "", "", "", "", ""],
    ["----+------------------+----------*-------", "", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["    |<-- d_arbol ----->|<--d_ojo-->|", "", "", "", "", ""],
    ["    |     (28.0 m)     | (35.0 cm) |", "", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["Triangulos semejantes:", "", "", "", "", ""],
    ["  Pequeno: base=d_ojo, altura=h_espejo", "", "", "", "", ""],
    ["  Grande:  base=d_total, altura=H", "", "", "", "", ""]
  ];

  hoja.getRange(fila, 1, diagrama.length, 6).setValues(diagrama)
    .setFontFamily("Courier New")
    .setBackground("#f8f9fa")
    .setBorder(true, true, true, true, false, false, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  hoja.getRange(fila, 1).setFontWeight("bold");
  fila += diagrama.length + 1;

  // ═══════════════════════════════════════════════════════════════════════════
  // F. DESARROLLO PASO A PASO
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("F. DESARROLLO PASO A PASO")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  // Paso 1
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 1: Calcular la distancia total imagen-ojo")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("La imagen virtual esta detras del espejo a distancia d_arbol:");
  fila++;

  hoja.getRange(fila, 1).setValue("d_total =");
  hoja.getRange(fila, 2).setValue("d_arbol + d_ojo");
  fila++;

  hoja.getRange(fila, 1).setValue("d_total =");
  hoja.getRange(fila, 2).setValue("28.0 m + 0.35 m");
  fila++;

  hoja.getRange(fila, 1).setValue("d_total =");
  hoja.getRange(fila, 2).setFormula("=" + celdaDarbol + "+" + celdaDojo)
    .setNumberFormat("0.00")
    .setBackground(COLORES.formula)
    .setFontWeight("bold");
  hoja.getRange(fila, 3).setValue("m");
  const filaDtotal = fila;
  fila++;

  // Paso 2
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 2: Aplicar semejanza de triangulos")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Por triangulos semejantes:");
  fila++;

  hoja.getRange(fila, 1, 1, 4).merge()
    .setValue("H / d_total = h_espejo / d_ojo")
    .setFontWeight("bold")
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  fila++;

  // Paso 3
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 3: Despejar H")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1, 1, 4).merge()
    .setValue("H = h_espejo x d_total / d_ojo")
    .setFontWeight("bold")
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  fila++;

  // Paso 4
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 4: Sustitucion numerica")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1).setValue("H =");
  hoja.getRange(fila, 2).setValue("(0.04 m) x (28.35 m) / (0.35 m)");
  fila++;

  hoja.getRange(fila, 1).setValue("H =");
  hoja.getRange(fila, 2).setFormula("=" + celdaHespejo + "*B" + filaDtotal + "/" + celdaDojo)
    .setNumberFormat("0.00")
    .setBackground(COLORES.formula)
    .setFontWeight("bold");
  hoja.getRange(fila, 3).setValue("m");
  const filaH = fila;
  fila += 2;

  // Resultado final
  hoja.getRange(fila, 1, 1, 4).merge()
    .setValue("RESULTADO FINAL")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORES.resultado);
  fila++;

  hoja.getRange(fila, 1).setValue("H =")
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 2).setFormula("=B" + filaH)
    .setNumberFormat("0.00")
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 3).setValue("m")
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 4).setValue("")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 1, 1, 4)
    .setBorder(true, true, true, true, false, false, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  fila++;

  hoja.getRange(fila, 1).setValue("Equivalente:")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 2).setFormula("=B" + filaH + "*100")
    .setNumberFormat("0.0")
    .setFontWeight("bold")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 3).setValue("cm")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 4).setValue("")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 1, 1, 4)
    .setBorder(true, true, true, true, false, false, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID);
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // G. COMPROBACION FISICA
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("G. COMPROBACION FISICA Y DIMENSIONAL")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  hoja.getRange(fila, 1).setValue("Analisis dimensional:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("[H] = [m] x [m] / [m] = [m]")
    .setFontFamily("Courier New")
    .setBackground(COLORES.formula);
  fila++;

  hoja.getRange(fila, 1).setValue("Orden de magnitud:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("H = 3.24 m es una altura razonable para un arbol pequeno.");
  fila++;

  hoja.getRange(fila, 1).setValue("Verificacion logica:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Factor de escala = d_total/d_ojo = 28.35/0.35 = 81. Por tanto H = 81 x 4 cm = 324 cm = 3.24 m");
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // CUADRO RESUMEN FINAL
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 4).merge()
    .setValue("RESUMEN DE RESULTADOS - EJERCICIO 4")
    .setFontWeight("bold")
    .setFontSize(10)
    .setBackground(COLORES.header)
    .setHorizontalAlignment("center");
  fila++;

  const resumenHeader = ["Magnitud", "Simbolo", "Valor", "Unidad"];
  hoja.getRange(fila, 1, 1, 4).setValues([resumenHeader])
    .setFontWeight("bold")
    .setBackground(COLORES.header)
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  hoja.getRange(fila, 1).setValue("Distancia total");
  hoja.getRange(fila, 2).setValue("d_total");
  hoja.getRange(fila, 3).setFormula("=B" + filaDtotal).setNumberFormat("0.00");
  hoja.getRange(fila, 4).setValue("m");
  hoja.getRange(fila, 1, 1, 4).setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  hoja.getRange(fila, 1).setValue("Factor de escala");
  hoja.getRange(fila, 2).setValue("d_total/d_ojo");
  hoja.getRange(fila, 3).setFormula("=B" + filaDtotal + "/" + celdaDojo).setNumberFormat("0.0");
  hoja.getRange(fila, 4).setValue("(adim)");
  hoja.getRange(fila, 1, 1, 4).setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  hoja.getRange(fila, 1).setValue("Altura del arbol").setFontWeight("bold");
  hoja.getRange(fila, 2).setValue("H").setFontWeight("bold");
  hoja.getRange(fila, 3).setFormula("=B" + filaH).setNumberFormat("0.00").setFontWeight("bold");
  hoja.getRange(fila, 4).setValue("m").setFontWeight("bold");
  hoja.getRange(fila, 1, 1, 4)
    .setBackground(COLORES.resultado)
    .setBorder(true, true, true, true, true, true, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

// =============================================================================
// EJERCICIO 5: DOBLE RENDIJA - INTERFERENCIA DE YOUNG
// =============================================================================

function crearEJ5() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let hoja = ss.getSheetByName("EJ5");
  if (!hoja) {
    hoja = ss.insertSheet("EJ5");
  }

  hoja.clear();
  hoja.clearFormats();

  crearEJ5_(hoja);

  ss.setActiveSheet(hoja);
}

function crearEJ5_(hoja) {
  hoja.setColumnWidths(1, 6, 160);
  hoja.setColumnWidth(1, 190);
  hoja.setColumnWidth(2, 260);

  let fila = 1;

  // ═══════════════════════════════════════════════════════════════════════════
  // FILA 1: Titulo corto en A1
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange("A1").setValue("EJ5: Doble rendija")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.titulo);
  fila = 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // FILA 2: Titulo largo academico
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange("A2:F2").merge()
    .setValue("EJERCICIO 5: DOBLE RENDIJA - INTERFERENCIA DE YOUNG")
    .setFontSize(13)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.titulo)
    .setHorizontalAlignment("center");
  fila = 4;

  // ═══════════════════════════════════════════════════════════════════════════
  // A. ENUNCIADO
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("A. ENUNCIADO")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  const enunciadoLineas = [
    "A traves de dos ranuras muy angostas separadas por una distancia de 0.200 mm se hace pasar",
    "luz coherente con longitud de onda de 400 nm, y el patron de interferencia se observa en una",
    "pantalla ubicada a 4.00 m de las ranuras.",
    "",
    "a) Cual es el ancho (en mm) del maximo central de interferencia?",
    "b) Cual es el ancho de la franja brillante de primer orden?"
  ];

  for (let i = 0; i < enunciadoLineas.length; i++) {
    hoja.getRange(fila, 1, 1, 6).merge()
      .setValue(enunciadoLineas[i])
      .setFontSize(10)
      .setWrap(true);
    fila++;
  }
  fila++;

  // ═══════════════════════════════════════════════════════════════════════════
  // B. IDENTIFICACION DEL MODELO FISICO
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("B. IDENTIFICACION DEL MODELO FISICO")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  hoja.getRange(fila, 1).setValue("Tipo de fenomeno:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Interferencia de luz - Experimento de doble rendija de Young");
  fila++;

  hoja.getRange(fila, 1).setValue("Modelo aplicado:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Interferencia constructiva y destructiva de ondas electromagneticas coherentes");
  fila++;

  hoja.getRange(fila, 1).setValue("Fundamento teorico:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("La luz coherente que pasa por dos rendijas produce un patron de interferencia. Los maximos (franjas brillantes) ocurren cuando la diferencia de camino optico es un multiplo entero de la longitud de onda. Se utiliza la aproximacion de angulo pequeno.")
    .setWrap(true);
  hoja.setRowHeight(fila, 50);
  fila++;

  hoja.getRange(fila, 1).setValue("Condiciones:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Rendijas muy angostas (aproximacion de fuentes puntuales), luz monocromatica y coherente, pantalla lejana (L >> d).")
    .setWrap(true)
    .setFontStyle("italic");
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // C. DATOS DEL PROBLEMA
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("C. DATOS DEL PROBLEMA")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  const headerDatos = ["Simbolo", "Magnitud fisica", "Valor dado", "Unidad", "Valor SI", "Unidad SI"];
  hoja.getRange(fila, 1, 1, 6).setValues([headerDatos])
    .setFontWeight("bold")
    .setBackground(COLORES.header)
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  const filaInicioData = fila;
  const datosProblema = [
    ["d", "Separacion entre rendijas", "0.200", "mm", 0.0002, "m"],
    ["lambda", "Longitud de onda", "400", "nm", 4e-7, "m"],
    ["L", "Distancia a la pantalla", "4.00", "m", 4.00, "m"]
  ];

  hoja.getRange(fila, 1, datosProblema.length, 6).setValues(datosProblema)
    .setBackground(COLORES.datos)
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  hoja.getRange(fila, 5, datosProblema.length, 1).setNumberFormat("0.00E+00");
  fila += datosProblema.length;

  fila++;
  hoja.getRange(fila, 1).setValue("Incognitas:").setFontWeight("bold");
  fila++;
  const incognitas = [
    ["Ancho_central", "Ancho del maximo central", "?", "mm", "", ""],
    ["Ancho_franja", "Ancho de franja brillante", "?", "mm", "", ""]
  ];
  hoja.getRange(fila, 1, incognitas.length, 6).setValues(incognitas)
    .setBackground("#e8f5e9")
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila += incognitas.length + 1;

  const celdaD = "E" + filaInicioData;
  const celdaLambda = "E" + (filaInicioData + 1);
  const celdaL = "E" + (filaInicioData + 2);

  // ═══════════════════════════════════════════════════════════════════════════
  // D. ECUACIONES FUNDAMENTALES
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("D. ECUACIONES FUNDAMENTALES")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  hoja.getRange(fila, 1).setValue("Condicion de maximos:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 2).merge()
    .setValue("d * sin(theta) = m * lambda")
    .setBackground(COLORES.formula)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  hoja.getRange(fila, 4, 1, 3).merge()
    .setValue("m = 0, +-1, +-2, ... (orden del maximo)");
  fila++;

  hoja.getRange(fila, 1).setValue("Aproximacion angulo pequeno:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 2).merge()
    .setValue("sin(theta) = tan(theta) = y / L")
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  hoja.getRange(fila, 4, 1, 3).merge()
    .setValue("Valida cuando theta << 1 rad");
  fila++;

  hoja.getRange(fila, 1).setValue("Posicion del maximo m:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 2).merge()
    .setValue("y_m = m * lambda * L / d")
    .setBackground(COLORES.formula)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  hoja.getRange(fila, 4, 1, 3).merge()
    .setValue("Distancia desde el centro al maximo m");
  fila++;

  hoja.getRange(fila, 1).setValue("Separacion entre maximos:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 2).merge()
    .setValue("Delta_y = lambda * L / d")
    .setBackground(COLORES.formula)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  hoja.getRange(fila, 4, 1, 3).merge()
    .setValue("Distancia entre franjas consecutivas");
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // E. DIAGRAMA / ESQUEMA
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("E. DIAGRAMA / ESQUEMA")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  const diagrama = [
    ["INTERFERENCIA DE DOBLE RENDIJA (YOUNG)", "", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["  Fuente      Rendijas              Pantalla", "", "", "", "", ""],
    ["              =======               ========", "", "", "", "", ""],
    ["    *  -->-->  | |                    |  m=+2", "", "", "", "", ""],
    ["   luz         | | d                  |  m=+1", "", "", "", "", ""],
    ["  coherente    | |        L           *  m=0 (central)", "", "", "", "", ""],
    ["    *  -->-->  | |  ------------->    |  m=-1", "", "", "", "", ""],
    ["               | |                    |  m=-2", "", "", "", "", ""],
    ["              =======               ========", "", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["  Maximos brillantes: d*sin(theta) = m*lambda", "", "", "", "", ""],
    ["  Ancho central: distancia entre m=-1 y m=+1 = 2*y_1", "", "", "", "", ""],
    ["  Ancho de franja: distancia entre maximos = Delta_y", "", "", "", "", ""]
  ];

  hoja.getRange(fila, 1, diagrama.length, 6).setValues(diagrama)
    .setFontFamily("Courier New")
    .setBackground("#f8f9fa")
    .setBorder(true, true, true, true, false, false, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  hoja.getRange(fila, 1).setFontWeight("bold");
  fila += diagrama.length + 1;

  // ═══════════════════════════════════════════════════════════════════════════
  // F. DESARROLLO PASO A PASO
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("F. DESARROLLO PASO A PASO")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  // === PARTE A ===
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("PARTE (a): ANCHO DEL MAXIMO CENTRAL")
    .setFontWeight("bold")
    .setFontColor(COLORES.titulo)
    .setFontSize(11)
    .setBackground("#e3f2fd");
  fila++;

  // Paso 1
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 1: Definicion del maximo central")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("El maximo central (m=0) se extiende desde el primer maximo negativo (m=-1) hasta el primer maximo positivo (m=+1).");
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("El ancho del maximo central es la distancia entre estos dos maximos:");
  fila++;

  hoja.getRange(fila, 1, 1, 3).merge()
    .setValue("Ancho_central = y_(+1) - y_(-1) = 2 * y_1")
    .setFontWeight("bold")
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  fila++;

  // Paso 2
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 2: Calcular la posicion del primer maximo (y_1)")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Usando la ecuacion de posicion del maximo con m = 1:");
  fila++;

  hoja.getRange(fila, 1, 1, 3).merge()
    .setValue("y_1 = lambda * L / d")
    .setFontWeight("bold")
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  fila++;

  // Paso 3
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 3: Sustitucion numerica para y_1")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1).setValue("lambda =");
  hoja.getRange(fila, 2).setFormula("=" + celdaLambda).setNumberFormat("0.00E+00");
  hoja.getRange(fila, 3).setValue("m");
  hoja.getRange(fila, 4).setValue("(400 nm)");
  fila++;

  hoja.getRange(fila, 1).setValue("L =");
  hoja.getRange(fila, 2).setFormula("=" + celdaL).setNumberFormat("0.00");
  hoja.getRange(fila, 3).setValue("m");
  fila++;

  hoja.getRange(fila, 1).setValue("d =");
  hoja.getRange(fila, 2).setFormula("=" + celdaD).setNumberFormat("0.0000");
  hoja.getRange(fila, 3).setValue("m");
  hoja.getRange(fila, 4).setValue("(0.200 mm)");
  fila++;

  fila++;
  hoja.getRange(fila, 1).setValue("y_1 =");
  hoja.getRange(fila, 2).setValue("(lambda * L) / d");
  fila++;

  hoja.getRange(fila, 1).setValue("y_1 =");
  hoja.getRange(fila, 2).setValue("(4.00 x 10^-7 m)(4.00 m) / (2.00 x 10^-4 m)");
  fila++;

  hoja.getRange(fila, 1).setValue("y_1 =");
  hoja.getRange(fila, 2).setFormula("=(" + celdaLambda + "*" + celdaL + ")/" + celdaD)
    .setNumberFormat("0.0000")
    .setBackground(COLORES.formula)
    .setFontWeight("bold");
  hoja.getRange(fila, 3).setValue("m");
  const filaY1 = fila;
  fila++;

  // Paso 4
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 4: Calcular el ancho del maximo central")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1).setValue("Ancho_central =");
  hoja.getRange(fila, 2).setValue("2 * y_1");
  fila++;

  hoja.getRange(fila, 1).setValue("Ancho_central =");
  hoja.getRange(fila, 2).setFormula("=2*B" + filaY1)
    .setNumberFormat("0.0000")
    .setBackground(COLORES.formula)
    .setFontWeight("bold");
  hoja.getRange(fila, 3).setValue("m");
  const filaAnchoCentral = fila;
  fila++;

  // Resultado parte a
  fila++;
  hoja.getRange(fila, 1, 1, 4).merge()
    .setValue("RESULTADO PARTE (a)")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORES.resultado);
  fila++;

  hoja.getRange(fila, 1).setValue("Ancho central =")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 2).setFormula("=B" + filaAnchoCentral + "*1000")
    .setNumberFormat("0.00")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 3).setValue("mm")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 4).setValue("")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 1, 1, 4)
    .setBorder(true, true, true, true, false, false, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  fila += 2;

  // === PARTE B ===
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("PARTE (b): ANCHO DE LA FRANJA BRILLANTE DE PRIMER ORDEN")
    .setFontWeight("bold")
    .setFontColor(COLORES.titulo)
    .setFontSize(11)
    .setBackground("#e3f2fd");
  fila++;

  // Paso 1
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 1: Definicion del ancho de franja")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("El ancho de cualquier franja brillante es la distancia entre dos maximos consecutivos:");
  fila++;

  hoja.getRange(fila, 1, 1, 3).merge()
    .setValue("Delta_y = y_(m+1) - y_m = lambda * L / d")
    .setFontWeight("bold")
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  fila++;

  // Paso 2
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 2: Observacion importante")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("La separacion entre franjas consecutivas (Delta_y) es CONSTANTE e igual a y_1.");
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Por lo tanto, el ancho de la franja de primer orden es igual a y_1:");
  fila++;

  hoja.getRange(fila, 1).setValue("Ancho_franja =");
  hoja.getRange(fila, 2).setFormula("=B" + filaY1)
    .setNumberFormat("0.0000")
    .setBackground(COLORES.formula)
    .setFontWeight("bold");
  hoja.getRange(fila, 3).setValue("m");
  const filaAnchoFranja = fila;
  fila++;

  // Resultado parte b
  fila++;
  hoja.getRange(fila, 1, 1, 4).merge()
    .setValue("RESULTADO PARTE (b)")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORES.resultado);
  fila++;

  hoja.getRange(fila, 1).setValue("Ancho franja =")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 2).setFormula("=B" + filaAnchoFranja + "*1000")
    .setNumberFormat("0.00")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 3).setValue("mm")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 4).setValue("")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 1, 1, 4)
    .setBorder(true, true, true, true, false, false, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // G. COMPROBACION FISICA
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("G. COMPROBACION FISICA Y DIMENSIONAL")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  hoja.getRange(fila, 1).setValue("Analisis dimensional:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("[y] = [lambda][L]/[d] = (m)(m)/(m) = m")
    .setFontFamily("Courier New")
    .setBackground(COLORES.formula);
  fila++;

  hoja.getRange(fila, 1).setValue("Verificacion angulo pequeno:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("theta = y/L = 0.008 m / 4 m = 0.002 rad = 0.11 grados << 1 rad (aproximacion valida)");
  fila++;

  hoja.getRange(fila, 1).setValue("Orden de magnitud:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Franjas de 8 mm son visibles y medibles en laboratorio, resultado coherente.");
  fila++;

  hoja.getRange(fila, 1).setValue("Coherencia fisica:").setFontWeight("bold");
  fila++;

  hoja.getRange(fila, 1).setValue("1. Relacion lambda/d:");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("lambda << d, condicion necesaria para patron de interferencia bien definido.")
    .setWrap(true);
  fila++;

  hoja.getRange(fila, 1).setValue("2. Franjas equiespaciadas:");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("En doble rendija, todas las franjas tienen el mismo ancho (Delta_y constante).")
    .setWrap(true);
  fila++;

  hoja.getRange(fila, 1).setValue("3. Maximo central:");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("El ancho del maximo central (16 mm) es el doble del ancho de franja (8 mm).")
    .setWrap(true);
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // CUADRO RESUMEN FINAL
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 5).merge()
    .setValue("RESUMEN DE RESULTADOS - EJERCICIO 5")
    .setFontWeight("bold")
    .setFontSize(10)
    .setBackground(COLORES.header)
    .setHorizontalAlignment("center");
  fila++;

  const resumenHeader = ["Magnitud", "Simbolo", "Valor", "Unidad", "Observacion"];
  hoja.getRange(fila, 1, 1, 5).setValues([resumenHeader])
    .setFontWeight("bold")
    .setBackground(COLORES.header)
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  hoja.getRange(fila, 1).setValue("Posicion primer maximo");
  hoja.getRange(fila, 2).setValue("y_1");
  hoja.getRange(fila, 3).setFormula("=B" + filaY1 + "*1000").setNumberFormat("0.00");
  hoja.getRange(fila, 4).setValue("mm");
  hoja.getRange(fila, 5).setValue("= lambda*L/d");
  hoja.getRange(fila, 1, 1, 5).setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  hoja.getRange(fila, 1).setValue("Ancho maximo central").setFontWeight("bold");
  hoja.getRange(fila, 2).setValue("2*y_1").setFontWeight("bold");
  hoja.getRange(fila, 3).setFormula("=B" + filaAnchoCentral + "*1000").setNumberFormat("0.00").setFontWeight("bold");
  hoja.getRange(fila, 4).setValue("mm").setFontWeight("bold");
  hoja.getRange(fila, 5).setValue("Parte (a)");
  hoja.getRange(fila, 1, 1, 5)
    .setBackground(COLORES.resultado)
    .setBorder(true, true, true, true, true, true, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  fila++;

  hoja.getRange(fila, 1).setValue("Ancho franja brillante").setFontWeight("bold");
  hoja.getRange(fila, 2).setValue("Delta_y").setFontWeight("bold");
  hoja.getRange(fila, 3).setFormula("=B" + filaAnchoFranja + "*1000").setNumberFormat("0.00").setFontWeight("bold");
  hoja.getRange(fila, 4).setValue("mm").setFontWeight("bold");
  hoja.getRange(fila, 5).setValue("Parte (b)");
  hoja.getRange(fila, 1, 1, 5)
    .setBackground(COLORES.resultado)
    .setBorder(true, true, true, true, true, true, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

// ═══════════════════════════════════════════════════════════════════════════════
// EJERCICIO 6: LENTES DELGADAS - FORMACION DE IMAGENES
// ═══════════════════════════════════════════════════════════════════════════════

function crearEJ6() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let hoja = ss.getSheetByName("EJ6");
  if (!hoja) {
    hoja = ss.insertSheet("EJ6");
  }

  hoja.clear();
  hoja.clearFormats();

  crearEJ6_(hoja);

  ss.setActiveSheet(hoja);
}

function crearEJ6_(hoja) {
  hoja.setColumnWidths(1, 6, 160);
  hoja.setColumnWidth(1, 190);
  hoja.setColumnWidth(2, 260);

  let fila = 1;

  // ═══════════════════════════════════════════════════════════════════════════
  // FILA 1: Titulo corto en A1
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange("A1").setValue("EJ6: Lentes delgadas")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.titulo);
  fila = 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // FILA 2: Titulo largo academico
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange("A2:F2").merge()
    .setValue("EJERCICIO 6: LENTES DELGADAS - FORMACION DE IMAGENES")
    .setFontSize(13)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.titulo)
    .setHorizontalAlignment("center");
  fila = 4;

  // ═══════════════════════════════════════════════════════════════════════════
  // A. ENUNCIADO
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("A. ENUNCIADO")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  const enunciadoLineas = [
    "Una lente delgada convergente tiene una distancia focal de 20.0 cm. Un objeto de 5.0 cm de altura",
    "se coloca a 30.0 cm de la lente, sobre el eje optico.",
    "",
    "a) Determine la posicion de la imagen.",
    "b) Calcule el aumento lateral y la altura de la imagen.",
    "c) Describa las caracteristicas de la imagen (real/virtual, derecha/invertida, mayor/menor)."
  ];

  for (let i = 0; i < enunciadoLineas.length; i++) {
    hoja.getRange(fila, 1, 1, 6).merge()
      .setValue(enunciadoLineas[i])
      .setFontSize(10)
      .setWrap(true);
    fila++;
  }
  fila++;

  // ═══════════════════════════════════════════════════════════════════════════
  // B. IDENTIFICACION DEL MODELO FISICO
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("B. IDENTIFICACION DEL MODELO FISICO")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  hoja.getRange(fila, 1).setValue("Tipo de fenomeno:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Refraccion de la luz - Formacion de imagenes por lentes delgadas");
  fila++;

  hoja.getRange(fila, 1).setValue("Modelo aplicado:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Optica geometrica paraxial - Ecuacion de lentes delgadas");
  fila++;

  hoja.getRange(fila, 1).setValue("Fundamento teorico:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("Las lentes delgadas refractan la luz siguiendo la ley de Snell. En la aproximacion paraxial, los rayos cercanos al eje optico convergen en puntos focales bien definidos. La ecuacion del fabricante de lentes relaciona las distancias objeto, imagen y focal.")
    .setWrap(true);
  hoja.setRowHeight(fila, 50);
  fila++;

  hoja.getRange(fila, 1).setValue("Convencion de signos:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("do > 0 (objeto real, mismo lado que luz incidente), di > 0 (imagen real, lado opuesto), di < 0 (imagen virtual, mismo lado), f > 0 (lente convergente), f < 0 (lente divergente).")
    .setWrap(true)
    .setFontStyle("italic");
  hoja.setRowHeight(fila, 40);
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // C. DATOS DEL PROBLEMA
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("C. DATOS DEL PROBLEMA")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  const headerDatos = ["Simbolo", "Magnitud fisica", "Valor dado", "Unidad", "Valor SI", "Unidad SI"];
  hoja.getRange(fila, 1, 1, 6).setValues([headerDatos])
    .setFontWeight("bold")
    .setBackground(COLORES.header)
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  const filaInicioData = fila;
  const datosProblema = [
    ["f", "Distancia focal de la lente", "20.0", "cm", 0.20, "m"],
    ["do", "Distancia objeto-lente", "30.0", "cm", 0.30, "m"],
    ["ho", "Altura del objeto", "5.0", "cm", 0.05, "m"]
  ];

  hoja.getRange(fila, 1, datosProblema.length, 6).setValues(datosProblema)
    .setBackground(COLORES.datos)
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila += datosProblema.length;

  fila++;
  hoja.getRange(fila, 1).setValue("Incognitas:").setFontWeight("bold");
  fila++;
  const incognitas = [
    ["di", "Distancia imagen-lente", "?", "cm", "", ""],
    ["M", "Aumento lateral", "?", "", "", ""],
    ["hi", "Altura de la imagen", "?", "cm", "", ""]
  ];
  hoja.getRange(fila, 1, incognitas.length, 6).setValues(incognitas)
    .setBackground("#e8f5e9")
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila += incognitas.length + 1;

  const celdaF = "E" + filaInicioData;
  const celdaDo = "E" + (filaInicioData + 1);
  const celdaHo = "E" + (filaInicioData + 2);

  // ═══════════════════════════════════════════════════════════════════════════
  // D. ECUACIONES FUNDAMENTALES
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("D. ECUACIONES FUNDAMENTALES")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  hoja.getRange(fila, 1).setValue("Ecuacion de lentes:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 2).merge()
    .setValue("1/f = 1/do + 1/di")
    .setBackground(COLORES.formula)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  hoja.getRange(fila, 4, 1, 3).merge()
    .setValue("Ecuacion fundamental de lentes delgadas");
  fila++;

  hoja.getRange(fila, 1).setValue("Despejando di:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 2).merge()
    .setValue("di = (do * f) / (do - f)")
    .setBackground(COLORES.formula)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  hoja.getRange(fila, 4, 1, 3).merge()
    .setValue("Distancia de la imagen");
  fila++;

  hoja.getRange(fila, 1).setValue("Aumento lateral:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 2).merge()
    .setValue("M = -di / do = hi / ho")
    .setBackground(COLORES.formula)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  hoja.getRange(fila, 4, 1, 3).merge()
    .setValue("Relacion de tamanos y orientacion");
  fila++;

  hoja.getRange(fila, 1).setValue("Altura de imagen:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 2).merge()
    .setValue("hi = M * ho")
    .setBackground(COLORES.formula)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  hoja.getRange(fila, 4, 1, 3).merge()
    .setValue("Altura de la imagen formada");
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // E. DIAGRAMA / ESQUEMA
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("E. DIAGRAMA / ESQUEMA")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  const diagrama = [
    ["LENTE CONVERGENTE - FORMACION DE IMAGEN", "", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["           |                    |", "", "", "", "", ""],
    ["     ho    |        Lente       |", "", "", "", "", ""],
    ["   -----   |          |         |   -----", "", "", "", "", ""],
    ["   | O |   |          |         |   | I |  hi", "", "", "", "", ""],
    ["   -----   |          |         |   -----", "", "", "", "", ""],
    ["     |     |          |         |     |", "", "", "", "", ""],
    ["  ---|-----|----F-----|----F----|-----|----> eje optico", "", "", "", "", ""],
    ["     |<--- do ------->|<------- di ------->|", "", "", "", "", ""],
    ["           |          |         |", "", "", "", "", ""],
    ["  Objeto   |    f     |    f    |  Imagen", "", "", "", "", ""],
    ["  (real)   |          |         |  (real, invertida)", "", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["  Rayos principales:", "", "", "", "", ""],
    ["  1. Paralelo al eje -> pasa por F'", "", "", "", "", ""],
    ["  2. Por el centro optico -> sin desviacion", "", "", "", "", ""],
    ["  3. Por F -> sale paralelo al eje", "", "", "", "", ""]
  ];

  hoja.getRange(fila, 1, diagrama.length, 6).setValues(diagrama)
    .setFontFamily("Courier New")
    .setBackground("#f8f9fa")
    .setBorder(true, true, true, true, false, false, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  hoja.getRange(fila, 1).setFontWeight("bold");
  fila += diagrama.length + 1;

  // ═══════════════════════════════════════════════════════════════════════════
  // F. DESARROLLO PASO A PASO
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("F. DESARROLLO PASO A PASO")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  // === PARTE A ===
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("PARTE (a): POSICION DE LA IMAGEN")
    .setFontWeight("bold")
    .setFontColor(COLORES.titulo)
    .setFontSize(11)
    .setBackground("#e3f2fd");
  fila++;

  // Paso 1
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 1: Identificar los datos en unidades consistentes")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1).setValue("f =");
  hoja.getRange(fila, 2).setValue("20.0 cm = 0.20 m");
  hoja.getRange(fila, 3).setValue("(lente convergente, f > 0)");
  fila++;

  hoja.getRange(fila, 1).setValue("do =");
  hoja.getRange(fila, 2).setValue("30.0 cm = 0.30 m");
  hoja.getRange(fila, 3).setValue("(objeto real, do > 0)");
  fila++;

  // Paso 2
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 2: Aplicar la ecuacion de lentes delgadas")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1, 1, 3).merge()
    .setValue("1/f = 1/do + 1/di")
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Despejando di:");
  fila++;

  hoja.getRange(fila, 1, 1, 3).merge()
    .setValue("1/di = 1/f - 1/do")
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  fila++;

  hoja.getRange(fila, 1, 1, 3).merge()
    .setValue("di = (do * f) / (do - f)")
    .setBackground(COLORES.formula)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  fila++;

  // Paso 3
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 3: Sustitucion numerica")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1).setValue("di =");
  hoja.getRange(fila, 2).setValue("(0.30 m)(0.20 m) / (0.30 m - 0.20 m)");
  fila++;

  hoja.getRange(fila, 1).setValue("di =");
  hoja.getRange(fila, 2).setValue("0.060 m^2 / 0.10 m");
  fila++;

  hoja.getRange(fila, 1).setValue("di =");
  hoja.getRange(fila, 2).setFormula("=(" + celdaDo + "*" + celdaF + ")/(" + celdaDo + "-" + celdaF + ")")
    .setNumberFormat("0.00")
    .setBackground(COLORES.formula)
    .setFontWeight("bold");
  hoja.getRange(fila, 3).setValue("m");
  const filaDi = fila;
  fila++;

  // Resultado parte a
  fila++;
  hoja.getRange(fila, 1, 1, 4).merge()
    .setValue("RESULTADO PARTE (a)")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORES.resultado);
  fila++;

  hoja.getRange(fila, 1).setValue("di =")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 2).setFormula("=B" + filaDi + "*100")
    .setNumberFormat("0.0")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 3).setValue("cm")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 4).setValue("(imagen real, al otro lado de la lente)")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 1, 1, 4)
    .setBorder(true, true, true, true, false, false, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  fila += 2;

  // === PARTE B ===
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("PARTE (b): AUMENTO LATERAL Y ALTURA DE LA IMAGEN")
    .setFontWeight("bold")
    .setFontColor(COLORES.titulo)
    .setFontSize(11)
    .setBackground("#e3f2fd");
  fila++;

  // Paso 1
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 1: Calcular el aumento lateral")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1, 1, 3).merge()
    .setValue("M = -di / do")
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  fila++;

  hoja.getRange(fila, 1).setValue("M =");
  hoja.getRange(fila, 2).setValue("-0.60 m / 0.30 m");
  fila++;

  hoja.getRange(fila, 1).setValue("M =");
  hoja.getRange(fila, 2).setFormula("=-B" + filaDi + "/" + celdaDo)
    .setNumberFormat("0.00")
    .setBackground(COLORES.formula)
    .setFontWeight("bold");
  hoja.getRange(fila, 3).setValue("(adimensional)");
  const filaM = fila;
  fila++;

  // Paso 2
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Paso 2: Calcular la altura de la imagen")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1, 1, 3).merge()
    .setValue("hi = M * ho")
    .setBackground(COLORES.formula)
    .setHorizontalAlignment("center");
  fila++;

  hoja.getRange(fila, 1).setValue("hi =");
  hoja.getRange(fila, 2).setValue("(-2.00)(0.05 m)");
  fila++;

  hoja.getRange(fila, 1).setValue("hi =");
  hoja.getRange(fila, 2).setFormula("=B" + filaM + "*" + celdaHo)
    .setNumberFormat("0.000")
    .setBackground(COLORES.formula)
    .setFontWeight("bold");
  hoja.getRange(fila, 3).setValue("m");
  const filaHi = fila;
  fila++;

  // Resultado parte b
  fila++;
  hoja.getRange(fila, 1, 1, 4).merge()
    .setValue("RESULTADO PARTE (b)")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORES.resultado);
  fila++;

  hoja.getRange(fila, 1).setValue("M =")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 2).setFormula("=B" + filaM)
    .setNumberFormat("0.00")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 3).setValue("")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 4).setValue("(aumento 2x, signo negativo = invertida)")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 1, 1, 4)
    .setBorder(true, true, true, true, false, false, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  fila++;

  hoja.getRange(fila, 1).setValue("hi =")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 2).setFormula("=ABS(B" + filaHi + ")*100")
    .setNumberFormat("0.0")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 3).setValue("cm")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 4).setValue("(altura absoluta)")
    .setBackground(COLORES.resultado);
  hoja.getRange(fila, 1, 1, 4)
    .setBorder(true, true, true, true, false, false, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  fila += 2;

  // === PARTE C ===
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("PARTE (c): CARACTERISTICAS DE LA IMAGEN")
    .setFontWeight("bold")
    .setFontColor(COLORES.titulo)
    .setFontSize(11)
    .setBackground("#e3f2fd");
  fila++;

  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("Analisis de los resultados obtenidos:")
    .setFontWeight("bold")
    .setFontColor(COLORES.subseccion);
  fila++;

  hoja.getRange(fila, 1).setValue("1. Real o Virtual:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("di > 0, por lo tanto la imagen es REAL (se forma al lado opuesto de la lente).");
  fila++;

  hoja.getRange(fila, 1).setValue("2. Orientacion:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("M < 0, por lo tanto la imagen esta INVERTIDA respecto al objeto.");
  fila++;

  hoja.getRange(fila, 1).setValue("3. Tamano relativo:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("|M| = 2 > 1, por lo tanto la imagen es MAYOR (magnificada) que el objeto.");
  fila++;

  // Resultado parte c
  fila++;
  hoja.getRange(fila, 1, 1, 5).merge()
    .setValue("RESULTADO PARTE (c): CARACTERISTICAS")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORES.resultado);
  fila++;

  const caracteristicas = [
    ["Caracteristica", "Valor", "Criterio", "Conclusion", ""],
    ["Naturaleza", "di = +60 cm", "di > 0", "IMAGEN REAL", ""],
    ["Orientacion", "M = -2.00", "M < 0", "INVERTIDA", ""],
    ["Tamano", "|M| = 2.00", "|M| > 1", "MAGNIFICADA (2x)", ""]
  ];

  hoja.getRange(fila, 1, caracteristicas.length, 5).setValues(caracteristicas)
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  hoja.getRange(fila, 1, 1, 5)
    .setFontWeight("bold")
    .setBackground(COLORES.header);
  hoja.getRange(fila + 1, 1, caracteristicas.length - 1, 5)
    .setBackground(COLORES.resultado);
  fila += caracteristicas.length + 1;

  // ═══════════════════════════════════════════════════════════════════════════
  // G. COMPROBACION FISICA
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 6).merge()
    .setValue("G. COMPROBACION FISICA Y DIMENSIONAL")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor(COLORES.blanco)
    .setBackground(COLORES.seccion);
  fila++;

  hoja.getRange(fila, 1).setValue("Analisis dimensional:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("[di] = [do][f]/[do-f] = (m)(m)/(m) = m   (correcto)")
    .setFontFamily("Courier New")
    .setBackground(COLORES.formula);
  fila++;

  hoja.getRange(fila, 1).setValue("Verificacion ecuacion:").setFontWeight("bold");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("1/f = 1/do + 1/di => 1/0.20 = 1/0.30 + 1/0.60 => 5.0 = 3.33 + 1.67 = 5.0 (verificado)");
  fila++;

  hoja.getRange(fila, 1).setValue("Coherencia fisica:").setFontWeight("bold");
  fila++;

  hoja.getRange(fila, 1).setValue("1. Posicion objeto:");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("do = 30 cm > f = 20 cm (objeto fuera del foco), imagen real esperada.")
    .setWrap(true);
  fila++;

  hoja.getRange(fila, 1).setValue("2. Caso especifico:");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("do = 1.5f produce di = 3f y M = -2, imagen real invertida y magnificada.")
    .setWrap(true);
  fila++;

  hoja.getRange(fila, 1).setValue("3. Conservacion energia:");
  hoja.getRange(fila, 2, 1, 5).merge()
    .setValue("La intensidad de la imagen disminuye al ser magnificada (misma luz en area mayor).")
    .setWrap(true);
  fila += 2;

  // ═══════════════════════════════════════════════════════════════════════════
  // CUADRO RESUMEN FINAL
  // ═══════════════════════════════════════════════════════════════════════════
  hoja.getRange(fila, 1, 1, 5).merge()
    .setValue("RESUMEN DE RESULTADOS - EJERCICIO 6")
    .setFontWeight("bold")
    .setFontSize(10)
    .setBackground(COLORES.header)
    .setHorizontalAlignment("center");
  fila++;

  const resumenHeader = ["Magnitud", "Simbolo", "Valor", "Unidad", "Observacion"];
  hoja.getRange(fila, 1, 1, 5).setValues([resumenHeader])
    .setFontWeight("bold")
    .setBackground(COLORES.header)
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  hoja.getRange(fila, 1).setValue("Distancia focal");
  hoja.getRange(fila, 2).setValue("f");
  hoja.getRange(fila, 3).setValue("20.0");
  hoja.getRange(fila, 4).setValue("cm");
  hoja.getRange(fila, 5).setValue("Dato (convergente)");
  hoja.getRange(fila, 1, 1, 5).setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  hoja.getRange(fila, 1).setValue("Distancia objeto");
  hoja.getRange(fila, 2).setValue("do");
  hoja.getRange(fila, 3).setValue("30.0");
  hoja.getRange(fila, 4).setValue("cm");
  hoja.getRange(fila, 5).setValue("Dato");
  hoja.getRange(fila, 1, 1, 5).setBorder(true, true, true, true, true, true, COLORES.borde, SpreadsheetApp.BorderStyle.SOLID);
  fila++;

  hoja.getRange(fila, 1).setValue("Distancia imagen").setFontWeight("bold");
  hoja.getRange(fila, 2).setValue("di").setFontWeight("bold");
  hoja.getRange(fila, 3).setFormula("=B" + filaDi + "*100").setNumberFormat("0.0").setFontWeight("bold");
  hoja.getRange(fila, 4).setValue("cm").setFontWeight("bold");
  hoja.getRange(fila, 5).setValue("Parte (a)");
  hoja.getRange(fila, 1, 1, 5)
    .setBackground(COLORES.resultado)
    .setBorder(true, true, true, true, true, true, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  fila++;

  hoja.getRange(fila, 1).setValue("Aumento lateral").setFontWeight("bold");
  hoja.getRange(fila, 2).setValue("M").setFontWeight("bold");
  hoja.getRange(fila, 3).setFormula("=B" + filaM).setNumberFormat("0.00").setFontWeight("bold");
  hoja.getRange(fila, 4).setValue("").setFontWeight("bold");
  hoja.getRange(fila, 5).setValue("Parte (b)");
  hoja.getRange(fila, 1, 1, 5)
    .setBackground(COLORES.resultado)
    .setBorder(true, true, true, true, true, true, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  fila++;

  hoja.getRange(fila, 1).setValue("Altura imagen").setFontWeight("bold");
  hoja.getRange(fila, 2).setValue("hi").setFontWeight("bold");
  hoja.getRange(fila, 3).setFormula("=ABS(B" + filaHi + ")*100").setNumberFormat("0.0").setFontWeight("bold");
  hoja.getRange(fila, 4).setValue("cm").setFontWeight("bold");
  hoja.getRange(fila, 5).setValue("Parte (b)");
  hoja.getRange(fila, 1, 1, 5)
    .setBackground(COLORES.resultado)
    .setBorder(true, true, true, true, true, true, "#2e7d32", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}
