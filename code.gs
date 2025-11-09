/**
 * Code.gs COMPLETO - Sistema de Censo con Reglas de Ahorro Energético
 * Integra reglas de iluminación y aire acondicionado + dimensionamiento BTU
 */

/* ------------------------
   UI / Helpers básicos
   ------------------------ */
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Censo de Cargas")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Conversión segura a número (admite coma decimal) */
function toNumber(value, defaultVal = 0) {
  if (value === undefined || value === null) return defaultVal;
  const num = parseFloat(value.toString().replace(",", "."));
  return isNaN(num) ? defaultVal : num;
}

/** Número o BLANCO: si está vacío, retorna "" (deja celda vacía) */
function toNumberOrBlank(value) {
  if (value === undefined || value === null) return "";
  const s = value.toString().trim();
  if (s === "") return "";
  const num = parseFloat(s.replace(",", "."));
  return isNaN(num) ? "" : num;
}

/* ------------------------
   Creación de spreadsheet
   ------------------------ */
function createNewSpreadsheet_(edificio) {
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const name = `Censo – ${edificio || "SinNombre"} – ${ts}`;
  const ss = SpreadsheetApp.create(name);
  let sheet = ss.getSheets()[0];
  sheet.setName("Respuestas");

  sheet.appendRow([
    "Fecha",                         // A
    "Edificio",                      // B
    "Tipo de edificio",              // C
    "Pisos",                         // D
    "Ubicación",                     // E
    "Área (m²)",                     // F
    "Set point (°C)",                // G
    "Controlador (Sí/No)",           // H
    "Carga",                         // I
    "Cantidad",                      // J
    "Descripción",                   // K
    "Demanda por unidad (kW)",       // L
    "Demanda (kW)",                  // M = L*J
    "Capacidad",                     // N
    "Unidad capacidad",              // O
    "Horas uso",                     // P
    "Días uso al mes",               // Q
    "Consumo diario (kWh)",          // R = M*P
    "Consumo mensual (kWh)",         // S = R*Q
    "Concurrencia",                  // T
    "Condición operativa"            // U
  ]);

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 21)
       .setBackground("#1E88E5")
       .setFontColor("#FFFFFF")
       .setFontWeight("bold");
  return { ss, sheet };
}

/* ------------------------
   Procesamiento del formulario
   ------------------------ */
function processForm(data) {
  const edificio = data.edificio || "";
  const tipo = data.tipo || "";
  const pisos = data.pisos || "";
  const cargas = data.cargas || [];

  const { ss, sheet } = createNewSpreadsheet_(edificio);

  cargas.forEach(c => {
    const r = sheet.getLastRow() + 1;

    const row = [
      new Date(),                        // A
      edificio,                          // B
      tipo,                              // C
      pisos,                             // D
      (c.ubicacion || "").trim(),        // E
      toNumber(c.areaM2),                // F
      toNumberOrBlank(c.setPoint),       // G
      (c.controlador || "").trim(),      // H
      (c.carga || "").trim(),            // I
      toNumber(c.cantidad, 1),           // J
      (c.descripcion || "").trim(),      // K
      toNumber(c.demanda),               // L
      "",                                // M (fórmula)
      (c.capacidad || "").trim(),        // N
      (c.unidadCapacidad || "").trim(),  // O
      toNumber(c.horasUso),              // P
      toNumber(c.diasMes),               // Q
      "",                                // R (fórmula)
      "",                                // S (fórmula)
      (c.concurrencia || "").trim(),     // T
      (c.condicionOperativa || "").trim()// U
    ];

    sheet.getRange(r, 1, 1, row.length).setValues([row]);

    sheet.getRange(r, 13).setFormula(`=IFERROR(L${r}*J${r},0)`); // M
    sheet.getRange(r, 18).setFormula(`=IFERROR(M${r}*P${r},0)`); // R
    sheet.getRange(r, 19).setFormula(`=IFERROR(R${r}*Q${r},0)`); // S
  });

  // Fila de TOTALES
  const lastRow = sheet.getLastRow();
  const totalsRow = lastRow + 1;
  sheet.getRange(totalsRow, 1, 1, 21).clearContent();
  sheet.getRange(totalsRow, 10).setValue("TOTALES →");
  sheet.getRange(totalsRow, 13).setFormula(`=SUM(M2:M${lastRow})`);
  sheet.getRange(totalsRow, 18).setFormula(`=SUM(R2:R${lastRow})`);
  sheet.getRange(totalsRow, 19).setFormula(`=SUM(S2:S${lastRow})`);

  applyFormatting_(sheet, totalsRow);

  // Procesamiento OpenAI
  try {
    processSheetWithOpenAI(ss.getId());
  } catch (e) {
    Logger.log('OpenAI processing skipped: ' + e.message);
  }

  // Email
  const email = "Kilowattiareportes@gmail.com";
  const subject = `Nuevo Censo de Cargas: ${edificio}`;
  const body = `Se ha recibido un nuevo censo para el edificio: ${edificio}\n\n` +
               `Archivo: ${ss.getName()}\nURL: ${ss.getUrl()}\n\n` +
               `Tipo: ${tipo}  |  Pisos: ${pisos}\n\nSaludos.`;
  MailApp.sendEmail(email, subject, body);

  return { url: ss.getUrl(), fileName: ss.getName(), count: cargas.length, id: ss.getId() };
}

/* ------------------------
   Formato y presentación
   ------------------------ */
function applyFormatting_(sheet, totalsRow) {
  const lastDataRow = totalsRow - 1;
  const lastCol = 24;
  if (lastDataRow < 1) return;

  sheet.getRange(2, 13, Math.max(lastDataRow - 1, 0), 1).setBackground("#FFF8DC");
  sheet.getRange(2, 18, Math.max(lastDataRow - 1, 0), 2).setBackground("#FFF8DC");

  sheet.getRange(totalsRow, 1, 1, lastCol)
       .setBackground("#E8F5E9")
       .setFontWeight("bold");

  sheet.getRange(2, 5,  lastDataRow - 1, 1).setWrap(true);
  sheet.getRange(2, 8,  lastDataRow - 1, 1).setWrap(true);
  sheet.getRange(2, 9,  lastDataRow - 1, 1).setWrap(true);
  sheet.getRange(2, 11, lastDataRow - 1, 1).setWrap(true);
  sheet.getRange(2, 15, lastDataRow - 1, 1).setWrap(true);
  sheet.getRange(2, 20, lastDataRow - 1, 1).setWrap(true);
  sheet.getRange(2, 21, lastDataRow - 1, 1).setWrap(true);

  sheet.getRange(2, 1,  lastDataRow - 1, 1).setNumberFormat("yyyy-mm-dd hh:mm");
  sheet.getRange(2, 6,  lastDataRow - 1, 1).setNumberFormat("#,##0.##");
  sheet.getRange(2, 7,  lastDataRow - 1, 1).setNumberFormat("#,##0.##");
  sheet.getRange(2, 10, lastDataRow - 1, 1).setNumberFormat("#,##0");
  sheet.getRange(2, 12, lastDataRow - 1, 1).setNumberFormat("#,##0.00");
  sheet.getRange(2, 13, lastDataRow - 1, 1).setNumberFormat("#,##0.00");
  sheet.getRange(2, 16, lastDataRow - 1, 2).setNumberFormat("#,##0.##");
  sheet.getRange(2, 18, lastDataRow - 1, 2).setNumberFormat("#,##0.00");

  sheet.getRange(2, 1,  lastDataRow - 1, 1).setHorizontalAlignment("left");
  sheet.getRange(2, 5,  lastDataRow - 1, 6).setHorizontalAlignment("left");
  sheet.getRange(2, 6,  lastDataRow - 1, 2).setHorizontalAlignment("right");
  sheet.getRange(2, 10, lastDataRow - 1, 1).setHorizontalAlignment("right");
  sheet.getRange(2, 12, lastDataRow - 1, 7).setHorizontalAlignment("right");
  sheet.getRange(2, 20, lastDataRow - 1, 2).setHorizontalAlignment("left");

  sheet.getRange(1, 1, totalsRow, lastCol)
       .setBorder(true, true, true, true, true, true, "#D7DFE8", SpreadsheetApp.BorderStyle.SOLID);
  sheet.autoResizeColumns(1, lastCol);
}

/* ------------------------
   Integración OpenAI
   ------------------------ */
function processSheetWithOpenAI(spreadsheetId) {
  const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!OPENAI_API_KEY) throw new Error('Falta OPENAI_API_KEY en Script Properties.');

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName('Respuestas');
  if (!sheet) throw new Error('Hoja "Respuestas" no encontrada.');

  const HEADER_ROW = 1;
  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) {
    SpreadsheetApp.getUi().alert('No hay filas de datos para procesar.');
    return;
  }

  const MAX_ROWS_TO_PROCESS = 50;
  const totalDataRows = lastRow - HEADER_ROW;
  const rowsToProcess = (MAX_ROWS_TO_PROCESS && totalDataRows > MAX_ROWS_TO_PROCESS) 
    ? MAX_ROWS_TO_PROCESS : totalDataRows;

  const dataRange = sheet.getRange(HEADER_ROW + 1, 1, rowsToProcess, 21);
  const rows = dataRange.getValues();

  // Columnas de salida V(22), W(23), X(24)
  const startColOut = 22;
  const outColsCount = 3;
  const desiredHdrs = ['% de ahorro alcanzable', 'kWh ahorrado', 'RAZON_MODEL'];
  
  const hdrs = sheet.getRange(1, startColOut, 1, outColsCount).getValues()[0];
  let needHeaders = hdrs.some(h => !h || h.toString().trim() === '');
  if (needHeaders) {
    sheet.getRange(1, startColOut, 1, desiredHdrs.length).setValues([desiredHdrs]);
    sheet.getRange(1, startColOut, 1, desiredHdrs.length)
         .setBackground("#1E88E5").setFontColor("#FFFFFF").setFontWeight("bold");
  }

  const outputs = [];

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const filaIndex = HEADER_ROW + 1 + i;

    // Extracción de datos (índices 0-based)
    const ubicacion = (r[4] || "").toString().trim().toLowerCase();     // E
    const area_m2 = toNumber(r[5], 0);                                  // F
    const setPoint = toNumberOrBlank(r[6]);                             // G
    const controlador = (r[7] || "").toString().trim().toLowerCase();   // H
    const carga = (r[8] || "").toString().trim().toLowerCase();         // I
    const cantidad = toNumber(r[9], 1);                                 // J
    const descripcion = (r[10] || "").toString().trim().toLowerCase();  // K
    const demanda_kW_unidad = toNumber(r[11], 0);                       // L
    const capacidad_unitaria = toNumber(r[13], 0);                      // N
    const unidad_capacidad = (r[14] || "").toString().trim().toLowerCase(); // O
    const horasUso = toNumber(r[15], 0);                                // P
    const kwhMensual = toNumber(r[18], 0);                              // S
    const concurrencia = (r[19] || "").toString().trim().toLowerCase(); // T
    const condicionOperativa = (r[20] || "").toString().trim().toLowerCase(); // U

    // Objeto completo para el modelo
    const modelInput = {
      fila: filaIndex,
      ubicacion: ubicacion,
      area_m2: area_m2,
      set_point: setPoint,
      controlador: controlador,
      carga: carga,
      cantidad: cantidad,
      descripcion: descripcion,
      capacidad_unitaria: capacidad_unitaria,
      unidad_capacidad: unidad_capacidad,
      demanda_kW_unidad: demanda_kW_unidad,
      horas_uso: horasUso,
      kwh_mensual: kwhMensual,
      concurrencia: concurrencia,
      condicion_operativa: condicionOperativa
    };

    // Llamada al modelo
    const resp = callOpenAI_forRow(modelInput, OPENAI_API_KEY);

    // Normalizar salida
    let pctAchievable = resp.pct_achievable ?? "";
    if (pctAchievable !== "" && !isNaN(Number(pctAchievable))) {
      pctAchievable = Math.max(0, Math.min(100, Number(pctAchievable)));
    } else {
      pctAchievable = "";
    }

    let kwhSaved = resp.kwh_saved ?? "";
    if ((kwhSaved === "" || isNaN(kwhSaved)) && kwhMensual && pctAchievable !== "") {
      kwhSaved = Number(kwhMensual) * (Number(pctAchievable) / 100);
    }
    if (!isFiniteNumber(kwhSaved)) kwhSaved = "";

    let razon = resp.reason || "";

    outputs.push([
      (pctAchievable !== "" ? roundTo(pctAchievable, 3) : ""),
      (kwhSaved !== "" ? roundTo(kwhSaved, 3) : ""),
      razon
    ]);

    Utilities.sleep(200);
  }

  if (outputs.length > 0) {
    sheet.getRange(HEADER_ROW + 1, startColOut, outputs.length, outputs[0].length).setValues(outputs);
  }

  // Totales
  const totalsRow = lastRow + 1;
  const colPct = columnLetter(startColOut);
  const colKwh = columnLetter(startColOut + 1);
  sheet.getRange(totalsRow, startColOut).setFormula(`=IFERROR(SUM(${colPct}2:${colPct}${lastRow}),"")`);
  sheet.getRange(totalsRow, startColOut + 1).setFormula(`=IFERROR(SUM(${colKwh}2:${colKwh}${lastRow}),"")`);
  sheet.getRange(totalsRow, startColOut, 1, 2).setFontWeight("bold");

  sheet.autoResizeColumns(startColOut, outColsCount);
  SpreadsheetApp.getUi().alert('Proceso OpenAI completado. Filas procesadas: ' + outputs.length +
                              (rowsToProcess < totalDataRows ? ` (solo primeros ${rowsToProcess})` : ''));
}

/**
 * Llamada a OpenAI con TODAS las reglas de negocio integradas
 */
function callOpenAI_forRow(modelInput, OPENAI_API_KEY) {
  const OPENAI_API_URL = 'https://api.openai.com/v1/chat/completions';

  const system = `Eres un ingeniero especializado en eficiencia energética. Recibirás datos de un censo de cargas eléctricas y debes calcular el porcentaje de ahorro alcanzable aplicando las siguientes reglas EXACTAS:

═══════════════════════════════════════════════════════════════
📌 REGLAS PARA ILUMINACIÓN (cuando carga contiene 'lumin', 'luz', 'bombill', 'led')
═══════════════════════════════════════════════════════════════

1. **55% de ahorro**:
   - CON o SIN controlador
   - Y concurrencia = "no concurrido"
   - Y ubicacion contiene: "vestibulo" O "estacionamiento" O "escalera" O "cuarto de maquinas"

2. **25% de ahorro**:
   - SIN controlador (controlador = "no" o vacío)
   - Y concurrencia = "no concurrido"
   - Y cantidad > 15 luminarias
   - Y NO aplica regla #1 (no es vestíbulo/estacionamiento/escalera/cuarto máquinas)

3. **Recomendación LED**:
   - Si descripcion contiene "bombillo incandescente" o "incandescente"
   - Agrega en reason: "Se recomienda cambiar a tecnología LED por baja eficiencia"

═══════════════════════════════════════════════════════════════
📌 REGLAS PARA AIRE ACONDICIONADO (cuando carga contiene 'aire', 'ac', 'hvac', 'climatiza')
═══════════════════════════════════════════════════════════════

**Cálculos base (SIEMPRE hacer):**
- BTU_requeridos = area_m2 × 800 (clima Panamá)
- Si unidad_capacidad contiene 'ton' o 'tr': capacidad_btu = capacidad_unitaria × 12000
- Si unidad_capacidad contiene 'btu': capacidad_btu = capacidad_unitaria
- total_capacity_btu = capacidad_btu × cantidad
- diff_btu = total_capacity_btu - BTU_requeridos

**Estado dimensionamiento:**
- Si diff_btu < -500: "subdimensionado"
- Si diff_btu > 500: "sobredimensionado"
- Si |diff_btu| ≤ 500: "correctamente dimensionado"

**Porcentajes de ahorro (ACUMULABLES donde se indique):**

1. **5% por cada grado bajo 23°C** (ACUMULABLE):
   - Si set_point < 23: ahorro_grados = (23 - set_point) × 5%
   - Ejemplo: set_point=20 → (23-20)×5 = 15%

2. **22% base** (INDEPENDIENTE, no acumulable con #3):
   - SIN controlador (controlador = "no" o vacío)
   - Y horas_uso > 8
   - Y capacidad < 24000 BTU (o < 2 toneladas)

3. **15% base** (INDEPENDIENTE, no acumulable con #2):
   - ES concurrido (concurrencia = "concurrido" o "sí")
   - Y capacidad < 24000 BTU (o < 2 toneladas)

4. **4% extra por mantenimiento** (ACUMULABLE):
   - Si condicion_operativa = "no favorable"
   - Se SUMA a cualquier otro porcentaje

**Regla final de cálculo:**
- pct_total = MAX(regla#2, regla#3) + ahorro_grados(#1) + mantenimiento(#4)

═══════════════════════════════════════════════════════════════
📤 SALIDA REQUERIDA (JSON)
═══════════════════════════════════════════════════════════════

Devuelve SÓLO un objeto JSON válido con:
{
  "tipo_carga": "iluminacion" | "aire_acondicionado" | "otro",
  "pct_achievable": número (0-100, suma total de ahorros),
  "kwh_saved": número (kWh/mes estimado ahorrado),
  "sizing_status": "subdimensionado" | "sobredimensionado" | "correctamente dimensionado" | null,
  "btu_required": número | null,
  "total_btu": número | null,
  "diff_btu": número | null,
  "reason": "Explicación clara de cálculo aplicado (máx 200 caracteres)",
  "confidence": número 0-1
}

IMPORTANTE:
- Si no puedes calcular algo, usa null
- reason debe ser conciso y explicar qué reglas aplicaste
- NO inventes datos, usa exactamente lo que recibes`;

  const userPrompt = `Analiza esta carga: ${JSON.stringify(modelInput)}

Devuelve SÓLO JSON válido.`;

  const payload = {
    model: "gpt-4o",
    messages: [
      { role: "system", content: system },
      { role: "user", content: userPrompt }
    ],
    temperature: 0.1,
    max_tokens: 500
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + OPENAI_API_KEY },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const resp = UrlFetchApp.fetch(OPENAI_API_URL, options);
    const code = resp.getResponseCode();
    const txt = resp.getContentText();

    if (code < 200 || code >= 300) {
      return { reason: `HTTP ${code}: ${txt.substring(0,300)}` };
    }

    let json = {};
    try { 
      json = JSON.parse(txt); 
    } catch (e) { 
      return { reason: 'Error parsing API response' };
    }

    // Extraer respuesta del modelo
    let assistantText = '';
    if (json.choices && json.choices[0] && json.choices[0].message) {
      assistantText = json.choices[0].message.content;
    } else {
      return { reason: 'Formato de respuesta inesperado' };
    }

    // Parsear JSON de la respuesta
    let parsed = null;
    try {
      const first = assistantText.indexOf('{');
      const last = assistantText.lastIndexOf('}');
      if (first !== -1 && last !== -1 && last > first) {
        const candidate = assistantText.slice(first, last + 1);
        parsed = JSON.parse(candidate);
      } else {
        parsed = JSON.parse(assistantText);
      }
    } catch (e) {
      return { reason: assistantText.substring(0,500) };
    }

    return {
      tipo_carga: parsed.tipo_carga ?? null,
      pct_achievable: parsed.pct_achievable ?? null,
      kwh_saved: parsed.kwh_saved ?? null,
      sizing_status: parsed.sizing_status ?? null,
      btu_required: parsed.btu_required ?? null,
      total_btu: parsed.total_btu ?? null,
      diff_btu: parsed.diff_btu ?? null,
      reason: parsed.reason ?? '',
      confidence: parsed.confidence ?? null
    };
  } catch (err) {
    return { reason: 'Error de conexión: ' + err.toString() };
  }
}

/* ------------------------
   Utilidades
   ------------------------ */
function isFiniteNumber(v) {
  return typeof v === 'number' && isFinite(v);
}

function roundTo(n, d) { 
  if (!isFiniteNumber(n)) return ""; 
  const p = Math.pow(10, d || 0); 
  return Math.round((n + Number.EPSILON) * p) / p; 
}

function columnLetter(col) {
  let temp, letter = '';
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}
