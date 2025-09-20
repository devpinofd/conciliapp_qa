/**
 * @fileoverview Lógica para el particionamiento de datos en Google Sheets.
 * Gestiona la creación y el acceso a hojas de cálculo particionadas por fecha,
 * vendedor o banco para mejorar el rendimiento y la escalabilidad.
 */

/**
 * Decide el tipo de partición a utilizar basado en el registro.
 * Esta es una implementación de ejemplo; la lógica real puede ser más compleja.
 * @param {Object} record El registro de datos.
 * @returns {string} El tipo de partición ('mes', 'vendedor', 'banco').
 */
function decidePartitionType(record) {
  const hasVendedor = !!record.vendedorCodigo;
  const hasBanco = !!record.bancoReceptor;

  if (hasVendedor && hasBanco) {
    return 'hibrido';
  }
  if (hasVendedor) {
    return 'vendedor';
  }
  if (hasBanco) {
    return 'banco';
  }
  return 'mes';
}

/**
 * Devuelve el nombre de la partición según la estrategia y los datos proporcionados.
 * @param {Date} date La fecha para la partición.
 * @param {Object} opts Opciones que incluyen el tipo y datos adicionales.
 * @param {string} opts.type Estrategia de partición: 'mes', 'vendedor', 'banco', 'hibrido'.
 * @param {string} [opts.vendedor] Código del vendedor.
 * @param {string} [opts.banco] Nombre del banco receptor.
 * @returns {string} El nombre de la hoja de partición.
 * @throws {Error} Si el tipo de partición es desconocido.
 */
function getPartitionName(date, { type, vendedor, banco }) {
  const yyyy = date.getFullYear();
  const monthNames = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
  const mm = monthNames[date.getMonth()];

  switch (type) {
    case 'mes':
      return `REG_${yyyy}_${mm}`;
    case 'vendedor':
      if (!vendedor) throw new Error('El código de vendedor es requerido para la partición por vendedor.');
      return `V_${vendedor}_${yyyy}_${mm}`;
    case 'banco':
      if (!banco) throw new Error('El banco es requerido para la partición por banco.');
      return `B_${banco}_${yyyy}_${mm}`;
    case 'hibrido':
      if (!vendedor || !banco) throw new Error('Vendedor y banco son requeridos para la partición híbrida.');
      return `V_${vendedor}_B_${banco}_${yyyy}_${mm}`;
    default:
      throw new Error(`Tipo de partición desconocido: ${type}`);
  }
}

/**
 * Asegura que una hoja de partición exista. Si no existe, la crea con su encabezado.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss El Spreadsheet activo.
 * @param {string} name El nombre de la hoja a asegurar.
 * @param {string[]} header El array de strings para el encabezado.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} La hoja de cálculo (existente o nueva).
 */
function ensurePartitionSheet(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (header && header.length > 0) {
      sh.getRange(1, 1, 1, header.length).setValues([header]);
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

/**
 * Función para la rotación mensual de la hoja de respuestas.
 * Crea o asegura la existencia de la hoja de cálculo para el mes actual y el siguiente.
 * Esta función está diseñada para ser ejecutada por un trigger de tiempo.
 */
function rotacionMensual_() {
  const ss = SpreadsheetApp.openById(SheetManager.SPREADSHEET_ID);
  const now = new Date();
  const header = SheetManager.SHEET_CONFIG['Respuestas'].headers;

  // Asegurar la partición del mes actual
  const currentMonthName = getPartitionName(now, {type:'mes'});
  ensurePartitionSheet(ss, currentMonthName, header);

  // Pre-crear la partición del mes siguiente para evitar cualquier demora el primer día
  const nextMonth = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const nextMonthName = getPartitionName(nextMonth, {type:'mes'});
  ensurePartitionSheet(ss, nextMonthName, header);
  
  Logger.log(`Ejecución de rotación mensual. Particiones aseguradas: ${currentMonthName}, ${nextMonthName}`);
}
