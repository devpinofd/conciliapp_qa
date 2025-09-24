/**
 * @fileoverview Lógica del servidor para la aplicación de cobranza.
 * Refactorizado con principios SOLID y mejores prácticas.
 * Incluye soporte para facturas múltiples (CSV) en una sola fila.
 */

// #region Autenticación (Funciones añadidas)
const TOKEN_EXPIRATION_SECONDS = 6 * 3600; // 6 horas

function checkAuth(token) {
  if (!token) return null;
  const userCache = CacheService.getUserCache();
  const cached = userCache.get(token);
  if (cached) {
    const user = JSON.parse(cached);
    // Renovar el token en caché para extender la sesión
    userCache.put(token, JSON.stringify(user), TOKEN_EXPIRATION_SECONDS);
    return user;
  }
  return null;
}

function logoutUser(token) {
  if (!token) return;
  try {
    const userCache = CacheService.getUserCache();
    userCache.remove(token);
    Logger.log('Cierre de sesión exitoso para el token.');
    return { status: 'success', message: 'Sesión cerrada.' };
  } catch (e) {
    Logger.error(`Error en logoutUser: ${e.message}`);
    throw new Error('No se pudo cerrar la sesión en el servidor.');
  }
}
// #endregion

/** Limpieza controlada de propiedades del script */
function clearScriptProperties() {
  const ultimoIndiceAConservar = 6;
  const propiedades = PropertiesService.getScriptProperties();
  const todasLasClaves = propiedades.getKeys();
  todasLasClaves.sort();
  Logger.log('Iniciando limpieza de propiedades desde el cliente.');
  todasLasClaves.forEach((clave, indice) => {
    if (indice > ultimoIndiceAConservar) {
      propiedades.deleteProperty(clave);
      Logger.log(`Se eliminó la propiedad en el índice ${indice} (clave: "${clave}").`);
    }
  });
  Logger.log(`Proceso de limpieza completado. Se conservaron las primeras ${ultimoIndiceAConservar + 1} propiedades.`);
}

// #region Utilidades
class Logger {
  static log(message, ...args) {
    console.log(message, ...args);
    this.appendLog('INFO', message);
  }
  static error(message, ...args) {
    console.error(message, ...args);
    this.appendLog('ERROR', message);
  }
  static appendLog(level, message) {
    const sheet = SheetManager.getSheet('Auditoria');
    sheet.appendRow([new Date(), Session.getActiveUser().getEmail(), level, message]);
  }
}

function normalizeEmail(email) {
  if (!email || typeof email !== 'string') {
    return '';
  }
  return email.trim().toLowerCase();
}

class CacheManager {
  static get(key, ttlSeconds, fetchFunction, ...args) {
    const cache = PropertiesService.getScriptProperties();
    const cached = cache.getProperty(key);
    if (cached) {
      const { timestamp, data } = JSON.parse(cached);
      if (new Date().getTime() - timestamp < ttlSeconds * 1000) {
        return data;
      }
    }
    const data = fetchFunction(...args);
    cache.setProperty(key, JSON.stringify({ timestamp: new Date().getTime(), data }));
    return data;
  }
}

class ApiHandler {
  constructor() {
    const props = PropertiesService.getScriptProperties();
    this.API_URL = props.getProperty('API_URL') || 'https://login.factorysoftve.com/api/generica/efactoryApiGenerica.asmx/Seleccionar';
    this.headers = {
      apikey: props.getProperty('API_KEY'),
      usuario: props.getProperty('API_USER'),
      empresa: props.getProperty('API_EMPRESA')
    };
    if (!this.headers.apikey || !this.headers.usuario || !this.headers.empresa) {
      Logger.error('Faltan credenciales de la API en las propiedades del script.');
      throw new Error('Faltan credenciales de la API.');
    }
  }
  fetchData(query) {
    const options = {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      headers: this.headers,
      payload: JSON.stringify({ lcResultado: 'json2', lcConsulta: query }),
      muteHttpExceptions: false,
      validateHttpsCertificates: true
    };
    try {
      const response = UrlFetchApp.fetch(this.API_URL, options);
      if (response.getResponseCode() !== 200) {
        throw new Error(`Error HTTP: ${response.getResponseCode()} - ${response.getContentText()}`);
      }
      const jsonResponse = JSON.parse(response.getContentText());
      if (!jsonResponse.d || !jsonResponse.d.laTablas || jsonResponse.d.laTablas.length === 0) {
        return [];
      }
      return jsonResponse.d.laTablas[0];
    } catch (e) {
      Logger.error(`Error al llamar a la API: ${e.message}`, { query });
      throw e;
    }
  }
}
// #endregion

// #region Gestores de datos
class SheetManager {
  static getSheet(sheetName) {
    const ss = SpreadsheetApp.openById(this.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet && this.SHEET_CONFIG[sheetName]?.headers) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, this.SHEET_CONFIG[sheetName].headers.length)
        .setValues([this.SHEET_CONFIG[sheetName].headers]);
    }
    if (!sheet) throw new Error(`Hoja ${sheetName} no encontrada y no se pudo crear.`);
    return sheet;
  }
}
SheetManager.SPREADSHEET_ID = '1jv3jlYSRfCj9VHPBla0g1l35AFEwNJzkWwrOML5oPY4';
SheetManager.SHEET_CONFIG = {
  'CorreosPermitidos': { headers: null },
  'obtenerVendedoresPorUsuario': { headers: ['correo', 'vendedorcompleto', 'codvendedor','Sucursal'] },
  'Administradores': { headers: ['correo_admin'] },
  'Bancos': { headers: ['Nombre del Banco'] },
  'analista': { headers: ['sucursal', 'codigousuario', 'email'] },
  'AsignacionOverrides': { headers: ['codvendedor', 'analista_email'] },
  'Respuestas': {
    headers: ['Timestamp', 'Vendedor', 'Codigo Cliente', 'Nombre Cliente', 'Factura',
      'Monto Pagado', 'Forma de Pago', 'Banco Emisor', 'Banco Receptor',
      'Nro. de Referencia', 'Tipo de Cobro', 'Fecha de la Transferencia o Pago',
      'Observaciones', 'Usuario Creador', 'EstadoAnalista', 'ComentarioAnalista', 'AnalistaAsignado','Sucursal','id_registro', 'FechaReconciliacion']
  },
  'Auditoria': { headers: ['Timestamp', 'Usuario', 'Nivel', 'Detalle'] },
  'Auditoria_Analistas': { headers: ['Timestamp', 'Analista', 'ID Registro', 'Estado Anterior', 'Estado Nuevo', 'Comentario'] },
  'RegistrosBancoEFactory': { headers: ['Timestamp', 'sucursal', 'banco_receptor', 'referencia', 'cobro', 'codvendedor', 'vendedor', 'factura', 'monto', 'comentario', 'Nombre'] },
  'Registros Eliminados': {
    headers: ['Fecha Eliminación', 'Usuario que Eliminó', 'Timestamp', 'Vendedor',
      'Codigo Cliente', 'Nombre Cliente', 'Factura', 'Monto Pagado',
      'Forma de Pago', 'Banco Emisor', 'Banco Receptor', 'Nro. de Referencia',
      'Tipo de Cobro', 'Fecha de la Transferencia o Pago', 'Observaciones', 'Usuario Creador', 'EstadoAnalista', 'ComentarioAnalista', 'AnalistaAsignado','Sucursal']
  },
  'Usuarios': {
    headers: ['Correo', 'Contraseña', 'Estado', 'Nombre', 'Fecha Registro']
  }
};

class DataFetcher {
  constructor() { this.api = new ApiHandler(); }
  fetchVendedoresFromSheetByUser(userEmail) {
    if (!userEmail) {
      Logger.error('Se intentó llamar a fetchVendedoresFromSheetByUser sin un email.');
      return [];
    }
    const normalizedUserEmail = normalizeEmail(userEmail);
    const sheet = SheetManager.getSheet('obtenerVendedoresPorUsuario');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const vendedoresFiltrados = data
      .map(row => ({
        email: normalizeEmail(row[0]),
        nombre: String(row[1]).trim(),
        codigo: String(row[2]).trim(),
        sucursal: String(row[3]).trim()
      }))
      .filter(v => v.email === normalizedUserEmail && v.nombre && v.codigo && v.sucursal);
    if (vendedoresFiltrados.length === 0) {
      Logger.log(`No se encontraron vendedores para el usuario: ${userEmail}`);
    }
    return vendedoresFiltrados;
  }
  fetchAllVendedoresFromSheet() {
    const sheet = SheetManager.getSheet('obtenerVendedoresPorUsuario');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    return data.map(row => ({
      nombre: String(row[1]).trim(),
      codigo: String(row[2]).trim(),
      sucursal: String(row[3]).trim()
    })).filter(v => v.nombre && v.codigo && v.sucursal);
  }
  isUserAdmin(userEmail) {
    if (!userEmail) return false;
    const normalizedUserEmail = normalizeEmail(userEmail);
    const sheet = SheetManager.getSheet('Administradores');
    if (sheet.getLastRow() < 2) return false;
    const adminEmails = sheet.getRange("A2:A" + sheet.getLastRow()).getValues()
      .flat().map(email => normalizeEmail(email));
    return adminEmails.includes(normalizedUserEmail);
  }
    fetchClientesFromApi(codVendedor) {
    // --- INICIO DE CAMBIOS PARA DEPURACIÓN ---
    Logger.log(`Iniciando fetchClientesFromApi con codVendedor: '${codVendedor}' (Tipo: ${typeof codVendedor})`);
    // --- FIN DE CAMBIOS PARA DEPURACIÓN ---

    if (!codVendedor || typeof codVendedor !== 'string') {
      Logger.error(`Código de vendedor inválido: ${codVendedor}`);
      return [];
    }
    const safeCodVendedor = codVendedor.replace(/['"]/g, '');
    const props = PropertiesService.getScriptProperties();
    const queryTemplate = props.getProperty('CLIENTES_QUERY');
    if (!queryTemplate) {
      Logger.error('La propiedad CLIENTES_QUERY no está definida.');
      throw new Error('No se encontró la consulta para cargar clientes.');
    }
    const query = queryTemplate.replace(/{codVendedor}/g, safeCodVendedor);
    
    // --- INICIO DE CAMBIOS PARA DEPURACIÓN ---
//    Logger.log(`Ejecutando consulta de clientes: ${query}`);
    // --- FIN DE CAMBIOS PARA DEPURACIÓN ---

    try {
      const data = this.api.fetchData(query);
      
      // --- INICIO DE CAMBIOS PARA DEPURACIÓN ---
  //    Logger.log(`API devolvió ${data.length} clientes para el vendedor ${codVendedor}.`);
      // --- FIN DE CAMBIOS PARA DEPURACIÓN ---

      return data.map(row => ({
        nombre: String(row.Nombre).trim(),
        codigo: String(row.Cod_Cliente).trim()
      }));
    } catch (e) {
      Logger.error(`Error en fetchClientesFromApi: ${e.message}`, { query });
      return [];
    }
  }
   fetchFacturasFromApi(codVendedor, codCliente) {
    // --- INICIO DE CAMBIOS PARA DEPURACIÓN DE FACTURAS ---
 //   Logger.log(`Iniciando fetchFacturasFromApi con codVendedor: '${codVendedor}' (Tipo: ${typeof codVendedor}), codCliente: '${codCliente}' (Tipo: ${typeof codCliente})`);
    // --- FIN DE CAMBIOS PARA DEPURACIÓN DE FACTURAS ---

    if (!codVendedor || !codCliente) {
      Logger.error(`Parámetros inválidos: codVendedor=${codVendedor}, codCliente=${codCliente}`);
      return [];
    }
    const safeCodVendedor = codVendedor.replace(/['"]/g, '');
    const safeCodCliente = codCliente.replace(/['"]/g, '');
    const props = PropertiesService.getScriptProperties();
    const queryTemplate = props.getProperty('FACTURAS_QUERY');
    if (!queryTemplate) {
      Logger.error('La propiedad FACTURAS_QUERY no está definida.');
      throw new Error('No se encontró la consulta para cargar facturas.');
    }
    const query = queryTemplate
      .replace(/{safeCodCliente}/g, safeCodCliente)
      .replace(/{safeCodVendedor}/g, safeCodVendedor);

    // --- INICIO DE CAMBIOS PARA DEPURACIÓN DE FACTURAS ---
 //   Logger.log(`Ejecutando consulta de facturas: ${query}`);
    // --- FIN DE CAMBIOS PARA DEPURACIÓN DE FACTURAS ---

    try {
      const data = this.api.fetchData(query);

      // --- INICIO DE CAMBIOS PARA DEPURACIÓN DE FACTURAS ---
  //    Logger.log(`API devolvió ${data.length} facturas para el vendedor ${codVendedor} y cliente ${codCliente}.`);
      // --- FIN DE CAMBIOS PARA DEPURACIÓN DE FACTURAS ---

      return data.map(row => ({
        documento: String(row.documento).trim(),
        mon_sal: parseFloat(row.mon_sal) || 0,
        fec_ini: row.fec_ini ? new Date(row.fec_ini).toISOString().split('T')[0] : '',
        cod_mon: String(row.cod_mon).trim() || 'USD'
      }));
    } catch (e) {
      Logger.error(`Error en fetchFacturasFromApi: ${e.message}`, { query });
      return [];
    }
  }
  fetchBcvRate() {
    const apiUrl = 'https://ve.dolarapi.com/v1/dolares/oficial';
    try {
      const response = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
      if (response.getResponseCode() !== 200) {
        Logger.error(`Error en fetchBcvRate: Código ${response.getResponseCode()}`);
        return 1;
      }
      const jsonResponse = JSON.parse(response.getContentText());
      const rate = parseFloat(jsonResponse.promedio);
      if (isNaN(rate) || rate <= 0) {
        Logger.error('Tasa BCV inválida');
        return 1;
      }
      return rate;
    } catch (e) {
      Logger.error(`Error en fetchBcvRate: ${e.message}`);
      return 1;
    }
  }
  fetchBancosFromSheet() {
    const sheet = SheetManager.getSheet('Bancos');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('No se encontraron bancos en la hoja.');
      return [];
    }
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    return data.map(row => ({
      nombre: String(row[0]).trim(),
      codigo: String(row[0]).trim()
    })).filter(b => b.nombre && b.codigo);
  }
  fetchSucursalesPorAnalistaFromApi() {
    const props = PropertiesService.getScriptProperties();
    const query = props.getProperty('SUCURSALES_USUARIOS_QUERY');
    if (!query) {
      Logger.error('La propiedad SUCURSALES_USUARIOS_QUERY no está definida.');
      throw new Error('No se encontró la consulta para cargar sucursales por usuario.');
    }
    try {
      const data = this.api.fetchData(query);
      return data.map(row => ({
        sucursal: String(row.sucursal || '').trim(),
        codigousuario: String(row.codigousuario || '').trim()
      }));
    } catch (e) {
      Logger.error(`Error en fetchSucursalesPorAnalistaFromApi: ${e.message}`, { query });
      return [];
    }
  }

  fetchRegistrosBancoFromApi(startDate, endDate) {
    const props = PropertiesService.getScriptProperties();
    const queryTemplate = props.getProperty('BANCOS_REGISTROS_QUERY');
    if (!queryTemplate) {
      Logger.error('La propiedad BANCOS_REGISTROS_QUERY no está definida.');
      throw new Error('No se encontró la consulta para cargar los registros del banco.');
    }
    const query = queryTemplate
      .replace(/{startDate}/g, startDate)
      .replace(/{endDate}/g, endDate);
    try {
      const data = this.api.fetchData(query);
      return data.map(row => ({
        sucursal: String(row.sucursal || '').trim(),
        banco_receptor: String(row.banco_receptor || '').trim(),
        referencia: String(row.referencia || '').trim(),
        cobro: String(row.cobro || '').trim(),
        codvendedor: String(row.codvendedor || '').trim(),
        vendedor: String(row.vendedor || '').trim(),
        factura: String(row.factura || '').trim(),
        monto: parseFloat(row.monto) || 0,
        comentario: String(row.comentario || '').trim(),
        Nombre: String(row.Nombre || '').trim()
      }));
    } catch (e) {
      Logger.error(`Error en fetchRegistrosBancoFromApi: ${e.message}`, { query });
      return [];
    }
  }
}

class CobranzaService {
  constructor(dataFetcher) {
    this.dataFetcher = dataFetcher;
    this.REGISTROS_POR_PAGINA = 10;
  }

  static get scriptUserEmail() {
    return Session.getActiveUser().getEmail();
  }

  /**
   * Normaliza una cadena CSV de facturas:
   * - split por coma
   * - trim de cada elemento
   * - filtra elementos vacíos
   * - join con coma sin espacios
   * @param {string} value
   * @return {string}
   */
  _normalizeFacturaCsv(value) {
    if (!value) return '';
    return value
      .split(',')
      .map(s => (s || '').trim())
      .filter(s => s.length > 0)
      .join(',');
  }

  getVendedores(userEmail, forceRefresh = false) {
    if (!userEmail) throw new Error("No se pudo obtener el email del usuario para cargar los vendedores.");
    const isAdmin = this.dataFetcher.isUserAdmin(userEmail);
    const cacheKey = `vendedores_html_${isAdmin ? 'admin' : userEmail}`;
    const fetchFunction = () => {
      let vendedores = isAdmin
        ? this.dataFetcher.fetchAllVendedoresFromSheet()
        : this.dataFetcher.fetchVendedoresFromSheetByUser(userEmail);
      if (vendedores.length === 0) {
        throw new Error(`No tiene vendedores asignados. Por favor, contacte al administrador.`);
      }
      let optionsHtml = isAdmin ? '<option value="Mostrar todos">Mostrar todos</option>' : '';
      optionsHtml += vendedores.map(v => `<option value="${v.codigo}">${v.nombre} (${v.sucursal})</option>`).join('');
      return optionsHtml;
    };
    if (forceRefresh) {
      const data = fetchFunction();
      PropertiesService.getScriptProperties()
        .setProperty(cacheKey, JSON.stringify({ timestamp: new Date().getTime(), data }));
      return data;
    }
    return CacheManager.get(cacheKey, 21600, fetchFunction);
  }

  getClientesHtml(codVendedor) {
    const clientes = CacheManager.get(`clientes_${codVendedor}`, 21600,
      () => this.dataFetcher.fetchClientesFromApi(codVendedor));
    return clientes.map(c => `<option value="${c.codigo}">${c.nombre}</option>`).join('');
  }

  getFacturas(codVendedor, codCliente) {
    return CacheManager.get(`facturas_${codVendedor}_${codCliente}`, 21600,
      () => this.dataFetcher.fetchFacturasFromApi(codVendedor, codCliente));
  }

  getBcvRate() {
    return CacheManager.get('bcv_rate', 21600, () => this.dataFetcher.fetchBcvRate());
  }

  getBancos() {
    return CacheManager.get('bancos', 86400, () => this.dataFetcher.fetchBancosFromSheet());
  }

  submitData(data, userEmail) {
    const ss = SpreadsheetApp.openById(SheetManager.SPREADSHEET_ID);
 // Normalización y validaciones mínimas
    const facturaCsvRaw = data.factura || data.documento || '';
    const facturaCsv = this._normalizeFacturaCsv(facturaCsvRaw);
    if (!facturaCsv) throw new Error('Debe indicar al menos una factura.');
    
    const montoNum = parseFloat(data.montoPagado);
    if (isNaN(montoNum) || montoNum <= 0) throw new Error('Monto inválido.');
    if (!data.vendedor) throw new Error('Vendedor requerido.');
    if (!data.cliente) throw new Error('Código de cliente requerido.');


 // Lógica de particionamiento

    const submissionDate = new Date(); // Fecha para determinar la partición (siempre la actual)
    const record = {
        vendedorCodigo: data.vendedor,
        bancoReceptor: data.bancoReceptor,
    };
    const partitionOpts = {
        type: decidePartitionType(record),
        vendedor: record.vendedorCodigo,
        banco: record.bancoReceptor
    };
    const partitionName = getPartitionName(submissionDate, partitionOpts);
    const header = SheetManager.SHEET_CONFIG['Respuestas'].headers;
    const partitionSheet = ensurePartitionSheet(ss, partitionName, header);

   // Validación de duplicidad de referencia (ahora en la partición correcta)
    let existingReferences = [];
    if (partitionSheet.getLastRow() > 1) {
      existingReferences = partitionSheet
        .getRange(2, 10, partitionSheet.getLastRow() - 1, 1)
        .getValues()
        .flat();
    }
    if (existingReferences.includes(data.nroReferencia)) {
      throw new Error('El número de referencia ya existe en esta partición.');
    }
    const facturaArray = facturaCsv.split(',');
    const todosLosVendedores = this.dataFetcher.fetchAllVendedoresFromSheet();
    const vendedorEncontrado = todosLosVendedores.find(v => v.codigo === data.vendedor);
    const nombreCompletoVendedor = vendedorEncontrado ? vendedorEncontrado.nombre : data.vendedor;
    const sucursal = vendedorEncontrado ? vendedorEncontrado.sucursal : '';
 // --- INICIO DE LA MODIFICACIÓN ---
    // Generar un ID único para el registro
    const id_registro = new Date().getTime().toString(36) + Math.random().toString(36).substring(2, 9);
    // --- FIN DE LA MODIFICACIÓN ---
    const row = [
      submissionDate,
      nombreCompletoVendedor,
      data.cliente,
      data.nombreCliente,
      facturaCsv,
      montoNum,
      data.formaPago,
      data.bancoEmisor,
      data.bancoReceptor,
      data.nroReferencia,
      data.tipoCobro,
      data.fechaTransferenciaPago,
      data.observaciones,
      userEmail,
      data.estadoAnalista || '',
      data.comentarioAnalista || '',
      data.analistaAsignado || '',
      sucursal,
      id_registro
    ];

    partitionSheet.appendRow(row);
    Logger.log(`Formulario enviado por ${userEmail} a la partición ${partitionName}. Facturas: ${facturaCsv}`);

    // Notificar a la vista de analista que hay una actualización
    const scriptCache = CacheService.getScriptCache();
    const newTimestamp = new Date().getTime().toString();
    scriptCache.put('lastUpdateTimestamp', newTimestamp, 21600); // Cache por 6 horas

    return '¡Datos recibidos con éxito!';
  }

  getRecentRecords(vendedor, userEmail) {
    const ss = SpreadsheetApp.openById(SheetManager.SPREADSHEET_ID);
    const allSheets = ss.getSheets();
    const partitionRegex = /^(V_.+|B_.+|REG_|V_.+_B_.+)_\d{4}_(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)$/;
    const monthOrder = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];

    const partitionSheets = allSheets
      .map(s => s.getName())
      .filter(name => partitionRegex.test(name))
      .sort((a, b) => {
        const [, , yearA, monthA] = a.match(partitionRegex);
        const [, , yearB, monthB] = b.match(partitionRegex);
        if (yearA !== yearB) return yearB.localeCompare(yearA);
        return monthOrder.indexOf(monthB) - monthOrder.indexOf(monthA);
      });

    let allRecords = [];
    const RECORDS_TO_FETCH = this.REGISTROS_POR_PAGINA * 5;

    for (const sheetName of partitionSheets) {
      if (allRecords.length >= RECORDS_TO_FETCH) break;
      const sheet = ss.getSheetByName(sheetName);
      if (sheet.getLastRow() <= 1) continue;
      
      const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      const recordsWithMeta = values.map((row, index) => ({
        data: row,
        sheetName: sheetName,
        rowIndex: index + 2 // El índice real de la fila en su hoja
      }));
      allRecords.push(...recordsWithMeta);
    }
    
    // Ordenar todos los registros por timestamp descendente
    allRecords.sort((a, b) => new Date(b.data[0]).getTime() - new Date(a.data[0]).getTime());

    const isAdmin = this.dataFetcher.isUserAdmin(userEmail);
    let filteredRecords;

    if (isAdmin) {
      if (vendedor && vendedor !== 'Mostrar todos') {
        const todosLosVendedores = this.dataFetcher.fetchAllVendedoresFromSheet();
        const vendedorSeleccionado = todosLosVendedores.find(v => v.codigo === vendedor);
        const nombreVendedorFiltro = vendedorSeleccionado ? vendedorSeleccionado.nombre : null;
        filteredRecords = allRecords.filter(r => r.data[1] === nombreVendedorFiltro);
      } else {
        filteredRecords = allRecords;
      }
    } else {
      const misVendedores = this.dataFetcher.fetchVendedoresFromSheetByUser(userEmail).map(v => v.nombre);
      filteredRecords = allRecords.filter(r => misVendedores.includes(r.data[1]));
    }

    const finalRecords = filteredRecords.slice(0, this.REGISTROS_POR_PAGINA);
    const now = new Date().getTime();
    const FIVE_MINUTES_IN_MS = 5 * 60 * 1000;

    return finalRecords.map(record => {
      const row = record.data;
      const timestamp = new Date(row[0]).getTime();
      const ageInMs = now - timestamp;
      const puedeEliminarPorTiempo = ageInMs < FIVE_MINUTES_IN_MS;
      
      return {
        rowIndex: JSON.stringify({ sheet: record.sheetName, row: record.rowIndex }),
        fechaEnvio: Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
        vendedor: row[1],
        clienteNombre: row[3],
        factura: row[4],
        monto: (typeof row[5] === 'number') ? row[5].toFixed(2) : row[5],
        bancoEmisor: row[7],
        bancoReceptor: row[8],
        referencia: row[9],
        creadoPor: row[13],
        puedeEliminar: (row[13] === userEmail && puedeEliminarPorTiempo)
      };
    });
  }

  deleteRecord(rowIndex, userEmail) {
    const { sheet: sheetName, row: rowNum } = JSON.parse(rowIndex);
    if (!sheetName || !rowNum) {
      throw new Error('Información de registro inválida para eliminación.');
    }
    
    const sheet = SpreadsheetApp.openById(SheetManager.SPREADSHEET_ID).getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`La hoja de partición '${sheetName}' no fue encontrada.`);
    }

    const rowToDelete = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (rowToDelete[13] !== userEmail) {
      throw new Error('No tienes permiso para eliminar este registro.');
    }
    
    const timestamp = new Date(rowToDelete[0]).getTime();
    const now = new Date().getTime();
    const ageInMs = now - timestamp;
    const FIVE_MINUTES_IN_MS = 5 * 60 * 1000;
    if (ageInMs > FIVE_MINUTES_IN_MS) {
      throw new Error('No se puede eliminar un registro después de 5 minutos de su creación.');
    }
    
    const auditSheet = SheetManager.getSheet('Registros Eliminados');
    auditSheet.appendRow([new Date(), userEmail, ...rowToDelete]);
    sheet.deleteRow(rowNum);
    
    Logger.log(`Registro eliminado por ${userEmail}. Fila: ${rowNum} en hoja: ${sheetName}`);
    return 'Registro eliminado y archivado con éxito.';
  }
}

// Reportes PDF


class ReportService {
  constructor(dataFetcher) { this.dataFetcher = dataFetcher; }
  
  getRecordsInDateRange(userEmail, vendedorFiltro, start, end) {
    const ss = SpreadsheetApp.openById(SheetManager.SPREADSHEET_ID);
    const allSheets = ss.getSheets();
    const partitionRegex = /^(V_.+|B_.+|REG_|V_.+_B_.+)_\d{4}_(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)$/;

    const partitionSheets = allSheets
      .map(s => s.getName())
      .filter(name => partitionRegex.test(name));

    let allRecords = [];
    for (const sheetName of partitionSheets) {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet.getLastRow() <= 1) continue;
      const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      const recordsInRange = values.filter(row => {
        const ts = new Date(row[0]).getTime();
        return ts >= start.getTime() && ts <= end.getTime();
      });
      allRecords.push(...recordsInRange);
    }
    
    const isAdmin = this.dataFetcher.isUserAdmin(userEmail);
    let filtered = allRecords;

    if (isAdmin) {
      if (vendedorFiltro && vendedorFiltro !== 'Mostrar todos') {
        const todos = this.dataFetcher.fetchAllVendedoresFromSheet();
        const ven = todos.find(v => v.codigo === vendedorFiltro);
        const nombreFiltro = ven ? ven.nombre : null;
        if (nombreFiltro) filtered = filtered.filter(row => row[1] === nombreFiltro);
      }
    } else {
      const misVendedores = this.dataFetcher.fetchVendedoresFromSheetByUser(userEmail).map(v => v.nombre);
      filtered = filtered.filter(row => misVendedores.includes(row[1]));
    }

    const tz = Session.getScriptTimeZone();
    return filtered.map(row => ({
      // Fecha de creación/timestamp (col A) para la primera columna del reporte
      fecha: Utilities.formatDate(new Date(row[0]), tz, 'dd/MM/yyyy HH:mm'),
      vendedor: String(row[1] ?? ''),
      clienteCodigo: String(row[2] ?? ''),
      clienteNombre: String(row[3] ?? ''),
      factura: String(row[4] ?? ''),
      monto: (typeof row[5] === 'number') ? row[5].toFixed(2) : String(row[5] ?? ''),
      formaPago: String(row[6] ?? ''),
      bancoEmisor: String(row[7] ?? ''),
      bancoReceptor: String(row[8] ?? ''),
      referencia: String(row[9] ?? ''),
      tipoCobro: String(row[10] ?? ''),
      // Fecha del pago (col L)
      fechaPago: row[11] ? Utilities.formatDate(new Date(row[11]), tz, 'dd/MM/yyyy') : '',
      // FORZAR A STRING PARA EVITAR .trim is not a function
      observaciones: String(row[12] ?? ''),
      creadoPor: String(row[13] ?? ''),
    }));
  }

  buildPdf(records, meta) {
    const template = HtmlService.createTemplateFromFile('Report');
    template.records = records;
    template.meta = meta;
    const html = template.evaluate().getContent();
    const blob = Utilities.newBlob(html, 'text/html', 'reporte.html').getAs(MimeType.PDF);
    blob.setName(meta.filename);
    return blob;
  }
}


// #endregion

// #region Lógica de Analistas

/**
 * Obtiene un mapa de los analistas disponibles y las sucursales que tienen asignadas.
 * @returns {Map<string, string[]>} Un mapa donde la clave es el email del analista y el valor es un array de sus sucursales.
 */
function getAvailableAnalysts() {
  try {
    const sheet = SheetManager.getSheet('analista');
    if (sheet.getLastRow() < 2) {
      Logger.log('No se encontraron analistas configurados en la hoja "analista".');
      return new Map();
    }
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    const analystMap = new Map();

    data.forEach(row => {
      const sucursal = String(row[0] || '').trim();
      const email = normalizeEmail(row[2]); // Usar la tercera columna (índice 2)

      if (email && sucursal) {
        if (!analystMap.has(email)) {
          analystMap.set(email, []);
        }
        analystMap.get(email).push(sucursal);
      }
    });
    Logger.log(`Analistas cargados: ${analystMap.size}`);
    return analystMap;
  } catch (e) {
    Logger.error(`Error en getAvailableAnalysts: ${e.message}`);
    return new Map();
  }
}

/**
 * Asigna registros pendientes a los analistas disponibles.
 * Sigue una lógica de "fair queue" por fecha, sucursal y vendedor.
 */
function assignPendingRecords() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('assignPendingRecords ya está en ejecución. Omitiendo esta llamada.');
    return;
  }

  try {
    Logger.log('Iniciando proceso de asignación de 2 etapas (Overrides + Reparto Justo)...');
    const allAnalystData = getAvailableAnalysts();
    if (allAnalystData.size === 0) {
      Logger.log('No hay analistas disponibles para asignar registros.');
      return;
    }

    // 1. Obtener reglas de override y crear un mapa de consulta rápida
    const overrideSheet = SheetManager.getSheet('AsignacionOverrides');
    const overrideData = overrideSheet.getLastRow() > 1 ? overrideSheet.getRange(2, 1, overrideSheet.getLastRow() - 1, 2).getValues() : [];
    const overrideMap = new Map(overrideData.map(([sellerCode, analystEmail]) => [String(sellerCode).trim(), normalizeEmail(analystEmail)]));
    Logger.log(`${overrideMap.size} reglas de override cargadas.`);

    // 2. Crear mapa inverso (Vendedor -> Codigo) para buscar overrides
    const dataFetcher = new DataFetcher();
    const allSellers = dataFetcher.fetchAllVendedoresFromSheet();
    const sellerNameToCodeMap = new Map(allSellers.map(seller => [seller.nombre, seller.codigo]));

    // 3. Recolectar todos los registros pendientes
    const ss = SpreadsheetApp.openById(SheetManager.SPREADSHEET_ID);
    const allSheets = ss.getSheets();
    const partitionRegex = /^(V_.+|B_.+|REG_|V_.+_B_.+)_\d{4}_(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)$/;
    const headers = SheetManager.SHEET_CONFIG['Respuestas'].headers;
    const analystColIndex = headers.indexOf('AnalistaAsignado');
    const statusColIndex = headers.indexOf('EstadoAnalista');
    const sucursalColIndex = headers.indexOf('Sucursal');
    const vendedorColIndex = headers.indexOf('Vendedor');

    let allPendingRecords = [];
    allSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (!partitionRegex.test(sheetName) || sheet.getLastRow() <= 1) return;
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      data.forEach((row, index) => {
        if (!row[analystColIndex] && !row[statusColIndex]) {
          const sellerName = String(row[vendedorColIndex] || '').trim();
          allPendingRecords.push({ 
              sheetName, 
              rowIndex: index + 2, 
              sucursal: String(row[sucursalColIndex] || '').trim(),
              sellerName: sellerName,
              sellerCode: sellerNameToCodeMap.get(sellerName) || null
          });
        }
      });
    });

    if (allPendingRecords.length === 0) {
      Logger.log('No hay registros pendientes para asignar.');
      return;
    }

    // 4. Separar registros: los que tienen override y los que van a reparto justo
    const recordsForFairShare = [];
    const directAssignments = [];
    allPendingRecords.forEach(record => {
        if (record.sellerCode && overrideMap.has(record.sellerCode)) {
            const designatedAnalyst = overrideMap.get(record.sellerCode);
            directAssignments.push({ ...record, assignedAnalyst: designatedAnalyst });
        } else {
            recordsForFairShare.push(record);
        }
    });

    const updates = new Map();

    // 5. Procesar asignaciones directas (override)
    if (directAssignments.length > 0) {
        Logger.log(`Procesando ${directAssignments.length} asignaciones directas por override.`);
        directAssignments.forEach(record => {
            if (!updates.has(record.sheetName)) updates.set(record.sheetName, []);
            updates.get(record.sheetName).push({ rowIndex: record.rowIndex, analyst: record.assignedAnalyst });
        });
    }

    // 6. Procesar el resto con el reparto justo por sucursal
    if (recordsForFairShare.length > 0) {
        Logger.log(`Procesando ${recordsForFairShare.length} registros por reparto justo.`);
        const recordsByBranch = recordsForFairShare.reduce((acc, record) => {
            const branch = record.sucursal || 'SIN_SUCURSAL';
            if (!acc[branch]) acc[branch] = [];
            acc[branch].push(record);
            return acc;
        }, {});

        const analystsWithAllAccess = Array.from(allAnalystData.entries()).filter(([, sucursales]) => sucursales.includes('TODAS')).map(([email]) => email);
        const analystsByBranch = {};
        allAnalystData.forEach((sucursales, email) => {
            sucursales.forEach(sucursal => {
                if (sucursal !== 'TODAS') {
                    if (!analystsByBranch[sucursal]) analystsByBranch[sucursal] = [];
                    analystsByBranch[sucursal].push(email);
                }
            });
        });

        const scriptCache = CacheService.getScriptCache();
        for (const branch in recordsByBranch) {
            const recordsToAssign = recordsByBranch[branch];
            const specificAnalysts = analystsByBranch[branch] || [];
            const qualifiedAnalysts = [...new Set([...specificAnalysts, ...analystsWithAllAccess])];

            if (qualifiedAnalysts.length === 0) {
                Logger.log(`No se encontraron analistas para la sucursal '${branch}' en reparto justo. ${recordsToAssign.length} registros serán omitidos.`);
                continue;
            }

            let lastIndexCacheKey = `last_analyst_index_${branch}`;
            let lastAssignedIndex = parseInt(scriptCache.get(lastIndexCacheKey) || '-1', 10);

            recordsToAssign.forEach(record => {
                const nextAnalystIndex = (lastAssignedIndex + 1) % qualifiedAnalysts.length;
                const assignedAnalyst = qualifiedAnalysts[nextAnalystIndex];
                if (!updates.has(record.sheetName)) updates.set(record.sheetName, []);
                updates.get(record.sheetName).push({ rowIndex: record.rowIndex, analyst: assignedAnalyst });
                lastAssignedIndex = nextAnalystIndex;
            });
            scriptCache.put(lastIndexCacheKey, lastAssignedIndex.toString(), 21600);
        }
    }

    // 7. Aplicar todas las actualizaciones en batch
    if (updates.size > 0) {
        updates.forEach((sheetUpdates, sheetName) => {
            const sheet = ss.getSheetByName(sheetName);
            Logger.log(`Aplicando ${sheetUpdates.length} actualizaciones a la hoja ${sheetName}`);
            sheetUpdates.forEach(update => {
                sheet.getRange(update.rowIndex, analystColIndex + 1).setValue(update.analyst);
            });
        });
        SpreadsheetApp.flush();
    }
    
    Logger.log(`Proceso de asignación de 2 etapas completado. ${allPendingRecords.length} registros evaluados.`);

  } catch (e) {
    Logger.error(`Error fatal en assignPendingRecords: ${e.message} \nStack: ${e.stack}`);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Obtiene los registros asignados a un analista, aplicando filtros.
 * Primero ejecuta la lógica de asignación de registros pendientes.
 * @param {string} token El token de sesión del usuario.
 * @param {object} filters Objeto con filtros (status, branch).
 * @returns {Array<Object>} Un array de objetos de registro para la vista.
 */
function getRecordsForAnalyst(token, filters) {
  return withAuth(token, (user) => {
    try {
      if (user.role !== 'Analista' && user.role !== 'Admin') {
        throw new Error('Acceso denegado. Se requiere rol de Analista o Administrador.');
      }

      // Ejecutar la asignación de registros. La función tiene su propio lock para evitar concurrencia.
      assignPendingRecords();
      
      const ss = SpreadsheetApp.openById(SheetManager.SPREADSHEET_ID);
      const allSheets = ss.getSheets();
      const partitionRegex = /^(V_.+|B_.+|REG_|V_.+_B_.+)_\d{4}_(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)$/;

      const headers = SheetManager.SHEET_CONFIG['Respuestas'].headers;
      const headerMap = headers.reduce((map, header, i) => {
          map[header] = i;
          return map;
      }, {});

      let assignedRecords = [];

      allSheets.forEach(sheet => {
        const sheetName = sheet.getName();
        if (!partitionRegex.test(sheetName) || sheet.getLastRow() <= 1) return;

        const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
        data.forEach((row, index) => {
          const analystAssigned = normalizeEmail(row[headerMap['AnalistaAsignado']]);
          const recordSucursal = String(row[headerMap['Sucursal']] || '').trim();
          const recordStatus = String(row[headerMap['EstadoAnalista']] || 'Pendiente').trim();

          // El Admin ve todo, el Analista solo lo suyo
          const isForCurrentUser = user.role === 'Admin' || analystAssigned === normalizeEmail(user.email);

          if (isForCurrentUser) {
            // Aplicar filtros
            const statusFilter = filters.status || 'Pendiente';
            const branchFilter = filters.branch || 'TODAS';

            const statusMatch = statusFilter === 'Todos' || recordStatus === statusFilter;
            const branchMatch = branchFilter === 'TODAS' || recordSucursal === branchFilter;

            if (statusMatch && branchMatch) {
              const recordObject = {};
              headers.forEach((header, i) => {
                recordObject[header] = row[i];
              });
              // Añadir un identificador único y robusto para acciones
              recordObject.recordIdentifier = JSON.stringify({ sheet: sheetName, row: index + 2 });
              assignedRecords.push(recordObject);
            }
          }
        });
      });

      // Ordenar por fecha descendente para mostrar lo más nuevo primero
      assignedRecords.sort((a, b) => new Date(b.Timestamp).getTime() - new Date(a.Timestamp).getTime());
      
      // Mapear a un formato más simple para el cliente, similar al original
      return assignedRecords.map(r => {
        const timestamp = r.Timestamp;
        const fechaPago = r['Fecha de la Transferencia o Pago'];

        return {
            'ID Registro': r.id_registro,
            'Timestamp': timestamp instanceof Date ? timestamp.toISOString() : timestamp,
            'Vendedor': r.Vendedor,
            'Nombre Cliente': r['Nombre Cliente'],
            'Factura': r.Factura,
            'Monto Pagado': r['Monto Pagado'],
            'Forma de Pago': r['Forma de Pago'],
            'Banco Emisor': r['Banco Emisor'],
            'Banco Receptor': r['Banco Receptor'],
            'Nro. de Referencia': r['Nro. de Referencia'],
            'Fecha de la Transferencia o Pago': fechaPago instanceof Date ? fechaPago.toISOString() : fechaPago,
            'Sucursal': r.Sucursal,
            'EstadoRegistro': r.EstadoAnalista || 'Pendiente',
            'recordIdentifier': r.recordIdentifier
        };
      });
    } catch (e) {
      Logger.error(`Error fatal dentro de getRecordsForAnalyst: ${e.message} Stack: ${e.stack}`);
      return []; // Devolver siempre un array vacío en caso de error.
    }
  });
}

/**
 * Actualiza el estado de un registro de cobranza.
 * @param {string} token El token de sesión.
 * @param {string} identifier El identificador JSON del registro (hoja y fila).
 * @param {string} newStatus El nuevo estado ('Procesado' o 'Rechazado').
 * @param {string|null} comment El comentario (requerido para rechazos).
 * @returns {string} Mensaje de confirmación.
 */

// function updateRecordStatus(token, identifier, newStatus, comment) {
//     return withAuth(token, (user) => {
//         if (user.role !== 'Analista' && user.role !== 'Admin') {
//             throw new Error('Acceso denegado.');
//         }
//         if (!['Procesado', 'Rechazado', 'Pendiente'].includes(newStatus)) {
//             throw new Error('Estado no válido.');
//         }
//         if (newStatus === 'Rechazado' && (!comment || comment.trim() === '')) {
//             throw new Error('Se requiere un comentario para rechazar un registro.');
//         }

//         const { sheet: sheetName, row: rowIndex } = JSON.parse(identifier);
//         if (!sheetName || !rowIndex) {
//             throw new Error('Identificador de registro inválido.');
//         }

//         const sheet = SpreadsheetApp.openById(SheetManager.SPREADSHEET_ID).getSheetByName(sheetName);
//         if (!sheet) {
//             throw new Error(`Hoja de partición '${sheetName}' no encontrada.`);
//         }

//         const headers = SheetManager.SHEET_CONFIG['Respuestas'].headers;
//         const statusCol = headers.indexOf('EstadoAnalista') + 1;
//         const commentCol = headers.indexOf('ComentarioAnalista') + 1;
//         const analystCol = headers.indexOf('AnalistaAsignado') + 1;

//         if (statusCol === 0 || commentCol === 0 || analystCol === 0) {
//             throw new Error('Columnas de estado/comentario no configuradas.');
//         }
        
//         // Verificación de seguridad: Asegurarse de que el analista solo modifique lo suyo (admin puede todo)
//         const assignedAnalyst = sheet.getRange(rowIndex, analystCol).getValue();
//         if (user.role === 'Analista' && normalizeEmail(assignedAnalyst) !== normalizeEmail(user.email)) {
//             throw new Error('No tiene permiso para modificar un registro que no le fue asignado.');
//         }

//         sheet.getRange(rowIndex, statusCol).setValue(newStatus);
        
//         // Si se revierte a Pendiente, limpiar el comentario. Si se rechaza, guardarlo.
//         if (newStatus === 'Pendiente') {
//             sheet.getRange(rowIndex, commentCol).setValue('');
//         } else if (comment) {
//             sheet.getRange(rowIndex, commentCol).setValue(comment);
//         }

//         Logger.log(`Registro en ${sheetName}, fila ${rowIndex} actualizado a ${newStatus} por ${user.email}`);
//         return `Registro actualizado a "${newStatus}" con éxito.`;
//     });
// }


/**
 * Actualiza el estado de un registro de cobranza y registra la acción en una hoja de auditoría.
 * @param {string} token El token de sesión.
 * @param {string} identifier El identificador JSON del registro (hoja y fila).
 * @param {string} newStatus El nuevo estado ('Procesado', 'Rechazado', 'Pendiente').
 * @param {string|null} comment El comentario (requerido para rechazos).
 * @returns {string} Mensaje de confirmación.
 */
function updateRecordStatus(token, identifier, newStatus, comment) {
    return withAuth(token, (user) => {
        // --- Validación de permisos y datos de entrada ---
        if (user.role !== 'Analista' && user.role !== 'Admin') {
            throw new Error('Acceso denegado. Se requiere rol de Analista o Administrador.');
        }
        if (!['Procesado', 'Rechazado', 'Pendiente'].includes(newStatus)) {
            throw new Error('El estado proporcionado no es válido.');
        }
        if (newStatus === 'Rechazado' && (!comment || comment.trim() === '')) {
            throw new Error('Se requiere un comentario para rechazar un registro.');
        }

        // --- Deserialización y validación del identificador del registro ---
        const { sheet: sheetName, row: rowIndex } = JSON.parse(identifier);
        if (!sheetName || !rowIndex) {
            throw new Error('El identificador del registro es inválido o está corrupto.');
        }

        const sheet = SpreadsheetApp.openById(SheetManager.SPREADSHEET_ID).getSheetByName(sheetName);
        if (!sheet) {
            throw new Error(`La hoja de partición '${sheetName}' no fue encontrada.`);
        }

        // --- Mapeo de columnas para evitar errores por cambios en la estructura ---
        const headers = SheetManager.SHEET_CONFIG['Respuestas'].headers;
        const statusCol = headers.indexOf('EstadoAnalista') + 1;
        const commentCol = headers.indexOf('ComentarioAnalista') + 1;
        const analystCol = headers.indexOf('AnalistaAsignado') + 1;
        const idRegistroCol = headers.indexOf('id_registro') + 1;

        if ([statusCol, commentCol, analystCol, idRegistroCol].includes(0)) {
            throw new Error('Una o más columnas críticas (Estado, Comentario, Analista, ID) no están configuradas en la hoja de respuestas.');
        }

        // --- Verificación de seguridad: El analista solo modifica sus registros asignados (Admin puede todo) ---
        const assignedAnalyst = sheet.getRange(rowIndex, analystCol).getValue();
        if (user.role === 'Analista' && normalizeEmail(assignedAnalyst) !== normalizeEmail(user.email)) {
            throw new Error('No tiene permiso para modificar un registro que no le ha sido asignado.');
        }

        // --- INICIO DE LA LÓGICA DE AUDITORÍA ---
        // 1. Obtener el estado actual y el ID del registro ANTES de modificarlo.
        const currentStatus = sheet.getRange(rowIndex, statusCol).getValue() || 'Pendiente';
        const recordId = sheet.getRange(rowIndex, idRegistroCol).getValue();

        // 2. Registrar la acción en la hoja de auditoría de analistas.
        const auditSheet = SheetManager.getSheet('Auditoria_Analistas');
        auditSheet.appendRow([
            new Date(),                      // Timestamp de la acción
            user.email,                      // Analista que realiza el cambio
            recordId,                        // ID único del registro afectado
            currentStatus,                   // Estado anterior
            newStatus,                       // Nuevo estado
            comment || ''                    // Comentario asociado (si existe)
        ]);
        // --- FIN DE LA LÓGICA DE AUDITORÍA ---

        // --- Actualización del registro ---
        sheet.getRange(rowIndex, statusCol).setValue(newStatus);
        
        // Limpiar o establecer el comentario según el nuevo estado.
        if (newStatus === 'Pendiente') {
            sheet.getRange(rowIndex, commentCol).setValue('');
        } else if (comment) {
            sheet.getRange(rowIndex, commentCol).setValue(comment);
        }

        Logger.log(`Registro ${recordId} en ${sheetName} (fila ${rowIndex}) actualizado de "${currentStatus}" a "${newStatus}" por ${user.email}.`);
        return `Registro actualizado a "${newStatus}" con éxito.`;
    });
}

/**
 * Obtiene las sucursales disponibles para el filtro del analista.
 * Un admin ve todas las sucursales de la hoja de vendedores.
 * Un analista ve solo las sucursales que tiene asignadas en la hoja 'analista'.
 * @param {string} token El token de sesión.
 * @returns {Array<string>} Un array de nombres de sucursales.
 */
function getSucursalesDisponibles(token) {
    return withAuth(token, (user) => {
        if (user.role === 'Admin') {
            // Admin ve todas las sucursales únicas de la lista de vendedores
            const sheet = SheetManager.getSheet('obtenerVendedoresPorUsuario');
            if (sheet.getLastRow() < 2) return [];
            const sucursalesData = sheet.getRange(2, 4, sheet.getLastRow() - 1, 1).getValues().flat();
            return [...new Set(sucursalesData.map(s => String(s).trim()).filter(s => s))];
        }
        
        if (user.role === 'Analista') {
            // Analista ve sus sucursales asignadas
            const analystMap = getAvailableAnalysts();
            const sucursales = analystMap.get(normalizeEmail(user.email)) || [];
            if (sucursales.includes('TODAS')) {
                // Si tiene 'TODAS', devolver todas las sucursales existentes
                const sheet = SheetManager.getSheet('obtenerVendedoresPorUsuario');
                if (sheet.getLastRow() < 2) return [];
                const sucursalesData = sheet.getRange(2, 4, sheet.getLastRow() - 1, 1).getValues().flat();
                return [...new Set(sucursalesData.map(s => String(s).trim()).filter(s => s))];
            }
            return sucursales;
        }

        return []; // Otros roles no ven sucursales
    });
}

function getVendedoresDisponibles(token) {
    return withAuth(token, (user) => {
        const records = getRecordsForAnalyst(token, { status: 'Todos', branch: 'TODAS' });
        const vendedores = [...new Set(records.map(r => r.Vendedor).filter(v => v))];
        return vendedores.sort();
    });
}

function getClientesDisponibles(token) {
    return withAuth(token, (user) => {
        const records = getRecordsForAnalyst(token, { status: 'Todos', branch: 'TODAS' });
        const clientes = [...new Set(records.map(r => r['Nombre Cliente']).filter(c => c))];
        return clientes.sort();
    });
}

function getBancosReceptoresDisponibles(token) {
    return withAuth(token, (user) => {
        const records = getRecordsForAnalyst(token, { status: 'Todos', branch: 'TODAS' });
        const bancos = [...new Set(records.map(r => r['Banco Receptor']).filter(b => b))];
        return bancos.sort();
    });
}

function checkForUpdates(token, clientTimestamp) {
    return withAuth(token, (user) => {
        const scriptCache = CacheService.getScriptCache();
        const serverTimestamp = scriptCache.get('lastUpdateTimestamp');
        
        const newUpdates = serverTimestamp && Number(serverTimestamp) > Number(clientTimestamp);

        return {
            newUpdates: newUpdates,
            serverTimestamp: serverTimestamp
        };
    });
}

// --- Funciones para la UI de Administración de Overrides ---

function getOverrideData(token, forceRefresh = false) {
    return withAuth(token, (user) => {
        if (user.role !== 'Admin') throw new Error('Acceso denegado.');

        if (forceRefresh) {
            PropertiesService.getScriptProperties().deleteProperty('allSellers');
            PropertiesService.getScriptProperties().deleteProperty('allAnalysts');
            PropertiesService.getScriptProperties().deleteProperty('currentRules');
        }

        const allSellers = CacheManager.get('allSellers', 3600, () => {
            const dataFetcher = new DataFetcher();
            return dataFetcher.fetchAllVendedoresFromSheet();
        });

        const allAnalysts = CacheManager.get('allAnalysts', 3600, () => {
            return Array.from(getAvailableAnalysts().keys());
        });

        const currentRules = CacheManager.get('currentRules', 60, () => {
            const overrideSheet = SheetManager.getSheet('AsignacionOverrides');
            const overrideData = overrideSheet.getLastRow() > 1 ? overrideSheet.getRange(2, 1, overrideSheet.getLastRow() - 1, 2).getValues() : [];
            return overrideData.map(([sellerCode, analystEmail]) => ({
                sellerCode: String(sellerCode).trim(),
                analystEmail: normalizeEmail(analystEmail)
            }));
        });

        return { allSellers, allAnalysts, currentRules };
    });
}

function addOverrideRule(token, sellerCode, analystEmail) {
    return withAuth(token, (user) => {
        if (user.role !== 'Admin') throw new Error('Acceso denegado.');
        if (!sellerCode || !analystEmail) throw new Error('Se requieren el código de vendedor y el email del analista.');

        const overrideSheet = SheetManager.getSheet('AsignacionOverrides');
        const lastRow = overrideSheet.getLastRow();
        let sellerCodes = [];
        if (lastRow > 1) {
            sellerCodes = overrideSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
        }

        const rowIndex = sellerCodes.indexOf(sellerCode);

        if (rowIndex !== -1) {
            // Si el vendedor ya existe, actualiza el analista
            overrideSheet.getRange(rowIndex + 2, 2).setValue(analystEmail);
            Logger.log(`Regla de override actualizada por ${user.email}: Vendedor ${sellerCode} -> Analista ${analystEmail}`);
            PropertiesService.getScriptProperties().deleteProperty('currentRules'); // Invalidate cache
            return "Regla actualizada con éxito.";
        } else {
            // Si no existe, añade una nueva fila
            overrideSheet.appendRow([sellerCode, analystEmail]);
            Logger.log(`Nueva regla de override creada por ${user.email}: Vendedor ${sellerCode} -> Analista ${analystEmail}`);
            PropertiesService.getScriptProperties().deleteProperty('currentRules'); // Invalidate cache
            return "Nueva regla creada con éxito.";
        }
    });
}

function deleteOverrideRule(token, sellerCode) {
    return withAuth(token, (user) => {
        if (user.role !== 'Admin') throw new Error('Acceso denegado.');
        if (!sellerCode) throw new Error('Se requiere el código de vendedor.');

        const overrideSheet = SheetManager.getSheet('AsignacionOverrides');
        const lastRow = overrideSheet.getLastRow();
        if (lastRow < 2) return "No hay reglas para eliminar.";

        const sellerCodes = overrideSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
        const rowIndex = sellerCodes.findIndex(code => String(code).trim() === sellerCode);

        if (rowIndex !== -1) {
            overrideSheet.deleteRow(rowIndex + 2);
            Logger.log(`Regla de override eliminada por ${user.email} para el vendedor ${sellerCode}`);
            PropertiesService.getScriptProperties().deleteProperty('currentRules'); // Invalidate cache
            return "Regla eliminada con éxito.";
        } else {
            throw new Error('No se encontró la regla para el vendedor especificado.');
        }
    });
}

// #endregion

// #region API pública Apps Script
const cobranzaService = new CobranzaService(new DataFetcher());

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const token = params.token;
  const url = getWebAppUrl();
  let user = null;

  if (token) {
    user = checkAuth(token); // Asumimos que checkAuth devuelve un objeto de usuario con una propiedad 'role'
  }

  if (user) {
    // --- LÓGICA DE ENRUTAMIENTO POR ROL ---
    let templateName;
    const page = String((params.view || params.page || '')).toLowerCase();

    // --- INICIO DE LA MODIFICACIÓN: Enrutamiento para rol Admin ---
    // Motivo: Permitir que los usuarios con rol 'Admin' accedan a la vista de analista.
    /*
    if (page === 'report') {
        templateName = 'Report';
    } else if (user.role === 'Analista') {
        templateName = 'AnalystView'; // Si el rol es Analista, carga su vista
    } else {
        templateName = 'Index'; // Por defecto, o si es Vendedor, carga Index
    }
    */
    if (page === 'report') {
        templateName = 'Report';
    } else if (user.role === 'Analista' || user.role === 'Admin') {
        templateName = 'AnalystView'; // Si el rol es Analista o Admin, carga su vista
    } else {
        templateName = 'Index'; // Por defecto, o si es Vendedor, carga Index
    }
    // --- FIN DE LA MODIFICACIÓN ---
    
    try {
        const template = HtmlService.createTemplateFromFile(templateName);
        template.user = user;
        template.url = url;
        template.token = token;

        if (templateName === 'Report') {
          template.meta = template.meta || { rangeLabel: 'Hoy y Ayer', user };
          template.records = template.records || [];
        }

        return template.evaluate()
          .setTitle(templateName)
          .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');

    } catch (err) {
        // Si no se encuentra un archivo de plantilla (p. ej., AnalystView.html no existe)
        // muestra un error claro en lugar de fallar silenciosamente.
        Logger.error(`Error al renderizar la plantilla '${templateName}': ${err.message}`);
        return HtmlService.createHtmlOutput(
            `<h2>Error del Servidor</h2><p>No se pudo cargar la vista: ${templateName}. Por favor, contacte al administrador.</p>`
        ).setTitle('Error');
    }
      
  } else {
    // Si no hay usuario, muestra la página de autenticación
    const template = HtmlService.createTemplateFromFile('Auth');
    template.url = url;
    return template.evaluate()
      .setTitle('Iniciar Sesión - Conciliapp')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function withAuth(token, action) {
  const user = checkAuth(token);
  if (!user) throw new Error("Sesión inválida o expirada. Por favor, inicie sesión de nuevo.");
  return action(user);
}

function loadVendedores(token, forceRefresh) {
  return withAuth(token, (user) => cobranzaService.getVendedores(user.email, forceRefresh));
}
function cargarClientesEnPregunta1(token, codVendedor) {
  return withAuth(token, () => {
    if (!codVendedor) return '<option value="" disabled selected>Seleccione un cliente</option>';
    return cobranzaService.getClientesHtml(codVendedor);
  });
}
function obtenerFacturas(token, codVendedor, codCliente) {
  return withAuth(token, () => cobranzaService.getFacturas(codVendedor, codCliente));
}
function obtenerTasaBCV(token) {
  return withAuth(token, () => cobranzaService.getBcvRate());
}
function obtenerBancos(token) {
  return withAuth(token, () => cobranzaService.getBancos());
}
function enviarDatos(token, datos) {
  return withAuth(token, (user) => cobranzaService.submitData(datos, user.email));
}
function obtenerRegistrosEnviados(token, vendedorFiltro) {
  return withAuth(token, (user) => cobranzaService.getRecentRecords(vendedorFiltro, user.email));
}
function eliminarRegistro(token, rowIndex) {
  return withAuth(token, (user) => cobranzaService.deleteRecord(rowIndex, user.email));
}

function getNotificacionesPagosConfirmados(token) {
  return withAuth(token, (user) => {
    const ss = SpreadsheetApp.openById(SheetManager.SPREADSHEET_ID);
    const allSheets = ss.getSheets();
    const partitionRegex = /^BANCO-.*_\d{4}$/;
    const headers = SheetManager.SHEET_CONFIG['RegistrosBancoEFactory'].headers;
    const userColIndex = headers.indexOf('codvendedor');
    const timestampColIndex = headers.indexOf('Timestamp');

    let notifications = [];

    const dataFetcher = new DataFetcher();
    const vendedores = dataFetcher.fetchVendedoresFromSheetByUser(user.email);
    const codVendedores = vendedores.map(v => v.codigo);

    allSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (partitionRegex.test(sheetName) && sheet.getLastRow() > 1) {
        const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
        data.forEach(row => {
          const codVendedor = row[userColIndex];
          const timestamp = row[timestampColIndex];

          if (codVendedores.includes(codVendedor) && timestamp) {
            const now = new Date();
            const diff = now.getTime() - new Date(timestamp).getTime();
            const hours = diff / (1000 * 60 * 60);
            if (hours <= 24) {
              notifications.push({
                factura: row[headers.indexOf('factura')],
                monto: row[headers.indexOf('monto')],
                cliente: row[headers.indexOf('Nombre')]
              });
            }
          }
        });
      }
    });
    return notifications;
  });
}
function descargarRegistrosPDF(token, vendedorFiltro) {
  return withAuth(token, (user) => {
    try {
      const tz = Session.getScriptTimeZone();
      
      // --- INICIO DE LA CORRECCIÓN ---
      const now = new Date();
      
      const end = new Date(now);
      end.setHours(23, 59, 59, 999);

      const start = new Date(now);
      start.setDate(start.getDate() - 1);
      start.setHours(0, 0, 0, 0);
      
      Logger.log(`Buscando registros para PDF desde: ${start.toISOString()} hasta: ${end.toISOString()}`);
      // --- FIN DE LA CORRECCIÓN ---

      const reportService = new ReportService(new DataFetcher());
      const records = reportService.getRecordsInDateRange(user.email, vendedorFiltro, start, end);
      
      Logger.log(`Encontrados ${records.length} registros para el PDF.`);

      const meta = {
        user,
        rangeLabel: `desde ${Utilities.formatDate(start, tz, 'dd/MM/yyyy HH:mm')} hasta ${Utilities.formatDate(end, tz, 'dd/MM/yyyy HH:mm')}`,
        filename: `Registros_${Utilities.formatDate(start, tz, 'yyyyMMdd')}_${Utilities.formatDate(end, tz, 'yyyyMMdd')}.pdf`,
        generatedDate: Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm')
      };

      const pdf = reportService.buildPdf(records, meta);
      
      return {
        filename: meta.filename,
        base64: Utilities.base64Encode(pdf.getBytes())
      };
    } catch (e) {
      Logger.error(`Error en descargarRegistrosPDF: ${e.message} en la línea ${e.lineNumber}`);
      throw e;
    }
  });
}

function exportarVistaFiltradaPDF(token, filters) {
    return withAuth(token, (user) => {
        try {
            // 1. Obtener los registros usando los filtros, igual que lo hace la vista.
            const registros = getRecordsForAnalyst(token, filters);

            if (!registros || registros.length === 0) {
                throw new Error("No se encontraron registros con los filtros seleccionados para generar el PDF.");
            }

            // 2. Reutilizar la lógica de generación de PDF de la función `exportarAnalisisPDF`
            const registrosPorVendedor = registros.reduce((acc, r) => {
                const vendedor = r.Vendedor || 'Sin Vendedor';
                if (!acc[vendedor]) acc[vendedor] = [];
                acc[vendedor].push(r);
                return acc;
            }, {});

            // --- INICIO DE LA MODIFICACIÓN: Ordenar por Banco Receptor ---
            // Motivo: El usuario solicitó que el orden principal en el PDF sea por banco receptor.
            /*
            for (const vendedor in registrosPorVendedor) {
                registrosPorVendedor[vendedor].sort((a, b) => {
                    // Criterio principal: Fecha de Registro (Timestamp) ascendente
                    const dateA = new Date(a.Timestamp);
                    const dateB = new Date(b.Timestamp);
                    if (dateA.getTime() !== dateB.getTime()) {
                        return dateA.getTime() - dateB.getTime();
                    }

                    // Criterio secundario: Banco Receptor alfabético
                    const bancoA = a['Banco Receptor'] || '';
                    const bancoB = b['Banco Receptor'] || '';
                    return bancoA.localeCompare(bancoB);
                });
            }
            */
            for (const vendedor in registrosPorVendedor) {
                registrosPorVendedor[vendedor].sort((a, b) => {
                    // Criterio principal: Banco Receptor alfabético
                    const bancoA = a['Banco Receptor'] || '';
                    const bancoB = b['Banco Receptor'] || '';
                    const bancoCompare = bancoA.localeCompare(bancoB);
                    if (bancoCompare !== 0) {
                        return bancoCompare;
                    }

                    // Criterio secundario: Fecha de Registro (Timestamp) ascendente
                    const dateA = new Date(a.Timestamp);
                    const dateB = new Date(b.Timestamp);
                    return dateA.getTime() - dateB.getTime();
                });
            }
            // --- FIN DE LA MODIFICACIÓN ---

            const tz = Session.getScriptTimeZone();
            const generatedDate = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
            let pageCounter = 1;
            const totalPages = Object.keys(registrosPorVendedor).length;
            let htmlContent = `<html><head><style>@page { size: A4 landscape; margin: 1cm; } body { font-family: Arial, sans-serif; font-size: 9pt; } .page-break { page-break-after: always; } .report-table { border-collapse: collapse; width: 100%; } .report-table th, .report-table td { border: 1px solid #ccc; padding: 4px; text-align: left; } .report-table th { background-color: #f2f2f2; } .header, .footer { width: 100%; position: fixed; } .header { top: -0.5cm; text-align: center; } .footer { bottom: 0; font-size: 8pt; } .text-right { text-align: right; } .text-center { text-align: center; } h1, h2, h3 { margin: 0; padding: 0; }</style></head><body>`;

            for (const vendedor in registrosPorVendedor) {
                const loteRegistros = registrosPorVendedor[vendedor];
                const totalRegistrosVendedor = loteRegistros.length;
                htmlContent += `<div class="header"><h2>Reporte de Cobranza</h2></div><h3>Vendedor: ${vendedor}</h3><p style="margin-top:0; padding-top:0;">Analista Asignado: ${user.name}</p><table class="report-table"><thead><tr><th>ID Reg.</th><th>F. Registro</th><th>Cliente</th><th>Factura(s)</th><th>Monto</th><th>Forma de Pago</th><th>Banco Emisor</th><th>Banco Receptor</th><th>Referencia</th><th>F. Pago</th><th>Sucursal</th><th>Estado</th></tr></thead><tbody>`;
                loteRegistros.forEach(r => {
                    const monto = typeof r['Monto Pagado'] === 'number' ? r['Monto Pagado'].toLocaleString('es-VE', { minimumFractionDigits: 2 }) : r['Monto Pagado'];
                    const fechaRegistro = r.Timestamp ? Utilities.formatDate(new Date(r.Timestamp), tz, 'dd/MM/yy') : 'N/A';
                    const fechaPago = r['Fecha de la Transferencia o Pago'] ? Utilities.formatDate(new Date(r['Fecha de la Transferencia o Pago']), tz, 'dd/MM/yy') : 'N/A';
                    htmlContent += `<tr><td>${r['ID Registro'] || 'N/A'}</td><td>${fechaRegistro}</td><td>${r['Nombre Cliente'] || 'N/A'}</td><td>${r.Factura || 'N/A'}</td><td class="text-right">${monto}</td><td>${r['Forma de Pago'] || 'N/A'}</td><td>${r['Banco Emisor'] || 'N/A'}</td><td>${r['Banco Receptor'] || 'N/A'}</td><td>${r['Nro. de Referencia'] || 'N/A'}</td><td>${fechaPago}</td><td>${r.Sucursal || 'N/A'}</td><td>${r.EstadoRegistro || 'Pendiente'}</td></tr>`;
                });
                htmlContent += `</tbody></table><div class="footer"><table><tr><td width="33.3%">Fecha de Impresión: ${generatedDate}</td><td width="33.3%" class="text-center">Total de Registros en esta página: ${totalRegistrosVendedor}</td><td width="33.3%" class="text-right">Página ${pageCounter} de ${totalPages}</td></tr></table></div>`;
                if (pageCounter < totalPages) {
                    htmlContent += '<div class="page-break"></div>';
                }
                pageCounter++;
            }
            htmlContent += '</body></html>';

            const blob = Utilities.newBlob(htmlContent, 'text/html', 'reporte.html').getAs(MimeType.PDF);
            const filename = `Reporte_Analisis_${Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmm')}.pdf`;
            blob.setName(filename);

            return {
                filename: filename,
                base64: Utilities.base64Encode(blob.getBytes())
            };

        } catch (e) {
            Logger.error(`Error en exportarVistaFiltradaPDF: ${e.message} \nStack: ${e.stack}`);
            throw new Error(`No se pudo generar el PDF: ${e.message}`);
        }
    });
}

function exportarAnalisisPDF(token, registros, analistaNombre) {
    return withAuth(token, (user) => {
        try {
            if (!registros || registros.length === 0) {
                throw new Error("No se proporcionaron registros para generar el PDF.");
            }

            // 1. Agrupar registros por Vendedor
            const registrosPorVendedor = registros.reduce((acc, r) => {
                const vendedor = r.Vendedor || 'Sin Vendedor';
                if (!acc[vendedor]) {
                    acc[vendedor] = [];
                }
                acc[vendedor].push(r);
                return acc;
            }, {});

            // 2. Ordenar registros dentro de cada grupo por fecha de pago
            for (const vendedor in registrosPorVendedor) {
                registrosPorVendedor[vendedor].sort((a, b) => {
                    const dateA = a['Fecha de la Transferencia o Pago'] ? new Date(a['Fecha de la Transferencia o Pago']) : 0;
                    const dateB = b['Fecha de la Transferencia o Pago'] ? new Date(b['Fecha de la Transferencia o Pago']) : 0;
                    if (!dateA) return 1; // Mover registros sin fecha al final
                    if (!dateB) return -1;
                    return dateA.getTime() - dateB.getTime();
                });
            }

            const tz = Session.getScriptTimeZone();
            const generatedDate = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
            let pageCounter = 1;
            const totalPages = Object.keys(registrosPorVendedor).length;
            let htmlContent = `
                <html>
                <head>
                    <style>
                        @page { size: A4 landscape; margin: 1cm; }
                        body { font-family: Arial, sans-serif; font-size: 9pt; }
                        .page-break { page-break-after: always; }
                        .report-table { border-collapse: collapse; width: 100%; }
                        .report-table th, .report-table td { border: 1px solid #ccc; padding: 4px; text-align: left; }
                        .report-table th { background-color: #f2f2f2; }
                        .header, .footer { width: 100%; position: fixed; }
                        .header { top: -0.5cm; text-align: center; }
                        .footer { bottom: 0; font-size: 8pt; }
                        .text-right { text-align: right; }
                        .text-center { text-align: center; }
                        h1, h2, h3 { margin: 0; padding: 0; }
                    </style>
                </head>
                <body>`;

            for (const vendedor in registrosPorVendedor) {
                const loteRegistros = registrosPorVendedor[vendedor];
                const totalRegistrosVendedor = loteRegistros.length;

                // --- Cabecera de la página ---
                htmlContent += `
                    <div class="header">
                        <h2>Reporte de Cobranza</h2>
                    </div>
                `;

                // --- Título del reporte para este vendedor ---
                htmlContent += `
                    <h3>Vendedor: ${vendedor}</h3>
                    <p style="margin-top:0; padding-top:0;">Analista Asignado: ${analistaNombre || user.name}</p>
                `;

                // --- Tabla de registros ---
                htmlContent += `
                    <table class="report-table">
                        <thead>
                            <tr>
                                <th>ID Reg.</th>
                                <th>F. Registro</th>
                                <th>Cliente</th>
                                <th>Factura(s)</th>
                                <th>Monto</th>
                                <th>Forma de Pago</th>
                                <th>Banco Emisor</th>
                                <th>Banco Receptor</th>
                                <th>Referencia</th>
                                <th>F. Pago</th>
                                <th>Sucursal</th>
                                <th>Estado</th>
                            </tr>
                        </thead>
                        <tbody>`;

                loteRegistros.forEach(r => {
                    const monto = typeof r['Monto Pagado'] === 'number' ? r['Monto Pagado'].toLocaleString('es-VE', { minimumFractionDigits: 2 }) : r['Monto Pagado'];
                    const fechaRegistro = r.Timestamp ? Utilities.formatDate(new Date(r.Timestamp), tz, 'dd/MM/yy') : 'N/A';
                    const fechaPago = r['Fecha de la Transferencia o Pago'] ? Utilities.formatDate(new Date(r['Fecha de la Transferencia o Pago']), tz, 'dd/MM/yy') : 'N/A';
                    
                    htmlContent += `
                        <tr>
                            <td>${r['ID Registro'] || 'N/A'}</td>
                            <td>${fechaRegistro}</td>
                            <td>${r['Nombre Cliente'] || 'N/A'}</td>
                            <td>${r.Factura || 'N/A'}</td>
                            <td class="text-right">${monto}</td>
                            <td>${r['Forma de Pago'] || 'N/A'}</td>
                            <td>${r['Banco Emisor'] || 'N/A'}</td>
                            <td>${r['Banco Receptor'] || 'N/A'}</td>
                            <td>${r['Nro. de Referencia'] || 'N/A'}</td>
                            <td>${fechaPago}</td>
                            <td>${r.Sucursal || 'N/A'}</td>
                            <td>${r.EstadoRegistro || 'Pendiente'}</td>
                        </tr>`;
                });

                htmlContent += `
                        </tbody>
                    </table>`;

                // --- Pie de página ---
                htmlContent += `
                    <div class="footer">
                        <table>
                            <tr>
                                <td width="33.3%">Fecha de Impresión: ${generatedDate}</td>
                                <td width="33.3%" class="text-center">Total de Registros en esta página: ${totalRegistrosVendedor}</td>
                                <td width="33.3%" class="text-right">Página ${pageCounter} de ${totalPages}</td>
                            </tr>
                        </table>
                    </div>
                `;

                if (pageCounter < totalPages) {
                    htmlContent += '<div class="page-break"></div>';
                }
                pageCounter++;
            }

            htmlContent += '</body></html>';

            const blob = Utilities.newBlob(htmlContent, 'text/html', 'reporte.html').getAs(MimeType.PDF);
            const filename = `Reporte_Analisis_${Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmm')}.pdf`;
            blob.setName(filename);

            return {
                filename: filename,
                base64: Utilities.base64Encode(blob.getBytes())
            };

        } catch (e) {
            Logger.error(`Error en exportarAnalisisPDF: ${e.message} \nStack: ${e.stack}`);
            throw new Error(`No se pudo generar el PDF: ${e.message}`);
        }
    });
}

// #endregion



// #region Lógica de Particionamiento (Restaurada)
const PARTITION_BY_VENDOR_PREFIX = 'V_';
const PARTITION_BY_BANK_PREFIX = 'B_';
const PARTITION_GENERAL_PREFIX = 'REG_';

function decidePartitionType(record) {
  const props = PropertiesService.getScriptProperties();
  const partitionStrategy = props.getProperty('PARTITION_STRATEGY') || 'MONTHLY'; // 'DAILY', 'WEEKLY', 'MONTHLY'
  const partitionBy = props.getProperty('PARTITION_BY') || 'NONE'; // 'NONE', 'VENDOR', 'BANK', 'VENDOR_AND_BANK'
  return { strategy: partitionStrategy, by: partitionBy };
}

function getPartitionName(date, opts) {
  const year = date.getFullYear();
  const monthNames = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
  const month = monthNames[date.getMonth()];
  let prefix = PARTITION_GENERAL_PREFIX;

  if (opts.type.by === 'VENDOR' && opts.vendedor) {
    prefix = `${PARTITION_BY_VENDOR_PREFIX}${opts.vendedor}_`;
  } else if (opts.type.by === 'BANK' && opts.banco) {
    prefix = `${PARTITION_BY_BANK_PREFIX}${opts.banco}_`;
  } else if (opts.type.by === 'VENDOR_AND_BANK' && opts.vendedor && opts.banco) {
    prefix = `${PARTITION_BY_VENDOR_PREFIX}${opts.vendedor}_${PARTITION_BY_BANK_PREFIX}${opts.banco}_`;
  }

  return `${prefix}${year}_${month}`;
}

function ensurePartitionSheet(ss, sheetName, header) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (header && header.length > 0) {
      sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    Logger.log(`Se ha creado la nueva hoja de partición: ${sheetName}`);
  }
  return sheet;
}

function rotacionMensual_() {
  const ss = SpreadsheetApp.openById(SheetManager.SPREADSHEET_ID);
  const newSheetName = getPartitionName(new Date(), { type: { by: 'NONE' } });
  const header = SheetManager.SHEET_CONFIG['Respuestas'].headers;
  ensurePartitionSheet(ss, newSheetName, header);
  Logger.log(`Ejecución de rotación mensual. Asegurada partición: ${newSheetName}`);
}

function crearTriggerRotacionMensual() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'rotacionMensual_') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('rotacionMensual_')
    .timeBased()
    .onMonthDay(1)
    .atHour(2)
    .create();
  Logger.log('Trigger de rotación mensual creado/actualizado correctamente.');
}
// #endregion

// #region Config helpers
function sincronizarVendedoresDesdeApi() {
  const dataFetcher = new DataFetcher();
  const api = dataFetcher.api;
  const sheet = SheetManager.getSheet('obtenerVendedoresPorUsuario');
  const query = `SELECT TRIM(v.correo) AS correo,  TRIM(v.cod_ven) AS codvendedor, 
  CONCAT(TRIM(v.cod_ven), '-', TRIM(v.nom_ven)) 
  AS vendedor_completo, TRIM(s.nom_suc) AS sucursal FROM vendedores v 
  JOIN sucursales s ON s.cod_suc = v.cod_suc where v.status='A';`;
  const vendedores = api.fetchData(query);
  if (vendedores && vendedores.length > 0) {
    sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent();
    const values = vendedores.map(v => [v.correo, v.vendedor_completo, v.codvendedor,v.sucursal]);
    sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
    Logger.log(`Sincronización de vendedores completada. ${vendedores.length} registros actualizados.`);
    return `Sincronización completada. ${vendedores.length} vendedores actualizados.`;
  } else {
    Logger.log('Sincronización de vendedores: No se encontraron registros.');
    return 'No se encontraron vendedores para sincronizar.';
  }
}

function sincronizarSucursalesPorAnalista() {
  const dataFetcher = new DataFetcher();
  const sheet = SheetManager.getSheet('analista');
  const sucursales = dataFetcher.fetchSucursalesPorAnalistaFromApi();
  
  if (sucursales && sucursales.length > 0) {
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
    
    const values = sucursales.map(s => [s.sucursal, s.codigousuario]);
    sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
    
    const message = `Sincronización de sucursales por analista completada. ${sucursales.length} registros actualizados.`;
    Logger.log(message);
    return message;
  } else {
    const message = 'Sincronización de sucursales por analista: No se encontraron registros para sincronizar.';
    Logger.log(message);
    return message;
  }
}

function setApiQueries() {
  const props = PropertiesService.getScriptProperties();
  const facturasQuery = `SELECT 
      TRIM(cc.documento) AS documento,
      CAST((cc.mon_net * cc.tasa) AS DECIMAL(18,2)) AS mon_sal,
      CAST(cc.fec_ini AS DATE) AS fec_ini,
      'USD' AS cod_mon
    FROM cuentas_cobrar cc
    JOIN clientes c ON c.cod_cli = cc.cod_cli
    WHERE cc.cod_tip = 'FACT' 
      AND cc.cod_cli = '{safeCodCliente}' 
      AND cc.cod_ven = '{safeCodVendedor}' 
      AND cc.mon_sal>0
      ORDER BY cc.fec_ini DESC`;
  props.setProperty('FACTURAS_QUERY', facturasQuery);

  const vendedoresQuery = `SELECT TRIM(v.correo) AS correo,  TRIM(v.cod_ven) AS codvendedor,
   CONCAT(TRIM(v.cod_ven), '-',
   TRIM(v.nom_ven)) AS vendedor_completo, TRIM(s.nom_suc) AS sucursal 
   FROM vendedores v JOIN sucursales s ON s.cod_suc = v.cod_suc;`;
  props.setProperty('VENDEDORES_QUERY', vendedoresQuery);

  const clientesQuery = `WITH clientes_filtrados AS (  SELECT cod_cli   FROM cuentas_cobrar   WHERE cod_tip = 'FACT'     AND cod_ven = '{codVendedor}'   GROUP BY cod_cli ) SELECT cf.cod_cli AS Cod_Cliente,       c.nom_cli  AS Nombre FROM clientes_filtrados cf JOIN clientes c ON c.cod_cli = cf.cod_cli order by 2 asc;`;
  props.setProperty('CLIENTES_QUERY', clientesQuery);

  const sucursalesUsuariosQuery = `select s.nom_suc as sucursal,su.cod_usu as codigousuario from Sucursales_Usuarios su left  join sucursales s on s.cod_suc=su.cod_suc order by 2 asc`;
  props.setProperty('SUCURSALES_USUARIOS_QUERY', sucursalesUsuariosQuery);

  const bancosRegistrosQuery = `select  trim(s.nom_suc)as sucursal, concat(trim(cb.cod_cue),'-', trim(cb.nom_cue)) as 'banco_receptor', trim(mv.referencia) as referencia, trim(rc.documento) as cobro, trim(rc.cod_ven) as codvendedor, trim(v.nom_ven)as vendedor, trim(isnull(rc.documento,0)) as factura, cast(mv.mon_deb as decimal (18,2))as monto, trim(mv.comentario)as comentario, trim(c.nom_cli) as Nombre from movimientos_cuentas mv join cuentas_bancarias cb on cb.cod_cue=mv.cod_cue join cobros rc on rc.documento=mv.doc_ori join clientes c on c.cod_cli=rc.cod_cli join sucursales s on s.cod_suc= rc.cod_suc join vendedores v on v.cod_ven=rc.cod_ven where mv.status='confirmado' and mv.fec_ini between '{startDate}' AND '{endDate}'`;
  props.setProperty('BANCOS_REGISTROS_QUERY', bancosRegistrosQuery);
}

function conservarPrimerasPropiedades() {
  var ultimoIndiceAConservar = 6;
  var propiedades = PropertiesService.getScriptProperties();
  var todasLasClaves = propiedades.getKeys();
  todasLasClaves.sort();
  todasLasClaves.forEach(function (clave, indice) {
    if (indice > ultimoIndiceAConservar) {
      propiedades.deleteProperty(clave);
    }
  });
}

function crearTriggerConservarPropiedades() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'conservarPrimerasPropiedades') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('conservarPrimerasPropiedades')
    .timeBased()
    .atHour(1)
    .everyDays(1)
    .create();
}

function crearTriggersSincronizacionYReconciliacion() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sincronizarRegistrosBanco') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('sincronizarRegistrosBanco')
    .timeBased()
    .daily()
    .atHour(1)
    .create();

  Logger.log('Triggers de sincronización y reconciliación creados/actualizados correctamente.');
}

function getPartitionNameForBanco(date, banco) {
  const year = date.getFullYear();
  return `BANCO-${banco}_${year}`;
}

function sincronizarRegistrosBanco() {
  const dataFetcher = new DataFetcher();
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);

  const formatDate = (date) => Utilities.formatDate(date, 'GMT', 'yyyy-MM-dd');

  const registros = dataFetcher.fetchRegistrosBancoFromApi(formatDate(yesterday), formatDate(today));

  if (registros && registros.length > 0) {
    const ss = SpreadsheetApp.openById(SheetManager.SPREADSHEET_ID);
    const header = SheetManager.SHEET_CONFIG['RegistrosBancoEFactory'].headers;

    const registrosPorParticion = registros.reduce((acc, registro) => {
      const partitionName = getPartitionNameForBanco(new Date(), registro.banco_receptor);
      if (!acc[partitionName]) {
        acc[partitionName] = [];
      }
      acc[partitionName].push(registro);
      return acc;
    }, {});

    for (const partitionName in registrosPorParticion) {
      const partitionSheet = ensurePartitionSheet(ss, partitionName, header);
      let existingReferences = [];
      if (partitionSheet.getLastRow() > 1) {
        existingReferences = partitionSheet.getRange(2, 4, partitionSheet.getLastRow() - 1, 1).getValues().flat();
      }
      const nuevosRegistros = registrosPorParticion[partitionName].filter(r => !existingReferences.includes(r.referencia));

      if (nuevosRegistros.length > 0) {
        const values = nuevosRegistros.map(r => [new Date(), r.sucursal, r.banco_receptor, r.referencia, r.cobro, r.codvendedor, r.vendedor, r.factura, r.monto, r.comentario, r.Nombre]);
        partitionSheet.getRange(partitionSheet.getLastRow() + 1, 1, values.length, values[0].length).setValues(values);
        Logger.log(`${values.length} nuevos registros de banco añadidos a la partición ${partitionName}.`);
      }
    }
  }
}

// #endregion