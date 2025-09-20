/**
 * @fileoverview Lógica del servidor para la aplicación de cobranza.
 * Refactorizado con principios SOLID y mejores prácticas.
 * Incluye soporte para facturas múltiples (CSV) en una sola fila.
 */

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
  'obtenerVendedoresPorUsuario': { headers: ['correo', 'vendedorcompleto', 'codvendedor'] },
  'Administradores': { headers: ['correo_admin'] },
  'Bancos': { headers: ['Nombre del Banco'] },
  'Respuestas': {
    headers: ['Timestamp', 'Vendedor', 'Codigo Cliente', 'Nombre Cliente', 'Factura',
      'Monto Pagado', 'Forma de Pago', 'Banco Emisor', 'Banco Receptor',
      'Nro. de Referencia', 'Tipo de Cobro', 'Fecha de la Transferencia o Pago',
      'Observaciones', 'Usuario Creador', 'EstadoRegistro', 'AnalistaAsignado', 
      'FechaProcesado', 'ComentarioAnalista']
  },
  'Auditoria': { headers: ['Timestamp', 'Usuario', 'Nivel', 'Detalle'] },
  'Registros Eliminados': {
    headers: ['Fecha Eliminación', 'Usuario que Eliminó', 'Timestamp', 'Vendedor',
      'Codigo Cliente', 'Nombre Cliente', 'Factura', 'Monto Pagado',
      'Forma de Pago', 'Banco Emisor', 'Banco Receptor', 'Nro. de Referencia',
      'Tipo de Cobro', 'Fecha de la Transferencia o Pago', 'Observaciones', 'Usuario Creador']
  },
  'Usuarios': {
    headers: ['Correo', 'Contraseña', 'Estado', 'Nombre', 'Fecha Registro', 
              'Rol', 'Empresa', 'Sucursal', 'GrupoVentas']
  }
};

class DataFetcher {
  constructor() { this.api = new ApiHandler(); }
  fetchVendedoresFromSheetByUser(userEmail) {
    if (!userEmail) {
      Logger.error('Se intentó llamar a fetchVendedoresFromSheetByUser sin un email.');
      return [];
    }
    const normalizedUserEmail = userEmail.trim().toLowerCase();
    const sheet = SheetManager.getSheet('obtenerVendedoresPorUsuario');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    const vendedoresFiltrados = data
      .map(row => ({
        email: String(row[0]).trim().toLowerCase(),
        nombre: String(row[1]).trim(),
        codigo: String(row[2]).trim()
      }))
      .filter(v => v.email === normalizedUserEmail && v.nombre && v.codigo);
    if (vendedoresFiltrados.length === 0) {
      Logger.log(`No se encontraron vendedores para el usuario: ${userEmail}`);
    }
    return vendedoresFiltrados;
  }
  fetchAllVendedoresFromSheet() {
    const sheet = SheetManager.getSheet('obtenerVendedoresPorUsuario');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    return data.map(row => ({
      nombre: String(row[1]).trim(),
      codigo: String(row[2]).trim()
    })).filter(v => v.nombre && v.codigo);
  }
  isUserAdmin(userEmail) {
    if (!userEmail) return false;
    const normalizedUserEmail = userEmail.trim().toLowerCase();
    const sheet = SheetManager.getSheet('Administradores');
    if (sheet.getLastRow() < 2) return false;
    const adminEmails = sheet.getRange("A2:A" + sheet.getLastRow()).getValues()
      .flat().map(email => email.trim().toLowerCase());
    return adminEmails.includes(normalizedUserEmail);
  }
  fetchClientesFromApi(codVendedor) {
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
    try {
      const data = this.api.fetchData(query);
      return data.map(row => ({
        nombre: String(row.Nombre).trim(),
        codigo: String(row.Codigo_Cliente).trim()
      }));
    } catch (e) {
      Logger.error(`Error en fetchClientesFromApi: ${e.message}`, { query });
      return [];
    }
  }
  fetchFacturasFromApi(codVendedor, codCliente) {
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
    try {
      const data = this.api.fetchData(query);
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
}

class CobranzaService {
  constructor(dataFetcher) {
    this.dataFetcher = dataFetcher;
    this.REGISTROS_POR_PAGINA = 10;
  }

  static get scriptUserEmail() {
    return Session.getActiveUser().getEmail();
  }

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
      optionsHtml += vendedores.map(v => `<option value="${v.codigo}">${v.nombre}</option>`).join('');
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

    const facturaCsvRaw = data.factura || data.documento || '';
    const facturaCsv = this._normalizeFacturaCsv(facturaCsvRaw);
    if (!facturaCsv) throw new Error('Debe indicar al menos una factura.');
    const montoNum = parseFloat(data.montoPagado);
    if (isNaN(montoNum) || montoNum <= 0) throw new Error('Monto inválido.');
    if (!data.vendedor) throw new Error('Vendedor requerido.');
    if (!data.cliente) throw new Error('Código de cliente requerido.');

    const fecha = new Date(data.fechaTransferenciaPago || Date.now());
    const record = {
        vendedorCodigo: data.vendedor,
        bancoReceptor: data.bancoReceptor,
    };
    const partitionOpts = {
        type: decidePartitionType(record),
        vendedor: record.vendedorCodigo,
        banco: record.bancoReceptor
    };
    const partitionName = getPartitionName(fecha, partitionOpts);
    const header = SheetManager.SHEET_CONFIG['Respuestas'].headers;
    const partitionSheet = ensurePartitionSheet(ss, partitionName, header);

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

    const todosLosVendedores = this.dataFetcher.fetchAllVendedoresFromSheet();
    const vendedorEncontrado = todosLosVendedores.find(v => v.codigo === data.vendedor);
    const nombreCompletoVendedor = vendedorEncontrado ? vendedorEncontrado.nombre : data.vendedor;

    const row = [
      new Date(),
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
      'Pendiente', // EstadoRegistro
      '',          // AnalistaAsignado
      '',          // FechaProcesado
      ''           // ComentarioAnalista
    ];

    partitionSheet.appendRow(row);
    Logger.log(`Formulario enviado por ${userEmail} a la partición ${partitionName}. Facturas: ${facturaCsv}`);
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
        rowIndex: index + 2 
      }));
      allRecords.push(...recordsWithMeta);
    }
    
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
    const sheet = SheetManager.getSheet('Respuestas');
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const inRange = values.filter(row => {
      const ts = new Date(row[0]).getTime();
      return ts >= start.getTime() && ts <= end.getTime();
    });
    const isAdmin = this.dataFetcher.isUserAdmin(userEmail);
    let filtered = inRange;
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
    return filtered.map(row => ({
      fecha: Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
      vendedor: row[1],
      clienteCodigo: row[2],
      clienteNombre: row[3],
      factura: row[4],
      monto: (typeof row[5] === 'number') ? row[5].toFixed(2) : row[5],
      formaPago: row[6],
      bancoEmisor: row[7],
      bancoReceptor: row[8],
      referencia: row[9],
      tipoCobro: row[10],
      fechaPago: row[11],
      observaciones: row[12],
      creadoPor: row[13],
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
  if (token) user = checkAuth(token);
  if (user) {
    const page = String((params.view || params.page || params.template || '')).toLowerCase();
    const templateName = page === 'report' ? 'Report' : 'Index';
    const template = HtmlService.createTemplateFromFile(templateName);
    template.user = user;
    template.url = url;
    template.token = token;
    if (templateName === 'Report') {
      template.meta = template.meta || { rangeLabel: 'Hoy y Ayer', user };
      template.records = template.records || [];
    }
    return template.evaluate()
      .setTitle(templateName === 'Report' ? 'Reporte de Registros' : 'Registro de Cobranzas')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  } else {
    const template = HtmlService.createTemplateFromFile('Auth');
    template.url = url;
    return template.evaluate()
      .setTitle('Iniciar Sesión - Registro de Cobranzas')
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
function descargarRegistrosPDF(token, vendedorFiltro) {
  return withAuth(token, (user) => {
    try {
      const tz = Session.getScriptTimeZone();
      const now = new Date();
      const end = new Date(Utilities.formatDate(now, tz, 'yyyy/MM/dd 23:59:59'));
      const y = new Date(now);
      y.setDate(now.getDate() - 1);
      const start = new Date(Utilities.formatDate(y, tz, 'yyyy/MM/dd 00:00:00'));
      const reportService = new ReportService(new DataFetcher());
      const records = reportService.getRecordsInDateRange(user.email, vendedorFiltro, start, end);
      const meta = {
        user,
        rangeLabel: `desde ${Utilities.formatDate(start, tz, 'dd/MM/yyyy HH:mm')} hasta ${Utilities.formatDate(end, tz, 'dd/MM/yyyy HH:mm')}`,
        filename: `Registros_${Utilities.formatDate(y, tz, 'yyyyMMdd')}_${Utilities.formatDate(now, tz, 'yyyyMMdd')}.pdf`
      };
      const pdf = reportService.buildPdf(records, meta);
      return {
        filename: meta.filename,
        base64: Utilities.base64Encode(pdf.getBytes())
      };
    } catch (e) {
      Logger.error(`Error en descargarRegistrosPDF: ${e.message}`);
      throw e;
    }
  });
}

// Config helpers
function sincronizarVendedoresDesdeApi() {
  const dataFetcher = new DataFetcher();
  const api = dataFetcher.api;
  const sheet = SheetManager.getSheet('obtenerVendedoresPorUsuario');
  const query = `SELECT	TRIM(v.correo) AS correo,TRIM(v.cod_ven) AS codvendedor, CONCAT(TRIM(v.cod_ven), '-', TRIM(v.nom_ven)) AS vendedor_completo, s.nom_suc FROM vendedores v join sucursales  s on s.cod_suc=v.cod_suc where v.status='A';`;
  const vendedores = api.fetchData(query);
  if (vendedores && vendedores.length > 0) {
    sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent();
    const values = vendedores.map(v => [v.correo, v.vendedor_completo, v.codvendedor, v.nom_suc]);
    sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
    Logger.log(`Sincronización de vendedores completada. ${vendedores.length} registros actualizados.`);
    return `Sincronización completada. ${vendedores.length} vendedores actualizados.`;
  } else {
    Logger.log('Sincronización de vendedores: No se encontraron registros.');
    return 'No se encontraron vendedores para sincronizar.';
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
      ORDER BY cc.fec_ini DESC`;
  props.setProperty('FACTURAS_QUERY', facturasQuery);
  const vendedoresQuery = `SELECT TRIM(correo) AS correo,  TRIM(cod_ven) AS codvendedor, CONCAT(TRIM(cod_ven), '-', TRIM(nom_ven)) AS vendedor_completo FROM vendedores;`;
  props.setProperty('VENDEDORES_QUERY', vendedoresQuery);
  const clientesQuery = `SELECT DISTINCT TRIM(COD_CLI) AS Codigo_Cliente, TRIM(NOM_CLI) AS Nombre 
FROM (
    SELECT COD_CLI, NOM_CLI 
    FROM CLIENTES 
    WHERE COD_VEN = '{codVendedor}' 
    UNION 
    SELECT precios_clientes.COD_REG AS Codigo_Cliente, clientes.NOM_CLI AS Nombre
    FROM precios_clientes 
    JOIN CLIENTES ON clientes.COD_CLI = precios_clientes.COD_REG 
    WHERE precios_clientes.COD_ART = '{codVendedor}' 
 ) AS Combinada 
 ORDER BY TRIM(NOM_CLI) ASC`;
  props.setProperty('CLIENTES_QUERY', clientesQuery);
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

/**
 * Crea o actualiza el trigger para la rotación mensual de particiones.
 * Se ejecutará el primer día de cada mes.
 */
function crearTriggerRotacionMensual() {
  // Eliminar triggers antiguos para evitar duplicados
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'rotacionMensual_') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Crear el nuevo trigger
  ScriptApp.newTrigger('rotacionMensual_')
    .timeBased()
    .onMonthDay(1)
    .atHour(2) // Ejecutar a las 2 AM para evitar horas pico
    .create();
    
  Logger.log('Trigger de rotación mensual creado/actualizado correctamente.');
}