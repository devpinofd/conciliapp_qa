/**
 * @fileoverview Lógica del servidor para la aplicación de cobranza.
 * Refactorizado con principios SOLID y mejores prácticas.
 * Incluye AnalystPanelService para la lógica del panel de analistas.
 */

// #region Funciones de Particionamiento
/**
 * Genera el nombre de una hoja de partición basado en el código del vendedor.
 * Formato: Respuestas_CODVENDEDOR
 */
function getPartitionName(codVendedor) {
  if (!codVendedor || typeof codVendedor !== 'string' || codVendedor.trim() === '') {
    return 'Respuestas_SIN_VENDEDOR';
  }
  const cleanCodVendedor = codVendedor.replace(/[^a-zA-Z0-9_-]/g, '').trim();
  return `Respuestas_${cleanCodVendedor}`;
}

/**
 * Asegura que una hoja de partición exista.
 */
function ensurePartitionSheet(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (header && header.length > 0) {
      sh.getRange(1, 1, 1, header.length).setValues([header]);
      sh.setFrozenRows(1);
      Logger.log(`Hoja de partición creada: ${name}`);
    }
  }
  return sh;
}
// #endregion

// ... (El resto de tus clases Logger, CacheManager, ApiHandler, SheetManager, DataFetcher y CobranzaService permanecen aquí sin cambios)
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
    try {
        const sheet = SheetManager.getSheet('Auditoria');
        sheet.appendRow([new Date(), Session.getActiveUser().getEmail(), level, String(message)]);
    } catch(e) { console.error("Error al escribir en Auditoria:", e); }
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
      const headers = this.SHEET_CONFIG[sheetName].headers;
      // Asegurarse de que la columna de estado del analista exista
      if (sheetName.startsWith('Respuestas_') && !headers.includes('EstadoAnalista')) {
          headers.push('EstadoAnalista', 'ComentarioAnalista', 'AnalistaAsignado');
      }
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    if (!sheet) {
        // Devolver null si la hoja no existe y no está en la config (para particiones)
        return null;
    }
    return sheet;
  }
}
SheetManager.SPREADSHEET_ID = '1jv3jlYSRfCj9VHPBla0g1l35AFEwNJzkWwrOML5oPY4';
SheetManager.SHEET_CONFIG = {
  'CorreosPermitidos': { headers: null },
  'obtenerVendedoresPorUsuario': { headers: ['correo', 'vendedorcompleto', 'codvendedor'] },
  'Administradores': { headers: ['correo_admin'] },
  'Bancos': { headers: ['Nombre del Banco'] },
  'Respuestas': { // Plantilla para nuevas particiones
    headers: ['Timestamp', 'Vendedor', 'Codigo Cliente', 'Nombre Cliente', 'Factura',
      'Monto Pagado', 'Forma de Pago', 'Banco Emisor', 'Banco Receptor',
      'Nro. de Referencia', 'Tipo de Cobro', 'Fecha de la Transferencia o Pago',
      'Observaciones', 'Usuario Creador', 'EstadoAnalista', 'ComentarioAnalista', 'AnalistaAsignado','Sucursal']
  },
  'Auditoria': { headers: ['Timestamp', 'Usuario', 'Nivel', 'Detalle'] },
  'Registros Eliminados': {
    headers: ['Fecha Eliminación', 'Usuario que Eliminó', 'Timestamp', 'Vendedor',
      'Codigo Cliente', 'Nombre Cliente', 'Factura', 'Monto Pagado',
      'Forma de Pago', 'Banco Emisor', 'Banco Receptor', 'Nro. de Referencia',
      'Tipo de Cobro', 'Fecha de la Transferencia o Pago', 'Observaciones', 'Usuario Creador']
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
    const normalizedUserEmail = userEmail.trim().toLowerCase();
    const sheet = SheetManager.getSheet('obtenerVendedoresPorUsuario');
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
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
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    return data.map(row => ({
      nombre: String(row[1]).trim(),
      codigo: String(row[2]).trim()
    })).filter(v => v.nombre && v.codigo);
  }
  isUserAdmin(userEmail) {
    if (!userEmail) return false;
    const normalizedUserEmail = userEmail.trim().toLowerCase();
    const sheet = SheetManager.getSheet('Administradores');
    if (!sheet || sheet.getLastRow() < 2) return false;
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
    if (!sheet || sheet.getLastRow() < 2) {
      Logger.log('No se encontraron bancos en la hoja.');
      return [];
    }
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
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
    const partitionName = getPartitionName(data.vendedor);
    const header = SheetManager.SHEET_CONFIG['Respuestas'].headers;
    const partitionSheet = ensurePartitionSheet(ss, partitionName, header);

    const facturaCsv = this._normalizeFacturaCsv(data.factura || data.documento || '');
    if (!facturaCsv) throw new Error('Debe indicar al menos una factura.');
    
    const montoNum = parseFloat(data.montoPagado);
    if (isNaN(montoNum) || montoNum <= 0) throw new Error('Monto inválido.');

    let existingReferences = [];
    if (partitionSheet.getLastRow() > 1) {
      existingReferences = partitionSheet.getRange(2, 10, partitionSheet.getLastRow() - 1, 1).getValues().flat();
    }
    if (existingReferences.includes(data.nroReferencia)) {
      throw new Error(`La referencia ya existe en la hoja del vendedor: ${partitionName}`);
    }

    const todosLosVendedores = this.dataFetcher.fetchAllVendedoresFromSheet();
    const vendedorEncontrado = todosLosVendedores.find(v => v.codigo === data.vendedor);
    const nombreCompletoVendedor = vendedorEncontrado ? vendedorEncontrado.nombre : data.vendedor;

    const row = [
      new Date(), nombreCompletoVendedor, data.cliente, data.nombreCliente, facturaCsv,
      montoNum, data.formaPago, data.bancoEmisor, data.bancoReceptor, data.nroReferencia,
      data.tipoCobro, data.fechaTransferenciaPago, data.observaciones, userEmail,
      'Pendiente', '', '' // EstadoAnalista, ComentarioAnalista, AnalistaAsignado
    ];

    partitionSheet.appendRow(row);
    Logger.log(`Formulario enviado por ${userEmail} a la partición ${partitionName}.`);
    return '¡Datos recibidos con éxito!';
  }

  getRecentRecords(vendedor, userEmail) {
    const ss = SpreadsheetApp.openById(SheetManager.SPREADSHEET_ID);
    const isAdmin = this.dataFetcher.isUserAdmin(userEmail);
    let sheetsToRead = [];

    if (isAdmin) {
        if (vendedor && vendedor !== 'Mostrar todos') {
            sheetsToRead.push(getPartitionName(vendedor));
        } else {
            const allSellers = this.dataFetcher.fetchAllVendedoresFromSheet();
            sheetsToRead = allSellers.map(v => getPartitionName(v.codigo));
        }
    } else {
        const mySellers = this.dataFetcher.fetchVendedoresFromSheetByUser(userEmail);
        sheetsToRead = mySellers.map(v => getPartitionName(v.codigo));
    }
    
    let allRecords = [];
    for (const sheetName of sheetsToRead) {
        const sheet = ss.getSheetByName(sheetName);
        if (sheet && sheet.getLastRow() > 1) {
            const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
            values.forEach((row, index) => {
                if(row[0]) allRecords.push({ data: row, sheetName: sheetName, rowIndex: index + 2 });
            });
        }
    }

    allRecords.sort((a, b) => new Date(b.data[0]).getTime() - new Date(a.data[0]).getTime());

    const recentRecords = allRecords.slice(0, this.REGISTROS_POR_PAGINA);
    const now = new Date().getTime();
    const FIVE_MINUTES_IN_MS = 5 * 60 * 1000;

    return recentRecords.map(record => {
      const rowData = record.data;
      const timestamp = new Date(rowData[0]).getTime();
      const puedeEliminar = (rowData[13] === userEmail) && (now - timestamp < FIVE_MINUTES_IN_MS);
      
      return {
        rowIndex: JSON.stringify({ sheetName: record.sheetName, row: record.rowIndex }),
        fechaEnvio: Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
        vendedor: rowData[1],
        clienteNombre: rowData[3],
        factura: rowData[4],
        monto: (typeof rowData[5] === 'number') ? rowData[5].toFixed(2) : rowData[5],
        bancoEmisor: rowData[7],
        bancoReceptor: rowData[8],
        referencia: rowData[9],
        creadoPor: rowData[13],
        puedeEliminar: puedeEliminar
      };
    });
  }

  deleteRecord(rowIndexStr, userEmail) {
    const { sheetName, row } = JSON.parse(rowIndexStr);
    const sheet = SpreadsheetApp.openById(SheetManager.SPREADSHEET_ID).getSheetByName(sheetName);

    if (!sheet) throw new Error(`No se encontró la hoja de partición: ${sheetName}`);

    const rowToDelete = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (rowToDelete[13] !== userEmail) throw new Error('No tienes permiso para eliminar este registro.');
    
    const timestamp = new Date(rowToDelete[0]).getTime();
    const now = new Date().getTime();
    const ageInMs = now - timestamp;
    const FIVE_MINUTES_IN_MS = 5 * 60 * 1000;

    if (ageInMs > FIVE_MINUTES_IN_MS) throw new Error('No se puede eliminar un registro después de 5 minutos.');

    const auditSheet = SheetManager.getSheet('Registros Eliminados');
    auditSheet.appendRow([new Date(), userEmail, ...rowToDelete]);
    sheet.deleteRow(row);
    Logger.log(`Registro eliminado por ${userEmail}. Fila: ${row} en hoja: ${sheetName}`);
    return 'Registro eliminado y archivado con éxito.';
  }
}
// #endregion

// #region Servicios de Analista
/**
 * @class AnalystPanelService
 * Encapsula la lógica para el panel de analistas.
 */
class AnalystPanelService {
    constructor(dataFetcher) {
        this.dataFetcher = dataFetcher;
        this.headerMap = this._createHeaderMap();
    }

    /** Crea un mapa de cabeceras para acceso por nombre en lugar de índice */
    _createHeaderMap() {
        const headers = SheetManager.SHEET_CONFIG['Respuestas'].headers;
        const map = {};
        headers.forEach((header, index) => {
            map[header] = index;
        });
        return map;
    }

    /**
     * Obtiene todas las sucursales (particiones de vendedor) disponibles.
     * @returns {string[]} Lista de nombres de sucursales/vendedores.
     */
    getSucursalesDisponibles() {
        const ss = SpreadsheetApp.openById(SheetManager.SPREADSHEET_ID);
        const allSheets = ss.getSheets();
        const sucursales = new Set();
        const prefix = 'Respuestas_';

        allSheets.forEach(sheet => {
            const sheetName = sheet.getName();
            if (sheetName.startsWith(prefix)) {
                const sucursal = sheetName.substring(prefix.length);
                if(sucursal !== 'SIN_VENDEDOR') {
                    sucursales.add(sucursal);
                }
            }
        });
        return Array.from(sucursales).sort();
    }

    /**
     * Obtiene registros para el panel de analista aplicando filtros.
     * @param {object} filters Filtros aplicados por el analista.
     * @param {string} filters.status 'Pendiente', 'Procesado', etc.
     * @param {string} filters.branch Sucursal/vendedor a filtrar o 'TODAS'.
     * @returns {object[]} Un array de objetos de registro.
     */
    getRecordsForAnalyst(filters) {
        const ss = SpreadsheetApp.openById(SheetManager.SPREADSHEET_ID);
        let sheetsToRead = [];

        if (filters.branch && filters.branch !== 'TODAS') {
            sheetsToRead.push(getPartitionName(filters.branch));
        } else {
            sheetsToRead = this.getSucursalesDisponibles().map(sucursal => getPartitionName(sucursal));
        }

        let allRecords = [];
        const statusFilter = filters.status;

        for (const sheetName of sheetsToRead) {
            const sheet = ss.getSheetByName(sheetName);
            if (sheet && sheet.getLastRow() > 1) {
                const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
                const values = range.getValues();
                
                values.forEach((row, index) => {
                    const recordStatus = row[this.headerMap['EstadoAnalista']] || 'Pendiente';
                    
                    if (statusFilter === 'Todos' || recordStatus === statusFilter) {
                        const recordObject = {
                            'ID Registro': `${sheetName}-${index + 2}`,
                            'recordIdentifier': JSON.stringify({ sheetName, row: index + 2 }),
                            Timestamp: row[this.headerMap['Timestamp']],
                            Vendedor: row[this.headerMap['Vendedor']],
                            'Nombre Cliente': row[this.headerMap['Nombre Cliente']],
                            'Monto Pagado': row[this.headerMap['Monto Pagado']],
                            Sucursal: sheetName.substring('Respuestas_'.length),
                            EstadoRegistro: recordStatus,
                        };
                        allRecords.push(recordObject);
                    }
                });
            }
        }

        allRecords.sort((a, b) => new Date(b.Timestamp).getTime() - new Date(a.Timestamp).getTime());
        return allRecords;
    }

    /**
     * Actualiza el estado de un registro.
     * @param {string} identifier JSON string que identifica la hoja y fila.
     * @param {string} newStatus Nuevo estado ('Procesado', 'Rechazado').
     * @param {string|null} comment Comentario opcional (para rechazos).
     * @param {string} analystEmail Email del analista que realiza la acción.
     * @returns {string} Mensaje de confirmación.
     */
    updateRecordStatus(identifier, newStatus, comment, analystEmail) {
        const { sheetName, row } = JSON.parse(identifier);
        const ss = SpreadsheetApp.openById(SheetManager.SPREADSHEET_ID);
        const sheet = ss.getSheetByName(sheetName);

        if (!sheet) {
            throw new Error(`No se encontró la hoja de partición: ${sheetName}`);
        }

        const statusCol = this.headerMap['EstadoAnalista'] + 1;
        const commentCol = this.headerMap['ComentarioAnalista'] + 1;
        const analystCol = this.headerMap['AnalistaAsignado'] + 1;

        sheet.getRange(row, statusCol).setValue(newStatus);
        sheet.getRange(row, analystCol).setValue(analystEmail);
        if (comment) {
            sheet.getRange(row, commentCol).setValue(comment);
        }
        
        Logger.log(`Analista ${analystEmail} actualizó el registro en ${sheetName}, fila ${row} a estado "${newStatus}"`);
        return `Registro actualizado a "${newStatus}" con éxito.`;
    }
}
// #endregion


// #region API pública Apps Script
const dataFetcher = new DataFetcher();
const cobranzaService = new CobranzaService(dataFetcher);
const analystPanelService = new AnalystPanelService(dataFetcher); // Nueva instancia

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
    // --- Lógica de Enrutamiento por Rol ---
    let templateName;
    const page = String((params.view || params.page || '')).toLowerCase();
    
    if (page === 'report') {
        templateName = 'Report';
    } else if (user.role === 'Analista' || dataFetcher.isUserAdmin(user.email)) {
        templateName = 'AnalystView'; // Redirige a Admin/Analista a su panel
    } else {
        templateName = 'Index'; // Vendedor va al formulario
    }
    
    try {
        const template = HtmlService.createTemplateFromFile(templateName);
        template.user = user;
        template.url = url;
        template.token = token;
        return template.evaluate()
          .setTitle(templateName === 'Index' ? 'Registro de Cobranzas' : 'Panel de Analista')
          .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
    } catch (err) {
        Logger.error(`Error al renderizar plantilla '${templateName}': ${err.message}`);
        return HtmlService.createHtmlOutput(`Error del servidor: ${err.message}`);
    }
  } else {
    // Usuario no autenticado
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
  if (!user) throw new Error("Sesión inválida o expirada.");
  // Adjuntar el email del usuario a la acción para auditoría
  return action(user);
}

// --- Funciones para CobranzaService ---
function loadVendedores(token, forceRefresh) { return withAuth(token, (user) => cobranzaService.getVendedores(user.email, forceRefresh)); }
function cargarClientesEnPregunta1(token, codVendedor) { return withAuth(token, () => cobranzaService.getClientesHtml(codVendedor)); }
function obtenerFacturas(token, codVendedor, codCliente) { return withAuth(token, () => cobranzaService.getFacturas(codVendedor, codCliente)); }
function obtenerTasaBCV(token) { return withAuth(token, () => cobranzaService.getBcvRate()); }
function obtenerBancos(token) { return withAuth(token, () => cobranzaService.getBancos()); }
function enviarDatos(token, datos) { return withAuth(token, (user) => cobranzaService.submitData(datos, user.email)); }
function obtenerRegistrosEnviados(token, vendedorFiltro) { return withAuth(token, (user) => cobranzaService.getRecentRecords(vendedorFiltro, user.email)); }
function eliminarRegistro(token, rowIndex) { return withAuth(token, (user) => cobranzaService.deleteRecord(rowIndex, user.email)); }
function descargarRegistrosPDF(token, vendedorFiltro) { /* ...código de PDF sin cambios... */ }

// --- Funciones para AnalystPanelService ---
function getSucursalesDisponibles(token) {
    return withAuth(token, () => analystPanelService.getSucursalesDisponibles());
}

function getRecordsForAnalyst(token, filters) {
    return withAuth(token, () => analystPanelService.getRecordsForAnalyst(filters));
}

function updateRecordStatus(token, identifier, newStatus, comment) {
    return withAuth(token, (user) => analystPanelService.updateRecordStatus(identifier, newStatus, comment, user.email));
}


// ... (El resto de las funciones de configuración como sincronizarVendedoresDesdeApi, setApiQueries, etc. permanecen aquí)
function sincronizarVendedoresDesdeApi() {
  const dataFetcher = new DataFetcher();
  const api = dataFetcher.api;
  const sheet = SheetManager.getSheet('obtenerVendedoresPorUsuario');
  const query = `SELECT TRIM(correo) AS correo, TRIM(cod_ven) AS codvendedor, CONCAT(TRIM(cod_ven), '-', TRIM(nom_ven)) AS vendedor_completo, trim(sucursales.nom_suc) AS sucursal FROM vendedores INNER JOIN sucursales ON vendedores.cod_suc = sucursales.cod_suc WHERE status='A';`;
  const vendedores = api.fetchData(query);
  if (vendedores && vendedores.length > 0) {
    const range = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn());
    if(range) range.clearContent();
    const values = vendedores.map(v => [v.correo, v.vendedor_completo, v.codvendedor]);
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
      'USD' AS cod_mon,
      S.NOM_SUC AS sucursal
    FROM cuentas_cobrar cc
    JOIN clientes c ON c.cod_cli = cc.cod_cli
    JOIN SUCURSALES s ON s.cod_suc = c.cod_suc
    WHERE cc.cod_tip = 'FACT' 
      AND cc.cod_cli = '{safeCodCliente}' 
      AND cc.cod_ven = '{safeCodVendedor}' 
      ORDER BY cc.fec_ini DESC`;
  props.setProperty('FACTURAS_QUERY', facturasQuery);
  const vendedoresQuery = `SELECT TRIM(v.correo) AS correo,  TRIM(v.cod_ven) AS codvendedor, CONCAT(TRIM(v.cod_ven), '-', TRIM(v.nom_ven)) AS vendedor_completo, trim(s.nom_suc) AS sucursal FROM vendedores v JOIN sucursales s ON s.cod_suc = v.cod_suc;`;
  props.setProperty('VENDEDORES_QUERY', vendedoresQuery);
  const clientesQuery = `WITH clientes_filtrados AS (  SELECT cod_cli  FROM cuentas_cobrar  WHERE cod_tip = 'FACT'    AND cod_ven = '{safeCodVendedor}' GROUP BY cod_cli) SELECT cf.cod_cli AS Codigo_Cliente,       c.nom_cli  AS Nombre FROM clientes_filtrados cf JOIN clientes c ON c.cod_cli = cf.cod_cli;`;
  props.setProperty('CLIENTES_QUERY', clientesQuery);
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

function crearTriggerRotacionMensual() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (trigger) {
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