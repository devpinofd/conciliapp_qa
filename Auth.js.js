// auth.js - Módulo de Autenticación con Tokens de Sesión

const SECRET_KEY_PROPERTY = 'AUTH_SECRET_KEY'; // Clave para hashear, guardada en PropertiesService

// Inicializa la clave secreta si no existe
function initializeAuthSecret() {
  const scriptProps = PropertiesService.getScriptProperties();
  if (!scriptProps.getProperty(SECRET_KEY_PROPERTY)) {
    scriptProps.setProperty(SECRET_KEY_PROPERTY, Utilities.getUuid());
  }
}
initializeAuthSecret();

/**
 * Determina el rol de un usuario.
 * @param {string} email El correo del usuario.
 * @returns {string} El rol del usuario ('Analista' o 'Vendedor').
 */
function getUserRole(email) {
    const dataFetcher = new DataFetcher(); // Usamos la clase de Codigo.js
    if (dataFetcher.isUserAdmin(email)) {
        return 'Analista';
    }
    // Si no es admin, asumimos que es vendedor. Se podría añadir más lógica si hay más roles.
    return 'Vendedor';
}

/**
 * Procesa el intento de login de un usuario.
 * Si es exitoso, genera un token de sesión único y almacena los datos del usuario, incluido el rol.
 * @param {string} email El correo del usuario.
 * @param {string} password La contraseña del usuario.
 * @returns {object} Un objeto con el estado del login y el token si es exitoso.
 */
function processLogin(email, password) {
  try {
    const userSheet = SheetManager.getSheet('Usuarios');
    const usersData = userSheet.getDataRange().getValues();
    const normalizedEmail = email.trim().toLowerCase();

    const userRow = usersData.find(row => row[0].toString().trim().toLowerCase() === normalizedEmail);

    if (!userRow) {
      Logger.log(`Intento de login fallido (usuario no encontrado): ${email}`);
      throw new Error("Usuario o contraseña incorrectos.");
    }

    const storedHash = userRow[1];
    const passwordHash = hashPassword(password);

    if (storedHash !== passwordHash) {
      Logger.log(`Intento de login fallido (contraseña incorrecta): ${email}`);
      throw new Error("Usuario o contraseña incorrectos.");
    }

    if (userRow[2] !== 'activo') {
      Logger.log(`Intento de login fallido (cuenta inactiva): ${email}`);
      throw new Error("La cuenta no está activa. Contacte al administrador.");
    }

    // --- Lógica de Token de Sesión MEJORADA ---
    const sessionCache = CacheService.getUserCache();
    const token = Utilities.getUuid();
    const userRole = getUserRole(normalizedEmail); // Obtenemos el rol

    // Guardar el token en caché, asociándolo con email, nombre y ROL.
    // El token expira en 6 horas (21600 segundos).
    const userData = { 
        email: normalizedEmail, 
        name: userRow[3] || normalizedEmail.split('@')[0],
        role: userRole // ¡Añadimos el rol a la sesión!
    };
    sessionCache.put(token, JSON.stringify(userData), 21600);

    Logger.log(`Login exitoso para: ${email} con rol: ${userRole}`);
    return { status: 'SUCCESS', token: token, role: userRole }; // Devolvemos el rol al cliente

  } catch (e) {
    Logger.error(`Error en processLogin: ${e.message}`);
    throw e; // Lanza el error para que el cliente lo maneje
  }
}

/**
 * Valida un token de sesión y devuelve los datos del usuario, incluido el rol.
 * @param {string} token El token a validar.
 * @returns {object|null} Los datos del usuario si el token es válido, de lo contrario null.
 */
function checkAuth(token) {
  if (!token) {
    return null;
  }
  const sessionCache = CacheService.getUserCache();
  const userDataJson = sessionCache.get(token);

  if (userDataJson) {
    const userData = JSON.parse(userDataJson);
    // Si por alguna razón el rol no está en la sesión, lo volvemos a calcular
    if (!userData.role) {
        userData.role = getUserRole(userData.email);
        sessionCache.put(token, JSON.stringify(userData), 21600); // Actualizamos la caché
    }
    return userData;
  }
  
  return null;
}

/**
 * Cierra la sesión de un usuario eliminando su token de la caché.
 * @param {string} token El token de sesión a invalidar.
 */
function logoutUser(token) {
  try {
    if (token) {
      const sessionCache = CacheService.getUserCache();
      sessionCache.remove(token);
      Logger.log(`Sesión cerrada para el token: ${token}`);
    }
  } catch (e) {
    Logger.error(`Error en logoutUser: ${e.message}`);
  }
}

function registerUser(name, email, password) {
  const normalizedEmail = email.trim().toLowerCase();
  if (!validateUserInVendedoresSheet(normalizedEmail)) {
    throw new Error("No está autorizado para registrarse. Su correo no se encuentra en la lista de vendedores.");
  }
  
  const userSheet = SheetManager.getSheet('Usuarios');
  const usersData = userSheet.getRange("A:A").getValues().flat();
  if (usersData.map(e => e.trim().toLowerCase()).includes(normalizedEmail)) {
    throw new Error("Este correo electrónico ya está registrado.");
  }

  const passwordHash = hashPassword(password);
  userSheet.appendRow([normalizedEmail, passwordHash, 'activo', name, new Date()]);
  Logger.log(`Nuevo usuario registrado: ${email}`);
  return "Usuario registrado con éxito.";
}

function validateUserInVendedoresSheet(email) {
    const sheet = SheetManager.getSheet('obtenerVendedoresPorUsuario');
    if (sheet.getLastRow() < 2) return false;
    const emailList = sheet.getRange("A2:A" + sheet.getLastRow()).getValues().flat();
    return emailList.map(e => e.trim().toLowerCase()).includes(email);
}

function hashPassword(password) {
  const secret = PropertiesService.getScriptProperties().getProperty(SECRET_KEY_PROPERTY);
  const signature = Utilities.computeHmacSha256Signature(password, secret);
  return Utilities.base64Encode(signature);
}