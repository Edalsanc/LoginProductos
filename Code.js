/*************** CONFIGURACIÓN ****************/
const SHEET_USERS = 'Usuarios';
const SHEET_PRODUCTS = 'Productos';
const SHEET_CLIENTS = 'Clientes';
const SHEET_KARDEX = 'Kardex';
const SHEET_ROLES = 'Roles';

const SECRET_KEY = 'MiClaveSecretaSuperSegura_2024!';

/*************** UTILIDADES Y SEGURIDAD ****************/
function generateUniqueId(prefix) {
  return prefix + '-' + Utilities.getUuid().substring(0, 8).toUpperCase();
}

function encryptPassword(text) {
  if (!text) return '';
  const salt = Utilities.getUuid().substring(0, 8);
  const derivedKey = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, SECRET_KEY + salt);
  const textBytes = Utilities.newBlob(text.toString()).getBytes();
  const cipherBytes = [];
  
  for (let i = 0; i < textBytes.length; i++) {
    const textByte = textBytes[i] & 0xFF;
    const keyByte = derivedKey[i % derivedKey.length] & 0xFF;
    const xorByte = (textByte ^ keyByte) & 0xFF;
    cipherBytes.push(xorByte > 127 ? xorByte - 256 : xorByte);
  }
  return salt + ':' + Utilities.base64Encode(cipherBytes);
}

function decryptPassword(cipherString) {
  if (!cipherString) return '';
  const str = cipherString.toString();
  const parts = str.split(':');
  if (parts.length !== 2) return str;
  
  try {
    const salt = parts[0];
    const cipherBase64 = parts[1];
    const derivedKey = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, SECRET_KEY + salt);
    const cipherBytes = Utilities.base64Decode(cipherBase64);
    const decryptedBytes = [];
    
    for (let i = 0; i < cipherBytes.length; i++) {
      const cipherByte = cipherBytes[i] & 0xFF;
      const keyByte = derivedKey[i % derivedKey.length] & 0xFF;
      const origByte = (cipherByte ^ keyByte) & 0xFF;
      decryptedBytes.push(origByte > 127 ? origByte - 256 : origByte);
    }
    return Utilities.newBlob(decryptedBytes).getDataAsString();
  } catch (error) {
    return str;
  }
}

/*************** INICIALIZACIÓN DE HOJAS ****************/
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Hoja Usuarios
  let sheetUsers = ss.getSheetByName(SHEET_USERS);
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet(SHEET_USERS);
    // Incluímos las columnas Rol y Estado
    sheetUsers.appendRow(['ID', 'Fecha de Registro', 'Nombre', 'Email', 'Contraseña', 'Rol', 'Estado']);
    sheetUsers.getRange('A1:G1').setFontWeight('bold').setBackground('#f3f4f6');
    sheetUsers.setFrozenRows(1);
  } else {
    // Si la hoja ya existía en la versión anterior, agregamos los encabezados faltantes (columna 6 y 7)
    if (sheetUsers.getLastColumn() === 5) {
      sheetUsers.getRange(1, 6, 1, 2).setValues([['Rol', 'Estado']]).setFontWeight('bold').setBackground('#f3f4f6');
    }
  }

  // 2. Hoja Productos (con carga inicial por defecto)
  let sheetProducts = ss.getSheetByName(SHEET_PRODUCTS);
  if (!sheetProducts) {
    sheetProducts = ss.insertSheet(SHEET_PRODUCTS);
    sheetProducts.appendRow(['código_id', 'nombre', 'descripcion']);
    sheetProducts.getRange('A1:C1').setFontWeight('bold').setBackground('#f3f4f6');
    sheetProducts.setFrozenRows(1);
    
    // Lista de productos solicitados
    const defaultProducts = [
      "Cobre", "Cobre Brillante", "Bronce", "Pelusa", "Quemado", "Tarro Rojo", 
      "Cobre 1B", "Radiador Aluminio", "Radiador Cobre", "Radiador Mixto", 
      "Baterías", "SCRA", "Aluminio Perfil", "Aluminio Olla", "Aluminio Grueso", 
      "Aluminio Sucio", "Plomo", "Limalla", "karla", "Antimonio", "Acero", 
      "Persiana", "Tarro Limpio", "Lamina Litografia", "Alambrón"
    ];
    
    defaultProducts.forEach((prod, index) => {
      let code = 'PRD-' + Utilities.formatString('%03d', index + 1);
      sheetProducts.appendRow([code, prod, 'Producto general ' + prod]);
    });
  }

  // 3. Hoja Clientes
  let sheetClients = ss.getSheetByName(SHEET_CLIENTS);
  if (!sheetClients) {
    sheetClients = ss.insertSheet(SHEET_CLIENTS);
    sheetClients.appendRow(['ID_Cliente', 'Nombre', 'Contacto']);
    sheetClients.getRange('A1:C1').setFontWeight('bold').setBackground('#f3f4f6');
    sheetClients.setFrozenRows(1);
    sheetClients.appendRow(['CLI-001', 'Cliente Mostrador', 'N/A']);
    sheetClients.appendRow(['CLI-002', 'Proveedor Local', 'N/A']);
  }

  // 4. Hoja Kardex
  let sheetKardex = ss.getSheetByName(SHEET_KARDEX);
  if (!sheetKardex) {
    sheetKardex = ss.insertSheet(SHEET_KARDEX);
    sheetKardex.appendRow([
      'Fecha', 'Tipo', 'Producto_ID', 'Producto_Nombre', 'Cliente', 
      'Unds_Mov', 'C_Unitar_Mov', 'C_Total_Mov', 
      'Unds_Saldo', 'C_Unitar_Saldo', 'C_Total_Saldo', 'Usuario'
    ]);
    sheetKardex.getRange('A1:L1').setFontWeight('bold').setBackground('#e0e7ff');
    sheetKardex.setFrozenRows(1);
  }

  // 5. Hoja Roles
  let sheetRoles = ss.getSheetByName(SHEET_ROLES);
  if (!sheetRoles) {
    sheetRoles = ss.insertSheet(SHEET_ROLES);
    sheetRoles.appendRow(['ID_Rol', 'Nombre', 'Estado']);
    sheetRoles.getRange('A1:C1').setFontWeight('bold').setBackground('#f3f4f6');
    sheetRoles.setFrozenRows(1);
    
    // Carga inicial de roles por defecto solicitados
    sheetRoles.appendRow(['ROL-001', 'Admin', 'Activo']);
    sheetRoles.appendRow(['ROL-002', 'Operador', 'Activo']);
    sheetRoles.appendRow(['ROL-003', 'Soporte', 'Activo']);
  }
}

/*************** FUNCIONES DE AUTENTICACIÓN ****************/
function doGet(e) {
  setupDatabase(); // Asegura que las hojas existan al cargar
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Sistema Kardex & Autenticación')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function handleRegister(params) {
  const { name, email, password } = params;
  if (!name || !email || !password) return { success: false, message: 'Faltan campos obligatorios.' };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][3].toString().toLowerCase() === email.toLowerCase()) {
      return { success: false, message: 'Este correo ya está registrado.' };
    }
  }
  
  const userId = generateUniqueId('USR');
  sheet.appendRow([userId, new Date(), name, email.toLowerCase(), encryptPassword(password)]);
  
  return { success: true, message: 'Registro exitoso.', user: { id: userId, name: name, email: email.toLowerCase() } };
}

function handleLogin(params) {
  const { email, password } = params;
  if (!email || !password) return { success: false, message: 'Debes ingresar correo y contraseña.' };

  setupDatabase();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    // Si el usuario está inactivo, lo saltamos
    if (data[i][6] === 'Inactivo') continue;

    if (data[i][3].toString().toLowerCase() === email.toLowerCase()) {
      if (decryptPassword(data[i][4]) === password) {
        return { 
          success: true, 
          user: { 
            id: data[i][0], 
            name: data[i][2], 
            email: data[i][3],
            role: data[i][5] || 'Admin' // A los usuarios viejos se les tratará como Admin para no perder acceso
          } 
        };
      } else {
        return { success: false, message: 'Contraseña incorrecta.' };
      }
    }
  }
  return { success: false, message: 'El usuario no existe o fue desactivado.' };
}

/*************** FUNCIONES DEL INVENTARIO KARDEX ****************/

// Obtiene todos los datos necesarios para cargar el Dashboard
function getAppData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Obtener Productos (Añadida captura de Estado)
  const prodData = ss.getSheetByName(SHEET_PRODUCTS).getDataRange().getValues();
  const products = prodData.slice(1).map(row => ({
    id: row[0], nombre: row[1], desc: row[2], estado: row[3] || 'Activo'
  }));
  
  // Obtener Clientes (Añadida captura de Estado)
  const cliData = ss.getSheetByName(SHEET_CLIENTS).getDataRange().getValues();
  const clients = cliData.slice(1).map(row => ({
    id: row[0], nombre: row[1], contacto: row[2], estado: row[3] || 'Activo'
  }));
  
  // Obtener Kardex
  const sheetKardex = ss.getSheetByName(SHEET_KARDEX);
  const karData = sheetKardex.getDataRange().getValues();
  const kardex = karData.slice(1).map(row => ({
    fecha: row[0] instanceof Date ? row[0].toISOString() : String(row[0]),
    tipo: row[1], prodId: row[2], prodNombre: row[3], cliente: row[4],
    unds: row[5], c_unitar: row[6], c_total: row[7],
    saldo_unds: row[8], saldo_c_unitar: row[9], saldo_c_total: row[10], usuario: row[11]
  })).reverse();
  
  // Obtener Usuarios
  const sheetUsers = ss.getSheetByName(SHEET_USERS);
  let users = [];
  if(sheetUsers) {
    const userData = sheetUsers.getDataRange().getValues();
    users = userData.slice(1).map(row => ({
      id: row[0], nombre: row[2], email: row[3], rol: row[5] || 'Admin', estado: row[6] || 'Activo'
    }));
  }
  
  // Obtener Roles
  const sheetRoles = ss.getSheetByName(SHEET_ROLES);
  let roles = [];
  if(sheetRoles) {
    const roleData = sheetRoles.getDataRange().getValues();
    roles = roleData.slice(1).map(row => ({
      id: row[0], nombre: row[1], estado: row[2] || 'Activo'
    }));
  }
  
  return { success: true, products: products, clients: clients, kardex: kardex, users: users, roles: roles };
}

// Guarda un movimiento de entrada o salida en el Kardex
function saveMovement(params) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    const { tipo, fecha, prodId, cliente, unds, c_unitar, c_total, userId } = params;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetKardex = ss.getSheetByName(SHEET_KARDEX);
    const karData = sheetKardex.getDataRange().getValues();
    
    // Obtener nombre del producto
    const prodData = ss.getSheetByName(SHEET_PRODUCTS).getDataRange().getValues();
    let prodNombre = 'Desconocido';
    for(let i=1; i<prodData.length; i++){
      if(prodData[i][0] === prodId) { prodNombre = prodData[i][1]; break; }
    }

    // Buscar el último saldo de ESTE producto específicamente
    let lastSaldoUnds = 0;
    let lastSaldoCTotal = 0;
    
    for (let i = karData.length - 1; i > 0; i--) {
      if (karData[i][2] === prodId) { // Coincide el Producto_ID
        lastSaldoUnds = parseFloat(karData[i][8] || 0);
        lastSaldoCTotal = parseFloat(karData[i][10] || 0);
        break;
      }
    }

    const movUnds = parseFloat(unds);
    const movCTotal = parseFloat(c_total);
    
    let newSaldoUnds = 0;
    let newSaldoCTotal = 0;
    let newSaldoCUnitar = 0;

    // FÓRMULAS DE KARDEX (Promedio Ponderado)
    if (tipo === 'Entrada') {
      newSaldoUnds = lastSaldoUnds + movUnds;
      newSaldoCTotal = lastSaldoCTotal + movCTotal;
    } else if (tipo === 'Salida') {
      // Validar stock
      if (movUnds > lastSaldoUnds) {
        return { success: false, message: 'Stock insuficiente para esta salida.' };
      }
      newSaldoUnds = lastSaldoUnds - movUnds;
      newSaldoCTotal = lastSaldoCTotal - movCTotal;
    }

    // Calcular nuevo costo unitario promedio (Evitar división por 0)
    if (newSaldoUnds > 0) {
      newSaldoCUnitar = newSaldoCTotal / newSaldoUnds;
    } else {
      newSaldoCUnitar = 0; 
      newSaldoCTotal = 0; // Ajuste para limpiar decimales basura si el stock es 0
    }

    // Guardar Fila
    sheetKardex.appendRow([
      fecha, tipo, prodId, prodNombre, cliente,
      movUnds, parseFloat(c_unitar), movCTotal,
      newSaldoUnds, newSaldoCUnitar, newSaldoCTotal, userId
    ]);

    return { success: true, message: 'Movimiento registrado correctamente.' };

  } catch (error) {
    return { success: false, message: 'Error interno: ' + error.message };
  } finally {
    lock.releaseLock();
  }
}

/*************** GESTIÓN DE PRODUCTOS Y CLIENTES ****************/

// Registra un nuevo producto en la hoja 'Productos'
function saveProduct(params) {
  const { id, nombre, descripcion } = params;
  if (!id || !nombre) return { success: false, message: 'El Código y el Nombre son obligatorios.' };
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PRODUCTS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === id.toLowerCase()) {
      return { success: false, message: 'Ya existe un producto con ese Código ID.' };
    }
  }
  
  sheet.appendRow([id.toUpperCase(), nombre, descripcion || '', 'Activo']);
  return { success: true, message: 'Producto registrado exitosamente.' };
}

// Actualiza un producto existente
function updateProduct(params) {
  const { id, nombre, descripcion } = params;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PRODUCTS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === id.toLowerCase()) {
      sheet.getRange(i + 1, 2).setValue(nombre);
      sheet.getRange(i + 1, 3).setValue(descripcion || '');
      return { success: true, message: 'Producto actualizado exitosamente.' };
    }
  }
  return { success: false, message: 'Producto no encontrado.' };
}

// Inactiva (Eliminado suave) un producto
function deleteProduct(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PRODUCTS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === id.toLowerCase()) {
      sheet.getRange(i + 1, 4).setValue('Inactivo');
      return { success: true, message: 'Producto eliminado.' };
    }
  }
  return { success: false, message: 'Producto no encontrado.' };
}

// Registra un nuevo cliente/proveedor en la hoja 'Clientes'
function saveClient(params) {
  const { id, nombre, contacto } = params;
  if (!id || !nombre) return { success: false, message: 'El ID y el Nombre son obligatorios.' };
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CLIENTS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === id.toLowerCase()) {
      return { success: false, message: 'Ya existe un cliente o proveedor con ese ID.' };
    }
  }
  
  sheet.appendRow([id.toUpperCase(), nombre, contacto || '', 'Activo']);
  return { success: true, message: 'Cliente/Proveedor registrado exitosamente.' };
}

// Actualiza un cliente existente
function updateClient(params) {
  const { id, nombre, contacto } = params;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CLIENTS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === id.toLowerCase()) {
      sheet.getRange(i + 1, 2).setValue(nombre);
      sheet.getRange(i + 1, 3).setValue(contacto || '');
      return { success: true, message: 'Cliente actualizado exitosamente.' };
    }
  }
  return { success: false, message: 'Cliente no encontrado.' };
}

// Inactiva (Eliminado suave) un cliente
function deleteClient(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CLIENTS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === id.toLowerCase()) {
      sheet.getRange(i + 1, 4).setValue('Inactivo');
      return { success: true, message: 'Cliente eliminado.' };
    }
  }
  return { success: false, message: 'Cliente no encontrado.' };
}

/*************** GESTIÓN DE USUARIOS (ADMIN) ****************/

// Guardar un nuevo usuario desde el panel
function saveUser(params) {
  const { nombre, email, password, rol } = params;
  if (!nombre || !email || !password) return { success: false, message: 'Faltan campos obligatorios.' };
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][3].toString().toLowerCase() === email.toLowerCase() && data[i][6] !== 'Inactivo') {
      return { success: false, message: 'Este correo ya está registrado.' };
    }
  }
  
  const userId = generateUniqueId('USR');
  sheet.appendRow([userId, new Date(), nombre, email.toLowerCase(), encryptPassword(password), rol, 'Activo']);
  
  return { success: true, message: 'Usuario ' + rol + ' creado exitosamente.' };
}

// Desactivar un usuario existente (Soft Delete)
function deleteUser(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === id) {
      sheet.getRange(i + 1, 7).setValue('Inactivo');
      return { success: true, message: 'Usuario desactivado.' };
    }
  }
  return { success: false, message: 'Usuario no encontrado.' };
}