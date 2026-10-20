const SHEET_USERS = 'Usuarios';
const SHEET_PRODS = 'Productos';
const SHEET_CLIENTS = 'Clientes';
const SHEET_KARDEX = 'Kardex';
const SHEET_ROLES = 'Roles';
const SHEET_LOGS = 'Logs';
const SHEET_FACTURAS = 'Facturas';
const SECRET_KEY = 'MiClaveSecretaSuperSegura_2024!';

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Usuarios
  let shUsers = ss.getSheetByName(SHEET_USERS);
  if (!shUsers) {
    shUsers = ss.insertSheet(SHEET_USERS);
    shUsers.appendRow(['ID', 'Fecha de Registro', 'Nombre', 'Email', 'Contraseña', 'Rol', 'Estado']);
    shUsers.getRange('A1:G1').setFontWeight('bold').setBackground('#f3f4f6');
    
    // Crear el usuario Administrador por defecto
    const adminId = generateUserId();
    const adminPass = encryptPassword('admin1234');
    shUsers.appendRow([adminId, new Date(), 'Administrador Principal', 'admin@kardex.com', adminPass, 'Admin', 'Activo']);
  }
  
  // 2. Productos
  let shProds = ss.getSheetByName(SHEET_PRODS);
  if (!shProds) {
    shProds = ss.insertSheet(SHEET_PRODS);
    shProds.appendRow(['ID', 'Nombre', 'Descripcion', 'Estado']);
    shProds.getRange('A1:D1').setFontWeight('bold').setBackground('#f3f4f6');
    
    // Poblar productos iniciales requeridos
    const prodList = ['Cobre', 'Cobre Brillante', 'Bronce', 'Pelusa', 'Quemado', 'Tarro Rojo', 'Cobre 1B', 'Radiador Aluminio', 'Radiador Cobre', 'Radiador Mixto', 'Baterías', 'SCRA', 'Aluminio Perfil', 'Aluminio Olla', 'Aluminio Grueso', 'Aluminio Sucio', 'Plomo', 'Limalla', 'karla', 'Antimonio', 'Acero', 'Persiana', 'Tarro Limpio', 'Lamina Litografia', 'Alambrón'];
    prodList.forEach((p, idx) => {
      let cod = 'PRD-' + (idx + 1).toString().padStart(3, '0');
      shProds.appendRow([cod, p, 'Producto de reciclaje', 'Activo']);
    });
  }

  // 3. Clientes
  let shClients = ss.getSheetByName(SHEET_CLIENTS);
  if (!shClients) {
    shClients = ss.insertSheet(SHEET_CLIENTS);
    shClients.appendRow(['ID', 'Nombre', 'Contacto', 'Estado']);
    shClients.getRange('A1:D1').setFontWeight('bold').setBackground('#f3f4f6');
  }

  // 4. Kardex
  let shKardex = ss.getSheetByName(SHEET_KARDEX);
  if (!shKardex) {
    shKardex = ss.insertSheet(SHEET_KARDEX);
    shKardex.appendRow(['ID_Mov', 'Fecha', 'Tipo', 'Producto_ID', 'Producto_Nombre', 'Cliente', 'Unds', 'C_Unitar', 'C_Total', 'Saldo_Unds', 'Saldo_C_Unitar', 'Saldo_C_Total', 'Usuario']);
    shKardex.getRange('A1:M1').setFontWeight('bold').setBackground('#f3f4f6');
  }

  // 5. Roles
  let shRoles = ss.getSheetByName(SHEET_ROLES);
  if (!shRoles) {
    shRoles = ss.insertSheet(SHEET_ROLES);
    shRoles.appendRow(['ID', 'Nombre', 'Estado']);
    shRoles.getRange('A1:C1').setFontWeight('bold').setBackground('#f3f4f6');
    shRoles.appendRow(['ROL-1', 'Admin', 'Activo']);
    shRoles.appendRow(['ROL-2', 'Operador', 'Activo']);
    shRoles.appendRow(['ROL-3', 'Soporte', 'Activo']);
  }

  // 6. Logs (Auditoría)
  let shLogs = ss.getSheetByName(SHEET_LOGS);
  if (!shLogs) {
    shLogs = ss.insertSheet(SHEET_LOGS);
    shLogs.appendRow(['Fecha_Hora', 'Usuario_Email', 'Accion', 'Detalle']);
    shLogs.getRange('A1:D1').setFontWeight('bold').setBackground('#f3f4f6');
  }

  // 7. Facturas (NUEVO MÓDULO)
  let shFacturas = ss.getSheetByName(SHEET_FACTURAS);
  if (!shFacturas) {
    shFacturas = ss.insertSheet(SHEET_FACTURAS);
    shFacturas.appendRow(['ID', 'Numero_Factura', 'Fecha_Factura', 'Descripcion', 'Proveedor_Cliente', 'Observacion', 'Fecha_Creacion', 'Estado']);
    shFacturas.getRange('A1:H1').setFontWeight('bold').setBackground('#f3f4f6');
  }

  // Asegurar que Kardex tenga la columna Factura_ID para la relación histórica
  if (shKardex) {
    const headers = shKardex.getRange(1, 1, 1, shKardex.getLastColumn()).getValues()[0];
    if (headers.indexOf('Factura_ID') === -1 && headers.indexOf('factura_id') === -1) {
      shKardex.getRange(1, headers.length + 1).setValue('Factura_ID').setFontWeight('bold').setBackground('#f3f4f6');
    }
  }
}

function generateUserId() {
  return 'USR-' + Utilities.getUuid().substring(0, 8).toUpperCase();
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

function doGet(e) {
  setupDatabase(); // Inicializa si es la primera vez
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Sistema Kardex Pro')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function registrarLog(usuario, accion, detalle) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetLogs = ss.getSheetByName(SHEET_LOGS);
    if(sheetLogs) {
      // Fecha, Usuario, Accion, Detalle
      sheetLogs.appendRow([new Date(), usuario, accion, detalle]);
    }
  } catch(e) {
    console.log("Error guardando log: " + e);
  }
}

function handleLogin(params) {
  const { email, password } = params;
  if (!email || !password) return { success: false, message: 'Faltan credenciales.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const storedEmail = (data[i][3] || '').toString().toLowerCase();
    const estado = data[i][6] || 'Activo';
    
    if (storedEmail === email.toLowerCase()) {
      if (estado !== 'Activo') return { success: false, message: 'Usuario inactivo o bloqueado.' };
      
      const storedPassword = data[i][4];
      const decryptedStoredPassword = decryptPassword(storedPassword);
      
      if (decryptedStoredPassword === password) {
        registrarLog(email, "LOGIN", "Inicio de sesión exitoso");
        const response = { 
          success: true, 
          message: 'Login exitoso.',
          user: { id: String(data[i][0] || ''), name: String(data[i][2] || ''), email: storedEmail, role: String(data[i][5] || '') }
        };
        return JSON.parse(JSON.stringify(response)); // Sanitización extrema para evitar errores Null
      } else {
        registrarLog(email, "LOGIN_FALLIDO", "Intento de contraseña incorrecta");
        return { success: false, message: 'Contraseña incorrecta.' };
      }
    }
  }
  return { success: false, message: 'El usuario no existe.' };
}

function getAppData(userEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Productos
  const sheetProds = ss.getSheetByName(SHEET_PRODS);
  let products = [];
  if(sheetProds) {
    const prodData = sheetProds.getDataRange().getValues();
    products = prodData.slice(1).filter(row => row[0]).map(row => ({
      id: String(row[0] || ''), nombre: String(row[1] || ''), desc: String(row[2] || ''), estado: String(row[3] || 'Activo')
    }));
  }
  
  // Clientes
  const sheetClients = ss.getSheetByName(SHEET_CLIENTS);
  let clients = [];
  if(sheetClients) {
    const cliData = sheetClients.getDataRange().getValues();
    clients = cliData.slice(1).filter(row => row[0]).map(row => ({
      id: String(row[0] || ''), nombre: String(row[1] || ''), contacto: String(row[2] || ''), estado: String(row[3] || 'Activo')
    }));
  }
  
  const sheetKardex = ss.getSheetByName(SHEET_KARDEX);
  let kardex = [];
  if(sheetKardex) {
    const kardexData = sheetKardex.getDataRange().getValues();
    if (kardexData.length > 0) {
      // Capturar encabezados para mapeo dinámico y evitar errores si las columnas se mueven
      const h = kardexData[0].map(x => String(x).toLowerCase().trim());
      const c = {
        id_mov: h.indexOf('id_mov') > -1 ? h.indexOf('id_mov') : 0,
        fecha: h.indexOf('fecha') > -1 ? h.indexOf('fecha') : 1,
        tipo: h.indexOf('tipo') > -1 ? h.indexOf('tipo') : 2,
        prodId: h.indexOf('producto_id') > -1 ? h.indexOf('producto_id') : 3,
        prodNombre: h.indexOf('producto_nombre') > -1 ? h.indexOf('producto_nombre') : 4,
        cliente: h.indexOf('cliente') > -1 ? h.indexOf('cliente') : 5,
        unds: h.indexOf('unds') > -1 ? h.indexOf('unds') : 6,
        c_unitar: h.indexOf('c_unitar') > -1 ? h.indexOf('c_unitar') : 7,
        c_total: h.indexOf('c_total') > -1 ? h.indexOf('c_total') : 8,
        saldo_unds: h.indexOf('saldo_unds') > -1 ? h.indexOf('saldo_unds') : 9,
        saldo_c_unitar: h.indexOf('saldo_c_unitar') > -1 ? h.indexOf('saldo_c_unitar') : 10,
        saldo_c_total: h.indexOf('saldo_c_total') > -1 ? h.indexOf('saldo_c_total') : 11,
        usuario: h.indexOf('usuario') > -1 ? h.indexOf('usuario') : 12,
        facturaId: h.indexOf('factura_id') > -1 ? h.indexOf('factura_id') : 13
      };

      kardex = kardexData.slice(1).filter(row => row[c.id_mov] || row[c.fecha]).map(row => {
        let parsedDate = row[c.fecha];
        if (parsedDate instanceof Date) parsedDate = parsedDate.toISOString();
        else parsedDate = String(parsedDate || '');
        
        return {
          id_mov: String(row[c.id_mov] || ''), 
          fecha: parsedDate, 
          tipo: String(row[c.tipo] || ''),
          prodId: String(row[c.prodId] || ''), 
          prodNombre: String(row[c.prodNombre] || ''), 
          cliente: String(row[c.cliente] || ''),
          unds: Number(row[c.unds]) || 0, 
          c_unitar: Number(row[c.c_unitar]) || 0, 
          c_total: Number(row[c.c_total]) || 0,
          saldo_unds: Number(row[c.saldo_unds]) || 0, 
          saldo_c_unitar: Number(row[c.saldo_c_unitar]) || 0, 
          saldo_c_total: Number(row[c.saldo_c_total]) || 0,
          usuario: String(row[c.usuario] || ''),
          facturaId: String(row[c.facturaId] || '')
        };
      }).reverse();
    }
  }

  // Facturas
  const sheetFacturas = ss.getSheetByName(SHEET_FACTURAS);
  let facturas = [];
  if(sheetFacturas) {
    const factData = sheetFacturas.getDataRange().getValues();
    facturas = factData.slice(1).filter(row => row[0]).map(row => {
      let parsedDate = row[2];
      if (parsedDate instanceof Date) parsedDate = parsedDate.toISOString();
      else parsedDate = String(parsedDate || '');
      return {
        id: String(row[0] || ''), 
        numero: String(row[1] || ''), 
        fecha: parsedDate, 
        descripcion: String(row[3] || ''), 
        cliente: String(row[4] || ''), 
        observacion: String(row[5] || ''), 
        estado: String(row[7] || 'Activa')
      };
    }).reverse();
  }
  
  // VERIFICACIÓN DE SEGURIDAD (Backend)
  let isAdmin = false;
  const sheetUsers = ss.getSheetByName(SHEET_USERS);
  let users = [];
  
  if(sheetUsers) {
    const userData = sheetUsers.getDataRange().getValues();
    
    // 1. Comprobar si el usuario que solicita tiene rol de Admin real en la base de datos
    for(let i=1; i<userData.length; i++) {
      if((userData[i][3] || '').toString().toLowerCase() === String(userEmail).toLowerCase() && userData[i][5] === 'Admin' && userData[i][6] !== 'Inactivo') {
        isAdmin = true;
        break;
      }
    }
    
    // 2. Solo enviar datos de usuarios si resultó ser Admin
    if (isAdmin) {
      users = userData.slice(1).filter(row => row[0]).map(row => ({
        id: String(row[0] || ''), nombre: String(row[2] || ''), email: String(row[3] || ''), rol: String(row[5] || ''), estado: String(row[6] || 'Activo')
      }));
    }
  }

  // Roles (Solo para Admin)
  const sheetRoles = ss.getSheetByName(SHEET_ROLES);
  let roles = [];
  if(sheetRoles && isAdmin) {
    const roleData = sheetRoles.getDataRange().getValues();
    roles = roleData.slice(1).filter(row => row[0]).map(row => ({
      id: String(row[0] || ''), nombre: String(row[1] || ''), estado: String(row[2] || 'Activo')
    }));
  }
  
  // Logs (Solo para Admin)
  const sheetLogs = ss.getSheetByName(SHEET_LOGS);
  let logs = [];
  if(sheetLogs && isAdmin) {
    const logData = sheetLogs.getDataRange().getValues();
    logs = logData.slice(1).filter(row => row[0]).map(row => {
      let logDate = row[0];
      if (logDate instanceof Date) logDate = logDate.toISOString();
      else logDate = String(logDate || '');
      
      return {
        fecha: logDate, usuario: String(row[1] || ''), accion: String(row[2] || ''), detalle: String(row[3] || '')
      };
    }).reverse();
  }
  
  const response = { success: true, products: products, clients: clients, kardex: kardex, users: users, roles: roles, logs: logs, facturas: facturas };
  // La clave para evitar el error "null" es forzar la sanitización completa:
  return JSON.parse(JSON.stringify(response));
}

function saveMovement(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_KARDEX);
  
  const id_mov = 'MOV-' + Date.now();
  const { tipo, fecha, prodId, cliente, unds, c_unitar, c_total, userEmail, facturaId } = params;
  
  const qUnds = parseFloat(unds);
  const qCUnitar = parseFloat(c_unitar);
  const qCTotal = parseFloat(c_total);
  
  // Encontrar el producto
  const pSheet = ss.getSheetByName(SHEET_PRODS);
  const pData = pSheet.getDataRange().getValues();
  let prodNombre = "Desconocido";
  for(let i=1; i<pData.length; i++) {
    if(pData[i][0] == prodId) { prodNombre = pData[i][1]; break; }
  }

  // Calcular saldos (buscar último movimiento de este producto)
  const kData = sheet.getDataRange().getValues();
  let last_saldo_unds = 0, last_saldo_c_total = 0;
  
  for(let i = kData.length - 1; i >= 1; i--) {
    if(kData[i][3] == prodId) {
      last_saldo_unds = parseFloat(kData[i][9]);
      last_saldo_c_total = parseFloat(kData[i][11]);
      break;
    }
  }

  let new_saldo_unds = 0, new_saldo_c_total = 0, new_saldo_c_unitar = 0;
  
  if(tipo === 'Entrada') {
    new_saldo_unds = last_saldo_unds + qUnds;
    new_saldo_c_total = last_saldo_c_total + qCTotal;
  } else {
    if(qUnds > last_saldo_unds) return { success: false, message: 'Stock insuficiente para esta salida.' };
    new_saldo_unds = last_saldo_unds - qUnds;
    new_saldo_c_total = last_saldo_c_total - qCTotal;
  }
  
  if(new_saldo_unds > 0) {
    new_saldo_c_unitar = new_saldo_c_total / new_saldo_unds;
  }
  
  sheet.appendRow([
    id_mov, fecha, tipo, prodId, prodNombre, cliente, 
    qUnds, qCUnitar, qCTotal, 
    new_saldo_unds, new_saldo_c_unitar, new_saldo_c_total,
    userEmail, facturaId || ''
  ]);
  
  registrarLog(userEmail, "MOVIMIENTO", `Registró ${tipo} de ${qUnds} unds para ${prodNombre}`);
  return { success: true, message: 'Movimiento registrado correctamente.' };
}

function saveProduct(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PRODS);
  const rows = sheet.getDataRange().getValues();
  for(let i=1; i<rows.length; i++) {
    if(rows[i][0].toString().toLowerCase() === data.id.toLowerCase()) return { success: false, message: 'El ID del producto ya existe.' };
  }
  sheet.appendRow([data.id.toUpperCase(), data.nombre, data.descripcion, 'Activo']);
  registrarLog(data.userEmail, "CREAR_PRODUCTO", `Producto creado: ${data.nombre} (${data.id})`);
  return { success: true, message: 'Producto guardado exitosamente.' };
}

function updateProduct(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PRODS);
  const rows = sheet.getDataRange().getValues();
  for(let i=1; i<rows.length; i++) {
    if(rows[i][0] === data.id) {
      sheet.getRange(i+1, 2).setValue(data.nombre);
      sheet.getRange(i+1, 3).setValue(data.descripcion);
      registrarLog(data.userEmail, "EDITAR_PRODUCTO", `Producto modificado: ${data.nombre} (${data.id})`);
      return { success: true, message: 'Producto actualizado.' };
    }
  }
  return { success: false, message: 'Producto no encontrado.' };
}

function deleteProduct(id, userEmail) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PRODS);
  const rows = sheet.getDataRange().getValues();
  for(let i=1; i<rows.length; i++) {
    if(rows[i][0] === id) {
      sheet.getRange(i+1, 4).setValue('Inactivo'); // Soft Delete
      registrarLog(userEmail, "ELIMINAR_PRODUCTO", `Producto desactivado: ${id}`);
      return { success: true, message: 'Producto desactivado (Oculto).' };
    }
  }
  return { success: false, message: 'Producto no encontrado.' };
}

function saveClient(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CLIENTS);
  const rows = sheet.getDataRange().getValues();
  for(let i=1; i<rows.length; i++) {
    if(rows[i][0].toString().toLowerCase() === data.id.toLowerCase()) return { success: false, message: 'El ID / NIT ya existe.' };
  }
  sheet.appendRow([data.id.toUpperCase(), data.nombre, data.contacto, 'Activo']);
  registrarLog(data.userEmail, "CREAR_CLIENTE", `Cliente creado: ${data.nombre}`);
  return { success: true, message: 'Cliente guardado exitosamente.' };
}

function updateClient(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CLIENTS);
  const rows = sheet.getDataRange().getValues();
  for(let i=1; i<rows.length; i++) {
    if(rows[i][0] === data.id) {
      sheet.getRange(i+1, 2).setValue(data.nombre);
      sheet.getRange(i+1, 3).setValue(data.contacto);
      registrarLog(data.userEmail, "EDITAR_CLIENTE", `Cliente modificado: ${data.nombre}`);
      return { success: true, message: 'Cliente actualizado.' };
    }
  }
  return { success: false, message: 'Cliente no encontrado.' };
}

function deleteClient(id, userEmail) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CLIENTS);
  const rows = sheet.getDataRange().getValues();
  for(let i=1; i<rows.length; i++) {
    if(rows[i][0] === id) {
      sheet.getRange(i+1, 4).setValue('Inactivo'); // Soft Delete
      registrarLog(userEmail, "ELIMINAR_CLIENTE", `Cliente desactivado: ${id}`);
      return { success: true, message: 'Cliente desactivado (Oculto).' };
    }
  }
  return { success: false, message: 'Cliente no encontrado.' };
}

function saveUser(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  const rows = sheet.getDataRange().getValues();
  for(let i=1; i<rows.length; i++) {
    if((rows[i][3] || '').toString().toLowerCase() === data.email.toLowerCase()) {
      return { success: false, message: 'El correo electrónico ya existe.' };
    }
  }
  const userId = generateUserId();
  const encryptedPassword = encryptPassword(data.password);
  sheet.appendRow([userId, new Date(), data.nombre, data.email.toLowerCase(), encryptedPassword, data.rol, 'Activo']);
  
  registrarLog(data.userEmail, "CREAR_USUARIO", `Nuevo usuario creado: ${data.email} (${data.rol})`);
  return { success: true, message: 'Usuario creado correctamente.' };
}

function updateUser(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  const rows = sheet.getDataRange().getValues();
  
  // Verificar si el nuevo correo se cruza con alguien más
  for(let i=1; i<rows.length; i++) {
    if(rows[i][0] !== data.id && (rows[i][3] || '').toString().toLowerCase() === data.email.toLowerCase() && rows[i][6] !== 'Inactivo') {
      return { success: false, message: 'Ese correo electrónico ya está en uso por otra persona.' };
    }
  }

  for(let i=1; i<rows.length; i++) {
    if(rows[i][0] === data.id) {
      sheet.getRange(i+1, 3).setValue(data.nombre);
      sheet.getRange(i+1, 4).setValue(data.email.toLowerCase());
      sheet.getRange(i+1, 6).setValue(data.rol);
      
      // Solo encriptar y actualizar la contraseña si el admin escribió una nueva
      if(data.password && data.password.trim() !== '') {
        sheet.getRange(i+1, 5).setValue(encryptPassword(data.password));
      }
      
      registrarLog(data.userEmail, "EDITAR_USUARIO", `Datos modificados para: ${data.email} (${data.rol})`);
      return { success: true, message: 'Cuenta de usuario actualizada exitosamente.' };
    }
  }
  return { success: false, message: 'Usuario no encontrado.' };
}

// --- MÓDULO FACTURAS ---

function saveFactura(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_FACTURAS);
  const rows = sheet.getDataRange().getValues();
  
  for(let i=1; i<rows.length; i++) {
    if(rows[i][1].toString().trim().toLowerCase() === data.numero.trim().toLowerCase() && rows[i][7] !== 'Anulada') {
      return { success: false, message: 'El número de factura ya existe y está activo.' };
    }
  }
  
  const id = 'FAC-' + Date.now();
  sheet.appendRow([id, data.numero, data.fecha, data.descripcion, data.cliente, data.observacion, new Date(), 'Activa']);
  registrarLog(data.userEmail, "CREAR_FACTURA", `Factura registrada: N° ${data.numero}`);
  return { success: true, message: 'Factura registrada exitosamente.' };
}

function updateFactura(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_FACTURAS);
  const rows = sheet.getDataRange().getValues();
  
  for(let i=1; i<rows.length; i++) {
    if(rows[i][0] !== data.id && rows[i][1].toString().trim().toLowerCase() === data.numero.trim().toLowerCase() && rows[i][7] !== 'Anulada') {
      return { success: false, message: 'El número de factura ya está en uso por otra factura activa.' };
    }
  }

  for(let i=1; i<rows.length; i++) {
    if(rows[i][0] === data.id) {
      sheet.getRange(i+1, 2).setValue(data.numero);
      sheet.getRange(i+1, 3).setValue(data.fecha);
      sheet.getRange(i+1, 4).setValue(data.descripcion);
      sheet.getRange(i+1, 5).setValue(data.cliente);
      sheet.getRange(i+1, 6).setValue(data.observacion);
      registrarLog(data.userEmail, "EDITAR_FACTURA", `Factura actualizada: N° ${data.numero}`);
      return { success: true, message: 'Datos de la factura actualizados.' };
    }
  }
  return { success: false, message: 'Factura no encontrada.' };
}

function deleteFactura(id, userEmail) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_FACTURAS);
  const rows = sheet.getDataRange().getValues();
  for(let i=1; i<rows.length; i++) {
    if(rows[i][0] === id) {
      sheet.getRange(i+1, 8).setValue('Anulada');
      registrarLog(userEmail, "ANULAR_FACTURA", `Factura anulada: ID ${id}`);
      return { success: true, message: 'Factura anulada correctamente. Sus movimientos se conservarán.' };
    }
  }
  return { success: false, message: 'Factura no encontrada.' };
}