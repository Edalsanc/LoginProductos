// Nombre de la pestaña en la hoja de cálculo donde se guardarán los datos
const SHEET_NAME = 'Usuarios';
const SSID = '1u4_PXoCO_Jh8KP9xHPH7Nam7T4qTy9KfOGYyaNtlkrE';

/**
 * Clave secreta maestra para encriptar y desencriptar.
 * Puedes personalizarla con tu propia frase secreta.
 */
const SECRET_KEY = 'Murcielago2026';

/**
 * Genera un ID único para el usuario con prefijo "USR-" y formato UUID.
 * Ejemplo: USR-4A7F9B12
 * @return {string} - ID único generado.
 */
function generateUserId() {
  return 'USR-' + Utilities.getUuid().substring(0, 8).toUpperCase();
}

/**
 * Encripta una contraseña usando cifrado simétrico con sal (salt) y SHA-256.
 * Genera un texto cifrado diferente cada vez incluso para la misma contraseña.
 * @param {string} text - Contraseña en texto plano.
 * @return {string} - Contraseña cifrada en formato "salt:base64".
 */
function encryptPassword(text) {
  if (!text) return '';
  // Vector de inicialización / salt único de 8 caracteres
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

/**
 * Desencripta una contraseña cifrada por encryptPassword.
 * Si la contraseña proviene de una cuenta previa (sin formato salt:base64), la retorna tal cual.
 * @param {string} cipherString - Contraseña cifrada guardada en la hoja.
 * @return {string} - Contraseña original en texto plano.
 */
function decryptPassword(cipherString) {
  if (!cipherString) return '';
  const str = cipherString.toString();
  const parts = str.split(':');
  
  // Soporte retrocompatible para contraseñas sin cifrar previas
  if (parts.length !== 2) {
    return str;
  }
  
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

/**
 * Muestra la interfaz gráfica (HTML) cuando se visita la URL web en el navegador
 */
function doGet(e) {
  var page = e.parameter.p || 'index';
  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle('Sistema de Autenticación')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
} 

/**
 * Esta función referncia las los archivos declarados como html
 * */
function include(filename) {

  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * Maneja el registro de nuevos usuarios
 */
function handleRegister(params) {
  const { name, email, password } = params;
  
  // Validar campos requeridos
  if (!name || !email || !password) {
    return { success: false, message: 'Faltan campos obligatorios.' };
  }

  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const cols = getColumnIndices(data[0]);
  
  // Verificar si el correo ya existe
  for (let i = 1; i < data.length; i++) {
    const storedEmail = (data[i][cols.email] || '').toString().toLowerCase();
    if (storedEmail === email.toLowerCase()) {
      return { success: false, message: 'Este correo electrónico ya está registrado.' };
    }
  }
  
  // Generar ID único y encriptar contraseña
  const userId = generateUserId();
  const encryptedPassword = encryptPassword(password);
  
  // Guardar fila: [ID, Fecha de Registro, Nombre, Email, Contraseña Cifrada]
  sheet.appendRow([userId, new Date(), name, email.toLowerCase(), encryptedPassword]);
  
  return { 
    success: true, 
    message: 'Registro exitoso.',
    user: {
      id: userId,
      name: name,
      email: email.toLowerCase()
    }
  };
}

/**
 * Maneja el inicio de sesión
 */
function handleLogin(params) {
  const { email, password } = params;
  
  if (!email || !password) {
    return { success: false, message: 'Debes ingresar correo y contraseña.' };
  }

  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const cols = getColumnIndices(data[0]);
  
  // Buscar el correo en los registros
  for (let i = 1; i < data.length; i++) {
    const storedEmail = (data[i][cols.email] || '').toString().toLowerCase();
    
    if (storedEmail === email.toLowerCase()) {
      const storedPassword = data[i][cols.password];
      const decryptedStoredPassword = decryptPassword(storedPassword);
      
      if (decryptedStoredPassword === password) {
        return { 
          success: true, 
          message: 'Login exitoso.',
          user: {
            id: data[i][cols.id] || 'N/A',
            name: data[i][cols.name] || '',
            email: storedEmail
          }
        };
      } else {
        return { success: false, message: 'Contraseña incorrecta.' };
      }
    }
  }
  
  return { success: false, message: 'El usuario no existe. Por favor, regístrate.' };
}

/**
 * Detecta dinámicamente las posiciones de las columnas por su cabecera.
 */
function getColumnIndices(headerRow) {
  const headers = (headerRow || []).map(h => (h || '').toString().toLowerCase().trim());
  
  return {
    id: headers.indexOf('id') !== -1 ? headers.indexOf('id') : 0,
    date: headers.indexOf('fecha de registro') !== -1 ? headers.indexOf('fecha de registro') : 1,
    name: headers.indexOf('nombre') !== -1 ? headers.indexOf('nombre') : 2,
    email: headers.indexOf('email') !== -1 ? headers.indexOf('email') : 3,
    password: headers.indexOf('contraseña') !== -1 ? headers.indexOf('contraseña') : 4
  };
}

/**
 * Obtiene la pestaña de la hoja de cálculo o la crea con sus encabezados si no existe.
 */
function getSheet() {
  const ss = SpreadsheetApp.openById(SSID);
  //const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['ID', 'Fecha de Registro', 'Nombre', 'Email', 'Contraseña']);
    
    const headerRange = sheet.getRange('A1:E1');
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f4f6');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 5);
  } else {
    // Si la hoja ya existía sin columna ID, se añade al inicio de forma transparente
    const lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      const hasIdCol = headers.some(h => (h || '').toString().toLowerCase().trim() === 'id');
      
      if (!hasIdCol) {
        sheet.insertColumnBefore(1);
        const idHeaderCell = sheet.getRange(1, 1);
        idHeaderCell.setValue('ID').setFontWeight('bold').setBackground('#f3f4f6');
        sheet.autoResizeColumns(1, sheet.getLastColumn());
      }
    }
  }
  
  return sheet;
}