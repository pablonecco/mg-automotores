// ==== CONFIGURACIÓN — editar antes de desplegar ====
const ADMIN_PASSWORD = 'CAMBIAR_ESTE_PASSWORD';
const SPREADSHEET_ID = '1ey2twsuVi03z733tOOezrAV84awWyGOCQetq06W0kb4';
const DRIVE_FOLDER_ID = '1IUG6aEZCjxPTT0XcjZ3ycNFm1ytPepHy';
const SHEET_NAME = 'Autos';
// =====================================================

function doPost(e) {
  var respuesta;
  try {
    var body = JSON.parse(e.postData.contents);
    var accion = body.accion;

    switch (accion) {
      case 'login':
        respuesta = { ok: true, data: body.password === ADMIN_PASSWORD };
        break;
      case 'crearAuto':
        checkPassword(body.password);
        respuesta = { ok: true, data: crearAuto(body.datos) };
        break;
      case 'editarAuto':
        checkPassword(body.password);
        respuesta = { ok: true, data: editarAuto(body.datos) };
        break;
      case 'borrarAuto':
        checkPassword(body.password);
        respuesta = { ok: true, data: borrarAuto(body.datos.id) };
        break;
      case 'subirFoto':
        checkPassword(body.password);
        respuesta = { ok: true, data: subirFoto(body.datos.base64, body.datos.nombreArchivo, body.datos.tipo) };
        break;
      default:
        respuesta = { ok: false, error: 'Acción desconocida: ' + accion };
    }
  } catch (err) {
    respuesta = { ok: false, error: err.message };
  }

  // Content-Type text/plain en el cliente evita el preflight CORS que Apps Script no maneja.
  return ContentService
    .createTextOutput(JSON.stringify(respuesta))
    .setMimeType(ContentService.MimeType.JSON);
}

function checkPassword(password) {
  if (password !== ADMIN_PASSWORD) {
    throw new Error('Password incorrecto');
  }
}

function getSheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
}

function buscarFilaPorId(sheet, id) {
  var valores = sheet.getDataRange().getValues();
  for (var i = 1; i < valores.length; i++) {
    if (valores[i][0] === id) return i;
  }
  return -1;
}

// Columnas: id, titulo, descripcion, marca, modelo, anio, km, precio, fotos, fotoPrincipal, fechaCreacion
function crearAuto(datos) {
  var sheet = getSheet();
  var id = Utilities.getUuid();
  sheet.appendRow([
    id,
    datos.titulo || '',
    datos.descripcion || '',
    datos.marca || '',
    datos.modelo || '',
    datos.anio || '',
    datos.km || '',
    datos.precio || '',
    JSON.stringify(datos.fotos || []),
    datos.fotoPrincipal || '',
    new Date().toISOString()
  ]);
  return { id: id };
}

function editarAuto(datos) {
  var sheet = getSheet();
  var filaIndex = buscarFilaPorId(sheet, datos.id);
  if (filaIndex === -1) throw new Error('Auto no encontrado: ' + datos.id);

  var fila = filaIndex + 1; // getRange es 1-indexado
  sheet.getRange(fila, 2, 1, 7).setValues([[
    datos.titulo || '',
    datos.descripcion || '',
    datos.marca || '',
    datos.modelo || '',
    datos.anio || '',
    datos.km || '',
    datos.precio || ''
  ]]);
  sheet.getRange(fila, 9, 1, 2).setValues([[
    JSON.stringify(datos.fotos || []),
    datos.fotoPrincipal || ''
  ]]);
  return { id: datos.id };
}

function borrarAuto(id) {
  var sheet = getSheet();
  var filaIndex = buscarFilaPorId(sheet, id);
  if (filaIndex === -1) throw new Error('Auto no encontrado: ' + id);
  sheet.deleteRow(filaIndex + 1);
  return { id: id };
}

function subirFoto(base64, nombreArchivo, tipo) {
  var carpeta = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var blob = Utilities.newBlob(
    Utilities.base64Decode(base64),
    tipo || 'image/jpeg',
    nombreArchivo || ('foto_' + new Date().getTime())
  );
  var archivo = carpeta.createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // El link de compartir normal de Drive no renderiza en <img>; este formato sí.
  return { url: 'https://drive.google.com/thumbnail?id=' + archivo.getId() + '&sz=w1000' };
}
