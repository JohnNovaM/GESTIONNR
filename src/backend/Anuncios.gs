/**
 * Anuncios.gs
 * Avisos cortos que el Administrador publica (ej. "Ya subí los pagos de la
 * quincena", "Ya pueden generar cuentas de cobro"). Aparecen como banner
 * arriba de la app para todos los que inicien sesión mientras estén activos.
 *
 * Requiere una hoja "Anuncios" con columnas:
 * id_anuncio | mensaje | fecha_creacion | activo | creado_por
 */

const HOJA_ANUNCIOS = 'Anuncios';

function getAnunciosSheet_() {
  return SpreadsheetApp.openById(getSheetId_()).getSheetByName(HOJA_ANUNCIOS);
}

/** Cualquier usuario autenticado puede ver los anuncios activos. */
function listarAnunciosActivos(token) {
  requiereRol_(token, null);

  const sheet = getAnunciosSheet_();
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];

  const idxId = headers.indexOf('id_anuncio');
  const idxMensaje = headers.indexOf('mensaje');
  const idxFecha = headers.indexOf('fecha_creacion');
  const idxActivo = headers.indexOf('activo');

  const anuncios = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][idxActivo]) continue;
    const fechaCelda = data[i][idxFecha];
    anuncios.push({
      idAnuncio: data[i][idxId],
      mensaje: data[i][idxMensaje],
      fecha: fechaCelda instanceof Date ? fechaCelda.toISOString() : String(fechaCelda || '')
    });
  }

  // más recientes primero
  anuncios.sort(function (a, b) { return new Date(b.fecha) - new Date(a.fecha); });
  return anuncios;
}

/** Solo Administrador publica anuncios nuevos. */
function crearAnuncio(token, mensaje) {
  const sesion = requiereRol_(token, [ROLES.ADMINISTRADOR]);
  if (!mensaje || !String(mensaje).trim()) {
    throw new Error('El mensaje no puede estar vacío.');
  }

  const sheet = getAnunciosSheet_();
  if (!sheet) throw new Error('No se encontró la hoja Anuncios.');

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxId = headers.indexOf('id_anuncio');

  let maxId = 0;
  for (let i = 1; i < data.length; i++) {
    const n = parseInt(data[i][idxId], 10);
    if (!isNaN(n) && n > maxId) maxId = n;
  }

  const fila = headers.map(function (h) {
    switch (h) {
      case 'id_anuncio': return maxId + 1;
      case 'mensaje': return String(mensaje).trim();
      case 'fecha_creacion': return new Date();
      case 'activo': return true;
      case 'creado_por': return sesion.usuario;
      default: return '';
    }
  });

  sheet.appendRow(fila);
  return { ok: true, idAnuncio: maxId + 1 };
}

/** Solo Administrador puede listar todos (para gestionarlos) y desactivarlos. */
function listarAnunciosTodos(token) {
  requiereRol_(token, [ROLES.ADMINISTRADOR]);

  const sheet = getAnunciosSheet_();
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];

  const idxId = headers.indexOf('id_anuncio');
  const idxMensaje = headers.indexOf('mensaje');
  const idxFecha = headers.indexOf('fecha_creacion');
  const idxActivo = headers.indexOf('activo');

  const anuncios = [];
  for (let i = 1; i < data.length; i++) {
    const fechaCelda = data[i][idxFecha];
    anuncios.push({
      idAnuncio: data[i][idxId],
      mensaje: data[i][idxMensaje],
      fecha: fechaCelda instanceof Date ? fechaCelda.toISOString() : String(fechaCelda || ''),
      activo: !!data[i][idxActivo]
    });
  }
  anuncios.sort(function (a, b) { return new Date(b.fecha) - new Date(a.fecha); });
  return anuncios;
}

function desactivarAnuncio(token, idAnuncio) {
  requiereRol_(token, [ROLES.ADMINISTRADOR]);

  const sheet = getAnunciosSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxId = headers.indexOf('id_anuncio');
  const idxActivo = headers.indexOf('activo');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === String(idAnuncio)) {
      sheet.getRange(i + 1, idxActivo + 1).setValue(false);
      return { ok: true };
    }
  }
  throw new Error('Anuncio no encontrado.');
}
