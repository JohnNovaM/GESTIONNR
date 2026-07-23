/**
 * Recepcion.gs
 *
 * Formulario abierto: cualquier usuario autenticado puede diligenciarlo
 * (celular, cliente sí/no, ciudad, recepcionado por, cliente de, asesor).
 * Único campo obligatorio: celular. La fecha se pone sola.
 *
 * Una vez enviado, solo lo ve el asesor que quedó en el campo "asesor"
 * (y Coordinador/Administrador, que ven todo y pueden reasignarlo cuando
 * quieran — sin restricción de cuántas veces).
 *
 * Requiere una hoja "Recepcion" con columnas:
 * id_registro | fecha | celular | cliente | ciudad | recepcionado_por | cliente_de | asesor
 */

const HOJA_RECEPCION = 'Recepcion';

function getRecepcionSheet_() {
  return SpreadsheetApp.openById(getSheetId_()).getSheetByName(HOJA_RECEPCION);
}

function listarRecepcion(token) {
  const sesion = requiereRol_(token, [ROLES.ASESOR, ROLES.COORDINADOR, ROLES.ADMINISTRADOR]);

  const sheet = getRecepcionSheet_();
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });

  const soloAsesor = sesion.rol === ROLES.ASESOR ? sesion.usuario : null;

  const registros = [];
  for (let i = 1; i < data.length; i++) {
    const f = data[i];
    const asesor = f[idx['asesor']];
    if (soloAsesor && asesor !== soloAsesor) continue;

    const fechaCelda = f[idx['fecha']];
    registros.push({
      idRegistro: f[idx['id_registro']],
      fecha: fechaCelda instanceof Date ? fechaCelda.toISOString() : String(fechaCelda || ''),
      celular: f[idx['celular']],
      nombreCliente: f[idx['nombre_cliente']],
      cliente: f[idx['cliente']],
      ciudad: f[idx['ciudad']],
      recepcionadoPor: f[idx['recepcionado_por']],
      clienteDe: f[idx['cliente_de']],
      asesor: asesor
    });
  }

  registros.sort(function (a, b) { return new Date(b.fecha) - new Date(a.fecha); });
  return registros;
}

function crearRegistroRecepcion(token, datos) {
  const sesion = requiereRol_(token, null); // cualquier rol autenticado

  if (!datos || !datos.celular || !String(datos.celular).trim()) {
    throw new Error('El celular es obligatorio.');
  }

  const sheet = getRecepcionSheet_();
  if (!sheet) throw new Error('No se encontró la hoja Recepcion.');

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxId = headers.indexOf('id_registro');

  let maxId = 0;
  for (let i = 1; i < data.length; i++) {
    const n = parseInt(data[i][idxId], 10);
    if (!isNaN(n) && n > maxId) maxId = n;
  }

  const fila = headers.map(function (h) {
    switch (h) {
      case 'id_registro': return maxId + 1;
      case 'fecha': return new Date();
      case 'celular': return String(datos.celular).trim();
      case 'nombre_cliente': return datos.nombreCliente || '';
      case 'cliente': return datos.cliente || '';
      case 'ciudad': return datos.ciudad || '';
      case 'recepcionado_por': return datos.recepcionadoPor || sesion.usuario;
      case 'cliente_de': return datos.clienteDe || '';
      case 'asesor': return datos.asesor || '';
      default: return '';
    }
  });

  sheet.appendRow(fila);
  return { ok: true, idRegistro: maxId + 1 };
}

/** Asignar o reasignar el asesor de un registro, sin restricción de veces. */
function reasignarAsesorRecepcion(token, idRegistro, nuevoAsesor) {
  requiereRol_(token, [ROLES.ADMINISTRADOR, ROLES.COORDINADOR]);

  const sheet = getRecepcionSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxId = headers.indexOf('id_registro');
  const idxAsesor = headers.indexOf('asesor');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === String(idRegistro)) {
      sheet.getRange(i + 1, idxAsesor + 1).setValue(nuevoAsesor || '');
      return { ok: true };
    }
  }
  throw new Error('Registro no encontrado.');
}
