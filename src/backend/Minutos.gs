/**
 * Minutos.gs
 *
 * Cada asesor reporta minutos PBX y minutos celular por día — ningún campo
 * es obligatorio (puede reportar solo uno de los dos).
 *
 * Minutos celular: solo queda como historial, sin ningún descuento.
 *
 * Minutos PBX: Coordinador/Administrador manejan una "bolsa" — un total de
 * minutos disponibles desde una fecha de inicio de ciclo. Lo consumido se
 * calcula sumando los minutos_pbx reportados desde esa fecha en adelante.
 * El reinicio es manual: Coordinador/Administrador editan la bolsa total
 * y/o la fecha de inicio cuando quieran (normalmente el día 5, pero no es
 * automático).
 *
 * Requiere dos hojas:
 * "Minutos": id_registro | fecha | asesor | minutos_pbx | minutos_celular
 * "Minutos_Config": bolsa_total | fecha_inicio_ciclo   (una sola fila de datos)
 */

const HOJA_MINUTOS = 'Minutos';
const HOJA_MINUTOS_CONFIG = 'Minutos_Config';

function getMinutosSheet_() {
  return SpreadsheetApp.openById(getSheetId_()).getSheetByName(HOJA_MINUTOS);
}

function getMinutosConfigSheet_() {
  return SpreadsheetApp.openById(getSheetId_()).getSheetByName(HOJA_MINUTOS_CONFIG);
}

function listarMinutos(token) {
  const sesion = requiereRol_(token, [ROLES.ASESOR, ROLES.COORDINADOR, ROLES.ADMINISTRADOR]);

  const sheet = getMinutosSheet_();
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
      asesor: asesor,
      minutosPbx: f[idx['minutos_pbx']],
      minutosCelular: f[idx['minutos_celular']]
    });
  }

  registros.sort(function (a, b) { return new Date(b.fecha) - new Date(a.fecha); });
  return registros;
}

function crearRegistroMinutos(token, datos) {
  const sesion = requiereRol_(token, null);

  const sheet = getMinutosSheet_();
  if (!sheet) throw new Error('No se encontró la hoja Minutos.');

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
      case 'asesor': return (datos && datos.asesor) || sesion.usuario;
      case 'minutos_pbx':
        return datos && datos.minutosPbx !== '' && datos.minutosPbx !== undefined ? Number(datos.minutosPbx) : '';
      case 'minutos_celular':
        return datos && datos.minutosCelular !== '' && datos.minutosCelular !== undefined ? Number(datos.minutosCelular) : '';
      default: return '';
    }
  });

  sheet.appendRow(fila);
  return { ok: true };
}

// ==================================================
// Bolsa de minutos PBX
// ==================================================

function leerConfigBolsa_() {
  const sheet = getMinutosConfigSheet_();
  if (!sheet) return { bolsaTotal: 0, fechaInicioCiclo: '' };

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { bolsaTotal: 0, fechaInicioCiclo: '' };

  const headers = data[0];
  const idxBolsa = headers.indexOf('bolsa_total');
  const idxFecha = headers.indexOf('fecha_inicio_ciclo');
  const fila = data[1];
  const fechaCelda = fila[idxFecha];

  return {
    bolsaTotal: Number(fila[idxBolsa]) || 0,
    fechaInicioCiclo: fechaCelda instanceof Date ? fechaCelda.toISOString() : String(fechaCelda || '')
  };
}

/** Cualquier rol puede consultar el estado de la bolsa (informativo). */
function obtenerEstadoBolsaMinutos(token) {
  requiereRol_(token, null);

  const config = leerConfigBolsa_();
  let consumido = 0;

  const sheetMinutos = getMinutosSheet_();
  if (sheetMinutos && config.fechaInicioCiclo) {
    const data = sheetMinutos.getDataRange().getValues();
    if (data.length > 1) {
      const headers = data[0];
      const idxFecha = headers.indexOf('fecha');
      const idxPbx = headers.indexOf('minutos_pbx');
      const inicio = new Date(config.fechaInicioCiclo);

      for (let i = 1; i < data.length; i++) {
        const fechaCelda = data[i][idxFecha];
        const fechaFila = fechaCelda instanceof Date ? fechaCelda : new Date(fechaCelda);
        if (!isNaN(fechaFila) && fechaFila >= inicio) {
          consumido += Number(data[i][idxPbx]) || 0;
        }
      }
    }
  }

  return {
    bolsaTotal: config.bolsaTotal,
    fechaInicioCiclo: config.fechaInicioCiclo,
    consumido: consumido,
    restante: config.bolsaTotal - consumido
  };
}

/** Coordinador/Administrador editan la bolsa total y/o reinician el ciclo cuando quieran. */
function actualizarConfigBolsaMinutos(token, bolsaTotal, fechaInicioCiclo) {
  requiereRol_(token, [ROLES.ADMINISTRADOR, ROLES.COORDINADOR]);

  const bolsa = Number(bolsaTotal);
  if (isNaN(bolsa) || bolsa < 0) {
    throw new Error('La bolsa debe ser un número válido.');
  }
  if (!fechaInicioCiclo) {
    throw new Error('La fecha de inicio de ciclo es obligatoria.');
  }

  const sheet = getMinutosConfigSheet_();
  if (!sheet) throw new Error('No se encontró la hoja Minutos_Config.');

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxBolsa = headers.indexOf('bolsa_total');
  const idxFecha = headers.indexOf('fecha_inicio_ciclo');

  if (data.length < 2) {
    const fila = headers.map(function (h) {
      if (h === 'bolsa_total') return bolsa;
      if (h === 'fecha_inicio_ciclo') return new Date(fechaInicioCiclo);
      return '';
    });
    sheet.appendRow(fila);
  } else {
    sheet.getRange(2, idxBolsa + 1).setValue(bolsa);
    sheet.getRange(2, idxFecha + 1).setValue(new Date(fechaInicioCiclo));
  }

  return { ok: true };
}
