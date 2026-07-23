/**
 * Matriz.gs
 * Cobertura por transportadora/distribuidor y destino (departamento/ciudad),
 * con frecuencia de despacho, tipo de trayecto, tiempo de entrega, tipo de
 * entrega, flete y sobreflete (%). Es la fuente de datos del Cotizador.
 *
 * La carga masiva inicial se hace importando un CSV directamente en la hoja
 * (no vía Apps Script), por volumen de datos. Estas funciones sirven para
 * consulta desde la app y para altas/ediciones puntuales posteriores.
 *
 * Lectura: cualquier usuario autenticado (el Cotizador la necesita).
 * Escritura: solo Administrador.
 */

const HOJA_MATRIZ = 'Matriz';

const TRANSPORTADORAS_DISTRIBUIDOR = ['Distribuidor', 'Servientrega', 'Coordinadora', 'TCC', 'Interrapidísimo'];
const TIPOS_ENTREGA_EN = ['Domicilio', 'Oficina', 'Domicilio y Oficina'];

function getMatrizSheet_() {
  return SpreadsheetApp.openById(getSheetId_()).getSheetByName(HOJA_MATRIZ);
}

function listarMatriz(token) {
  requiereRol_(token, null); // cualquier rol autenticado puede leer

  const sheet = getMatrizSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idx = {
    id: headers.indexOf('id_matriz'),
    transportadora: headers.indexOf('transportadora_distribuidor'),
    departamento: headers.indexOf('departamento'),
    ciudad: headers.indexOf('ciudad'),
    frecuencia: headers.indexOf('frecuencia_entrega'),
    trayecto: headers.indexOf('tipo_trayecto'),
    tiempo: headers.indexOf('tiempo_entrega'),
    entregaEn: headers.indexOf('entrega_en'),
    flete: headers.indexOf('flete'),
    sobreflete: headers.indexOf('sobreflete'),
    restriccion: headers.indexOf('restriccion'),
    activo: headers.indexOf('activo')
  };

  const filas = [];
  for (let i = 1; i < data.length; i++) {
    const f = data[i];
    filas.push({
      idMatriz: f[idx.id],
      transportadoraDistribuidor: f[idx.transportadora],
      departamento: f[idx.departamento],
      ciudad: f[idx.ciudad],
      frecuenciaEntrega: f[idx.frecuencia],
      tipoTrayecto: f[idx.trayecto],
      tiempoEntrega: f[idx.tiempo],
      entregaEn: f[idx.entregaEn],
      flete: f[idx.flete],
      sobreflete: f[idx.sobreflete],
      restriccion: f[idx.restriccion],
      activo: !!f[idx.activo]
    });
  }
  return filas;
}

function crearCobertura(token, datos) {
  requiereRol_(token, [ROLES.ADMINISTRADOR]);

  if (!datos || !datos.transportadoraDistribuidor || !datos.departamento || !datos.ciudad ||
      !datos.entregaEn || datos.tiempoEntrega === '' || datos.tiempoEntrega === undefined) {
    throw new Error('Transportadora/Distribuidor, departamento, ciudad, tipo de entrega y tiempo de entrega son obligatorios.');
  }
  if (TRANSPORTADORAS_DISTRIBUIDOR.indexOf(datos.transportadoraDistribuidor) === -1) {
    throw new Error('Transportadora/Distribuidor no válido.');
  }
  if (TIPOS_ENTREGA_EN.indexOf(datos.entregaEn) === -1) {
    throw new Error('Tipo de entrega no válido.');
  }
  const tiempo = Number(datos.tiempoEntrega);
  if (isNaN(tiempo) || tiempo < 0) {
    throw new Error('El tiempo de entrega debe ser un número válido.');
  }
  const flete = datos.flete === '' || datos.flete === undefined ? 0 : Number(datos.flete);
  if (isNaN(flete) || flete < 0) {
    throw new Error('El flete debe ser un número válido.');
  }
  const sobreflete = datos.sobreflete === '' || datos.sobreflete === undefined ? 0 : Number(datos.sobreflete);
  if (isNaN(sobreflete) || sobreflete < 0) {
    throw new Error('El sobreflete debe ser un número válido (ej. 0.035 para 3.5%).');
  }

  const sheet = getMatrizSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idx = {
    id: headers.indexOf('id_matriz'),
    transportadora: headers.indexOf('transportadora_distribuidor'),
    departamento: headers.indexOf('departamento'),
    ciudad: headers.indexOf('ciudad'),
    activo: headers.indexOf('activo')
  };

  const claveNueva = normalizarClave_(datos.transportadoraDistribuidor, datos.departamento, datos.ciudad);
  let maxId = 0;

  for (let i = 1; i < data.length; i++) {
    const f = data[i];
    const idNum = parseInt(f[idx.id], 10);
    if (!isNaN(idNum) && idNum > maxId) maxId = idNum;

    if (f[idx.activo]) {
      const claveExistente = normalizarClave_(f[idx.transportadora], f[idx.departamento], f[idx.ciudad]);
      if (claveExistente === claveNueva) {
        throw new Error('Ya existe una cobertura activa para esa transportadora/distribuidor y ese destino.');
      }
    }
  }

  const nuevoId = maxId + 1;
  const fila = headers.map(function (h) {
    switch (h) {
      case 'id_matriz': return nuevoId;
      case 'transportadora_distribuidor': return datos.transportadoraDistribuidor;
      case 'departamento': return datos.departamento;
      case 'ciudad': return datos.ciudad;
      case 'frecuencia_entrega': return datos.frecuenciaEntrega || '';
      case 'tipo_trayecto': return datos.tipoTrayecto || '';
      case 'tiempo_entrega': return tiempo;
      case 'entrega_en': return datos.entregaEn;
      case 'flete': return flete;
      case 'sobreflete': return sobreflete;
      case 'restriccion': return datos.restriccion || '';
      case 'activo': return true;
      default: return '';
    }
  });

  sheet.appendRow(fila);
  return { ok: true, idMatriz: nuevoId };
}

function actualizarCobertura(token, idMatriz, datos) {
  requiereRol_(token, [ROLES.ADMINISTRADOR]);

  if (!datos || !datos.entregaEn || datos.tiempoEntrega === '' || datos.tiempoEntrega === undefined) {
    throw new Error('Tipo de entrega y tiempo de entrega son obligatorios.');
  }
  if (TIPOS_ENTREGA_EN.indexOf(datos.entregaEn) === -1) {
    throw new Error('Tipo de entrega no válido.');
  }
  const tiempo = Number(datos.tiempoEntrega);
  if (isNaN(tiempo) || tiempo < 0) {
    throw new Error('El tiempo de entrega debe ser un número válido.');
  }
  const flete = datos.flete === '' || datos.flete === undefined ? 0 : Number(datos.flete);
  if (isNaN(flete) || flete < 0) {
    throw new Error('El flete debe ser un número válido.');
  }
  const sobreflete = datos.sobreflete === '' || datos.sobreflete === undefined ? 0 : Number(datos.sobreflete);
  if (isNaN(sobreflete) || sobreflete < 0) {
    throw new Error('El sobreflete debe ser un número válido (ej. 0.035 para 3.5%).');
  }

  const sheet = getMatrizSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxId = headers.indexOf('id_matriz');
  const idxFrecuencia = headers.indexOf('frecuencia_entrega');
  const idxTiempo = headers.indexOf('tiempo_entrega');
  const idxEntregaEn = headers.indexOf('entrega_en');
  const idxFlete = headers.indexOf('flete');
  const idxSobreflete = headers.indexOf('sobreflete');
  const idxRestriccion = headers.indexOf('restriccion');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === String(idMatriz)) {
      if (datos.frecuenciaEntrega !== undefined) sheet.getRange(i + 1, idxFrecuencia + 1).setValue(datos.frecuenciaEntrega);
      sheet.getRange(i + 1, idxTiempo + 1).setValue(tiempo);
      sheet.getRange(i + 1, idxEntregaEn + 1).setValue(datos.entregaEn);
      sheet.getRange(i + 1, idxFlete + 1).setValue(flete);
      sheet.getRange(i + 1, idxSobreflete + 1).setValue(sobreflete);
      if (datos.restriccion !== undefined) sheet.getRange(i + 1, idxRestriccion + 1).setValue(datos.restriccion);
      return { ok: true };
    }
  }
  throw new Error('Registro de cobertura no encontrado.');
}

function cambiarEstadoCobertura(token, idMatriz, activo) {
  requiereRol_(token, [ROLES.ADMINISTRADOR]);

  const sheet = getMatrizSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxId = headers.indexOf('id_matriz');
  const idxActivo = headers.indexOf('activo');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === String(idMatriz)) {
      sheet.getRange(i + 1, idxActivo + 1).setValue(activo);
      return { ok: true };
    }
  }
  throw new Error('Registro de cobertura no encontrado.');
}

function normalizarClave_(transportadora, departamento, ciudad) {
  return [transportadora, departamento, ciudad].map(function (v) {
    return String(v).trim().toUpperCase();
  }).join('|');
}
