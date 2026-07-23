/**
 * Cotizador.gs
 * El armado del kit y el cálculo del valor sugerido se hacen en el cliente
 * (a partir de listarCatalogo, ya existente). Aquí solo lo que necesita
 * datos frescos del servidor: departamentos/ciudades con cobertura, y el
 * cálculo de opciones de envío una vez el asesor confirma el valor de venta.
 *
 * Disponible para cualquier rol autenticado.
 */

/**
 * Quita acentos y normaliza mayúsculas/espacios, para poder comparar y
 * agrupar "Fusagasugá" con "Fusagasuga" como el mismo destino, sin importar
 * cómo lo haya escrito cada transportadora en el archivo original.
 */
function normalizarTexto_(v) {
  return String(v || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
}

function tieneAcento_(v) {
  const original = String(v || '').trim();
  return original !== original.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function listarDepartamentos(token) {
  requiereRol_(token, null);

  const sheet = getMatrizSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxDepto = headers.indexOf('departamento');
  const idxActivo = headers.indexOf('activo');

  const mapa = {}; // clave normalizada -> mejor texto para mostrar
  for (let i = 1; i < data.length; i++) {
    const valor = data[i][idxDepto];
    if (!data[i][idxActivo] || !valor) continue;
    const clave = normalizarTexto_(valor);
    if (!mapa[clave] || (tieneAcento_(valor) && !tieneAcento_(mapa[clave]))) {
      mapa[clave] = valor;
    }
  }
  return Object.keys(mapa).map(function (k) { return mapa[k]; }).sort();
}

function listarCiudadesPorDepartamento(token, departamento) {
  requiereRol_(token, null);

  const sheet = getMatrizSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxDepto = headers.indexOf('departamento');
  const idxCiudad = headers.indexOf('ciudad');
  const idxActivo = headers.indexOf('activo');
  const claveDepto = normalizarTexto_(departamento);

  const mapa = {};
  for (let i = 1; i < data.length; i++) {
    const ciudad = data[i][idxCiudad];
    if (!data[i][idxActivo] || !ciudad) continue;
    if (normalizarTexto_(data[i][idxDepto]) !== claveDepto) continue;

    const clave = normalizarTexto_(ciudad);
    if (!mapa[clave] || (tieneAcento_(ciudad) && !tieneAcento_(mapa[clave]))) {
      mapa[clave] = ciudad;
    }
  }
  return Object.keys(mapa).map(function (k) { return mapa[k]; }).sort();
}

/**
 * Para el departamento/ciudad elegidos, arma una opción por cada transportadora
 * o distribuidor con cobertura activa ahí. El valor de envío = flete fijo +
 * (valorReal de la venta * % de sobreflete de esa transportadora).
 */
function cotizarEnvio(token, departamento, ciudad, valorReal) {
  requiereRol_(token, null);

  const valor = Number(valorReal);
  if (isNaN(valor) || valor < 0) {
    throw new Error('El valor de la venta no es válido.');
  }
  if (!departamento || !ciudad) {
    throw new Error('Departamento y ciudad son obligatorios.');
  }

  const sheet = getMatrizSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idx = {
    id: headers.indexOf('id_matriz'),
    transportadora: headers.indexOf('transportadora_distribuidor'),
    departamento: headers.indexOf('departamento'),
    ciudad: headers.indexOf('ciudad'),
    frecuencia: headers.indexOf('frecuencia_entrega'),
    tiempo: headers.indexOf('tiempo_entrega'),
    entregaEn: headers.indexOf('entrega_en'),
    flete: headers.indexOf('flete'),
    sobreflete: headers.indexOf('sobreflete'),
    restriccion: headers.indexOf('restriccion'),
    activo: headers.indexOf('activo')
  };

  const opciones = [];
  const claveDepto = normalizarTexto_(departamento);
  const claveCiudad = normalizarTexto_(ciudad);

  for (let i = 1; i < data.length; i++) {
    const f = data[i];
    if (!f[idx.activo]) continue;
    if (normalizarTexto_(f[idx.departamento]) !== claveDepto || normalizarTexto_(f[idx.ciudad]) !== claveCiudad) continue;

    const flete = Number(f[idx.flete]) || 0;
    const sobreflete = Number(f[idx.sobreflete]) || 0;
    const valorSobreflete = Math.round(valor * sobreflete);
    const valorEnvio = flete + valorSobreflete;

    opciones.push({
      idMatriz: f[idx.id],
      transportadoraDistribuidor: f[idx.transportadora],
      frecuenciaEntrega: f[idx.frecuencia],
      tiempoEntrega: f[idx.tiempo],
      entregaEn: f[idx.entregaEn],
      restriccion: f[idx.restriccion],
      flete: flete,
      sobreflete: sobreflete,
      valorSobreflete: valorSobreflete,
      valorEnvio: valorEnvio
    });
  }

  opciones.sort(function (a, b) { return a.tiempoEntrega - b.tiempoEntrega; });
  return opciones;
}
