/**
 * Catalogo.gs
 * Catálogo de productos individuales — es la fuente de precios para el
 * Cotizador (armado libre de kits: producto + cantidad).
 *
 * Lectura: cualquier usuario autenticado (el Cotizador la necesita).
 * Escritura (crear, editar precio, activar/desactivar): solo Administrador.
 */

const HOJA_CATALOGO = 'Catalogo_Articulos';

function getCatalogoSheet_() {
  return SpreadsheetApp.openById(getSheetId_()).getSheetByName(HOJA_CATALOGO);
}

function listarCatalogo(token) {
  requiereRol_(token, null); // cualquier rol autenticado puede leer

  const sheet = getCatalogoSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idxId = headers.indexOf('id_articulo');
  const idxNombre = headers.indexOf('nombre');
  const idxPrecio = headers.indexOf('precio_unitario');
  const idxActivo = headers.indexOf('activo');

  const articulos = [];
  for (let i = 1; i < data.length; i++) {
    articulos.push({
      idArticulo: data[i][idxId],
      nombre: data[i][idxNombre],
      precioUnitario: data[i][idxPrecio],
      activo: !!data[i][idxActivo]
    });
  }
  return articulos;
}

function crearArticulo(token, datos) {
  requiereRol_(token, [ROLES.ADMINISTRADOR]);

  if (!datos || !datos.nombre || datos.precioUnitario === undefined || datos.precioUnitario === '') {
    throw new Error('Nombre y precio son obligatorios.');
  }
  const precio = Number(datos.precioUnitario);
  if (isNaN(precio) || precio < 0) {
    throw new Error('El precio debe ser un número válido.');
  }

  const sheet = getCatalogoSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxId = headers.indexOf('id_articulo');
  const idxNombre = headers.indexOf('nombre');

  const nombreNuevo = String(datos.nombre).trim().toUpperCase();
  let maxId = 0;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxNombre]).trim().toUpperCase() === nombreNuevo) {
      throw new Error('Ya existe un producto con ese nombre.');
    }
    const idNum = parseInt(data[i][idxId], 10);
    if (!isNaN(idNum) && idNum > maxId) maxId = idNum;
  }

  const nuevoId = maxId + 1;
  const fila = headers.map(function (h) {
    switch (h) {
      case 'id_articulo': return nuevoId;
      case 'nombre': return datos.nombre;
      case 'precio_unitario': return precio;
      case 'activo': return true;
      case 'fecha_actualizacion': return new Date();
      default: return '';
    }
  });

  sheet.appendRow(fila);
  return { ok: true, idArticulo: nuevoId };
}

function actualizarPrecioArticulo(token, idArticulo, nuevoPrecio) {
  requiereRol_(token, [ROLES.ADMINISTRADOR]);

  const precio = Number(nuevoPrecio);
  if (isNaN(precio) || precio < 0) {
    throw new Error('El precio debe ser un número válido.');
  }

  const sheet = getCatalogoSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxId = headers.indexOf('id_articulo');
  const idxPrecio = headers.indexOf('precio_unitario');
  const idxFecha = headers.indexOf('fecha_actualizacion');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === String(idArticulo)) {
      sheet.getRange(i + 1, idxPrecio + 1).setValue(precio);
      sheet.getRange(i + 1, idxFecha + 1).setValue(new Date());
      return { ok: true };
    }
  }
  throw new Error('Producto no encontrado.');
}

function cambiarEstadoArticulo(token, idArticulo, activo) {
  requiereRol_(token, [ROLES.ADMINISTRADOR]);

  const sheet = getCatalogoSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxId = headers.indexOf('id_articulo');
  const idxActivo = headers.indexOf('activo');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === String(idArticulo)) {
      sheet.getRange(i + 1, idxActivo + 1).setValue(activo);
      return { ok: true };
    }
  }
  throw new Error('Producto no encontrado.');
}

/**
 * Carga inicial con los 9 productos reales. Correr UNA VEZ desde el editor
 * (seleccionar "cargarCatalogoInicial" en el desplegable y ▶ Ejecutar).
 * Si la hoja ya tiene datos, no hace nada — evita duplicar si se corre 2 veces.
 * Andronal Forte queda pendiente (falta su precio); se agrega luego desde
 * la pantalla de Catálogo cuando lo tengan.
 */
function cargarCatalogoInicial() {
  const sheet = getCatalogoSheet_();
  const data = sheet.getDataRange().getValues();
  if (data.length > 1) {
    Logger.log('El catálogo ya tiene datos — no se cargó nada, para evitar duplicados.');
    return;
  }

  const productos = [
    ['ANDRONAL', 100000],
    ['HARPAGOFITO', 100000],
    ['HYDROCOLLAG', 20000],
    ['HYDRO 250', 10000],
    ['HYDRO 500', 10000],
    ['GELOSALIN 120', 20000],
    ['ANDROGEL', 20000],
    ['RESVERATROL 450GR', 100000],
    ['RESVERATROL LT', 20000]
  ];

  productos.forEach(function (p, i) {
    sheet.appendRow([i + 1, p[0], p[1], true, new Date()]);
  });

  Logger.log('Catálogo cargado con ' + productos.length + ' productos.');
}
