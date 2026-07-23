/**
 * Ventas.gs
 *
 * Tres hojas de venta:
 * - Ventas_Transportadora: numero_guia lo asigna Logística/Administrador.
 *   Al asignarse, el semáforo pasa a "Creado". El semáforo (Creado/Enviado/
 *   Novedad/Entregado/Caída) lo cambia después el propio asesor dueño de la
 *   venta (o Administrador), según lo que vea al rastrear.
 * - Ventas_Distribuidor: remision se asigna MANUALMENTE por el Administrador
 *   (no es consecutivo automático — ya se genera en la otra app y el número
 *   base puede cambiar). El estado de entrega no se edita aquí: se trae en
 *   vivo desde la hoja de reporte de distribuidores vía IMPORTRANGE + VLOOKUP
 *   directamente en el Sheet (columnas estado_entrega / observacion_entrega).
 * - Ventas_Consignacion: mismo contenido que Transportadora (incluye a qué
 *   transportadora/distribuidor iría el envío), pero queda con
 *   estado_aprobacion = "Pendiente" hasta que Administrador o Coordinador la
 *   aprueben. Ya aprobada, se le asigna numero_guia_remision igual que a las
 *   otras (Logística si es transportadora, Administrador si es Distribuidor),
 *   y el semáforo aplica solo si terminó yendo por una transportadora real.
 *
 * fve, pagado a, vlr fac, quincena y vlr comision se dejan en blanco en las
 * tres: llegan después por la importación del módulo de Pagos.
 */

const HOJA_VENTAS_TRANSPORTADORA = 'Ventas_Transportadora';
const HOJA_VENTAS_DISTRIBUIDOR = 'Ventas_Distribuidor';
const HOJA_VENTAS_CONSIGNACION = 'Ventas_Consignacion';
const COMISION_POR_DEFECTO = 0.11; // 11%, el asesor la puede cambiar al vender

const ESTADOS_SEMAFORO = ['Creado', 'En ruta', 'Novedad', 'Entregado', 'Caída'];
const ESTADOS_APROBACION_CONSIGNACION = ['Pendiente', 'Aprobada', 'Rechazada'];

function getVentasTransportadoraSheet_() {
  return SpreadsheetApp.openById(getSheetId_()).getSheetByName(HOJA_VENTAS_TRANSPORTADORA);
}
function getVentasDistribuidorSheet_() {
  return SpreadsheetApp.openById(getSheetId_()).getSheetByName(HOJA_VENTAS_DISTRIBUIDOR);
}
function getVentasConsignacionSheet_() {
  return SpreadsheetApp.openById(getSheetId_()).getSheetByName(HOJA_VENTAS_CONSIGNACION);
}

/**
 * items: [{ nombre: 'ANDRONAL', cantidad: 3 }, ...] tal como vienen del Catálogo.
 * Arma el texto del resumen y un mapa columna(minúsculas) -> cantidad, para
 * que cada producto caiga en su columna correspondiente de la hoja de venta.
 */
function armarResumenYColumnas_(items) {
  const partes = [];
  const columnas = {};
  (items || []).forEach(function (it) {
    const cantidad = Number(it.cantidad) || 0;
    if (cantidad > 0) {
      partes.push(cantidad + 'x ' + it.nombre);
      columnas[String(it.nombre).trim().toLowerCase()] = cantidad;
    }
  });
  return { resumen: partes.join(' + '), columnas: columnas };
}

function siguienteIdVenta_(sheet, headers) {
  const data = sheet.getDataRange().getValues();
  const idx = headers.indexOf('id_venta');
  let maxId = 0;
  for (let i = 1; i < data.length; i++) {
    const n = parseInt(data[i][idx], 10);
    if (!isNaN(n) && n > maxId) maxId = n;
  }
  return maxId + 1;
}

function validarDatosVenta_(datos) {
  if (!datos || !datos.ciudad || !datos.nombre || !datos.cedula || datos.valorReal === undefined) {
    throw new Error('Faltan datos obligatorios de la venta (ciudad, nombre, cédula o valor).');
  }
  const valorReal = Number(datos.valorReal);
  if (isNaN(valorReal) || valorReal < 0) {
    throw new Error('El valor de la venta no es válido.');
  }
  const comision = datos.comisionPorcentaje === undefined || datos.comisionPorcentaje === ''
    ? COMISION_POR_DEFECTO : Number(datos.comisionPorcentaje);
  if (isNaN(comision) || comision < 0) {
    throw new Error('La comisión no es válida.');
  }
  return { valorReal: valorReal, valorSugerido: Number(datos.valorSugerido) || 0, comision: comision };
}

/**
 * Busca una cédula en las 3 hojas de venta y devuelve los datos de la
 * compra más reciente (nombre, dirección, celular, teléfono, correo y el
 * asesor al que quedó asignada) — o null si nunca ha comprado.
 * No se restringe por asesor: cualquiera puede ver que un cliente ya tiene
 * asesor asignado, para poder devolverle la venta a quien corresponde.
 */
function buscarClientePorCedula(token, cedula) {
  requiereRol_(token, null);
  return buscarClientePorCedula_(cedula);
}

function buscarClientePorCedula_(cedula) {
  if (!cedula) return null;
  const cedulaBuscada = String(cedula).trim();
  if (!cedulaBuscada) return null;

  const candidatos = [];

  [getVentasTransportadoraSheet_, getVentasDistribuidorSheet_, getVentasConsignacionSheet_].forEach(function (obtener) {
    try {
      const sheet = obtener();
      if (!sheet) return;
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return;
      const headers = data[0];
      const idx = {};
      headers.forEach(function (h, i) { idx[h] = i; });

      for (let i = 1; i < data.length; i++) {
        const f = data[i];
        if (String(f[idx['cedula']]).trim() === cedulaBuscada) {
          const fechaCelda = f[idx['fecha']];
          candidatos.push({
            nombre: f[idx['nombre']],
            direccion: f[idx['direccion']],
            celular: f[idx['celular']],
            telefono: f[idx['telefono']],
            correoElectronico: f[idx['correo_electronico']],
            asesor: f[idx['asesor']],
            fecha: fechaCelda instanceof Date ? fechaCelda.toISOString() : String(fechaCelda || '')
          });
        }
      }
    } catch (e) {
      // si una hoja falla, seguimos con las demás en vez de tumbar la búsqueda completa
    }
  });

  if (!candidatos.length) return null;
  candidatos.sort(function (a, b) { return new Date(b.fecha) - new Date(a.fecha); });
  return candidatos[0];
}

/**
 * Decide a nombre de qué asesor queda la venta:
 * - Por defecto, quien la está subiendo (sesion.usuario)
 * - Si se pide asignarla a alguien más, SOLO se permite si es justo el
 *   asesor que ya tenía asignado ese cliente (se revalida acá, no se
 *   confía en lo que mande el navegador)
 */
function resolverAsesorVenta_(sesion, cedula, asesorSolicitado) {
  if (!asesorSolicitado || asesorSolicitado === sesion.usuario) {
    return sesion.usuario;
  }
  const cliente = buscarClientePorCedula_(cedula);
  if (!cliente || !cliente.asesor || cliente.asesor !== asesorSolicitado) {
    throw new Error('Solo puedes asignar la venta al asesor que ya tenía asignado ese cliente.');
  }
  return asesorSolicitado;
}

function encontrarFila_(sheet, headers, idVenta) {
  const data = sheet.getDataRange().getValues();
  const idxId = headers.indexOf('id_venta');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === String(idVenta)) {
      return { numeroFila: i + 1, fila: data[i] };
    }
  }
  return null;
}

// ==================================================
// Crear ventas
// ==================================================

function crearVentaTransportadora(token, datos) {
  const sesion = requiereRol_(token, null); // cualquier rol autenticado que use el Cotizador

  if (!datos || !datos.transportadora || !datos.departamento) {
    throw new Error('Falta la transportadora o el departamento.');
  }
  const v = validarDatosVenta_(datos);
  const nombreAsesor = resolverAsesorVenta_(sesion, datos.cedula, datos.asesorAsignado);

  const sheet = getVentasTransportadoraSheet_();
  const headers = sheet.getDataRange().getValues()[0];
  const nuevoId = siguienteIdVenta_(sheet, headers);
  const armado = armarResumenYColumnas_(datos.items);

  const fila = headers.map(function (h) {
    switch (h) {
      case 'id_venta': return nuevoId;
      case 'fecha': return new Date();
      case 'numero_guia': return '';
      case 'semaforo': return '';
      case 'asesor': return nombreAsesor;
      case 'subido_por': return sesion.usuario;
      case 'transportadora': return datos.transportadora;
      case 'cedula': return datos.cedula;
      case 'nombre': return datos.nombre;
      case 'departamento': return datos.departamento;
      case 'ciudad': return datos.ciudad;
      case 'direccion': return datos.direccion || '';
      case 'celular': return datos.celular || '';
      case 'telefono': return datos.telefono || '';
      case 'correo_electronico': return datos.correoElectronico || '';
      case 'observacion': return datos.observacion || '';
      case 'id_kit': return datos.idKit || 'personalizado';
      case 'resumen_kit': return armado.resumen;
      case 'valor_sugerido': return v.valorSugerido;
      case 'valor_real': return v.valorReal;
      case 'comision_(%)': return v.comision;
      case 'fve': return '';
      case 'pagado a': return '';
      case 'vlr fac': return '';
      case 'quincena': return '';
      case 'vlr comision': return '';
      default:
        return armado.columnas[h] !== undefined ? armado.columnas[h] : '';
    }
  });

  sheet.appendRow(fila);
  return { ok: true, idVenta: nuevoId };
}

function crearVentaDistribuidor(token, datos) {
  const sesion = requiereRol_(token, null);

  if (!datos || !datos.distribuidor) {
    throw new Error('Falta el distribuidor.');
  }
  const v = validarDatosVenta_(datos);
  const nombreAsesor = resolverAsesorVenta_(sesion, datos.cedula, datos.asesorAsignado);

  const sheet = getVentasDistribuidorSheet_();
  const headers = sheet.getDataRange().getValues()[0];
  const nuevoId = siguienteIdVenta_(sheet, headers);
  const armado = armarResumenYColumnas_(datos.items);

  const fila = headers.map(function (h) {
    switch (h) {
      case 'id_venta': return nuevoId;
      case 'remision': return '';
      case 'asesor': return nombreAsesor;
      case 'subido_por': return sesion.usuario;
      case 'fecha': return new Date();
      case 'cedula': return datos.cedula;
      case 'nombre': return datos.nombre;
      case 'ciudad': return datos.ciudad;
      case 'distribuidor': return datos.distribuidor;
      case 'direccion': return datos.direccion || '';
      case 'celular': return datos.celular || '';
      case 'telefono': return datos.telefono || '';
      case 'correo_electronico': return datos.correoElectronico || '';
      case 'observacion': return datos.observacion || '';
      case 'id_kit': return datos.idKit || 'personalizado';
      case 'resumen_kit': return armado.resumen;
      case 'valor_sugerido': return v.valorSugerido;
      case 'valor_real': return v.valorReal;
      case 'comision_(%)': return v.comision;
      case 'fve': return '';
      case 'pagado a': return '';
      case 'vlr fac': return '';
      case 'quincena': return '';
      case 'vlr comision': return '';
      default:
        return armado.columnas[h] !== undefined ? armado.columnas[h] : '';
    }
  });

  sheet.appendRow(fila);
  return { ok: true, idVenta: nuevoId };
}

/**
 * Consignación: guarda a qué transportadora/distribuidor iría (para saber
 * después si le toca guía o remisión, y si aplica semáforo), pero queda
 * "Pendiente" de aprobación antes de que logística/admin hagan nada más.
 */
function crearVentaConsignacion(token, datos) {
  const sesion = requiereRol_(token, null);

  if (!datos || !datos.transportadoraDistribuidor || !datos.ciudad) {
    throw new Error('Falta la transportadora/distribuidor o la ciudad.');
  }
  const v = validarDatosVenta_(datos);
  const nombreAsesor = resolverAsesorVenta_(sesion, datos.cedula, datos.asesorAsignado);

  const sheet = getVentasConsignacionSheet_();
  const headers = sheet.getDataRange().getValues()[0];
  const nuevoId = siguienteIdVenta_(sheet, headers);
  const armado = armarResumenYColumnas_(datos.items);

  const fila = headers.map(function (h) {
    switch (h) {
      case 'id_venta': return nuevoId;
      case 'fecha': return new Date();
      case 'asesor': return nombreAsesor;
      case 'subido_por': return sesion.usuario;
      case 'transportadora_distribuidor': return datos.transportadoraDistribuidor;
      case 'departamento': return datos.departamento || '';
      case 'ciudad': return datos.ciudad;
      case 'cedula': return datos.cedula;
      case 'nombre': return datos.nombre;
      case 'direccion': return datos.direccion || '';
      case 'celular': return datos.celular || '';
      case 'telefono': return datos.telefono || '';
      case 'correo_electronico': return datos.correoElectronico || '';
      case 'observacion': return datos.observacion || '';
      case 'id_kit': return datos.idKit || 'personalizado';
      case 'resumen_kit': return armado.resumen;
      case 'valor_sugerido': return v.valorSugerido;
      case 'valor_real': return v.valorReal;
      case 'comision_(%)': return v.comision;
      case 'estado_aprobacion': return 'Pendiente';
      case 'numero_guia_remision': return '';
      case 'semaforo': return '';
      case 'fve': return '';
      case 'pagado a': return '';
      case 'vlr fac': return '';
      case 'quincena': return '';
      case 'vlr comision': return '';
      default:
        return armado.columnas[h] !== undefined ? armado.columnas[h] : '';
    }
  });

  sheet.appendRow(fila);
  return { ok: true, idVenta: nuevoId };
}

// ==================================================
// Asignación de guía / remisión y semáforo
// ==================================================

/** Transportadora: asigna número de guía y pone el semáforo en "Creado". */
function asignarNumeroGuia(token, idVenta, numeroGuia) {
  requiereRol_(token, [ROLES.LOGISTICA, ROLES.ADMINISTRADOR]);
  if (!numeroGuia) throw new Error('El número de guía no puede estar vacío.');

  const sheet = getVentasTransportadoraSheet_();
  const headers = sheet.getDataRange().getValues()[0];
  const encontrada = encontrarFila_(sheet, headers, idVenta);
  if (!encontrada) throw new Error('Venta no encontrada.');

  const idxGuia = headers.indexOf('numero_guia');
  const idxSemaforo = headers.indexOf('semaforo');
  sheet.getRange(encontrada.numeroFila, idxGuia + 1).setValue(numeroGuia);
  if (!encontrada.fila[idxSemaforo]) {
    sheet.getRange(encontrada.numeroFila, idxSemaforo + 1).setValue('Creado');
  }
  return { ok: true };
}

/** Distribuidor: remisión manual, editable por Administrador en cualquier momento. */
function asignarRemision(token, idVenta, remision) {
  requiereRol_(token, [ROLES.ADMINISTRADOR]);
  if (!remision) throw new Error('El número de remisión no puede estar vacío.');

  const sheet = getVentasDistribuidorSheet_();
  const headers = sheet.getDataRange().getValues()[0];
  const encontrada = encontrarFila_(sheet, headers, idVenta);
  if (!encontrada) throw new Error('Venta no encontrada.');

  const idxRemision = headers.indexOf('remision');
  sheet.getRange(encontrada.numeroFila, idxRemision + 1).setValue(remision);
  return { ok: true };
}

/**
 * El asesor (dueño de la venta) cambia el semáforo según lo que ve al
 * rastrear. Administrador también puede, como excepción de supervisión.
 * Solo aplica a Ventas_Transportadora.
 */
function actualizarSemaforo(token, idVenta, nuevoSemaforo) {
  const sesion = requiereRol_(token, null);
  if (ESTADOS_SEMAFORO.indexOf(nuevoSemaforo) === -1) {
    throw new Error('Estado de semáforo no válido.');
  }

  const sheet = getVentasTransportadoraSheet_();
  const headers = sheet.getDataRange().getValues()[0];
  const encontrada = encontrarFila_(sheet, headers, idVenta);
  if (!encontrada) throw new Error('Venta no encontrada.');

  const idxAsesor = headers.indexOf('asesor');
  const dueño = encontrada.fila[idxAsesor];
  if (sesion.rol !== ROLES.ADMINISTRADOR && sesion.usuario !== dueño) {
    throw new Error('Solo el asesor dueño de la venta puede cambiar su semáforo.');
  }

  const idxSemaforo = headers.indexOf('semaforo');
  sheet.getRange(encontrada.numeroFila, idxSemaforo + 1).setValue(nuevoSemaforo);
  return { ok: true };
}

/** Administrador o Coordinador aprueban/rechazan una consignación pendiente. */
function aprobarConsignacion(token, idVenta, nuevoEstado) {
  requiereRol_(token, [ROLES.ADMINISTRADOR, ROLES.COORDINADOR]);
  if (ESTADOS_APROBACION_CONSIGNACION.indexOf(nuevoEstado) === -1) {
    throw new Error('Estado de aprobación no válido.');
  }

  const sheet = getVentasConsignacionSheet_();
  const headers = sheet.getDataRange().getValues()[0];
  const encontrada = encontrarFila_(sheet, headers, idVenta);
  if (!encontrada) throw new Error('Consignación no encontrada.');

  const idxEstado = headers.indexOf('estado_aprobacion');
  sheet.getRange(encontrada.numeroFila, idxEstado + 1).setValue(nuevoEstado);
  return { ok: true };
}

/**
 * Ya aprobada la consignación: asigna guía o remisión según a quién iba
 * dirigida (Logística/Administrador si es transportadora real, solo
 * Administrador si terminó siendo Distribuidor). Si aplica semáforo
 * (no es Distribuidor), lo deja en "Creado".
 */
function asignarGuiaRemisionConsignacion(token, idVenta, valor) {
  const sesion = requiereRol_(token, [ROLES.LOGISTICA, ROLES.ADMINISTRADOR]);
  if (!valor) throw new Error('El número no puede estar vacío.');

  const sheet = getVentasConsignacionSheet_();
  const headers = sheet.getDataRange().getValues()[0];
  const encontrada = encontrarFila_(sheet, headers, idVenta);
  if (!encontrada) throw new Error('Consignación no encontrada.');

  const idxEstadoAprob = headers.indexOf('estado_aprobacion');
  if (encontrada.fila[idxEstadoAprob] !== 'Aprobada') {
    throw new Error('Solo se puede asignar guía/remisión a una consignación ya aprobada.');
  }

  const idxEntidad = headers.indexOf('transportadora_distribuidor');
  const esDistribuidor = encontrada.fila[idxEntidad] === 'Distribuidor';
  if (esDistribuidor && sesion.rol !== ROLES.ADMINISTRADOR) {
    throw new Error('Solo el Administrador asigna la remisión cuando es Distribuidor.');
  }

  const idxValor = headers.indexOf('numero_guia_remision');
  sheet.getRange(encontrada.numeroFila, idxValor + 1).setValue(valor);

  if (!esDistribuidor) {
    const idxSemaforo = headers.indexOf('semaforo');
    if (!encontrada.fila[idxSemaforo]) {
      sheet.getRange(encontrada.numeroFila, idxSemaforo + 1).setValue('Creado');
    }
  }
  return { ok: true };
}

/** Semáforo de una consignación ya con guía asignada (no aplica a Distribuidor). */
function actualizarSemaforoConsignacion(token, idVenta, nuevoSemaforo) {
  const sesion = requiereRol_(token, null);
  if (ESTADOS_SEMAFORO.indexOf(nuevoSemaforo) === -1) {
    throw new Error('Estado de semáforo no válido.');
  }

  const sheet = getVentasConsignacionSheet_();
  const headers = sheet.getDataRange().getValues()[0];
  const encontrada = encontrarFila_(sheet, headers, idVenta);
  if (!encontrada) throw new Error('Consignación no encontrada.');

  const idxAsesor = headers.indexOf('asesor');
  const dueño = encontrada.fila[idxAsesor];
  if (sesion.rol !== ROLES.ADMINISTRADOR && sesion.usuario !== dueño) {
    throw new Error('Solo el asesor dueño de la venta puede cambiar su semáforo.');
  }

  const idxSemaforo = headers.indexOf('semaforo');
  sheet.getRange(encontrada.numeroFila, idxSemaforo + 1).setValue(nuevoSemaforo);
  return { ok: true };
}

// ==================================================
// Historial combinado (módulo "Ventas" de revisión/control)
// ==================================================

function listarVentas(token) {
  try {
    const sesion = requiereRol_(token, [ROLES.ASESOR, ROLES.COORDINADOR, ROLES.ADMINISTRADOR, ROLES.LOGISTICA]);
    const soloAsesor = sesion.rol === ROLES.ASESOR ? sesion.usuario : null;

    const ventas = [];
    const erroresPorHoja = [];

    [
      { nombre: 'Ventas_Transportadora', obtener: getVentasTransportadoraSheet_, tipo: 'Transportadora' },
      { nombre: 'Ventas_Distribuidor', obtener: getVentasDistribuidorSheet_, tipo: 'Distribuidor' },
      { nombre: 'Ventas_Consignacion', obtener: getVentasConsignacionSheet_, tipo: 'Consignacion' }
    ].forEach(function (config) {
      try {
        const sheet = config.obtener();
        if (!sheet) {
          erroresPorHoja.push(config.nombre + ': no se encontró la pestaña (revisa que el nombre sea exactamente ese).');
          return;
        }
        ventas.push.apply(ventas, leerVentasDeHoja_(sheet, config.tipo, soloAsesor));
      } catch (e) {
        erroresPorHoja.push(config.nombre + ': ' + e.message);
      }
    });

    ventas.sort(function (a, b) { return new Date(b.fecha) - new Date(a.fecha); });

    return { ventas: ventas, avisos: erroresPorHoja };
  } catch (errorGeneral) {
    // Cualquier error no previsto (sesión, permisos, lo que sea) llega igual
    // al navegador como un mensaje legible, en vez de perderse en el camino.
    return { ventas: [], avisos: ['ERROR: ' + errorGeneral.message] };
  }
}

function leerVentasDeHoja_(sheet, tipo, soloAsesor) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];

  const idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });

  const resultado = [];
  for (let i = 1; i < data.length; i++) {
    const f = data[i];
    const asesor = f[idx['asesor']];
    if (soloAsesor && asesor !== soloAsesor) continue;

    const fechaCelda = f[idx['fecha']];
    const fechaTexto = fechaCelda instanceof Date ? fechaCelda.toISOString() : String(fechaCelda || '');

    const base = {
      idVenta: f[idx['id_venta']],
      tipo: tipo,
      fecha: fechaTexto,
      asesor: asesor,
      subidoPor: idx['subido_por'] !== undefined ? f[idx['subido_por']] : '',
      departamento: idx['departamento'] !== undefined ? f[idx['departamento']] : '',
      ciudad: f[idx['ciudad']],
      cedula: f[idx['cedula']],
      nombreCliente: f[idx['nombre']],
      direccion: f[idx['direccion']],
      celular: f[idx['celular']],
      telefono: f[idx['telefono']],
      correoElectronico: f[idx['correo_electronico']],
      observacion: f[idx['observacion']],
      resumenKit: f[idx['resumen_kit']],
      valorSugerido: f[idx['valor_sugerido']],
      valorReal: f[idx['valor_real']],
      comisionPorcentaje: f[idx['comision_(%)']],
      fve: f[idx['fve']],
      pagadoA: f[idx['pagado a']],
      vlrFac: f[idx['vlr fac']],
      quincena: f[idx['quincena']],
      vlrComision: f[idx['vlr comision']]
    };

    if (tipo === 'Transportadora') {
      base.entidad = f[idx['transportadora']];
      base.referencia = f[idx['numero_guia']];
      base.semaforo = idx['semaforo'] !== undefined ? f[idx['semaforo']] : '';
    } else if (tipo === 'Distribuidor') {
      base.entidad = f[idx['distribuidor']];
      base.referencia = f[idx['remision']];
      base.estadoEntrega = idx['estado_entrega'] !== undefined ? f[idx['estado_entrega']] : '';
      base.observacionEntrega = idx['observacion_entrega'] !== undefined ? f[idx['observacion_entrega']] : '';
    } else if (tipo === 'Consignacion') {
      base.entidad = f[idx['transportadora_distribuidor']];
      base.referencia = f[idx['numero_guia_remision']];
      base.estadoAprobacion = f[idx['estado_aprobacion']];
      base.semaforo = idx['semaforo'] !== undefined ? f[idx['semaforo']] : '';
    }

    resultado.push(base);
  }
  return resultado;
}
