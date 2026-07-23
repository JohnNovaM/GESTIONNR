/**
 * Correr manualmente desde el editor (sin necesidad de sesión/token).
 * Revisa Ver > Registros después.
 */
function diagnosticoListarVentas() {
  [
    { nombre: 'Ventas_Transportadora', obtener: getVentasTransportadoraSheet_ },
    { nombre: 'Ventas_Distribuidor', obtener: getVentasDistribuidorSheet_ },
    { nombre: 'Ventas_Consignacion', obtener: getVentasConsignacionSheet_ }
  ].forEach(function (config) {
    try {
      const sheet = config.obtener();
      if (!sheet) {
        Logger.log(config.nombre + ': NO ENCONTRADA (revisa el nombre exacto de la pestaña)');
        return;
      }
      const filas = sheet.getDataRange().getValues().length;
      Logger.log(config.nombre + ': OK, encontrada. ' + filas + ' filas (incluye encabezado).');
    } catch (e) {
      Logger.log(config.nombre + ': ERROR -> ' + e.message);
    }
  });

  Logger.log('¿Existe la función listarVentas en este proyecto? ' + (typeof listarVentas === 'function'));
}
