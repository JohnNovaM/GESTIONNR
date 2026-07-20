/**
 * Auth.gs
 * Login, manejo de sesión y control de rol.
 *
 * Requiere una hoja "Usuarios" con columnas:
 * id_usuario | nombre | usuario | password_hash | rol | activo | fecha_creacion
 *
 * Requiere una Script Property llamada SHEET_ID con el ID del Google Sheet.
 * (Configuración del proyecto > Script Properties, o correr configurarSheetId()
 * una sola vez desde el editor).
 */

const HOJA_USUARIOS = 'Usuarios';
const SESSION_DURATION_SECONDS = 21600; // 6 horas — máximo permitido por CacheService

const ROLES = {
  ASESOR: 'Asesor',
  COORDINADOR: 'Coordinador',
  ADMINISTRADOR: 'Administrador',
  LOGISTICA: 'Logistica'
};

/**
 * Helper de configuración inicial. Correr UNA VEZ desde el editor de Apps Script,
 * reemplazando el ID por el de tu Sheet.
 */
function configurarSheetId() {
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', 'PON_AQUI_EL_ID_DE_TU_SHEET');
}

function getSheetId_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) {
    throw new Error('Falta configurar SHEET_ID. Corre configurarSheetId() desde el editor.');
  }
  return id;
}

function getUsuariosSheet_() {
  return SpreadsheetApp.openById(getSheetId_()).getSheetByName(HOJA_USUARIOS);
}

/**
 * Genera el hash SHA-256 de una contraseña en texto plano.
 */
function hashPassword_(password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  return bytes.map(function (b) {
    return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0');
  }).join('');
}

/**
 * Utilidad para generar el hash de la contraseña de un usuario nuevo.
 * Correr desde el editor, revisar el Log (Ver > Registros) y pegar el
 * resultado en la columna password_hash de la hoja Usuarios.
 */
function generarHashParaUsuario(passwordEnTextoPlano) {
  Logger.log(hashPassword_(passwordEnTextoPlano));
}

/**
 * Login. Llamado desde el cliente vía google.script.run.
 */
function login(usuario, password) {
  if (!usuario || !password) {
    return { ok: false, error: 'Usuario y contraseña son obligatorios.' };
  }

  const sheet = getUsuariosSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idxId = headers.indexOf('id_usuario');
  const idxNombre = headers.indexOf('nombre');
  const idxUsuario = headers.indexOf('usuario');
  const idxPassword = headers.indexOf('password_hash');
  const idxRol = headers.indexOf('rol');
  const idxActivo = headers.indexOf('activo');

  const passwordHash = hashPassword_(password);
  const usuarioBuscado = String(usuario).trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const fila = data[i];
    if (String(fila[idxUsuario]).trim().toLowerCase() === usuarioBuscado) {
      if (!fila[idxActivo]) {
        return { ok: false, error: 'Usuario inactivo. Contacta al administrador.' };
      }
      if (String(fila[idxPassword]) !== passwordHash) {
        return { ok: false, error: 'Usuario o contraseña incorrectos.' };
      }

      const token = Utilities.getUuid();
      const sesion = {
        idUsuario: fila[idxId],
        usuario: fila[idxUsuario],
        nombre: fila[idxNombre],
        rol: fila[idxRol]
      };
      CacheService.getScriptCache().put(token, JSON.stringify(sesion), SESSION_DURATION_SECONDS);

      return {
        ok: true,
        token: token,
        nombre: sesion.nombre,
        rol: sesion.rol,
        menu: obtenerMenuPorRol(sesion.rol)
      };
    }
  }

  return { ok: false, error: 'Usuario o contraseña incorrectos.' };
}

/**
 * Valida un token de sesión. Devuelve los datos de sesión o null si no es válido/expiró.
 */
function validarSesion(token) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get(token);
  if (!raw) return null;
  return JSON.parse(raw);
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove(token);
  return { ok: true };
}

/**
 * Lanza un error si la sesión no es válida o el rol no está permitido.
 * Cada función de cada módulo (Ventas.gs, Matriz.gs, etc.) debe llamar esto
 * como primera línea, pasando los roles que sí pueden ejecutarla.
 *
 * Ejemplo:
 *   function guardarVenta(token, datosVenta) {
 *     const sesion = requiereRol_(token, [ROLES.ASESOR, ROLES.COORDINADOR]);
 *     ...
 *   }
 */
function requiereRol_(token, rolesPermitidos) {
  const sesion = validarSesion(token);
  if (!sesion) {
    throw new Error('Sesión inválida o expirada. Vuelve a iniciar sesión.');
  }
  if (rolesPermitidos && rolesPermitidos.indexOf(sesion.rol) === -1) {
    throw new Error('No tienes permiso para esta acción.');
  }
  return sesion;
}

/**
 * Menú placeholder por rol. Los módulos aún no existen — esto solo define
 * qué debe ver cada rol para que el frontend lo pinte. Se irá conectando
 * a pantallas reales a medida que se construya cada módulo.
 */
function obtenerMenuPorRol(rol) {
  const menus = {
    Asesor: ['Cotizador', 'Ventas', 'Recepción', 'Minutos', 'Novedades', 'Desplazamientos', 'Pagos', 'Cuentas de Cobro'],
    Coordinador: ['Ventas (equipo)', 'Novedades (equipo)', 'Desplazamientos (equipo)', 'Reportes'],
    Administrador: ['Usuarios', 'Catálogo y Kits', 'Matriz', 'Archivado', 'Reportes'],
    Logistica: ['Descarga de Ventas']
  };
  return menus[rol] || [];
}
