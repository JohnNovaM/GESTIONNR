/**
 * Usuarios.gs
 * Gestión de usuarios: listar, crear (con hash automático) y activar/desactivar.
 * Todas las funciones son exclusivas del rol Administrador.
 */

/**
 * Lista liviana de asesores activos (usuario + nombre), para módulos que
 * necesitan asignarles cosas (ej. Recepción) sin dar acceso completo a la
 * gestión de usuarios.
 */
function listarAsesoresActivos(token) {
  requiereRol_(token, null);

  const sheet = getUsuariosSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idxUsuario = headers.indexOf('usuario');
  const idxNombre = headers.indexOf('nombre');
  const idxRol = headers.indexOf('rol');
  const idxActivo = headers.indexOf('activo');

  const asesores = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][idxRol] === ROLES.ASESOR && data[i][idxActivo]) {
      asesores.push({ usuario: data[i][idxUsuario], nombre: data[i][idxNombre] });
    }
  }
  return asesores;
}

function listarUsuarios(token) {
  requiereRol_(token, [ROLES.ADMINISTRADOR]);

  const sheet = getUsuariosSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idxId = headers.indexOf('id_usuario');
  const idxNombre = headers.indexOf('nombre');
  const idxUsuario = headers.indexOf('usuario');
  const idxRol = headers.indexOf('rol');
  const idxActivo = headers.indexOf('activo');

  const usuarios = [];
  for (let i = 1; i < data.length; i++) {
    usuarios.push({
      idUsuario: data[i][idxId],
      nombre: data[i][idxNombre],
      usuario: data[i][idxUsuario],
      rol: data[i][idxRol],
      activo: !!data[i][idxActivo]
    });
  }
  return usuarios;
}

/**
 * Crea un usuario nuevo. Recibe la contraseña en texto plano desde el
 * formulario, la hashea aquí antes de guardar — el texto plano nunca
 * llega a tocar la hoja de cálculo.
 */
function crearUsuario(token, datos) {
  requiereRol_(token, [ROLES.ADMINISTRADOR]);

  if (!datos || !datos.usuario || !datos.nombre || !datos.password || !datos.rol) {
    throw new Error('Todos los campos son obligatorios.');
  }
  if (Object.values(ROLES).indexOf(datos.rol) === -1) {
    throw new Error('Rol no válido.');
  }
  if (String(datos.password).length < 4) {
    throw new Error('La contraseña debe tener al menos 4 caracteres.');
  }

  const sheet = getUsuariosSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxUsuario = headers.indexOf('usuario');
  const idxId = headers.indexOf('id_usuario');

  const usuarioNuevo = String(datos.usuario).trim().toLowerCase();
  let maxId = 0;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxUsuario]).trim().toLowerCase() === usuarioNuevo) {
      throw new Error('Ya existe un usuario con ese nombre de usuario.');
    }
    const idNum = parseInt(data[i][idxId], 10);
    if (!isNaN(idNum) && idNum > maxId) maxId = idNum;
  }

  const nuevoId = maxId + 1;
  const passwordHash = hashPassword_(datos.password);

  const fila = headers.map(function (h) {
    switch (h) {
      case 'id_usuario': return nuevoId;
      case 'nombre': return datos.nombre;
      case 'usuario': return datos.usuario;
      case 'password_hash': return passwordHash;
      case 'rol': return datos.rol;
      case 'activo': return true;
      case 'fecha_creacion': return new Date();
      default: return '';
    }
  });

  sheet.appendRow(fila);
  return { ok: true, idUsuario: nuevoId };
}

function cambiarEstadoUsuario(token, idUsuario, activo) {
  requiereRol_(token, [ROLES.ADMINISTRADOR]);

  const sheet = getUsuariosSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxId = headers.indexOf('id_usuario');
  const idxActivo = headers.indexOf('activo');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === String(idUsuario)) {
      sheet.getRange(i + 1, idxActivo + 1).setValue(activo);
      return { ok: true };
    }
  }
  throw new Error('Usuario no encontrado.');
}
