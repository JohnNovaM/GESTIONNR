# App Gestión de Asesores — Módulo de Login

Primer módulo del proyecto: autenticación y control de rol. Los demás módulos
(Cotizador, Ventas, Matriz, etc.) se construyen sobre esta base.

## 1. Configurar el proyecto

1. Crea un proyecto de Apps Script (script.google.com) o vincúlalo a uno existente.
2. Copia `.clasp.json.example` a `.clasp.json` y pon el `scriptId` de tu proyecto
   (Configuración del proyecto > IDs > Identificador del script).
3. Desde la carpeta del proyecto:
   ```
   clasp push
   ```

## 2. Conectar el Google Sheet

1. Abre tu Google Sheet (el que ya tiene la pestaña `Usuarios`).
2. Copia su ID (la parte de la URL entre `/d/` y `/edit`).
3. En el editor de Apps Script, abre `Auth.gs`, entra a la función
   `configurarSheetId()`, reemplaza el texto por tu ID real, y ejecútala
   **una sola vez** desde el editor (▶ Ejecutar).
   - Esto guarda el ID en las Script Properties del proyecto, no queda
     escrito en el código — así puedes subir el código a Git sin exponerlo.

## 3. Crear tu primer usuario administrador

La hoja `Usuarios` no tiene registro con qué entrar todavía. Para crear el
primero:

1. En el editor de Apps Script, abre `Auth.gs`.
2. En la función `generarHashParaUsuario`, ejecútala pasando la contraseña
   que quieras (o corre `generarHashParaUsuario('tu_clave_aqui')` desde el
   editor con el cursor dentro de la función, luego ▶ Ejecutar).
3. Ve a **Ver > Registros** (View > Logs) y copia el hash que aparece.
4. En la hoja `Usuarios`, agrega una fila:
   `id_usuario=1 | nombre=Tu Nombre | usuario=admin | password_hash=<el hash copiado> | rol=Administrador | activo=TRUE | fecha_creacion=hoy`

## 4. Publicar y probar

1. En el editor: **Implementar > Nueva implementación > Aplicación web**.
2. "Ejecutar como": tu cuenta. "Quién tiene acceso": según prefieras
   (dominio, o cuentas específicas). Así los asesores nunca necesitan
   acceso directo al Sheet — solo a la URL del Web App.
3. Abre la URL de la implementación, ingresa con el usuario `admin` y la
   contraseña que definiste.
4. Deberías ver el menú correspondiente al rol Administrador (por ahora
   son botones placeholder — cada uno se conecta a un módulo real más
   adelante).

## Qué sigue

- Cada módulo nuevo (Ventas, Matriz, Cotizador, etc.) agrega su propio
  archivo `.gs`, y sus funciones deben empezar validando la sesión con
  `requiereRol_(token, [ROLES.ASESOR, ...])` — ver el ejemplo comentado
  en `Auth.gs`.
- El menú por rol (`obtenerMenuPorRol` en `Auth.gs`) se va conectando a
  pantallas reales a medida que existan — hoy son botones que muestran
  "módulo en construcción".
- Las sesiones duran 6 horas (máximo permitido por `CacheService`). Si
  más adelante necesitas sesiones más largas ("recordarme"), se puede
  cambiar a un esquema con `PropertiesService` en vez de `CacheService`.
