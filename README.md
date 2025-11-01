# App – Reserva de Salones

Aplicación completa basada en Google Apps Script que permite administrar la
reserva de salones corporativos desde una hoja de cálculo, con interfaces web
para solicitantes y administradores. El proyecto centraliza la lógica de
negocio en `Código.js` y utiliza plantillas HTML publicadas como Web App para
distribuir la experiencia a diferentes perfiles.

## Características principales

- **Autenticación por dominio** y validación de usuarios registrados con estado
  y rol definidos en la pestaña `Usuarios`.
- **Panel público** para solicitantes y administradores con formulario dinámico
  de creación de reservas y consulta de disponibilidad.
- **Panel administrativo** con herramientas para revisar solicitudes, aprobar,
  denegar o cancelar reservas existentes.
- **Gestión de catálogos** de salones, con capacidad, sede, banderas de
  habilitación y reglas de restricción horaria.
- **Multiadministración por salón**, permitiendo delegar la operación a
  distintos equipos con configuraciones (horarios, duración, contactos y
  branding de correo) independientes por administración.
- **Notificaciones por correo** en cada transición relevante de la reserva.
- **Integración con Google Drive** para incrustar el logotipo institucional en
  correos e interfaces.
- **Prioridades condicionadas por salón**, permitiendo que ciertos usuarios
  solo tengan ventaja en los espacios explícitamente asignados en su ficha.

## Estructura del repositorio

| Ruta              | Descripción                                                                 |
|-------------------|-----------------------------------------------------------------------------|
| `Código.js`       | Lógica principal de Apps Script: helpers, control de acceso, API interna y
|                   | envío de notificaciones.                                                    |
| `Admin.html`      | Plantilla HTML para la vista de administración.                             |
| `Public.html`     | Plantilla HTML para el flujo de solicitudes de reserva.                     |
| `Cancel.html`     | Plantilla para la cancelación mediante enlace directo.                      |
| `Denied.html`     | Mensajes de acceso denegado para estados no permitidos.                     |
| `admin.js.html`   | Recursos JavaScript específicos de la vista administrativa.                 |
| `public.js.html`  | Lógica frontend usada por solicitantes.                                     |
| `css.html`        | Estilos compartidos entre vistas.                                           |
| `appsscript.json` | Manifest del proyecto Apps Script (scopes, archivos y configuración).       |

## Requisitos previos

1. **Cuenta de Google Workspace** con permisos para crear y desplegar Apps Script.
2. **Hoja de cálculo asociada** con las pestañas mínimas: `Config`, `Usuarios`,
   `Conserjes`, `Salones` y `Reservas`. Cada pestaña debe seguir el layout
   esperado por el script.
3. Definir la zona horaria del proyecto en `File → Project properties → Script
   properties` (`America/Santo_Domingo` sugerida).【F:Código.js†L6-L13】
4. Habilitar la **Drive API avanzada** en `Services` y en Google Cloud Console
   para obtener miniaturas/logos en correos.【F:Código.js†L8-L14】

## Configuración inicial

1. Duplicar la hoja de cálculo original y ajustar los datos base en cada pestaña.
2. En `Config`, completar las claves utilizadas por el script (`ADMIN_EMAILS`,
   `HORARIO_INICIO`, `HORARIO_FIN`, `DURATION_MIN`, `DURATION_MAX`,
   `DURATION_STEP`, etc.).
3. En `Usuarios`, registrar a todas las personas que usarán la herramienta,
   definiendo `rol` (`ADMIN1`, `ADMIN2`, `ADMIN3`, …, `SOLICITANTE`), `estado`
   (`ACTIVO`, `PENDIENTE`, `INACTIVO`), `prioridad` y el campo
   `prioridad_salones`. Este último permite restringir la prioridad a ciertos
   salones escribiendo sus códigos separados por `;` (ej. `S00004;S00005`).
   Para los demás espacios la prioridad se tratará como 0. Completa también los
   campos opcionales como `extension`. El número que acompaña al rol `ADMIN`
   determina la administración a la que pertenece el usuario: `ADMIN1` accede a
   la administración general (Config), `ADMIN2` a la hoja `Config2`, y así
   sucesivamente. Cada administrador `ADMIN#` únicamente verá y gestionará los
   salones y reservas asociados a su propia administración en el panel
   administrativo. Si necesitas un perfil con acceso total a todas las
   administraciones, asigna el rol `ADMIN` (sin sufijo) o `SUPERADMIN`, o bien
   incluye su correo en la clave opcional `ADMIN_ALL_ACCESS_EMAILS` de la hoja
   `Config` (separando múltiples direcciones con `;`).
4. En `Salones`, mantener el catálogo de espacios disponibles, con campos de
   capacidad, sede y las columnas `restriccion`, `conserje` y `administracion`.
   La columna `restriccion` acepta:
   - Intervalos separados por `;` con formato `HH:MM-HH:MM` para bloquear
     horarios específicos (por ejemplo `11:00-14:00;08:00-09:00`).
   - El valor `CONFIRM` (o `COFNIRM`) para forzar que las solicitudes entren en
     estado `PENDIENTE` y requieran aprobación manual.
   La columna `conserje` controla si, además de que el evento se extienda más
   allá de las 16:00, el sistema debe solicitar asignación de conserje (`SI`)
   o si puede omitir ese flujo (`NO`). La columna `administracion` determina
   qué configuración aplicará a ese salón: `1` utiliza la hoja `Config`
   tradicional, mientras que `2`, `3`, etc. consultan hojas adicionales como
   `Config2`, `Config3` para obtener parámetros específicos (horarios,
   duraciones, contactos, remitentes) de cada administración.
   Cada hoja adicional debe incluir al menos las claves `ADMIN_EMAILS`,
   `HORARIO_INICIO`, `HORARIO_FIN`, `DURATION_MIN`, `DURATION_STEP`,
   `DURATION_MAX`, `MAIL_SENDER_NAME`, `MAIL_REPLY_TO`, `ADMIN_CONTACT_NAME`,
   `ADMIN_CONTACT_EMAIL` y `ADMIN_CONTACT_EXTENSION`; cualquier clave faltante
   heredará automáticamente el valor definido en la hoja `Config` principal.
5. Verificar la pestaña `Reservas` para asegurar que las columnas coinciden con
   la estructura consumida por el backend (ID, solicitante, fechas, horarios,
   estado, etc.).
6. Personalizar los archivos HTML y CSS si se requiere adaptar la identidad
   gráfica corporativa.

## Despliegue como Web App

1. Abrir el editor de Apps Script y seleccionar `Deploy → New deployment`.
2. Elegir el tipo **Web app**, añadir una descripción y seleccionar:
   - *Execute as*: `Me` (propietario del script).
   - *Who has access*: `Only people in your organization` o el ámbito requerido.
3. Guardar y autorizar los permisos OAuth solicitados.
4. Compartir la URL resultante con los usuarios. El script controlará el acceso
   según los registros de la pestaña `Usuarios` y la lista `ADMIN_EMAILS`.

## Desarrollo y mantenimiento

- El control de acceso se maneja mediante `getUser_`, `isAdminEmail_` y la
  lógica del `doGet`, que enruta a las vistas correctas según el rol y estado
  del usuario.【F:Código.js†L27-L134】
- Los catálogos (`apiListSalones`, etc.) leen datos directamente de cada pestaña
  y devuelven objetos JSON usados por las interfaces HTML.【F:Código.js†L136-L154】
- Las restricciones de salones se interpretan con `parseSalonRestriction_` y se
  aplican al consultar disponibilidad (`apiListDisponibilidad`) y al crear
  reservas (`apiCrearReserva`). Los intervalos bloqueados no aparecen como
  opciones seleccionables y los salones marcados con `CONFIRM/COFNIRM` generan
  solicitudes en estado `PENDIENTE`.【F:Código.js†L340-L707】【F:Código.js†L974-L1076】
- El frontend público y el panel administrativo consumen la metadata de
  restricción para mostrar avisos, deshabilitar horarios restringidos y permitir
  la aprobación manual de reservas desde la interfaz.【F:public.js.html†L24-L360】【F:admin.js.html†L20-L219】
- Las notificaciones por correo se encuentran agrupadas en funciones `notify*`
  dentro del mismo archivo.
- Para modificar estilos o componentes del frontend, editar los archivos
  `*.html` correspondientes y volver a desplegar el Web App.

## Automatizaciones

El directorio `workflows/` contiene pipelines opcionales (por ejemplo,
`deploy.yml`) que pueden integrarse con GitHub Actions para mantener la
sincronización del código Apps Script mediante clasp u otras herramientas.

## Contribuir

1. Crear una rama a partir de `main` y aplicar los cambios.
2. Ejecutar las pruebas manuales necesarias (creación de reserva, aprobación,
   cancelación).
3. Abrir un Pull Request describiendo el alcance y adjuntar evidencias.
4. Tras la aprobación, desplegar una nueva versión del Web App y comunicar la
   actualización a los usuarios.

¡Listo! Con estos pasos el equipo puede mantener y evolucionar la App de
Reserva de Salones con trazabilidad y documentación centralizada.
