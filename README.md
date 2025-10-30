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
- **Gestión de catálogos** de salones, con capacidad, sede y banderas de
  habilitación.
- **Notificaciones por correo** en cada transición relevante de la reserva.
- **Integración con Google Drive** para incrustar el logotipo institucional en
  correos e interfaces.

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
   definiendo `rol` (`ADMIN`, `SOLICITANTE`), `estado` (`ACTIVO`, `PENDIENTE`,
   `INACTIVO`) y campos opcionales como `extension`.
4. En `Salones`, mantener el catálogo de espacios disponibles, con campos de
   capacidad y sede.
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
