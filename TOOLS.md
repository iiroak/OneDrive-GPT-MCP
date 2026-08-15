# OneDrive-GPT-MCP: Tools

Referencia de las tools definidas en `index.js` y los módulos `auth/`,
`calendar/`, `email/`, `folder/`, `rules/` y `onedrive/`.

## Transporte Y Visibilidad

| Transporte | Endpoint | Tools |
|---|---|---:|
| MCP remoto | `https://mcp.example.com/outlook/mcp` | 50 |
| MCP stdio/local | `node index.js` | 50 |

El remoto usa el flujo OAuth/PKCE del propio MCP para autorizar al cliente y
conserva aparte el OAuth delegado de Microsoft Graph.

## Scopes

Cada tool remota publica uno de estos scopes OAuth:

| Scope | Uso |
|---|---|
| `outlook:read` | Lecturas sin modificar Outlook, calendario, reglas ni OneDrive |
| `outlook:write` | Crear, editar, mover, marcar, subir o compartir |
| `outlook:destructive` | Borrado irreversible o acciones explícitamente destructivas |

Las annotations MCP también marcan las tools de escritura y las destructivas.

## Respuesta Común

El servidor remoto devuelve el texto legible en `content` y un sobre estructurado:

```json
{
  "message": "Resultado legible para humanos",
  "data": {}
}
```

Cuando una tool no tiene datos adicionales, `data` puede omitirse. Las tools de
OneDrive que manejan archivos incluyen sus metadatos dentro de `data`.

Los argumentos indicados como `required` deben enviarse. El resto son opcionales.

# Auth

## `about`

**Scope:** `outlook:read`
**Remote:** sí
**Efecto:** ninguno. Devuelve nombre, versión, servicios y transporte.

**Parámetros:** ninguno.

## `check-auth-status`

**Scope:** `outlook:read`
**Remote:** sí
**Efecto:** ninguno. Comprueba si existe un access token Graph válido.

**Parámetros:** ninguno.

# Calendar

Los identificadores de calendario y evento son IDs de Microsoft Graph. Si se
omite `calendarId`, se usa el calendario predeterminado.

## `list-calendars`

**Scope:** `outlook:read`
**Efecto:** ninguno.

| Parámetro | Tipo | Requerido | Descripción |
|---|---|---:|---|
| `count` | integer, 1-100 | no | Máximo de calendarios a devolver. |

## `create-calendar`

**Scope:** `outlook:write`
**Efecto:** crea un calendario nuevo.

| Parámetro | Tipo | Requerido | Descripción |
|---|---|---:|---|
| `name` | string | sí | Nombre del calendario. |

## `update-calendar`

**Scope:** `outlook:write`
**Efecto:** cambia el nombre o color de un calendario.

| Parámetro | Tipo | Requerido | Descripción |
|---|---|---:|---|
| `calendarId` | string | sí | ID del calendario. |
| `name` | string | no | Nuevo nombre. |
| `color` | string | no | Nuevo color de Outlook. |

## `delete-calendar`

**Scope:** `outlook:destructive`
**Efecto:** elimina un calendario no predeterminado. No usar sobre el calendario
principal.

| Parámetro | Tipo | Requerido |
|---|---|---:|
| `calendarId` | string | sí |

## `list-events`

**Scope:** `outlook:read`
**Efecto:** ninguno. Lista eventos en una vista de calendario.

| Parámetro | Tipo | Requerido | Descripción |
|---|---|---:|---|
| `calendarId` | string | no | Calendario; predeterminado si se omite. |
| `count` | integer, 1-50 | no | Máximo de eventos. |
| `startDateTime` | string | no | Inicio del rango, normalmente ISO 8601. |
| `endDateTime` | string | no | Fin del rango, normalmente ISO 8601. |

## `get-event`

**Scope:** `outlook:read`
**Efecto:** ninguno. Obtiene el evento completo.

| Parámetro | Tipo | Requerido |
|---|---|---:|
| `eventId` | string | sí |
| `calendarId` | string | no |

## `create-event`

**Scope:** `outlook:write`
**Efecto:** crea un evento en el calendario indicado o en el predeterminado.

| Parámetro | Tipo | Requerido | Descripción |
|---|---|---:|---|
| `subject` | string | sí | Asunto del evento. |
| `start` | string/object | sí | ISO 8601 o `{dateTime, timeZone}`. |
| `end` | string/object | sí | ISO 8601 o `{dateTime, timeZone}`. |
| `calendarId` | string | no | Calendario destino. |
| `body` | string/object | no | Texto o `{contentType, content}`. |
| `location` | string/object | no | Ubicación o objeto de ubicación. |
| `attendees` | array | no | Emails u objetos de asistentes. |
| `categories` | string[] | no | Categorías de Outlook. |
| `isAllDay` | boolean | no | Evento de día completo. |
| `isReminderOn` | boolean | no | Activa el recordatorio. |
| `reminderMinutesBeforeStart` | integer | no | Minutos antes del inicio. |
| `responseRequested` | boolean | no | Solicita respuesta de asistentes. |
| `allowNewTimeProposals` | boolean | no | Permite proponer otro horario. |
| `hideAttendees` | boolean | no | Oculta asistentes. |
| `isOnlineMeeting` | boolean | no | Crea reunión online. |
| `onlineMeetingProvider` | string | no | Proveedor de reunión online. |
| `importance` | string | no | Importancia Graph. |
| `sensitivity` | string | no | Sensibilidad Graph. |
| `showAs` | string | no | Estado de disponibilidad. |
| `recurrence` | object | no | Patrón de recurrencia Graph. |
| `transactionId` | string | no | Identificador Graph para idempotencia. |

## `update-event`

**Scope:** `outlook:write`
**Efecto:** actualiza un evento existente, incluidas sus categorías.

Acepta todos los campos opcionales de `create-event` y añade:

| Parámetro | Tipo | Requerido |
|---|---|---:|
| `eventId` | string | sí |

## `copy-event`

**Scope:** `outlook:write`
**Efecto:** copia un evento sin eliminar el original.

| Parámetro | Tipo | Requerido | Valores |
|---|---|---:|---|
| `sourceEventId` | string | sí | ID origen. |
| `targetCalendarId` | string | sí | ID destino. |
| `sourceCalendarId` | string | no | Calendario origen. |
| `categoriesMode` | string | no | `preserve` o `clear`. |
| `includeAttendees` | boolean | no | Copia asistentes. |
| `onlineMeetingMode` | string | no | `reject`, `preserveBodyOnly` o `newOnlineMeeting`. |
| `transactionId` | string | no | Idempotencia Graph. |

## `migrate-events`

**Scope:** `outlook:destructive`
**Efecto:** copia, verifica y elimina los eventos originales. Es irreversible
después de la eliminación.

| Parámetro | Tipo | Requerido | Valores |
|---|---|---:|---|
| `sourceEventIds` | string[] | sí | Uno o más IDs origen. |
| `targetCalendarId` | string | sí | ID calendario destino. |
| `confirm` | string | sí | Debe ser exactamente `MIGRATE_EVENTS`. |
| `sourceCalendarId` | string | no | Calendario origen. |
| `categoriesMode` | string | no | `preserve` o `clear`. |
| `includeAttendees` | boolean | no | Copia asistentes. |
| `onlineMeetingMode` | string | no | `reject`, `preserveBodyOnly` o `newOnlineMeeting`. |
| `expectedChangeKeys` | object | no | Verificación de versión Graph. |
| `transactionIdPrefix` | string | no | Prefijo de idempotencia. |

## Categorías De Outlook

| Tool | Scope | Parámetros | Efecto |
|---|---|---|---|
| `list-master-categories` | read | ninguno | Lista categorías y colores. |
| `create-master-category` | write | `displayName` requerido, `color` requerido | Crea una categoría. |
| `update-master-category` | write | `categoryId` requerido, `color` requerido | Cambia el color. |
| `delete-master-category` | destructive | `categoryId` requerido | Elimina una categoría. |

## Respuestas De Eventos

| Tool | Scope | Parámetros | Efecto |
|---|---|---|---|
| `accept-event` | write | `eventId` requerido; `calendarId`, `comment` opcionales | Acepta y responde al organizador. |
| `decline-event` | write | `eventId` requerido; `calendarId`, `comment` opcionales | Rechaza el evento. |
| `cancel-event` | destructive | `eventId` requerido; `calendarId`, `comment` opcionales | Cancela el evento. |
| `delete-event` | destructive | `eventId` requerido; `calendarId` opcional | Elimina el evento. |

# Email

## `list-emails`

**Scope:** `outlook:read`
**Efecto:** ninguno.

| Parámetro | Tipo | Requerido | Descripción |
|---|---|---:|---|
| `folder` | string | no | `inbox`, `sent`, `drafts` u otra carpeta válida. Predeterminado: `inbox`. |
| `count` | number | no | Predeterminado 10, máximo 50. |

## `search-emails`

**Scope:** `outlook:read`
**Efecto:** ninguno.

| Parámetro | Tipo | Requerido | Descripción |
|---|---|---:|---|
| `query` | string | no | Texto de búsqueda. |
| `folder` | string | no | Carpeta; predeterminada `inbox`. |
| `from` | string | no | Remitente, email o nombre. |
| `to` | string | no | Destinatario, email o nombre. |
| `subject` | string | no | Filtro de asunto. |
| `hasAttachments` | boolean | no | Solo mensajes con adjuntos. |
| `unreadOnly` | boolean | no | Solo mensajes no leídos. |
| `count` | number | no | Predeterminado 10, máximo 50. |

## `read-email`

**Scope:** `outlook:read`
**Efecto:** ninguno. El HTML se sanitiza para devolver texto visible y reducir
prompt injection mediante contenido oculto.

| Parámetro | Tipo | Requerido | Descripción |
|---|---|---:|---|
| `id` | string | sí | ID del mensaje. |
| `includeRawHtml` | boolean | no | Devuelve HTML crudo. Solo debugging; puede contener contenido oculto no seguro. |

## `send-email`

**Scope:** `outlook:write`
**Efecto:** envía un email real.

| Parámetro | Tipo | Requerido | Descripción |
|---|---|---:|---|
| `to` | string | sí | Destinatarios separados por coma. |
| `subject` | string | sí | Asunto. |
| `body` | string | sí | Cuerpo plain text o HTML. |
| `cc` | string | no | Copia separados por coma. |
| `bcc` | string | no | Copia oculta separados por coma. |
| `isHtml` | boolean | no | Fuerza HTML; si se omite, autodetecta. |
| `importance` | string | no | `normal`, `high` o `low`. |
| `saveToSentItems` | boolean | no | Guarda en enviados. |

## `draft-email`

**Scope:** `outlook:write`
**Efecto:** guarda un borrador; no lo envía.

Acepta `to`, `cc`, `bcc`, `subject`, `body` e `importance` de
`send-email`, pero todos son opcionales. `importance` acepta `normal`, `high` o
`low`.

## `mark-as-read`

**Scope:** `outlook:write`
**Efecto:** cambia el estado de lectura.

| Parámetro | Tipo | Requerido | Descripción |
|---|---|---:|---|
| `id` | string | sí | ID del mensaje. |
| `isRead` | boolean | no | `true` leído, `false` no leído. Predeterminado: `true`. |

## `trash-email`

**Scope:** `outlook:write`
**Efecto:** mueve el mensaje a Deleted Items. Es reversible desde Outlook.

| Parámetro | Tipo | Requerido |
|---|---|---:|
| `id` | string | sí |

## `permanently-delete-email`

**Scope:** `outlook:destructive`
**Efecto:** elimina permanentemente el mensaje. No se puede deshacer.

| Parámetro | Tipo | Requerido |
|---|---|---:|
| `id` | string | sí |

# Folders

## `list-folders`

**Scope:** `outlook:read`
**Efecto:** ninguno.

| Parámetro | Tipo | Requerido |
|---|---|---:|
| `includeItemCounts` | boolean | no |
| `includeChildren` | boolean | no |

## `create-folder`

**Scope:** `outlook:write`
**Efecto:** crea una carpeta de correo.

| Parámetro | Tipo | Requerido |
|---|---|---:|
| `name` | string | sí |
| `parentFolder` | string | no; raíz por defecto |

## `move-emails`

**Scope:** `outlook:write`
**Efecto:** mueve uno o varios emails.

| Parámetro | Tipo | Requerido | Descripción |
|---|---|---:|---|
| `emailIds` | string | sí | IDs separados por coma. |
| `targetFolder` | string | sí | Carpeta destino. |
| `sourceFolder` | string | no | Carpeta origen; `inbox` por defecto. |

# Rules

## `list-rules`

**Scope:** `outlook:read`
**Efecto:** ninguno.

| Parámetro | Tipo | Requerido |
|---|---|---:|
| `includeDetails` | boolean | no |

## `create-rule`

**Scope:** `outlook:write`
**Efecto:** crea una regla de bandeja de entrada.

| Parámetro | Tipo | Requerido | Descripción |
|---|---|---:|---|
| `name` | string | sí | Nombre de la regla. |
| `fromAddresses` | string | no | Remitentes separados por coma. |
| `containsSubject` | string | no | Texto que debe contener el asunto. |
| `hasAttachments` | boolean | no | Requiere adjuntos. |
| `moveToFolder` | string | no | Carpeta destino. |
| `markAsRead` | boolean | no | Marca coincidencias como leídas. |
| `isEnabled` | boolean | no | Activa la regla; por defecto `true`. |
| `sequence` | number | no | Orden de ejecución; menor número tiene prioridad. Predeterminado 100. |

## `edit-rule-sequence`

**Scope:** `outlook:write`
**Efecto:** cambia el orden de ejecución de una regla existente.

| Parámetro | Tipo | Requerido |
|---|---|---:|
| `ruleName` | string | sí |
| `sequence` | number | sí; positivo |

# OneDrive

Los identificadores son `DriveItem` IDs de Microsoft Graph. Para tools que aceptan
`itemId` o `path`, basta uno de los dos.

## `onedrive-list`

**Scope:** `outlook:read`
**Efecto:** ninguno.

| Parámetro | Tipo | Requerido | Descripción |
|---|---|---:|---|
| `path` | string | no | Ruta, por ejemplo `/Documents` o `/Photos`; raíz por defecto. |
| `count` | number | no | Predeterminado 25, máximo 50. |

## `onedrive-search`

**Scope:** `outlook:read`
**Efecto:** ninguno.

| Parámetro | Tipo | Requerido |
|---|---|---:|
| `query` | string | sí |
| `count` | number | no; predeterminado 25, máximo 50 |

## `onedrive-download`

**Scope:** `outlook:read`
**Efecto:** ninguno. Devuelve una capability URL temporal; **no devuelve el
contenido del archivo**.

| Parámetro | Tipo | Requerido |
|---|---|---:|
| `itemId` | string | uno de `itemId`/`path` |
| `path` | string | uno de `itemId`/`path` |

Usar esta tool cuando otro servidor consumidor puede dereferenciar la URL, por
ejemplo un backend de transcripción. Para que el MCP lea el archivo directamente,
usar `onedrive-read-file`.

## `onedrive-read-file`

**Scope:** `outlook:read`
**Efecto:** descarga bytes temporalmente y devuelve texto extraído.

| Parámetro | Tipo | Requerido | Descripción |
|---|---|---:|---|
| `itemId` | string | uno de `itemId`, `path`, `fileId` | ID Graph. |
| `path` | string | uno de `itemId`, `path`, `fileId` | Ruta OneDrive. |
| `fileId` | string | uno de `itemId`, `path`, `fileId` | Handle de una descarga anterior; evita otra descarga. |
| `maxChars` | number | no | Predeterminado 50.000, máximo 200.000. |

Formatos de extracción: texto plano, Markdown, HTML, PDF, DOCX, PPTX y XLSX.
Revisar siempre el estado:

| `status` | Significado |
|---|---|
| `complete` | Todo el contenido significativo se representó como texto. |
| `partial` | Faltan imágenes, gráficos, páginas escaneadas u otros objetos visuales. |
| `failed` | No se pudo extraer; usar `onedrive-export-file`. |

Ejemplo:

```json
{
  "path": "/Resumen Clases/Instrucciones_Resumen.md",
  "maxChars": 50000
}
```

El texto aparece en `structuredContent.data.text` en el transporte remoto.

## `onedrive-export-file`

**Scope:** `outlook:read`
**Efecto:** descarga y entrega bytes como `EmbeddedResource` MCP. Está limitado
por `OUTLOOK_FILE_INLINE_MAX_BYTES` (8 MiB por defecto).

| Parámetro | Tipo | Requerido |
|---|---|---:|
| `itemId` | string | uno de `itemId`, `path`, `fileId` |
| `path` | string | uno de `itemId`, `path`, `fileId` |
| `fileId` | string | uno de `itemId`, `path`, `fileId` |

Para MIME textual, el recurso usa `text`. Para binarios, usa `blob` Base64.
Usarlo para imágenes, audio, ZIP, formatos Office antiguos o documentos donde la
extracción devuelva `failed`.

## `onedrive-upload`

**Scope:** `outlook:write`
**Efecto:** crea o reemplaza un archivo pequeño en OneDrive. La implementación
está pensada para archivos menores de 4 MiB.

| Parámetro | Tipo | Requerido | Valores |
|---|---|---:|---|
| `path` | string | sí | Ruta destino incluyendo nombre. |
| `content` | string | uno de `content`/`contentBase64`/`file` | Texto UTF-8 que se subirá. |
| `contentBase64` | string | uno de `content`/`contentBase64`/`file` | Bytes binarios codificados en Base64 estándar, por ejemplo un PDF. |
| `file` | object | uno de `content`/`contentBase64`/`file` | Archivo seleccionado o subido en ChatGPT; incluye `download_url` y `file_id`. |
| `conflictBehavior` | string | no | `rename`, `replace` o `fail`; predeterminado `rename`. |

## `onedrive-upload-large`

**Scope:** `outlook:write`
**Efecto:** sube un archivo grande usando una upload session y chunks.

Acepta los mismos parámetros que `onedrive-upload`: `path`, `content`,
`contentBase64` o `file`, y `conflictBehavior` (`rename`, `replace` o `fail`).
El campo `file` usa el soporte de ChatGPT para parámetros de archivo y se
descarga directamente al servidor antes de iniciar la sesión de Graph.

El objeto `file` tiene esta forma:

```json
{
  "download_url": "https://files.oaiusercontent.com/...",
  "file_id": "file_...",
  "mime_type": "application/pdf",
  "file_name": "document.pdf"
}
```

La descarga solo permite HTTPS, los hosts `files.oaiusercontent.com`,
`files.openaiusercontent.com` y `oaisdmnt*.blob.core.windows.net`, redirects
limitados y el máximo configurado por `OUTLOOK_ONEDRIVE_IMPORT_MAX_BYTES`.

## `onedrive-upload-session-start`

**Scope:** `outlook:write`

Inicia una carga binaria reanudable. Requiere `path` y `totalBytes`, y devuelve
un `uploadId` y el tamaño máximo recomendado de cada chunk.

## `onedrive-upload-session-chunk`

**Scope:** `outlook:write`

Añade un chunk Base64 estándar. Requiere `uploadId`, `offset` y `chunkBase64`.
Los chunks deben enviarse en orden y `offset` debe coincidir con
`bytesReceived` de la respuesta anterior.

## `onedrive-upload-session-complete`

**Scope:** `outlook:write`

Finaliza la sesión y sube el temporal completo a OneDrive. Si faltan bytes, no
se ejecuta ninguna subida.

## `onedrive-upload-session-abort`

**Scope:** `outlook:write`

Cancela la sesión y elimina sus bytes temporales. Las sesiones abandonadas
expiran automáticamente.

## `onedrive-import-url`

**Scope:** `outlook:write`
**Efecto:** descarga una capability URL HTTPS aprobada por el servidor y sube el
resultado a OneDrive sin incrustar los bytes en la petición MCP.

| Parámetro | Tipo | Requerido | Valores |
|---|---|---:|---|
| `sourceUrl` | string | sí | URL HTTPS temporal de un servicio confiable. |
| `path` | string | sí | Ruta destino incluyendo nombre. |
| `conflictBehavior` | string | no | `rename`, `replace` o `fail`; predeterminado `rename`. |

## `onedrive-share`

**Scope:** `outlook:write`
**Efecto:** crea un enlace de uso compartido en OneDrive.

| Parámetro | Tipo | Requerido | Valores |
|---|---|---:|---|
| `itemId` | string | uno de `itemId`/`path` |
| `path` | string | uno de `itemId`/`path` |
| `type` | string | no | `view`, `edit` o `embed`; predeterminado `view`. |
| `scope` | string | no | `anonymous` u `organization`; predeterminado `anonymous`. |

Crear enlaces `anonymous` expone el archivo fuera del control OAuth del MCP.

## `onedrive-create-folder`

**Scope:** `outlook:write`
**Efecto:** crea una carpeta.

| Parámetro | Tipo | Requerido |
|---|---|---:|
| `name` | string | sí |
| `path` | string | no; carpeta padre, raíz por defecto |

## `onedrive-move`

**Scope:** `outlook:write`
**Efecto:** mueve y/o renombra un archivo o carpeta sin descargarlo.

| Parámetro | Tipo | Requerido | Descripción |
|---|---|---:|---|
| `itemId` | string | uno de `itemId`/`path` | ID del item. |
| `path` | string | uno de `itemId`/`path` | Ruta del item. |
| `destinationPath` | string | uno de `destinationPath`/`newName` | Carpeta destino; `/` o `root` para raíz. |
| `newName` | string | uno de `destinationPath`/`newName` | Nuevo nombre. |

## `onedrive-delete`

**Scope:** `outlook:destructive`
**Efecto:** elimina un archivo o carpeta de OneDrive. No se debe usar sin
confirmación explícita.

| Parámetro | Tipo | Requerido |
|---|---|---:|
| `itemId` | string | uno de `itemId`/`path` |
| `path` | string | uno de `itemId`/`path` |

# Protocolo De Archivos

Las tools de lectura/exportación guardan temporalmente los bytes bajo un nombre
opaco. Devuelven un URI como:

```text
m365-file:///0123456789abcdef0123456789abcdef
```

Los clientes MCP que soportan recursos pueden llamar:

```json
{
  "method": "resources/read",
  "params": {
    "uri": "m365-file:///0123456789abcdef0123456789abcdef"
  }
}
```

`resources/read` devuelve `text` para MIME textuales y `blob` Base64 para
binarios. Los handles son temporales y locales al proceso; no son referencias
durables. El servidor elimina los archivos cuando expiran.

La descarga aplica HTTPS obligatorio, allowlist de hosts, límite por archivo,
límite total, verificación de tamaño y SHA-256. Al cruzar de Graph a un CDN se
elimina el header `Authorization`.

## Configuración De Archivos

| Variable | Predeterminado | Uso |
|---|---:|---|
| `OUTLOOK_FILE_STORE_DIR` | `${DATA_DIR}/files` | Directorio de staging. |
| `OUTLOOK_FILE_MAX_BYTES` | 100 MiB | Límite por archivo. |
| `OUTLOOK_FILE_MAX_TOTAL_BYTES` | 512 MiB | Cuota total en staging. |
| `OUTLOOK_FILE_INLINE_MAX_BYTES` | 8 MiB | Límite de `EmbeddedResource`. |
| `OUTLOOK_FILE_RETENTION_MS` | 1 hora | Vida del handle. |
| `OUTLOOK_FILE_DOWNLOAD_TIMEOUT_MS` | 120 s | Timeout de descarga. |
| `OUTLOOK_FILE_DOWNLOAD_ALLOWED_HOSTS` | Graph + SharePoint/OneDrive CDN | Allowlist de redirects. |

La cuenta del servicio debe ser propietaria de `OUTLOOK_FILE_STORE_DIR` con
permisos de directorio `0700`; los archivos deben quedar con `0600`.

## No Confundir

| Necesidad | Tool/protocolo correcto |
|---|---|
| Saber si el archivo existe | `onedrive-search` o `onedrive-list` |
| Obtener contenido Markdown/texto | `onedrive-read-file` |
| Transferir imagen/audio/ZIP/binario | `onedrive-export-file` |
| Entregar bytes a otro backend | `onedrive-download` |
| Leer un recurso ya descargado | `resources/read` |
| Compartir públicamente | `onedrive-share` |
