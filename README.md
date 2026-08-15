# OneDrive-GPT-MCP

Servidor MCP para ChatGPT que conecta Outlook y OneDrive con Microsoft Graph.
Expone transporte Streamable HTTP con OAuth para clientes remotos y transporte
stdio para desarrollo local.

El nombre del proyecto es branding: el servidor mantiene tanto las funciones de
Outlook como las de OneDrive.

## Funcionalidades

- Correo: listar, buscar, leer, redactar, enviar, marcar y mover mensajes.
- Calendario: calendarios, eventos, categorías, invitaciones y migraciones.
- Carpetas y reglas: listar, crear, mover mensajes y administrar reglas.
- OneDrive: listar, buscar, leer, descargar, subir, exportar, compartir y mover
  archivos y carpetas.
- OAuth remoto MCP/PKCE y OAuth delegado de Microsoft Graph.
- Transferencia segura de archivos mediante recursos MCP temporales.
- Annotations y scopes separados para operaciones de lectura, escritura y
  destructivas.

## Transporte

El servidor remoto utiliza Streamable HTTP:

```text
https://mcp.iroak.dev/outlook/mcp
```

El endpoint público y la ruta `/outlook` forman parte de la configuración actual
de producción. Los ejemplos de `deploy/` documentan el servicio systemd, nginx y
el adaptador de administración MCP.

Para ejecutar localmente:

```bash
npm install
npm run start:http
```

El servidor stdio está disponible con:

```bash
npm start
```

## Estructura

```text
├── index.js                 # Registro de tools y servidor MCP stdio
├── server.js                # Transporte HTTP, OAuth remoto y health endpoint
├── admin.js                 # Adaptador de administración del servicio
├── config.js                # Configuración de Graph, OAuth y archivos
├── AGENTS.md                # Guía operativa para cambios en el repositorio
├── auth/                    # OAuth MCP y almacenamiento de tokens Graph
├── calendar/                # Calendarios y eventos de Outlook
├── email/                   # Operaciones de correo
├── folder/                  # Carpetas y movimiento de mensajes
├── rules/                   # Reglas de bandeja de entrada
├── onedrive/                # Archivos, carpetas y transferencia OneDrive
├── utils/                   # Graph API, staging y extracción de documentos
├── deploy/                  # Ejemplos de systemd, nginx y MCP Admin
└── docs/TOOLS.md            # Referencia completa de tools y parámetros
```

## Configuración local

```bash
cp .env.example .env
```

Configura las credenciales de una aplicación Microsoft Entra registrada para
Microsoft Graph. El secreto debe ser el **valor** del secreto, no su identificador.

Permisos delegados principales:

- `offline_access`
- `User.Read`
- `Mail.ReadWrite`
- `Mail.Send`
- `MailboxSettings.ReadWrite`
- `Calendars.ReadWrite`
- `Files.ReadWrite`

La configuración de producción se almacena fuera del repositorio mediante MCP
Admin. No se deben commitear credenciales, tokens, claves OAuth ni archivos `.env`.

## Tools

La lista completa de tools, parámetros, scopes y efectos está en
[`docs/TOOLS.md`](docs/TOOLS.md).

Las operaciones remotas se clasifican con estos scopes:

| Scope | Uso |
|---|---|
| `outlook:read` | Lecturas de Outlook, calendario, reglas y OneDrive |
| `outlook:write` | Crear, editar, mover, subir o compartir |
| `outlook:destructive` | Borrado o acciones explícitamente destructivas |

Las tools de escritura y destrucción incluyen annotations MCP para que el cliente
pueda solicitar confirmación antes de cambiar o borrar datos.

## Transferencia De Archivos

`onedrive-read-file` extrae texto de archivos compatibles como texto plano,
Markdown, HTML, PDF, DOCX, PPTX y XLSX.

`onedrive-export-file` permite transferir bytes sin interpretar para imágenes,
audio, archivos comprimidos y documentos no compatibles.

Los archivos se guardan temporalmente con identificadores opacos y se exponen como
recursos `m365-file:///...`. Los clientes MCP compatibles pueden leerlos mediante
`resources/read`. Las descargas requieren HTTPS, allowlist de hosts, límites de
tamaño y verificación SHA-256.

## Desarrollo

Comandos disponibles:

```bash
npm install
npm start
npm run start:http
npm run inspect
```

Este despliegue independiente no incluye una suite automatizada. La verificación
mínima consiste en cargar `index.js` y `server.js`, iniciar el servidor HTTP y
comprobar `/health`.

## Agradecimientos

Este proyecto comenzó a partir de [ryaker/outlook-mcp](https://github.com/ryaker/outlook-mcp).
Se conserva la atribución correspondiente en el historial y en la licencia MIT.

## Licencia

MIT. Consulta [`LICENSE`](LICENSE).
