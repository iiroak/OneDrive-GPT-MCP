/**
 * Configuration for Outlook MCP Server
 */
const path = require('path');
const os = require('os');
const crypto = require('crypto');
const fs = require('fs');

// Ensure we have a home directory path even if process.env.HOME is undefined
const homeDir = process.env.HOME || process.env.USERPROFILE || os.homedir() || '/tmp';
const dataDir = process.env.OUTLOOK_DATA_DIR
  ? path.resolve(process.env.OUTLOOK_DATA_DIR.replace(/^~/, homeDir))
  : path.join(homeDir, '.local', 'share', 'outlook-mcp');

function loadAdminOverrides() {
  if (typeof fs.readFileSync !== 'function') return {};
  try {
    const key = fs.readFileSync(path.join(dataDir, 'admin-config.key'));
    const stored = JSON.parse(fs.readFileSync(path.join(dataDir, 'admin-config.json'), 'utf8'));
    const values = { ...(stored.values || {}) };
  for (const field of ['ms_client_secret', 'mcp_token', 'oauth_secret', 'transcription_delivery_token']) {
      if (!values[field]) continue;
      const payload = JSON.parse(values[field]);
      const decipher = crypto.createDecipheriv('aes-256-gcm', key, Buffer.from(payload.iv, 'base64url'));
      decipher.setAuthTag(Buffer.from(payload.tag, 'base64url'));
      values[field] = Buffer.concat([
        decipher.update(Buffer.from(payload.data, 'base64url')),
        decipher.final()
      ]).toString('utf8');
    }
    return values;
  } catch (error) {
    if (error.code !== 'ENOENT') console.warn(`Unable to load admin configuration: ${error.message}`);
    return {};
  }
}

const adminOverrides = loadAdminOverrides();
const override = (name, fallback) => Object.prototype.hasOwnProperty.call(adminOverrides, name)
  ? adminOverrides[name]
  : fallback;
const httpPort = Number.parseInt(adminOverrides.http_port || process.env.OUTLOOK_HTTP_PORT || '8767', 10);
const publicBaseUrl = (process.env.OUTLOOK_PUBLIC_BASE_URL || `http://127.0.0.1:${httpPort}/outlook`).replace(/\/+$/, '');
const publicHostname = (() => {
  try {
    return new URL(publicBaseUrl).hostname;
  } catch (error) {
    return '127.0.0.1';
  }
})();
const parseHostList = (value, fallback) => String(value || fallback)
  .split(',')
  .map(host => host.trim().toLowerCase().replace(/\.+$/, ''))
  .filter(Boolean);

module.exports = {
  // Server information
  SERVER_NAME: "m365-assistant",
  SERVER_VERSION: "2.1.0-chatgpt",
  
  // Authentication configuration
  AUTH_CONFIG: {
    clientId: override('ms_client_id', process.env.MS_CLIENT_ID || process.env.OUTLOOK_CLIENT_ID || ''),
    clientSecret: override('ms_client_secret', process.env.MS_CLIENT_SECRET || process.env.OUTLOOK_CLIENT_SECRET || ''),
    redirectUri: process.env.MS_REDIRECT_URI || `${publicBaseUrl}/microsoft/callback`,
    scopes: String(override('ms_scopes', process.env.MS_SCOPES || 'offline_access openid profile User.Read Mail.ReadWrite Mail.Send MailboxSettings.ReadWrite Calendars.ReadWrite Files.ReadWrite')).split(/\s+/).filter(Boolean),
    tokenStorePath: process.env.MS_TOKEN_STORE_PATH || path.join(dataDir, 'microsoft-tokens.json'),
  },
  MS_TENANT_ID: override('ms_tenant_id', process.env.MS_TENANT_ID || 'consumers'),
  AUTHORITY_URL: (process.env.MS_AUTHORITY_HOST || 'https://login.microsoftonline.com').replace(/\/+$/, ''),

  // ChatGPT remote MCP transport and OAuth resource configuration
  DATA_DIR: dataDir,
  HTTP_HOST: override('http_host', process.env.OUTLOOK_HTTP_HOST || '127.0.0.1'),
  HTTP_PORT: httpPort,
  PUBLIC_BASE_URL: publicBaseUrl,
  MCP_RESOURCE_URL: `${publicBaseUrl}/mcp`,
  MCP_BOOTSTRAP_TOKEN: override('mcp_token', process.env.OUTLOOK_MCP_TOKEN || process.env.MCP_TOKEN || ''),
  MCP_OAUTH_SECRET: override('oauth_secret', process.env.OUTLOOK_OAUTH_SECRET || process.env.MCP_OAUTH_SECRET || ''),
  MCP_ACCESS_TOKEN_TTL: Number.parseInt(process.env.OUTLOOK_ACCESS_TOKEN_TTL || '3600', 10),
  MCP_REFRESH_TOKEN_TTL: Number.parseInt(process.env.OUTLOOK_REFRESH_TOKEN_TTL || `${30 * 24 * 3600}`, 10),
  MCP_AUTHORIZATION_TTL: Number.parseInt(process.env.OUTLOOK_AUTHORIZATION_TTL || '600', 10),
  // Microsoft Graph API
  GRAPH_API_ENDPOINT: 'https://graph.microsoft.com/v1.0/',
  
  // Email constants
  EMAIL_SELECT_FIELDS: 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,hasAttachments,importance,isRead',
  EMAIL_DETAIL_FIELDS: 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,bodyPreview,body,hasAttachments,importance,isRead,internetMessageHeaders',
  
  // Calendar constants
  CALENDAR_SELECT_FIELDS: 'id,subject,bodyPreview,body,start,end,location,organizer,attendees,categories,isAllDay,isCancelled,type,seriesMasterId,recurrence,originalStart,changeKey,isOrganizer,isOnlineMeeting,onlineMeetingProvider,isReminderOn,reminderMinutesBeforeStart,importance,sensitivity,showAs,responseRequested,allowNewTimeProposals,hideAttendees',
  
  // Pagination
  DEFAULT_PAGE_SIZE: 25,
  MAX_RESULT_COUNT: 50,

  // Timezone
  DEFAULT_TIMEZONE: "Central European Standard Time",

  // OneDrive constants
  ONEDRIVE_SELECT_FIELDS: 'id,name,size,lastModifiedDateTime,webUrl,folder,file,parentReference',
  ONEDRIVE_UPLOAD_THRESHOLD: 4 * 1024 * 1024, // 4MB - files larger than this need chunked upload
  ONEDRIVE_UPLOAD_SESSION_CHUNK_BYTES: Number.parseInt(
    process.env.OUTLOOK_ONEDRIVE_UPLOAD_SESSION_CHUNK_BYTES || `${1024 * 1024}`,
    10
  ),
  ONEDRIVE_UPLOAD_SESSION_MAX_BYTES: Number.parseInt(
    process.env.OUTLOOK_ONEDRIVE_UPLOAD_SESSION_MAX_BYTES || `${512 * 1024 * 1024}`,
    10
  ),
  ONEDRIVE_UPLOAD_SESSION_TTL_MS: Number.parseInt(
    process.env.OUTLOOK_ONEDRIVE_UPLOAD_SESSION_TTL_MS || `${15 * 60 * 1000}`,
    10
  ),
  ONEDRIVE_IMPORT_ALLOWED_HOSTS: parseHostList(
    process.env.OUTLOOK_ONEDRIVE_IMPORT_ALLOWED_HOSTS,
    publicHostname
  ),
  CHATGPT_FILE_ALLOWED_HOSTS: parseHostList(
    process.env.OUTLOOK_CHATGPT_FILE_ALLOWED_HOSTS,
    'files.oaiusercontent.com,files.openaiusercontent.com'
  ),
  ONEDRIVE_IMPORT_MAX_BYTES: Number.parseInt(
    process.env.OUTLOOK_ONEDRIVE_IMPORT_MAX_BYTES || `${200 * 1024 * 1024}`,
    10
  ),
  ONEDRIVE_DOWNLOAD_CAPABILITY_TTL: Number.parseInt(
    process.env.OUTLOOK_ONEDRIVE_DOWNLOAD_CAPABILITY_TTL || `${30 * 60}`,
    10
  ),
  // File transfer / staging. Files are kept under opaque ids and consumed via
  // text extraction, resources/read, or an inline EmbeddedResource.
  FILE_MAX_BYTES: Number.parseInt(
    process.env.OUTLOOK_FILE_MAX_BYTES || `${100 * 1024 * 1024}`,
    10
  ),
  FILE_MAX_TOTAL_BYTES: Number.parseInt(
    process.env.OUTLOOK_FILE_MAX_TOTAL_BYTES || `${512 * 1024 * 1024}`,
    10
  ),
  FILE_INLINE_MAX_BYTES: Number.parseInt(
    process.env.OUTLOOK_FILE_INLINE_MAX_BYTES || `${8 * 1024 * 1024}`,
    10
  ),
  FILE_RETENTION_MS: Number.parseInt(
    process.env.OUTLOOK_FILE_RETENTION_MS || `${60 * 60 * 1000}`,
    10
  ),
  FILE_DOWNLOAD_TIMEOUT_MS: Number.parseInt(
    process.env.OUTLOOK_FILE_DOWNLOAD_TIMEOUT_MS || '120000',
    10
  ),
  FILE_DOWNLOAD_ALLOWED_HOSTS: parseHostList(
    process.env.OUTLOOK_FILE_DOWNLOAD_ALLOWED_HOSTS,
    'graph.microsoft.com,*.sharepoint.com,*.sharepointusercontent.com,*.onedrive.com,onedrive.live.com,*.onedriveusercontent.com,*.microsoftpersonalcontent.com,*.1drv.com,*.svc.ms'
  ),
  TRANSCRIPTION_DELIVERY_TOKEN: override(
    'transcription_delivery_token',
    process.env.OUTLOOK_TRANSCRIPTION_DELIVERY_TOKEN || ''
  ),
  TRANSCRIPTION_DELIVERY_MAX_OUTPUTS: Number.parseInt(
    process.env.OUTLOOK_TRANSCRIPTION_DELIVERY_MAX_OUTPUTS || '8',
    10
  ),

};
