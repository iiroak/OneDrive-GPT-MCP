const crypto = require('crypto');
const fs = require('fs').promises;
const path = require('path');

const ADMIN_PREFIX = '/internal-admin/v1';
const SECRET_FIELDS = new Set([
  'ms_client_secret',
  'mcp_token',
  'oauth_secret',
  'transcription_delivery_token'
]);
const CONFIG_FIELDS = {
  microsoft: new Set(['ms_client_id', 'ms_client_secret', 'ms_tenant_id', 'ms_scopes']),
  security: new Set(['mcp_token', 'oauth_secret', 'transcription_delivery_token']),
  runtime: new Set(['http_host', 'http_port'])
};

function manifest(config) {
  return {
    protocol: 'iroak.mcp-admin/v1',
    service: {
      id: 'outlook',
      name: 'Outlook and OneDrive',
      version: config.SERVER_VERSION,
      description: 'MCP remoto para Outlook y OneDrive con cuenta Microsoft personal.',
      public_url: config.PUBLIC_BASE_URL + '/mcp'
    },
    config: [
      {
        id: 'microsoft',
        title: 'Cuenta Microsoft',
        description: 'Configuración de la aplicación Entra y permisos Graph delegados.',
        schema: {
          type: 'object',
          properties: {
            ms_client_id: { type: 'string', title: 'Client ID', minLength: 1 },
            ms_client_secret: { type: 'string', title: 'Client secret', writeOnly: true, minLength: 1 },
            ms_tenant_id: { type: 'string', title: 'Tenant', minLength: 1 },
            ms_scopes: { type: 'string', title: 'Scopes', minLength: 1 }
          }
        }
      },
      {
        id: 'security',
        title: 'Seguridad MCP',
        description: 'Credenciales que invalidan clientes ChatGPT existentes al rotarse.',
        schema: {
          type: 'object',
          properties: {
            mcp_token: { type: 'string', title: 'Token bootstrap MCP', writeOnly: true, minLength: 32 },
            oauth_secret: { type: 'string', title: 'Clave de firma OAuth', writeOnly: true, minLength: 32 },
            transcription_delivery_token: { type: 'string', title: 'Token de entrega de transcripciones', writeOnly: true, minLength: 1 }
          }
        }
      },
      {
        id: 'runtime',
        title: 'Endpoint remoto',
        description: 'URL pública y listener local del transporte MCP.',
        schema: {
          type: 'object',
          properties: {
            http_host: { type: 'string', title: 'Host local', minLength: 1 },
            http_port: { type: 'string', title: 'Puerto local', minLength: 1 }
          }
        }
      }
    ],
    actions: [
      { id: 'validate', title: 'Validar configuración', description: 'Comprueba que la configuración requerida está presente.' },
      { id: 'restart', title: 'Reiniciar MCP', description: 'Aplica los cambios de configuración.', confirm: true, impact: { requires_restart: true } },
      { id: 'rotate-oauth', title: 'Rotar firma OAuth', description: 'Genera una nueva clave e invalida clientes existentes.', confirm: true, impact: { destructive: true, requires_restart: true, invalidates_oauth_clients: true } }
    ]
  };
}

function encrypted(value, key) {
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv('aes-256-gcm', key, iv);
  const ciphertext = Buffer.concat([cipher.update(value, 'utf8'), cipher.final()]);
  return JSON.stringify({
    v: 1,
    iv: iv.toString('base64url'),
    tag: cipher.getAuthTag().toString('base64url'),
    data: ciphertext.toString('base64url')
  });
}

function decrypted(value, key) {
  const payload = JSON.parse(value);
  const decipher = crypto.createDecipheriv('aes-256-gcm', key, Buffer.from(payload.iv, 'base64url'));
  decipher.setAuthTag(Buffer.from(payload.tag, 'base64url'));
  return Buffer.concat([
    decipher.update(Buffer.from(payload.data, 'base64url')),
    decipher.final()
  ]).toString('utf8');
}

class AdminStore {
  constructor(config) {
    this.config = config;
    this.file = path.join(config.DATA_DIR, 'admin-config.json');
    this.keyFile = path.join(config.DATA_DIR, 'admin-config.key');
    this.values = {};
    this.updated = {};
    this.deleted = {};
    this.key = null;
  }

  async load() {
    await fs.mkdir(this.config.DATA_DIR, { recursive: true, mode: 0o700 });
    try {
      this.key = await fs.readFile(this.keyFile);
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
      this.key = crypto.randomBytes(32);
      await fs.writeFile(this.keyFile, this.key, { mode: 0o600 });
    }
    try {
      const stored = JSON.parse(await fs.readFile(this.file, 'utf8'));
      this.values = stored.values || {};
      this.updated = stored.updated || {};
      this.deleted = stored.deleted || {};
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
    }
  }

  async save() {
    const temporary = `${this.file}.tmp`;
    await fs.writeFile(temporary, JSON.stringify({ values: this.values, updated: this.updated, deleted: this.deleted }, null, 2), { mode: 0o600 });
    await fs.rename(temporary, this.file);
  }

  value(field) {
    const value = this.values[field];
    if (value === undefined) return '';
    return SECRET_FIELDS.has(field) ? decrypted(value, this.key) : value;
  }

  has(field) {
    return Object.prototype.hasOwnProperty.call(this.values, field);
  }

  status(field) {
    return { configured: !this.deleted[field] && Boolean(this.values[field]), updated_at: this.updated[field] || null };
  }

  async set(field, value) {
    this.values[field] = SECRET_FIELDS.has(field) ? encrypted(value, this.key) : value;
    this.updated[field] = new Date().toISOString();
    delete this.deleted[field];
    await this.save();
  }

  async delete(field) {
    this.values[field] = SECRET_FIELDS.has(field) ? encrypted('', this.key) : '';
    this.updated[field] = new Date().toISOString();
    this.deleted[field] = true;
    await this.save();
  }
}

function createAdminController(config, options = {}) {
  const store = options.store || new AdminStore(config);
  const prefix = options.prefix || ADMIN_PREFIX;
  const expectedToken = () => process.env.MCP_ADMIN_OUTLOOK_TOKEN || '';
  const authorized = req => {
    const supplied = req.get('x-mcp-admin-token') || '';
    const expected = expectedToken();
    return Boolean(expected && supplied && supplied.length === expected.length && crypto.timingSafeEqual(Buffer.from(supplied), Buffer.from(expected)));
  };
  const guard = (req, res) => {
    if (!authorized(req)) {
      res.status(401).json({ error: 'no autorizado' });
      return false;
    }
    return true;
  };

  function secretStatus(name, fallback) {
    const stored = store.status(name);
    return { configured: stored.configured || Boolean(fallback), updated_at: stored.updated_at };
  }

  async function configResponse() {
    const value = (name, fallback = '') => SECRET_FIELDS.has(name)
      ? secretStatus(name, store.has(name) ? '' : fallback)
      : (store.has(name) ? store.value(name) : fallback);
    return {
      sections: {
        microsoft: {
          ms_client_id: value('ms_client_id', config.AUTH_CONFIG.clientId),
          ms_client_secret: value('ms_client_secret', config.AUTH_CONFIG.clientSecret),
          ms_tenant_id: value('ms_tenant_id', config.MS_TENANT_ID),
          ms_scopes: value('ms_scopes', config.AUTH_CONFIG.scopes.join(' '))
        },
        security: {
          mcp_token: value('mcp_token', config.MCP_BOOTSTRAP_TOKEN),
          oauth_secret: value('oauth_secret', config.MCP_OAUTH_SECRET),
          transcription_delivery_token: value('transcription_delivery_token', config.TRANSCRIPTION_DELIVERY_TOKEN)
        },
        runtime: {
          http_host: value('http_host', config.HTTP_HOST),
          http_port: value('http_port', String(config.HTTP_PORT))
        }
      }
    };
  }

  function current(field, fallback) {
    return store.has(field) ? store.value(field) : fallback;
  }

  function status() {
    const clientId = current('ms_client_id', config.AUTH_CONFIG.clientId);
    const clientSecret = current('ms_client_secret', config.AUTH_CONFIG.clientSecret);
    const mcpToken = current('mcp_token', config.MCP_BOOTSTRAP_TOKEN);
    const oauthSecret = current('oauth_secret', config.MCP_OAUTH_SECRET);
    const ready = Boolean(clientId && clientSecret && mcpToken && oauthSecret);
    return {
      state: ready ? 'healthy' : 'unconfigured',
      liveness: 'ok',
      readiness: ready ? 'ready' : 'not_ready',
      sources: { microsoft_configured: Boolean(clientId && clientSecret), security_configured: Boolean(mcpToken && oauthSecret) },
      message: ready ? 'Configuración local cargada.' : 'Faltan credenciales de Microsoft o seguridad MCP.'
    };
  }

  async function update(section, payload) {
    if (!CONFIG_FIELDS[section] || !payload || typeof payload !== 'object' || Array.isArray(payload)) throw new Error('sección o payload inválido');
    for (const [field, value] of Object.entries(payload)) {
      if (!CONFIG_FIELDS[section].has(field) || typeof value !== 'string' || !value.trim()) throw new Error('campo inválido');
      if ((field === 'mcp_token' || field === 'oauth_secret') && value.length < 32) throw new Error('los secretos MCP deben tener al menos 32 caracteres');
      await store.set(field, value);
    }
    return { state: 'stored', requires_restart: true };
  }

  async function remove(section, field) {
    if (!CONFIG_FIELDS[section] || !CONFIG_FIELDS[section].has(field)) throw new Error('campo no encontrado');
    await store.delete(field);
    return { state: 'deleted', requires_restart: true };
  }

  async function action(name) {
    if (name === 'validate') return status();
    if (name === 'restart') return { state: 'accepted', requires_restart: true, message: 'El broker del panel aplicará el reinicio.' };
    if (name === 'rotate-oauth') {
      await store.set('oauth_secret', crypto.randomBytes(48).toString('base64url'));
      return { state: 'rotated', requires_restart: true, invalidates_oauth_clients: true };
    }
    throw new Error('acción no encontrada');
  }

  return {
    store,
    async initialize() {
      await store.load();
    },
    routes: [
      { method: 'get', path: `${prefix}/manifest`, handler: async (req, res) => guard(req, res) && res.json(manifest(config)) },
      { method: 'get', path: `${prefix}/status`, handler: async (req, res) => guard(req, res) && res.json(status()) },
      { method: 'get', path: `${prefix}/config`, handler: async (req, res) => guard(req, res) && res.json(await configResponse()) },
      { method: 'patch', path: `${prefix}/config/:section`, handler: async (req, res) => guard(req, res) && res.json(await update(req.params.section, req.body)) },
      { method: 'delete', path: `${prefix}/config/:section/:field`, handler: async (req, res) => guard(req, res) && res.json(await remove(req.params.section, req.params.field)) },
      { method: 'post', path: `${prefix}/actions/:action`, handler: async (req, res) => guard(req, res) && res.json(await action(req.params.action)) },
      { method: 'get', path: `${prefix}/logs`, handler: async (req, res) => guard(req, res) && res.json({ lines: [], message: 'Los logs se consultan mediante systemd.' }) }
    ]
  };
}

module.exports = { AdminStore, createAdminController, manifest };
