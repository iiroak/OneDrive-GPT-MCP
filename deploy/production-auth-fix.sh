#!/bin/sh
set -eu

APP=/opt/mcp/outlook
STATE=/var/lib/mcp/outlook
STAMP=$(date -u +%Y%m%dT%H%M%SZ)
BACKUP="$STATE/deploy-backups/$STAMP"
DEPLOYED=0

if [ "$(id -u)" -ne 0 ]; then
  echo 'This script must run as root' >&2
  exit 1
fi

rollback() {
  if [ "$DEPLOYED" -ne 1 ]; then
    return
  fi
  echo 'Deployment failed; restoring previous authentication files' >&2
  cp -p "$BACKUP/index.js" "$APP/auth/index.js"
  cp -p "$BACKUP/tools.js" "$APP/auth/tools.js"
  if [ -f "$BACKUP/token-storage-instance.js" ]; then
    cp -p "$BACKUP/token-storage-instance.js" "$APP/auth/token-storage-instance.js"
  else
    rm -f "$APP/auth/token-storage-instance.js"
  fi
  systemctl restart mcp-outlook.service || true
}

trap rollback EXIT

install -d -m 700 "$BACKUP"
cp -p "$APP/auth/index.js" "$BACKUP/index.js"
cp -p "$APP/auth/tools.js" "$BACKUP/tools.js"
if [ -f "$APP/auth/token-storage-instance.js" ]; then
  cp -p "$APP/auth/token-storage-instance.js" "$BACKUP/token-storage-instance.js"
fi

TOKEN_HASH=$(sha256sum "$STATE/microsoft-tokens.json" 2>/dev/null || true)

rm -f "$APP/auth/index.js.new" "$APP/auth/tools.js.new" "$APP/auth/token-storage-instance.js.new"
rm -f "$APP/auth/index.js.tmp.js" "$APP/auth/tools.js.tmp.js" "$APP/auth/token-storage-instance.js.tmp.js"

node - "$APP" <<'NODE'
const fs = require('fs');
const path = require('path');

const app = process.argv[2];
const indexPath = path.join(app, 'auth/index.js');
const toolsPath = path.join(app, 'auth/tools.js');
const instancePath = path.join(app, 'auth/token-storage-instance.js');

function replaceOnce(source, before, after, label) {
  const matches = source.split(before).length - 1;
  if (matches !== 1) throw new Error(`${label}: expected one match, found ${matches}`);
  return source.replace(before, after);
}

let index = fs.readFileSync(indexPath, 'utf8');
index = replaceOnce(
  index,
  "const TokenStorage = require('./token-storage');",
  "const tokenStorage = require('./token-storage-instance');",
  indexPath
);
index = replaceOnce(
  index,
  "// Singleton TokenStorage instance for automatic token refresh\nconst tokenStorage = new TokenStorage();\n",
  '',
  indexPath
);
fs.writeFileSync(`${indexPath}.tmp.js`, index);

let tools = fs.readFileSync(toolsPath, 'utf8');
tools = replaceOnce(
  tools,
  "const tokenManager = require('./token-manager');",
  "const tokenManager = require('./token-manager');\nconst tokenStorage = require('./token-storage-instance');",
  toolsPath
);
const startMarker = '  const tokens = tokenManager.loadTokenCache();';
const endMarker = '  return {\n    content: [{ type: "text", text: "Authenticated and ready" }]\n  };';
const start = tools.indexOf(startMarker);
const end = tools.indexOf(endMarker, start);
if (start < 0 || end < 0 || end <= start) {
  throw new Error(`${toolsPath}: authentication status handler not found`);
}
const statusBody = [
  '  const accessToken = await tokenStorage.getValidAccessToken();',
  '',
  "  console.error('[CHECK-AUTH-STATUS] Access token available: ' + (accessToken ? 'YES' : 'NO'));",
  '',
  '  if (!accessToken) {',
  "    console.error('[CHECK-AUTH-STATUS] No valid access token found');",
  '    return {',
  '      content: [{ type: "text", text: "Not authenticated" }]',
  '    };',
  '  }',
  '',
  "  console.error('[CHECK-AUTH-STATUS] Access token present');",
  "  console.error('[CHECK-AUTH-STATUS] Token expires at: ' + tokenStorage.getExpiryTime());",
  "  console.error('[CHECK-AUTH-STATUS] Current time: ' + Date.now());",
  ''
].join('\n');
tools = tools.slice(0, start) + statusBody + tools.slice(end);
fs.writeFileSync(`${toolsPath}.tmp.js`, tools);
fs.writeFileSync(`${instancePath}.tmp.js`, "const TokenStorage = require('./token-storage');\n\nmodule.exports = new TokenStorage();\n");
NODE

node --check "$APP/auth/index.js.tmp.js"
node --check "$APP/auth/tools.js.tmp.js"
node --check "$APP/auth/token-storage-instance.js.tmp.js"

for file in index.js tools.js; do
  chmod --reference="$APP/auth/$file" "$APP/auth/$file.tmp.js"
  chown --reference="$APP/auth/$file" "$APP/auth/$file.tmp.js"
done
chmod 0644 "$APP/auth/token-storage-instance.js.tmp.js"
chown --reference="$APP/auth/token-storage.js" "$APP/auth/token-storage-instance.js.tmp.js"

mv "$APP/auth/index.js.tmp.js" "$APP/auth/index.js"
mv "$APP/auth/tools.js.tmp.js" "$APP/auth/tools.js"
mv "$APP/auth/token-storage-instance.js.tmp.js" "$APP/auth/token-storage-instance.js"
DEPLOYED=1

systemctl restart mcp-outlook.service
systemctl is-active --quiet mcp-outlook.service

ready=0
for _ in $(seq 1 20); do
  if curl -fsS http://127.0.0.1:8767/outlook/health >/dev/null; then
    ready=1
    break
  fi
  sleep 1
done
if [ "$ready" -ne 1 ]; then
  echo 'Outlook health endpoint did not become ready' >&2
  exit 1
fi

AFTER_TOKEN_HASH=$(sha256sum "$STATE/microsoft-tokens.json" 2>/dev/null || true)
if [ "$TOKEN_HASH" != "$AFTER_TOKEN_HASH" ]; then
  echo 'Token file changed unexpectedly' >&2
  exit 1
fi

DEPLOYED=0
printf '%s\n' "$BACKUP"
