FROM node:22-bookworm-slim

ENV NODE_ENV=production \
    OUTLOOK_HTTP_HOST=0.0.0.0 \
    OUTLOOK_HTTP_PORT=8767 \
    OUTLOOK_PUBLIC_BASE_URL=https://mcp.example.com/outlook \
    OUTLOOK_DATA_DIR=/data

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends ca-certificates curl \
    && rm -rf /var/lib/apt/lists/* \
    && mkdir -p /data \
    && chown node:node /app /data

COPY --chown=node:node package.json pnpm-lock.yaml ./
RUN corepack pnpm install --frozen-lockfile --prod

COPY --chown=node:node . .

USER node

EXPOSE 8767

HEALTHCHECK --interval=30s --timeout=5s --start-period=30s --retries=3 \
    CMD path="$(node -e 'const url = new URL(process.env.OUTLOOK_PUBLIC_BASE_URL); process.stdout.write(url.pathname.replace(/\/+$/, ""))')/health" \
    && curl -fsS "http://127.0.0.1:${OUTLOOK_HTTP_PORT}${path}" >/dev/null || exit 1

CMD ["node", "server.js"]
