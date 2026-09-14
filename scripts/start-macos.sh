#!/usr/bin/env bash
set -euo pipefail

project_dir="$(cd "$(dirname "$0")/.." && pwd)"
cd "$project_dir"

if [[ ! -f .env ]]; then
  cp .env.example .env
  encryption_key="$(openssl rand -hex 32)"
  sed -i '' "s/REPLACE_WITH_A_RANDOM_64_CHARACTER_HEX_VALUE/$encryption_key/" .env
  echo "Created .env with a new n8n encryption key."
fi

set -a
source .env
set +a

if [[ ! "${N8N_ENCRYPTION_KEY:-}" =~ ^[0-9a-fA-F]{64}$ ]]; then
  echo "N8N_ENCRYPTION_KEY in .env must contain exactly 64 hexadecimal characters." >&2
  exit 1
fi

if ! command -v docker >/dev/null 2>&1; then
  echo "Docker CLI is not installed. Install Docker Desktop first." >&2
  exit 1
fi
if ! docker info >/dev/null 2>&1; then
  echo "Docker Desktop is not running. Open it and retry." >&2
  exit 1
fi

mkdir -p data/output
docker compose up -d --build
docker compose ps
echo "Open http://localhost:${N8N_PORT:-5678}"
