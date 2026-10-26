#!/usr/bin/env bash
# Sobe o SupoClip 100% local (Whisper + Ollama, sem API paga) em http://localhost:3107
# Uso: bash rodar-supoclip.sh [modelo_ollama]   (padrão: qwen2.5:7b; PC fraco: llama3.2:3b)
set -euo pipefail

MODELO="${1:-qwen2.5:7b}"
DIR="${SUPOCLIP_DIR:-$HOME/supoclip}"

command -v docker >/dev/null || { echo "Instale o Docker Desktop e abra-o antes."; exit 1; }
docker info >/dev/null 2>&1 || { echo "Docker não está rodando. Abra o Docker Desktop."; exit 1; }
command -v ollama >/dev/null || { echo "Instale o Ollama: https://ollama.com/download"; exit 1; }
command -v git >/dev/null || { echo "Instale o git."; exit 1; }

# Ollama precisa escutar em todas as interfaces para o container enxergar (Linux).
if ! curl -fs http://localhost:11434/api/tags >/dev/null 2>&1; then
  echo "Iniciando o Ollama..."
  (OLLAMA_HOST=0.0.0.0 nohup ollama serve >/dev/null 2>&1 &)
  sleep 4
fi
ollama pull "$MODELO"

[ -d "$DIR/.git" ] || git clone https://github.com/FujiwaraChoki/supoclip "$DIR"
cd "$DIR"
[ -f .env ] || cp .env.example .env

# set_env CHAVE VALOR: troca a linha existente ou acrescenta no fim
set_env() {
  if grep -q "^$1=" .env; then
    sed -i.bak "s|^$1=.*|$1=$2|" .env
  else
    echo "$1=$2" >> .env
  fi
}
set_env TRANSCRIPTION_PROVIDER whisper
set_env WHISPER_MODEL_SIZE small
set_env TRANSCRIPTION_LANGUAGE pt
set_env LLM "ollama:$MODELO"
set_env SELF_HOST true
set_env FRONTEND_PORT 3107
set_env NEXT_PUBLIC_APP_URL http://localhost:3107
rm -f .env.bak

docker compose up -d --build

echo "Aguardando o frontend..."
for _ in $(seq 1 90); do
  curl -fs -o /dev/null http://localhost:3107 && break
  sleep 5
done
echo "Pronto: http://localhost:3107"
(xdg-open http://localhost:3107 || open http://localhost:3107) >/dev/null 2>&1 || true
