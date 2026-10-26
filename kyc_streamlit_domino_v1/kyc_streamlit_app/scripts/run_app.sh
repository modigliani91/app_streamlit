#!/usr/bin/env bash
set -euo pipefail
umask 077
APP_ROOT="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")/.." && pwd)"
cd "${APP_ROOT}"
# La capture donne /opt/conda/envs/[MY VENV NAME]/bin/python : nom réel à configurer.
PYTHON_BIN="${KYC_PYTHON_BIN:-python}"
APP_PORT="${PORT:-8888}"
UPLOAD_MB="${KYC_MAX_UPLOAD_MB:-1024}"
[[ "${APP_PORT}" =~ ^[0-9]+$ ]] && (( APP_PORT >= 1 && APP_PORT <= 65535 )) || { echo "Port invalide" >&2; exit 1; }
[[ "${UPLOAD_MB}" =~ ^[1-9][0-9]*$ ]] || { echo "KYC_MAX_UPLOAD_MB invalide" >&2; exit 1; }
export HF_HUB_OFFLINE=1 TRANSFORMERS_OFFLINE=1 HF_HUB_DISABLE_TELEMETRY=1 DO_NOT_TRACK=1
export STREAMLIT_BROWSER_GATHER_USAGE_STATS=false
export TOKENIZERS_PARALLELISM=false
[[ "${KYC_CORE_CPU_TEST:-0}" != "1" ]] || { echo "Le mode CPU test ne peut pas lancer l'App." >&2; exit 1; }
"${PYTHON_BIN}" scripts/validate_environment.py --paths-only
command -v flock >/dev/null || { echo "flock (util-linux) est requis." >&2; exit 1; }
mkdir -p /mnt/output_qwen38/logs /mnt/output_qwen38/run_metadata /mnt/output_qwen38/reports
# Gardé pendant toute la vie du serveur, pas seulement pendant un job.
# Évite un second serveur qui conserverait sa propre copie du modèle.
exec flock --nonblock --no-fork /mnt/output_qwen38/.kyc_server.lock \
  "${PYTHON_BIN}" -m streamlit run app.py \
  --theme.base=dark \
  --server.fileWatcherType=none \
  --browser.gatherUsageStats=false \
  --server.address=0.0.0.0 \
  --server.port="${APP_PORT}" \
  --server.headless=true \
  --server.maxUploadSize="${UPLOAD_MB}" \
  --client.showErrorDetails=false
