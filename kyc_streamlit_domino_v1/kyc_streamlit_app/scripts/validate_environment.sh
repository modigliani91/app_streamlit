#!/usr/bin/env bash
set -euo pipefail
APP_ROOT="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")/.." && pwd)"
cd "${APP_ROOT}"
exec "${KYC_PYTHON_BIN:-python}" scripts/validate_environment.py "$@"
