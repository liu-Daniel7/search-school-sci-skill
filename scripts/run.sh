#!/usr/bin/env bash
# 自动挑选可用的 Python（优先 DSH 自带运行时，含 pandas/openpyxl）
set -euo pipefail
CFG="${1:-../reference/config.example.json}"
OUT="${2:-college_sci.xlsx}"
pick() {
  for c in "${DSH_PYTHON:-}" \
           "$(ls -d "$HOME"/.dsh/dsh-runtimes/*/dependencies/python/bin/python3 2>/dev/null | head -1)" \
           "$(command -v python3 || true)"; do
    [ -n "${c:-}" ] && [ -x "$c" ] && "$c" -c "import openpyxl, pandas" 2>/dev/null && { echo "$c"; return; }
  done
}
PY="$(pick || true)"
[ -z "${PY:-}" ] && { echo "未找到含 openpyxl+pandas 的 Python，请通过 load_workspace_dependencies 获取解释器后手动执行" >&2; exit 1; }
echo "使用 $PY"
exec "$PY" "$(dirname "$0")/college_sci.py" --config "$CFG" --out "$OUT" --cache "$(dirname "$0")/../.cache"
