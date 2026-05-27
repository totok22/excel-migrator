#!/bin/zsh
# macOS one-click launcher for Excel Migrator.
set -e
cd "$(dirname "$0")"

VENV_DIR=".venv"
PY=""

# Find a python interpreter (prefer 3.11+).
for candidate in python3.13 python3.12 python3.11 python3.10 python3 python; do
  if command -v "$candidate" >/dev/null 2>&1; then
    if "$candidate" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)' 2>/dev/null; then
      PY="$candidate"
      break
    fi
  fi
done

if [[ -z "$PY" ]]; then
  echo "未找到 Python 3.10+，请先安装：https://www.python.org/downloads/"
  read -k1 "?按任意键退出..."
  exit 1
fi

if [[ ! -d "$VENV_DIR" ]]; then
  echo "首次运行：建立虚拟环境..."
  "$PY" -m venv "$VENV_DIR"
fi

source "$VENV_DIR/bin/activate"

# Only install dependencies if requirements changed or never installed
STAMP="$VENV_DIR/.requirements.stamp"
if [[ ! -f "$STAMP" ]] || ! diff -q requirements.txt "$STAMP" >/dev/null 2>&1; then
  echo "安装/更新依赖..."
  python -m pip install --upgrade pip >/dev/null
  python -m pip install -r requirements.txt >/dev/null
  cp requirements.txt "$STAMP"
else
  echo "依赖已就绪，跳过安装。"
fi

export PYTHONPATH="$PWD/src:${PYTHONPATH:-}"
python -m excel_migrator.server "$@"
