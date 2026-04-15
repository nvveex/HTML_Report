#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
SKILL_SRC="$ROOT_DIR/skill"
CODEX_HOME_DIR="${CODEX_HOME:-$HOME/.codex}"
SKILL_DEST="$CODEX_HOME_DIR/skills/seiue-usage-report"

if [ ! -d "$SKILL_SRC" ]; then
  echo "未找到 skill 源目录：$SKILL_SRC" >&2
  exit 1
fi

mkdir -p "$CODEX_HOME_DIR/skills"
rm -rf "$SKILL_DEST"
mkdir -p "$SKILL_DEST"
rsync -a --delete "$SKILL_SRC"/ "$SKILL_DEST"/

echo "seiue-usage-report 已安装/更新到：$SKILL_DEST"
echo "现在可直接在 Codex 中调用：\$seiue-usage-report"
