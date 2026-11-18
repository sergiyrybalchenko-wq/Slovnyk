#!/usr/bin/env bash
# Конвертує вихідний DOC у Markdown: original/slovnyk-vykhidnyi.doc -> original/slovnyk-vykhidnyi.md
# Потрібні: LibreOffice (soffice), pandoc.
set -euo pipefail
cd "$(dirname "$0")/.."
tmp=$(mktemp -d)
cp original/slovnyk-vykhidnyi.doc "$tmp/s.doc"
soffice --headless --convert-to docx --outdir "$tmp" "$tmp/s.doc" >/dev/null
pandoc "$tmp/s.docx" -t gfm --wrap=none -o original/slovnyk-vykhidnyi.md
rm -rf "$tmp"
echo "Готово: original/slovnyk-vykhidnyi.md"
