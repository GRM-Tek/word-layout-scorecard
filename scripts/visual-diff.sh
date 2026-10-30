#!/usr/bin/env bash
# visual-diff.sh <reference.pdf> <test.pdf> <out-dir> [--private]
# Runs the pinned tree's own page-by-page comparison (`cargo xtask visual-diff`) sealed: no network, the two PDFs
# read-only, results in <out-dir>/result.json. Builds the tool's image the first time (a few minutes).
# --private writes numbers and line hashes only; use it for any real agency document.
set -euo pipefail
cd "$(dirname "$0")/.."
REF=$1; TEST=$2; OUT=$3; shift 3
TAG=studio-editor/wordcraft-xtask:latest
if ! docker image inspect "$TAG" >/dev/null 2>&1; then
  echo "building $TAG"; docker build --target xtask -t "$TAG" .
fi
IN=$(mktemp -d); cp "$REF" "$IN/reference.pdf"; cp "$TEST" "$IN/test.pdf"
mkdir -p "$OUT"
abs() { (cd "$1" && pwd -W 2>/dev/null || pwd); }
MSYS_NO_PATHCONV=1 docker run --rm --network none --cap-drop ALL --security-opt no-new-privileges --read-only --tmpfs /tmp \
  -v "$(abs "$IN"):/in:ro" -v "$(abs "$OUT"):/out" "$TAG" visual-diff /in/reference.pdf /in/test.pdf /out "$@"
rm -rf "$IN"
