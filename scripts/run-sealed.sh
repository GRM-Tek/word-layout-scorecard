#!/usr/bin/env bash
# run-sealed.sh <in-dir> <out-dir> [fonts-dirs] -- <wordcraft-cli args...>
# Runs the engine with no network, <in-dir> read-only at /in, <out-dir> writable at /out, and the optional
# <fonts-dirs> (one or more folders separated by ';', for example "C:/Windows/Fonts") read-only under
# /usr/share/fonts/truetype/host*. Nothing else of this computer is visible to the container.
set -euo pipefail
IN=$1; OUT=$2; shift 2
FONTS=""
if [ "${1:-}" != "--" ]; then FONTS=$1; shift; fi
[ "${1:-}" = "--" ] && shift
mkdir -p "$OUT"
TAG=${WORDCRAFT_IMAGE:-studio-editor/wordcraft-cli:latest}
abs() { (cd "$1" && pwd -W 2>/dev/null || pwd); }
ARGS=(--rm --network none --cap-drop ALL --security-opt no-new-privileges --read-only --tmpfs /tmp
      -v "$(abs "$IN"):/in:ro" -v "$(abs "$OUT"):/out")
n=0
IFS=';' read -ra DIRS <<< "$FONTS"
for d in "${DIRS[@]}"; do
  [ -n "$d" ] || continue
  ARGS+=(-v "$(abs "$d"):/usr/share/fonts/truetype/host$n:ro"); n=$((n + 1))
done
MSYS_NO_PATHCONV=1 exec docker run "${ARGS[@]}" "$TAG" "$@"
