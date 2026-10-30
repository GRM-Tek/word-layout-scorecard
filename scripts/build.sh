#!/usr/bin/env bash
# Build the engine image from the pinned submodule. Refuses a dirty or moved submodule.
set -euo pipefail
cd "$(dirname "$0")/.."
PIN=$(git ls-files -s wordcraft | awk '{print $2}')
HEAD=$(git -C wordcraft rev-parse HEAD)
if [ -z "$PIN" ]; then echo "wordcraft is not a recorded submodule"; exit 1; fi
if [ "$PIN" != "$HEAD" ]; then echo "wordcraft checkout $HEAD is not the pin $PIN"; exit 1; fi
if [ -n "$(git -C wordcraft status --porcelain --untracked-files=no)" ]; then echo "wordcraft has local changes; commit a patch or reset"; exit 1; fi
TAG="studio-editor/wordcraft-cli:${PIN:0:8}"
echo "pin $PIN -> $TAG"
docker build --build-arg "WORDCRAFT_COMMIT=$PIN" -t "$TAG" -t studio-editor/wordcraft-cli:latest .
echo "built $TAG"
