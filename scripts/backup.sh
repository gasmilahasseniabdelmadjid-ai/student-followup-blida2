#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
mkdir -p "$ROOT/backups"
ts="$(date +%Y%m%d_%H%M%S)"
tar -czf "$ROOT/backups/blida2_publications_$ts.tar.gz" -C "$ROOT" data
find "$ROOT/backups" -type f -mtime +30 -delete
echo "Backup created: $ROOT/backups/blida2_publications_$ts.tar.gz"
