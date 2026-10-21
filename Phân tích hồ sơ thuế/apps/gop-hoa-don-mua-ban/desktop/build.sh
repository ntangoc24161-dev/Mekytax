#!/bin/sh
set -eu
HERE="$(CDPATH= cd -- "$(dirname "$0")" && pwd)"
ROOT="$(dirname "$HERE")"
WEB="$HERE/web"
rm -rf "$WEB"
mkdir -p "$WEB/vendor"
cp "$ROOT/index.html" "$ROOT/app.js" "$ROOT/core.js" "$ROOT/excel-out.js" "$WEB/"
cp "$ROOT/vendor/exceljs.min.js" "$ROOT/vendor/jszip.min.js" "$WEB/vendor/"
GO="${GO:-go}"
export GOOS=windows
export GOARCH=amd64
export CGO_ENABLED=0
cd "$HERE"
"$GO" build -trimpath -ldflags="-s -w -H windowsgui" -o "$ROOT/GopHoaDon.exe" .
echo "XONG $ROOT/GopHoaDon.exe"
ls -lh "$ROOT/GopHoaDon.exe"
