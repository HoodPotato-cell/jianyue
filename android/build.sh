#!/usr/bin/env bash
# 简阅 1.4 构建：替换 WebView 资产 + 改版本号 + 对齐 + 签名
set -euo pipefail
cd "$(dirname "$0")"

SRC="${SRC:-/sdcard/Download/QQ/base.apk}"
KS="${JY_KS:-build/jianyue.jks}"
OUT="build/简阅-1.4.apk"

rm -rf build/stage build/app.apk build/aligned.apk "$OUT"
mkdir -p build/stage/assets/web/js build/stage/assets/web/css

cp assets/web/js/app.js assets/web/js/ole.js assets/web/js/doc97.js \
   assets/web/js/xls97.js assets/web/js/ppt97.js build/stage/assets/web/js/
cp assets/web/css/app.css build/stage/assets/web/css/

python3 - "$SRC" <<'PY'
import sys, zipfile
open('build/AndroidManifest.xml', 'wb').write(zipfile.ZipFile(sys.argv[1]).read('AndroidManifest.xml'))
PY

python3 tools/patch_manifest.py build/AndroidManifest.xml build/stage/AndroidManifest.xml 5 1.4

cp "$SRC" build/app.apk
(cd build/stage && zip -q -X ../app.apk AndroidManifest.xml \
  assets/web/js/app.js assets/web/js/ole.js assets/web/js/doc97.js \
  assets/web/js/xls97.js assets/web/js/ppt97.js assets/web/css/app.css)

zipalign -f -p 4 build/app.apk build/aligned.apk
apksigner sign --ks "$KS" --ks-key-alias jianyue \
  --ks-pass env:JY_KS_PASS --key-pass env:JY_KEY_PASS \
  --out "$OUT" build/aligned.apk

apksigner verify --print-certs "$OUT" | head -4
aapt2 dump badging "$OUT" | head -1
ls -la "$OUT"
