#!/bin/sh
# Actualizeaza versiunea din adresele fisierelor (?v=...), ca browserul sa nu mai serveasca din cache
# codul vechi dupa o publicare. Ruleaza automat din .githooks/pre-commit; manual: sh tools/versiune.sh
set -e
cd "$(dirname "$0")/.."
v=$(date +%Y%m%d-%H%M%S)
perl -pi -e "s/\?v=[0-9A-Za-z._-]+/?v=$v/g" index.html assets/app.js assets/xlsx.js assets/store.js
echo "$v"
