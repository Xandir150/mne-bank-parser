#!/bin/bash
# Валидация выгруженных XML по официальным XSD: ./validate.sh <fi.xml> <pd.xml>
set -e
DIR="$(cd "$(dirname "$0")/../.." && pwd)"
[ -n "$1" ] && xmllint --schema "$DIR/xsd/fi.xsd" "$1" --noout && echo "FI: valid"
[ -n "$2" ] && xmllint --schema "$DIR/xsd/pd.xsd" "$2" --noout && echo "PD: valid"
