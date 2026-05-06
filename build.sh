#!/bin/bash
set -e

PYINSTALLER="/Users/momo/ClaudeCode/MBANK-konvertor/.venv/bin/pyinstaller"

echo "🔧 Čistím predchádzajúci build..."
rm -rf build/ dist/

echo "📦 Budujem aplikáciu..."
$PYINSTALLER bank_converter.spec

echo "📂 Kopírujem aplikáciu do priečinka projektu..."
rm -rf "Bank Converter.app"
cp -r "dist/Bank Converter.app" "Bank Converter.app"

echo ""
echo "✅ Hotovo!"
echo "   Aplikácia: Bank Converter.app  (priamo v priečinku projektu)"
echo "   Spusti:    open 'Bank Converter.app'"
