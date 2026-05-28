#!/bin/bash
# Abre o sistema de Solicitação de Emprego no navegador
cd "$(dirname "$0")"

echo "╔══════════════════════════════════════════════════╗"
echo "║   SISTEMA DE SOLICITAÇÃO DE EMPREGO - FIBRA      ║"
echo "╚══════════════════════════════════════════════════╝"
echo ""
echo "Abrindo no navegador em http://localhost:8501"
echo "Para fechar, pressione Ctrl+C"
echo ""

streamlit run app.py \
  --server.headless false \
  --browser.gatherUsageStats false \
  --theme.primaryColor "#1f4e79" \
  --theme.backgroundColor "#ffffff" \
  --theme.secondaryBackgroundColor "#f0f4f8"
