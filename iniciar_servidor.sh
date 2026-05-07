#!/bin/bash
echo ""
echo " ============================================================"
echo "  CHUMACERO ESPADA S.R.L. - CONTROL DE INVENTARIOS v2.0"
echo " ============================================================"
echo ""
if ! command -v python3 &>/dev/null; then
    echo " [ERROR] Python3 no encontrado."
    echo " Instala con: sudo apt install python3 python3-pip"
    exit 1
fi
echo " Instalando dependencias..."
pip3 install flask openpyxl --quiet
IP=$(hostname -I | awk '{print $1}')
echo ""
echo " ============================================================"
echo "  Local:       http://localhost:5000"
echo "  En red WiFi: http://${IP}:5000"
echo "  Ctrl+C para detener"
echo " ============================================================"
echo ""
python3 app.py
