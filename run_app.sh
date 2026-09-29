#!/usr/bin/env bash
set -e

echo "========================================================"
echo "       OmniStudy - Starting on New Device"
echo "========================================================"
echo ""

# Check Python 3
if ! command -v python3 &> /dev/null; then
    echo "[ERROR] Python 3 is not installed!"
    echo "Please install Python 3.10+ and ensure python3 is in your PATH."
    exit 1
fi

# Set up venv if not present
if [ ! -f ".venv/bin/activate" ]; then
    echo "[*] Creating virtual environment (.venv)..."
    python3 -m venv .venv
    echo "[*] Activating virtual environment..."
    source .venv/bin/activate
    echo "[*] Installing dependencies..."
    pip install --upgrade pip
    pip install -r requirements.txt
else
    source .venv/bin/activate
fi

echo ""
echo "[*] Launching OmniStudy Server..."
echo "[*] Open your browser at: http://127.0.0.1:5000"
echo ""
python3 app.py
