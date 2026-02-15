#!/bin/bash
# Set up Python virtual environment for pst-to-markdown skill
set -e

SKILL_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "==> Setting up pst-to-markdown Python environment..."

# Create venv if it doesn't exist
if [ ! -d "$SKILL_DIR/.venv" ]; then
    python3 -m venv "$SKILL_DIR/.venv"
    echo "Created virtual environment"
fi

# Install/upgrade dependencies
echo "Installing dependencies..."
"$SKILL_DIR/.venv/bin/pip" install --upgrade pip -q
"$SKILL_DIR/.venv/bin/pip" install -r "$SKILL_DIR/requirements.txt" -q

echo "==> pst-to-markdown environment ready"
echo
echo "Dependencies installed:"
"$SKILL_DIR/.venv/bin/pip" list --format=columns 2>/dev/null | grep -iE "libratom|html2text|dateutil|tqdm" || true
echo
echo "System tool check:"
if command -v readpst &>/dev/null; then
    echo "  readpst: $(readpst -V 2>&1 | head -1)"
else
    echo "  readpst: NOT FOUND (optional - install pst-utils for fallback extraction)"
    echo "    Ubuntu/Debian: sudo apt install pst-utils"
    echo "    macOS: brew install libpst"
fi
