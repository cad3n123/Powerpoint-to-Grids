#!/bin/bash
set -e

echo "==== Grid Maker Builder (macOS) ===="

STARTDIR="$(pwd)"
TEMPDIR="$STARTDIR/build_temp"

if [[ ! -f "grid_maker.py" ]]; then
    echo "Error: grid_maker.py not found. Please save the python code first."
    exit 1
fi

# 1. Clean previous build
rm -rf "$TEMPDIR"
mkdir -p "$TEMPDIR"
cd "$TEMPDIR"

# 2. Check for Homebrew (Install if missing)
if ! command -v brew &>/dev/null; then
    echo "Installing Homebrew..."
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)" </dev/null
fi

# 3. Install System Dependencies (LibreOffice + Poppler)
echo "Checking system dependencies..."
brew install poppler
brew install --cask libreoffice

# 4. Setup Python
PYTHON_EXEC="$(which python3)"
echo "Using Python: $PYTHON_EXEC"
"$PYTHON_EXEC" -m venv venv
source venv/bin/activate

# 5. Install Python Libraries
pip install --upgrade pip
pip install py2app python-pptx pdf2image

# 6. Prepare Setup
cp "$STARTDIR/grid_maker.py" .
cat <<EOF > setup.py
from setuptools import setup
APP = ['grid_maker.py']
OPTIONS = {
    'argv_emulation': True,
    'packages': ['python-pptx', 'pdf2image', 'pptx']
}
setup(
    app=APP,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
EOF

# 7. Build App
python setup.py py2app

# 8. Move App and Clean
mv dist/grid_maker.app "$STARTDIR"
cd "$STARTDIR"
rm -rf "$TEMPDIR"

echo
echo "==== Build Complete ===="
echo "Created: grid_maker.app"
echo "Note: LibreOffice and Poppler were installed via Homebrew."