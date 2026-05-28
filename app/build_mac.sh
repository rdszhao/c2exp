#!/bin/bash
# Build script for Concept2 Exporter (macOS version for testing)

echo "============================================"
echo "Building Concept2 Logbook Exporter (macOS)"
echo "============================================"

# Check if Python is available
if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python 3 is not installed"
    exit 1
fi

# Create virtual environment if it doesn't exist
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
fi

# Activate virtual environment
source venv/bin/activate

# Install dependencies
echo "Installing dependencies..."
pip install -r requirements.txt

# Build executable
echo "Building executable..."
pyinstaller concept2_export.spec --clean

echo ""
echo "============================================"
echo "Build complete!"
echo "============================================"
echo ""
echo "The executable is at: dist/Concept2Exporter"
echo ""
