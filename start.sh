#!/bin/bash

# Print environment information
echo "Starting PDF Converter API..."
echo "Current directory: $(pwd)"
echo "Node version: $(node -v)"
echo "NPM version: $(npm -v)"
echo "PORT environment variable: ${PORT:-3000}"
echo "NODE_ENV: ${NODE_ENV:-development}"
echo "TEMP_DIR: ${TEMP_DIR:-/tmp/pdf-converter}"

# Check for LibreOffice and Ghostscript installations
echo "Checking for LibreOffice..."
libreoffice --version || echo "WARNING: LibreOffice not found or not working properly"

echo "Checking for Ghostscript..."
gs --version || echo "WARNING: Ghostscript not found or not working properly"

# Create temporary directory
TEMP_DIR=${TEMP_DIR:-/tmp/pdf-converter}
mkdir -p $TEMP_DIR
chmod 777 $TEMP_DIR
echo "Created temporary directory at $TEMP_DIR"

# List files in the current directory
echo "Files in current directory:"
ls -la

# Check if dist/main.js exists
if [ ! -f dist/main.js ]; then
  echo "ERROR: dist/main.js not found. Build may have failed."
  exit 1
fi

# Start the application
echo "Starting application on port ${PORT:-3000}..."
exec node dist/main.js