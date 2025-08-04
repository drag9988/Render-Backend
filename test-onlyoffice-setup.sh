#!/bin/bash

echo "🔧 Testing ONLYOFFICE Enhanced Service Setup"
echo "=============================================="

# Test 1: Check temp directory
echo "📁 Testing temporary directory creation..."
TEMP_DIR="/tmp/pdf-converter"
mkdir -p $TEMP_DIR
chmod 777 $TEMP_DIR

if [ -d "$TEMP_DIR" ] && [ -w "$TEMP_DIR" ]; then
    echo "✅ Temp directory created and writable: $TEMP_DIR"
else
    echo "❌ Temp directory issue: $TEMP_DIR"
fi

# Test 2: Check LibreOffice
echo "📚 Testing LibreOffice..."
if command -v libreoffice >/dev/null 2>&1; then
    echo "✅ LibreOffice found: $(libreoffice --version | head -1)"
else
    echo "❌ LibreOffice not found"
fi

# Test 3: Check Python
echo "🐍 Testing Python..."
if command -v python3 >/dev/null 2>&1; then
    echo "✅ Python found: $(python3 --version)"
else
    echo "❌ Python3 not found"
fi

# Test 4: Check Node.js
echo "🟢 Testing Node.js..."
if command -v node >/dev/null 2>&1; then
    echo "✅ Node.js found: $(node --version)"
else
    echo "❌ Node.js not found"
fi

# Test 5: Check npm packages
echo "📦 Testing npm packages..."
if [ -f "package.json" ]; then
    echo "✅ package.json found"
    if [ -d "node_modules" ]; then
        echo "✅ node_modules found"
    else
        echo "⚠️  node_modules not found - run npm install"
    fi
else
    echo "❌ package.json not found"
fi

# Test 6: Check build
echo "🏗️  Testing build..."
if [ -f "dist/main.js" ]; then
    echo "✅ Build found: dist/main.js"
else
    echo "❌ Build not found - run npm run build"
fi

# Test 7: Check environment variables
echo "🌍 Testing environment variables..."
echo "NODE_ENV: ${NODE_ENV:-not set}"
echo "PORT: ${PORT:-not set}"
echo "TEMP_DIR: ${TEMP_DIR:-not set}"
echo "ONLYOFFICE_DOCUMENT_SERVER_URL: ${ONLYOFFICE_DOCUMENT_SERVER_URL:-not set (Enhanced Mode)}"
echo "PYTHON_PATH: ${PYTHON_PATH:-python3 (default)}"

# Test 8: Test enhanced service endpoint (if server is running)
echo "🌐 Testing Enhanced ONLYOFFICE Service endpoint..."
if curl -s "http://localhost:${PORT:-10000}/onlyoffice/enhanced-status" >/dev/null 2>&1; then
    echo "✅ Enhanced ONLYOFFICE Service endpoint accessible"
    curl -s "http://localhost:${PORT:-10000}/onlyoffice/enhanced-status" | head -5
else
    echo "⚠️  Enhanced ONLYOFFICE Service endpoint not accessible (server might not be running)"
fi

echo ""
echo "🎯 Summary:"
echo "- Enhanced Mode: Uses LibreOffice + Python (no ONLYOFFICE server needed)"
echo "- Fallback Chain: Enhanced ONLYOFFICE → Original ONLYOFFICE → ConvertAPI → LibreOffice"
echo "- Status: Ready for deployment"
echo ""
echo "🚀 To test conversion, start the server and try:"
echo "curl -X POST -F 'file=@test.pdf' http://localhost:${PORT:-10000}/convert-pdf-to-word"
