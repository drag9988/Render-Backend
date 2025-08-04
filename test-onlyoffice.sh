#!/bin/bash

# Test script for ONLYOFFICE integration
echo "🧪 Testing ONLYOFFICE Document Server Integration"
echo "================================================"

# Check if ONLYOFFICE server is running
echo "1. Checking ONLYOFFICE server availability..."
ONLYOFFICE_URL=${ONLYOFFICE_DOCUMENT_SERVER_URL:-"http://localhost:8000"}

if curl -s "$ONLYOFFICE_URL/healthcheck" > /dev/null; then
    echo "✅ ONLYOFFICE server is accessible at $ONLYOFFICE_URL"
else
    echo "❌ ONLYOFFICE server is not accessible at $ONLYOFFICE_URL"
    echo "💡 Try starting it with: docker-compose -f docker-compose.dev.yml up -d"
    exit 1
fi

# Check backend server
echo "2. Checking backend server..."
BACKEND_URL=${SERVER_URL:-"http://localhost:10000"}

if curl -s "$BACKEND_URL/health" > /dev/null; then
    echo "✅ Backend server is accessible at $BACKEND_URL"
else
    echo "❌ Backend server is not accessible at $BACKEND_URL"
    echo "💡 Try starting it with: npm run build && npm start"
    exit 1
fi

# Test ONLYOFFICE status endpoint
echo "3. Testing ONLYOFFICE status endpoint..."
ONLYOFFICE_STATUS=$(curl -s "$BACKEND_URL/onlyoffice/status" | jq -r '.onlyoffice.available // false')

if [ "$ONLYOFFICE_STATUS" = "true" ]; then
    echo "✅ ONLYOFFICE integration is working"
    curl -s "$BACKEND_URL/onlyoffice/status" | jq '.onlyoffice'
else
    echo "❌ ONLYOFFICE integration is not working"
    echo "📊 Status response:"
    curl -s "$BACKEND_URL/onlyoffice/status" | jq
fi

# Test ConvertAPI status (for comparison)
echo "4. Testing ConvertAPI status..."
CONVERTAPI_STATUS=$(curl -s "$BACKEND_URL/convertapi/status" | jq -r '.convertapi.available // false')

if [ "$CONVERTAPI_STATUS" = "true" ]; then
    echo "✅ ConvertAPI is also available (fallback)"
else
    echo "ℹ️  ConvertAPI is not available (will use LibreOffice as final fallback)"
fi

echo ""
echo "🎯 Integration Summary:"
echo "----------------------"
echo "ONLYOFFICE Available: $ONLYOFFICE_STATUS"
echo "ConvertAPI Available: $CONVERTAPI_STATUS"
echo ""
echo "📝 Conversion Priority:"
echo "1. ONLYOFFICE Document Server (Primary)"
echo "2. ConvertAPI (Fallback)"
echo "3. LibreOffice (Final fallback)"
echo ""
echo "🧪 To test PDF conversion:"
echo "curl -X POST -F 'file=@your-test.pdf' $BACKEND_URL/convert-pdf-to-word"
