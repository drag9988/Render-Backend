#!/bin/bash

echo "🎯 Testing Enhanced PDF to Word Conversion"
echo "=========================================="

# Start the server in the background
echo "🚀 Starting the PDF converter server..."
npm start &
SERVER_PID=$!

# Wait for server to start
echo "⏳ Waiting for server to start..."
sleep 10

# Test 1: Check Enhanced ONLYOFFICE Service status
echo "📊 Checking Enhanced ONLYOFFICE Service status..."
if curl -s "http://localhost:10000/onlyoffice/enhanced-status" > /dev/null; then
    echo "✅ Enhanced ONLYOFFICE Service is running"
    curl -s "http://localhost:10000/onlyoffice/enhanced-status" | head -10
else
    echo "❌ Enhanced ONLYOFFICE Service not accessible"
fi

echo ""

# Test 2: Check original ONLYOFFICE Service status
echo "📊 Checking Original ONLYOFFICE Service status..."
if curl -s "http://localhost:10000/onlyoffice/status" > /dev/null; then
    echo "✅ Original ONLYOFFICE Service is running"
    curl -s "http://localhost:10000/onlyoffice/status" | head -10
else
    echo "❌ Original ONLYOFFICE Service not accessible"
fi

echo ""

# Test 3: Create a test PDF (if one doesn't exist)
if [ ! -f "test-sample.pdf" ]; then
    echo "📄 Creating test PDF..."
    # Create a simple test PDF using LibreOffice
    echo "This is a test document for PDF to Word conversion.

This document contains:
- Multiple paragraphs
- Different text formatting
- Sample content for testing

The Enhanced ONLYOFFICE Service should convert this to a high-quality Word document." > test-content.txt

    if command -v libreoffice >/dev/null 2>&1; then
        libreoffice --headless --convert-to pdf test-content.txt --outdir .
        mv test-content.pdf test-sample.pdf
        rm test-content.txt
        echo "✅ Test PDF created: test-sample.pdf"
    else
        echo "⚠️ LibreOffice not found, please provide a test PDF file named 'test-sample.pdf'"
    fi
fi

# Test 4: Test PDF to Word conversion
if [ -f "test-sample.pdf" ]; then
    echo ""
    echo "🔄 Testing PDF to Word conversion..."
    
    # Test with the enhanced endpoint
    if curl -X POST -F "file=@test-sample.pdf" \
         -o "converted-output.docx" \
         -w "HTTP Status: %{http_code}, Size: %{size_download} bytes, Time: %{time_total}s\n" \
         "http://localhost:10000/convert-pdf-to-word"; then
        
        if [ -f "converted-output.docx" ] && [ -s "converted-output.docx" ]; then
            FILE_SIZE=$(stat -f%z "converted-output.docx" 2>/dev/null || stat -c%s "converted-output.docx" 2>/dev/null)
            echo "✅ PDF to Word conversion successful!"
            echo "📄 Output file: converted-output.docx"
            echo "📊 File size: $FILE_SIZE bytes"
            
            # Validate DOCX file format
            if file "converted-output.docx" | grep -q "Microsoft Word\|Zip archive\|DOCX"; then
                echo "✅ File format validation: Valid DOCX file"
            else
                echo "⚠️ File format validation: May not be a valid DOCX file"
                echo "   File type: $(file converted-output.docx)"
            fi
        else
            echo "❌ Conversion failed: No output file or empty file"
        fi
    else
        echo "❌ Conversion request failed"
    fi
else
    echo "❌ No test PDF file available"
fi

echo ""
echo "🧹 Cleanup..."

# Stop the server
kill $SERVER_PID 2>/dev/null
wait $SERVER_PID 2>/dev/null

echo "✅ Test completed!"
echo ""
echo "📋 Summary:"
echo "- Enhanced ONLYOFFICE Service: Available with multiple conversion methods"
echo "- Priority Order: ONLYOFFICE Server → Premium Python → Advanced LibreOffice → Fallback"
echo "- PDF to Word conversion should now provide high-quality output"
echo ""
echo "🎯 Next steps:"
echo "1. If ONLYOFFICE Document Server is available, it will be used first"
echo "2. Premium Python libraries (pdf2docx, PyMuPDF) provide excellent quality"
echo "3. Advanced LibreOffice methods as reliable fallback"
echo "4. Deploy to Railway/Render with enhanced fallbacks"
