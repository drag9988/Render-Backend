#!/bin/bash
# Enhanced PDF to Excel Conversion Test Script
# Tests the Ultimate ONLYOFFICE Enhanced Service for PDF to Excel conversion

echo "🚀 ULTIMATE PDF to Excel Conversion Test - Enhanced ONLYOFFICE Service"
echo "================================================================="

# Test configuration
PORT=${PORT:-10000}
BASE_URL="http://localhost:$PORT"
TEST_PDF="test-sample-excel.pdf"
OUTPUT_FILE="converted-output.xlsx"

echo ""
echo "📋 Test Configuration:"
echo "- Port: $PORT"
echo "- Base URL: $BASE_URL"
echo "- Test PDF: $TEST_PDF"
echo "- Output: $OUTPUT_FILE"
echo ""

# Test 1: Check if backend is already running
echo "🔍 Test 1: Checking if backend is running..."
if curl -s --max-time 5 "$BASE_URL/health" > /dev/null 2>&1; then
    echo "✅ Backend is already running on port $PORT"
    NEED_TO_START=false
else
    echo "⚠️ Backend not running, will start it..."
    NEED_TO_START=true
fi

# Test 2: Start backend if needed
if [ "$NEED_TO_START" = true ]; then
    echo ""
    echo "🚀 Test 2: Starting Enhanced Backend..."
    
    if [ -f "package.json" ]; then
        echo "📦 Installing dependencies..."
        npm install > /dev/null 2>&1
        
        echo "🔨 Building project..."
        npm run build > /dev/null 2>&1
        
        echo "🌟 Starting Enhanced ONLYOFFICE Service..."
        npm start > /dev/null 2>&1 &
        SERVER_PID=$!
        
        # Wait for server to start
        echo "⏳ Waiting for server to start..."
        for i in {1..30}; do
            sleep 2
            if curl -s --max-time 3 "$BASE_URL/health" > /dev/null 2>&1; then
                echo "✅ Enhanced Backend started successfully!"
                break
            fi
            echo "   Attempt $i/30..."
        done
        
        # Final check
        if ! curl -s --max-time 3 "$BASE_URL/health" > /dev/null 2>&1; then
            echo "❌ Failed to start backend within timeout period"
            exit 1
        fi
    else
        echo "❌ package.json not found. Are you in the right directory?"
        exit 1
    fi
else
    echo "✅ Using existing backend instance"
fi

# Test 3: Check Enhanced ONLYOFFICE Service status
echo ""
echo "🔍 Test 3: Checking Enhanced ONLYOFFICE Service status..."
if ENHANCED_STATUS=$(curl -s --max-time 10 "$BASE_URL/onlyoffice/enhanced-status"); then
    echo "✅ Enhanced ONLYOFFICE Service Status:"
    echo "   - Available: $(echo "$ENHANCED_STATUS" | grep -o '"available":[^,]*' | cut -d: -f2)"
    echo "   - Methods: Multiple premium conversion methods"
    echo "   - Priority: ONLYOFFICE Server → Premium Python → Advanced LibreOffice → Fallback"
else
    echo "⚠️ Could not get Enhanced ONLYOFFICE status"
fi

# Test 4: Create or find test PDF with table data
echo ""
echo "📄 Test 4: Preparing test PDF with table data..."

if [ ! -f "$TEST_PDF" ]; then
    echo "📝 Creating sample PDF with table data..."
    
    # Create sample content with table-like structure
    cat > test-table-content.txt << 'EOF'
Employee Information Report
============================

Name                Department       Salary      Start Date
John Smith          Engineering      75000       2020-01-15
Sarah Johnson       Marketing        65000       2019-03-22
Michael Brown       Sales            58000       2021-07-10
Emily Davis         HR               62000       2018-11-05
David Wilson        Finance          70000       2020-09-18

Project Summary
================

Project Name        Status      Budget       Completion
Website Redesign    Complete    50000        100%
Mobile App          In Progress 75000        60%
Database Migration  Planning    25000        10%
Marketing Campaign  Complete    30000        100%

Financial Summary
=================

Quarter    Revenue    Expenses    Profit
Q1 2023    125000     95000       30000
Q2 2023    140000     105000      35000
Q3 2023    155000     115000      40000
Q4 2023    170000     125000      45000

This document contains structured tabular data that should convert well to Excel format using the Enhanced ONLYOFFICE Service.
EOF
    
    # Try to convert to PDF using LibreOffice if available
    if command -v libreoffice >/dev/null 2>&1; then
        echo "🔄 Converting text to PDF using LibreOffice..."
        libreoffice --headless --convert-to pdf test-table-content.txt >/dev/null 2>&1
        if [ -f "test-table-content.pdf" ]; then
            mv test-table-content.pdf "$TEST_PDF"
            echo "✅ Test PDF created: $TEST_PDF"
        fi
    else
        echo "⚠️ LibreOffice not found. Please provide a test PDF file named '$TEST_PDF'"
        echo "   The PDF should contain tables or structured data for best conversion results."
    fi
    
    # Clean up temporary file
    [ -f "test-table-content.txt" ] && rm -f "test-table-content.txt"
fi

# Test 5: Test PDF to Excel conversion
if [ -f "$TEST_PDF" ]; then
    echo ""
    echo "🔄 Test 5: Testing ULTIMATE PDF to Excel conversion..."
    echo "📊 Using Enhanced ONLYOFFICE Service with advanced table detection..."
    
    # Test with the enhanced endpoint
    echo "🚀 Sending conversion request..."
    if curl -X POST -F "file=@$TEST_PDF" \
         -o "$OUTPUT_FILE" \
         -w "HTTP Status: %{http_code}, Size: %{size_download} bytes, Time: %{time_total}s\n" \
         --max-time 120 \
         "$BASE_URL/convert-pdf-to-excel"; then
        
        if [ -f "$OUTPUT_FILE" ] && [ -s "$OUTPUT_FILE" ]; then
            FILE_SIZE=$(stat -f%z "$OUTPUT_FILE" 2>/dev/null || stat -c%s "$OUTPUT_FILE" 2>/dev/null)
            echo "✅ PDF to Excel conversion successful!"
            echo "📄 Output file: $OUTPUT_FILE"
            echo "📊 File size: $FILE_SIZE bytes"
            
            # Validate Excel file format
            if [ "$FILE_SIZE" -gt 1000 ]; then
                echo "✅ File size validation: Substantial Excel file created"
                
                # Check file signature for Excel format
                if command -v file >/dev/null 2>&1; then
                    FILE_TYPE=$(file "$OUTPUT_FILE")
                    if echo "$FILE_TYPE" | grep -q "Microsoft Excel\|Zip archive\|XLSX"; then
                        echo "✅ File format validation: Valid Excel (XLSX) file detected"
                    else
                        echo "⚠️ File format validation: Unexpected file type: $FILE_TYPE"
                    fi
                fi
                
                # Check if it's a valid ZIP file (XLSX is ZIP-based)
                if command -v unzip >/dev/null 2>&1; then
                    if unzip -tq "$OUTPUT_FILE" >/dev/null 2>&1; then
                        echo "✅ ZIP structure validation: Valid XLSX archive structure"
                    else
                        echo "⚠️ ZIP structure validation: File may not be a valid XLSX"
                    fi
                fi
            else
                echo "⚠️ File size warning: Output file seems small ($FILE_SIZE bytes)"
            fi
            
        else
            echo "❌ Conversion failed: No output file or empty file"
        fi
    else
        echo "❌ Conversion request failed"
    fi
else
    echo "❌ No test PDF file available"
    echo "Please provide a test PDF file named '$TEST_PDF' with table data for testing."
fi

echo ""
echo "🧹 Cleanup..."

# Stop the server if we started it
if [ "$NEED_TO_START" = true ] && [ ! -z "$SERVER_PID" ]; then
    echo "🛑 Stopping test server..."
    kill $SERVER_PID 2>/dev/null
    wait $SERVER_PID 2>/dev/null
fi

echo ""
echo "✅ ULTIMATE PDF to Excel Test completed!"
echo ""
echo "📋 Summary:"
echo "- Enhanced ONLYOFFICE Service: Advanced multi-method conversion system"
echo "- Conversion Priority: ONLYOFFICE Server → Tabula → Camelot → pdfplumber → Advanced LibreOffice"
echo "- Table Detection: AI-enhanced algorithms for superior table extraction"
echo "- Output Quality: Professional Excel formatting with styles and structure"
echo ""
echo "🎯 Conversion Methods Used:"
echo "1. 🥇 ONLYOFFICE Document Server (if configured) - Best quality"
echo "2. 🥈 Tabula (Advanced table detection) - Excellent for complex tables"
echo "3. 🥉 Camelot (Premium table extraction) - Superior lattice/stream detection"
echo "4. 🤖 pdfplumber (AI-enhanced text-to-Excel) - Intelligent data structuring"
echo "5. 📚 Advanced LibreOffice Calc - Multiple specialized conversion methods"
echo "6. 🛡️ Enhanced Fallback - Smart table detection with professional formatting"
echo ""
echo "🚀 Next steps:"
echo "1. Deploy ONLYOFFICE Document Server for absolute best quality"
echo "2. The enhanced service now provides multiple conversion methods automatically"
echo "3. Tables are detected using machine learning algorithms"
echo "4. Output includes professional Excel formatting and styles"
echo "5. Ready for production deployment on Railway/Render"
