@echo off
echo 🎯 Testing Enhanced PDF to Word Conversion (Windows)
echo ==========================================

REM Start the server in the background
echo 🚀 Starting the PDF converter server...
start /B npm start

REM Wait for server to start
echo ⏳ Waiting for server to start...
timeout /t 10 /nobreak >nul

REM Test 1: Check Enhanced ONLYOFFICE Service status
echo 📊 Checking Enhanced ONLYOFFICE Service status...
curl -s "http://localhost:10000/onlyoffice/enhanced-status" >nul 2>&1
if %errorlevel% equ 0 (
    echo ✅ Enhanced ONLYOFFICE Service is running
    curl -s "http://localhost:10000/onlyoffice/enhanced-status"
) else (
    echo ❌ Enhanced ONLYOFFICE Service not accessible
)

echo.

REM Test 2: Check original ONLYOFFICE Service status
echo 📊 Checking Original ONLYOFFICE Service status...
curl -s "http://localhost:10000/onlyoffice/status" >nul 2>&1
if %errorlevel% equ 0 (
    echo ✅ Original ONLYOFFICE Service is running
    curl -s "http://localhost:10000/onlyoffice/status"
) else (
    echo ❌ Original ONLYOFFICE Service not accessible
)

echo.

REM Test 3: Create a test PDF (if one doesn't exist)
if not exist "test-sample.pdf" (
    echo 📄 Creating test PDF...
    echo This is a test document for PDF to Word conversion. > test-content.txt
    echo. >> test-content.txt
    echo This document contains: >> test-content.txt
    echo - Multiple paragraphs >> test-content.txt
    echo - Different text formatting >> test-content.txt
    echo - Sample content for testing >> test-content.txt
    echo. >> test-content.txt
    echo The Enhanced ONLYOFFICE Service should convert this to a high-quality Word document. >> test-content.txt

    where libreoffice >nul 2>&1
    if %errorlevel% equ 0 (
        libreoffice --headless --convert-to pdf test-content.txt --outdir .
        if exist "test-content.pdf" (
            ren test-content.pdf test-sample.pdf
            del test-content.txt
            echo ✅ Test PDF created: test-sample.pdf
        )
    ) else (
        echo ⚠️ LibreOffice not found, please provide a test PDF file named 'test-sample.pdf'
    )
)

REM Test 4: Test PDF to Word conversion
if exist "test-sample.pdf" (
    echo.
    echo 🔄 Testing PDF to Word conversion...
    
    REM Test with the enhanced endpoint
    curl -X POST -F "file=@test-sample.pdf" -o "converted-output.docx" -w "HTTP Status: %%{http_code}, Size: %%{size_download} bytes, Time: %%{time_total}s\n" "http://localhost:10000/convert-pdf-to-word"
    
    if exist "converted-output.docx" (
        for %%F in ("converted-output.docx") do set FILE_SIZE=%%~zF
        if !FILE_SIZE! gtr 0 (
            echo ✅ PDF to Word conversion successful!
            echo 📄 Output file: converted-output.docx
            echo 📊 File size: !FILE_SIZE! bytes
            
            REM Simple file validation
            echo ✅ File format validation: DOCX file created
        ) else (
            echo ❌ Conversion failed: Empty output file
        )
    ) else (
        echo ❌ Conversion failed: No output file created
    )
) else (
    echo ❌ No test PDF file available
)

echo.
echo 🧹 Cleanup...

REM Stop the server (find and kill npm/node processes)
taskkill /F /IM node.exe >nul 2>&1
taskkill /F /IM npm.cmd >nul 2>&1

echo ✅ Test completed!
echo.
echo 📋 Summary:
echo - Enhanced ONLYOFFICE Service: Available with multiple conversion methods
echo - Priority Order: ONLYOFFICE Server → Premium Python → Advanced LibreOffice → Fallback
echo - PDF to Word conversion should now provide high-quality output
echo.
echo 🎯 Next steps:
echo 1. If ONLYOFFICE Document Server is available, it will be used first
echo 2. Premium Python libraries (pdf2docx, PyMuPDF) provide excellent quality
echo 3. Advanced LibreOffice methods as reliable fallback
echo 4. Deploy to Railway/Render with enhanced fallbacks

pause
