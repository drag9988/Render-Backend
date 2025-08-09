# PowerShell script to test PDF to PowerPoint conversion locally
# This helps debug Python script issues on Windows

Write-Host "🚀 Starting local PDF to PowerPoint conversion test..." -ForegroundColor Green
Write-Host ""

# Configuration
$testPdfPath = ".\test.pdf"
$outputDir = ".\test-output"
$pythonPath = if ($env:PYTHON_PATH) { $env:PYTHON_PATH } else { "python" }

try {
    # Create output directory
    if (!(Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        Write-Host "✅ Created output directory: $outputDir" -ForegroundColor Green
    }

    # Check if test PDF exists
    if (Test-Path $testPdfPath) {
        $pdfSize = (Get-Item $testPdfPath).Length
        Write-Host "📄 Found test PDF: $testPdfPath ($pdfSize bytes)" -ForegroundColor Green
    } else {
        Write-Host "❌ Test PDF not found. Please place a PDF file named 'test.pdf' in the current directory." -ForegroundColor Red
        Write-Host "💡 You can use any PDF file you want to test with." -ForegroundColor Yellow
        return
    }

    # Test Python availability
    Write-Host ""
    Write-Host "🐍 Testing Python environment..." -ForegroundColor Cyan
    try {
        $pythonVersion = & $pythonPath --version 2>&1
        Write-Host "✅ Python version: $pythonVersion" -ForegroundColor Green
    } catch {
        Write-Host "❌ Python not found at: $pythonPath" -ForegroundColor Red
        Write-Host "💡 Try setting PYTHON_PATH environment variable or install Python" -ForegroundColor Yellow
        return
    }

    # Test Python packages
    Write-Host ""
    Write-Host "📦 Testing Python packages..." -ForegroundColor Cyan
    $requiredPackages = @("pdf2docx", "PyMuPDF", "pdfplumber", "python-pptx", "openpyxl")
    
    foreach ($pkg in $requiredPackages) {
        $importName = $pkg.Replace("-", "_").Replace("PyMuPDF", "fitz")
        try {
            $result = & $pythonPath -c "import $importName; print('$pkg available')" 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✅ $pkg`: Available" -ForegroundColor Green
            } else {
                Write-Host "❌ $pkg`: Not available" -ForegroundColor Red
            }
        } catch {
            Write-Host "❌ $pkg`: Error - $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Run the Node.js test script
    Write-Host ""
    Write-Host "🚀 Running detailed test..." -ForegroundColor Cyan
    Write-Host "Command: node test-local-conversion.js" -ForegroundColor Gray
    Write-Host ""
    
    if (Test-Path "test-local-conversion.js") {
        node "test-local-conversion.js"
    } else {
        Write-Host "❌ test-local-conversion.js not found. Make sure you're in the correct directory." -ForegroundColor Red
    }

    Write-Host ""
    Write-Host "🎯 Next Steps:" -ForegroundColor Yellow
    Write-Host "1. Install missing Python packages: pip install pdf2docx PyMuPDF pdfplumber python-pptx openpyxl" -ForegroundColor White
    Write-Host "2. If packages are installed but still failing, check the detailed logs above" -ForegroundColor White
    Write-Host "3. Try the server conversion and check the enhanced logs" -ForegroundColor White

} catch {
    Write-Host "❌ Test failed: $($_.Exception.Message)" -ForegroundColor Red
}
