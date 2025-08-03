# PowerShell script to test CORS fixes
Write-Host "🚀 Building and testing CORS fixes..." -ForegroundColor Green

# Build the project
Write-Host "📦 Building TypeScript..." -ForegroundColor Yellow
npm run build

if ($LASTEXITCODE -eq 0) {
    Write-Host "✅ Build successful!" -ForegroundColor Green
    
    # Start the server in background
    Write-Host "🌐 Starting server..." -ForegroundColor Yellow
    $serverJob = Start-Job -ScriptBlock { 
        Set-Location "c:\Users\saada\Downloads\PDF App Versions\render backend\pdf_converter_backend_render"
        npm start 
    }
    
    # Wait for server to start
    Start-Sleep -Seconds 8
    
    # Test CORS
    Write-Host "🧪 Testing CORS..." -ForegroundColor Yellow
    
    try {
        Write-Host "Testing GET /cors-test..." -ForegroundColor Cyan
        $response1 = Invoke-WebRequest -Uri "http://localhost:3000/cors-test" -Headers @{"Origin"="http://localhost:3000"} -Method GET
        Write-Host "Status: $($response1.StatusCode)" -ForegroundColor Green
        Write-Host "CORS Headers:" -ForegroundColor Yellow
        $response1.Headers | Where-Object {$_.Key -like "*Access-Control*"} | ForEach-Object { Write-Host "$($_.Key): $($_.Value)" }
        
        Write-Host "`nTesting POST /cors-test..." -ForegroundColor Cyan
        $body = '{"test": "data"}' | ConvertTo-Json
        $response2 = Invoke-WebRequest -Uri "http://localhost:3000/cors-test" -Headers @{"Origin"="http://localhost:3000"; "Content-Type"="application/json"} -Method POST -Body $body
        Write-Host "Status: $($response2.StatusCode)" -ForegroundColor Green
        
        Write-Host "`nTesting /health endpoint..." -ForegroundColor Cyan
        $response3 = Invoke-WebRequest -Uri "http://localhost:3000/health" -Headers @{"Origin"="http://localhost:3000"} -Method GET
        Write-Host "Status: $($response3.StatusCode)" -ForegroundColor Green
        
        Write-Host "`n✅ All CORS tests passed!" -ForegroundColor Green
        
    } catch {
        Write-Host "❌ CORS test failed: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # Stop the server
    Stop-Job $serverJob
    Remove-Job $serverJob
    Write-Host "`n🛑 Server stopped" -ForegroundColor Yellow
    
} else {
    Write-Host "❌ Build failed!" -ForegroundColor Red
}

Write-Host "`n📋 Next Steps:" -ForegroundColor Yellow
Write-Host "1. If tests pass, deploy to Railway" -ForegroundColor White
Write-Host "2. Update your frontend to use the Railway URL" -ForegroundColor White
Write-Host "3. Test from your frontend application" -ForegroundColor White
