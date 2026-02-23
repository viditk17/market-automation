# Market Automation Server - Easy Launcher

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "Market Automation - Office Network Setup" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Get local IP
$ip = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object {$_.InterfaceAlias -notmatch "Loopback"} | Select-Object -First 1).IPAddress

if ($null -eq $ip) {
    $ip = "localhost"
}

Write-Host "Step 1: Installing dependencies..." -ForegroundColor Yellow
python -m pip install -r requirements.txt --quiet
Write-Host "[DONE] All packages ready" -ForegroundColor Green
Write-Host ""

Write-Host "Step 2: Starting server..." -ForegroundColor Yellow
Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "✅ Server is running!" -ForegroundColor Green
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host ""
Write-Host "📍 Your IP Address: $ip" -ForegroundColor Yellow
Write-Host ""
Write-Host "🔗 Colleagues can access at:" -ForegroundColor Green
Write-Host "   http://$($ip):5000" -ForegroundColor Cyan
Write-Host ""
Write-Host "💻 Or on this machine:" -ForegroundColor Green
Write-Host "   http://localhost:5000" -ForegroundColor Cyan
Write-Host ""
Write-Host "⚠️  Press Ctrl+C to stop the server" -ForegroundColor Red
Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host ""

python app4.py
pause
