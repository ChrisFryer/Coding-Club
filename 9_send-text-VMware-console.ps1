Add-Type -AssemblyName System.Windows.Forms

# The command to send
$command = @"
Insert whatever you want to push to the VMware Console here
"@

Write-Host "INSTRUCTIONS:" -ForegroundColor Yellow
Write-Host "1. Click on your VMware console window to make it active" -ForegroundColor Cyan
Write-Host "2. Make sure you have permission to execute your command e.g. root command prompt" -ForegroundColor Cyan
Write-Host "3. Press any key here to start the 5-second countdown" -ForegroundColor Cyan
Write-Host "4. Quickly switch back to your VMware console" -ForegroundColor Cyan

$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

Write-Host "`nStarting countdown..." -ForegroundColor Red
for ($i = 5; $i -gt 0; $i--) {
    Write-Host "$i..." -ForegroundColor Red
    Start-Sleep -Seconds 1
}

Write-Host "Sending command now!" -ForegroundColor Green

# Send the command to whatever window is currently active
[System.Windows.Forms.SendKeys]::SendWait($command)
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

Write-Host "`nCommand sent! Check your VMware console." -ForegroundColor Green
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")