$ErrorActionPreference = "Stop"

$projectRoot = $PSScriptRoot
$frontendDir = Join-Path $projectRoot "Frontend\frontend"
$healthUrl = "http://127.0.0.1:8000/api/health"
$backend = $null

try {
    Write-Host "Starting Python API..." -ForegroundColor Cyan
    $backend = Start-Process python `
        -ArgumentList "-m", "uvicorn", "web_api:app", "--host", "127.0.0.1", "--port", "8000" `
        -WorkingDirectory $projectRoot `
        -WindowStyle Hidden `
        -PassThru

    $backendReady = $false
    for ($attempt = 1; $attempt -le 20; $attempt++) {
        if ($backend.HasExited) {
            throw "Python API failed to start. Run 'python -m uvicorn web_api:app' to view the error."
        }

        try {
            $response = Invoke-RestMethod -Uri $healthUrl -TimeoutSec 1
            if ($response.status -eq "ok") {
                $backendReady = $true
                break
            }
        }
        catch {
            Start-Sleep -Milliseconds 500
        }
    }

    if (-not $backendReady) {
        throw "Python API was not ready within 10 seconds."
    }

    Write-Host "Python API is ready: http://127.0.0.1:8000" -ForegroundColor Green
    Write-Host "Starting frontend. Open the displayed URL in Chrome." -ForegroundColor Cyan
    Set-Location $frontendDir
    npm start
}
finally {
    if ($backend -and -not $backend.HasExited) {
        Write-Host "Stopping Python API..." -ForegroundColor Yellow
        Stop-Process -Id $backend.Id -Force -ErrorAction SilentlyContinue
    }
    Set-Location $projectRoot
}
