$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$logDir = Join-Path $projectRoot "logs"

if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $logDir ("streamlit_" + $timestamp + ".log")

Write-Host "启动 Streamlit，并保存日志到: $logFile"
Write-Host "停止服务请按 Ctrl+C"

python -m streamlit run main.py --server.port 8501 --server.headless true 2>&1 |
    Tee-Object -FilePath $logFile -Append
