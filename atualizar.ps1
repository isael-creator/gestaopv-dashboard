# atualizar.ps1 — Gera o dashboard SolarZ com dados atualizados do Google Sheets
# Uso: clique com botão direito → "Executar com PowerShell"

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  SolarZ Dashboard — Atualizacao" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ID da planilha mestre do Google Sheets
$env:MASTER_SHEET_ID = "1i9IS-3SEtxYna5AM7c5RQEA7UORtEaneAvClkKs9Yus"
$env:OUTPUT_DIR      = $ScriptDir

Write-Host "Conectando ao Google Sheets e baixando planilhas..." -ForegroundColor Yellow
Write-Host "(Isso pode levar 2-3 minutos para 107 empresas)" -ForegroundColor Gray
Write-Host ""

$start = Get-Date

try {
    & python "$ScriptDir\gerar_dashboard.py"
    if ($LASTEXITCODE -ne 0) { throw "Python retornou codigo $LASTEXITCODE" }
} catch {
    Write-Host ""
    Write-Host "ERRO ao gerar o dashboard:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "Pressione qualquer tecla para fechar..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

$elapsed = [math]::Round(((Get-Date) - $start).TotalSeconds)
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Dashboard atualizado em ${elapsed}s!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Abrindo index.html no navegador..." -ForegroundColor Cyan

Start-Process "$ScriptDir\index.html"

Start-Sleep -Seconds 2
