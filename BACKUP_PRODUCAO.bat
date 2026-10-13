@echo off
setlocal
cd /d "%~dp0"
title Backup dos arquivos de producao

set "PASTA=\\10.10.2.55\planejamento\0. Time Plan - Mis (Lucas,Mikas,Pam e Italo)\02. Mikas\NUBANK\PAGAMENTOS\MAIO"

powershell -NoProfile -Command "$d = Join-Path (Get-Location) ('backup_producao\' + (Get-Date -Format 'yyyyMMdd_HHmmss')); New-Item -ItemType Directory -Force $d | Out-Null; foreach ($n in 'pagamentos_CC_122.xlsx','pagamentos_LL_122.xlsx') { $o = Join-Path $env:PASTA $n; if (Test-Path -LiteralPath $o) { Copy-Item -LiteralPath $o -Destination $d; Write-Host ('Copiado: ' + $n + ' (' + [math]::Round((Get-Item -LiteralPath (Join-Path $d $n)).Length / 1MB, 1) + ' MB)') } else { Write-Host ('NAO ENCONTRADO: ' + $o) } }; Write-Host ''; Write-Host ('Backup em: ' + $d)"

echo.
pause
