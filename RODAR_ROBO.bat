@echo off
setlocal
cd /d "%~dp0"
title Robo QuickSight - PRODUCAO

if not exist ".venv\Scripts\activate.bat" (
  echo Nao encontrei a pasta .venv aqui. Deixe este arquivo dentro da pasta RoboQuickSight.
  pause
  exit /b 1
)
if not exist "config_producao.yaml" (
  echo Nao encontrei o config_producao.yaml nesta pasta.
  pause
  exit /b 1
)

call ".venv\Scripts\activate.bat"
echo ==========================================================
echo  ROBO QUICKSIGHT - PRODUCAO
echo  Atualiza os arquivos OFICIAIS de pagamentos (CC e LL).
echo  Se pedir login: entre na janela que abrir e depois volte
echo  aqui e aperte Enter.
echo ==========================================================
echo.

python qs_download.py --config config_producao.yaml
set CODIGO=%ERRORLEVEL%

echo.
echo ==========================================================
if "%CODIGO%"=="0" (
  echo  RESULTADO: OK
) else (
  echo  RESULTADO: ATENCAO - codigo %CODIGO%. Leia as linhas ERROR e WARNING acima.
)
echo ==========================================================
echo.
pause
exit /b %CODIGO%
