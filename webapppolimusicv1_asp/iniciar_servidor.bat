@echo off
title PoliMusic - Servidor IIS Express
echo =====================================================================
echo          PoliMusic - Clon de Spotify en ASP Clasico (.asp)
echo =====================================================================
echo.
echo Servidor : IIS Express
echo Puerto   : 8085
echo Directorio: %~dp0
echo.
echo Iniciando servidor y abriendo navegador en http://localhost:8085/login.asp ...
echo Para detener el servidor, cierre esta ventana o presione Ctrl + C.
echo.
set "APP_PATH=%~dp0"
if "%APP_PATH:~-1%"=="\" set "APP_PATH=%APP_PATH:~0,-1%"
start "" "http://localhost:8085/login.asp"
"C:\Program Files\IIS Express\iisexpress.exe" /path:"%APP_PATH%" /port:8085
pause
