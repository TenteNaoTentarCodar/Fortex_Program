@echo off
setlocal
cd /d "%~dp0"
title Compilando ExtratoComprovantes (C#)
echo.
echo ===============================================================
echo   ExtratoComprovantes - versao C# (etapa 2)
echo   Gera o programa em: publicar\ExtratoComprovantes.exe
echo ===============================================================
echo.

where dotnet >nul 2>nul
if errorlevel 1 goto semdotnet
dotnet --list-sdks | findstr /b "8. 9. 10." >nul
if errorlevel 1 goto semdotnet

echo Compilando (a primeira vez baixa os componentes do .NET, precisa de internet)...
echo.
dotnet publish "App\ExtratoComprovantes.csproj" -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true -p:DebugType=none -o publicar
if errorlevel 1 goto erro

echo.
echo ===============================================================
echo   PRONTO: publicar\ExtratoComprovantes.exe
echo   Esse .exe roda sozinho (nao precisa instalar .NET nos outros PCs).
echo ===============================================================
start "" explorer "%~dp0publicar"
pause
exit /b 0

:semdotnet
echo O .NET 8 SDK nao foi encontrado neste computador.
echo Baixe e instale o ".NET 8.0 SDK (x64)" na pagina que vai abrir,
echo depois feche esta janela e rode o build.bat de novo.
start "" "https://dotnet.microsoft.com/download/dotnet/8.0"
pause
exit /b 1

:erro
echo.
echo *** ERRO NA COMPILACAO ***
echo Copie as mensagens acima (ou tire um print) e envie para correcao.
pause
exit /b 1
