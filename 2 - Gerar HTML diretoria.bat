@echo off
setlocal
rem Guarda o caminho completo da planilha arrastada (se houver) antes de trocar de pasta.
set "PLANILHA=%~f1"
cd /d "%~dp0"

if defined PLANILHA (
    echo Lendo a planilha arrastada: %PLANILHA%
    ".venv\Scripts\python.exe" gerar_html_diretoria.py "%PLANILHA%"
) else (
    echo Lendo a planilha mais recente da pasta "planilhas"...
    ".venv\Scripts\python.exe" gerar_html_diretoria.py
)
if errorlevel 1 goto erro

for /f "delims=" %%F in ('dir /b /o-d "html_diretoria\*.html"') do (
    set "ULTIMO=%%F"
    goto abrir
)

:abrir
echo.
echo HTML da diretoria gerado: html_diretoria\%ULTIMO%
start "" "html_diretoria\%ULTIMO%"
goto fim

:erro
echo.
echo Nao foi possivel gerar o HTML. Veja a mensagem acima.

:fim
pause
