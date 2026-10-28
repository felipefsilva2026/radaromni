@echo off
setlocal
rem Uso: duplo clique (modo normal) ou "1 - Gerar planilha.bat" auto (agendador: sem pausa, grava log).
set "MODO=%~1"
cd /d "%~dp0"
if not exist logs mkdir logs

echo Pesquisando noticias (ultimos 30 dias) e montando a planilha de curadoria...
if /i "%MODO%"=="auto" (
    echo [%date% %time%] Inicio da geracao automatica >> "logs\agendamento.log"
    ".venv\Scripts\python.exe" gerar_planilha.py --days 30 >> "logs\agendamento.log" 2>&1
) else (
    ".venv\Scripts\python.exe" gerar_planilha.py --days 30
)
if errorlevel 1 goto erro

for /f "delims=" %%F in ('dir /b /o-d "planilhas\*.xlsx"') do (
    set "ULTIMA=%%F"
    goto abrir
)

:abrir
echo.
echo Planilha pronta: planilhas\%ULTIMA%
echo Revise, ajuste as colunas laranjas, marque SIM/NAO na coluna STATUS e salve.
echo Depois rode "2 - Gerar HTML diretoria.bat".
start "" "planilhas\%ULTIMA%"
goto fim

:erro
echo.
echo Ocorreu um erro ao gerar a planilha. Veja as mensagens acima.
if /i "%MODO%"=="auto" echo [%date% %time%] ERRO na geracao automatica >> "logs\agendamento.log"

:fim
if /i not "%MODO%"=="auto" pause
