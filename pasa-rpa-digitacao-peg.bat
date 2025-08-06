@echo off
setlocal

REM Caminho do seu executável
set "APP=C:\projects\pasa-rpa-digitacao-peg\pasa-rpa-digitacao-peg.exe"

REM Executa o .exe
%APP%
set "RESULT=%ERRORLEVEL%"

REM Testa o código de saída
if %RESULT% NEQ 0 (
    echo Falhou com código %RESULT%, agendando nova tentativa...
    timeout /t 60 >nul
    schtasks /run /tn "pasa-rpa-digitacao-peg"
) else (
    echo Sucesso.
)

endlocal
exit /b %RESULT%
