@echo off
setlocal enabledelayedexpansion

REM Verifica che sia stato fornito un percorso
if "%~1"=="" (
    echo Uso: %~nx0 "percorso_cartella" "file_excel" [opzioni]
    echo.
    echo Esempio: %~nx0 "C:\Video" "struttura.xlsx" --save-frames
    exit /b 1
)

REM Verifica che sia stato fornito il file Excel
if "%~2"=="" (
    echo Errore: File Excel non specificato
    echo.
    echo Uso: %~nx0 "percorso_cartella" "file_excel" [opzioni]
    exit /b 1
)

REM Salva il percorso della cartella e del file Excel
set "VIDEO_DIR=%~1"
set "EXCEL_FILE=%~2"

REM Sposta tutti gli argomenti aggiuntivi in una variabile
set "EXTRA_ARGS="
shift
shift
:parse_args
if "%~1"=="" goto :end_args
set "EXTRA_ARGS=!EXTRA_ARGS! %~1"
shift
goto :parse_args
:end_args

REM Verifica che la cartella esista
if not exist "%VIDEO_DIR%" (
    echo Errore: La cartella "%VIDEO_DIR%" non esiste
    exit /b 1
)

REM Verifica che il file Excel esista
if not exist "%EXCEL_FILE%" (
    echo Errore: Il file Excel "%EXCEL_FILE%" non esiste
    exit /b 1
)

echo Elaborazione video nella cartella: %VIDEO_DIR%
echo File Excel: %EXCEL_FILE%
echo Opzioni aggiuntive: %EXTRA_ARGS%
echo.

REM Conta quanti file video ci sono
set "video_count=0"
for %%F in ("%VIDEO_DIR%\*.mp4" "%VIDEO_DIR%\*.mkv" "%VIDEO_DIR%\*.avi" "%VIDEO_DIR%\*.mov") do (
    set /a "video_count+=1"
)

if %video_count% equ 0 (
    echo Nessun file video trovato nella cartella
    exit /b 1
)

echo Trovati %video_count% file video
echo.

REM Processa ogni file video
for %%F in ("%VIDEO_DIR%\*.mp4" "%VIDEO_DIR%\*.mkv" "%VIDEO_DIR%\*.avi" "%VIDEO_DIR%\*.mov") do (
    echo Elaborazione di: %%~nxF
    
    REM Controlla se il nome del file inizia con "Section"
    set "section_param="
    set "filename=%%~nF"
    REM Rimuovi gli spazi dopo "Section" se presenti
    set "filename=!filename:Section =Section!"
    echo !filename! | findstr /i /r "^Section[0-9-]*$" >nul
    if not errorlevel 1 (
        REM Estrai il numero o range dopo "Section"
        for /f "tokens=2 delims=n" %%S in ("!filename!") do (
            set "section_param=--section %%S"
        )
    )
    
    if defined section_param (
        echo Rilevata sezione: !section_param!
        python app.py "%%F" "%EXCEL_FILE%" !section_param! %EXTRA_ARGS%
    ) else (
        python app.py "%%F" "%EXCEL_FILE%" %EXTRA_ARGS%
    )
    
    if errorlevel 1 (
        echo Errore nell'elaborazione di %%~nxF
        echo.
    ) else (
        echo Elaborazione di %%~nxF completata con successo
        echo.
    )
)

echo Elaborazione completata
exit /b 0
