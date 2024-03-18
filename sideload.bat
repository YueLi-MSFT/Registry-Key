@echo off

REM Set the registry key path and value
set "keyPath=HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentEcs\Overrides"
set "key1=Microsoft.Office.OEP.MosProviderEnabled"
set "value1=%2"
set "key2=Microsoft.Office.OEP.MosManifest"
set "value2=%3"

if "%1"=="add" (
    REM Add the registry key and value
    reg add "%keyPath%" /v "%key1%" /t REG_SZ /d "%value1%" /f
    reg add "%keyPath%" /v "%key2%" /t REG_SZ /d "%value2%" /f

    REM Check if the registry key was added successfully
    if %errorlevel% equ 0 (
        echo Registry key added successfully.
    ) else (
        echo Failed to add registry key.
    )
)

if "%1"=="delete" (
    REM Delete the registry key and value
    reg delete "%keyPath%" /v "%key1%" /f
    reg delete "%keyPath%" /v "%key2%" /f

    REM Check if the registry key was deleted successfully
    if %errorlevel% equ 0 (
        echo Registry key deleted successfully.
    ) else (
        echo Failed to delete registry key.
    )
)

REM Pause the script to keep the command prompt window open
pause