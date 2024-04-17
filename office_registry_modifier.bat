@echo off

REM Set the registry key path and value
set "keyPath=HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Common\ExperimentEcs\Overrides"
set "key1=Microsoft.Office.OEP.MosProviderEnabled"
set "value1=true"
set "key2=Microsoft.Office.OEP.MosManifest"
set "value2=true"
set "key3=Microsoft.Office.OEP.CG_MosPopulateContentEnabled"
set "value3=true"
set "key4=Microsoft.Office.OEP.EnableOsfMosAppFlyout"
set "value4=true"
set "key5=Microsoft.Office.OEP.MosPopulateContentEnabled"
set "value5=true"
set "key6=Microsoft.Office.OEP.ChangeGate.DedupeXmlAddInForMosExtension"
set "value6=false"

if "%1"=="add" (
    REM Add the registry key and value
    reg add "%keyPath%" /v "%key1%" /t REG_SZ /d "%value1%" /f
    reg add "%keyPath%" /v "%key2%" /t REG_SZ /d "%value2%" /f
    reg add "%keyPath%" /v "%key3%" /t REG_SZ /d "%value3%" /f
    reg add "%keyPath%" /v "%key4%" /t REG_SZ /d "%value4%" /f
    reg add "%keyPath%" /v "%key5%" /t REG_SZ /d "%value5%" /f
    reg add "%keyPath%" /v "%key6%" /t REG_SZ /d "%value6%" /f

    REM Check if the registry key was added successfully
    if %errorlevel% equ 0 (
        echo Registry keys added successfully.
    ) else (
        echo Failed to add registry keys.
    )
)

if "%1"=="delete" (
    REM Delete the registry key and value
    reg delete "%keyPath%" /v "%key1%" /f
    reg delete "%keyPath%" /v "%key2%" /f
    reg delete "%keyPath%" /v "%key3%" /f
    reg delete "%keyPath%" /v "%key4%" /f
    reg delete "%keyPath%" /v "%key5%" /f
    reg delete "%keyPath%" /v "%key6%" /f

    REM Check if the registry key was deleted successfully
    if %errorlevel% equ 0 (
        echo Registry keys deleted successfully.
    ) else (
        echo Failed to delete registry keys.
    )
)

REM Pause the script to keep the command prompt window open
pause