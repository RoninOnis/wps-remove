@echo off
setlocal EnableDelayedExpansion

echo.
echo ===================================================
echo   Removing WPS traces and restoring Office
echo ===================================================
echo.

:: ========================
:: 1. Search for Office and icons paths
:: ========================
echo [Searching for Office and icons folder...]

set "OFFICE_ROOT="
set "ICONS_FOLDER="

for /f "delims=" %%A in ('where /r c:\ WINWORD.EXE 2^>nul') do (
    set "candidate=%%~dpA"
    set "candidate=!candidate:~0,-1!"
    echo "!candidate!" | findstr /i /c:"ClickToRun" /c:"Downloaded" /c:"Packages" >nul || (
        set "OFFICE_ROOT=!candidate!"
        goto :office_found
    )
)
:office_found

if not defined OFFICE_ROOT (
    echo Office not found on disk.
    echo Trying standard paths...
    for %%p in (
        "C:\Program Files\Microsoft Office\root\Office16"
        "C:\Program Files (x86)\Microsoft Office\root\Office16"
        "C:\Program Files\Microsoft Office\root\Office15"
        "C:\Program Files (x86)\Microsoft Office\root\Office15"
    ) do (
        if exist "%%~p\WINWORD.EXE" (
            set "OFFICE_ROOT=%%~p"
            goto :office_done
        )
    )
    :office_done
    if not defined OFFICE_ROOT (
        echo Warning: Office not found anywhere. File associations may not work.
        set "OFFICE_ROOT=C:\Program Files\Microsoft Office\root\Office16"
    )
)

for /f "delims=" %%F in ('where /r c:\ wordicon.* 2^>nul') do (
    set "dir=%%~dpF"
    set "dir=!dir:~0,-1!"
    echo "!dir!" | findstr /r /c:"{[0-9A-F-]*}" >nul && (
        set "ICONS_FOLDER=!dir!"
        goto :icons_found
    )
)
:icons_found

if not defined ICONS_FOLDER (
    echo GUID folder with icons not found → using Office folder as fallback.
    set "ICONS_FOLDER=%OFFICE_ROOT%"
)

echo.
echo Detected paths:
echo   Office      → %OFFICE_ROOT%
echo   Icons       → %ICONS_FOLDER%
echo.

:: ========================
:: 2. Remove WPS traces
:: ========================
echo Removing WPS processes and folders...

powershell -Command "Start-Process taskkill -ArgumentList '/f /im wpscenter.exe' -Verb RunAs -WindowStyle Hidden -Wait" 2>nul
powershell -Command "Start-Process taskkill -ArgumentList '/f /im wpscloudsvr.exe' -Verb RunAs -WindowStyle Hidden -Wait" 2>nul

set "local_kingsoft=%LOCALAPPDATA%\Kingsoft"
set "roaming_kingsoft=%APPDATA%\Kingsoft"

if exist "%local_kingsoft%"   powershell -Command "Start-Process cmd -ArgumentList '/c rmdir /s /q \"%local_kingsoft%\"' -Verb RunAs -WindowStyle Hidden -Wait" 2>nul
if exist "%roaming_kingsoft%" powershell -Command "Start-Process cmd -ArgumentList '/c rmdir /s /q \"%roaming_kingsoft%\"' -Verb RunAs -WindowStyle Hidden -Wait" 2>nul

:: ========================
:: 3. Clean registry + restore VBA TypeLib
:: ========================
echo Cleaning registry and restoring VBA Extensibility...

reg delete "HKCU\Software\Kingsoft" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "WPS Office" /f >nul 2>&1

set "wps_prefixes=ET. KET. KWPP. KWPS. WPP. WPS. WPSCloudSv WPSFileSync"
for %%p in (%wps_prefixes%) do (
    reg delete "HKCU\Software\Classes\%%p" /f >nul 2>&1
    reg delete "HKEY_CLASSES_ROOT\%%p" /f >nul 2>&1
)

:: ─── Restore VBA Extensibility 5.3 TypeLib ───
set "SYSTEM_DRIVE=%SystemDrive%"

set "TYPELIB_KEY=HKEY_CLASSES_ROOT\TypeLib\{0002E157-0000-0000-C000-000000000046}"
set "VERSION_KEY=%TYPELIB_KEY%\5.3"
set "ZERO_KEY=%VERSION_KEY%\0"
set "WIN32_KEY=%ZERO_KEY%\win32"
set "FLAGS_KEY=%VERSION_KEY%\FLAGS"
set "HELPDIR_KEY=%VERSION_KEY%\HELPDIR"

echo Restoring VBA TypeLib entries...

reg query "%TYPELIB_KEY%" >nul 2>&1 || reg add "%TYPELIB_KEY%" /f >nul
reg query "%VERSION_KEY%"   >nul 2>&1 || reg add "%VERSION_KEY%"   /f >nul

reg add "%VERSION_KEY%" /ve /t REG_SZ /d "Microsoft Visual Basic for Applications Extensibility 5.3" /f >nul
reg add "%VERSION_KEY%" /v "PrimaryInteropAssemblyName" /t REG_SZ /d "Microsoft.Vbe.Interop, Version=12.0.0.0, Culture=neutral, PublicKeyToken=71E9BCE111E9429C" /f >nul

reg query "%ZERO_KEY%" >nul 2>&1 || reg add "%ZERO_KEY%" /f >nul
reg query "%WIN32_KEY%" >nul 2>&1 || reg add "%WIN32_KEY%" /f >nul

set "EXPECTED_OLB=%SYSTEM_DRIVE%\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
reg add "%WIN32_KEY%" /ve /t REG_SZ /d "%EXPECTED_OLB%" /f >nul

reg add "%FLAGS_KEY%"   /ve /t REG_SZ /d "0" /f >nul
reg add "%HELPDIR_KEY%" /ve /t REG_SZ /d "[{0002E157-0000-0000-C000-000000000046}]" /f >nul

echo VBA TypeLib entries restored.

:: ========================
:: 4. Restore file associations
:: ========================
echo Restoring file associations...

:: Word
reg add "HKCU\Software\Classes\.doc"  /ve /t REG_SZ /d "Word.Document.8" /f >nul
reg add "HKCU\Software\Classes\.docx" /ve /t REG_SZ /d "Word.Document.12" /f >nul
reg add "HKCU\Software\Classes\Word.Document.8\shell\open\command"  /ve /t REG_SZ /d "\"%OFFICE_ROOT%\WINWORD.EXE\" \"%%1\"" /f >nul
reg add "HKCU\Software\Classes\Word.Document.12\shell\open\command" /ve /t REG_SZ /d "\"%OFFICE_ROOT%\WINWORD.EXE\" \"%%1\"" /f >nul
reg add "HKCU\Software\Classes\Word.Document.8\DefaultIcon"  /ve /t REG_SZ /d "%ICONS_FOLDER%\wordicon.exe,1" /f >nul
reg add "HKCU\Software\Classes\Word.Document.12\DefaultIcon" /ve /t REG_SZ /d "%ICONS_FOLDER%\wordicon.exe,1" /f >nul

:: Excel
reg add "HKCU\Software\Classes\.xls"  /ve /t REG_SZ /d "Excel.Sheet.8" /f >nul
reg add "HKCU\Software\Classes\.xlsx" /ve /t REG_SZ /d "Excel.Sheet.12" /f >nul
reg add "HKCU\Software\Classes\Excel.Sheet.8\shell\open\command"  /ve /t REG_SZ /d "\"%OFFICE_ROOT%\EXCEL.EXE\" \"%%1\"" /f >nul
reg add "HKCU\Software\Classes\Excel.Sheet.12\shell\open\command" /ve /t REG_SZ /d "\"%OFFICE_ROOT%\EXCEL.EXE\" \"%%1\"" /f >nul
reg add "HKCU\Software\Classes\Excel.Sheet.8\DefaultIcon"  /ve /t REG_SZ /d "%ICONS_FOLDER%\xlicons.exe,1" /f >nul
reg add "HKCU\Software\Classes\Excel.Sheet.12\DefaultIcon" /ve /t REG_SZ /d "%ICONS_FOLDER%\xlicons.exe,1" /f >nul

:: PowerPoint
reg add "HKCU\Software\Classes\.ppt"  /ve /t REG_SZ /d "PowerPoint.Show.8" /f >nul
reg add "HKCU\Software\Classes\.pptx" /ve /t REG_SZ /d "PowerPoint.Show.12" /f >nul
reg add "HKCU\Software\Classes\PowerPoint.Show.8\shell\open\command"  /ve /t REG_SZ /d "\"%OFFICE_ROOT%\POWERPNT.EXE\" \"%%1\"" /f >nul
reg add "HKCU\Software\Classes\PowerPoint.Show.12\shell\open\command" /ve /t REG_SZ /d "\"%OFFICE_ROOT%\POWERPNT.EXE\" \"%%1\"" /f >nul
reg add "HKCU\Software\Classes\PowerPoint.Show.8\DefaultIcon"  /ve /t REG_SZ /d "%ICONS_FOLDER%\pptico.exe,1" /f >nul
reg add "HKCU\Software\Classes\PowerPoint.Show.12\DefaultIcon" /ve /t REG_SZ /d "%ICONS_FOLDER%\pptico.exe,1" /f >nul

:: ========================
:: 5. Restore "New → Word Document" context menu
:: ========================
echo Restoring "New → Word Document" option...

set "TEMPLATES_FOLDER=%USERPROFILE%\Documents\Custom Office Templates"
if not exist "%TEMPLATES_FOLDER%" mkdir "%TEMPLATES_FOLDER%"

set "BLANK_DOC=%TEMPLATES_FOLDER%\Blank.docx"
if not exist "%BLANK_DOC%" echo. > "%BLANK_DOC%"

reg add "HKCU\Software\Classes\.docx\ShellNew" /v "FileName" /t REG_SZ /d "%BLANK_DOC%" /f >nul
reg add "HKCU\Software\Classes\.doc\ShellNew"  /v "FileName" /t REG_SZ /d "%BLANK_DOC%" /f >nul
reg add "HKCU\Software\Classes\.docx\ShellNew" /v "NullFile" /t REG_SZ /d "" /f >nul 2>nul
reg add "HKCU\Software\Classes\.doc\ShellNew"  /v "NullFile" /t REG_SZ /d "" /f >nul 2>nul

:: ========================
:: 6. Update icon cache
:: ========================
echo Updating icon cache...

del /q /f "%LocalAppData%\IconCache.db" >nul 2>&1
del /q /f "%LocalAppData%\Microsoft\Windows\Explorer\iconcache_*" >nul 2>&1

powershell -Command "Start-Process taskkill -ArgumentList '/f /im explorer.exe' -Verb RunAs -WindowStyle Hidden -Wait" 2>nul
start explorer.exe >nul 2>&1
timeout /t 4 >nul

:: ========================
:: Done
:: ========================
echo.
echo ===================================================
echo   Completed!
echo   WPS traces removed, Office associations restored.
echo   Office path:    %OFFICE_ROOT%
echo   Icons from:     %ICONS_FOLDER%
echo ===================================================
echo.
pause
