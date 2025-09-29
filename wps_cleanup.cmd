@echo off
setlocal enabledelayedexpansion

:: ========================
:: 1. Удаление WPS
:: ========================
echo Removing WPS from AppData...

:: Run taskkill with admin privileges
powershell -Command "Start-Process taskkill -ArgumentList '/f /im wpscenter.exe' -Verb RunAs -WindowStyle Hidden -Wait"
powershell -Command "Start-Process taskkill -ArgumentList '/f /im wpscloudsvr.exe' -Verb RunAs -WindowStyle Hidden -Wait"

set "user_local=%LOCALAPPDATA%\Kingsoft"
set "user_roaming=%APPDATA%\Kingsoft"

:: Run folder deletion with admin privileges
if exist "%user_local%" powershell -Command "Start-Process cmd -ArgumentList '/c rmdir /s /q \"%user_local%\"' -Verb RunAs -WindowStyle Hidden -Wait"
if exist "%user_roaming%" powershell -Command "Start-Process cmd -ArgumentList '/c rmdir /s /q \"%user_roaming%\"' -Verb RunAs -WindowStyle Hidden -Wait"

:: ========================
:: 2. Очистка реестра WPS
:: ========================
echo Cleaning WPS registry entries...

:: Проверка и создание записей TypeLib в реестре
:: Определяем диск с Windows (используется для системных путей)
set "SYS_DRIVE=%SystemDrive%"

set "TYPED_FOLDER=HKEY_CLASSES_ROOT\TypeLib\{0002E157-0000-0000-C000-000000000046}"
set "VER_FOLDER=%TYPED_FOLDER%\5.3"
set "ZERO_FOLDER=%VER_FOLDER%\0"
set "WIN32_FOLDER=%ZERO_FOLDER%\win32"
set "FLAGS_FOLDER=%VER_FOLDER%\FLAGS"
set "HELPDIR_FOLDER=%VER_FOLDER%\HELPDIR"

echo System drive detected: %SYS_DRIVE%

:: Создаём главный ключ TypeLib если отсутствует
reg query "%TYPED_FOLDER%" >nul 2>&1 || (
    reg add "%TYPED_FOLDER%" /f >nul
)

:: Создаём версию 5.3
reg query "%VER_FOLDER%" >nul 2>&1 || (
    reg add "%VER_FOLDER%" /f >nul
)

:: Устанавливаем имя по умолчанию для 5.3 (Default)
reg add "%VER_FOLDER%" /ve /t REG_SZ /d "Microsoft Visual Basic for Applications Extensibility 5.3" /f >nul

:: Устанавливаем PrimaryInteropAssemblyName
reg add "%VER_FOLDER%" /v "PrimaryInteropAssemblyName" /t REG_SZ /d "Microsoft.Vbe.Interop, Version=12.0.0.0, Culture=neutral, PublicKeyToken=71E9BCE111E9429C" /f >nul

:: Создаём подраздел 0
reg query "%ZERO_FOLDER%" >nul 2>&1 || (
    reg add "%ZERO_FOLDER%" /f >nul
)

:: Создаём win32 и задаём путь к OLB файлу, используя системный диск
set "EXPECTED_OLB=%SYS_DRIVE%\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
reg query "%WIN32_FOLDER%" >nul 2>&1 || (
    reg add "%WIN32_FOLDER%" /f >nul
)

reg add "%WIN32_FOLDER%" /ve /t REG_SZ /d "%EXPECTED_OLB%" /f >nul

:: FLAGS = "0"
reg add "%FLAGS_FOLDER%" /f >nul
reg add "%FLAGS_FOLDER%" /ve /t REG_SZ /d "0" /f >nul

:: HELPDIR = "[{...}]"
reg add "%HELPDIR_FOLDER%" /f >nul
reg add "%HELPDIR_FOLDER%" /ve /t REG_SZ /d "[{0002E157-0000-0000-C000-000000000046}]" /f >nul

echo TypeLib registry entries ensured.

reg delete "HKCU\Software\Kingsoft" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "WPS Office" /f >nul 2>&1

set wps_classes=ET. KET. KWPP. KWPS. WPP. WPS. WPSCloudSv WPSFileSync
for %%c in (%wps_classes%) do (
    reg delete "HKCU\Software\Classes\%%c" /f >nul 2>&1
    reg delete "HKCU\Software\Classes\%%c*" /f >nul 2>&1
)

:: Префиксы для удаления (должны быть в начале имени раздела)
set "prefixes=ET. KET. KWPP. KWPS. WPP. WPS."

echo Searching and deleting registry keys in HKEY_CLASSES_ROOT...
echo.

:: Получаем список всех подразделов в HKCR
for /f "tokens=*" %%a in ('reg query "HKEY_CLASSES_ROOT"') do (
    set "key=%%a"
    
    :: Убираем "HKEY_CLASSES_ROOT\" из пути
    set "key=!key:HKEY_CLASSES_ROOT\=!"
    
    :: Проверяем, начинается ли ключ с одного из префиксов
    for %%p in (%prefixes%) do (
        if "!key!"=="%%p!key:*%%p=!" (
            echo Deleting: !key!
            reg delete "HKEY_CLASSES_ROOT\!key!" /f >nul 2>&1
        )
    )
)

echo Searching and deleting registry keys in HKCU\Software\Classes...
echo.

:: Получаем список всех подразделов в HKCU\Software\Classes
for /f "tokens=*" %%a in ('reg query "HKCU\Software\Classes"') do (
    set "key=%%a"
    
    :: Убираем "HKEY_CURRENT_USER\Software\Classes\" из пути
    set "key=!key:HKEY_CURRENT_USER\Software\Classes\=!"
    
    :: Проверяем, начинается ли ключ с одного из префиксов
    for %%p in (%prefixes%) do (
        if "!key!"=="%%p!key:*%%p=!" (
            echo Deleting: !key!
            reg delete "HKCU\Software\Classes\!key!" /f >nul 2>&1
        )
    )
)

:: Удаление ассоциаций файлов
set extensions=.doc .docx .xls .xlsx .ppt .pptx .rtf .txt .csv .dot .dotx .xlt .xltx .pot .potx
for %%e in (%extensions%) do (
    reg delete "HKCU\Software\Classes\%%e" /f >nul 2>&1
    reg delete "HKCU\Software\Classes\%%e\OpenWithProgids" /f >nul 2>&1
)

:: ========================
:: 3. Восстановление Office
:: ========================
echo Restoring Office file associations...

set "office_path=C:\Program Files\Microsoft Office\root\Office16"
set "msi_icon_path=C:\Windows\Installer\{90160000-0011-0000-1000-0000000FF1CE}"
set "pdf_reader_path=C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"

:: Проверка существования папки с иконками
if not exist "%msi_icon_path%" (
    echo Warning: Office icons folder not found!
    echo Using standard EXE file icons.
    set "msi_icon_path=%office_path%"
)

:: Word
reg add "HKCU\Software\Classes\.doc" /ve /t REG_SZ /d "Word.Document.8" /f >nul
reg add "HKCU\Software\Classes\Word.Document.8\shell\open\command" /ve /t REG_SZ /d "\"%office_path%\WINWORD.EXE\" \"%%1\"" /f >nul
reg add "HKCU\Software\Classes\Word.Document.8\DefaultIcon" /ve /t REG_SZ /d "%msi_icon_path%\wordicon.exe,1" /f >nul

reg add "HKCU\Software\Classes\.docx" /ve /t REG_SZ /d "Word.Document.12" /f >nul
reg add "HKCU\Software\Classes\Word.Document.12\shell\open\command" /ve /t REG_SZ /d "\"%office_path%\WINWORD.EXE\" \"%%1\"" /f >nul
reg add "HKCU\Software\Classes\Word.Document.12\DefaultIcon" /ve /t REG_SZ /d "%msi_icon_path%\wordicon.exe,1" /f >nul

:: Excel
reg add "HKCU\Software\Classes\.xls" /ve /t REG_SZ /d "Excel.Sheet.8" /f >nul
reg add "HKCU\Software\Classes\Excel.Sheet.8\shell\open\command" /ve /t REG_SZ /d "\"%office_path%\EXCEL.EXE\" \"%%1\"" /f >nul
reg add "HKCU\Software\Classes\Excel.Sheet.8\DefaultIcon" /ve /t REG_SZ /d "%msi_icon_path%\xlicons.exe,1" /f >nul

reg add "HKCU\Software\Classes\.xlsx" /ve /t REG_SZ /d "Excel.Sheet.12" /f >nul
reg add "HKCU\Software\Classes\Excel.Sheet.12\shell\open\command" /ve /t REG_SZ /d "\"%office_path%\EXCEL.EXE\" \"%%1\"" /f >nul
reg add "HKCU\Software\Classes\Excel.Sheet.12\DefaultIcon" /ve /t REG_SZ /d "%msi_icon_path%\xlicons.exe,1" /f >nul

:: PowerPoint
reg add "HKCU\Software\Classes\.ppt" /ve /t REG_SZ /d "PowerPoint.Show.8" /f >nul
reg add "HKCU\Software\Classes\PowerPoint.Show.8\shell\open\command" /ve /t REG_SZ /d "\"%office_path%\POWERPNT.EXE\" \"%%1\"" /f >nul
reg add "HKCU\Software\Classes\PowerPoint.Show.8\DefaultIcon" /ve /t REG_SZ /d "%msi_icon_path%\ppticons.exe,1" /f >nul

reg add "HKCU\Software\Classes\.pptx" /ve /t REG_SZ /d "PowerPoint.Show.12" /f >nul
reg add "HKCU\Software\Classes\PowerPoint.Show.12\shell\open\command" /ve /t REG_SZ /d "\"%office_path%\POWERPNT.EXE\" \"%%1\"" /f >nul
reg add "HKCU\Software\Classes\PowerPoint.Show.12\DefaultIcon" /ve /t REG_SZ /d "%msi_icon_path%\ppticons.exe,1" /f >nul

:: PDF
reg add "HKCU\Software\Classes\.pdf" /ve /t REG_SZ /d "AcroExch.Document.DC" /f >nul
reg add "HKCU\Software\Classes\AcroExch.Document.DC\shell\open\command" /ve /t REG_SZ /d "\"%pdf_reader_path%\" \"%%1\"" /f >nul
reg add "HKCU\Software\Classes\AcroExch.Document.DC\DefaultIcon" /ve /t REG_SZ /d "%pdf_reader_path%,0" /f >nul

:: ========================
:: 4. Восстановление "Создать -> Документ Word"
:: ========================
echo Restoring "New -> Word Document" option...

:: Шаблон для нового документа Word
set "word_template=%USERPROFILE%\Documents\Custom Office Templates\Blank.docx"
if not exist "%USERPROFILE%\Documents\Custom Office Templates\" mkdir "%USERPROFILE%\Documents\Custom Office Templates\"

:: Если шаблона нет - создаём пустой
if not exist "%word_template%" (
    echo. > "%word_template%"
)

:: Добавляем в контекстное меню
reg add "HKCU\Software\Classes\.docx\ShellNew" /v "FileName" /t REG_SZ /d "%word_template%" /f >nul
reg add "HKCU\Software\Classes\.doc\ShellNew" /v "FileName" /t REG_SZ /d "%word_template%" /f >nul

:: Альтернативный способ (если первый не сработает)
reg add "HKCU\Software\Classes\.docx\ShellNew" /v "NullFile" /t REG_SZ /d "" /f >nul
reg add "HKCU\Software\Classes\.doc\ShellNew" /v "NullFile" /t REG_SZ /d "" /f >nul

:: ========================
:: 5. Обновление иконок
:: ========================
echo Updating icon cache...

:: Очистка кэша иконок
del /q "%userprofile%\AppData\Local\IconCache.db" >nul 2>&1
del /q "%userprofile%\AppData\Local\Microsoft\Windows\Explorer\iconcache*" >nul 2>&1

:: Перезапуск проводника с правами администратора
powershell -Command "Start-Process taskkill -ArgumentList '/f /im explorer.exe' -Verb RunAs -WindowStyle Hidden -Wait"
start explorer.exe >nul 2>&1
timeout /t 3 >nul

:: ========================
:: Готово!
:: ========================
echo.
echo [+] WPS Office has been removed.
echo [+] File associations have been restored.
echo [+] Icons have been updated.
echo.
pause
