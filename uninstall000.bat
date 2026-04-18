@echo off
setlocal
title Uninstall Status... Do not Close Window...
cd /d "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; $choice=[System.Windows.MessageBox]::Show('Are you sure you want to uninstall Barnaby and all components?','Uninstall Barnaby','YesNo','Warning'); if ($choice -eq 'Yes') { exit 0 } else { exit 1 }"
if not errorlevel 1 goto uninstall_confirmed
echo Barnaby uninstall cancelled.
timeout /t 2 >nul
exit /b 1

:uninstall_confirmed
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; $choice=[System.Windows.MessageBox]::Show('Would you also like to delete data and preferences?','Barnaby saved data','YesNo','Question'); if ($choice -eq 'Yes') { exit 0 } else { exit 2 }"
set DATA_CHOICE=%ERRORLEVEL%
set DELETE_DATA=no
if "%DATA_CHOICE%"=="0" set DELETE_DATA=yes

echo Uninstall Status... Do not Close Window...
echo Closing Barnaby if he is running...
taskkill /f /im python.exe /fi "WINDOWTITLE eq Barnaby the Octopus" >nul 2>nul
taskkill /f /im pythonw.exe /fi "WINDOWTITLE eq Barnaby the Octopus" >nul 2>nul

echo Removing Barnaby helper environment...
if not exist ".venv" goto shortcut_remove
rmdir /s /q ".venv"

:shortcut_remove
echo Removing desktop shortcut...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$shortcut=Join-Path ([Environment]::GetFolderPath('Desktop')) 'Barnaby the Octopus.lnk'; if (Test-Path $shortcut) { Remove-Item $shortcut -Force }"

if /i "%DELETE_DATA%"=="yes" goto delete_saved_data
echo Keeping Barnaby data and preferences.
goto remove_program_files

:delete_saved_data
echo Removing Barnaby data and preferences...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$data=Join-Path $env:APPDATA 'BarnabyDesktopHelper'; if (Test-Path $data) { Remove-Item $data -Recurse -Force }"

:remove_program_files
echo Removing Barnaby program files...
if exist "__pycache__" rmdir /s /q "__pycache__"
if exist "Run.bat" del /f /q "Run.bat"
if exist "barnaby.py" del /f /q "barnaby.py"
if exist "requirements.txt" del /f /q "requirements.txt"
if exist "install.bat" del /f /q "install.bat"

:uninstall_done
echo Barnaby has been uninstalled from this folder.
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Barnaby has been uninstalled. His program files, helper environment, and shortcut were removed.','Barnaby uninstalled','OK','Information')"
echo This uninstall file will remove itself after you press a key.
pause
del /f /q "%~f0"
