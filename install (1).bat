@echo off
setlocal
title Barnaby Installer
cd /d "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; $choice=[System.Windows.MessageBox]::Show('Install Barnaby?','Barnaby the Octopus','YesNo','Question'); if ($choice -eq 'Yes') { exit 0 } else { exit 1 }"
if errorlevel 1 (
  echo Barnaby install cancelled.
  timeout /t 2 >nul
  exit /b 1
)

echo Installing Barnaby... Do not close this window.
echo Barnaby is creating his home, installing helper parts, and preparing Run.bat.

where py >nul 2>nul
if errorlevel 1 (
  echo Python launcher was not found.
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; $choice=[System.Windows.MessageBox]::Show('Barnaby needs Python 3. Would you like Windows to try installing Python automatically with winget?','Barnaby needs Python','YesNo','Question'); if ($choice -eq 'Yes') { exit 0 } else { exit 1 }"
  if errorlevel 1 (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Please install Python 3 from python.org, then run install.bat again.','Barnaby install paused','OK','Warning')"
    exit /b 1
  )
  where winget >nul 2>nul
  if errorlevel 1 (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Windows package installer winget was not found. Please install Python 3 from python.org, then run install.bat again.','Barnaby needs Python','OK','Warning')"
    exit /b 1
  )
  echo Installing Python 3...
  winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
  if errorlevel 1 (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Python could not be installed automatically. Please install Python 3 from python.org, then run install.bat again.','Barnaby needs Python','OK','Warning')"
    exit /b 1
  )
)

where py >nul 2>nul
if errorlevel 1 (
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Python was installed, but Windows may need a new command window. Close this window and run install.bat again.','Restart installer','OK','Information')"
  exit /b 1
)

if not exist "requirements.txt" (
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('requirements.txt is missing. Barnaby cannot install packages.','Barnaby install error','OK','Error')"
  exit /b 1
)

if not exist "barnaby.py" (
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('barnaby.py is missing. Barnaby cannot install.','Barnaby install error','OK','Error')"
  exit /b 1
)

py -3 -m venv ".venv"
if errorlevel 1 (
  echo Barnaby could not create his Python environment.
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Barnaby could not create his Python environment. Try running install.bat again.','Barnaby install error','OK','Error')"
  exit /b 1
)

".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 (
  echo Barnaby could not update pip.
  exit /b 1
)

".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo Barnaby could not finish installing packages from requirements.txt.
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Barnaby could not install requirements.txt. Check your internet connection and run install.bat again.','Barnaby install error','OK','Error')"
  exit /b 1
)

(
  echo @echo off
  echo title Barnaby the Octopus
  echo cd /d "%%~dp0"
  echo if not exist ".venv\Scripts\python.exe" ^(
  echo   powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Barnaby is not installed yet. Please run install.bat first.','Barnaby not installed','OK','Warning')"
  echo   exit /b 1
  echo ^)
  echo if not exist "barnaby.py" ^(
  echo   powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('barnaby.py is missing. Please reinstall Barnaby.','Barnaby file missing','OK','Error')"
  echo   exit /b 1
  echo ^)
  echo ".venv\Scripts\python.exe" "barnaby.py"
) > "Run.bat"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$desktop=[Environment]::GetFolderPath('Desktop'); $shortcut=Join-Path $desktop 'Barnaby the Octopus.lnk'; $shell=New-Object -ComObject WScript.Shell; $link=$shell.CreateShortcut($shortcut); $link.TargetPath=(Join-Path (Get-Location) 'Run.bat'); $link.WorkingDirectory=(Get-Location).Path; $link.Description='Run Barnaby the Octopus desktop helper'; $link.Save()"

echo Barnaby installed.
echo Run.bat is ready.
echo uninstall000.bat is ready.
echo A desktop shortcut was created when Windows allowed it.
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Barnaby installed! Open Run.bat or the desktop shortcut to start him. uninstall000.bat can remove him later.','Barnaby is ready','OK','Information')"
pause
