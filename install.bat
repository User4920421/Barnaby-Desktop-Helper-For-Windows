@echo off
setlocal
title Barnaby Installer
cd /d "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; $choice=[System.Windows.MessageBox]::Show('Install Barnaby?','Barnaby the Octopus','YesNo','Question'); if ($choice -eq 'Yes') { exit 0 } else { exit 1 }"
if not errorlevel 1 goto install_confirmed
echo Barnaby install cancelled.
timeout /t 2 >nul
exit /b 1

:install_confirmed
echo Installing Barnaby... Do not close this window.
echo Barnaby is creating his home, installing helper parts, and preparing Run.bat.

where py >nul 2>nul
if not errorlevel 1 goto python_ready
echo Python launcher was not found.
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; $choice=[System.Windows.MessageBox]::Show('Barnaby needs Python 3. Would you like Windows to try installing Python automatically with winget?','Barnaby needs Python','YesNo','Question'); if ($choice -eq 'Yes') { exit 0 } else { exit 1 }"
if not errorlevel 1 goto try_winget
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Please install Python 3 from python.org, then run install.bat again.','Barnaby install paused','OK','Warning')"
exit /b 1

:try_winget
where winget >nul 2>nul
if not errorlevel 1 goto winget_ready
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Windows package installer winget was not found. Please install Python 3 from python.org, then run install.bat again.','Barnaby needs Python','OK','Warning')"
exit /b 1

:winget_ready
echo Installing Python 3...
winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
if not errorlevel 1 goto python_recheck
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Python could not be installed automatically. Please install Python 3 from python.org, then run install.bat again.','Barnaby needs Python','OK','Warning')"
exit /b 1

:python_recheck
where py >nul 2>nul
if not errorlevel 1 goto python_ready
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Python was installed, but Windows may need a new command window. Close this window and run install.bat again.','Restart installer','OK','Information')"
exit /b 1

:python_ready
> "requirements.txt" echo psutil^>=5.9.8
>> "requirements.txt" echo pyttsx3^>=2.90
>> "requirements.txt" echo pywin32^>=306; sys_platform == "win32"

if exist "barnaby.py" goto app_file_ok
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('barnaby.py is missing. Barnaby cannot install.','Barnaby install error','OK','Error')"
exit /b 1

:app_file_ok
py -3 -m venv ".venv"
if not errorlevel 1 goto venv_ok
echo Barnaby could not create his Python environment.
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Barnaby could not create his Python environment. Try running install.bat again.','Barnaby install error','OK','Error')"
exit /b 1

:venv_ok
".venv\Scripts\python.exe" -m pip install --upgrade pip setuptools wheel
if not errorlevel 1 goto pip_ok
echo Barnaby could not update pip.
exit /b 1

:pip_ok
".venv\Scripts\python.exe" -m pip install --upgrade --prefer-binary -r requirements.txt
if not errorlevel 1 goto packages_ok
echo Barnaby could not finish installing packages from requirements.txt.
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Barnaby could not install requirements.txt. Check your internet connection, make sure Python is version 3.10 or newer, then run install.bat again.','Barnaby install error','OK','Error')"
exit /b 1

:packages_ok
if exist "Run.bat" goto shortcut_create
> "Run.bat" echo @echo off
>> "Run.bat" echo title Barnaby the Octopus
>> "Run.bat" echo cd /d "%%~dp0"
>> "Run.bat" echo if exist ".venv\Scripts\python.exe" goto python_ok
>> "Run.bat" echo echo Barnaby is not installed yet. Please run install.bat first.
>> "Run.bat" echo pause
>> "Run.bat" echo exit /b 1
>> "Run.bat" echo.
>> "Run.bat" echo :python_ok
>> "Run.bat" echo if exist "barnaby.py" goto app_ok
>> "Run.bat" echo echo barnaby.py is missing. Please reinstall Barnaby.
>> "Run.bat" echo pause
>> "Run.bat" echo exit /b 1
>> "Run.bat" echo.
>> "Run.bat" echo :app_ok
>> "Run.bat" echo ".venv\Scripts\python.exe" "barnaby.py"

:shortcut_create
powershell -NoProfile -ExecutionPolicy Bypass -Command "$desktop=[Environment]::GetFolderPath('Desktop'); $shortcut=Join-Path $desktop 'Barnaby the Octopus.lnk'; $shell=New-Object -ComObject WScript.Shell; $link=$shell.CreateShortcut($shortcut); $link.TargetPath=(Join-Path (Get-Location) 'Run.bat'); $link.WorkingDirectory=(Get-Location).Path; $link.Description='Run Barnaby the Octopus desktop helper'; $link.Save()"

echo Barnaby installed.
echo Run.bat is ready.
echo uninstall000.bat is ready.
echo A desktop shortcut was created when Windows allowed it.
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Barnaby installed! Open Run.bat or the desktop shortcut to start him. uninstall000.bat can remove him later.','Barnaby is ready','OK','Information')"
pause
