@echo off
SETLOCAL EnableDelayedExpansion

:: 设置Python版本和安装路径变量
set PYTHON_VERSION=3.12.0
set PYTHON_HOME=%USERPROFILE%\Python%PYTHON_VERSION%

:: 检查当前已安装的Python版本
echo Checking for existing Python installations...
for /f "tokens=2 delims=:" %%v in ('python -V 2^>nul') do set INSTALLED_VERSION=%%v
set INSTALLED_VERSION=!INSTALLED_VERSION: =!

if defined INSTALLED_VERSION (
    echo Installed Python version: !INSTALLED_VERSION!
    if "!INSTALLED_VERSION!"=="Python %PYTHON_VERSION%" (
        echo The installed version is already up-to-date. Skipping uninstall and install steps.
        goto install_packages
    )
)

:: 卸载现有的Python版本（如果存在）
where msiexec >nul 2>&1
if %ERRORLEVEL% == 0 (
    echo Attempting to uninstall existing Python installation using msiexec...
    msiexec /x {E99C8B5C-81E2-4B94-B1E6-23F0D674F6F0} /qn
) else (
    echo msiexec not found. Unable to proceed with uninstallation.
    echo Please manually uninstall Python and rerun this script.
    pause
    exit /b 1
)

:: 下载并安装Python 3.13.0（安装过程非静默化）
echo Downloading Python %PYTHON_VERSION%...
powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/%PYTHON_VERSION%/python-%PYTHON_VERSION%-amd64.exe' -OutFile python-%PYTHON_VERSION%-amd64.exe"
if %ERRORLEVEL% neq 0 (
    echo Download failed. Please check your internet connection.
    pause
    exit /b 1
)

echo Installing Python %PYTHON_VERSION%...
start /wait python-%PYTHON_VERSION%-amd64.exe /passive InstallAllUsers=0 PrependPath=1
if %ERRORLEVEL% neq 0 (
    echo Python installation failed.
    pause
    exit /b 1
)

:: 等待Python安装完成
:wait_for_python
echo Waiting for Python installation to complete...
timeout /t 5 >nul
where python >nul 2>&1
if %ERRORLEVEL% neq 0 (
    goto wait_for_python
)
echo Python installed successfully.

:: 将Python路径添加到PATH
echo Adding Python to PATH...
setx PATH "%PATH%;C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python%PYTHON_VERSION%;C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python%PYTHON_VERSION%\Scripts"

:install_packages
:: 确保pip已安装并升级
echo Upgrading pip...
python -m ensurepip --upgrade
if %ERRORLEVEL% neq 0 (
    echo Pip upgrade failed.
    pause
    exit /b 1
)

:: 安装akshare
echo Installing akshare...
python -m pip install akshare --user
if %ERRORLEVEL% neq 0 (
    echo Akshare installation failed.
    pause
    exit /b 1
)

:: 安装xlwings
echo Installing xlwings...
python -m pip install xlwings --user
if %ERRORLEVEL% neq 0 (
    echo xlwings installation failed.
    pause
    exit /b 1
)

:: 检查安装是否成功
where pip >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Pip installation failed. Exiting.
    pause
    exit /b %ERRORLEVEL%
)

where akshare >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo akshare installation failed. Exiting.
    pause
    exit /b %ERRORLEVEL%
)

where xlwings >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo xlwings installation failed. Exiting.
    pause
    exit /b %ERRORLEVEL%
)

echo Installation completed successfully.
pause
ENDLOCAL
