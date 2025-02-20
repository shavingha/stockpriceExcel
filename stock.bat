@echo off
SETLOCAL EnableDelayedExpansion

:: 设置Python版本和安装路径变量
set PYTHON_VERSION=3.13.0
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
where python >nul 2>&1
if %ERRORLEVEL% == 0 (
    echo Uninstalling existing Python installation...
    :: Uninstall Python using the uninstall command (this assumes Python is installed with an MSI installer)
    msiexec /x {E99C8B5C-81E2-4B94-B1E6-23F0D674F6F0} /quiet
    echo Python uninstalled successfully.
)

:: 下载并安装Python 3.13.0
echo Downloading Python %PYTHON_VERSION%...
powershell -Command "Invoke-WebRequest -Uri https://www.python.org/ftp/python/%PYTHON_VERSION%/python-%PYTHON_VERSION%-amd64.exe -OutFile python-%PYTHON_VERSION%-amd64.exe"
echo Installing Python %PYTHON_VERSION%...
start /wait "" "python-%PYTHON_VERSION%-amd64.exe" /quiet InstallAllUsers=0 PrependPath=1

:: 等待Python安装完成
:wait_for_python
timeout /t 5
where python >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Waiting for Python installation to complete...
    goto wait_for_python
)

echo Python installed successfully.

:: 删除安装文件
echo Deleting the downloaded Python installer...
del /f /q python-%PYTHON_VERSION%-amd64.exe

:install_packages
:: 确保pip已安装
python -m ensurepip --default-pip

:: 升级pip到最新版本
echo Upgrading pip...
python -m pip install --upgrade pip --user

:: 安装akshare
echo Installing akshare...
python -m pip install akshare --user

:: 安装xlwings
echo Installing xlwings...
python -m pip install xlwings --user

:: 检查安装是否成功
where pip >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Pip installation failed.
    exit /b %ERRORLEVEL%
)

where akshare >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo akshare installation failed.
    exit /b %ERRORLEVEL%
)

where xlwings >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo xlwings installation failed.
    exit /b %ERRORLEVEL%
)

echo Installation completed successfully.
ENDLOCAL

pause