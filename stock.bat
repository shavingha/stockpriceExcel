@echo off
SETLOCAL EnableDelayedExpansion

:: 设置Python版本和安装路径变量
set PYTHON_VERSION=3.11.0
set PYTHON_HOME=C:\Python%PYTHON_VERSION%

:: 下载并安装Python 3.11.0
echo Downloading Python %PYTHON_VERSION%...
powershell -Command "Invoke-WebRequest -Uri https://www.python.org/ftp/python/%PYTHON_VERSION%/python-%PYTHON_VERSION%-amd64.exe -OutFile python-%PYTHON_VERSION%-amd64.exe"
echo Installing Python %PYTHON_VERSION%...
start /wait "" "python-%PYTHON_VERSION%-amd64.exe" /quiet InstallAllUsers=1 PrependPath=1

:: 等待Python安装完成
:wait_for_python
timeout /t 5
where python >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Waiting for Python installation to complete...
    goto wait_for_python
)

echo Python installed successfully.

:: 确保pip已安装
python -m ensurepip --default-pip

:: 升级pip到最新版本
echo Upgrading pip...
python -m pip install --upgrade pip

:: 安装akshare
echo Installing akshare...
python -m pip install akshare

:: 安装xlwings
echo Installing xlwings...
python -m pip install xlwings

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