@echo off
setlocal

:: 检测 node 是否已安装
node -v >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Node.js 未安装. 正在安装 Node.js...
    :: 使用 Chocolatey 安装 Node.js
    @powershell -NoProfile -ExecutionPolicy Bypass -Command "Set-ExecutionPolicy Bypass -Scope Process; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))"
    choco install -y nodejs
)

:: 再次检测 node 是否已安装成功
node -v >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Node.js 安装失败.
    exit /b 1
)

echo Node.js 已安装. 正在启动 server.js...
node server.js

endlocal
