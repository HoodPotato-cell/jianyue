@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set "LOG=启动日志.txt"

echo === 简阅启动 %DATE% %TIME% === > "%LOG%"

where node >> "%LOG%" 2>&1
if errorlevel 1 goto NO_NODE

echo [node] >> "%LOG%"
node -v >> "%LOG%" 2>&1
npm -v >> "%LOG%" 2>&1

if not exist "node_modules" goto INSTALL
if not exist "node_modules\electron\dist\electron.exe" goto REINSTALL
goto RUN

:REINSTALL
echo node_modules 不完整，重新安装依赖... 
echo [reinstall 原因：electron.exe 缺失] >> "%LOG%"
goto INSTALL

:INSTALL
echo.
echo 正在安装依赖，首次约 100 MB，需要联网，请耐心等待...
echo [install] >> "%LOG%"
set "ELECTRON_MIRROR=https://npmmirror.com/mirrors/electron/"
call npm install >> "%LOG%" 2>&1
if errorlevel 1 goto INSTALL_FAIL
if not exist "node_modules\electron\dist\electron.exe" goto INSTALL_FAIL

:RUN
echo.
echo 正在启动简阅...
echo [run] >> "%LOG%"
call npm start >> "%LOG%" 2>&1
echo.
echo 简阅已退出。日志：%CD%\%LOG%
goto END

:NO_NODE
echo.
echo 没有找到 Node.js。请先安装 LTS 版本：https://nodejs.org/
echo 装完重新双击本文件。
goto END

:INSTALL_FAIL
echo.
echo 依赖安装失败。详情已写入 %CD%\%LOG%
echo 常见原因：网络或镜像不通。可以把日志发给我。
goto END

:END
echo.
pause
