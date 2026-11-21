@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 goto NO_NODE

echo 浏览器模式：不需要 Electron，直接用默认浏览器打开简阅。
echo 关掉这个窗口就会停止服务。
node server.js --port 8792 --token dev --open
goto END

:NO_NODE
echo 没有找到 Node.js，请先安装：https://nodejs.org/
pause

:END
