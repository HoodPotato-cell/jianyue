@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

if not exist "node_modules\electron\dist\electron.exe" (
  echo 先安装依赖...
  set "ELECTRON_MIRROR=https://npmmirror.com/mirrors/electron/"
  call npm install
)

call npm run dist
echo.
echo 打包完成，产物在 dist\ 目录。
echo 如果失败，把上面的报错发给我。
pause
