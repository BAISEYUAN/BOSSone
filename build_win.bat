@echo off
REM ============================================================
REM BOSSone 打包脚本（Windows）
REM 产物：
REM   dist\scraper.exe   抓取子进程（V15B，onefile 单文件）
REM   dist\BOSSone\      主程序目录（BOSSone.exe + _internal 依赖）
REM 打包完成后把 dist\scraper.exe 复制到 dist\BOSSone\ 下即可整包发送
REM ============================================================
chcp 65001 >nul
cd /d %~dp0

echo [1/2] 打包抓取子进程 scraper.exe ...
python -m PyInstaller --onefile --name scraper --console --clean --noconfirm ^
  "V15B_终极纯CDP_DOM抓_0内部API_不超时.py"
if errorlevel 1 goto :err

echo [2/2] 打包主程序 BOSSone.exe ...
python -m PyInstaller --onedir --name BOSSone --console --clean --noconfirm ^
  --add-data "index.html;." launcher.py
if errorlevel 1 goto :err

echo.
echo 打包完成！
echo   产物：dist\scraper.exe  = 抓取子进程（单文件）
echo        dist\BOSSone\      = 主程序目录
echo 下一步：把 dist\scraper.exe 复制到 dist\BOSSone\ 里，整目录压缩发送。
exit /b 0

:err
echo.
echo 打包失败，请查看上方错误信息。
exit /b 1