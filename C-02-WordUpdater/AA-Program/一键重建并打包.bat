@echo off

chcp 65001 >nul

setlocal

cd /d "%~dp0.."

rem 共用构建入口。仅生成、编译、签名、打包，不安装、不关闭 WPS。

set "WPS_TASK_PYTHON="

if exist "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" set "WPS_TASK_PYTHON=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if not defined WPS_TASK_PYTHON if exist "%USERPROFILE%\.workbuddy\binaries\python\versions\3.13.12\python.exe" set "WPS_TASK_PYTHON=%USERPROFILE%\.workbuddy\binaries\python\versions\3.13.12\python.exe"

if not defined WPS_TASK_PYTHON set "WPS_TASK_PYTHON=python"

set "PYTHONDONTWRITEBYTECODE=1"
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"

"%WPS_TASK_PYTHON%" rebuild.py --background %*

if errorlevel 1 goto fail

"%WPS_TASK_PYTHON%" validate.py --background %*

if errorlevel 1 goto validatefail

echo.

echo 安装包已经生成。安装前请自行保存文档并完全退出 WPS。

set "WPS_TASK_ANSWER="

set /p "WPS_TASK_ANSWER=现在打开安装器？[Y/N] "

if /i "%WPS_TASK_ANSWER%"=="Y" "%WPS_TASK_PYTHON%" rebuild.py --open-installer-only

pause

exit /b 0

:fail

echo.

echo 构建失败，请检查上方错误信息。

pause

exit /b 1

:validatefail

echo.

echo 收尾验证失败，请检查上方错误信息。

pause

exit /b 1

