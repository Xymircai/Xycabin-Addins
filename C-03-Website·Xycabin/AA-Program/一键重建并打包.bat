@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0.."
rem 统一网站入口。后台重建验证两个安装包，不自动安装、不关闭 WPS。
set "WPS_TASK_PYTHON="
if exist "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" set "WPS_TASK_PYTHON=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if not defined WPS_TASK_PYTHON if exist "%USERPROFILE%\.workbuddy\binaries\python\versions\3.13.12\python.exe" set "WPS_TASK_PYTHON=%USERPROFILE%\.workbuddy\binaries\python\versions\3.13.12\python.exe"
if not defined WPS_TASK_PYTHON set "WPS_TASK_PYTHON=python"
set "PYTHONDONTWRITEBYTECODE=1"
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"
"%WPS_TASK_PYTHON%" -c "print('\u8bf7\u9009\u62e9\u9a8c\u8bc1\u65b9\u5f0f\uff1a\n1. \u589e\u91cf\u68c0\u67e5\uff08\u590d\u7528\u8f93\u5165\u672a\u53d8\u5316\u7684\u5df2\u9a8c\u8bc1\u7ed3\u679c\uff09\n2. \u5b8c\u6574\u9a8c\u8bc1\uff08\u91cd\u65b0\u6784\u5efa\uff0c\u4e24\u4e2a\u4ea7\u54c1\u540e\u53f0\u56de\u5f52\u5404\u4e00\u8f6e\uff09')"
choice /c 12 /n
if errorlevel 255 goto fail
if errorlevel 2 goto full
if errorlevel 1 goto incremental
exit /b 130
:full
set "WPS_BUILD_MODE=full"
goto run
:incremental
set "WPS_BUILD_MODE=incremental"
:run
"%WPS_TASK_PYTHON%" rebuild.py --background --mode %WPS_BUILD_MODE% %*
if errorlevel 1 goto fail
"%WPS_TASK_PYTHON%" validate.py --background --mode %WPS_BUILD_MODE% %*
if errorlevel 1 goto validatefail
echo.
chcp 65001 >nul
"%WPS_TASK_PYTHON%" -c "print('\u540e\u53f0\u9a8c\u8bc1\u5df2\u5b8c\u6210\uff1b\u524d\u53f0\u7a97\u53e3\u4e0e\u771f\u5b9e WPS \u64cd\u4f5c\u8bf7\u53e6\u884c\u9a8c\u6536\u3002'); print('\u4e24\u4e2a\u5b89\u88c5\u5305\u5df2\u540c\u6b65\u66f4\u65b0\u3002\u8bf7\u4fdd\u5b58\u6587\u4ef6\u3001\u5b8c\u5168\u9000\u51fa WPS \u540e\u5206\u522b\u91cd\u88c5\u3002')"
pause
exit /b 0
:fail
echo.
chcp 65001 >nul
"%WPS_TASK_PYTHON%" -c "print('\u7edf\u4e00\u6784\u5efa\u5931\u8d25\uff0c\u8bf7\u68c0\u67e5\u4e0a\u65b9\u9519\u8bef\u53ca\u65e5\u5fd7\u3002\u672a\u5b8c\u6210\u7684\u6784\u5efa\u4e0d\u4f1a\u66ff\u6362\u6b63\u5f0f\u5b89\u88c5\u5305\u3002')"
pause
exit /b 1
:validatefail
echo.
chcp 65001 >nul
"%WPS_TASK_PYTHON%" -c "print('\u6536\u5c3e\u9a8c\u8bc1\u5931\u8d25\uff0c\u8bf7\u68c0\u67e5\u4e0a\u65b9\u9519\u8bef\u53ca\u65e5\u5fd7\u3002\u5b89\u88c5\u5305\u5df2\u751f\u6210\u4f46\u672a\u901a\u8fc7\u6b63\u5f0f\u4ea7\u7269\u6838\u5bf9\uff0c\u8bf7\u4fee\u590d\u540e\u91cd\u8dd1\u3002')"
pause
exit /b 1
