@echo off
REM 发票酱 桌面 exe 编译脚本（OCR 版，含 PP-OCRv5）
call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat" >nul 2>&1
set "PATH=%USERPROFILE%\.cargo\bin;%PATH%"
cd /D "%~dp0"
echo [1/2] npm install ...
call npm install
if %ERRORLEVEL% NEQ 0 goto :npm_install_failed
echo [2/2] tauri build --features ocr (release) ...
call npm run build:ocr
if %ERRORLEVEL% NEQ 0 goto :tauri_build_failed
echo BUILD_DONE
exit /b 0

:npm_install_failed
echo npm install FAILED
exit /b 1

:tauri_build_failed
echo tauri build FAILED
exit /b 1
