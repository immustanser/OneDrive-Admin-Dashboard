@echo off
cd /d "%~dp0"
call npm run build
echo BUILD_DONE
