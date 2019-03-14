@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

rem "引数の有無によって処理を切り分け"
if "%~1"=="" (
	rem "引数無しの場合はinputディレクトリの中を変換"
	powershell -NoProfile -ExecutionPolicy Unrestricted -File .\src\Convert-AllCsvInInputDir2Excel.ps1
	exit /b
)

rem "引数有りの場合は引数のパスのファイルを変換"
set ARGS=
:LOOP
set ARGS=!ARGS! "%~1"
shift
if not "%~1"=="" goto LOOP
powershell -NoProfile -ExecutionPolicy Unrestricted -File .\src\Convert-AllCsvOfArgs2Excel.ps1 %ARGS%
exit /b