@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

rem "�����̗L���ɂ���ď�����؂蕪��"
if "%~1"=="" (
	rem "���������̏ꍇ��input�f�B���N�g���̒���ϊ�"
	powershell -NoProfile -ExecutionPolicy Unrestricted -File .\src\Convert-AllCsvInInputDir2Excel.ps1
	exit /b
)

rem "�����L��̏ꍇ�͈����̃p�X�̃t�@�C����ϊ�"
set ARGS=
:LOOP
set ARGS=!ARGS! "%~1"
shift
if not "%~1"=="" goto LOOP
powershell -NoProfile -ExecutionPolicy Unrestricted -File .\src\Convert-AllCsvOfArgs2Excel.ps1 %ARGS%
exit /b