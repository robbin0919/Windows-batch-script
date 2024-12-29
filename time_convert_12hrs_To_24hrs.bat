@echo off
setlocal enabledelayedexpansion
set "delimiter=/: "
set  "hour12_str=10/22/2024 6:44:30 PM"
call :ConvertTo24Hour "!hour12_str!"
echo 12�p�ɨ�ɶ�: !hour12_str!
echo 24�p�ɨ�ɶ�: %result%
goto :eof

:ConvertTo24Hour
set "timeString=%~1"
:: ����ɶ��r�Ŧ�
for /f "tokens=1-7 delims=%delimiter%" %%a in ("%timeString%") do (
    set month=%%a
    set day=%%b
    set year=%%c
    set hour=%%d
    set minute=%%e
    set second=%%f
    set ampm=%%g
)

:: �P�_ AM/PM �ýվ�p��
if "%ampm%"=="PM" (
    if %hour% neq 12 (
        set /a hour=%hour%+12
    )
) else (
    if %hour% equ 12 (
        set hour=0
    )
)

:: �զX 24 �p�ɨ�ɶ�
set "newTime=%year%/%month%/%day% %hour%:%minute%:%second%"
 
set result=%newTime%
goto :eof