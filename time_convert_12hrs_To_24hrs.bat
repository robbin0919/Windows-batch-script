@echo off
setlocal enabledelayedexpansion
set "delimiter=/: "
set  "hour12_str=10/22/2024 6:44:30 PM"
call :ConvertTo24Hour "!hour12_str!"
echo 12小時制時間: !hour12_str!
echo 24小時制時間: %result%
goto :eof

:ConvertTo24Hour
set "timeString=%~1"
:: 拆分時間字符串
for /f "tokens=1-7 delims=%delimiter%" %%a in ("%timeString%") do (
    set month=%%a
    set day=%%b
    set year=%%c
    set hour=%%d
    set minute=%%e
    set second=%%f
    set ampm=%%g
)

:: 判斷 AM/PM 並調整小時
if "%ampm%"=="PM" (
    if %hour% neq 12 (
        set /a hour=%hour%+12
    )
) else (
    if %hour% equ 12 (
        set hour=0
    )
)

:: 組合 24 小時制時間
set "newTime=%year%/%month%/%day% %hour%:%minute%:%second%"
 
set result=%newTime%
goto :eof