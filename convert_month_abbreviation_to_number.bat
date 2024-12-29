@echo off
setlocal enabledelayedexpansion
set "delimiter=/: "
set  "hour12_str=11/Nov/24 3:24 PM"
:: 建立月份對照表（使用陣列）
set "month[Jan]=01"
set "month[Feb]=02"
set "month[Mar]=03"
set "month[Apr]=04"
set "month[May]=05"
set "month[Jun]=06"
set "month[Jul]=07"
set "month[Aug]=08"
set "month[Sep]=09"
set "month[Oct]=10"
set "month[Nov]=11"
set "month[Dec]=12"

:: main
call :ConvertTo24Hour "!hour12_str!"
echo 12小時制時間: !hour12_str!
echo 24小時制時間: %result%
goto :eof

:ConvertTo24Hour
set "timeString=%~1"
:: 拆分時間字符串
for /f "tokens=1-7 delims=%delimiter%" %%a in ("%timeString%") do (
    set day=%%a
    set month=%%b
    set year=%%c
    set hour=%%d
    set minute=%%e
    set ampm=%%f
)
:: 根據月份縮寫查找對應數字
set monthNum=!month[%month%]!

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
set "newTime=20%year%/%monthNum%/%day% %hour%:%minute%"
 
set result=%newTime%
goto :eof