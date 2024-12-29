@echo off
setlocal enabledelayedexpansion
set "delimiter=/: "
set  "hour12_str=11/Nov/24 3:24 PM"
:: �إߤ����Ӫ�]�ϥΰ}�C�^
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
echo 12�p�ɨ�ɶ�: !hour12_str!
echo 24�p�ɨ�ɶ�: %result%
goto :eof

:ConvertTo24Hour
set "timeString=%~1"
:: ����ɶ��r�Ŧ�
for /f "tokens=1-7 delims=%delimiter%" %%a in ("%timeString%") do (
    set day=%%a
    set month=%%b
    set year=%%c
    set hour=%%d
    set minute=%%e
    set ampm=%%f
)
:: �ھڤ���Y�g�d������Ʀr
set monthNum=!month[%month%]!

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
set "newTime=20%year%/%monthNum%/%day% %hour%:%minute%"
 
set result=%newTime%
goto :eof