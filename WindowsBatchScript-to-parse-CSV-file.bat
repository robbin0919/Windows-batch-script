@echo off
setlocal enabledelayedexpansion
set work_disk=f:
set work_dir=F:\LAB
set checkmarx=F:\LAB\checkmarx_report.csv
set "delimiter=,"
set "delimiter2=^"
set var1=""
set "var2=^^" 
set "var3=^ ^" 
set var4=""
set "keyword2="" 
set "replacement2=" 
set "keyword1=","" 
set "replacement1=^" 
set "keyword=^^" 
set "replacement=^ ^" 
!work_disk!
cd !work_dir!

for /F "tokens=* delims=" %%a in ('powershell "import-csv %checkmarx%  |convertto-csv "') do (
set var1=%%a
echo source:[!var1!] 
set "var4=!var1:%keyword1%=%replacement1%!" 
set "var4=!var4:%keyword%=%replacement%!" 
set "var4=!var4:%keyword%=%replacement%!" 
set "var4=!var4:%keyword2%=%replacement2%!" 
echo b:[!var4!]  
for /f "tokens=1-29* delims=%delimiter2%" %%@ in ("!var4!") do (
    echo:
    echo  1=%%@
    echo  2=%%A
    echo  3=%%B
    echo  4=%%C
    echo  5=%%D
    echo  6=%%E
    echo  7=%%F
    echo  8=%%G
    echo  9=%%H
    echo 10=%%I
    echo 11=%%J
    echo 12=%%K
    echo 13=%%L
    echo 14=%%M
    echo 15=%%N
    echo 16=%%O
    echo 17=%%P
    echo 18=%%Q
    echo 19=%%R
    echo 20=%%S
    echo 21=%%T
    echo 22=%%U
    echo 23=%%V
    echo 24=%%W
    echo 25=%%X
    echo 26=%%Y
    echo 27=%%Z
    echo 28=%%[
    echo 29=%%\
    echo 30=%%]
    echo 31=%%^^
	for /F "tokens=1-30* delims=%delimiter2%" %%a in ("%%]") do (
        echo A01=%%a
        echo A02=%%b
        echo A03=%%c
        echo A04=%%d
        echo A05=%%e
        echo A06=%%f
        echo A07=%%g
        echo A08=%%h
        echo A09=%%i
        echo A10=%%j
        echo A11=%%k
        echo A12=%%l
        echo A13=%%m
        echo A14=%%n
        echo A15=%%o
        echo A16=%%p
        echo A17=%%q
        echo A18=%%r
        echo A19=%%s
        echo A20=%%t
        echo A21=%%u
        echo A22=%%v
        echo A23=%%w
        echo A24=%%x
        echo A25=%%y
        echo A26=%%z
	)
)
)
  
 
:: 設定原始字符串、關鍵字和替換字符串
rem set "originalString=這是我的原始字串abc,其中包含關鍵字abc"
rem set "keyword=abc"
rem set "replacement=AAA"

:: 進行替換
rem set "newString=!originalString:%keyword%=%replacement%!"

:: 輸出結果
rem echo 新的字符串：%newString%