rem 
@echo off
setlocal enabledelayedexpansion
set work_disk=D:
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

）

