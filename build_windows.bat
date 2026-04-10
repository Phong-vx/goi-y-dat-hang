@echo off
chcp 65001 > nul
title Build - Goi Y Dat Hang

echo.
echo  ============================================
echo    GOI Y DAT HANG - BUILD EXE
echo  ============================================
echo.

:: Tìm Python (thử py, python, python3)
set PYTHON=
for %%P in (py python python3) do (
    if not defined PYTHON (
        %%P --version > nul 2>&1
        if not errorlevel 1 set PYTHON=%%P
    )
)

if not defined PYTHON (
    echo  LOI: Khong tim thay Python tren may!
    echo.
    echo  Vui long:
    echo    1. Tai Python tai: https://www.python.org/downloads/
    echo    2. Khi cai, TICK vao o "Add Python to PATH"
    echo    3. Sau do chay lai file nay
    echo.
    pause & exit /b
)

echo  Tim thay Python: %PYTHON%
echo.

echo  [1/3] Cai dat thu vien can thiet...
%PYTHON% -m pip install pandas openpyxl Pillow pyinstaller --quiet
if %errorlevel% neq 0 (
    echo  LOI: Khong cai duoc thu vien. Kiem tra ket noi mang.
    pause & exit /b
)

echo  [2/3] Dang build file .exe (cho 1-2 phut)...
%PYTHON% -m PyInstaller --onefile --windowed --name "GoiYDatHang" --clean main.py > nul 2>&1
if %errorlevel% neq 0 (
    echo  LOI: Build that bai.
    pause & exit /b
)

echo  [3/3] Copy file .exe ra thu muc chinh...
copy /y "dist\GoiYDatHang.exe" "GoiYDatHang.exe" > nul

echo.
echo  ============================================
echo   HOAN TAT!
echo   File: GoiYDatHang.exe da san sang.
echo   Double-click de mo app.
echo  ============================================
echo.

start "" "GoiYDatHang.exe"
