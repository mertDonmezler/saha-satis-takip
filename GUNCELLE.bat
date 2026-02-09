@echo off
chcp 65001 >nul
echo.
echo ===================================================
echo   MASTER DATA GUNCELLEME
echo   Klasordeki tum Excel dosyalari taranacak
echo ===================================================
echo.
python "%~dp0master_data_olustur.py"
echo.
echo ---------------------------------------------------
echo   MASTER_DATA.xlsx guncellendi!
echo   Dosya: %~dp0MASTER_DATA.xlsx
echo ---------------------------------------------------
echo.
pause
