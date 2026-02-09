@echo off
chcp 65001 >nul
echo.
echo ===================================================
echo   DOSYA IZLEME MODU
echo   Yeni Excel dosyasi eklendiginde otomatik gunceller
echo   Durdurmak icin Ctrl+C
echo ===================================================
echo.
python "%~dp0master_data_olustur.py" --izle
