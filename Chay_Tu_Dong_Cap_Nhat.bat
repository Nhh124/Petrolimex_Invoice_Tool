@echo off
chcp 65001 > nul
title Tự Động Đồng Bộ Hóa Đơn Petrolimex - POVINA
cd /d "%~dp0"

echo ===================================================================
echo   HỆ THỐNG TỰ ĐỘNG QUÉT MAIL & TẢI HÓA ĐƠN PETROLIMEX
echo   CÔNG TY CỔ PHẦN THÉP POVINA
echo ===================================================================
echo Chọn chế độ quét hóa đơn:
echo   [1] Quét toàn bộ tháng này (Mặc định)
echo   [2] Quét tuần này (7 ngày gần nhất)
echo   [3] Quét hôm nay
echo   [4] Quét tháng trước
echo   [5] Quét TOÀN BỘ hòm thư (Tất cả từ trước đến nay)
echo ===================================================================

set /p choice="Nhập lựa chọn [1-5] (Nhấn Enter để chọn 1): "
if "%choice%"=="" set choice=1

set FILTER=month
if "%choice%"=="1" set FILTER=month
if "%choice%"=="2" set FILTER=week
if "%choice%"=="3" set FILTER=today
if "%choice%"=="4" set FILTER=last_month
if "%choice%"=="5" set FILTER=all

echo.
echo Đang quét hộp thư Outlook theo chế độ: %FILTER%...
python cli.py --mailbox "vantai povina" --filter %FILTER% --limit 0 --open-excel

echo ===================================================================
echo   HOÀN THÀNH ĐỒNG BỘ!
echo ===================================================================
pause
