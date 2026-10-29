@echo off
chcp 65001 > nul
title Tự Động Đồng Bộ Hóa Đơn Petrolimex - POVINA
cd /d "%~dp0"

echo ===================================================================
echo   HỆ THỐNG TỰ ĐỘNG QUÉT MAIL & TẢI HÓA ĐƠN PETROLIMEX
echo   CÔNG TY CỔ PHẦN THÉP POVINA
echo ===================================================================
echo Chọn chế độ quét hóa đơn:
echo   [1] Quét TOÀN BỘ hòm thư 'vantai povina' (Tất cả từ trước đến nay - Mặc định)
echo   [2] Quét TẤT CẢ hòm thư trong Outlook (Cả vantai povina + Phòng kế toán...)
echo   [3] Quét toàn bộ tháng này (Tháng 10)
echo   [4] Quét tuần này (7 ngày gần nhất)
echo   [5] Quét hôm nay
echo ===================================================================

set /p choice="Nhập lựa chọn [1-5] (Nhấn Enter để chọn 1): "
if "%choice%"=="" set choice=1

set MAILBOX="vantai povina"
set FILTER=all

if "%choice%"=="1" (
    set MAILBOX="vantai povina"
    set FILTER=all
)
if "%choice%"=="2" (
    set MAILBOX="all"
    set FILTER=all
)
if "%choice%"=="3" (
    set MAILBOX="vantai povina"
    set FILTER=month
)
if "%choice%"=="4" (
    set MAILBOX="vantai povina"
    set FILTER=week
)
if "%choice%"=="5" (
    set MAILBOX="vantai povina"
    set FILTER=today
)

echo.
echo ===================================================================
echo Đang quét Outlook...
echo Hộp thư: %MAILBOX%  ^| Chế độ lọc: %FILTER%
echo Vui lòng đợi trong giây lát...
echo ===================================================================
python cli.py --mailbox %MAILBOX% --filter %FILTER% --limit 0 --open-excel

echo ===================================================================
echo   HOÀN THÀNH ĐỒNG BỘ! Dữ liệu đã cập nhật lên Excel và GitHub Pages.
echo ===================================================================
pause
