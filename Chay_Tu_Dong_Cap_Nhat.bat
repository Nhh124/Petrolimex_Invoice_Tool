@echo off
chcp 65001 > nul
title Tự Động Đồng Bộ Hóa Đơn Petrolimex
cd /d "%~dp0"

echo ===================================================================
echo   HỆ THỐNG TỰ ĐỘNG QUÉT MAIL & TẢI HÓA ĐƠN PETROLIMEX
echo   CÔNG TY CỔ PHẦN THÉP POVINA
echo ===================================================================
echo Đang quét hộp thư Outlook và tải hóa đơn mới...

python cli.py --mailbox "vantai povina" --limit 30 --open-excel

echo ===================================================================
echo   HOÀN THÀNH ĐỒNG BỘ!
echo ===================================================================
pause
