@echo off
chcp 65001 > nul
title Hệ Thống Quản Lý Hóa Đơn Petrolimex - POVINA
cd /d "%~dp0"

echo ===================================================================
echo   HỆ THỐNG TỰ ĐỘNG TẢI & TỔNG HỢP HÓA ĐƠN PETROLIMEX
echo   CÔNG TY CỔ PHẦN THÉP POVINA
echo ===================================================================
echo Đang khởi động máy chủ Web Dashboard...
echo Đang mở trình duyệt tại: http://localhost:5005

start "" http://localhost:5005
python app.py

pause
