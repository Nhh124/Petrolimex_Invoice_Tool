@echo off
chcp 65001 > nul
cd /d "C:\Users\haiho\OneDrive\Desktop\Petrolimex_Invoice_Tool"
echo [%date% %time%] Bat dau quet toan bo hoa don... > sync.log
python cli.py --mailbox "vantai povina" --filter all --limit 0 >> sync.log 2>&1
echo [%date% %time%] Hoan thanh! >> sync.log
