import os
import sys
import argparse
import subprocess

# Set encoding utf-8 for Windows console
sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from core.manager import PetrolimexInvoiceManager

def main():
    parser = argparse.ArgumentParser(description="Tự động tải và tổng hợp hóa đơn điện tử Petrolimex từ Outlook")
    parser.add_argument("--mailbox", default="vantai povina", help="Tên hòm thư Outlook cần quét (mặc định: 'vantai povina')")
    parser.add_argument("--limit", type=int, default=10, help="Số lượng email gần nhất để quét (mặc định: 10, đặt 0 để quét tất cả)")
    parser.add_argument("--open-excel", action="store_true", help="Tự động mở file Excel sau khi hoàn thành")

    args = parser.parse_args()

    limit_val = None if args.limit == 0 else args.limit

    manager = PetrolimexInvoiceManager(BASE_DIR)
    result = manager.run_sync(mailbox_name=args.mailbox, limit=limit_val)

    if args.open_excel and os.path.exists(result.get('excel_path', '')):
        try:
            os.startfile(result['excel_path'])
        except Exception as e:
            print(f"Không thể mở file Excel tự động: {e}")

if __name__ == "__main__":
    main()
