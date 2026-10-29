import os
import sys
import argparse

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from core.manager import PetrolimexInvoiceManager

def main():
    parser = argparse.ArgumentParser(description="Tự động tải và tổng hợp hóa đơn điện tử Petrolimex từ Outlook")
    parser.add_argument("--mailbox", default="vantai povina", help="Tên hòm thư Outlook cần quét (mặc định: 'vantai povina')")
    parser.add_argument("--filter", default="month", choices=["all", "today", "week", "month", "last_month", "custom"], 
                        help="Bộ lọc thời gian: 'today' (hôm nay), 'week' (7 ngày qua), 'month' (tháng này), 'last_month' (tháng trước), 'all' (toàn bộ), 'custom' (tùy chọn)")
    parser.add_argument("--start-date", default=None, help="Ngày bắt đầu (YYYY-MM-DD) khi chọn --filter custom")
    parser.add_argument("--end-date", default=None, help="Ngày kết thúc (YYYY-MM-DD) khi chọn --filter custom")
    parser.add_argument("--limit", type=int, default=0, help="Giới hạn số lượng email cần xử lý (mặc định: 0 = không giới hạn)")
    parser.add_argument("--no-push", action="store_true", help="Không tự động push lên GitHub sau khi đồng bộ")
    parser.add_argument("--open-excel", action="store_true", help="Tự động mở file Excel sau khi hoàn thành")

    args = parser.parse_args()

    limit_val = None if args.limit <= 0 else args.limit

    manager = PetrolimexInvoiceManager(BASE_DIR)
    result = manager.run_sync(
        mailbox_name=args.mailbox,
        limit=limit_val,
        time_filter=args.filter,
        start_date=args.start_date,
        end_date=args.end_date,
        auto_push_github=not args.no_push
    )

    if args.open_excel and os.path.exists(result.get('excel_path', '')):
        try:
            os.startfile(result['excel_path'])
        except Exception as e:
            print(f"Không thể mở file Excel tự động: {e}")

if __name__ == "__main__":
    main()
