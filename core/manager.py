import os
import sys
import subprocess
from datetime import datetime

from .outlook_reader import OutlookReader
from .imap_reader import IMAPReader
from .portal_downloader import PortalDownloader
from .invoice_parser import parse_invoice_xml
from .excel_manager import ExcelManager
from .export_json import export_excel_to_json

class PetrolimexInvoiceManager:
    def __init__(self, base_dir):
        self.base_dir = base_dir
        self.pdf_dir = os.path.join(base_dir, "invoices_pdf")
        self.xml_dir = os.path.join(base_dir, "invoices_xml")
        self.data_dir = os.path.join(base_dir, "data")
        self.excel_path = os.path.join(self.data_dir, "Bao_Cao_Hoa_Don_Petrolimex.xlsx")
        self.json_path = os.path.join(self.data_dir, "invoices.json")

        self.downloader = PortalDownloader(self.pdf_dir, self.xml_dir)
        self.excel_mgr = ExcelManager(self.excel_path)

    def sync_to_github(self, log_callback=None):
        """Tự động commit và push dữ liệu mới lên GitHub Pages"""
        def log(msg):
            print(msg)
            if log_callback:
                log_callback(msg)

        git_exe = r"C:\Users\haiho\AppData\Local\GitHubDesktop\app-3.6.5\resources\app\git\cmd\git.exe"
        if not os.path.exists(git_exe):
            git_exe = "git"

        try:
            log("Đang đồng bộ dữ liệu lên GitHub Pages...")
            subprocess.run([git_exe, "add", "."], cwd=self.base_dir, capture_output=True)
            status_res = subprocess.run([git_exe, "status", "--porcelain"], cwd=self.base_dir, capture_output=True, text=True)
            if not status_res.stdout.strip():
                log("Dữ liệu trên GitHub đã là mới nhất.")
                return

            commit_msg = f"Cap nhat hoa don Petrolimex {datetime.now().strftime('%d/%m/%Y %H:%M')}"
            subprocess.run([git_exe, "commit", "-m", commit_msg], cwd=self.base_dir, capture_output=True)

            token = None
            try:
                import win32cred
                cred = win32cred.CredRead('GitHub - https://api.github.com/Nhh124', win32cred.CRED_TYPE_GENERIC)
                token = cred['CredentialBlob'].decode('utf-16') if b'\x00' in cred['CredentialBlob'] else cred['CredentialBlob'].decode('utf-8')
            except Exception:
                pass

            if token:
                auth_url = f"https://{token}@github.com/Nhh124/Petrolimex_Invoice_Tool.git"
                subprocess.run([git_exe, "remote", "set-url", "origin", auth_url], cwd=self.base_dir, capture_output=True)

            push_res = subprocess.run([git_exe, "push", "origin", "main"], cwd=self.base_dir, capture_output=True, text=True)

            if token:
                clean_url = "https://github.com/Nhh124/Petrolimex_Invoice_Tool.git"
                subprocess.run([git_exe, "remote", "set-url", "origin", clean_url], cwd=self.base_dir, capture_output=True)

            if push_res.returncode == 0:
                log("Đã cập nhật Web GitHub Pages thành công!")
            else:
                log(f"Lỗi khi push GitHub: {push_res.stderr[:200]}")
        except Exception as e:
            log(f"Cảnh báo đồng bộ GitHub: {e}")

    def run_sync(self, mailbox_name="vantai povina", limit=None, time_filter="all", start_date=None, end_date=None, auto_push_github=True, log_callback=None, email_user=None, email_pass=None, email_server=None):
        """
        Tiến hành quy trình đồng bộ hóa đơn từ Outlook hoặc IMAP Server -> Tải Portal -> Xuất Excel -> Cập nhật Web GitHub Pages
        - time_filter: 'all', 'today', 'week', 'month', 'last_month', 'custom'
        - start_date, end_date: chuỗi YYYY-MM-DD
        - limit: số lượng hóa đơn cần tải (None hoặc 0 = không giới hạn)
        - email_user, email_pass, email_server: Thông tin đăng nhập trực tiếp hòm thư qua IMAP
        """
        def log(msg):
            print(msg)
            if log_callback:
                log_callback(msg)

        log("=" * 60)
        log(f"BẮT ĐẦU ĐỒNG BỘ HÓA ĐƠN PETROLIMEX - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        if email_user:
            log(f"Chế độ: Đăng nhập trực tiếp máy chủ Email IMAP ({email_user}) | Bộ lọc: {time_filter}")
        else:
            log(f"Chế độ: Đọc từ Outlook ứng dụng | Hộp thư: {mailbox_name} | Bộ lọc: {time_filter}")
        if time_filter == "custom":
            log(f"Khoảng ngày: Từ {start_date or 'bắt đầu'} đến {end_date or 'hiện tại'}")
        log("=" * 60)

        # 1. Đọc danh sách đã có trong Excel để không quét trùng
        existing_keys = self.excel_mgr.get_existing_invoices()
        inv_count = self.excel_mgr.get_invoice_count()
        log(f"Dữ liệu hiện tại trong Excel: {inv_count} hóa đơn đã lưu.")

        # 2. Đọc mail theo phương thức người dùng chọn
        try:
            if email_user and email_pass:
                log(f"Đang kết nối tới máy chủ hòm thư ({email_server or 'Tự động'})...")
                reader = IMAPReader(username=email_user, password=email_pass, server=email_server)
                emails = reader.fetch_petrolimex_emails(
                    folder="INBOX",
                    limit=limit,
                    time_filter=time_filter,
                    start_date=start_date,
                    end_date=end_date,
                    log_cb=log
                )
            else:
                reader = OutlookReader()
                emails = reader.fetch_petrolimex_emails(
                    mailbox_name=mailbox_name,
                    limit=limit,
                    time_filter=time_filter,
                    start_date=start_date,
                    end_date=end_date
                )
        except Exception as e:
            log(f"LỖI KHI ĐỌC EMAIL: {e}")
            return {
                'status': 'error',
                'message': str(e),
                'processed': 0,
                'total_found': 0
            }

        log(f"Tìm thấy {len(emails)} email hóa đơn Petrolimex hợp lệ.")

        # 3. Lọc ra các hóa đơn chưa được xử lý
        pending_emails = []
        for em in emails:
            key1 = f"{em['ky_hieu']}_{em['so_hd']}"
            fkey = em['ma_tra_cuu']
            if key1 not in existing_keys and fkey not in existing_keys:
                pending_emails.append(em)
            else:
                log(f"-> Đã có trong Excel: {em['ky_hieu']}-{em['so_hd']} (Fkey: {fkey})")

        log(f"\nSố lượng hóa đơn mới cần tải và xử lý: {len(pending_emails)}")

        processed_invoices = []
        for idx, em in enumerate(pending_emails, 1):
            log(f"\n--- [{idx}/{len(pending_emails)}] Đang xử lý: {em['subject']} ---")
            log(f"Mã tra cứu: {em['ma_tra_cuu']} | Số HĐ: {em['so_hd']} | Ký hiệu: {em['ky_hieu']}")

            # Tải XML và PDF từ portal
            res = self.downloader.download_invoice(
                fkey=em['ma_tra_cuu'],
                mau_so=em['mau_so'],
                ky_hieu=em['ky_hieu'],
                so_hd=em['so_hd']
            )

            if not res or not res.get('xml_path'):
                log(f"Không thể tải hóa đơn cho mã {em['ma_tra_cuu']}, bỏ qua.")
                continue

            # Phân tích cú pháp file XML
            parsed = parse_invoice_xml(res['xml_path'])
            if not parsed:
                log(f"Không thể đọc thông tin từ file XML: {res['xml_path']}")
                continue

            # Bổ sung thông tin từ email và file
            parsed['received_time'] = em.get('received_time', '')
            parsed['pdf_path'] = res.get('pdf_path', '')
            parsed['xml_path'] = res.get('xml_path', '')

            if not parsed.get('mau_so'):
                parsed['mau_so'] = em.get('mau_so', '')
            if not parsed.get('ky_hieu'):
                parsed['ky_hieu'] = em.get('ky_hieu', '')
            if not parsed.get('so_hd'):
                parsed['so_hd'] = em.get('so_hd', '')
            if not parsed.get('ma_tra_cuu'):
                parsed['ma_tra_cuu'] = em.get('ma_tra_cuu', '')

            processed_invoices.append(parsed)

            # Cập nhật từng hóa đơn vào Excel ngay lập tức
            self.excel_mgr.append_invoices([parsed])
            log(f"Đã lưu vào Excel: Số HĐ {parsed['so_hd']} - Xe {parsed.get('bien_so_xe','')} - {parsed.get('ten_hang','')} - {parsed.get('tong_tien_tt',0):,.0f} VNĐ")

            # Cập nhật JSON theo thời gian thực mỗi 5 hóa đơn
            if len(processed_invoices) % 5 == 0:
                try:
                    export_excel_to_json(self.excel_path, self.json_path, self.pdf_dir, self.xml_dir)
                except Exception:
                    pass

        # 4. Xuất dữ liệu sang JSON cho Web tĩnh GitHub Pages
        export_excel_to_json(self.excel_path, self.json_path, self.pdf_dir, self.xml_dir)

        # 5. Tự động đồng bộ lên GitHub Pages nếu có dữ liệu mới
        if auto_push_github and len(processed_invoices) > 0:
            self.sync_to_github(log_callback=log)

        log("\n" + "=" * 60)
        log(f"HOÀN THÀNH ĐỒNG BỘ! Tổng số hóa đơn mới đã xử lý: {len(processed_invoices)}")
        log(f"File Excel tổng hợp: {self.excel_path}")
        log(f"Trang Web GitHub Pages: https://nhh124.github.io/Petrolimex_Invoice_Tool/")
        log("=" * 60)

        return {
            'status': 'success',
            'total_found': len(emails),
            'new_processed': len(processed_invoices),
            'excel_path': self.excel_path,
            'github_pages_url': 'https://nhh124.github.io/Petrolimex_Invoice_Tool/',
            'invoices': processed_invoices
        }
