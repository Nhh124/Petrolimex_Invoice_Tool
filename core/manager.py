import os
import sys
from datetime import datetime

from .outlook_reader import OutlookReader
from .portal_downloader import PortalDownloader
from .invoice_parser import parse_invoice_xml
from .excel_manager import ExcelManager

class PetrolimexInvoiceManager:
    def __init__(self, base_dir):
        self.base_dir = base_dir
        self.pdf_dir = os.path.join(base_dir, "invoices_pdf")
        self.xml_dir = os.path.join(base_dir, "invoices_xml")
        self.data_dir = os.path.join(base_dir, "data")
        self.excel_path = os.path.join(self.data_dir, "Bao_Cao_Hoa_Don_Petrolimex.xlsx")

        self.downloader = PortalDownloader(self.pdf_dir, self.xml_dir)
        self.excel_mgr = ExcelManager(self.excel_path)

    def run_sync(self, mailbox_name="vantai povina", limit=20, log_callback=None):
        """
        Tiến hành quy trình đồng bộ hóa đơn từ Outlook -> Tải Portal -> Xuất Excel
        """
        def log(msg):
            print(msg)
            if log_callback:
                log_callback(msg)

        log("=" * 60)
        log(f"BẮT ĐẦU ĐỒNG BỘ HÓA ĐƠN PETROLIMEX - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        log(f"Hộp thư: {mailbox_name} | Giới hạn quét: {limit} email gần nhất")
        log("=" * 60)

        # 1. Đọc danh sách đã có trong Excel để không quét trùng
        existing_keys = self.excel_mgr.get_existing_invoices()
        log(f"Dữ liệu hiện tại trong Excel: {len(existing_keys)} hóa đơn đã lưu.")

        # 2. Đọc mail từ Outlook
        try:
            reader = OutlookReader()
            emails = reader.fetch_petrolimex_emails(mailbox_name=mailbox_name, limit=limit)
        except Exception as e:
            log(f"LỖI KHI KẾT NỐI OUTLOOK: {e}")
            return {
                'status': 'error',
                'message': str(e),
                'processed': 0,
                'total_found': 0
            }

        log(f"Tìm thấy {len(emails)} email hóa đơn Petrolimex trong Outlook.")

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

            # Nếu XML thiếu một số trường, lấy dự phòng từ email
            if not parsed.get('mau_so'):
                parsed['mau_so'] = em.get('mau_so', '')
            if not parsed.get('ky_hieu'):
                parsed['ky_hieu'] = em.get('ky_hieu', '')
            if not parsed.get('so_hd'):
                parsed['so_hd'] = em.get('so_hd', '')
            if not parsed.get('ma_tra_cuu'):
                parsed['ma_tra_cuu'] = em.get('ma_tra_cuu', '')

            processed_invoices.append(parsed)

            # Cập nhật từng hóa đơn vào Excel ngay lập tức để không mất dữ liệu
            self.excel_mgr.append_invoices([parsed])
            log(f"Đã lưu vào Excel: Số HĐ {parsed['so_hd']} - Xe {parsed.get('bien_so_xe','')} - {parsed.get('ten_hang','')} - {parsed.get('tong_tien_tt',0):,.0f} VNĐ")

        log("\n" + "=" * 60)
        log(f"HOÀN THÀNH ĐỒNG BỘ! Tổng số hóa đơn mới đã xử lý: {len(processed_invoices)}")
        log(f"File Excel tổng hợp: {self.excel_path}")
        log("=" * 60)

        return {
            'status': 'success',
            'total_found': len(emails),
            'new_processed': len(processed_invoices),
            'excel_path': self.excel_path,
            'invoices': processed_invoices
        }
