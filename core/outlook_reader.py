import os
import re
import sys
import subprocess
from datetime import datetime, timedelta

def _ensure_desktop():
    """Đảm bảo thread kết nối đúng vào interactive desktop Default của người dùng"""
    try:
        import ctypes
        if hasattr(ctypes, 'windll'):
            user32 = ctypes.windll.user32
            hdesk = user32.OpenDesktopW('Default', 0, False, 0x10000000)
            if hdesk:
                user32.SetThreadDesktop(hdesk)
    except Exception:
        pass

def cleanup_zombie_outlook():
    """Dọn dẹp các tiến trình Outlook chạy ngầm bị treo (-Embedding)"""
    try:
        # Lọc và tắt các tiến trình OUTLOOK.EXE có switch -Embedding
        ps_cmd = (
            "Get-CimInstance Win32_Process -Filter \"Name = 'OUTLOOK.EXE'\" | "
            "Where-Object { $_.CommandLine -like '*-Embedding*' } | "
            "ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"
        )
        subprocess.run(["powershell", "-NoProfile", "-Command", ps_cmd], capture_output=True, timeout=5)
    except Exception:
        pass

def parse_email_content(body_text, html_text=""):
    """
    Trích xuất các thông tin hóa đơn từ nội dung email Petrolimex
    """
    clean_body = body_text.replace('\xa0', ' ')
    
    # Tìm Mẫu số
    m_mau = re.search(r'(?:Mẫu|Mau)\s*s(?:ố|o)\s*:\s*([^\r\n<]+)', clean_body, re.IGNORECASE)
    mau_so = m_mau.group(1).strip() if m_mau else ""
    mau_so = re.sub(r'<[^>]+>', '', mau_so).strip()

    # Tìm Ký hiệu
    m_kh = re.search(r'(?:Ký|Ky)\s*hi(?:ệ|e)u\s*:\s*([^\r\n<]+)', clean_body, re.IGNORECASE)
    ky_hieu = m_kh.group(1).strip() if m_kh else ""
    ky_hieu = re.sub(r'<[^>]+>', '', ky_hieu).strip()

    # Tìm Số hóa đơn
    m_shd = re.search(r'(?:Số|So)\s*(?:hóa\s*đơn|hoa\s*don|HĐ|HD)\s*:\s*([^\r\n<]+)', clean_body, re.IGNORECASE)
    so_hd = m_shd.group(1).strip() if m_shd else ""
    so_hd = re.sub(r'<[^>]+>', '', so_hd).strip()

    # Tìm Mã tra cứu (Fkey)
    m_mtc = re.search(r'(?:Mã|Ma)\s*tra\s*c(?:ứ|u)u\s*:\s*([^\r\n<]+)', clean_body, re.IGNORECASE)
    ma_tra_cuu = m_mtc.group(1).strip() if m_mtc else ""
    ma_tra_cuu = re.sub(r'<[^>]+>', '', ma_tra_cuu).strip()

    # Link portal tra cứu
    portal_link = "https://hoadon.petrolimex.com.vn"
    if html_text:
        links = re.findall(r'href=[\'"]?(https?://[^\'">\s]+)', html_text)
        for l in links:
            if 'petrolimex' in l.lower() or 'hoadon' in l.lower():
                portal_link = l
                break

    return {
        'mau_so': mau_so,
        'ky_hieu': ky_hieu,
        'so_hd': so_hd,
        'ma_tra_cuu': ma_tra_cuu,
        'portal_link': portal_link
    }

class OutlookReader:
    def __init__(self):
        _ensure_desktop()
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()

        # 1. Dọn dẹp tiến trình zombie kẹt ngầm nếu có
        cleanup_zombie_outlook()

        self.app = None

        # 2. Kết nối trực tiếp vào ứng dụng Outlook đang mở trên máy tính
        try:
            self.app = win32com.client.GetActiveObject('Outlook.Application')
        except Exception:
            self.app = None

        if not self.app:
            raise Exception("Chưa phát hiện phần mềm Microsoft Outlook đang mở! Vui lòng nhấp mở ứng dụng Microsoft Outlook trên máy tính (đảm bảo cửa sổ hòm thư Outlook đang hiển thị) rồi bấm 'Bắt Đầu Quét' lại.")

        try:
            self.ns = self.app.GetNamespace('MAPI')
        except Exception as e:
            raise Exception(f"Lỗi truy cập dữ liệu MAPI: {e}. Vui lòng kiểm tra lại ứng dụng Microsoft Outlook.")

    def get_mailboxes(self):
        """Liệt kê các tài khoản/hòm thư trong Outlook"""
        mailboxes = []
        try:
            for f in self.ns.Folders:
                mailboxes.append(f.Name)
        except Exception:
            pass
        return mailboxes

    def get_inbox_folders(self, mailbox_name="vantai povina"):
        """Lấy danh sách các thư mục Hộp thư đến (Inbox) cần quét"""
        target_folders = []
        if mailbox_name and mailbox_name.lower() != "all":
            try:
                for f in self.ns.Folders:
                    if mailbox_name.lower() in f.Name.lower():
                        for sub in f.Folders:
                            if any(k in sub.Name.lower() for k in ['inbox', 'hộp thư đến', 'đến']):
                                target_folders.append(sub)
                if target_folders:
                    return target_folders
            except Exception:
                pass

        # Nếu chọn "all" hoặc không tìm thấy theo tên, lấy tất cả các Inbox
        try:
            for f in self.ns.Folders:
                for sub in f.Folders:
                    if any(k in sub.Name.lower() for k in ['inbox', 'hộp thư đến', 'đến']):
                        target_folders.append(sub)
        except Exception:
            pass
        
        if not target_folders:
            try:
                def_ib = self.ns.GetDefaultFolder(6)
                if def_ib:
                    target_folders.append(def_ib)
            except Exception:
                pass

        return target_folders

    def fetch_petrolimex_emails(self, mailbox_name="vantai povina", limit=None, time_filter="all", start_date=None, end_date=None, log_cb=None):
        """
        Quét các email Petrolimex trong hòm thư với tốc độ cao bằng Cursor MAPI:
        - time_filter: 'all' (toàn bộ), 'today' (hôm nay), 'week' (7 ngày qua), 'month' (tháng này), 'last_month' (tháng trước), 'custom' (tùy chọn ngày)
        - start_date, end_date: Định dạng YYYY-MM-DD khi time_filter='custom'
        - limit: Giới hạn số lượng email (None hoặc 0 = không giới hạn)
        """
        def log(msg):
            if log_cb:
                log_cb(msg)
            else:
                print(msg)

        inboxes = self.get_inbox_folders(mailbox_name)
        if not inboxes:
            raise Exception(f"Không tìm thấy hộp thư đến cho: {mailbox_name}")

        now = datetime.now()
        start_dt = None
        end_dt = None

        if time_filter == "today":
            start_dt = datetime(now.year, now.month, now.day, 0, 0, 0)
        elif time_filter == "week":
            d7 = now - timedelta(days=7)
            start_dt = datetime(d7.year, d7.month, d7.day, 0, 0, 0)
        elif time_filter == "month":
            start_dt = datetime(now.year, now.month, 1, 0, 0, 0)
        elif time_filter == "last_month":
            first_this_month = datetime(now.year, now.month, 1)
            last_day_prev = first_this_month - timedelta(days=1)
            start_dt = datetime(last_day_prev.year, last_day_prev.month, 1, 0, 0, 0)
            end_dt = datetime(last_day_prev.year, last_day_prev.month, last_day_prev.day, 23, 59, 59)
        elif time_filter == "custom":
            if start_date:
                parts = [int(x) for x in start_date.split('-')]
                start_dt = datetime(parts[0], parts[1], parts[2], 0, 0, 0)
            if end_date:
                parts = [int(x) for x in end_date.split('-')]
                end_dt = datetime(parts[0], parts[1], parts[2], 23, 59, 59)

        log(f"Áp dụng bộ lọc thời gian: {time_filter} (Từ: {start_dt or 'Không giới hạn'} Đến: {end_dt or 'Hiện tại'})")

        results = []

        for inbox in inboxes:
            folder_title = f"{inbox.Parent.Name} / {inbox.Name}"
            items = inbox.Items
            try:
                # Sắp xếp từ thư mới nhất đến cũ nhất
                items.Sort('[ReceivedTime]', True)
            except Exception:
                pass

            total_items = 0
            try:
                total_items = items.Count
            except Exception:
                pass

            log(f"Bắt đầu quét hòm thư '{folder_title}' ({total_items} thư)...")

            scanned = 0
            try:
                item = items.GetFirst()
            except Exception:
                item = None

            while item:
                scanned += 1
                if limit and limit > 0 and len(results) >= limit:
                    log(f"Đã đạt giới hạn tối đa {limit} hóa đơn.")
                    return results

                if scanned % 25 == 0:
                    log(f"  -> Đang quét thư [{scanned}/{total_items or '?'}] - Đã tìm thấy {len(results)} HĐ...")

                try:
                    received_time = getattr(item, 'ReceivedTime', None)

                    # Kiểm tra lọc theo thời gian
                    if received_time:
                        try:
                            item_dt = datetime(
                                received_time.year, received_time.month, received_time.day,
                                received_time.hour, received_time.minute, received_time.second
                            )
                            # Nếu thư mới hơn khoảng lọc (khi có end_dt) thì bỏ qua thư này
                            if end_dt and item_dt > end_dt:
                                item = items.GetNext()
                                continue
                            # Vì danh sách đã xếp giảm dần theo ngày, nếu gặp thư cũ hơn start_dt
                            # thì TẤT CẢ các thư tiếp theo đều cũ hơn -> Dừng quét ngay lập tức để tiết kiệm thời gian!
                            if start_dt and item_dt < start_dt:
                                log(f"  -> Đã quét tới ngày {item_dt.strftime('%d/%m/%Y')} (trước mốc {start_dt.strftime('%d/%m/%Y')}), hoàn thành quét hòm thư.")
                                break
                        except Exception:
                            pass

                    subj = getattr(item, 'Subject', '') or ''
                    subj_lower = subj.lower()

                    # Kiểm tra tiêu đề trước để tránh tải nội dung thư không liên quan (tăng tốc x100 lần)
                    is_petrol = (
                        'petrolimex' in subj_lower or 
                        'hóa đơn' in subj_lower or 
                        'hoa don' in subj_lower or
                        'hđđt' in subj_lower
                    )

                    if not is_petrol:
                        sender = getattr(item, 'SenderEmailAddress', '') or ''
                        sender_lower = sender.lower()
                        if 'petrolimex' in sender_lower or 'hoadon' in sender_lower:
                            is_petrol = True

                    if is_petrol:
                        sender = getattr(item, 'SenderEmailAddress', '') or ''
                        body = getattr(item, 'Body', '') or ''
                        html = getattr(item, 'HTMLBody', '') or ''

                        info = parse_email_content(body, html)
                        if info['ma_tra_cuu']:
                            results.append({
                                'subject': subj,
                                'sender': sender,
                                'received_time': str(received_time) if received_time else '',
                                'mau_so': info['mau_so'],
                                'ky_hieu': info['ky_hieu'],
                                'so_hd': info['so_hd'],
                                'ma_tra_cuu': info['ma_tra_cuu'],
                                'portal_link': info['portal_link']
                            })
                            log(f"  [+] Tìm thấy HĐ: {info['ky_hieu']}-{info['so_hd']} | Mã tra cứu: {info['ma_tra_cuu']}")
                except Exception as ex:
                    pass

                try:
                    item = items.GetNext()
                except Exception:
                    break

        log(f"Hoàn thành quét! Tìm thấy {len(results)} email hóa đơn Petrolimex thỏa mãn điều kiện.")
        return results
