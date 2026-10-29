import os
import re
import sys
import ctypes
from datetime import datetime, timedelta

def _ensure_desktop():
    try:
        user32 = ctypes.windll.user32
        hdesk = user32.OpenDesktopW('Default', 0, False, 0x10000000)
        if hdesk:
            user32.SetThreadDesktop(hdesk)
    except Exception:
        pass

def parse_email_content(body_text, html_text=""):
    """
    Trích xuất các thông tin hóa đơn từ nội dung email Petrolimex
    """
    clean_body = body_text.replace('\xa0', ' ')
    
    # Tìm Mẫu số
    m_mau = re.search(r'Mẫu\s*số\s*:\s*([^\r\n<]+)', clean_body, re.IGNORECASE)
    mau_so = m_mau.group(1).strip() if m_mau else ""
    mau_so = re.sub(r'<[^>]+>', '', mau_so).strip()

    # Tìm Ký hiệu
    m_kh = re.search(r'Ký\s*hiệu\s*:\s*([^\r\n<]+)', clean_body, re.IGNORECASE)
    ky_hieu = m_kh.group(1).strip() if m_kh else ""
    ky_hieu = re.sub(r'<[^>]+>', '', ky_hieu).strip()

    # Tìm Số hóa đơn
    m_shd = re.search(r'Số\s*hóa\s*đơn\s*:\s*([^\r\n<]+)', clean_body, re.IGNORECASE)
    so_hd = m_shd.group(1).strip() if m_shd else ""
    so_hd = re.sub(r'<[^>]+>', '', so_hd).strip()

    # Tìm Mã tra cứu (Fkey)
    m_mtc = re.search(r'Mã\s*tra\s*cứu\s*:\s*([^\r\n<]+)', clean_body, re.IGNORECASE)
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
        self.app = win32com.client.Dispatch('Outlook.Application')
        self.ns = self.app.GetNamespace('MAPI')

    def get_mailboxes(self):
        """Liệt kê các tài khoản/hòm thư trong Outlook"""
        mailboxes = []
        for f in self.ns.Folders:
            mailboxes.append(f.Name)
        return mailboxes

    def get_inbox_folder(self, mailbox_name=None):
        """Lấy thư mục Hộp thư đến (Inbox)"""
        if mailbox_name:
            for f in self.ns.Folders:
                if mailbox_name.lower() in f.Name.lower():
                    for sub in f.Folders:
                        if any(k in sub.Name.lower() for k in ['inbox', 'hộp thư đến', 'đến']):
                            return sub
        # Fallback to default
        try:
            return self.ns.GetDefaultFolder(6)
        except Exception:
            return None

    def fetch_petrolimex_emails(self, mailbox_name="vantai povina", limit=None, time_filter="all", start_date=None, end_date=None):
        """
        Quét các email Petrolimex trong hòm thư với bộ lọc linh hoạt:
        - time_filter: 'all' (toàn bộ), 'today' (hôm nay), 'week' (7 ngày qua), 'month' (tháng này), 'last_month' (tháng trước), 'custom' (tùy chọn ngày)
        - start_date, end_date: Định dạng YYYY-MM-DD khi time_filter='custom'
        - limit: Giới hạn số lượng email (None hoặc 0 = không giới hạn)
        """
        inbox = self.get_inbox_folder(mailbox_name)
        if not inbox:
            raise Exception(f"Không tìm thấy hộp thư đến cho: {mailbox_name}")

        items = inbox.Items
        try:
            items.Sort('[ReceivedTime]', True)
        except Exception:
            pass

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

        print(f"Áp dụng bộ lọc thời gian: {time_filter} (Từ: {start_dt or 'Không giới hạn'} Đến: {end_dt or 'Hiện tại'})")

        results = []
        total_items = items.Count
        max_limit = total_items if (limit is None or limit <= 0) else min(limit, total_items)

        print(f"Bắt đầu quét hòm thư '{inbox.Parent.Name}' ({total_items} thư)...")

        for i in range(1, total_items + 1):
            if limit and limit > 0 and len(results) >= limit:
                print(f"Đã đạt giới hạn tối đa {limit} hóa đơn.")
                break

            try:
                item = items.Item(i)
                received_time = getattr(item, 'ReceivedTime', None)

                # Kiểm tra lọc theo thời gian
                if received_time:
                    try:
                        item_dt = datetime(
                            received_time.year, received_time.month, received_time.day,
                            received_time.hour, received_time.minute, received_time.second
                        )
                        # Nếu thư mới hơn end_dt thì bỏ qua sang thư tiếp theo
                        if end_dt and item_dt > end_dt:
                            continue
                        # Vì đã sắp xếp mới nhất lên đầu, nếu thư cũ hơn start_dt thì dừng lại
                        if start_dt and item_dt < start_dt:
                            print(f"Đã duyệt đến mốc ngày {item_dt.strftime('%d/%m/%Y')}, dừng quét.")
                            break
                    except Exception:
                        pass

                subj = getattr(item, 'Subject', '') or ''
                sender = getattr(item, 'SenderEmailAddress', '') or ''

                if 'petrolimex' in subj.lower() or 'petrolimex' in sender.lower() or 'hóa đơn' in subj.lower():
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
            except Exception:
                continue

        print(f"Hoàn thành quét! Tìm thấy {len(results)} email hóa đơn Petrolimex thỏa mãn điều kiện.")
        return results
