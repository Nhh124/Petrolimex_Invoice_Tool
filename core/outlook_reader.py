import os
import re
import sys
import ctypes
from datetime import datetime

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
    # Loại bỏ thẻ html nếu có
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

    def fetch_petrolimex_emails(self, mailbox_name="vantai povina", limit=50):
        """
        Quét các email Petrolimex trong hòm thư chỉ định.
        limit: số lượng email gần nhất để quét (hoặc None để quét toàn bộ)
        """
        inbox = self.get_inbox_folder(mailbox_name)
        if not inbox:
            raise Exception(f"Không tìm thấy hộp thư đến cho: {mailbox_name}")

        items = inbox.Items
        # Sắp xếp thời gian nhận mới nhất lên đầu
        try:
            items.Sort('[ReceivedTime]', True)
        except Exception:
            pass

        results = []
        total_items = items.Count
        scan_limit = total_items if limit is None else min(limit, total_items)

        print(f"Bắt đầu quét {scan_limit}/{total_items} email trong hộp thư '{inbox.Parent.Name}'...")
        
        for i in range(1, scan_limit + 1):
            try:
                item = items.Item(i)
                subj = getattr(item, 'Subject', '') or ''
                sender = getattr(item, 'SenderEmailAddress', '') or ''

                # Kiểm tra có phải email Petrolimex không
                if 'petrolimex' in subj.lower() or 'petrolimex' in sender.lower() or 'hóa đơn' in subj.lower():
                    body = getattr(item, 'Body', '') or ''
                    html = getattr(item, 'HTMLBody', '') or ''
                    received_time = getattr(item, 'ReceivedTime', None)

                    # Trích xuất thông tin
                    info = parse_email_content(body, html)
                    
                    # Chỉ lấy các mail có mã tra cứu
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
            except Exception as e:
                continue

        print(f"Hoàn thành quét! Tìm thấy {len(results)} email hóa đơn Petrolimex.")
        return results
