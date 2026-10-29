import imaplib
import email
from email.header import decode_header
import re
import ssl
from datetime import datetime, timedelta

from .outlook_reader import parse_email_content

def decode_mime_words(header_val):
    if not header_val:
        return ""
    decoded_fragments = decode_header(header_val)
    text = ""
    for frag, enc in decoded_fragments:
        if isinstance(frag, bytes):
            text += frag.decode(enc or 'utf-8', errors='replace')
        else:
            text += str(frag)
    return text

def extract_email_body(msg):
    body_text = ""
    html_text = ""
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            cdispo = str(part.get('Content-Disposition') or '')
            if 'attachment' in cdispo:
                continue
            payload = part.get_payload(decode=True)
            if not payload:
                continue
            charset = part.get_content_charset() or 'utf-8'
            try:
                decoded_str = payload.decode(charset, errors='replace')
            except Exception:
                decoded_str = payload.decode('utf-8', errors='replace')
                
            if ctype == 'text/plain' and not body_text:
                body_text = decoded_str
            elif ctype == 'text/html' and not html_text:
                html_text = decoded_str
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or 'utf-8'
            try:
                body_text = payload.decode(charset, errors='replace')
            except Exception:
                body_text = payload.decode('utf-8', errors='replace')
        if msg.get_content_type() == 'text/html':
            html_text = body_text

    return body_text, html_text

class IMAPReader:
    def __init__(self, username, password, server=None, port=993):
        self.username = username.strip() if username else ""
        self.password = password.strip() if password else ""
        
        # Tự động nhận diện máy chủ IMAP dựa trên tên miền email
        if not server:
            if '@povina.com.vn' in self.username.lower():
                self.server = "outlook.office365.com"
            elif '@gmail.com' in self.username.lower():
                self.server = "imap.gmail.com"
            else:
                self.server = "outlook.office365.com"
        else:
            self.server = server.strip()
            
        self.port = int(port or 993)
        self.client = None

    def connect(self):
        ctx = ssl.create_default_context()
        self.client = imaplib.IMAP4_SSL(self.server, self.port, ssl_context=ctx)
        self.client.login(self.username, self.password)
        return True

    def test_connection(self):
        try:
            self.connect()
            self.client.select("INBOX", readonly=True)
            self.client.logout()
            return True, "Kết nối hòm thư máy chủ thành công!"
        except Exception as e:
            return False, f"Lỗi đăng nhập hòm thư: {e}"

    def fetch_petrolimex_emails(self, folder="INBOX", limit=None, time_filter="all", start_date=None, end_date=None, log_cb=None):
        def log(m):
            print(m)
            if log_cb:
                log_cb(m)

        if not self.client:
            log(f"Đang kết nối tới máy chủ IMAP: {self.server}:{self.port}...")
            self.connect()
            log("Đăng nhập tài khoản email thành công!")

        self.client.select(f'"{folder}"', readonly=True)
        
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

        log("Đang tìm kiếm email Petrolimex trên máy chủ...")
        typ, data = self.client.search(None, '(OR SUBJECT "Petrolimex" FROM "Petrolimex")')
        msg_ids = data[0].split()

        # Nếu không thấy theo từ khóa, thử tìm các thư có tiêu đề "Hoa don"
        if not msg_ids:
            typ, data = self.client.search(None, '(SUBJECT "Hoa don")')
            msg_ids = data[0].split()

        log(f"Tổng số email khớp từ khóa trên máy chủ: {len(msg_ids)}")

        # Sắp xếp từ thư mới nhất lùi về trước
        msg_ids.reverse()

        results = []
        for idx, mid in enumerate(msg_ids, 1):
            if limit and limit > 0 and len(results) >= limit:
                log(f"Đã đạt giới hạn tối đa {limit} hóa đơn.")
                break

            try:
                # Chỉ lấy cấu trúc và nội dung thư
                typ, mdata = self.client.fetch(mid, '(BODY.PEEK[])')
                raw_bytes = None
                for response_part in mdata:
                    if isinstance(response_part, tuple):
                        raw_bytes = response_part[1]
                        break

                if not raw_bytes:
                    continue

                msg = email.message_from_bytes(raw_bytes)
                subj = decode_mime_words(msg.get('Subject', ''))
                sender = decode_mime_words(msg.get('From', ''))
                date_str = msg.get('Date', '')

                # Parse date
                item_dt = None
                try:
                    parsed_tuple = email.utils.parsedate_to_datetime(date_str)
                    if parsed_tuple:
                        item_dt = datetime(parsed_tuple.year, parsed_tuple.month, parsed_tuple.day, parsed_tuple.hour, parsed_tuple.minute, parsed_tuple.second)
                except Exception:
                    pass

                # Lọc ngày
                if item_dt:
                    if end_dt and item_dt > end_dt:
                        continue
                    if start_dt and item_dt < start_dt:
                        continue

                body_text, html_text = extract_email_body(msg)
                info = parse_email_content(body_text, html_text)

                if info.get('ma_tra_cuu'):
                    results.append({
                        'subject': subj,
                        'sender': sender,
                        'received_time': str(item_dt or date_str),
                        'mau_so': info['mau_so'],
                        'ky_hieu': info['ky_hieu'],
                        'so_hd': info['so_hd'],
                        'ma_tra_cuu': info['ma_tra_cuu'],
                        'portal_link': info['portal_link']
                    })

                    if len(results) % 20 == 0:
                        log(f"  -> Đã tìm thấy {len(results)} hóa đơn hợp lệ...")
            except Exception as e:
                continue

        log(f"Hoàn thành quét máy chủ! Tìm thấy {len(results)} email hóa đơn Petrolimex.")
        try:
            self.client.logout()
        except Exception:
            pass

        return results
