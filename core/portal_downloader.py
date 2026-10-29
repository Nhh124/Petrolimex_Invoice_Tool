import os
import re
import sys
import ssl
import json
import urllib.request
import urllib.parse
import http.cookiejar
import ddddocr
from playwright.sync_api import sync_playwright

class PortalDownloader:
    def __init__(self, output_pdf_dir, output_xml_dir):
        self.output_pdf_dir = output_pdf_dir
        self.output_xml_dir = output_xml_dir
        os.makedirs(self.output_pdf_dir, exist_ok=True)
        os.makedirs(self.output_xml_dir, exist_ok=True)

        self.ctx = ssl.create_default_context()
        self.ctx.check_hostname = False
        self.ctx.verify_mode = ssl.CERT_NONE

        # Initialize OCR once
        self.ocr = ddddocr.DdddOcr(show_ad=False)

    def _create_session(self):
        cj = http.cookiejar.CookieJar()
        opener = urllib.request.build_opener(
            urllib.request.HTTPSHandler(context=self.ctx),
            urllib.request.HTTPCookieProcessor(cj)
        )
        return opener

    def download_invoice(self, fkey, mau_so="", ky_hieu="", so_hd="", max_retries=4):
        """
        Tra cứu mã fkey trên portal Petrolimex, tự động giải captcha và tải XML + PDF.
        Trả về dictionary chứa kết quả và đường dẫn file đã tải.
        """
        clean_fkey = fkey.strip()
        print(f"\n[PortalDownloader] Bắt đầu tra cứu hóa đơn: Mã Fkey = {clean_fkey}")

        for attempt in range(1, max_retries + 1):
            try:
                opener = self._create_session()

                # 1. Tải trang chủ để lấy Token
                req_home = urllib.request.Request(
                    'https://hoadon.petrolimex.com.vn',
                    headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
                )
                with opener.open(req_home, timeout=15) as resp:
                    home_html = resp.read().decode('utf-8', errors='ignore')

                m_token = re.search(r'name="__RequestVerificationToken"\s+type="hidden"\s+value="([^"]+)"', home_html)
                if not m_token:
                    print(f"  [Lần {attempt}] Không tìm thấy Token xác thực, thử lại...")
                    continue
                token = m_token.group(1)

                # 2. Tải và nhận diện Captcha
                req_captcha = urllib.request.Request(
                    'https://hoadon.petrolimex.com.vn/Captcha/Show',
                    headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
                )
                with opener.open(req_captcha, timeout=10) as resp:
                    captcha_bytes = resp.read()

                captcha_text = self.ocr.classification(captcha_bytes)
                print(f"  [Lần {attempt}] Nhận diện Captcha: '{captcha_text}'")

                # 3. Gửi form tìm kiếm
                form_data = {
                    '__RequestVerificationToken': token,
                    'tab': 'content1',
                    'strFkey': clean_fkey,
                    'captch': captcha_text,
                    'submit': 'Tìm kiếm'
                }
                encoded_data = urllib.parse.urlencode(form_data).encode('utf-8')
                req_search = urllib.request.Request(
                    'https://hoadon.petrolimex.com.vn/SearchInvoicebycode/Index',
                    data=encoded_data,
                    headers={
                        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
                        'Content-Type': 'application/x-www-form-urlencoded',
                        'Referer': 'https://hoadon.petrolimex.com.vn'
                    }
                )

                with opener.open(req_search, timeout=20) as resp:
                    search_html = resp.read().decode('utf-8', errors='ignore')

                # Kiểm tra lỗi Captcha sai
                if "Mã xác thực không đúng" in search_html or ("Captcha" in search_html and "sai" in search_html):
                    print(f"  [Lần {attempt}] Captcha không khớp, đang thử lại...")
                    continue

                # 4. Tìm checkCode trong kết quả tìm kiếm
                m_xml = re.search(r'/SearchInvoicebycode/DownloadXml/\?checkCode=([^"\'\s>]+)', search_html)
                if not m_xml:
                    # Kiểm tra xem có phải mã không tồn tại không
                    if "Không tìm thấy dữ liệu" in search_html or "không tồn tại" in search_html:
                        print(f"  [Thông báo] Mã tra cứu {clean_fkey} không tìm thấy trên hệ thống!")
                        return None
                    print(f"  [Lần {attempt}] Không thấy liên kết tải về, đang thử lại...")
                    continue

                check_code = m_xml.group(1)
                print(f"  -> Tra cứu thành công! checkCode = {check_code[:25]}...")

                # Chuẩn bị tên file
                safe_kh = re.sub(r'[^\w\-]', '_', ky_hieu) or "KH"
                safe_shd = re.sub(r'[^\w\-]', '_', so_hd) or "SHD"
                safe_fkey = re.sub(r'[^\w\-]', '_', clean_fkey)
                
                filename_base = f"HD_{safe_kh}_{safe_shd}_{safe_fkey}"
                xml_path = os.path.join(self.output_xml_dir, f"{filename_base}.xml")
                pdf_path = os.path.join(self.output_pdf_dir, f"{filename_base}.pdf")

                # 5. Tải file XML
                xml_url = f"https://hoadon.petrolimex.com.vn/SearchInvoicebycode/DownloadXml/?checkCode={urllib.parse.quote(check_code)}"
                with opener.open(xml_url, timeout=15) as resp_xml:
                    xml_data = resp_xml.read()
                    with open(xml_path, "wb") as f_out:
                        f_out.write(xml_data)
                print(f"  -> Đã lưu file XML ({len(xml_data)} bytes): {xml_path}")

                # 6. Tải preview HTML để kết xuất thành PDF
                preview_url = "https://hoadon.petrolimex.com.vn/SearchInvoicebycode/ajxPreview/"
                post_data = urllib.parse.urlencode({'checkCode': check_code}).encode('utf-8')
                req_prev = urllib.request.Request(
                    preview_url,
                    data=post_data,
                    headers={'User-Agent': 'Mozilla/5.0', 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'}
                )
                with opener.open(req_prev, timeout=20) as resp_prev:
                    prev_json_str = resp_prev.read().decode('utf-8', errors='ignore')
                    prev_json = json.loads(prev_json_str)
                    html_content = prev_json.get('str', '')

                # Render HTML sang file PDF bằng Playwright
                if html_content:
                    try:
                        with sync_playwright() as p:
                            browser = p.chromium.launch()
                            page = browser.new_page()
                            page.set_content(html_content, wait_until="networkidle")
                            page.pdf(
                                path=pdf_path,
                                format="A4",
                                print_background=True,
                                margin={"top": "10mm", "bottom": "10mm", "left": "10mm", "right": "10mm"}
                            )
                            browser.close()
                        print(f"  -> Đã xuất file PDF hóa đơn: {pdf_path}")
                    except Exception as pe:
                        print(f"  [Cảnh báo] Lỗi kết xuất PDF: {pe}")
                        pdf_path = None

                return {
                    'status': 'success',
                    'xml_path': xml_path,
                    'pdf_path': pdf_path,
                    'check_code': check_code
                }

            except Exception as e:
                print(f"  [Lần {attempt}] Gặp lỗi khi xử lý: {e}")
                continue

        print(f"  [Thất bại] Đã thử {max_retries} lần nhưng không thành công cho mã {clean_fkey}")
        return None
