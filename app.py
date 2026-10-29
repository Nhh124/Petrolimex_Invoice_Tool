import os
import sys
import re
import subprocess
import threading
from flask import Flask, render_template, jsonify, request, send_from_directory, redirect, url_for, session
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from core.manager import PetrolimexInvoiceManager
from core.outlook_reader import OutlookReader
from core.auth import verify_login, DEFAULT_USER

app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, "web", "templates"),
    static_folder=os.path.join(BASE_DIR, "web", "static")
)
app.secret_key = os.urandom(32)

manager = PetrolimexInvoiceManager(BASE_DIR)

# Trạng thái tiến trình đồng bộ
sync_state = {
    'is_running': False,
    'logs': [],
    'last_result': None
}

@app.before_request
def check_auth():
    # Các route công khai không cần đăng nhập
    public_endpoints = ['login', 'static']
    if request.endpoint and any(request.endpoint.startswith(p) for p in public_endpoints):
        return None

    if not session.get('logged_in'):
        if request.path.startswith('/api/'):
            return jsonify({'status': 'unauthorized', 'message': 'Vui lòng đăng nhập'}), 401
        return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if session.get('logged_in'):
        return redirect(url_for('index'))

    error = None
    username = ""
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')

        if verify_login(username, password):
            session['logged_in'] = True
            session['user'] = username
            return redirect(url_for('index'))
        else:
            error = "Tên tài khoản hoặc mật khẩu không chính xác!"

    return render_template('login.html', error=error, username=username)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

def read_excel_data():
    excel_path = manager.excel_path
    if not os.path.exists(excel_path):
        return []

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        if not rows or len(rows) < 2:
            return []

        invoices = []
        for r in rows[1:]:
            if not r or not any(r):
                continue
            stt = r[0] if len(r) > 0 else ""
            ngay_mail = r[1] if len(r) > 1 else ""
            ngay_hd = r[2] if len(r) > 2 else ""
            mau_so = r[3] if len(r) > 3 else ""
            ky_hieu = r[4] if len(r) > 4 else ""
            so_hd = r[5] if len(r) > 5 else ""
            bien_so = r[6] if len(r) > 6 else ""
            ten_hang = r[7] if len(r) > 7 else ""
            dvt = r[8] if len(r) > 8 else "Lít"
            so_luong = float(r[9] or 0.0) if len(r) > 9 else 0.0
            don_gia = float(r[10] or 0.0) if len(r) > 10 else 0.0
            thanh_tien = float(r[11] or 0.0) if len(r) > 11 else 0.0
            tien_thue = float(r[12] or 0.0) if len(r) > 12 else 0.0
            tong_tien = float(r[13] or 0.0) if len(r) > 13 else 0.0
            ben_ban = r[14] if len(r) > 14 else ""
            dia_chi_ban = r[15] if len(r) > 15 else ""
            ben_mua = r[16] if len(r) > 16 else ""
            fkey = r[17] if len(r) > 17 else ""

            # Tìm file PDF thực tế trong thư mục invoices_pdf
            safe_fkey = re.sub(r'[^\w\-]', '_', str(fkey or ''))
            pdf_filename = ""
            if safe_fkey and os.path.exists(manager.pdf_dir):
                target_suffix = f"_{safe_fkey}.pdf"
                for fn in os.listdir(manager.pdf_dir):
                    if fn.endswith(target_suffix):
                        pdf_filename = fn
                        break

            invoices.append({
                'stt': stt,
                'ngay_mail': str(ngay_mail or ''),
                'ngay_hd': str(ngay_hd or ''),
                'mau_so': str(mau_so or ''),
                'ky_hieu': str(ky_hieu or ''),
                'so_hd': str(so_hd or ''),
                'bien_so': str(bien_so or ''),
                'ten_hang': str(ten_hang or ''),
                'dvt': str(dvt or 'Lít'),
                'so_luong': so_luong,
                'don_gia': don_gia,
                'thanh_tien': thanh_tien,
                'tien_thue': tien_thue,
                'tong_tien': tong_tien,
                'ben_ban': str(ben_ban or ''),
                'dia_chi_ban': str(dia_chi_ban or ''),
                'ben_mua': str(ben_mua or ''),
                'fkey': str(fkey or ''),
                'pdf_filename': pdf_filename
            })

        invoices.reverse()
        return invoices
    except Exception as e:
        print(f"Lỗi khi đọc file Excel: {e}")
        return []

@app.route('/')
def index():
    return render_template('index.html', user=session.get('user', DEFAULT_USER))

@app.route('/api/status')
def get_status():
    invoices = read_excel_data()
    total_count = len(invoices)
    total_amount = sum(inv['tong_tien'] for inv in invoices)
    total_volume = sum(inv['so_luong'] for inv in invoices)

    return jsonify({
        'total_count': total_count,
        'total_amount': total_amount,
        'total_volume': total_volume,
        'invoices': invoices,
        'is_syncing': sync_state['is_running']
    })

@app.route('/api/mailboxes')
def get_mailboxes():
    default_boxes = ['vantai povina', 'Phòng kế toán', 'all']
    try:
        import ctypes
        if hasattr(ctypes, 'windll'):
            user32 = ctypes.windll.user32
            hdesk = user32.OpenDesktopW('Default', 0, False, 0x10000000)
            if hdesk:
                user32.SetThreadDesktop(hdesk)
        import win32com.client
        app_obj = win32com.client.GetActiveObject('Outlook.Application')
        ns = app_obj.GetNamespace('MAPI')
        mbs = [f.Name for f in ns.Folders]
        if mbs:
            return jsonify({'status': 'success', 'mailboxes': mbs})
    except Exception:
        pass
    return jsonify({'status': 'success', 'mailboxes': default_boxes})

@app.route('/api/test-imap', methods=['POST'])
def test_imap():
    data = request.json or {}
    email_user = data.get('email_user', '').strip()
    email_pass = data.get('email_pass', '')
    email_server = data.get('email_server', '').strip() or None

    if not email_user or not email_pass:
        return jsonify({'status': 'error', 'message': 'Vui lòng nhập đầy đủ email và mật khẩu!'})

    try:
        from core.imap_reader import IMAPReader
        reader = IMAPReader(username=email_user, password=email_pass, server=email_server)
        ok, msg = reader.test_connection()
        if ok:
            return jsonify({'status': 'success', 'message': f'Kết nối thành công tới máy chủ ({reader.server})!'})
        else:
            return jsonify({'status': 'error', 'message': msg})
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'Lỗi kết nối: {e}'})

@app.route('/api/sync', methods=['POST'])
def start_sync():
    if sync_state['is_running']:
        return jsonify({'status': 'error', 'message': 'Tiến trình đồng bộ đang chạy!'})

    data = request.json or {}
    mailbox = data.get('mailbox', 'vantai povina')
    time_filter = data.get('time_filter', 'all')
    start_date = data.get('start_date', None)
    end_date = data.get('end_date', None)
    limit = int(data.get('limit', 0))
    email_user = data.get('email_user', '').strip()
    email_pass = data.get('email_pass', '')
    email_server = data.get('email_server', '').strip()

    def run_worker():
        import subprocess
        sync_state['is_running'] = True
        sync_state['logs'] = ["Khởi chạy tiến trình quét hóa đơn..."]

        cmd = [
            sys.executable, "-u", os.path.join(BASE_DIR, "cli.py"),
            "--mailbox", str(mailbox),
            "--filter", str(time_filter),
            "--limit", str(limit)
        ]
        if start_date:
            cmd.extend(["--start-date", str(start_date)])
        if end_date:
            cmd.extend(["--end-date", str(end_date)])
        if email_user:
            cmd.extend(["--email-user", str(email_user)])
        if email_pass:
            cmd.extend(["--email-pass", str(email_pass)])
        if email_server:
            cmd.extend(["--email-server", str(email_server)])

        try:
            process = subprocess.Popen(
                cmd,
                cwd=BASE_DIR,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding='utf-8',
                errors='replace',
                bufsize=1
            )
            for line in iter(process.stdout.readline, ''):
                cleaned = line.rstrip()
                if cleaned:
                    sync_state['logs'].append(cleaned)
                    if len(sync_state['logs']) > 500:
                        sync_state['logs'].pop(0)
            process.wait()
            sync_state['last_result'] = {'status': 'success' if process.returncode == 0 else 'error'}
        except Exception as e:
            sync_state['logs'].append(f"Lỗi ngoại lệ: {e}")
        finally:
            sync_state['is_running'] = False

    t = threading.Thread(target=run_worker)
    t.daemon = True
    t.start()

    return jsonify({'status': 'started'})

@app.route('/api/logs')
def get_logs():
    return jsonify({
        'is_running': sync_state['is_running'],
        'logs': sync_state['logs'],
        'last_result': sync_state['last_result']
    })

@app.route('/api/pdf/<path:filename>')
def serve_pdf(filename):
    return send_from_directory(manager.pdf_dir, filename)

@app.route('/api/open-folder', methods=['POST'])
def open_folder():
    try:
        os.startfile(manager.pdf_dir)
        return jsonify({'status': 'success'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/api/open-excel', methods=['POST'])
def open_excel():
    try:
        if os.path.exists(manager.excel_path):
            os.startfile(manager.excel_path)
            return jsonify({'status': 'success'})
        else:
            return jsonify({'status': 'error', 'message': 'File Excel chưa được tạo.'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

if __name__ == '__main__':
    port = 5005
    print(f"Khởi động Web Dashboard tại: http://localhost:{port}")
    app.run(host='0.0.0.0', port=port, debug=False)
