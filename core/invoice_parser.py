import xml.etree.ElementTree as ET
import os
import re

def _get_text(element, tag_name, default=""):
    if element is None:
        return default
    child = element.find(tag_name)
    if child is not None and child.text:
        return child.text.strip()
    return default

def _get_ttin_value(parent_element, ttruong_name):
    """Lấy giá trị DLieu trong danh sách TTin / TTKhac"""
    if parent_element is None:
        return ""
    ttkhac = parent_element.find('TTKhac')
    if ttkhac is None:
        return ""
    for ttin in ttkhac.findall('TTin'):
        ttruong = _get_text(ttin, 'TTruong')
        if ttruong.lower() == ttruong_name.lower():
            return _get_text(ttin, 'DLieu')
    return ""

def parse_invoice_xml(xml_path):
    """
    Đọc và trích xuất toàn bộ dữ liệu có cấu trúc từ file XML hóa đơn điện tử Petrolimex (chuẩn TT78).
    """
    if not os.path.exists(xml_path):
        return None

    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
    except Exception as e:
        print(f"Lỗi khi đọc file XML {xml_path}: {e}")
        return None

    # Tìm nhánh DLHDon
    dlhdon = root.find('.//DLHDon')
    if dlhdon is None:
        dlhdon = root

    ttchung = dlhdon.find('TTChung')
    ndhdon = dlhdon.find('NDHDon')
    if ndhdon is None:
        ndhdon = dlhdon

    nban = ndhdon.find('NBan')
    nmua = ndhdon.find('NMua')
    dshhdvu = ndhdon.find('DSHHDVu')
    ttoan = ndhdon.find('TToan')

    # 1. Thông tin chung hóa đơn
    mau_so = _get_text(ttchung, 'KHMSHDon')
    ky_hieu = _get_text(ttchung, 'KHHDon')
    so_hd = _get_text(ttchung, 'SHDon')
    ngay_lap = _get_text(ttchung, 'NLap')
    hinh_thuc_tt = _get_text(ttchung, 'HTTToan')
    loai_tien = _get_text(ttchung, 'DVTTe', 'VND')

    # 2. Bên bán
    ten_ben_ban = _get_text(nban, 'Ten')
    mst_ben_ban = _get_text(nban, 'MST')
    dia_chi_ban = _get_text(nban, 'DChi')
    dien_thoai_ban = _get_text(nban, 'SDThoai')
    bien_so_xe = _get_ttin_value(nban, 'SoPhuongTien')
    ma_ddkd = _get_ttin_value(nban, 'MDDKDoanh')
    dia_chi_ddkd = _get_ttin_value(nban, 'DCDDKDoanh')

    # 3. Bên mua
    ten_ben_mua = _get_text(nmua, 'Ten')
    mst_ben_mua = _get_text(nmua, 'MST')
    dia_chi_mua = _get_text(nmua, 'DChi')
    email_mua = _get_text(nmua, 'DCTDTu')

    # 4. Hàng hóa dịch vụ
    danh_sach_hang = []
    if dshhdvu is not None:
        for item in dshhdvu.findall('HHDVu'):
            ten_hh = _get_text(item, 'THHDVu')
            ma_hh = _get_text(item, 'MHHDVu')
            dvt = _get_text(item, 'DVTinh')
            
            try:
                sl = float(_get_text(item, 'SLuong', '0'))
            except ValueError:
                sl = 0.0
                
            try:
                dg = float(_get_text(item, 'DGia', '0'))
            except ValueError:
                dg = 0.0
                
            try:
                tt = float(_get_text(item, 'ThTien', '0'))
            except ValueError:
                tt = 0.0

            tsuat = _get_text(item, 'TSuat')
            
            danh_sach_hang.append({
                'ten_hang': ten_hh,
                'ma_hang': ma_hh,
                'dvt': dvt,
                'so_luong': sl,
                'don_gia': dg,
                'thanh_tien': tt,
                'thue_suat': tsuat
            })

    # Hàng hóa đại diện chính (hầu hết là 1 dòng xăng/dầu)
    hang_chinh = danh_sach_hang[0] if danh_sach_hang else {
        'ten_hang': '', 'ma_hang': '', 'dvt': 'Lít', 'so_luong': 0.0, 'don_gia': 0.0, 'thanh_tien': 0.0, 'thue_suat': ''
    }

    # 5. Tổng thanh toán
    try:
        tien_chua_thue = float(_get_text(ttoan, 'TgTCThue', '0'))
    except ValueError:
        tien_chua_thue = hang_chinh['thanh_tien']

    try:
        tien_thue = float(_get_text(ttoan, 'TgTThue', '0'))
    except ValueError:
        tien_thue = 0.0

    try:
        tong_tien_tt = float(_get_text(ttoan, 'TgTTTBSo', '0'))
    except ValueError:
        tong_tien_tt = tien_chua_thue + tien_thue

    tong_tien_chu = _get_text(ttoan, 'TgTTTBChu')

    # 6. Mã tra cứu (Fkey)
    fkey = _get_text(root, 'Fkey')
    if not fkey:
        fkey = _get_ttin_value(dlhdon, 'Fkey')

    # 7. Thời gian ký số
    signing_time = ""
    for elem in root.iter():
        if 'SigningTime' in elem.tag and elem.text:
            signing_time = elem.text.strip()
            break

    return {
        'mau_so': mau_so,
        'ky_hieu': ky_hieu,
        'so_hd': so_hd,
        'ngay_lap': ngay_lap,
        'hinh_thuc_tt': hinh_thuc_tt,
        'loai_tien': loai_tien,
        'ten_ben_ban': ten_ben_ban,
        'mst_ben_ban': mst_ben_ban,
        'dia_chi_ban': dia_chi_ban,
        'bien_so_xe': bien_so_xe,
        'dia_chi_ddkd': dia_chi_ddkd,
        'ma_ddkd': ma_ddkd,
        'ten_ben_mua': ten_ben_mua,
        'mst_ben_mua': mst_ben_mua,
        'dia_chi_mua': dia_chi_mua,
        'email_mua': email_mua,
        'ten_hang': hang_chinh['ten_hang'],
        'dvt': hang_chinh['dvt'],
        'so_luong': hang_chinh['so_luong'],
        'don_gia': hang_chinh['don_gia'],
        'thanh_tien': hang_chinh['thanh_tien'],
        'thue_suat': hang_chinh['thue_suat'],
        'tien_chua_thue': tien_chua_thue,
        'tien_thue': tien_thue,
        'tong_tien_tt': tong_tien_tt,
        'tong_tien_chu': tong_tien_chu,
        'ma_tra_cuu': fkey,
        'signing_time': signing_time,
        'danh_sach_hang': danh_sach_hang
    }
