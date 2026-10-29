import os
import json
import re
import openpyxl

def export_excel_to_json(excel_path, json_path, pdf_dir=None, xml_dir=None):
    """
    Chuyển đổi dữ liệu từ file Excel sang file JSON phục vụ cho Web tĩnh GitHub Pages.
    """
    if not os.path.exists(excel_path):
        return []

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not rows or len(rows) < 2:
        return []

    # Map existing files in pdf_dir
    pdf_files = os.listdir(pdf_dir) if pdf_dir and os.path.exists(pdf_dir) else []
    xml_files = os.listdir(xml_dir) if xml_dir and os.path.exists(xml_dir) else []

    def find_file(files, prefix_folder, ky_hieu, so_hd, fkey):
        clean_fkey = re.sub(r'[^\w\-]', '_', fkey).strip('_')
        clean_shd = re.sub(r'[^\w\-]', '', so_hd)
        for f in files:
            if clean_fkey and clean_fkey in f:
                return f"{prefix_folder}/{f}"
            if clean_shd and clean_shd in f:
                return f"{prefix_folder}/{f}"
        return ""

    invoices = []
    for r in rows[1:]:
        if not r or not any(r):
            continue
        stt = r[0] if len(r) > 0 else ""
        ngay_mail = str(r[1] or '').strip()
        ngay_hd = str(r[2] or '').strip()
        mau_so = str(r[3] or '').strip()
        ky_hieu = str(r[4] or '').strip()
        so_hd = str(r[5] or '').strip()
        bien_so = str(r[6] or '').strip()
        ten_hang = str(r[7] or '').strip()
        dvt = str(r[8] or 'Lít').strip()
        so_luong = float(r[9] or 0.0) if len(r) > 9 else 0.0
        don_gia = float(r[10] or 0.0) if len(r) > 10 else 0.0
        thanh_tien = float(r[11] or 0.0) if len(r) > 11 else 0.0
        tien_thue = float(r[12] or 0.0) if len(r) > 12 else 0.0
        tong_tien = float(r[13] or 0.0) if len(r) > 13 else 0.0
        ben_ban = str(r[14] or '').strip()
        dia_chi_ban = str(r[15] or '').strip()
        ben_mua = str(r[16] or '').strip()
        fkey = str(r[17] or '').strip()

        pdf_rel = find_file(pdf_files, "invoices_pdf", ky_hieu, so_hd, fkey)
        xml_rel = find_file(xml_files, "invoices_xml", ky_hieu, so_hd, fkey)

        invoices.append({
            'stt': stt,
            'ngay_mail': ngay_mail,
            'ngay_hd': ngay_hd,
            'mau_so': mau_so,
            'ky_hieu': ky_hieu,
            'so_hd': so_hd,
            'bien_so': bien_so,
            'ten_hang': ten_hang,
            'dvt': dvt,
            'so_luong': so_luong,
            'don_gia': don_gia,
            'thanh_tien': thanh_tien,
            'tien_thue': tien_thue,
            'tong_tien': tong_tien,
            'ben_ban': ben_ban,
            'dia_chi_ban': dia_chi_ban,
            'ben_mua': ben_mua,
            'fkey': fkey,
            'pdf_url': pdf_rel,
            'xml_url': xml_rel
        })

    # Sort newest first
    invoices.reverse()

    os.makedirs(os.path.dirname(json_path), exist_ok=True)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(invoices, f, ensure_ascii=False, indent=2)

    print(f"[ExportJSON] Exported {len(invoices)} invoices to {json_path}")
    return invoices
