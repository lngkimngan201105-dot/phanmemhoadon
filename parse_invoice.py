import os
import glob
import xml.etree.ElementTree as ET
import pandas as pd

def get_node_text(element, tag_name, default=""):
    if element is None:
        return default
    node = element.find(f".//{tag_name}")
    return node.text.strip() if (node is not None and node.text) else default

def parse_invoice_xml(file_path):
    """
    Trả về 2 danh sách:
    - summary_row: 1 dòng tóm tắt toàn bộ hóa đơn (Sheet Tổng Hợp)
    - detail_rows: danh sách từng dòng hàng hóa (Sheet Chi Tiết)
    """
    try:
        tree = ET.parse(file_path)
        root = tree.getroot()
    except Exception as e:
        print(f"Lỗi đọc {file_path}: {e}")
        return None, []

    file_name = os.path.basename(file_path)
    
    # 1. Thông tin chung
    kh_mau = get_node_text(root, "KHMSHDon")
    kh_hd = get_node_text(root, "KHHDon")
    so_hd = get_node_text(root, "SHDon")
    ngay_lap = get_node_text(root, "NLap")

    # Người bán & Người mua
    nban = root.find(".//NBan")
    ten_nban = get_node_text(nban, "Ten")
    mst_nban = get_node_text(nban, "MST")
    dchi_nban = get_node_text(nban, "DChi")

    nmua = root.find(".//NMua")
    ten_nmua = get_node_text(nmua, "Ten")
    mst_nmua = get_node_text(nmua, "MST")

    # Tổng tiền toàn hóa đơn (thường nằm trong thẻ TToan)
    ttoan = root.find(".//TToan")
    def to_float(val):
        try:
            return float(val) if val else 0.0
        except ValueError:
            return 0.0

    tong_tien_chua_thue = to_float(get_node_text(ttoan, "TgTCThue"))
    tong_tien_thue = to_float(get_node_text(ttoan, "TgTThue"))
    tong_thanh_toan = to_float(get_node_text(ttoan, "TgTTTBSo"))

    summary_row = {
        "File Nguồn": file_name,
        "Mẫu Số": kh_mau,
        "Ký Hiệu": kh_hd,
        "Số HĐ": so_hd,
        "Ngày Lập": ngay_lap,
        "MST Người Bán": mst_nban,
        "Tên Người Bán": ten_nban,
        "Địa Chỉ Người Bán": dchi_nban,
        "MST Người Mua": mst_nmua,
        "Tên Người Mua": ten_nmua,
        "Tiền Chưa Thuế": tong_tien_chua_thue,
        "Tiền Thuế GTGT": tong_tien_thue,
        "Tổng Thanh Toán": tong_thanh_toan
    }

    # 2. Chi tiết từng mặt hàng
    detail_rows = []
    hh_nodes = root.findall(".//DSHHang/HHDVu")

    for idx, item in enumerate(hh_nodes, start=1):
        sl = to_float(get_node_text(item, "SLuong"))
        dg = to_float(get_node_text(item, "DGia"))
        th_tien = to_float(get_node_text(item, "ThTien"))
        st_ckhau = to_float(get_node_text(item, "STCKhau"))

        detail_rows.append({
            "File Nguồn": file_name,
            "Ký Hiệu": kh_hd,
            "Số HĐ": so_hd,
            "Ngày Lập": ngay_lap,
            "MST Người Bán": mst_nban,
            "Tên Người Bán": ten_nban,
            "STT": get_node_text(item, "STT") or str(idx),
            "Tên Hàng Hóa / Dịch Vụ": get_node_text(item, "THHDon"),
            "ĐVT": get_node_text(item, "DVTinh"),
            "Số Lượng": sl,
            "Đơn Giá": dg,
            "Tiền Chiết Khấu": st_ckhau,
            "Thành Tiền Chưa Thuế": th_tien,
            "Thuế Suất": get_node_text(item, "TSuat")
        })

    return summary_row, detail_rows

def export_to_excel(summary_data, detail_data, output_excel_path):
    """Ghi dữ liệu vào 1 file Excel gồm 2 Sheet riêng biệt"""
    df_summary = pd.DataFrame(summary_data)
    df_detail = pd.DataFrame(detail_data)

    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        if not df_summary.empty:
            df_summary.to_excel(writer, sheet_name="Tong_Hop_Hoa_Don", index=False)
        if not df_detail.empty:
            df_detail.to_excel(writer, sheet_name="Chi_Tiet_Hang_Hoa", index=False)