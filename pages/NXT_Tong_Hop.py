import io
import re
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

st.set_page_config(page_title="Báo Cáo NXT NVL Chính", page_icon="📊", layout="wide")

# ==========================================
# GIAO DIỆN
# ==========================================
st.markdown("""
<style>
    .report-card {
        background-color: #ffffff;
        border: 2px solid #000000;
        padding: 20px;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        margin-bottom: 25px;
        font-family: 'Times New Roman', Times, serif;
    }
    .company-title {
        font-size: 20px;
        font-weight: bold;
        color: #000000;
        text-align: center;
        margin: 0;
    }
    .report-title {
        font-size: 22px;
        font-weight: 800;
        text-align: center;
        color: #000000;
        margin-top: 5px;
        margin-bottom: 4px;
    }
    .report-date {
        font-size: 14px;
        font-style: italic;
        color: #d90429;
        text-align: center;
        margin-bottom: 15px;
        font-weight: bold;
    }
    .section-title {
        font-size: 15px;
        font-weight: bold;
        color: #d90429;
        margin-top: 15px;
        margin-bottom: 8px;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 BÁO CÁO TỔNG HỢP NXT NVL CHÍNH")
st.caption("Tải hai nguồn NLC-NPL và 105. Ứng dụng chỉ tạo sheet NXT Tong Hop, không tạo sheet Report.")

# ==========================================
# 1. HÀM DÙNG CHUNG ĐỂ ĐỌC FILE
# ==========================================
def normalize_text(value):
    """Chuyển text về chữ thường, bỏ dấu và khoảng trắng thừa để dò tên cột."""
    if value is None or (not isinstance(value, (list, tuple, dict)) and pd.isna(value)):
        return ""
    text = str(value).replace("Đ", "D").replace("đ", "d")
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", text).strip().casefold()


def read_raw_excel(uploaded_file):
    """Đọc sheet đầu tiên, không tự lấy dòng tiêu đề để tránh lệch cột trong báo cáo ERP."""
    uploaded_file.seek(0)
    raw = pd.read_excel(uploaded_file, sheet_name=0, header=None, dtype=object, engine="openpyxl")
    uploaded_file.seek(0)
    return raw


def find_header_row(raw, required_tokens, scan_rows=80):
    required = [normalize_text(x) for x in required_tokens]
    for row_idx in range(min(scan_rows, len(raw))):
        vals = [normalize_text(v) for v in raw.iloc[row_idx].tolist()]
        if all(any(token in val for val in vals) for token in required):
            return row_idx
    return None


def find_column(row_values, token, fallback=None):
    token = normalize_text(token)
    for idx, value in enumerate(row_values):
        if token and token in normalize_text(value):
            return idx
    return fallback


def to_numeric_series(series):
    """Đổi cột khối lượng sang số; ô trống hoặc giá trị không phải số được tính bằng 0."""
    return pd.to_numeric(series, errors="coerce").fillna(0.0).astype(float)


def extract_inventory_date(raw):
    """Lấy ngày từ ô có dạng 'Ngày tồn kho: dd/mm/yyyy'."""
    for value in raw.iloc[:15, :].to_numpy().ravel():
        text = str(value) if value is not None else ""
        if "ngay ton kho" in normalize_text(text):
            match = re.search(r"(\d{1,2}[/-]\d{1,2}[/-]\d{4})", text)
            if match:
                return match.group(1).replace("-", "/")
    return "Không xác định"


# ==========================================
# 2. PHÂN LOẠI NGUYÊN LIỆU: NLC-NPL
#    Bám theo logic công thức cột A trong sheet NLC-NPL gốc.
# ==========================================
def classify_stock_item(item_name, item_code):
    name = normalize_text(item_name)
    code = normalize_text(item_code).replace(" ", "")
    is_code_a = code.startswith("a")

    # Cần xét PO trước HRC vì tên PO cũng có cụm "thép cán nóng".
    if "thep can nguoi" in name and not is_code_a:
        return "CM"
    if "thep can nong tay ri" in name and not is_code_a:
        return "PO"
    if "thep can nong" in name and not is_code_a:
        return "HRC"

    if "hop kim nhom" in name and "kem" in name:
        return "Hợp kim nhôm kẽm"
    if "kem khong hop kim" in name:
        return "Kẽm không hợp kim"
    if "hop kim kem nhom magie" in name or "hop kim kem nhom magi" in name:
        return "Hợp kim Kẽm Nhôm Magie"
    if "antimon" in name:
        return "Nguyên liệu pha chế"

    # Ví dụ "Hợp kim kẽm (dạng xỉ)" không khớp các nhóm báo cáo ở trên nên bị loại.
    return ""


def load_nlc_npl(uploaded_files):
    frames = []
    dates = []

    for file in uploaded_files:
        raw = read_raw_excel(file)
        header_idx = find_header_row(raw, ["Vị trí kho", "Mã hàng", "Tên hàng"])
        if header_idx is None:
            raise ValueError(
                f"Không tìm thấy dòng tiêu đề Vị trí kho / Mã hàng / Tên hàng trong file '{file.name}'."
            )

        header_values = raw.iloc[header_idx].tolist()
        location_col = find_column(header_values, "Vị trí kho", fallback=1)
        code_col = find_column(header_values, "Mã hàng", fallback=2)
        name_col = find_column(header_values, "Tên hàng", fallback=3)

        # Báo cáo NLC-NPL có dòng tiêu đề phụ Gross/Net ngay dưới tiêu đề chính.
        subheader_idx = header_idx + 1
        subheader_net_col = None
        if subheader_idx < len(raw):
            subheader_net_col = find_column(raw.iloc[subheader_idx].tolist(), "Net")

        if subheader_net_col is not None:
            # Có dòng tiêu đề phụ Gross/Net nên dữ liệu bắt đầu sau hai dòng tiêu đề.
            net_col = subheader_net_col
            data_start = header_idx + 2
        else:
            # Một số bản xuất có thể chỉ có một dòng tiêu đề.
            net_col = find_column(header_values, "Net", fallback=23)
            data_start = header_idx + 1
        data = raw.iloc[data_start:].copy()

        # Phòng trường hợp báo cáo có thêm các dòng trắng cuối file.
        max_col = max(location_col, code_col, name_col, net_col)
        if data.shape[1] <= max_col:
            raise ValueError(f"Thiếu cột dữ liệu cần thiết trong file '{file.name}'.")

        out = pd.DataFrame({
            "location": data.iloc[:, location_col].fillna("").astype(str).str.strip(),
            "item_code": data.iloc[:, code_col].fillna("").astype(str).str.strip(),
            "item_name": data.iloc[:, name_col].fillna("").astype(str).str.strip(),
            "net_kg": to_numeric_series(data.iloc[:, net_col]),
        })
        out["category"] = [
            classify_stock_item(name, code)
            for name, code in zip(out["item_name"], out["item_code"])
        ]
        frames.append(out)

        report_date = extract_inventory_date(raw)
        if report_date != "Không xác định":
            dates.append(report_date)

    if not frames:
        raise ValueError("Bạn chưa tải file NLC-NPL.")

    report_date = dates[0] if dates else "Không xác định"
    if len(set(dates)) > 1:
        st.warning("Các file NLC-NPL có ngày tồn kho khác nhau. Báo cáo đang hiển thị ngày của file đầu tiên.")
    return pd.concat(frames, ignore_index=True), report_date


# ==========================================
# 3. PHÂN LOẠI GIAO DỊCH 105
#    Bám theo logic hai cột phụ A và B trong sheet 105 gốc.
# ==========================================
def classify_transaction_item(item_name, item_code, transaction_group):
    if not transaction_group:
        return ""

    name = normalize_text(item_name)
    code = normalize_text(item_code).replace(" ", "")

    # Theo công thức Excel: mã mặt hàng bắt đầu bằng A được đưa vào nhóm A.
    if code.startswith("a"):
        return "A"
    if "thep can nguoi" in name:
        return "CM"
    if "thep can nong tay ri" in name:
        return "PO"
    if "thep can nong" in name:
        return "HRC"
    if "hop kim nhom" in name and "kem" in name:
        return "Hợp kim nhôm kẽm"
    if "kem khong hop kim" in name:
        return "Kẽm không hợp kim"
    if "hop kim kem nhom magie" in name or "hop kim kem nhom magi" in name:
        return "Hợp kim Kẽm Nhôm Magie"
    if "antimon" in name:
        return "Nguyên liệu pha chế"
    return ""


def get_transaction_group(transaction, production_line):
    trans = normalize_text(transaction)
    line = normalize_text(production_line)

    # Đúng theo công thức cột B của file gốc: bỏ qua xử lý, gia công, tái sinh.
    if any(word in line for word in ("xu ly", "gia cong", "tai sinh")):
        return ""

    if trans in {"wip issue", "wip return"}:
        return "1"
    if trans in {
        "wip completion",
        "wip completion return",
        "wip byproduct completion",
        "wip byproduct return",
    }:
        return "2"
    if trans == "move order issue":
        return "3"
    if trans == "miscellaneous receipt":
        return "4"
    return ""


def load_105(uploaded_files):
    frames = []

    for file in uploaded_files:
        raw = read_raw_excel(file)
        header_idx = find_header_row(
            raw,
            ["Loại Giao Dịch", "Dây Chuyền", "Mã Mặt Hàng", "Tên Mặt Hàng", "Khối Lượng"],
        )
        if header_idx is None:
            raise ValueError(
                f"Không tìm thấy dòng tiêu đề Loại Giao Dịch / Dây Chuyền / Mã Mặt Hàng / Tên Mặt Hàng / Khối Lượng trong file '{file.name}'."
            )

        header_values = raw.iloc[header_idx].tolist()
        transaction_col = find_column(header_values, "Loại Giao Dịch", fallback=2)
        line_col = find_column(header_values, "Dây Chuyền", fallback=4)
        code_col = find_column(header_values, "Mã Mặt Hàng", fallback=5)
        name_col = find_column(header_values, "Tên Mặt Hàng", fallback=6)
        weight_col = find_column(header_values, "Khối Lượng", fallback=11)

        # Sheet 105 có dòng "Đánh số giao dịch" / [1], [2]... giữa tiêu đề và dữ liệu.
        data_start = header_idx + 1
        if header_idx + 1 < len(raw):
            second_header = [normalize_text(v) for v in raw.iloc[header_idx + 1].tolist()]
            if any("danh so giao dich" in v for v in second_header) or any(re.search(r"\[\d+\]", v) for v in second_header):
                data_start = header_idx + 2

        data = raw.iloc[data_start:].copy()
        max_col = max(transaction_col, line_col, code_col, name_col, weight_col)
        if data.shape[1] <= max_col:
            raise ValueError(f"Thiếu cột dữ liệu cần thiết trong file '{file.name}'.")

        trans = data.iloc[:, transaction_col].fillna("").astype(str).str.strip()
        line = data.iloc[:, line_col].fillna("").astype(str).str.strip()
        code = data.iloc[:, code_col].fillna("").astype(str).str.strip()
        name = data.iloc[:, name_col].fillna("").astype(str).str.strip()
        weight = to_numeric_series(data.iloc[:, weight_col])

        group = [get_transaction_group(t, l) for t, l in zip(trans, line)]
        category = [
            classify_transaction_item(n, c, g)
            for n, c, g in zip(name, code, group)
        ]

        out = pd.DataFrame({
            "transaction": trans.map(normalize_text),
            "production_line": line,
            "item_code": code,
            "item_name": name,
            "group": group,
            "category": category,
            "weight_kg": weight,
        })
        # Dòng trống / dòng tiêu đề lọt vào không ảnh hưởng phép SUMIFS, nhưng bỏ đi để gọn dữ liệu.
        out = out[out["transaction"].ne("")].copy()
        frames.append(out)

    if not frames:
        raise ValueError("Bạn chưa tải file 105.")
    return pd.concat(frames, ignore_index=True)


# ==========================================
# 4. TÍNH KẾT QUẢ NXT TỔNG HỢP
#    Công thức được đối chiếu theo sheet NXT Tong Hop của file mẫu.
# ==========================================
def calculate_nxt(stock_df, trans_df, report_date):
    location_norm = stock_df["location"].map(normalize_text)
    stock_category = stock_df["category"]
    stock_kg = stock_df["net_kg"]
    is_port_035 = location_norm.str.endswith("035")

    def inventory_count(category, port=None):
        mask = stock_category.eq(category)
        if port is True:
            mask &= is_port_035
        elif port is False:
            mask &= ~is_port_035
        return int(mask.sum())

    def inventory_tons(category, port=None):
        mask = stock_category.eq(category)
        if port is True:
            mask &= is_port_035
        elif port is False:
            mask &= ~is_port_035
        return float(stock_kg[mask].sum() / 1000.0)

    # Bảng 1: Tồn kho HRC, HRC tại cảng (vị trí kết thúc bằng 035), PO, CM.
    stock_steel = [
        {"item": "Tồn kho HRC", "qty": inventory_count("HRC", port=False), "weight": inventory_tons("HRC", port=False), "note": ""},
        {"item": "Tồn kho HRC (hàng ngoài cảng)", "qty": inventory_count("HRC", port=True), "weight": inventory_tons("HRC", port=True), "note": ""},
        {"item": "Tồn kho PO", "qty": inventory_count("PO"), "weight": inventory_tons("PO"), "note": ""},
        {"item": "Tồn kho CM + CR", "qty": inventory_count("CM"), "weight": inventory_tons("CM"), "note": ""},
    ]
    stock_steel_total = {
        "item": "Tổng",
        "qty": sum(x["qty"] for x in stock_steel),
        "weight": sum(x["weight"] for x in stock_steel),
        "note": "",
    }

    # Bảng 2: Tồn kho kim loại màu. Công thức gốc dùng cột Net (X), đổi kg sang tấn.
    stock_nonferrous_map = [
        ("Tồn kho CGG", "Kẽm không hợp kim"),
        ("Tồn kho Hợp kim nhôm kẽm", "Hợp kim nhôm kẽm"),
        ("Tồn kho Hợp kim kẽm Magie", "Hợp kim Kẽm Nhôm Magie"),
        ("Tồn kho nguyên liệu pha chế", "Nguyên liệu pha chế"),
    ]
    stock_nonferrous = [
        {"item": label, "weight": inventory_tons(category), "note": ""}
        for label, category in stock_nonferrous_map
    ]

    def sum_105_kg(category, group):
        mask = trans_df["category"].eq(category) & trans_df["group"].eq(group)
        return float(trans_df.loc[mask, "weight_kg"].sum())

    def count_105(category, transaction):
        mask = (
            trans_df["category"].eq(category)
            & trans_df["group"].eq("1")
            & trans_df["transaction"].eq(transaction)
        )
        return int(mask.sum())

    # Bảng 3: NVL chính. Số lượng = số WIP Issue trừ số WIP Return.
    primary_items = [
        ("HRC", "HRC"),
        ("PO", "PO"),
        ("CM", "CM"),
    ]
    usage_primary = []
    for label, category in primary_items:
        qty = count_105(category, "wip issue") - count_105(category, "wip return")
        weight_tons = abs(sum_105_kg(category, "1") / 1000.0)
        usage_primary.append({"item": label, "qty": qty, "weight": weight_tons, "note": ""})

    # Theo công thức gốc: tổng nhóm A, nhóm giao dịch 2; không lấy ABS.
    scrap_tons = sum_105_kg("A", "2") / 1000.0
    usage_primary.append({"item": "Hao hụt+ phế liệu", "qty": None, "weight": scrap_tons, "note": ""})
    usage_primary_total = {
        "item": "Tổng",
        "qty": sum(x["qty"] or 0 for x in usage_primary),
        "weight": sum(x["weight"] for x in usage_primary),
        "note": "",
    }

    # Bảng 4: NPL. Đúng công thức ABS(tổng Move Order Issue - tổng Miscellaneous receipt) / 1000.
    npl_usage_map = [
        ("Hợp kim nhôm kẽm", "Hợp kim nhôm kẽm"),
        ("Kẽm không hợp kim", "Kẽm không hợp kim"),
        ("Hợp kim kẽm Magie", "Hợp kim Kẽm Nhôm Magie"),
        ("Nguyên liệu pha chế", "Nguyên liệu pha chế"),
    ]
    usage_npl = []
    for label, category in npl_usage_map:
        used_kg = sum_105_kg(category, "3")
        received_kg = sum_105_kg(category, "4")
        usage_npl.append({"item": label, "weight": abs(used_kg - received_kg) / 1000.0, "note": ""})

    return {
        "report_date": report_date,
        "stock_steel": stock_steel,
        "stock_steel_total": stock_steel_total,
        "stock_nonferrous": stock_nonferrous,
        "usage_primary": usage_primary,
        "usage_primary_total": usage_primary_total,
        "usage_npl": usage_npl,
    }


# ==========================================
# 5. XUẤT FILE EXCEL CHỈ CÓ SHEET NXT TONG HOP
# ==========================================
def create_nxt_excel(result):
    wb = Workbook()
    ws = wb.active
    ws.title = "NXT Tong Hop"
    ws.sheet_view.showGridLines = False

    # Màu sắc / viền gần với biểu mẫu NXT Tong Hop.
    blue_fill = PatternFill("solid", fgColor="CCE5FF")
    yellow_fill = PatternFill("solid", fgColor="FFFF00")
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")

    ws["C1"] = "NHÀ MÁY TÔN HOA SEN PHÚ MỸ"
    ws["M1"] = "NHÀ MÁY TÔN HOA SEN PHÚ MỸ"
    for cell in (ws["C1"], ws["M1"]):
        cell.font = Font(name="Times New Roman", size=12, bold=True)
        cell.alignment = center

    ws["A3"] = "BÁO CÁO KHỐI LƯỢNG NLC, NPL TẠI CTC/NM"
    ws["K3"] = "BÁO CÁO KHỐI LƯỢNG SỬ DỤNG NLC, NPL TẠI CTC/NM"
    ws["A4"] = f"Đến 23 giờ 59 phút ngày {result['report_date']}"
    ws["K4"] = f"Đến 23 giờ 59 phút ngày {result['report_date']}"
    for cell in (ws["A3"], ws["K3"]):
        cell.font = Font(name="Times New Roman", size=12, bold=True)
        cell.alignment = center
    for cell in (ws["A4"], ws["K4"]):
        cell.font = Font(name="Times New Roman", size=10, italic=True, color="D90429", bold=True)
        cell.alignment = center

    ws["A6"] = "1. Tồn kho HRC-PO-CR-CM"
    ws["K6"] = "3. Tình hình sử dụng NVL chính :"
    ws["A16"] = "2. Kim loại màu tại nhà máy NH"
    ws["K16"] = "4. Tình hình sử dụng NPL :"
    for cell in (ws["A6"], ws["K6"], ws["A16"], ws["K16"]):
        cell.font = Font(name="Times New Roman", size=11, bold=True, color="D90429")

    # Tiêu đề bảng tồn kho NLC.
    for cell, value in {
        "B8": "Tên hàng", "D8": "Số lượng", "E8": "Khối lượng", "F8": "Ghi chú",
        "D9": "Cuộn", "E9": "Tấn",
    }.items():
        ws[cell] = value
    # Tiêu đề bảng sử dụng NLC.
    for cell, value in {
        "M8": "Tên hàng", "O8": "Số lượng", "P8": "Khối lượng", "Q8": "Ghi chú",
        "O9": "Cuộn", "P9": "Tấn",
    }.items():
        ws[cell] = value

    for row_num, row in enumerate(result["stock_steel"], start=10):
        ws.cell(row_num, 2, row["item"])
        ws.cell(row_num, 4, row["qty"])
        ws.cell(row_num, 5, row["weight"])
        ws.cell(row_num, 6, row["note"])
    ws["B14"] = result["stock_steel_total"]["item"]
    ws["D14"] = result["stock_steel_total"]["qty"]
    ws["E14"] = result["stock_steel_total"]["weight"]

    for row_num, row in enumerate(result["usage_primary"], start=10):
        ws.cell(row_num, 13, row["item"])
        if row["qty"] is not None:
            ws.cell(row_num, 15, row["qty"])
        ws.cell(row_num, 16, row["weight"])
        ws.cell(row_num, 17, row["note"])
    ws["M14"] = result["usage_primary_total"]["item"]
    ws["O14"] = result["usage_primary_total"]["qty"]
    ws["P14"] = result["usage_primary_total"]["weight"]

    # Tiêu đề bảng tồn kho NPL / kim loại màu.
    for cell, value in {
        "B18": "Tên hàng", "D18": "Khối lượng", "F18": "Ghi chú", "D19": "Tấn",
        "M18": "Tên hàng", "O18": "Khối lượng", "P18": "Ghi chú", "O19": "Tấn",
    }.items():
        ws[cell] = value

    for row_num, row in enumerate(result["stock_nonferrous"], start=20):
        ws.cell(row_num, 2, row["item"])
        ws.cell(row_num, 4, row["weight"])
        ws.cell(row_num, 6, row["note"])
    for row_num, row in enumerate(result["usage_npl"], start=20):
        ws.cell(row_num, 13, row["item"])
        ws.cell(row_num, 15, row["weight"])
        ws.cell(row_num, 16, row["note"])

    # Định dạng bảng.
    for row in (8, 9, 18, 19):
        for col in range(2, 7):
            cell = ws.cell(row, col)
            if cell.value is not None:
                cell.fill = blue_fill
                cell.font = Font(name="Times New Roman", size=10, bold=True)
                cell.alignment = center
                cell.border = border
        for col in range(13, 18):
            cell = ws.cell(row, col)
            if cell.value is not None:
                cell.fill = blue_fill
                cell.font = Font(name="Times New Roman", size=10, bold=True)
                cell.alignment = center
                cell.border = border

    for row in list(range(10, 15)) + list(range(20, 24)):
        for col in list(range(2, 7)) + list(range(13, 18)):
            cell = ws.cell(row, col)
            if cell.value is not None:
                cell.border = border
                cell.font = Font(name="Times New Roman", size=10, bold=(row == 14))
                cell.alignment = left if col in (2, 6, 13, 17) else right

    for col in range(2, 7):
        ws.cell(14, col).fill = yellow_fill
    for col in range(13, 18):
        ws.cell(14, col).fill = yellow_fill

    for row in range(10, 15):
        ws.cell(row, 4).number_format = '#,##0;[Red](#,##0);-'
        ws.cell(row, 5).number_format = '#,##0.000;[Red](#,##0.000);-'
        ws.cell(row, 15).number_format = '#,##0;[Red](#,##0);-'
        ws.cell(row, 16).number_format = '#,##0.000;[Red](#,##0.000);-'
    for row in range(20, 24):
        ws.cell(row, 4).number_format = '#,##0.000;[Red](#,##0.000);-'
        ws.cell(row, 15).number_format = '#,##0.000;[Red](#,##0.000);-'

    widths = {
        "A": 5, "B": 34, "C": 3, "D": 14, "E": 16, "F": 18,
        "G": 3, "H": 3, "I": 3, "J": 3, "K": 5, "L": 3,
        "M": 32, "N": 3, "O": 14, "P": 16, "Q": 18,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for row_num in [1, 3, 4, 6, 8, 9, 14, 16, 18, 19]:
        ws.row_dimensions[row_num].height = 22 if row_num in (3, 6, 16) else 18
    ws.freeze_panes = "B10"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ==========================================
# 6. GIAO DIỆN TẢI FILE VÀ TẠO NXT
# ==========================================
col1, col2 = st.columns(2)
with col1:
    nlc_files = st.file_uploader(
        "① Tải file NLC-NPL (tồn kho theo lô/cuộn)",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        key="upload_nlc_npl",
    )
with col2:
    files_105 = st.file_uploader(
        "② Tải file 105 (giao dịch sử dụng nguyên liệu)",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        key="upload_105",
    )

if st.button("🔄 TÍNH NXT TỔNG HỢP", type="primary", use_container_width=True):
    if not nlc_files or not files_105:
        st.error("Bạn cần tải đủ cả hai nguồn: NLC-NPL và 105.")
        st.stop()

    try:
        with st.spinner("Đang đọc và đối chiếu dữ liệu..."):
            stock_df, report_date = load_nlc_npl(nlc_files)
            trans_df = load_105(files_105)
            result = calculate_nxt(stock_df, trans_df, report_date)
            xlsx_bytes = create_nxt_excel(result)

            st.session_state["nxt_result"] = result
            st.session_state["nxt_xlsx_bytes"] = xlsx_bytes
            st.session_state["nxt_source_date"] = report_date
            st.session_state["nxt_data_counts"] = {
                "stock_rows": len(stock_df),
                "transaction_rows": len(trans_df),
            }
    except Exception as exc:
        st.error(f"Không thể tạo báo cáo: {exc}")
        st.stop()

if st.session_state.get("nxt_result"):
    result = st.session_state["nxt_result"]
    counts = st.session_state.get("nxt_data_counts", {})
    st.success(f"Đã tính xong NXT Tổng hợp — ngày tồn kho: {result['report_date']}.")
    st.caption(
        f"Đã đọc {counts.get('stock_rows', 0):,} dòng NLC-NPL và "
        f"{counts.get('transaction_rows', 0):,} dòng giao dịch 105."
    )

    st.subheader("1. Tồn kho NLC-NPL")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**HRC / PO / CM**")
        st.dataframe(
            pd.DataFrame(result["stock_steel"] + [result["stock_steel_total"]]),
            use_container_width=True,
            hide_index=True,
        )
    with c2:
        st.markdown("**Kim loại màu / nguyên liệu pha chế**")
        st.dataframe(pd.DataFrame(result["stock_nonferrous"]), use_container_width=True, hide_index=True)

    st.subheader("2. Tình hình sử dụng nguyên liệu từ file 105")
    c3, c4 = st.columns(2)
    with c3:
        st.markdown("**Nguyên liệu chính**")
        st.dataframe(
            pd.DataFrame(result["usage_primary"] + [result["usage_primary_total"]]),
            use_container_width=True,
            hide_index=True,
        )
    with c4:
        st.markdown("**Nguyên liệu phụ / kim loại màu**")
        st.dataframe(pd.DataFrame(result["usage_npl"]), use_container_width=True, hide_index=True)

    st.download_button(
        "⬇️ TẢI FILE EXCEL — NXT TONG HOP",
        data=st.session_state["nxt_xlsx_bytes"],
        file_name=f"NXT_Tong_Hop_{result['report_date'].replace('/', '-')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
