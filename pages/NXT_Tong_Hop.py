import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="Báo Cáo NXT NVL Chính", page_icon="📊", layout="wide")

# ==========================================
# CSS ĐỊNH DẠNG GIAO DIỆN CHUẨN FORM EXCEL
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
    .header-logo-title {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 15px;
        margin-bottom: 10px;
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
    .excel-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 14px;
        color: #000000;
    }
    .excel-table th {
        background-color: #cce5ff;
        border: 1px solid #000000;
        padding: 6px 8px;
        text-align: center;
        font-weight: bold;
    }
    .excel-table td {
        border: 1px solid #000000;
        padding: 5px 8px;
    }
    .text-center { text-align: center; }
    .text-right { text-align: right; }
    .text-left { text-align: left; }
    .row-total {
        background-color: #ffff00;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 BÁO CÁO TỔNG HỢP NXT NVL CHÍNH")

# ==========================================
# 1. HÀM TÍNH TOÁN DỮ LIỆU TỒN KHO (INV 110)
# ==========================================
def extract_date_inv(file):
    try:
        df_head = pd.read_excel(file, header=None, nrows=10)
        for val in df_head.values.flatten():
            if pd.notna(val) and "Ngày tồn kho:" in str(val):
                return str(val).split("Ngày tồn kho:")[-1].strip()
    except:
        pass
    return "09/10/2026"

def process_inv110(files):
    df_list = []
    report_date = "09/10/2026"
    for file in files:
        cur_date = extract_date_inv(file)
        if cur_date:
            report_date = cur_date
        df = pd.read_excel(file, skiprows=10, header=None)
        df_list.append(df)
        
    if not df_list:
        return None, None, report_date
        
    df_raw = pd.concat(df_list, ignore_index=True)
    
    # Xác định các cột chính
    # Col 1: Vị trí kho, Col 2: Mã hàng, Col 3: Tên hàng, Col 23: Net weight (KG)
    col_kho = df_raw[1].astype(str).str.strip()
    col_ma = df_raw[2].astype(str).str.strip()
    col_ten = df_raw[3].astype(str).str.strip()
    col_kg = pd.to_numeric(df_raw[23], errors='coerce').fillna(0)

    # Phân loại tự động theo công thức chuẩn Excel
    phan_loai = pd.Series("", index=df_raw.index)
    
    is_not_A = ~col_ma.str.startswith("A")
    is_cm = is_not_A & col_ten.str.contains("Thép cán nguội", case=False, na=False)
    is_po = is_not_A & col_ten.str.contains("Thép cán nóng tẩy rỉ", case=False, na=False)
    is_hrc = is_not_A & (~is_po) & col_ten.str.contains("Thép cán nóng", case=False, na=False)
    is_nhom_kem = col_ten.str.contains("Hợp kim Nhôm - Kẽm|nhôm kẽm", case=False, na=False)
    is_kem = col_ten.str.contains("Kẽm không hợp kim", case=False, na=False)
    is_magie = col_ten.str.contains("Hợp kim Kẽm Nhôm Magie|kẽm nhôm magie", case=False, na=False)
    is_pha_che = col_ten.str.contains("Antimon|pha chế", case=False, na=False)
    
    phan_loai[is_cm] = "CM"
    phan_loai[is_po] = "PO"
    phan_loai[is_hrc] = "HRC"
    phan_loai[is_nhom_kem] = "Hợp kim nhôm kẽm"
    phan_loai[is_kem] = "Kẽm không hợp kim"
    phan_loai[is_magie] = "Hợp kim Kẽm Nhôm Magie"
    phan_loai[is_pha_che] = "Nguyên liệu pha chế"
    
    # Bảng 1: Tồn kho Thép
    mask_hrc_cang = (phan_loai == "HRC") & (col_kho.str.contains("035"))
    mask_hrc_kho = (phan_loai == "HRC") & (~col_kho.str.contains("035"))
    mask_po = (phan_loai == "PO")
    mask_cm = (phan_loai == "CM")
    
    tb1 = [
        {"item": "Tồn kho HRC", "qty": int(mask_hrc_kho.sum()), "weight": col_kg[mask_hrc_kho].sum() / 1000.0, "note": ""},
        {"item": "Tồn kho HRC ̣̣̣( hàng ngoài cảng)", "qty": int(mask_hrc_cang.sum()), "weight": col_kg[mask_hrc_cang].sum() / 1000.0, "note": ""},
        {"item": "Tồn kho PO", "qty": int(mask_po.sum()), "weight": col_kg[mask_po].sum() / 1000.0, "note": ""},
        {"item": "Tồn kho CM + CR", "qty": int(mask_cm.sum()), "weight": col_kg[mask_cm].sum() / 1000.0, "note": ""}
    ]
    
    # Bảng 2: Tồn kho Kim loại màu
    tb2 = [
        {"item": "Tồn kho CGG", "weight": col_kg[phan_loai == "Kẽm không hợp kim"].sum() / 1000.0, "note": ""},
        {"item": "Tồn kho Hợp kim nhôm kẽm", "weight": col_kg[phan_loai == "Hợp kim nhôm kẽm"].sum() / 1000.0, "note": ""},
        {"item": "Tồn kho Hợp kim kẽm Magie", "weight": col_kg[phan_loai == "Hợp kim Kẽm Nhôm Magie"].sum() / 1000.0, "note": ""},
        {"item": "Tồn kho nguyên liệu pha chế", "weight": col_kg[phan_loai == "Nguyên liệu pha chế"].sum() / 1000.0, "note": ""}
    ]
    
    return tb1, tb2, report_date

# ==========================================
# 2. HÀM TÍNH TOÁN DỮ LIỆU SỬ DỤNG (105)
# ==========================================
def process_105(files):
    df_list = []
    for file in files:
        df = pd.read_excel(file, skiprows=8, header=None)
        df_list.append(df)
        
    if not df_list:
        return None, None
        
    df_raw = pd.concat(df_list, ignore_index=True)
    
    # Cột trong 105: 2: Loại GD, 4: Dây Chuyền, 5: Mã Hàng, 6: Tên Hàng, 11: Khối lượng (KG)
    col_gd = df_raw[2].astype(str).str.strip()
    col_line = df_raw[4].astype(str).str.lower()
    col_ma = df_raw[5].astype(str).str.strip()
    col_ten = df_raw[6].astype(str).str.strip()
    col_kg = pd.to_numeric(df_raw[11], errors='coerce').fillna(0)

    # Đánh số nhóm giao dịch (Cột B trong Excel)
    cond_ignore = col_line.str.contains("xử lý|gia công|tái sinh", regex=True, na=False)
    col_b = pd.Series("", index=df_raw.index)
    
    col_b[(~cond_ignore) & (col_gd.isin
