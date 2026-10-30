import io
import re
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Đối Chiếu Nghiệm Thu HRC - KSNB", page_icon="🛡️", layout="wide"
)


def tim_dong_tieu_de(file_bytes, tu_khoa_list, max_scan=20):
    """Tự động tìm dòng header thực tế trong file Excel, không sợ file bị đổi dòng header."""
    df_raw = pd.read_excel(file_bytes, header=None, nrows=max_scan)
    for idx, row in df_raw.iterrows():
        row_text = " ".join(
            [str(v).upper() for v in row.values if pd.notna(v)]
        )
        if any(k.upper() in row_text for k in tu_khoa_list):
            return idx
    return 0


def xu_ly_doi_chieu(f_bb, f_pl, f_tc, f_bv, f_opm):
    # 1. MỞ TRỰC TIẾP FILE BIÊN BẢN GỐC (GIỮ NGUYÊN 100% FORMAT GỐC)
    wb = load_workbook(f_bb, data_only=False)
    ws = wb.active

    # Đọc bản data_only để lấy giá trị số thực của 13 cột đầu (tránh dính công thức vlookup bị lỗi)
    wb_val = load_workbook(f_bb, data_only=True)
    ws_val = wb_val.active

    # Xác định dòng tiêu đề của BBNT (thường là dòng 8)
    row_header = 8
    for r in range(1, 15):
        if (
            ws.cell(row=r, column=4).value
            and "CUỘN" in str(ws.cell(row=r, column=4).value).upper()
        ):
            row_header = r
            break

    # Lấy danh sách các dòng dữ liệu (loại bỏ dòng tổng cộng, kết luận)
    data_rows = []
    for r in range(row_header + 1, ws.max_row + 1):
        c_val = ws_val.cell(row=r, column=4).value
        stt_val = str(ws_val.cell(row=r, column=1).value).strip()
        if (
            c_val is not None
            and str(c_val).strip() != ""
            and "TỔNG" not in stt_val.upper()
            and "KẾT LUẬN" not in stt_val.upper()
        ):
            cuon_clean = str(c_val).replace(".0", "").strip()
            data_rows.append((r, cuon_clean))

    list_cuon = [x[1] for x in data_rows]

    # 2. SCAN FILE PACKING LIST
    dict_pl = {}
    if f_pl:
        idx_pl = tim_dong_tieu_de(f_pl, ["MÃ SỐ NSX", "MÃ SỐ HSPM", "KL(NET)"])
        f_pl.seek(0)
        df_pl = pd.read_excel(f_pl, skiprows=idx_pl)
        col_nsx = next(
            c
            for c in df_pl.columns
            if "MÃ SỐ NSX" in str(c).upper() or "MÃ SỐ" in str(c).upper()
        )
        col_net = next(c for c in df_pl.columns if "NET" in str(c).upper())

        df_pl_clean = df_pl.dropna(subset=[col_nsx, col_net])
        keys = (
            df_pl_clean[col_nsx]
            .astype(str)
            .str.replace(r"\.0$", "", regex=True)
            .str.strip()
            + "A"
        )
        nets = pd.to_numeric(df_pl_clean[col_net], errors="coerce").fillna(0)
        dict_pl = (
            pd.DataFrame({"Key": keys, "Net": nets})
            .groupby("Key")["Net"]
            .sum()
            .to_dict()
        )

    # 3. SCAN FILE DỮ LIỆU BẢO VỆ (Quét mã cuộn trong text)
    dict_bv_phieu = {}
    dict_bv_kl = {}
    if f_bv:
        f_bv.seek(0)
        df_bv = pd.read_excel(f_bv)
        col_hang = next(
            (
                c
                for c in df_bv.columns
                if "MÃ HÀNG HÓA" in str(c).upper()
                or "TÊN HÀNG HÓA" in str(c).upper()
            ),
            None,
        )
        col_pcan = next(
            (
                c
                for c in df_bv.columns
                if "MÃ PHIẾU CÂN" in str(c).upper()
                or "PHIẾU CÂN" in str(c).upper()
            ),
            None,
        )
        col_klcan = next(
            (c for c in df_bv.columns if "KHỐI LƯỢNG" in str(c).upper()), None
        )

        for _, r in df_bv.iterrows():
            txt = f"{r.get('MÃ HÀNG HÓA', '')} {r.get('TÊN HÀNG HÓA', '')}"
            p_can = str(r.get(col_pcan, "")).replace(".0", "").strip()
            kl_can = r.get(col_klcan, 0)
            for m in list_cuon:
                if m in txt and m not in dict_bv_phieu:
                    dict_bv_phieu[m] = p_can
                    dict_bv_kl[m] = kl_can

    # 4. SCAN FILE TRẠM CÂN
    dict_tc_kl = {}
    dict_tc_pl = {}
    if f_tc:
        f_tc.seek(0)
        df_tc = pd.read_excel(f_tc)
        col_p = next(
            c for c in df_tc.columns if str(c).strip().upper() == "PHIẾU CÂN"
        )
        col_k = next(c for c in df_tc.columns if "KHỐI LƯỢNG" in str(c).upper())
        col_pl = next(
            (c for c in df_tc.columns if "PHÂN LOẠI" in str(c).upper()), None
        )

        p_series = (
            df_tc[col_p]
            .astype(str)
            .str.replace(r"\.0$", "", regex=True)
            .str.strip()
        )
        kl_series = pd.to_numeric(df_tc[col_k], errors="coerce").fillna(0)
        dict_tc_kl = dict(zip(p_series, kl_series))
        if col_pl:
            dict_tc_pl = dict(
                zip(p_series, df_tc[col_pl].fillna("Nhập hàng").astype(str))
            )

    # 5. SCAN FILE ERP OPM 105 (Bắt đúng cột Khối lượng và Mã Kho)
    opm_035_po, opm_035_rtv, opm_031, opm_03g_po, opm_03g_rtv = (
        {},
        {},
        {},
        {},
        {},
    )
    if f_opm:
        idx_opm = tim_dong_tieu_de(f_opm, ["LOẠI GIAO DỊCH", "MÃ LÔ"])
        f_opm.seek(0)
        df_opm = pd.read_excel(f_opm, skiprows=idx_opm)
        if len(df_opm) > 0 and "[" in str(df_opm.iloc[0, 1]):
            df_opm = df_opm.iloc[1:].copy()

        c_loai = next(
            c for c in df_opm.columns if "LOẠI GIAO DỊCH" in str(c).upper()
        )
        c_malo = next(c for c in df_opm.columns if "MÃ LÔ" in str(c).upper())
        # Cột khối lượng thực chất là cột tên 'Khối Lượng' (cột số 9)
        c_kl = next(
            c
            for c in df_opm.columns
            if "KHỐI LƯỢNG" in str(c).upper() and "GROSS" not in str(c).upper()
        )
        # Cột kho thực chất là cột 'Mă Kho' (cột 34)
        c_kho = next(
            c
            for c in df_opm.columns
            if "MĂ KHO" in str(c).upper() or str(c).strip().upper() == "MÃ KHO"
        )

        df_opm_clean = df_opm.dropna(subset=[c_malo, c_loai]).copy()
        raw_lo = df_opm_clean[c_malo].astype(str).str.strip().str.upper()
        df_opm_clean["Key"] = raw_lo.apply(
            lambda x: x if x.endswith("A") else x + "A"
        )
        df_opm_clean["KL_Val"] = pd.to_numeric(
            df_opm_clean[c_kl], errors="coerce"
        ).fillna(0)
        df_opm_clean["Kho_Val"] = (
            df_opm_clean[c_kho].astype(str).str.strip().str.upper()
        )

        opm_035_po = (
            df_opm_clean[
                (df_opm_clean["Kho_Val"] == "035")
                & (
                    df_opm_clean[c_loai].str.contains(
                        "PO Receipt", case=False, na=False
                    )
                )
            ]
            .groupby("Key")["KL_Val"]
            .sum()
            .to_dict()
        )
        opm_035_rtv = (
            df_opm_clean[
                (df_opm_clean["Kho_Val"] == "035")
                & (
                    df_opm_clean[c_loai].str.contains(
                        "Return to Vendor", case=False, na=False
                    )
                )
            ]
            .groupby("Key")["KL_Val"]
            .sum()
            .to_dict()
        )
        opm_031 = (
            df_opm_clean[
                (df_opm_clean["Kho_Val"].isin(["03G", "03H"]))
                & (
                    df_opm_clean[c_loai].str.contains(
                        "Intransit Receipt", case=False, na=False
                    )
                )
            ]
            .groupby("Key")["KL_Val"]
            .sum()
            .to_dict()
        )
        opm_03g_po = (
            df_opm_clean[
                (df_opm_clean["Kho_Val"] == "03G")
                & (
                    df_opm_clean[c_loai].str.contains(
                        "PO Receipt", case=False, na=False
                    )
                )
            ]
            .groupby("Key")["KL_Val"]
            .sum()
            .to_dict()
        )
        opm_03g_rtv = (
            df_opm_clean[
                (df_opm_clean["Kho_Val"] == "03G")
                & (
                    df_opm_clean[c_loai].str.contains(
                        "Return to Vendor", case=False, na=False
                    )
                )
            ]
            .groupby("Key")["KL_Val"]
            .sum()
            .to_dict()
        )

    # 6. GHI ĐÈ TIÊU ĐỀ CÁC CỘT TỪ N ĐẾN AH (CỘT 14 ĐẾN 34)
    headers = [
        (14, "NHẬP/XUẤT"),
        (15, "Mã hóa"),
        (16, "Khối lượng net theo PL"),
        (17, "Đối chiếu "),
        (18, "KL theo bc nhập hàng "),
        (19, "Số phiếu cân "),
        (20, "Đối chiếu số phiếu cân"),
        (21, "KL theo bc trạm cân"),
        (22, ""),
        (23, "Đối chiếu KL TC & KL BCBV"),
        (24, "Đối chiếu KL TC & KL BBNT"),
        (25, "Chênh lệch (kg)"),
        (26, " Chênh lệch dung sai"),
        (27, "Đối chiếu dung sai"),
        (28, " So với quy định ( >+-0.3%)"),
        (29, "KL Nhập kho\n035"),
        (30, "Chênh lệch"),
        (31, "KL Nhập kho\n031"),
        (32, "Chênh lệch"),
        (33, "KL Nhập kho\n03G-3H"),
        (34, "Chênh lệch"),
    ]

    yellow_fill = PatternFill(
        start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid"
    )
    red_fill = PatternFill(
        start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid"
    )
    red_font = Font(color="9C0006", bold=True)
    border_thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    for col_idx, h_text in headers:
        cell = ws.cell(row=row_header, column=col_idx)
        cell.value = h_text
        cell.fill = yellow_fill
        cell.font = Font(bold=True)
        cell.alignment = Alignment(
            wrap_text=True, horizontal="center", vertical="center"
        )

    # 7. SCAN TỪNG DÒNG VÀ ĐIỀN GIÁ TRỊ VÀO PHÔI BBNT
    for r, m in data_rows:
        m_hoa = m + "A"
        p_can_bbnt = str(ws_val.cell(row=r, column=5).value or "").strip()
        net_ct = float(ws_val.cell(row=r, column=8).value or 0)
        gross_ct = float(ws_val.cell(row=r, column=9).value or 1)
        can_ban = float(ws_val.cell(row=r, column=10).value or 0)
        dsai_bbnt = round(float(ws_val.cell(row=r, column=12).value or 0), 6)

        # Tính toán giá trị từ các nguồn đã scan
        nhap_xuat = dict_tc_pl.get(p_can_bbnt, "Nhập hàng")
        kl_net_pl = float(dict_pl.get(m_hoa, 0))
        doi_chieu_pl = net_ct - kl_net_pl
        kl_bv = float(dict_bv_kl.get(m, 0))
        p_can_bv = str(dict_bv_phieu.get(m, "")).strip()
        chk_pcan = p_can_bbnt == p_can_bv if p_can_bv else False

        # Khối lượng trạm cân tra theo số phiếu bảo vệ (nếu ko có lấy phiếu bbnt)
        p_tra_tc = p_can_bv if p_can_bv else p_can_bbnt
        kl_tc = float(dict_tc_kl.get(p_tra_tc, 0))

        chk_tc_bv = kl_tc == kl_bv if kl_bv > 0 else False
        chk_tc_bbnt = kl_tc == can_ban if can_ban > 0 else False
        chenh_lech_kg = kl_tc - gross_ct
        chenh_lech_ds = chenh_lech_kg / gross_ct if gross_ct != 0 else 0
        chk_dsai = round(chenh_lech_ds, 6) == dsai_bbnt
        so_quy_dinh = (
            "VUOT DUNG SAI" if abs(chenh_lech_ds) > 0.003 else "OK"
        )

        # OPM 105
        kl_035 = float(
            opm_035_po.get(m_hoa, 0) - abs(opm_035_rtv.get(m_hoa, 0))
        )
        chenh_035 = kl_net_pl - kl_035
        kl_031 = float(opm_031.get(m_hoa, 0))
        chenh_031 = 0 if kl_031 == 0 else kl_net_pl - kl_031
        kl_03g = float(
            opm_03g_po.get(m_hoa, 0) - abs(opm_03g_rtv.get(m_hoa, 0))
        )
        chenh_03g = kl_net_pl - kl_03g

        # Ghi vào từng ô
        row_values = {
            14: nhap_xuat,
            15: m_hoa,
            16: kl_net_pl,
            17: doi_chieu_pl,
            18: kl_bv,
            19: p_can_bv,
            20: chk_pcan,
            21: kl_tc,
            22: p_can_bv + "A" if p_can_bv else "",
            23: chk_tc_bv,
            24: chk_tc_bbnt,
            25: chenh_lech_kg,
            26: chenh_lech_ds,
            27: chk_dsai,
            28: so_quy_dinh,
            29: kl_035,
            30: chenh_035,
            31: kl_031,
            32: chenh_031,
            33: kl_03g,
            34: chenh_03g,
        }

        for col_idx, val in row_values.items():
            c = ws.cell(row=r, column=col_idx)
            c.value = val
            c.border = border_thin

            # Tô đỏ nếu có lỗi vi phạm
            is_err = False
            if col_idx in [20, 23, 24, 27] and val is False:
                is_err = True
            elif col_idx == 28 and val == "VUOT DUNG SAI":
                is_err = True
            elif (
                col_idx in [17, 30]
                and isinstance(val, (int, float))
                and val != 0
            ):
                is_err = True

            if is_err:
                c.fill = red_fill
                c.font = red_font

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
