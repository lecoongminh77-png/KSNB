import io
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import plotly.express as px
import streamlit as st

# -------------------------------------------------------------
# CẤU HÌNH GIAO DIỆN CHUẨN APPLE & THANH TIỆN ÍCH
# -------------------------------------------------------------
st.set_page_config(page_title='Tiện Ích - KSNB', page_icon='🛡️', layout='wide')

# CSS ẩn menu sidebar và tạo phong cách Apple
st.markdown(
    """
<style>
    [data-testid="stSidebarNav"] { display: none !important; }
    .stApp {
        background-color: #f5f5f7;
        font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", sans-serif;
        color: #1d1d1f;
    }
    .apple-header-box {
        display: flex;
        justify-content: space-between;
        align-items: flex-end;
        padding: 30px 0 15px 0;
        margin-bottom: 10px;
    }
    .apple-main-title {
        font-size: 46px;
        font-weight: 700;
        color: #1d1d1f;
        letter-spacing: -0.025em;
        margin: 0;
    }
    .apple-sub-text {
        font-size: 19px;
        font-weight: 600;
        color: #1d1d1f;
        text-align: right;
        margin: 0;
        line-height: 1.35;
    }
    .apple-sub-text span {
        color: #0071e3;
        font-size: 15px;
        display: block;
        margin-top: 4px;
        font-weight: 500;
    }
    div.stButton > button:first-child {
        background: #0071e3 !important;
        color: #ffffff !important;
        border-radius: 980px !important;
        border: none !important;
        padding: 12px 24px !important;
        font-weight: 600 !important;
        font-size: 15px !important;
        box-shadow: 0 4px 12px rgba(0, 113, 227, 0.2) !important;
    }
</style>
""",
    unsafe_allow_html=True,
)

# Header Tiện Ích
st.markdown(
    """
<div class="apple-header-box">
    <div>
        <h1 class="apple-main-title">🛠️ Tiện Ích</h1>
    </div>
    <div>
        <div class="apple-sub-text">
            Nền tảng công cụ hỗ trợ<br>công việc thông minh.
            <span>Lê Công Minh Workspace ›</span>
        </div>
    </div>
</div>
""",
    unsafe_allow_html=True,
)

# Quản lý trạng thái chuyển đổi giữa KSNB Tồn Kho và Nghiệm Thu HRC
if 'menu_chinh' not in st.session_state:
  st.session_state.menu_chinh = 'KSNB'

# Thanh nút bấm tiện ích trên cùng màn hình
col_nav1, col_nav2, col_nav3, _ = st.columns([1.5, 1.8, 2.2, 4.5])
with col_nav1:
  if st.button('✍️ Sforum', use_container_width=True):
    st.switch_page('pages/Sforum.py')

with col_nav2:
  if st.button('🛡️ KSNB Tồn Kho', use_container_width=True):
    st.session_state.menu_chinh = 'KSNB'

with col_nav3:
  if st.button('📑 Đối Chiếu Nghiệm Thu HRC', use_container_width=True):
    st.session_state.menu_chinh = 'HRC'

st.markdown(
    "<hr style='border: none; border-top: 1px solid #e2e8f0; margin: 15px 0"
    " 25px 0;'>",
    unsafe_allow_html=True,
)


# ==============================================================================
# CHỨC NĂNG 1: ĐỐI CHIẾU NGHIỆM THU HRC (5 FILE RỜI - CHUẨN 20 CỘT VÀNG & TÔ ĐỎ)
# ==============================================================================
def doi_chieu_hrc_tung_file():
  st.title('📑 ĐỐI CHIẾU NGHIỆM THU THÉP CUỘN CÁN NÓNG (HRC)')
  st.caption('Kiểm soát nội bộ • Tự động đối soát 4 bên và tô đỏ ô sai phạm')

  col1, col2 = st.columns(2)
  with col1:
    f_bbnt = st.file_uploader(
        '1. File Biên bản nghiệm thu (BBNT)', type=['xlsx', 'xls']
    )
    f_pl = st.file_uploader(
        '2. File Packing List nhà cung cấp', type=['xlsx', 'xls']
    )
    f_tc = st.file_uploader('3. File Dữ liệu Trạm cân', type=['xlsx', 'xls'])
  with col2:
    f_bv = st.file_uploader('4. File Dữ liệu Nhập Bảo vệ', type=['xlsx', 'xls'])
    f_opm = st.file_uploader(
        '5. File Báo cáo ERP OPM 105', type=['xlsx', 'xls']
    )

  if f_bbnt is not None:
    if st.button('🚀 Bắt đầu đối chiếu tự động', use_container_width=True):
      with st.spinner('Đang đối soát dữ liệu và định dạng màu...'):
        try:
          # 1. ĐỌC BBNT
          df_bbnt = pd.read_excel(f_bbnt, skiprows=7)
          df_bbnt = df_bbnt.dropna(subset=[df_bbnt.columns[3]]).copy()

          # Giữ lại các cột thông tin gốc (STT -> Khối lượng cân treo)
          base_cols = df_bbnt.columns[:13].tolist()
          res_df = df_bbnt[base_cols].copy()

          col_ma_cuon = df_bbnt.columns[3]
          col_phieu_can = df_bbnt.columns[4]
          col_net_ct = df_bbnt.columns[7]
          col_gross_ct = df_bbnt.columns[8]
          col_can_ban = df_bbnt.columns[9]
          col_lech_pt_bbnt = df_bbnt.columns[11]

          ma_cuon_clean = (
              df_bbnt[col_ma_cuon]
              .astype(str)
              .str.replace(r'\.0$', '', regex=True)
              .str.strip()
              .str.upper()
          )
          phieu_can_clean = (
              df_bbnt[col_phieu_can]
              .astype(str)
              .str.replace(r'\.0$', '', regex=True)
              .str.strip()
          )
          ma_hoa = ma_cuon_clean + 'A'

          # 2. XỬ LÝ PACKING LIST
          dict_pl = {}
          if f_pl is not None:
            df_pl = pd.read_excel(f_pl, skiprows=6)
            col_pl_ma = df_pl.columns[7]
            col_pl_net = df_pl.columns[11]
            df_pl_clean = df_pl.dropna(subset=[col_pl_ma, col_pl_net]).copy()
            pl_key = (
                df_pl_clean[col_pl_ma]
                .astype(str)
                .str.replace(r'\.0$', '', regex=True)
                .str.strip()
                .str.upper()
                + 'A'
            )
            pl_net_val = pd.to_numeric(
                df_pl_clean[col_pl_net], errors='coerce'
            ).fillna(0)
            dict_pl = (
                pd.DataFrame({'Key': pl_key, 'Net': pl_net_val})
                .groupby('Key')['Net']
                .sum()
                .to_dict()
            )

          # 3. XỬ LÝ DỮ LIỆU NHẬP BẢO VỆ
          dict_bv_phieu = {}
          dict_bv_kl = {}
          if f_bv is not None:
            df_bv = pd.read_excel(f_bv)
            for _, r in df_bv.iterrows():
              txt = f"{r.get('MÃ HÀNG HÓA', '')} {r.get('TÊN HÀNG HÓA', '')}".upper()
              p_can = str(r.get('MÃ PHIẾU CÂN', '')).replace('.0', '').strip()
              kl_can = r.get('KHỐI LƯỢNG PHIẾU CÂN', 0)
              for m in ma_cuon_clean:
                if m in txt and m not in dict_bv_phieu:
                  dict_bv_phieu[m] = p_can
                  dict_bv_kl[m] = kl_can

          # 4. XỬ LÝ TRẠM CÂN
          dict_tc_kl = {}
          dict_tc_phanloai = {}
          if f_tc is not None:
            df_tc = pd.read_excel(f_tc)
            tc_p = (
                df_tc['PHIẾU CÂN']
                .astype(str)
                .str.replace(r'\.0$', '', regex=True)
                .str.strip()
            )
            tc_kl = pd.to_numeric(
                df_tc['KHỐI LƯỢNG PHIẾU CÂN'], errors='coerce'
            ).fillna(0)
            tc_pl = (
                df_tc['PHÂN LOẠI'].fillna('Nhập hàng').astype(str).str.strip()
                if 'PHÂN LOẠI' in df_tc.columns
                else 'Nhập hàng'
            )
            dict_tc_kl = dict(zip(tc_p, tc_kl))
            if isinstance(tc_pl, str):
              dict_tc_phanloai = {k: tc_pl for k in tc_p}
            else:
              dict_tc_phanloai = dict(zip(tc_p, tc_pl))

          # 5. XỬ LÝ OPM 105
          opm_035_po, opm_035_rtv, opm_031, opm_03g = {}, {}, {}, {}
          if f_opm is not None:
            df_opm = pd.read_excel(f_opm, skiprows=8)
            col_opm_ma = df_opm.columns[0]
            col_opm_loai = df_opm.columns[1]
            col_opm_kl = df_opm.columns[10]
            col_opm_kho = df_opm.columns[35]

            df_opm_clean = df_opm.dropna(
                subset=[col_opm_ma, col_opm_loai, col_opm_kl]
            ).copy()
            df_opm_clean['Mã Lô'] = (
                df_opm_clean[col_opm_ma].astype(str).str.strip().str.upper()
            )
            df_opm_clean['Khối Lượng'] = pd.to_numeric(
                df_opm_clean[col_opm_kl], errors='coerce'
            ).fillna(0)
            df_opm_clean['Mã Kho'] = (
                df_opm_clean[col_opm_kho].astype(str).str.strip()
            )

            opm_035_po = (
                df_opm_clean[
                    (df_opm_clean['Mã Kho'] == '035')
                    & (df_opm_clean[col_opm_loai] == 'PO Receipt')
                ]
                .groupby('Mã Lô')['Khối Lượng']
                .sum()
                .to_dict()
            )
            opm_035_rtv = (
                df_opm_clean[
                    (df_opm_clean['Mã Kho'] == '035')
                    & (df_opm_clean[col_opm_loai] == 'Return to Vendor')
                ]
                .groupby('Mã Lô')['Khối Lượng']
                .sum()
                .to_dict()
            )
            opm_031 = (
                df_opm_clean[
                    (df_opm_clean['Mã Kho'].isin(['03G', '03H']))
                    & (df_opm_clean[col_opm_loai] == 'Intransit Receipt')
                ]
                .groupby('Mã Lô')['Khối Lượng']
                .sum()
                .to_dict()
            )
            opm_03g = (
                df_opm_clean[
                    (df_opm_clean['Mã Kho'] == '03G')
                    & (df_opm_clean[col_opm_loai] == 'PO Receipt')
                ]
                .groupby('Mã Lô')['Khối Lượng']
                .sum()
                .to_dict()
            )

          # ==============================================================================
          # GẮN ĐÚNG 20 CỘT VÀNG THEO YÊU CẦU
          # ==============================================================================
          res_df['NHẬP/XUẤT'] = phieu_can_clean.map(
              lambda x: dict_tc_phanloai.get(x, 'Nhập hàng')
          )
          res_df['Mã hóa'] = ma_hoa
          res_df['Khối lượng net theo PL'] = ma_hoa.map(dict_pl).fillna(0)

          net_ct_val = pd.to_numeric(
              df_bbnt[col_net_ct], errors='coerce'
          ).fillna(0)
          res_df['Đối chiếu'] = (
              net_ct_val - res_df['Khối lượng net theo PL']
          ).astype(float)

          res_df['KL theo bc nhập hàng'] = ma_cuon_clean.map(
              dict_bv_kl
          ).fillna(0)
          res_df['Số phiếu cân'] = ma_cuon_clean.map(dict_bv_phieu).fillna('')
          res_df['Đối chiếu số phiếu cân'] = (
              phieu_can_clean == res_df['Số phiếu cân'].astype(str)
          )

          res_df['KL theo bc trạm cân'] = (
              phieu_can_clean.map(dict_tc_kl).fillna(0).astype(float)
          )
          res_df['Đối chiếu KL TC & KL BCBV'] = res_df[
              'KL theo bc trạm cân'
          ] == pd.to_numeric(res_df['KL theo bc nhập hàng'], errors='coerce')

          can_ban_val = (
              pd.to_numeric(df_bbnt[col_can_ban], errors='coerce')
              .fillna(0)
              .astype(float)
          )
          res_df['Đối chiếu KL TC & KL BBNT'] = (
              res_df['KL theo bc trạm cân'] == can_ban_val
          )

          gross_ct_val = (
              pd.to_numeric(df_bbnt[col_gross_ct], errors='coerce')
              .replace(0, pd.NA)
              .fillna(1)
              .astype(float)
          )
          res_df['Chênh lệch (kg)'] = res_df['KL theo bc trạm cân'] - gross_ct_val
          res_df['Chênh lệch dung sai'] = (
              res_df['Chênh lệch (kg)'] / gross_ct_val
          )

          lech_pt_val = pd.to_numeric(
              df_bbnt[col_lech_pt_bbnt], errors='coerce'
          ).fillna(0)
          res_df['Đối chiếu dung sai'] = (
              res_df['Chênh lệch dung sai'].round(6) == lech_pt_val.round(6)
          )
          res_df['So với quy định ( >+-0.3%)'] = res_df[
              'Chênh lệch dung sai'
          ].apply(lambda x: 'VUOT DUNG SAI' if abs(x) > 0.003 else 'OK')

          res_df['KL Nhập kho 035'] = ma_hoa.map(
              lambda x: opm_035_po.get(x, 0) - abs(opm_035_rtv.get(x, 0))
          )
          res_df['Chênh lệch '] = (
              res_df['Khối lượng net theo PL'] - res_df['KL Nhập kho 035']
          )

          res_df['KL Nhập kho 031'] = ma_hoa.map(lambda x: opm_031.get(x, 0))
          res_df['Chênh lệch  '] = res_df.apply(
              lambda r: 0
              if r['KL Nhập kho 031'] == 0
              else r['Khối lượng net theo PL'] - r['KL Nhập kho 031'],
              axis=1,
          )

          res_df['KL Nhập kho 03G-3H'] = ma_hoa.map(lambda x: opm_03g.get(x, 0))
          res_df['Chênh lệch   '] = (
              res_df['Khối lượng net theo PL'] - res_df['KL Nhập kho 03G-3H']
          )

          # ==============================================================================
          # XUẤT EXCEL: TÔ VÀNG HEADER & TÔ ĐỎ Ô SAI PHẠM
          # ==============================================================================
          output = io.BytesIO()
          with pd.ExcelWriter(output, engine='openpyxl') as writer:
            res_df.to_excel(writer, index=False, sheet_name='BBNT_Doi_Chieu')
            ws = writer.sheets['BBNT_Doi_Chieu']

            yellow_fill = PatternFill(
                start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid'
            )
            red_fill = PatternFill(
                start_color='FFFFC7CE', end_color='FFFFC7CE', fill_type='solid'
            )
            red_font = Font(color='9C0006', bold=True)
            border_thin = Border(
                left=Side(style='thin', color='D9D9D9'),
                right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'),
                bottom=Side(style='thin', color='D9D9D9'),
            )

            # Tô màu vàng đúng 20 cột đối chiếu ở Header Row 1
            for c in range(14, 34):
              cell = ws.cell(row=1, column=c)
              cell.fill = yellow_fill
              cell.font = Font(bold=True)

            # Quét kiểm tra dữ liệu và TÔ ĐỎ ô bị SAI
            col_names = list(res_df.columns)
            for r_idx in range(2, len(res_df) + 2):
              for c_idx in range(1, len(col_names) + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                c_name = col_names[c_idx - 1]
                val = cell.value

                cell.border = border_thin
                is_error = False

                if (
                    c_name
                    in [
                        'Đối chiếu số phiếu cân',
                        'Đối chiếu KL TC & KL BCBV',
                        'Đối chiếu KL TC & KL BBNT',
                        'Đối chiếu dung sai',
                    ]
                    and val is False
                ):
                  is_error = True
                elif (
                    c_name == 'So với quy định ( >+-0.3%)'
                    and str(val).strip() == 'VUOT DUNG SAI'
                ):
                  is_error = True
                elif (
                    c_name in ['Đối chiếu', 'Chênh lệch ']
                    and isinstance(val, (int, float))
                    and val != 0
                ):
                  is_error = True

                if is_error:
                  cell.fill = red_fill
                  cell.font = red_font

          st.success('✅ Đã đối soát xong!')
          st.dataframe(res_df, use_container_width=True)

          st.download_button(
              '📥 Tải File Kết Quả Nghiệm Thu (Đã Tô Đỏ Lỗi)',
              data=output.getvalue(),
              file_name='BBNT_Nghiem_Thu_Da_Doi_Chieu.xlsx',
              mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          )

        except Exception as e:
          st.error(f'Lỗi xử lý: {e}')
  else:
    st.info('👉 Hãy nạp tối thiểu File BBNT để bắt đầu đối chiếu.')


# ==============================================================================
# ĐIỀU HƯỚNG HIỂN THỊ CHỨC NĂNG
# ==============================================================================
if st.session_state.menu_chinh == 'HRC':
  doi_chieu_hrc_tung_file()

else:
  # CHỨC NĂNG 2: KIỂM SOÁT TỒN KHO HÀNG NGÀY
  st.title('🛡️ HỆ THỐNG KIỂM SOÁT TỒN KHO HÀNG NGÀY')
  st.caption(
      'Kiểm soát nội bộ (KSNB) • Tự động quét lỗi tỷ trọng & kiểm tra phẩm cấp'
      ' cuộn'
  )

  uploaded_file = st.file_uploader(
      'Kéo thả file Excel tồn kho (.xlsx):', type=['xlsx']
  )

  def load_data(file_source):
    headers = [
        'Kho',
        'Ma_hang',
        'Ten_hang',
        'Ma_loai',
        'Phan_loai',
        'Phan_loai_SX',
        'Tieu_chuan',
        'Grade',
        'BMT',
        'TCT',
        'APT',
        'Kho_mm',
        'Do_ma',
        'Phu_son',
        'Ma_mau',
        'Ma_cuon',
        'DVT',
        'Gross',
        'Net',
        'Loi',
        'So_met',
        'A',
        'B',
        'C',
        'Ty_trong',
        'Don_gia',
        'Thanh_tien',
        'Xuat_xu',
        'Don_hang',
    ]
    df = pd.read_excel(file_source, sheet_name=0, skiprows=13, header=None)
    df = df.iloc[:, : len(headers)]
    df.columns = headers

    df = df.dropna(subset=['Kho', 'Net']).copy()

    df['Kho'] = df['Kho'].astype(str).str.strip()
    df['Phan_loai'] = df['Phan_loai'].fillna('Khác').astype(str).str.strip()
    df['Net_KG'] = pd.to_numeric(df['Net'], errors='coerce').fillna(0)
    df['Net_Tan'] = df['Net_KG'] / 1000
    df['So_met'] = pd.to_numeric(df['So_met'], errors='coerce').fillna(0)
    df['BMT'] = pd.to_numeric(df['BMT'], errors='coerce')
    df['Ty_trong'] = pd.to_numeric(df['Ty_trong'], errors='coerce').fillna(0)
    return df

  if uploaded_file is None:
    st.info(
        '👉 Hãy kéo thả file báo cáo tồn kho Excel (.xlsx) vào khung bên dưới'
        ' để bắt đầu phân tích.'
    )
  else:
    try:
      df_raw = load_data(uploaded_file)
      st.sidebar.success(f'Đã nạp file: {uploaded_file.name}')

      with st.sidebar:
        st.header('🔍 Bộ Lọc Nhanh')
        kho_list = sorted([str(x) for x in df_raw['Kho'].unique()])
        selected_kho = st.multiselect(
            'Lọc Vị trí kho:', options=kho_list, default=kho_list
        )

        pl_list = sorted([str(x) for x in df_raw['Phan_loai'].unique()])
        selected_pl = st.multiselect(
            'Lọc Phân loại:', options=pl_list, default=pl_list
        )

        search_coil = st.text_input('Tìm kiếm Mã cuộn / Mã hàng:')

      df = df_raw[
          (df_raw['Kho'].isin(selected_kho))
          & (df_raw['Phan_loai'].isin(selected_pl))
      ]
      if search_coil:
        df = df[
            df['Ma_cuon'].astype(str).str.contains(search_coil, case=False, na=False)
            | df['Ma_hang']
            .astype(str)
            .str.contains(search_coil, case=False, na=False)
        ]

      c1, c2, c3, c4 = st.columns(4)
      c1.metric(
          'Tổng Tồn Kho', f"{df['Net_Tan'].sum():,.2f} Tấn", f'{len(df):,} Cuộn/Lô'
      )
      c2.metric(
          'Hàng Loại 1',
          f"{df[df['Phan_loai'] == '1']['Net_Tan'].sum():,.2f} Tấn",
          f"{(df['Phan_loai'] == '1').sum():,} Cuộn",
      )
      c3.metric(
          'Chờ XL & Phế phẩm',
          f"{df[df['Phan_loai'].isin(['CXL', 'PP'])]['Net_Tan'].sum():,.2f} Tấn",
          f"{df['Phan_loai'].isin(['CXL', 'PP']).sum():,} Cuộn",
      )

      rot_xk = df[df['Phan_loai'].isin(['A1', 'A2'])]
      c4.metric(
          'Hàng Rớt XK (A1, A2)',
          f"{rot_xk['Net_Tan'].sum():,.2f} Tấn",
          f'{len(rot_xk):,} Cuộn',
      )

      st.markdown('---')

      col_a, col_b = st.columns(2)
      with col_a:
        kho_sum = df.groupby('Kho')['Net_Tan'].sum().reset_index()
        fig1 = px.bar(
            kho_sum,
            x='Kho',
            y='Net_Tan',
            text_auto='.2f',
            color='Kho',
            title='Khối lượng theo Kho (Tấn)',
        )
        st.plotly_chart(fig1, use_container_width=True)

      with col_b:
        pl_sum = df.groupby('Phan_loai')['Net_Tan'].sum().reset_index()
        fig2 = px.pie(
            pl_sum,
            names='Phan_loai',
            values='Net_Tan',
            hole=0.35,
            title='Cơ cấu theo Phân loại chất lượng',
        )
        st.plotly_chart(fig2, use_container_width=True)

      st.markdown('---')

      st.subheader('🚨 DANH SÁCH BẤT THƯỜNG CẦN KIỂM SOÁT')
      t1, t2, t3 = st.tabs([
          '⚠️ Lệch Tỷ Trọng / Thiếu Mét',
          '⚠️ Kiểm Tra Quy Chuẩn A2',
          '📋 Bảng Dữ Liệu Chi Tiết',
      ])

      with t1:
        err_tytrong = df[(df['Net_KG'] > 0) & (df['So_met'] == 0)][[
            'Kho',
            'Ma_cuon',
            'Ma_hang',
            'Ten_hang',
            'Net_KG',
            'So_met',
            'Ty_trong',
        ]]
        st.write(
            f'Tìm thấy **{len(err_tytrong)}** dòng có khối lượng Net nhưng'
            ' thiếu số mét (tỷ trọng = 0):'
        )
        st.dataframe(err_tytrong, use_container_width=True)

      with t2:
        a2_all = df[df['Phan_loai'] == 'A2']
        st.write(
            f'Tổng số cuộn A2: **{len(a2_all)}** cuộn (Quy định KSNB: BMT phải'
            ' ≥ 0.57mm).'
        )
        a2_err = a2_all[a2_all['BMT'] < 0.57]
        if len(a2_err) == 0:
          st.success('✅ Toàn bộ cuộn hàng A2 đều đúng quy định độ dày!')
        else:
          st.error(
              f'⚠️ Có {len(a2_err)} cuộn A2 vi phạm tiêu chuẩn độ dày:'
          )
          st.dataframe(
              a2_err[['Kho', 'Ma_cuon', 'Ten_hang', 'BMT', 'Net_KG']],
              use_container_width=True,
          )

      with t3:
        st.dataframe(
            df[[
                'Kho',
                'Ma_cuon',
                'Ma_hang',
                'Ten_hang',
                'Phan_loai',
                'BMT',
                'Net_KG',
                'So_met',
                'Ty_trong',
            ]],
            use_container_width=True,
        )

    except Exception as e:
      st.error(f'Lỗi khi đọc file: {e}')
