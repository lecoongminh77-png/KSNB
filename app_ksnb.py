df_bbnt[col_lech_pt_bbnt], errors='coerce'
          ).fillna(0)
          res_df['Đối chiếu dung sai'] = (
              res_df['Chênh lệch dung sai'].round(6) == lech_pt_val.round(6)
          )
          res_df['So với quy định ( >+-0.3%)'] = res_df[
              'Chênh lệch dung sai'
          ].apply(lambda x: 'VUOT DUNG SAI' if abs(x) > 0.003 else 'OK')

          # Các cột kho ERP
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
          # XUẤT EXCEL: TÔ VÀNG HEADER ĐỐI CHIẾU & TÔ ĐỎ Ô PHÁT HIỆN SAI
          # ===========================================================
