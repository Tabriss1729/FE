import streamlit as st
import pandas as pd
import numpy as np
import json
import re
import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font
import extra_streamlit_components as stx

# --- 1. 頁面與常數設定 ---
st.set_page_config(page_title="Faradaic Efficiency 計算機", layout="wide")

@st.experimental_dialog("🎉 歡迎使用 FE 數據計算機！(新手快速導覽)")
def show_tutorial(cookie_manager):
    st.write("偵測到您是第一次使用本系統！讓我們花 30 秒了解核心功能：")
    st.info("💡 **小提示**：本系統的所有計算都在您的瀏覽器與雲端伺服器間安全進行。")
    st.markdown("""
    1. **📂 讀取舊有設定**：將之前的 `.json` 檔拖曳到左側側邊欄。
    2. **➕ 批量新增空行**：在下方設定需要幾行數據，可一次展開表格。
    3. **🪄 批量修改**：快速將所有數據列套用相同的催化劑或稀釋倍率。
    4. **📥 下載專業報表**：點擊計算後，下載的 Excel 會自動為您合併儲存格並加上化學下標！
    """)
    st.write("準備好開始了嗎？")
    if st.button("我了解了！開始使用", type="primary", use_container_width=True):
        cookie_manager.set('has_seen_tutorial', 'true', max_age=3650*24*60*60)
        st.rerun()

cookie_manager = stx.CookieManager()
if cookie_manager.get_all() is not None:
    if cookie_manager.get('has_seen_tutorial') != 'true':
        show_tutorial(cookie_manager)

st.title("⚡ Faradaic Efficiency 數據計算機")

F_const = 96485
HCELL_FORMULA = "(Conc * 50 * Dilution * 1e-6 * n_e * F) / Q * 100"
GDE_N_FORMULA = "(C1 * V_acid + C2 * V_re) * Dilution"
GDE_FE_FORMULA = "(Total_n * 1e-6 * n_e * F) / Q * 100"
today_str = datetime.date.today().strftime("%Y%m%d")

# --- 初始化 Session State ---
if 'mode' not in st.session_state: st.session_state.mode = "GDE (雙槽)"
if 'q_toggle' not in st.session_state: st.session_state.q_toggle = False
if 'total_q' not in st.session_state: st.session_state.total_q = 100.0
if 'electrolyte' not in st.session_state: st.session_state.electrolyte = "0.5 M KNO3 + 0.1 M KOH" 
if 'acid_vol' not in st.session_state: st.session_state.acid_vol = 10.0
if 're_vol' not in st.session_state: st.session_state.re_vol = 50.0

if 'hcell_data' not in st.session_state:
    st.session_state.hcell_data = pd.DataFrame({
        "Product": pd.Series(dtype='str'), "Catalyst": pd.Series(dtype='str'), "Loading (μl)": pd.Series(dtype='float'),
        "V vs RHE": pd.Series(dtype='float'), "稀釋倍率": pd.Series(dtype='float'), "Conc. (μmol)": pd.Series(dtype='float')
    })
if 'gde_data' not in st.session_state:
    st.session_state.gde_data = pd.DataFrame({
        "Product": pd.Series(dtype='str'), "Catalyst": pd.Series(dtype='str'), "Loading (μl)": pd.Series(dtype='float'),
        "V vs RHE": pd.Series(dtype='float'), "稀釋倍率": pd.Series(dtype='float'), "Acid C1 (mM)": pd.Series(dtype='float'), "RE C2 (mM)": pd.Series(dtype='float')
    })

# --- 2. 側邊欄：設定管理與實驗參數 ---
with st.sidebar:
    st.header("💾 設定與檔案管理")
    with st.expander("JSON 設定檔操作", expanded=True):
        uploaded_json = st.file_uploader("📂 讀取設定 (.json)", type="json")
        if uploaded_json is not None:
            try:
                data = json.load(uploaded_json)
                gp = data.get('global_params', {})
                if 'mode' in gp: st.session_state.mode = "H-cell (單槽)" if gp['mode'] == "H-cell" else "GDE (雙槽)"
                st.session_state.q_toggle = gp.get('gde_q_toggle', False)
                st.session_state.total_q = gp.get('total_coulomb', 100.0)
                st.session_state.electrolyte = gp.get('electrolyte', "0.5 M KNO3 + 0.1 M KOH")
                st.session_state.acid_vol = gp.get('acid_vol', 10.0)
                st.session_state.re_vol = gp.get('re_vol', 50.0)
                
                rows = data.get('rows', [])
                if rows:
                    df = pd.DataFrame(rows)
                    mapping = {'product': 'Product', 'catalyst': 'Catalyst', 'volume_ul': 'Loading (μl)', 'v_rhe': 'V vs RHE', 'dilution': '稀釋倍率', 'conc': 'Conc. (μmol)', 'acid_c1': 'Acid C1 (mM)', 're_c2': 'RE C2 (mM)'}
                    df = df.rename(columns=mapping)
                    if gp.get('mode') == "H-cell": st.session_state.hcell_data = df[[c for c in st.session_state.hcell_data.columns if c in df.columns]]
                    else: st.session_state.gde_data = df[[c for c in st.session_state.gde_data.columns if c in df.columns]]
                st.success("✅ 已載入設定")
            except: st.error("讀取失敗")

        st.divider()
        st.markdown("##### 💾 儲存設定")
        custom_json_name = st.text_input("自訂 JSON 檔名", value="FE_Config")
        json_filename_final = f"{today_str}_{custom_json_name}.json"

        save_data = {
            'global_params': {'mode': "H-cell" if "H-cell" in st.session_state.mode else "GDE", 'total_coulomb': st.session_state.total_q, 'electrolyte': st.session_state.electrolyte, 'acid_vol': st.session_state.acid_vol, 're_vol': st.session_state.re_vol, 'gde_q_toggle': st.session_state.q_toggle},
            'rows': st.session_state.hcell_data.to_dict('records') if "H-cell" in st.session_state.mode else st.session_state.gde_data.to_dict('records')
        }
        st.download_button("💾 儲存為 JSON", data=json.dumps(save_data, ensure_ascii=False, indent=4), file_name=json_filename_final)

    st.markdown("---")
    st.header("🧪 實驗參數")
    mode = st.radio("實驗模式", ["H-cell (單槽)", "GDE (雙槽)"], index=0 if "H-cell" in st.session_state.mode else 1)
    st.session_state.mode = mode
    
    if "GDE" in mode:
        is_n2_mode = st.toggle("通入氮氣 (N2 Mode)", value=st.session_state.q_toggle)
        st.session_state.q_toggle = is_n2_mode
    else:
        is_n2_mode = False 

    total_q = st.number_input("總電量 Q (C)", value=float(st.session_state.total_q), step=10.0)
    st.session_state.total_q = total_q
    electrolyte = st.text_input("電解液", value=st.session_state.electrolyte)
    st.session_state.electrolyte = electrolyte

    if "GDE" in mode:
        st.session_state.acid_vol = st.number_input("Acid 側體積 (mL)", value=float(st.session_state.acid_vol))
        st.session_state.re_vol = st.number_input("RE 側體積 (mL)", value=float(st.session_state.re_vol))

# --- 3. 表格操作 ---
with st.expander("🛠️ 表格操作 (新增行數 / 批量編輯)", expanded=False):
    st.markdown("##### ➕ 批量新增空行")
    col_add1, col_add2, _ = st.columns([1, 1, 3])
    with col_add1:
        add_count = st.number_input("輸入要新增的行數", min_value=1, max_value=50, value=1, step=1, label_visibility="collapsed")
    with col_add2:
        if st.button("確認新增"):
            new_rows = []
            for _ in range(add_count):
                if "H-cell" in mode:
                    new_rows.append({"Product": "NH3", "Catalyst": "", "Loading (μl)": None, "V vs RHE": None, "稀釋倍率": 1.0, "Conc. (μmol)": None})
                else:
                    new_rows.append({"Product": "NH3", "Catalyst": "", "Loading (μl)": None, "V vs RHE": None, "稀釋倍率": 1.0, "Acid C1 (mM)": None, "RE C2 (mM)": None})
            
            new_df = pd.DataFrame(new_rows)
            if "H-cell" in mode: st.session_state.hcell_data = pd.concat([st.session_state.hcell_data, new_df], ignore_index=True)
            else: st.session_state.gde_data = pd.concat([st.session_state.gde_data, new_df], ignore_index=True)
            st.rerun()

    st.divider()

    st.markdown("##### 🪄 批量修改 (留空則不修改)")
    col1, col2, col3, col4 = st.columns(4)
    with col1: b_cat = st.text_input("批量更新催化劑")
    with col2: b_load = st.text_input("批量更新 Loading (μl)")
    with col3: b_vrhe = st.text_input("批量更新 V vs RHE")
    with col4: b_dil = st.text_input("批量更新稀釋倍率")
    
    if st.button("套用批量修改"):
        # 💡 核心修正：使用 .copy() 創造全新資料表，並明確覆寫回 session_state
        try:
            if "H-cell" in mode:
                target_df = st.session_state.hcell_data.copy()
                if b_cat: target_df["Catalyst"] = b_cat
                if b_load: target_df["Loading (μl)"] = float(b_load)
                if b_vrhe: target_df["V vs RHE"] = float(b_vrhe)
                if b_dil: target_df["稀釋倍率"] = float(b_dil)
                st.session_state.hcell_data = target_df
            else:
                target_df = st.session_state.gde_data.copy()
                if b_cat: target_df["Catalyst"] = b_cat
                if b_load: target_df["Loading (μl)"] = float(b_load)
                if b_vrhe: target_df["V vs RHE"] = float(b_vrhe)
                if b_dil: target_df["稀釋倍率"] = float(b_dil)
                st.session_state.gde_data = target_df
            st.rerun()
        except ValueError:
            st.error("數值格式輸入錯誤！請確保 Loading, V vs RHE, 稀釋倍率 欄位中輸入的是數字。")

# --- 4. 數據表格 ---
st.subheader(f"📊 數據輸入 - {mode}")
current_df = st.session_state.hcell_data if "H-cell" in mode else st.session_state.gde_data

edited_df = st.data_editor(
    current_df,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "Product": st.column_config.SelectboxColumn("產物", options=["NH3"] if is_n2_mode else ["NH3", "NO2"], required=True)
    }
)
if "H-cell" in mode: st.session_state.hcell_data = edited_df
else: st.session_state.gde_data = edited_df

# --- 5. 計算與匯出設定 ---
st.divider()
st.subheader("📥 計算與匯出")

col_name1, _ = st.columns([1, 2])
with col_name1:
    if "H-cell" in mode:
        default_excel_name = "FE_Result_H-cell"
    else:
        gas_str = "N2_Gas" if is_n2_mode else "Ar_Gas"
        default_excel_name = f"FE_Result_GDE_{gas_str}"
        
    custom_excel_name = st.text_input("自訂 Excel 檔名", value=default_excel_name)
    excel_filename_final = f"{today_str}_{custom_excel_name}.xlsx"

if st.button("🔄 開始計算 FE", type="primary"):
    res_df = edited_df.copy()
    fe_res, tn_res = [], []
    for _, row in res_df.iterrows():
        try:
            prod = row["Product"]
            n_e = 6 if (is_n2_mode and prod == 'NH3') else (8 if prod == 'NH3' else (2 if prod == 'NO2' else np.nan))
            dil = float(row["稀釋倍率"])
            if "H-cell" in mode:
                env = {'Conc': float(row["Conc. (μmol)"]), 'Dilution': dil, 'Q': total_q, 'n_e': n_e, 'F': F_const}
                fe_res.append(round(eval(HCELL_FORMULA, {"__builtins__": {}}, env), 2))
            else:
                env_n = {'C1': float(row["Acid C1 (mM)"]), 'C2': float(row["RE C2 (mM)"]), 'V_acid': st.session_state.acid_vol, 'V_re': st.session_state.re_vol, 'Dilution': dil}
                tn = eval(GDE_N_FORMULA, {"__builtins__": {}}, env_n)
                tn_res.append(round(tn, 3))
                env_fe = {'Total_n': tn, 'Q': total_q, 'n_e': n_e, 'F': F_const}
                fe_res.append(round(eval(GDE_FE_FORMULA, {"__builtins__": {}}, env_fe), 2))
        except:
            fe_res.append("Error")
            if "GDE" in mode: tn_res.append("Error")
    
    if "GDE" in mode: res_df["Total n (μmol)"] = tn_res
    res_df["FE (%)"] = fe_res
    st.dataframe(res_df, use_container_width=True)

    def to_pro_excel(df, curr_mode, curr_electrolyte, curr_Q, is_n2):
        def apply_subscript(text):
            if pd.isna(text): return ""
            text_str = str(text)
            try:
                f_val = float(text_str)
                if np.isnan(f_val): return ""
                return int(f_val) if f_val.is_integer() else f_val
            except ValueError: pass
            subscript_map = str.maketrans("0123456789", "₀₁₂₃₄₅₆₇₈₉")
            return re.sub(r'([a-zA-Z])(\d+)', lambda m: m.group(1) + m.group(2).translate(subscript_map), text_str)

        wb = Workbook()
        ws = wb.active
        ws.title = "FE_Results"
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        export_data = []
        gas_mode_str = "N2_Gas" if is_n2 else "Ar_Gas"

        for _, row in df.iterrows():
            if "H-cell" in curr_mode:
                export_data.append({
                    'Cell': f"H-cell({row['Product']})", 'Electrolyte': curr_electrolyte, 'Total Coulomb (Q)': curr_Q,
                    'Product Type': row['Product'], 'Catalyst': row['Catalyst'], 'Loading (μl)': row['Loading (μl)'],
                    '稀釋倍率': row['稀釋倍率'], 'V vs RHE': row['V vs RHE'], 'Total Concentration (μmol)': row['Conc. (μmol)'],
                    'Faradaic Efficiency (%)': row['FE (%)']
                })
            else:
                export_data.append({
                    'Cell': f"GDE_{gas_mode_str}({row['Product']})", 'Electrolyte': curr_electrolyte, 'Total Coulomb (Q)': curr_Q,
                    'Product Type': row['Product'], 'Catalyst': row['Catalyst'], 'Loading (μl)': row['Loading (μl)'],
                    '稀釋倍率': row['稀釋倍率'], 'V vs RHE': row['V vs RHE'], 'Acid C1 (mM)': row['Acid C1 (mM)'], 
                    'RE C2 (mM)': row['RE C2 (mM)'], 'Total Concentration (μmol)': row['Total n (μmol)'], 'Faradaic Efficiency (%)': row['FE (%)']
                })
                
        df_export = pd.DataFrame(export_data)

        cur_row = 1
        for prod, group in df_export.groupby("Product Type"):
            cols = list(group.columns)
            for c_idx, c_name in enumerate(cols):
                cell = ws.cell(row=cur_row, column=c_idx+1, value=apply_subscript(c_name))
                cell.font = Font(bold=True)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            
            start_r = cur_row + 1
            for r_idx, r_data in enumerate(group.values.tolist()):
                for c_idx, val in enumerate(r_data):
                    cell = ws.cell(row=start_r+r_idx, column=c_idx+1, value=apply_subscript(val))
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal='center', vertical='center')
            
            end_r = start_r + len(group) - 1

            if end_r > start_r:
                for col_idx in [1, 2, 3, 4, 7]:
                    ws.merge_cells(start_row=start_r, end_row=end_r, start_column=col_idx, end_column=col_idx)

                catalyst_starts = start_r
                current_cat = ws.cell(row=start_r, column=5).value
                for r in range(start_r + 1, end_r + 2):
                    cell_val = ws.cell(row=r, column=5).value if r <= end_r else None
                    if cell_val != current_cat:
                        if (r - 1) > catalyst_starts:
                            ws.merge_cells(start_row=catalyst_starts, end_row=r-1, start_column=5, end_column=5)
                            ws.merge_cells(start_row=catalyst_starts, end_row=r-1, start_column=6, end_column=6)
                        catalyst_starts = r
                        current_cat = cell_val

            cur_row = end_r + 2

        for col in ws.columns:
            max_length = 0
            column_letter = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except: pass
            ws.column_dimensions[column_letter].width = (max_length + 2) * 1.2

        out = BytesIO()
        wb.save(out)
        return out.getvalue()

    st.download_button("📥 下載 Excel", data=to_pro_excel(res_df, mode, electrolyte, total_q, is_n2_mode), file_name=excel_filename_final)
