import streamlit as st
import pandas as pd
import numpy as np
import json
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font

# --- 1. 頁面與常數設定 ---
st.set_page_config(page_title="Faradaic Efficiency 計算機", layout="wide")
st.title("⚡ Faradaic Efficiency 數據計算機")

F_const = 96485
# 固定計算公式
HCELL_FORMULA = "(Conc * 50 * Dilution * 1e-6 * n_e * F) / Q * 100"
GDE_N_FORMULA = "(C1 * V_acid + C2 * V_re) * Dilution"
GDE_FE_FORMULA = "(Total_n * 1e-6 * n_e * F) / Q * 100"

# --- 初始化 Session State ---
if 'mode' not in st.session_state: st.session_state.mode = "GDE (雙槽)"
if 'q_toggle' not in st.session_state: st.session_state.q_toggle = False
if 'total_q' not in st.session_state: st.session_state.total_q = 100.0
if 'electrolyte' not in st.session_state: st.session_state.electrolyte = "" # 預設清空
if 'acid_vol' not in st.session_state: st.session_state.acid_vol = 10.0
if 're_vol' not in st.session_state: st.session_state.re_vol = 50.0

# 💡 修改點：將初始表格改為帶有型態的「完全空白 DataFrame」
if 'hcell_data' not in st.session_state:
    st.session_state.hcell_data = pd.DataFrame({
        "Product": pd.Series(dtype='str'),
        "Catalyst": pd.Series(dtype='str'),
        "Loading (μl)": pd.Series(dtype='float'),
        "V vs RHE": pd.Series(dtype='float'),
        "稀釋倍率": pd.Series(dtype='float'),
        "Conc. (μmol)": pd.Series(dtype='float')
    })

if 'gde_data' not in st.session_state:
    st.session_state.gde_data = pd.DataFrame({
        "Product": pd.Series(dtype='str'),
        "Catalyst": pd.Series(dtype='str'),
        "Loading (μl)": pd.Series(dtype='float'),
        "V vs RHE": pd.Series(dtype='float'),
        "稀釋倍率": pd.Series(dtype='float'),
        "Acid C1 (mM)": pd.Series(dtype='float'),
        "RE C2 (mM)": pd.Series(dtype='float')
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
                st.session_state.electrolyte = gp.get('electrolyte', "")
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
        json_filename_final = f"{custom_json_name}.json" if not custom_json_name.endswith(".json") else custom_json_name

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
            # 💡 修改點：新增的行數預設改為空值 (None)
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
        target_df = st.session_state.hcell_data if "H-cell" in mode else st.session_state.gde_data
        if b_cat: target_df["Catalyst"] = b_cat
        if b_load: target_df["Loading (μl)"] = float(b_load)
        if b_vrhe: target_df["V vs RHE"] = float(b_vrhe)
        if b_dil: target_df["稀釋倍率"] = float(b_dil)
        st.rerun()

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
    excel_filename_final = f"{custom_excel_name}.xlsx" if not custom_excel_name.endswith(".xlsx") else custom_excel_name

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

    def to_pro_excel(df, m, el, q_val, n2):
        def sub(t): return re.sub(r'([a-zA-Z])(\d+)', lambda m: m.group(1) + m.group(2).translate(str.maketrans("0123456789", "₀₁₂₃₄₅₆₇₈₉")), str(t))
        wb = Workbook()
        ws = wb.active
        ws.title = "FE_Results"
        cur_row = 1
        for prod, group in df.groupby("Product"):
            cols = list(group.columns)
            for c_idx, c_name in enumerate(cols):
                cell = ws.cell(row=cur_row, column=c_idx+1, value=sub(c_name))
                cell.font = Font(bold=True)
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            start_r = cur_row + 1
            for r_idx, r_data in enumerate(group.values.tolist()):
                for c_idx, val in enumerate(r_data):
                    cell = ws.cell(row=start_r+r_idx, column=c_idx+1, value=sub(val))
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                    cell.alignment = Alignment(horizontal='center')
            cur_row = start_r + len(group) + 1
        
        out = BytesIO()
        wb.save(out)
        return out.getvalue()

    st.download_button("📥 下載專業排版 Excel", data=to_pro_excel(res_df, mode, electrolyte, total_q, is_n2_mode), file_name=excel_filename_final)