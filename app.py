import streamlit as st
import pandas as pd
import numpy as np
import json
import re
import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font

# --- 1. 頁面與常數設定 ---
st.set_page_config(page_title="Faradaic Efficiency 計算機", layout="wide")
st.title("⚡ Faradaic Efficiency 數據計算機")

F_const = 96485
today_str = datetime.date.today().strftime("%Y%m%d")

# --- 初始化 Session State ---
if 'editor_key' not in st.session_state: st.session_state.editor_key = 0
if 'loaded_file_id' not in st.session_state: st.session_state.loaded_file_id = ""

if 'mode' not in st.session_state: st.session_state.mode = "GDE (雙槽)"
if 'q_toggle' not in st.session_state: st.session_state.q_toggle = False
if 'custom_ne_toggle' not in st.session_state: st.session_state.custom_ne_toggle = False
if 'sidebar_ne' not in st.session_state: st.session_state.sidebar_ne = 8.0 
if 'total_q' not in st.session_state: st.session_state.total_q = 100.0
if 'electrolyte' not in st.session_state: st.session_state.electrolyte = "0.5 M KNO3 + 0.1 M KOH" 
if 'acid_vol' not in st.session_state: st.session_state.acid_vol = 10.0
if 're_vol' not in st.session_state: st.session_state.re_vol = 50.0

if 'hcell_formula' not in st.session_state: st.session_state.hcell_formula = "(Conc * 50 * Dilution * 1e-6 * n_e * F) / Q * 100"
if 'gde_n_formula' not in st.session_state: st.session_state.gde_n_formula = "(C1 * V_acid + C2 * V_re) * Dilution"
if 'gde_fe_formula' not in st.session_state: st.session_state.gde_fe_formula = "(Total_n * 1e-6 * n_e * F) / Q * 100"

if 'hcell_data' not in st.session_state:
    st.session_state.hcell_data = pd.DataFrame({
        "選取": pd.Series(dtype='bool'),
        "Product": pd.Series(dtype='str'), "n_e": pd.Series(dtype='float'), "Catalyst": pd.Series(dtype='str'), "Loading (μl)": pd.Series(dtype='float'),
        "V vs RHE": pd.Series(dtype='float'), "Dilution Factor": pd.Series(dtype='float'), "Conc. (μmol)": pd.Series(dtype='float')
    })
if 'gde_data' not in st.session_state:
    st.session_state.gde_data = pd.DataFrame({
        "選取": pd.Series(dtype='bool'),
        "Product": pd.Series(dtype='str'), "n_e": pd.Series(dtype='float'), "Catalyst": pd.Series(dtype='str'), "Loading (μl)": pd.Series(dtype='float'),
        "V vs RHE": pd.Series(dtype='float'), "Dilution Factor": pd.Series(dtype='float'), "Acid C1 (mM)": pd.Series(dtype='float'), "RE C2 (mM)": pd.Series(dtype='float')
    })

# 相容性防呆
for df_name in ['hcell_data', 'gde_data']:
    if '選取' not in st.session_state[df_name].columns: 
        st.session_state[df_name].insert(0, '選取', False)
    if 'n_e' not in st.session_state[df_name].columns:
        st.session_state[df_name].insert(2, 'n_e', np.nan)

# JSON 讀取中繼處理
if 'pending_json_data' in st.session_state:
    data = st.session_state.pending_json_data
    gp = data.get('global_params', {})
    
    if 'mode' in gp: st.session_state.mode = "H-cell (單槽)" if gp['mode'] == "H-cell" else "GDE (雙槽)"
    st.session_state.q_toggle = gp.get('gde_q_toggle', False)
    st.session_state.custom_ne_toggle = gp.get('custom_ne_toggle', False)
    st.session_state.sidebar_ne = gp.get('sidebar_ne', 8.0)
    st.session_state.total_q = gp.get('total_coulomb', 100.0)
    st.session_state.electrolyte = gp.get('electrolyte', "0.5 M KNO3 + 0.1 M KOH")
    st.session_state.acid_vol = gp.get('acid_vol', 10.0)
    st.session_state.re_vol = gp.get('re_vol', 50.0)
    
    st.session_state.hcell_formula = gp.get('hcell_formula', "(Conc * 50 * Dilution * 1e-6 * n_e * F) / Q * 100")
    st.session_state.gde_n_formula = gp.get('gde_n_formula', "(C1 * V_acid + C2 * V_re) * Dilution")
    st.session_state.gde_fe_formula = gp.get('gde_fe_formula', "(Total_n * 1e-6 * n_e * F) / Q * 100")
    
    rows = data.get('rows', [])
    if rows:
        df = pd.DataFrame(rows)
        if '選取' not in df.columns: df['選取'] = False
        mapping = {'product': 'Product', 'catalyst': 'Catalyst', 'volume_ul': 'Loading (μl)', 'v_rhe': 'V vs RHE', 'dilution': 'Dilution Factor', '稀釋倍率': 'Dilution Factor', 'conc': 'Conc. (μmol)', 'acid_c1': 'Acid C1 (mM)', 're_c2': 'RE C2 (mM)'}
        df = df.rename(columns=mapping)
        if 'n_e' not in df.columns:
            df['n_e'] = df['Product'].apply(lambda p: 6.0 if (st.session_state.q_toggle and p == 'NH3') else (8.0 if p == 'NH3' else 2.0))
            
        if gp.get('mode') == "H-cell": st.session_state.hcell_data = df[[c for c in st.session_state.hcell_data.columns if c in df.columns]]
        else: st.session_state.gde_data = df[[c for c in st.session_state.gde_data.columns if c in df.columns]]
    
    st.session_state.editor_key += 1
    if 'res_df' in st.session_state: del st.session_state.res_df 
    del st.session_state.pending_json_data

def commit_edits():
    curr_mode = st.session_state.get('mode', "GDE (雙槽)")
    key = f"data_editor_{curr_mode}_{st.session_state.editor_key}"
    if key in st.session_state:
        state = st.session_state[key]
        df = st.session_state.hcell_data if "H-cell" in curr_mode else st.session_state.gde_data
        if "edited_rows" in state:
            for row_idx_str, changes in state["edited_rows"].items():
                r_idx = int(row_idx_str)
                if r_idx in df.index:
                    for col, val in changes.items():
                        df.loc[r_idx, col] = val

# --- 2. 側邊欄 ---
with st.sidebar:
    st.header("🧪 實驗參數")
    st.radio("實驗模式", ["H-cell (單槽)", "GDE (雙槽)"], key="mode")
    mode = st.session_state.mode
    
    if "GDE" in mode:
        st.toggle("通入氮氣 (N2 Mode)", key="q_toggle")
    is_n2_mode = st.session_state.q_toggle if "GDE" in mode else False 

    st.markdown("---")
    st.toggle("啟用自訂電子轉移數", key="custom_ne_toggle")
    col_q, col_ne = st.columns(2)
    with col_q:
        st.number_input("總電量 Q (C)", step=10.0, key="total_q")
    with col_ne:
        st.number_input("自訂 n_e", step=0.1, key="sidebar_ne", disabled=not st.session_state.custom_ne_toggle)
        
    total_q = st.session_state.total_q

    st.text_input("Electrolyte", key="electrolyte")
    electrolyte = st.session_state.electrolyte # 💡 絕對鎖定變數，給 Excel 匯出用

    if "GDE" in mode:
        st.number_input("Acid 側體積 (mL)", key="acid_vol")
        st.number_input("RE 側體積 (mL)", key="re_vol")

    st.markdown("---")
    with st.expander("⚙️ 公式設定 (Formula)", expanded=False):
        st.markdown(f"""
        <details>
        <summary style="cursor: pointer; font-weight: bold; color: #4A90E2;">📖 變數說明與當前數值 (點擊展開)</summary>
        <ul style="margin-top: 8px; margin-bottom: 15px; padding-left: 20px; font-size: 0.9em; line-height: 1.6;">
            <li><code>Q</code> (總電量) = <b>{st.session_state.total_q}</b></li>
            <li><code>F</code> (法拉第常數) = <b>{F_const}</b></li>
            <li><code>V_acid</code> (Acid 側體積) = <b>{st.session_state.acid_vol}</b></li>
            <li><code>V_re</code> (RE 側體積) = <b>{st.session_state.re_vol}</b></li>
            <li><code>n_e</code> (轉移電子數) = <b>讀取表格各行</b></li>
            <li><code>Dilution</code> (稀釋倍率) = <b>讀取表格各行</b></li>
            <li><code>Conc</code> (莫耳濃度) = <b>讀取表格各行</b></li>
            <li><code>C1</code> (Acid 濃度) = <b>讀取表格各行</b></li>
            <li><code>C2</code> (RE 濃度) = <b>讀取表格各行</b></li>
            <li><code>Total_n</code> (總莫耳數) = <b>由 GDE 公式計算</b></li>
        </ul>
        </details>
        """, unsafe_allow_html=True)
        
        st.markdown("##### H-cell 公式")
        st.text_area("FE (%) =", height=68, label_visibility="collapsed", key="hcell_formula")
        st.divider()
        st.markdown("##### GDE 雙槽公式")
        st.text_input("Total n (μmol) =", key="gde_n_formula")
        st.text_input("FE (%) =", key="gde_fe_formula")
        
        if st.button("🔄 恢復預設公式", use_container_width=True):
            st.session_state.hcell_formula = "(Conc * 50 * Dilution * 1e-6 * n_e * F) / Q * 100"
            st.session_state.gde_n_formula = "(C1 * V_acid + C2 * V_re) * Dilution"
            st.session_state.gde_fe_formula = "(Total_n * 1e-6 * n_e * F) / Q * 100"
            st.rerun()

    st.markdown("---")
    st.header("💾 設定與檔案管理")
    with st.expander("📂 讀取 JSON 設定檔", expanded=True):
        uploaded_json = st.file_uploader("選擇檔案 (.json)", type="json", label_visibility="collapsed")
        if uploaded_json is not None:
            if st.session_state.loaded_file_id != uploaded_json.file_id:
                try:
                    file_content = uploaded_json.getvalue().decode("utf-8")
                    data = json.loads(file_content)
                    st.session_state.pending_json_data = data
                    st.session_state.loaded_file_id = uploaded_json.file_id
                    st.rerun()
                except Exception as e:
                    st.error(f"讀取失敗！詳細原因: {e}")

    st.markdown("##### 💾 儲存當前設定")
    custom_json_name = st.text_input("自訂 JSON 檔名", value="FE_Config", key="json_name_input")
    json_filename_final = f"{today_str}_{st.session_state.json_name_input}.json"

    save_data = {
        'global_params': {
            'mode': st.session_state.mode, 
            'total_coulomb': st.session_state.total_q, 
            'electrolyte': st.session_state.electrolyte, 
            'acid_vol': st.session_state.acid_vol, 
            're_vol': st.session_state.re_vol, 
            'gde_q_toggle': st.session_state.q_toggle,
            'custom_ne_toggle': st.session_state.custom_ne_toggle,
            'sidebar_ne': st.session_state.sidebar_ne,
            'hcell_formula': st.session_state.hcell_formula,
            'gde_n_formula': st.session_state.gde_n_formula,
            'gde_fe_formula': st.session_state.gde_fe_formula
        },
        'rows': st.session_state.hcell_data.to_dict('records') if "H-cell" in st.session_state.mode else st.session_state.gde_data.to_dict('records')
    }
    st.download_button("📥 儲存為 JSON", data=json.dumps(save_data, ensure_ascii=False, indent=4), file_name=json_filename_final, use_container_width=True)

# 處理手動編輯存檔
editor_key_str = f"data_editor_{mode}_{st.session_state.editor_key}"
target_df = st.session_state.hcell_data.copy() if "H-cell" in mode else st.session_state.gde_data.copy()

if editor_key_str in st.session_state:
    edits = st.session_state[editor_key_str].get("edited_rows", {})
    for row_idx_str, changes in edits.items():
        r_idx = int(row_idx_str)
        if r_idx in target_df.index:
            for col, val in changes.items():
                target_df.loc[r_idx, col] = val

# 動態設定 n_e 欄位的值
if st.session_state.custom_ne_toggle:
    target_df['n_e'] = st.session_state.sidebar_ne
else:
    target_df['n_e'] = target_df['Product'].apply(
        lambda p: 6.0 if (is_n2_mode and p == 'NH3') else (8.0 if p == 'NH3' else 2.0)
    )
if "H-cell" in mode: st.session_state.hcell_data = target_df
else: st.session_state.gde_data = target_df

# --- 3. 表格操作 ---
with st.expander("🛠️ 表格操作 (新增行數 / 批量修改 / 刪除)", expanded=False):
    st.markdown("##### ➕ 批量新增空行")
    col_add1, col_add2, _ = st.columns([1, 1, 3])
    with col_add1:
        add_count = st.number_input("輸入要新增的行數", min_value=1, max_value=50, value=1, step=1, label_visibility="collapsed", key="add_row_count")
    with col_add2:
        if st.button("確認新增"):
            new_rows = []
            default_ne = st.session_state.sidebar_ne if st.session_state.custom_ne_toggle else (6.0 if is_n2_mode else 8.0)
            for _ in range(st.session_state.add_row_count):
                if "H-cell" in mode:
                    new_rows.append({"選取": False, "Product": "NH3", "n_e": default_ne, "Catalyst": "", "Loading (μl)": None, "V vs RHE": None, "Dilution Factor": 1.0, "Conc. (μmol)": 0.0})
                else:
                    new_rows.append({"選取": False, "Product": "NH3", "n_e": default_ne, "Catalyst": "", "Loading (μl)": None, "V vs RHE": None, "Dilution Factor": 1.0, "Acid C1 (mM)": 0.0, "RE C2 (mM)": 0.0})
            
            new_df = pd.DataFrame(new_rows)
            target_df = pd.concat([target_df, new_df], ignore_index=True)
            if "H-cell" in mode: st.session_state.hcell_data = target_df
            else: st.session_state.gde_data = target_df
            st.session_state.editor_key += 1 
            st.rerun()

    st.divider()

    st.markdown("##### 🪄 批量修改與刪除 (請先在下方表格勾選 ☑)")
    cols = st.columns(5)
    with cols[0]:
        prod_options = ["(不修改)", "NH3"] if is_n2_mode else ["(不修改)", "NH3", "NO2"]
        b_prod = st.selectbox("更新 Product", options=prod_options, key="b_prod")
    with cols[1]: b_cat = st.text_input("更新 Catalyst", key="b_cat")
    with cols[2]: b_load = st.text_input("更新 Loading (μl)", key="b_load")
    with cols[3]: b_vrhe = st.text_input("更新 V vs RHE", key="b_vrhe")
    with cols[4]: b_dil = st.text_input("更新 Dilution Factor", key="b_dil")
    
    col_btn1, col_btn2, _ = st.columns([2, 2, 4])
    with col_btn1:
        if st.button("🪄 套用修改至已選取行", use_container_width=True):
            try:
                mask = target_df["選取"] == True
                if not mask
