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
# 固定計算公式
HCELL_FORMULA = "(Conc * 50 * Dilution * 1e-6 * n_e * F) / Q * 100"
GDE_N_FORMULA = "(C1 * V_acid + C2 * V_re) * Dilution"
GDE_FE_FORMULA = "(Total_n * 1e-6 * n_e * F) / Q * 100"

# 獲取當天日期 (格式如: 20240520)
today_str = datetime.date.today().strftime("%Y%m%d")

# --- 初始化 Session State ---
if 'editor_key' not in st.session_state: st.session_state.editor_key = 0
if 'loaded_file_id' not in st.session_state: st.session_state.loaded_file_id = ""

if 'mode' not in st.session_state: st.session_state.mode = "GDE (雙槽)"
if 'q_toggle' not in st.session_state: st.session_state.q_toggle = False
if 'total_q' not in st.session_state: st.session_state.total_q = 100.0
if 'electrolyte' not in st.session_state: st.session_state.electrolyte = "0.5 M KNO3 + 0.1 M KOH" 
if 'acid_vol' not in st.session_state: st.session_state.acid_vol = 10.0
if 're_vol' not in st.session_state: st.session_state.re_vol = 50.0

if 'hcell_data' not in st.session_state:
    st.session_state.hcell_data = pd.DataFrame({
        "選取": pd.Series(dtype='bool'),
        "Product": pd.Series(dtype='str'), "Catalyst": pd.Series(dtype='str'), "Loading (μl)": pd.Series(dtype='float'),
        "V vs RHE": pd.Series(dtype='float'), "Dilution Factor": pd.Series(dtype='float'), "Conc. (μmol)": pd.Series(dtype='float')
    })
if 'gde_data' not in st.session_state:
    st.session_state.gde_data = pd.DataFrame({
        "選取": pd.Series(dtype='bool'),
        "Product": pd.Series(dtype='str'), "Catalyst": pd.Series(dtype='str'), "Loading (μl)": pd.Series(dtype='float'),
        "V vs RHE": pd.Series(dtype='float'), "Dilution Factor": pd.Series(dtype='float'), "Acid C1 (mM)": pd.Series(dtype='float'), "RE C2 (mM)": pd.Series(dtype='float')
    })

if '選取' not in st.session_state.hcell_data.columns:
    st.session_state.hcell_data.insert(0, '選取', False)
if '選取' not in st.session_state.gde_data.columns:
    st.session_state.gde_data.insert(0, '選取', False)

# --- 2. 側邊欄：設定管理與實驗參數 ---
with st.sidebar:
    st.header("💾 設定與檔案管理")
    with st.expander("JSON 設定檔操作", expanded=True):
        uploaded_json = st.file_uploader("📂 讀取設定 (.json)", type="json")
        if uploaded_json is not None:
            if st.session_state.loaded_file_id != uploaded_json.file_id:
                try:
                    # 💡 核心修復：使用更安全的字串讀取方式，避免檔案指標遺失
                    file_content = uploaded_json.getvalue().decode("utf-8")
                    data = json.loads(file_content)
                    
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
                        if '選取' not in df.columns: df['選取'] = False
                        
                        mapping = {'product': 'Product', 'catalyst': 'Catalyst', 'volume_ul': 'Loading (μl)', 'v_rhe': 'V vs RHE', 'dilution': 'Dilution Factor', '稀釋倍率': 'Dilution Factor', 'conc': 'Conc. (μmol)', 'acid_c1': 'Acid C1 (mM)', 're_c2': 'RE C2 (mM)'}
                        df = df.rename(columns=mapping)
                        if gp.get('mode') == "H-cell": st.session_state.hcell_data = df[[c for c in st.session_state.hcell_data.columns if c in df.columns]]
                        else: st.session_state.gde_data = df[[c for c in st.session_state.gde_data.columns if c in df.columns]]
                    
                    st.session_state.loaded_file_id = uploaded_json.file_id
                    st.session_state.editor_key += 1
                    st.rerun()
                except Exception as e:
                    # 💡 核心修復：如果失敗，印出真正的紅色錯誤代碼
                    st.error(f"讀取失敗！詳細原因: {e}")

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
    
    electrolyte = st.text_input("Electrolyte", value=st.session_state.electrolyte)
    st.session_state.electrolyte = electrolyte

    if "GDE" in mode:
        st.session_state.acid_vol = st.number_input("Acid 側體積 (mL)", value=float(st.session_state.acid_vol))
        st.session_state.re_vol = st.number_input("RE 側體積 (mL)", value=float(st.session_state.re_vol))

# 處理按鈕之前，先捕捉表格內「來不及存檔」的手動編輯
editor_key_str = f"data_editor_{mode}_{st.session_state.editor_key}"
target_df = st.session_state.hcell_data.copy() if "H-cell" in mode else st.session_state.gde_data.copy()

if editor_key_str in st.session_state:
    edits = st.session_state[editor_key_str].get("edited_rows", {})
    for row_idx_str, changes in edits.items():
        r_idx = int(row_idx_str)
        if r_idx in target_df.index:
            for col, val in changes.items():
                target_df.loc[r_idx, col] = val

# --- 3. 表格操作 ---
with st.expander("🛠️ 表格操作 (新增行數 / 批量修改 / 刪除)", expanded=False):
    st.markdown("##### ➕ 批量新增空行")
    col_add1, col_add2, _ = st.columns([1, 1, 3])
    with col_add1:
        add_count = st.number_input("輸入要新增的行數", min_value=1, max_value=50, value=1, step=1, label_visibility="collapsed")
    with col_add2:
        if st.button("確認新增"):
            new_rows = []
            for _ in range(add_count):
                if "H-cell" in mode:
                    new_rows.append({"選取": False, "Product": "NH3", "Catalyst": "", "Loading (μl)": None, "V vs RHE": None, "Dilution Factor": 1.0, "Conc. (μmol)": None})
                else:
                    new_rows.append({"選取": False, "Product": "NH3", "Catalyst": "", "Loading (μl)": None, "V vs RHE": None, "Dilution Factor": 1.0, "Acid C1 (mM)": None, "RE C2 (mM)": None})
            
            new_df = pd.DataFrame(new_rows)
            target_df = pd.concat([target_df, new_df], ignore_index=True)
            if "H-cell" in mode: st.session_state.hcell_data = target_df
            else: st.session_state.gde_data = target_df
            st.session_state.editor_key += 1 
            st.rerun()

    st.divider()

    st.markdown("##### 🪄 批量修改與刪除 (請先在下方表格勾選 ☑)")
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1: 
        prod_options = ["(不修改)", "NH3"] if is_n2_mode else ["(不修改)", "NH3", "NO2"]
        b_prod = st.selectbox("更新 Product", options=prod_options)
    with col2: b_cat = st.text_input("更新 Catalyst")
    with col3: b_load = st.text_input("更新 Loading (μl)")
    with col4: b_vrhe = st.text_input("更新 V vs RHE")
    with col5: b_dil = st.text_input("更新 Dilution Factor")
    
    col_btn1, col_btn2, _ = st.columns([2, 2, 4])
    with col_btn1:
        if st.button("🪄 套用修改至已選取行", use_container_width=True):
            try:
                mask = target_df["選取"] == True
                if not mask.any():
                    st.warning("⚠️ 請先在下方表格左側勾選 (☑) 您要修改的行！")
                else:
                    if b_prod != "(不修改)": target_df.loc[mask, "Product"] = b_prod
                    if b_cat: target_df.loc[mask, "Catalyst"] = b_cat
                    if b_load: target_df.loc[mask, "Loading (μl)"] = float(b_load)
                    if b_vrhe: target_df.loc[mask, "V vs RHE"] = float(b_vrhe)
                    if b_dil: target_df.loc[mask, "Dilution Factor"] = float(b_dil)
                    
                    target_df["選取"] = False
                    if "H-cell" in mode: st.session_state.hcell_data = target_df
                    else: st.session_state.gde_data = target_df
                    st.session_state.editor_key += 1
                    st.rerun()
            except ValueError:
                st.error("數值格式輸入錯誤！請確保 Loading, V vs RHE, Dilution Factor 欄位中輸入的是數字。")

    with col_btn2:
        if st.button("❌ 刪除已選取行", use_container_width=True):
            try:
                mask = target_df["選取"] == True
                if not mask.any():
                    st.warning("⚠️ 請先在下方表格左側勾選 (☑) 您要刪除的行！")
                else:
                    target_df = target_df[~mask].reset_index(drop=True)
                    if "H-cell" in mode: st.session_state.hcell_data = target_df
                    else: st.session_state.gde_data = target_df
                    st.session_state.editor_key += 1
                    st.rerun()
            except Exception as e:
                st.error(f"刪除失敗: {e}")

# --- 4. 數據表格 ---
st.subheader(f"📊 數據輸入 - {mode}")

col_sel1, col_sel2, _ = st.columns([1, 1, 8])
with col_sel1:
    if st.button("☑ 全選", use_container_width=True):
        target_df["選取"] = True
        if "H-cell" in mode: st.session_state.hcell_data = target_df
        else: st.session_state.gde_data = target_df
        st.session_state.editor_key += 1
        st.rerun()
with col_sel2:
    if st.button("☐ 全取消", use_container_width=True):
        target_df["選取"] = False
        if "H-cell" in mode: st.session_state.hcell_data = target_df
        else: st.session_state.gde_data = target_df
        st.session_state.editor_key += 1
        st.rerun()

base_render_df = st.session_state.hcell_data if "H-cell" in mode else st.session_state.gde_data

edited_df = st.data_editor(
    base_render_df,
    key=editor_key_str,
    num_rows="fixed",
    use_container_width=True,
    hide_index=True,
    column_config={
        "選取": st.column_config.CheckboxColumn("☑ 選取", default=False, width="small"),
        "Product": st.column_config.SelectboxColumn("Product", options=["NH3"] if is_n2_mode else ["NH3", "NO2"], required=True)
    }
)

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
            dil = float(row["Dilution Factor"])
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
    
    st.dataframe(res_df.drop(columns=["選取"], errors='ignore'), use_container_width=True)

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
                    'Dilution Factor': row['Dilution Factor'], 'V vs RHE': row['V vs RHE'], 'Total Concentration (μmol)': row['Conc. (μmol)'],
                    'Faradaic Efficiency (%)': row['FE (%)']
                })
            else:
                export_data.append({
                    'Cell': f"GDE_{gas_mode_str}({row['Product']})", 'Electrolyte': curr_electrolyte, 'Total Coulomb (Q)': curr_Q,
                    'Product Type': row['Product'], 'Catalyst': row['Catalyst'], 'Loading (μl)': row['Loading (μl)'],
                    'Dilution Factor': row['Dilution Factor'], 'V vs RHE': row['V vs RHE'], 'Acid C1 (mM)': row['Acid C1 (mM)'], 
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
