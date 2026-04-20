import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# --- 1. 頁面與全域設定 ---
st.set_page_config(page_title="Faradaic Efficiency 計算機", layout="wide")
st.title("⚡ Faradaic Efficiency 數據計算機")

# 數學常數與預設公式
F_const = 96485
DEFAULT_HCELL_FORMULA = "(Conc * 50 * Dilution * 1e-6 * n_e * F) / Q * 100"
DEFAULT_GDE_N_FORMULA = "(C1 * V_acid + C2 * V_re) * Dilution"
DEFAULT_GDE_FE_FORMULA = "(Total_n * 1e-6 * n_e * F) / Q * 100"

# --- 2. 側邊欄：模式與參數設定 ---
st.sidebar.header("🧪 實驗參數設定")

# 模式切換
mode = st.sidebar.radio("實驗模式", ["H-cell (單槽)", "GDE (雙槽)"])

# 通氣模式切換
st.sidebar.markdown("---")
is_n2_mode = st.sidebar.toggle("通入氮氣 (N2 Mode)", value=False, help="開啟後，NH3 的轉移電子數將設為 6，且產物鎖定為 NH3")

st.sidebar.markdown("---")
# 共用參數
Q = st.sidebar.number_input("總電量 Q (C)", value=100.0, step=10.0)
electrolyte = st.sidebar.text_input("電解液", value="0.5 M KNO3 + 0.1 M KOH")

# 體積參數 (GDE 模式需要)
if "GDE" in mode:
    acid_vol = st.sidebar.number_input("Acid 側體積 (mL)", value=10.0, step=1.0)
    re_vol = st.sidebar.number_input("RE 側體積 (mL)", value=50.0, step=5.0)
else:
    # 針對 H-cell 隱藏在 UI 裡，但如果你公式裡有用到可以設置
    pass

st.sidebar.markdown("---")
st.sidebar.subheader("📐 自訂計算公式")
if "H-cell" in mode:
    formula_hcell = st.sidebar.text_input("H-cell FE (%)", value=DEFAULT_HCELL_FORMULA)
else:
    formula_gde_n = st.sidebar.text_input("GDE Total n (μmol)", value=DEFAULT_GDE_N_FORMULA)
    formula_gde_fe = st.sidebar.text_input("GDE FE (%)", value=DEFAULT_GDE_FE_FORMULA)

# --- 3. 核心數據表格設定 ---
# 初始化資料結構 (如果你希望預設帶入你常用的參數)
if 'hcell_data' not in st.session_state:
    st.session_state.hcell_data = pd.DataFrame({
        "Product": ["NH3"],
        "Catalyst": ["純鎳纖維紙"],
        "Loading (μl)": [80.0],
        "V vs RHE": [-0.3],
        "稀釋倍率": [1.0],
        "Conc. (μmol)": [0.0]
    })

if 'gde_data' not in st.session_state:
    st.session_state.gde_data = pd.DataFrame({
        "Product": ["NH3"],
        "Catalyst": ["純鎳纖維紙"],
        "Loading (μl)": [80.0],
        "V vs RHE": [-0.3],
        "稀釋倍率": [1.0],
        "Acid C1 (mM)": [0.0],
        "RE C2 (mM)": [0.0]
    })

st.subheader(f"📊 實驗數據輸入 ({mode})")
st.info("💡 提示：你可以直接點擊下方表格進行編輯，點擊最右側的「+」或「垃圾桶」圖示來新增/刪除行。")

# 根據模式顯示並編輯表格
if "H-cell" in mode:
    # 使用 st.data_editor 讓使用者能在網頁上直接像 Excel 一樣編輯
    edited_df = st.data_editor(
        st.session_state.hcell_data,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Product": st.column_config.SelectboxColumn(
                "產物",
                help="選擇反應產物",
                options=["NH3"] if is_n2_mode else ["NH3", "NO2"],
                required=True
            )
        }
    )
else:
    edited_df = st.data_editor(
        st.session_state.gde_data,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Product": st.column_config.SelectboxColumn(
                "產物",
                options=["NH3"] if is_n2_mode else ["NH3", "NO2"],
                required=True
            )
        }
    )

# --- 4. 進行計算 ---
if st.button("🔄 計算 FE", type="primary"):
    result_df = edited_df.copy()
    
    # 準備存放計算結果的 List
    fe_results = []
    total_n_results = []

    for index, row in result_df.iterrows():
        try:
            prod = row["Product"]
            
            # 判斷電子數 n_e
            if is_n2_mode and prod == 'NH3':
                n_e = 6
            else:
                n_e = 8 if prod == 'NH3' else (2 if prod == 'NO2' else np.nan)
            
            dilution = float(row["稀釋倍率"])
            
            if "H-cell" in mode:
                conc = float(row["Conc. (μmol)"])
                if not np.isnan(n_e):
                    # 建立變數環境給 eval 使用
                    env = {'Conc': conc, 'Dilution': dilution, 'Q': Q, 'n_e': n_e, 'F': F_const}
                    fe_val = eval(formula_hcell, {"__builtins__": {}}, env)
                    fe_results.append(round(fe_val, 2))
                else:
                    fe_results.append(None)
            
            else: # GDE 模式
                c1, c2 = float(row["Acid C1 (mM)"]), float(row["RE C2 (mM)"])
                env_n = {'C1': c1, 'C2': c2, 'V_acid': acid_vol, 'V_re': re_vol, 'Dilution': dilution}
                total_n = eval(formula_gde_n, {"__builtins__": {}}, env_n)
                total_n_results.append(round(total_n, 3))
                
                if not np.isnan(n_e):
                    env_fe = {'Total_n': total_n, 'Q': Q, 'n_e': n_e, 'F': F_const}
                    fe_val = eval(formula_gde_fe, {"__builtins__": {}}, env_fe)
                    fe_results.append(round(fe_val, 2))
                else:
                    fe_results.append(None)
                    
        except Exception as e:
            st.error(f"第 {index+1} 行計算錯誤: {e}")
            fe_results.append("Error")
            if "GDE" in mode: total_n_results.append("Error")

    # 將結果合併回 DataFrame
    if "GDE" in mode:
        result_df["Total n (μmol)"] = total_n_results
    result_df["FE (%)"] = fe_results
    
    # --- 5. 顯示結果與匯出 ---
    st.success("✅ 計算完成！")
    st.dataframe(result_df, use_container_width=True)
    
    # 匯出成 Excel
    def to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='FE_Data')
        processed_data = output.getvalue()
        return processed_data
        
    excel_data = to_excel(result_df)
    
    gas_str = "N2_Gas" if is_n2_mode else "Ar_Gas"
    file_name = f"FE_Result_{mode}_{gas_str}.xlsx"
    
    st.download_button(
        label="📥 下載 Excel 結果",
        data=excel_data,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
