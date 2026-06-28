import streamlit as st
import pandas as pd
import json

STORE_ID_NAMES = ["店代碼", "店代", "門店代號", "門店代碼", "門市代號", "門市代碼", "店號"]
DAILY_METRIC_NAMES = [
    "預估業績達成率", "當日實際業績", "班表_年曆工時", "班表_實際時數", 
    "班表_高低預估", "生產力_標準", "生產力_實際", "生產力_高低標準", 
    "標班工時", "超額工時", "實際營業額標工", "工時差異", "備註"
]

def clean_date_value(val):
    if pd.isna(val) or val == "" or "Unnamed" in str(val): return pd.NaT
    try:
        num = float(val)
        if num > 40000: return pd.to_datetime(num, unit='D', origin='1899-12-30')
    except: pass
    dt = pd.to_datetime(val, errors='coerce')
    return dt if dt is not pd.NaT and dt.year >= 2020 else pd.NaT

def transform_to_perfect_format(uploaded_file, selected_ids, keep_mode):
    # 讀取上傳的記憶體內檔案
    full_df = pd.read_excel(uploaded_file, header=None)
    header_idx, id_col_idx = -1, -1
    for i, row in full_df.head(20).iterrows():
        for c_idx, val in enumerate(row):
            if any(name in str(val) for name in STORE_ID_NAMES):
                header_idx, id_col_idx = i, c_idx
                break
        if header_idx != -1: break

    if header_idx == -1:
        return None

    date_row_idx = -1
    for r_idx in range(max(0, header_idx-5), header_idx):
        if sum(1 for v in full_df.iloc[r_idx, :] if clean_date_value(v) is not pd.NaT) > 10:
            date_row_idx = r_idx
            break

    dates_row_values = full_df.iloc[date_row_idx, :].values
    start_date_col = -1
    for c_idx, val in enumerate(dates_row_values):
        if clean_date_value(val) is not pd.NaT:
            start_date_col = c_idx
            break
    
    df_data = pd.read_excel(uploaded_file, header=header_idx)
    df_data.iloc[:, id_col_idx] = df_data.iloc[:, id_col_idx].astype(str).str.strip()
    sel_set = set(str(sid) for sid in selected_ids)
    if keep_mode:
        df_data = df_data[df_data.iloc[:, id_col_idx].isin(sel_set)]
    else:
        df_data = df_data[~df_data.iloc[:, id_col_idx].isin(sel_set)]

    all_days_data = []
    base_info = df_data.iloc[:, 0:5] 
    
    curr_col = start_date_col
    while curr_col + 12 < len(df_data.columns):
        raw_date_val = clean_date_value(dates_row_values[curr_col])
        if raw_date_val is pd.NaT:
            for back_idx in range(curr_col, start_date_col-1, -1):
                dt = clean_date_value(dates_row_values[back_idx])
                if dt is not pd.NaT: raw_date_val = dt; break
        
        if raw_date_val is not pd.NaT:
            day_block = df_data.iloc[:, curr_col : curr_col + 13].copy()
            day_block.columns = DAILY_METRIC_NAMES
            
            day_block["預估業績達成率"] = pd.to_numeric(day_block["預估業績達成率"], errors='coerce')
            day_block["預估業績達成率"] = day_block["預估業績達成率"].apply(
                lambda x: f"{int(round(x * 100))}%" if pd.notnull(x) else ""
            )
            day_block["生產力_標準"] = pd.to_numeric(day_block["生產力_標準"], errors='coerce').round(0).fillna(0).astype(int)
            
            temp_combined = base_info.copy()
            temp_combined.insert(0, '星期', raw_date_val.strftime('%A'))
            temp_combined.insert(0, '日期', raw_date_val.strftime('%Y-%m-%d'))
            week_map = {'Monday':'週一','Tuesday':'週二','Wednesday':'週三','Thursday':'週四','Friday':'週五','Saturday':'週六','Sunday':'週日'}
            temp_combined['星期'] = temp_combined['星期'].map(week_map)
            
            all_days_data.append(pd.concat([temp_combined.reset_index(drop=True), day_block.reset_index(drop=True)], axis=1))
        curr_col += 13

    if all_days_data:
        final_df = pd.concat(all_days_data, axis=0, ignore_index=True)
        final_df = final_df.dropna(subset=[df_data.columns[id_col_idx]])
        return final_df
    return None

# --- Streamlit 手機友善網頁介面 ---
st.set_page_config(page_title="Excel 工具 v4.7 (行動優化版)", layout="centered")
st.title("📊 Excel 格式自動優化工具")

# 1. 檔案上傳
uploaded_file = st.file_uploader("1. 請上傳 Excel 檔案", type=["xlsx"])

if uploaded_file:
    # 讀取店號清單供使用者選擇
    try:
        temp_df = pd.read_excel(uploaded_file, header=None, nrows=15)
        h_idx, c_idx = -1, -1
        for i, row in temp_df.iterrows():
            for j, v in enumerate(row):
                if any(n in str(v) for n in STORE_ID_NAMES):
                    h_idx, c_idx = i, j; break
            if h_idx != -1: break
        
        df = pd.read_excel(uploaded_file, header=h_idx)
        store_ids_all = sorted(df.iloc[:, c_idx].dropna().unique().astype(str))
        
        # 2. 功能設定
        st.write("---")
        keep_mode_str = st.radio("篩選模式：", ["勾選＝保留店代", "勾選＝排除店代"])
        keep_mode = True if "保留" in keep_mode_str else False
        
        # JSON 配置導入/導出 (轉換為網頁文字框複製貼上)
        with st.expander("⚙️ JSON 快速配置"):
            json_input = st.text_area("在此貼上舊的 JSON 內容來自動勾選 (或留空)：")
            preset_ids = []
            if json_input.strip():
                try:
                    preset_ids = json.loads(json_input).get("selected_ids", [])
                    st.success("JSON 讀取成功！")
                except:
                    st.error("JSON 格式有誤。")

        # 3. 店號選擇器 (支援關鍵字搜尋與多選)
        selected = st.multiselect(
            "2. 請選擇或搜尋店號：", 
            options=store_ids_all, 
            default=[x for x in preset_ids if x in store_ids_all]
        )
        
        # 顯示當前已選擇的 JSON，方便複製備份
        if selected:
            st.code(json.dumps({"selected_ids": selected}, ensure_ascii=False), language="json")

        # 4. 開始處理與下載
        st.write("---")
        if st.button("🚀 執行篩選並生成優化檔案", type="primary", use_container_width=True):
            with st.spinner("資料處理中，請稍候..."):
                output_df = transform_to_perfect_format(uploaded_file, selected, keep_mode)
                if output_df is not None:
                    # 將 DataFrame 轉為 Excel 二進位資料提供下載
                    import io
                    towrite = io.BytesIO()
                    output_df.to_excel(towrite, index=False, engine='openpyxl')
                    towrite.seek(0)
                    
                    st.success("✨ 處理成功！")
                    st.download_button(
                        label="📥 點擊下載「格式優化版」Excel",
                        data=towrite,
                        file_name="格式優化版.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.error("處理失敗，找不到對應的店代碼欄位。")
    except Exception as e:
        st.error(f"讀取檔案時發生錯誤: {e}")