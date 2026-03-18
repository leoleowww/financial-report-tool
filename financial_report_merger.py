import streamlit as st
import pandas as pd
import io
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

st.set_page_config(page_title="出帳合併系統", layout="wide")
st.title("📂 出帳合併系統")
st.info("上傳Excel檔，並設定搜尋範圍，即可下載合併檔案。")

# --- SIDEBAR ---
st.sidebar.header("設定")
start_date = st.sidebar.text_input("查詢開始日 (YYYMMDD)", value="1150101")
end_date = st.sidebar.text_input("查詢結束日 (YYYMMDD)", value="1151231")
output_filename = st.sidebar.text_input("存檔檔名", value="合併結果")

# --- UPLOAD ---
uploaded_files = st.file_uploader(
    "選擇或拖曳檔案至此",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    all_dfs = []

    with st.spinner('處理中...'):
        for uploaded_file in uploaded_files:
            try:
                df = pd.read_excel(uploaded_file, sheet_name="主持人－計畫簡稱（明細）")
                df.columns = df.iloc[2]
                df = df.iloc[3:].reset_index(drop=True)
                df = df.iloc[1:, :10]
                all_dfs.append(df)  # ✅ THIS LINE WAS MISSING

            except Exception as e:
                st.error(f"處理 {uploaded_file.name} 時發生錯誤: {e}")

    if all_dfs:
        combined_df = pd.concat(all_dfs, ignore_index=True)

        def filter_data(df, col, start, end):
            if col in df.columns:
                col_str = df[col].astype(str).str.strip()
                return df[(col_str >= str(start)) & (col_str <= str(end))]
            else:
                st.warning(f"Column '{col}' not found.")
                return df

        matches = filter_data(combined_df, "請購日期", start_date, end_date)

        st.success(f"已找到範圍內 {len(matches)} 筆資料")
        st.dataframe(matches.head(50))

        # --- DOWNLOAD ---
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            matches.to_excel(writer, index=False, sheet_name='合併結果')

        st.download_button(
            label="💾 下載合併檔案",
            data=buffer.getvalue(),
            file_name=output_filename if output_filename.endswith('.xlsx') else f"{output_filename}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("所有檔案處理失敗，請確認工作表名稱是否為「主持人－計畫簡稱（明細）」")
else:
    st.write("請至少上傳一個Excel檔")