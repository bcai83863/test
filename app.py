import streamlit as st
import pandas as pd
import plotly.express as px

# 設定網頁基本狀態
st.set_page_config(page_title="資料視覺化儀表板", layout="wide")

st.title("📊 專案資料視覺化儀表板")
st.markdown("這是一個結合 Pandas 資料處理與 Plotly 互動圖表的基礎**範例**。")

# 建立模擬的資料集 (可替換為實際的 CSV 或資料庫來源)
data = {
    "月份": ["一月", "二月", "三月", "四月", "五月", "六月"],
    "資料處理量 (萬筆)": [120, 150, 180, 130, 210, 250],
    "活躍度指標": [85, 88, 92, 87, 95, 98]
}
df = pd.DataFrame(data)

# 將畫面分為左右兩欄，提升視覺品質
col1, col2 = st.columns(2)

with col1:
    st.subheader("📋 原始資料表")
    # 顯示互動式資料表
    st.dataframe(df, use_container_width=True)

with col2:
    st.subheader("📈 處理量趨勢圖")
    # 使用 Plotly 繪製互動式折線圖
    fig = px.line(
        df, 
        x="月份", 
        y="資料處理量 (萬筆)", 
        markers=True, 
        title="上半年度趨勢"
    )
    st.plotly_chart(fig, use_container_width=True)
