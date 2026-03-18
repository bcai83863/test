import re
from pathlib import Path
from typing import List

import pandas as pd
import streamlit as st # ✨ 新增 Streamlit 套件
from font_utils import apply_streamlit_cjk_css

# =========================================================
# 1) 來源檔自動帶入 (挑選：*有效港澳)
# =========================================================
def pick_latest_source_file(base: Path) -> Path:
    FILE_PATTERNS = ["*有效港澳.xlsx", "*有效港澳.xls"]
    FILENAME_YMD_RE = re.compile(r"(20\d{2})(0[1-9]|1[0-2])([0-3]\d).*有效港澳", re.IGNORECASE)
    
    candidates = []
    for pat in FILE_PATTERNS:
        candidates.extend(base.glob(pat))

    if not candidates:
        raise FileNotFoundError(f"找不到符合『有效港澳』格式的 Excel 檔案")

    with_ymd = []
    for p in candidates:
        m = FILENAME_YMD_RE.search(p.name)
        if m:
            ymd = int(m.group(1) + m.group(2) + m.group(3))
            with_ymd.append((ymd, p))

    if with_ymd:
        with_ymd.sort(key=lambda x: x[0], reverse=True)
        return with_ymd[0][1]

    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[0]

# =========================================================
# 2) 數據解析工具
# =========================================================
def to_num(series):
    return pd.to_numeric(series.astype(str).str.replace(",", "").replace({"-": "0", "nan": "0"}), errors='coerce').fillna(0)

@st.cache_data # ✨ 加上快取機制
def load_and_build_table8(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    engine = "openpyxl" if path.suffix == ".xlsx" else "xlrd"
    raw = pd.read_excel(path, header=None, engine=engine)
    
    # 定位表頭
    keys = ["統計區間", "外籍人士", "港澳人士", "總計"]
    h_row = 0
    for i in range(min(30, len(raw))):
        row_str = " ".join(raw.iloc[i].astype(str).fillna(""))
        if all(k in row_str for k in keys):
            h_row = i
            break
    
    df = raw.iloc[h_row+1:].copy()
    df.columns = raw.iloc[h_row].astype(str).str.strip().str.replace("\n", "")
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
    
    # 解析日期 (格式：2025年11月)
    YM_RE = re.compile(r"(20\d{2})\s*年\s*(\d{1,2})\s*月")
    parsed_dates = []
    for val in df["統計區間"]:
        m = YM_RE.search(str(val))
        if m:
            parsed_dates.append(pd.Timestamp(int(m.group(1)), int(m.group(2)), 1))
        else:
            parsed_dates.append(pd.NaT)
    
    df["date"] = parsed_dates
    df = df.dropna(subset=["date"]).sort_values("date").reset_index(drop=True)
    
    max_year = df["date"].dt.year.max()
    
    # 篩選邏輯：舊年份取該年最大月份，最新年份取全月
    d_old = df[df["date"].dt.year < max_year].copy()
    if not d_old.empty:
        d_old = d_old.loc[d_old.groupby(d_old["date"].dt.year)["date"].idxmax()]
        
    d_new = df[df["date"].dt.year == max_year].copy()
    final_df = pd.concat([d_old, d_new], ignore_index=True)
    
    # 格式化輸出
    labels = []
    for _, r in final_df.iterrows():
        y, m = r["date"].year, r["date"].month
        labels.append(f"{y}年" if y < max_year else f"{y}年{m}月")
    
    output = pd.DataFrame({
        "年/月底": labels,
        "外籍人士": final_df["外籍人士"].apply(lambda x: f"{int(to_num(pd.Series([x])).iloc[0]):,}"),
        "港澳人士": final_df["港澳人士"].apply(lambda x: f"{int(to_num(pd.Series([x])).iloc[0]):,}"),
        "總計": final_df["總計"].apply(lambda x: f"{int(to_num(pd.Series([x])).iloc[0]):,}")
    })
    
    return output

# =========================================================
# 3) ✨ Streamlit 專屬渲染函式 (入口)
# =========================================================
def render_streamlit(data_dir: Path):
    apply_streamlit_cjk_css()
    st.subheader("📊 表8：歷年有效許可人次 (外籍、港澳)")

    try:
        source_file = pick_latest_source_file(data_dir)
        # 顯示使用的檔案資訊
        m = re.search(r"(20\d{2})(0[1-9]|1[0-2])([0-3]\d)", source_file.stem)
        tag = m.group(0) if m else "最新"
        st.markdown(f"**目前顯示資料檔案日期：{tag}**")

        # 執行數據解析
        table8_df = load_and_build_table8(str(source_file))
        
        # ✨ 渲染網頁表格
        # 使用 use_container_width=True 讓表格填滿畫面
        st.dataframe(table8_df, hide_index=True, width="stretch")
        
        st.info("💡 註：舊年份顯示該年底數據，最新年份則列出所有月份明細。")

    except Exception as e:
        st.error(f"讀取或處理資料時發生錯誤：{e}")