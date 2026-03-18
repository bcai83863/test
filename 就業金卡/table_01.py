import re
from pathlib import Path
from datetime import datetime, date
from typing import Tuple, Union

import pandas as pd
import streamlit as st 
from font_utils import apply_streamlit_cjk_css

# =========================================================
# 1) Excel 讀取與快取
# =========================================================
@st.cache_data
def read_excel_cached(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    engine = "openpyxl" if path.suffix.lower() == ".xlsx" else "xlrd"
    df = pd.read_excel(path, sheet_name=0, engine=engine)
    df.columns = [str(c).strip() for c in df.columns]
    return df

# =========================================================
# 2) 日期與數值解析
# =========================================================
YM_FULL_RE = re.compile(r"^\s*(\d{2,4})\s*[\/\-.年]\s*(\d{1,2})\s*(?:月)?\s*$")

def parse_ym_strict(x):
    if pd.isna(x): return pd.NaT
    if isinstance(x, (datetime, pd.Timestamp, date)):
        return pd.Timestamp(x.year, x.month, 1)

    s = str(x).replace("\u00a0", "").replace("\u3000", "").strip()
    if any(tok in s for tok in ["～", "〜", "~", "至"]): return pd.NaT
    if re.fullmatch(r"\d{6}", s): return pd.to_datetime(s, format="%Y%m", errors="coerce")
    m = YM_FULL_RE.match(s)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if y < 1900: y += 1911
        if 1 <= mo <= 12: return pd.Timestamp(y, mo, 1)
    return pd.NaT

def resolve_cutoff_date(dates: pd.Series, requested: str | None) -> Tuple[pd.Timestamp, str]:
    valid_dates = dates.dropna()
    if valid_dates.empty: raise RuntimeError("Excel 內找不到有效的『統計年月』")

    max_dt = pd.to_datetime(valid_dates, errors="coerce").max()
    max_dt = pd.Timestamp(max_dt.year, max_dt.month, 1)

    target_dt = max_dt
    if requested:
        req = parse_ym_strict(requested)
        if not pd.isna(req):
            target_dt = min(req, max_dt) # 不允許超過資料現有最大日期

    # ✨ 關鍵修復：確保傳回的用於顯示的字串格式正確
    used_ym_str = f"{target_dt.year}/{int(target_dt.month):02d}"
    return target_dt, used_ym_str

# =========================================================
# 3) 資料樞紐邏輯
# =========================================================
def build_table1_data(df: pd.DataFrame, requested: str | None) -> Tuple[pd.DataFrame, str]:
    # 尋找「總計」列所在的維度欄
    dim_col = next((c for c in df.columns if c not in ["統計年月", "總計"] 
                    and (df[c].astype(str).str.strip() == "總計").any()), None)
    
    if not dim_col: raise RuntimeError("找不到維度欄位來定位『總計』列")

    df["統計年月"] = df["統計年月"].ffill()
    df["date"] = df["統計年月"].apply(parse_ym_strict)
    
    cutoff_date, used_ym_str = resolve_cutoff_date(df["date"], requested)

    df_filtered = df[(df["date"] <= cutoff_date) & (df[dim_col].astype(str).str.strip() == "總計")].copy()
    df_filtered["year"] = df_filtered["date"].dt.year
    df_filtered["month"] = df_filtered["date"].dt.month
    df_filtered["val"] = pd.to_numeric(df_filtered["總計"].astype(str).str.replace(",", ""), errors='coerce').fillna(0)

    # 樞紐分析
    piv = df_filtered.pivot_table(index="month", columns="year", values="val", aggfunc="sum").reindex(range(1, 13))
    years = sorted(piv.columns)

    out_rows = []
    for m in range(1, 13):
        row_data = {"時間": f"{m}月"}
        for y in years:
            val = piv.loc[m, y]
            # ✨ 格式化數值，處理 NaN
            row_data[f"{int(y)}年"] = "-" if pd.isna(val) or val == 0 else f"{int(val):,}"
        out_rows.append(row_data)

    annual_row = {"時間": "當年度核發數"}
    total_row = {"時間": "累計核發數"}
    running_total = 0

    for y in years:
        annual_sum = piv[y].sum()
        running_total += annual_sum
        annual_row[f"{int(y)}年"] = f"{int(annual_sum):,}"
        total_row[f"{int(y)}年"] = f"{int(running_total):,}"

    out_rows.append(annual_row)
    out_rows.append(total_row)

    return pd.DataFrame(out_rows), used_ym_str

# =========================================================
# 4) ✨ Streamlit 渲染
# =========================================================
def render_streamlit(data_dir: Path):
    apply_streamlit_cjk_css()
    st.subheader("📊 表1：歷年累計許可人次")

    requested = st.text_input("您可以手動輸入截止月份 (例：2025/11)，若留白則自動抓取最新資料：",
                              placeholder="例如：2025/11", key="table1_input")
    requested = requested.strip() or None

    try:
        excel_files = list(data_dir.glob("*.xlsx")) + list(data_dir.glob("*.xls"))
        candidates = [p for p in excel_files if ("累計核發" in p.name) and ("初次申請" not in p.name)]

        if not candidates:
            st.error("找不到符合規格的來源檔 (需含『累計核發』且不含『初次申請』)")
            return

        # 讀取最新的檔案
        latest_file = sorted(candidates, key=lambda x: x.name, reverse=True)[0]
        df_raw = read_excel_cached(str(latest_file))

        # 執行原本的樞紐分析與計算
        table1_df, used_ym_str = build_table1_data(df_raw, requested)

        y, m = used_ym_str.split("/")
        st.info(f"📅 目前顯示資料時間：截至 {y}年{int(m)}月")

        # ✨ 新增：下載按鈕
        csv = table1_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 下載此報表 (CSV)",
            data=csv,
            file_name=f"表1_歷年累計許可人次_{y}{int(m):02d}.csv",
            mime="text/csv",
        )

        # ✨ 顯示美化後的表格
        st.dataframe(table1_df, hide_index=True, use_container_width=True)

    except Exception as e:
        st.error(f"讀取或處理資料時發生錯誤：{e}")