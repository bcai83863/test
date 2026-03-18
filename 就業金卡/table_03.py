import re
from pathlib import Path
from datetime import datetime, date
from typing import Tuple

import pandas as pd
import streamlit as st # ✨ 新增 Streamlit 套件
from font_utils import apply_streamlit_cjk_css

# =========================================================
# 0) 參數設定
# =========================================================
# 年份範圍：表 3 固定顯示範圍
YEARS = list(range(2018, 2027))

# 領域順序
DOMAIN_ORDER = ["經濟","科技","教育","數位","金融","文化、藝術","專案會商","建築設計","國防","運動","法律","環境","生技"]

# =========================================================
# 1) 來源檔自動搜尋
# =========================================================
FILE_PATTERNS = ["*累計核發-領域.xlsx", "*累計核發-領域.xls"]
FILENAME_YM_RE = re.compile(r"(\d{6})-(\d{6})累計核發-領域\.(xlsx|xls)$", re.IGNORECASE)

def pick_latest_source_file(base: Path) -> Path:
    candidates = []
    for pat in FILE_PATTERNS:
        candidates.extend(base.glob(pat))

    if not candidates:
        raise FileNotFoundError(f"找不到『累計核發-領域』來源檔")

    with_ym = []
    for p in candidates:
        m = FILENAME_YM_RE.search(p.name)
        if m:
            with_ym.append((int(m.group(2)), p))

    if with_ym:
        with_ym.sort(key=lambda x: x[0], reverse=True)
        return with_ym[0][1]

    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[0]

# =========================================================
# 2) 日期與資料清洗
# =========================================================
YM_FULL_RE = re.compile(r"^\s*(\d{2,4})\s*[\/\-.年]\s*(\d{1,2})\s*(?:月)?\s*$")

def parse_ym_strict(x):
    if pd.isna(x): return pd.NaT
    if isinstance(x, (datetime, pd.Timestamp, date)): return pd.Timestamp(x.year, x.month, 1)
    s = str(x).replace("\u00a0", "").replace("\u3000", "").strip()
    if any(tok in s for tok in ["～", "〜", "~", "至"]): return pd.NaT
    if re.fullmatch(r"\d{6}", s): return pd.to_datetime(s, format="%Y%m", errors="coerce")
    m = YM_FULL_RE.match(s)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if y < 1900: y += 1911
        return pd.Timestamp(y, mo, 1)
    return pd.NaT

def to_num(series):
    return pd.to_numeric(series.astype(str).str.replace(",", "").replace({"-": "0", "nan": "0"}), errors='coerce').fillna(0)

@st.cache_data # ✨ 加上快取機制
def read_and_clean_excel(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    engine = "openpyxl" if path.suffix.lower() == ".xlsx" else "xlrd"
    raw_xls = pd.read_excel(path, header=None, engine=engine)
    
    header_row = 0
    for i in range(min(50, len(raw_xls))):
        row_str = " ".join(raw_xls.iloc[i].astype(str))
        if "統計年月" in row_str and "男" in row_str:
            header_row = i
            break
            
    df = raw_xls.iloc[header_row+1:].copy()
    df.columns = raw_xls.iloc[header_row].astype(str).str.strip()
    return df

# =========================================================
# 3) 表 3 核心邏輯
# =========================================================
def build_table3_data(df: pd.DataFrame, requested_ym: str | None):
    df.columns = [str(c).strip() for c in df.columns]
    
    # 找領域欄
    ignore = {"統計年月", "男", "女", "總計"}
    domain_col = next((c for c in df.columns if c not in ignore), df.columns[1])
    
    df["統計年月"] = df["統計年月"].ffill()
    df["date"] = df["統計年月"].apply(parse_ym_strict)
    
    valid_dates = df["date"].dropna()
    cutoff = valid_dates.max()
    if requested_ym:
        req = parse_ym_strict(requested_ym)
        if not pd.isna(req): cutoff = min(req, cutoff)

    # 過濾資料
    d = df[df["date"] <= cutoff].copy()
    d[domain_col] = d[domain_col].astype(str).str.strip()
    d = d[~d[domain_col].isin(["", "總計", "合計"])].copy()
    
    d["val"] = to_num(d["男"]) + to_num(d["女"])
    d["year"] = d["date"].dt.year
    
    # 樞紐分析
    # 使用 pivot_table 時，若遇到重複索引，預設會計算平均，這裡我們確保用 sum
    piv = d.pivot_table(index=domain_col, columns="year", values="val", aggfunc="sum")
    
    # 確保所有領域與年份都存在
    piv = piv.reindex(index=DOMAIN_ORDER, columns=YEARS).fillna(0)
    
    # 計算合計
    piv["累計"] = piv.sum(axis=1)
    total_sum = piv["累計"].sum()
    piv["占比"] = piv["累計"] / (total_sum if total_sum > 0 else 1)
    
    # 加入總計列
    total_row = piv.sum(axis=0)
    total_row.name = "總計"
    total_row["占比"] = 1.0
    
    # 將 total_row 轉換為 DataFrame 再與 piv 合併
    total_df = pd.DataFrame([total_row])
    total_df.index = ["總計"]
    piv = pd.concat([piv, total_df])
    
    # 格式化輸出
    rows = []
    for domain, data in piv.iterrows():
        entry = {"領域別": domain}
        for y in YEARS:
            entry[f"{y}年"] = f"{int(data[y]):,}" if data[y] > 0 else "-"
        entry["累計"] = f"{int(data['累計']):,}"
        entry["占比"] = f"{data['占比']*100:.2f}%"
        rows.append(entry)
        
    return pd.DataFrame(rows), cutoff.strftime("%Y/%m")

# =========================================================
# 4) ✨ Streamlit 專屬渲染函式 (給 app.py 呼叫的入口)
# =========================================================
def render_streamlit(data_dir: Path):
    apply_streamlit_cjk_css()
    st.subheader("📊 表3：歷年累計許可人次 (核發領域別分)")
    
    requested = st.text_input("您可以手動輸入統計截止月份 (例：2025/11)，若留白則自動抓取最新資料：", placeholder="例如：2025/11", key="table3_input")
    requested = requested.strip() or None

    try:
        source_file = pick_latest_source_file(data_dir)
        df_raw = read_and_clean_excel(str(source_file))
        
        table3_df, used_ym_str = build_table3_data(df_raw, requested)
        
        y, m = used_ym_str.split("/")
        st.markdown(f"**目前顯示資料時間：截至 {y}年{int(m)}月**")
        
        # ✨ 將表格渲染在網頁上
        st.dataframe(table3_df, hide_index=True, width="stretch")
        
    except Exception as e:
        st.error(f"讀取或處理資料時發生錯誤：{e}")