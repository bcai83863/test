import re
from pathlib import Path
from datetime import datetime, date
from typing import Tuple

import pandas as pd
import streamlit as st 
from font_utils import apply_streamlit_cjk_css

# =========================================================
# 0) 參數設定
# =========================================================
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
        raise FileNotFoundError("找不到符合命名規則的『累計核發-領域』來源檔")

    with_ym = []
    for p in candidates:
        m = FILENAME_YM_RE.search(p.name)
        if m: with_ym.append((int(m.group(2)), p))

    if with_ym:
        with_ym.sort(key=lambda x: x[0], reverse=True)
        return with_ym[0][1]
    return max(candidates, key=lambda p: p.stat().st_mtime)

# =========================================================
# 2) 日期解析與快取讀取
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
        return pd.Timestamp(y, mo, 1)
    return pd.NaT

def to_num(series):
    return pd.to_numeric(series.astype(str).str.replace(",", "").replace({"-": "0", "nan": "0"}), errors='coerce').fillna(0)

@st.cache_data
def read_and_clean_excel(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    engine = "openpyxl" if path.suffix.lower() == ".xlsx" else "xlrd"
    raw_xls = pd.read_excel(path, header=None, engine=engine)
    
    header_row = 0
    for i in range(min(40, len(raw_xls))):
        row_str = " ".join(raw_xls.iloc[i].astype(str))
        if "統計年月" in row_str and "男" in row_str:
            header_row = i
            break
            
    df = raw_xls.iloc[header_row+1:].copy()
    df.columns = [str(c).strip() for c in raw_xls.iloc[header_row]]
    return df

# =========================================================
# 3) 資料處理邏輯
# =========================================================
def build_table4_data(df: pd.DataFrame, requested_ym: str | None) -> Tuple[pd.DataFrame, str]:
    df.columns = [str(c).strip() for c in df.columns]
    ignore = {"統計年月", "男", "女", "總計", "date"}
    domain_cols = [c for c in df.columns if c not in ignore]
    if not domain_cols: raise ValueError("找不到領域分類欄位")
    domain_col = domain_cols[0]
    
    df["統計年月"] = df["統計年月"].ffill()
    df["date"] = df["統計年月"].apply(parse_ym_strict)
    
    valid_dates = df["date"].dropna()
    max_date = valid_dates.max()
    cutoff = max_date
    if requested_ym:
        req = parse_ym_strict(requested_ym)
        if not pd.isna(req): cutoff = min(req, max_date)

    d = df[df["date"] <= cutoff].copy()
    d = d[d[domain_col].astype(str).str.strip() != "總計"]
    
    d["male_v"] = to_num(d["男"])
    d["female_v"] = to_num(d["女"])
    d["total_v"] = d["male_v"] + d["female_v"]

    g = d.groupby(domain_col)[["male_v", "female_v", "total_v"]].sum().reindex(DOMAIN_ORDER).fillna(0)
    
    grand_total = g["total_v"].sum()
    grand_male = g["male_v"].sum()
    grand_female = g["female_v"].sum()
    
    # 格式化輸出
    rows = []
    for idx, r in g.iterrows():
        rows.append({
            "領域別": idx,
            "男性": f"{int(r['male_v']):,}",
            "女性": f"{int(r['female_v']):,}",
            "總計": f"{int(r['total_v']):,}",
            "領域占比": f"{r['total_v']/(grand_total if grand_total>0 else 1)*100:.2f}%"
        })
    
    # 總計列
    rows.append({
        "領域別": "總計",
        "男性": f"{int(grand_male):,}",
        "女性": f"{int(grand_female):,}",
        "總計": f"{int(grand_total):,}",
        "領域占比": "100.00%"
    })
    
    # 性別占比列
    rows.append({
        "領域別": "性別占比",
        "男性": f"{grand_male/(grand_total if grand_total>0 else 1)*100:.1f}%",
        "女性": f"{grand_female/(grand_total if grand_total>0 else 1)*100:.1f}%",
        "總計": "100.0%",
        "領域占比": "-"
    })
    
    used_ym_str = f"{cutoff.year}/{int(cutoff.month):02d}"
    return pd.DataFrame(rows), used_ym_str

# =========================================================
# 4) ✨ Streamlit 渲染入口
# =========================================================
def render_streamlit(data_dir: Path):
    apply_streamlit_cjk_css()
    st.subheader("📊 表4：累計許可人次 (按領域別及性別分)")
    
    requested = st.text_input("您可以手動輸入截止月份 (例：2025/11)，若留白則自動抓取最新資料：", 
                              placeholder="例如：2025/11", key="table4_input")
    requested = requested.strip() or None

    try:
        source_file = pick_latest_source_file(data_dir)
        df_raw = read_and_clean_excel(str(source_file))
        table4_df, used_ym_str = build_table4_data(df_raw, requested)
        
        y, m = used_ym_str.split("/")
        st.info(f"📅 目前顯示資料時間：截至 {y}年{int(m)}月底")
        
        # ✨ CSV 下載功能 (UTF-8 with BOM 以利 Excel)
        csv = table4_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 下載領域性別報表 (CSV)",
            data=csv,
            file_name=f"表4_累計許可人次_領域及性別_{y}{int(m):02d}.csv",
            mime="text/csv",
        )

        st.dataframe(table4_df, hide_index=True, use_container_width=True)
        
    except Exception as e:
        st.error(f"讀取或處理資料時發生錯誤：{e}")