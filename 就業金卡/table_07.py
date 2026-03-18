import re
from pathlib import Path
from datetime import datetime, date
from typing import Tuple, Optional

import pandas as pd
import streamlit as st # ✨ 新增 Streamlit 套件
from font_utils import apply_streamlit_cjk_css

# =========================================================
# 0) 參數設定
# =========================================================
# 年齡順序
AGE_ORDER = [
    "未滿20歲", "20歲-24歲", "25歲-29歲", "30歲-34歲", "35歲-39歲",
    "40歲-44歲", "45歲-49歲", "50歲-54歲", "55歲-59歲", "60歲-64歲", "超過65歲",
]

# =========================================================
# 1) 來源檔自動搜尋 (挑選：*累計核發-年齡)
# =========================================================
def pick_latest_source_file(base: Path) -> Path:
    FILE_PATTERNS = ["*累計核發-年齡.xlsx", "*累計核發-年齡.xls"]
    FILENAME_YM_RE = re.compile(r"(\d{6})-(\d{6})累計核發-年齡\.(xlsx|xls)$", re.IGNORECASE)
    
    candidates = []
    for pat in FILE_PATTERNS:
        candidates.extend(base.glob(pat))

    if not candidates:
        raise FileNotFoundError(f"在資料夾找不到來源檔：{FILE_PATTERNS}")

    with_ym = []
    for p in candidates:
        m = FILENAME_YM_RE.search(p.name)
        if m:
            end_ym = int(m.group(2))
            with_ym.append((end_ym, p))

    if with_ym:
        with_ym.sort(key=lambda x: x[0], reverse=True)
        return with_ym[0][1]

    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[0]

# =========================================================
# 2) Excel 讀取與資料解析 (加上快取)
# =========================================================
@st.cache_data # ✨ 加上快取機制
def read_and_detect_header(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    engine = "openpyxl" if path.suffix.lower() == ".xlsx" else "xlrd"
    raw = pd.read_excel(path, header=None, engine=engine)

    header_row = 0
    for i in range(min(40, len(raw))):
        joined = " ".join(raw.iloc[i].astype(str).tolist())
        if ("統計年月" in joined) and ("男" in joined) and ("女" in joined):
            header_row = i
            break

    header = raw.iloc[header_row].astype(str).str.replace("\n", "", regex=False).str.strip()
    df = raw.iloc[header_row + 1 :].copy()
    df.columns = header
    df = df.dropna(axis=1, how="all").reset_index(drop=True)
    return df

# =========================================================
# 3) 資料處理邏輯
# =========================================================
def parse_ym_strict(x):
    YM_FULL_RE = re.compile(r"^\s*(\d{2,4})\s*[\/\-.年]\s*(\d{1,2})\s*(?:月)?\s*$")
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

def to_num(series: pd.Series) -> pd.Series:
    s = series.copy().astype(str)
    s = s.str.replace(",", "", regex=False).str.strip().replace({"nan": "", "NaN": "", "-": ""})
    return pd.to_numeric(s, errors="coerce").fillna(0)

def build_table7(df: pd.DataFrame, requested_ym: str | None) -> Tuple[pd.DataFrame, str]:
    d = df.copy()
    d.columns = [str(c).strip().replace("\n", "") for c in d.columns]
    
    # 找年齡欄位
    ignore = {"統計年月", "男", "女", "總計"}
    age_col = [c for c in d.columns if c not in ignore][0]

    # 移除統計資料累計塊
    hit = d.astype(str).apply(lambda col: col.str.contains("統計資料累計", na=False)).any(axis=1)
    if hit.any():
        d = d.loc[: hit[hit].index[0] - 1].copy()

    d["date"] = d["統計年月"].ffill().apply(parse_ym_strict)
    
    # 決定截止日期
    max_date = pd.to_datetime(d["date"].dropna(), errors="coerce").max()
    cutoff_date = max_date
    if requested_ym:
        req = parse_ym_strict(requested_ym)
        if not pd.isna(req): cutoff_date = min(req, max_date)
    
    used_ym = f"{cutoff_date.year}/{cutoff_date.month:02d}"

    d = d[d["date"] <= cutoff_date].copy()
    d[age_col] = d[age_col].astype(str).str.replace("\u3000", "", regex=False).str.strip()
    d = d[(d[age_col] != "") & (d[age_col] != "總計")].copy()

    d["male"] = to_num(d["男"])
    d["female"] = to_num(d["女"])
    d["total"] = d["male"] + d["female"]

    g = d.groupby(age_col)[["male", "female", "total"]].sum()
    g = g.reindex(AGE_ORDER).fillna(0.0)

    grand_total = g["total"].sum() if g["total"].sum() > 0 else 1.0
    g["年齡占比"] = g["total"] / grand_total

    # 建立最終輸出列 (含總計與性別占比)
    rows = []
    # 1. 各年齡層數據
    for idx, r in g.iterrows():
        rows.append({
            "年齡區間": idx,
            "男性": f"{int(r['male']):,}",
            "女性": f"{int(r['female']):,}",
            "總計": f"{int(r['total']):,}",
            "年齡占比": f"{r['年齡占比']*100:.2f}%"
        })
    
    # 2. 總計列
    grand_male = g["male"].sum()
    grand_female = g["female"].sum()
    rows.append({
        "年齡區間": "總計",
        "男性": f"{int(grand_male):,}",
        "女性": f"{int(grand_female):,}",
        "總計": f"{int(grand_total):,}",
        "年齡占比": "100.00%"
    })
    
    # 3. 性別占比列
    rows.append({
        "年齡區間": "性別占比",
        "男性": f"{grand_male/grand_total*100:.1f}%",
        "女性": f"{grand_female/grand_total*100:.1f}%",
        "總計": "100.0%",
        "年齡占比": "-"
    })

    return pd.DataFrame(rows), used_ym

# =========================================================
# 4) ✨ Streamlit 專屬渲染函式 (入口)
# =========================================================
def render_streamlit(data_dir: Path):
    apply_streamlit_cjk_css()
    st.subheader("📊 表7：月底累計許可人次 (按年齡別及性別分)")
    
    requested = st.text_input("您可以手動輸入統計截止月份 (例：2025/11)，若留白則自動抓取最新資料：", placeholder="例如：2025/11", key="table7_input")
    requested = requested.strip() or None

    try:
        source_file = pick_latest_source_file(data_dir)
        df_raw = read_and_detect_header(str(source_file))
        
        table7_df, used_ym = build_table7(df_raw, requested)
        
        y, m = used_ym.split("/")
        st.markdown(f"**目前顯示資料時間：截至 {y}年{int(m)}月底**")
        
        # ✨ 渲染網頁表格
        st.dataframe(table7_df, hide_index=True, width="stretch")
        
    except Exception as e:
        st.error(f"讀取或處理資料時發生錯誤：{e}")