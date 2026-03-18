import re
from pathlib import Path
from typing import Optional, Tuple, List

import pandas as pd
import streamlit as st # ✨ 新增 Streamlit 套件

# =========================================================
# 0) 參數與順序設定
# =========================================================
TOP_N = 10  # 十大國別
DOMAIN_ORDER = [
    "科技", "經濟", "教育", "文化、藝術", "運動", "金融", "法律",
    "建築設計", "國防", "數位", "專案會商", "環境", "生技"
]

# =========================================================
# 1) 檔案自動偵測邏輯
# =========================================================
def pick_latest_report_file(base: Path) -> Path:
    REPORT_PATTERNS = ["*國籍和領域統計報表.xlsx", "*國籍和領域統計報表.xls"]
    REPORT_DATE_RE = re.compile(r"^(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])")
    
    candidates: List[Path] = []
    for pat in REPORT_PATTERNS:
        candidates.extend(base.glob(pat))
    if not candidates:
        raise FileNotFoundError(f"找不到國籍領域統計報表")

    with_date = []
    for p in candidates:
        m = REPORT_DATE_RE.match(p.name)
        if m:
            ymd = int(m.group(1) + m.group(2) + m.group(3))
            with_date.append((ymd, p))
    
    if with_date:
        with_date.sort(key=lambda x: x[0], reverse=True)
        return with_date[0][1]
    
    return sorted(candidates, key=lambda p: p.stat().st_mtime)[-1]

# =========================================================
# 2) 資料標準化工具 (處理領域名稱差異)
# =========================================================
PUNC_RE = re.compile(r"[ \t\r\n\u3000·•．\.\-—_／/\\,，、&＆:：;；()（）]")

def norm_key(s: str) -> str:
    return PUNC_RE.sub("", str(s).strip())

CANON_DOMAIN_MAP = {norm_key(x): x for x in DOMAIN_ORDER}

def to_num(series):
    return pd.to_numeric(series.astype(str).str.replace(",", "").replace({"-": "0", "nan": "0"}), errors='coerce').fillna(0)

@st.cache_data # ✨ 強大快取：包含工作表掃描與格式判斷邏輯
def read_and_detect_best_table(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    xls = pd.ExcelFile(path)
    best_sheet, best_df_processed = None, None
    max_score = -1
    
    # 掃描所有工作表，找出最像我們要的那張
    for sn in xls.sheet_names:
        raw = pd.read_excel(path, sheet_name=sn, header=None)
        for i in range(min(50, len(raw))):
            row_str = " ".join(raw.iloc[i].astype(str))
            score = sum(2 for d in DOMAIN_ORDER if d in row_str)
            if "國籍" in row_str or "國別" in row_str: score += 10
            if score > max_score:
                max_score, best_sheet = score, sn
    
    # 讀取選定的工作表並定位表頭
    raw_df = pd.read_excel(path, sheet_name=best_sheet, header=None)
    h_row = 0
    for i in range(min(60, len(raw_df))):
        row_str = " ".join(raw_df.iloc[i].astype(str))
        if "國" in row_str and any(d in row_str for d in DOMAIN_ORDER):
            h_row = i
            break
    
    df_out = raw_df.iloc[h_row+1:].copy()
    df_out.columns = raw_df.iloc[h_row].astype(str).str.strip()
    return df_out

# =========================================================
# 3) 核心：產表邏輯 (長寬表自動轉換)
# =========================================================
def build_table6_data(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip() for c in df.columns]
    
    # 國別欄定位
    country_col = next((c for c in df.columns if any(k in c for k in ["國別", "國籍"])), None)
    if not country_col: raise RuntimeError("找不到國別欄位")

    # 標準化欄位名稱
    rename_map = {}
    for c in df.columns:
        k = norm_key(c)
        if k in CANON_DOMAIN_MAP: rename_map[c] = CANON_DOMAIN_MAP[k]
    df = df.rename(columns=rename_map)

    # 寬表與長表邏輯
    wide_domains = [d for d in DOMAIN_ORDER if d in df.columns]
    
    if len(wide_domains) >= 3:
        # A) 寬表邏輯
        d = df.copy()
        d[country_col] = d[country_col].astype(str).str.strip()
        d = d[~d[country_col].isin(["", "總計", "合計"])].copy()
        for dom in DOMAIN_ORDER:
            if dom not in d.columns: d[dom] = 0
            d[dom] = to_num(d[dom])
        d["row_total"] = d[DOMAIN_ORDER].sum(axis=1)
        g = d.groupby(country_col)[DOMAIN_ORDER + ["row_total"]].sum().reset_index()
    else:
        # B) 長表邏輯
        domain_col = next((c for c in df.columns if "領域" in c or "專長" in c), None)
        if not domain_col: raise RuntimeError("找不到領域欄位")
        d = df.copy()
        d[country_col] = d[country_col].astype(str).str.strip()
        d[domain_col] = d[domain_col].apply(lambda x: CANON_DOMAIN_MAP.get(norm_key(x), norm_key(x)))
        d = d[~d[country_col].isin(["", "總計", "合計"])].copy()
        val_col = "總計" if "總計" in d.columns else d.columns[-1]
        d["val"] = to_num(d[val_col])
        piv = d.pivot_table(index=country_col, columns=domain_col, values="val", aggfunc="sum").fillna(0)
        for dom in DOMAIN_ORDER:
            if dom not in piv.columns: piv[dom] = 0
        piv = piv[DOMAIN_ORDER].copy()
        piv["row_total"] = piv.sum(axis=1)
        g = piv.reset_index()

    # 十大排序
    g = g.sort_values("row_total", ascending=False).reset_index(drop=True)
    top = g.head(TOP_N).copy()
    rest = g.iloc[TOP_N:].copy()
    
    if not rest.empty:
        other_row = pd.DataFrame([{
            country_col: "其他",
            **{dom: rest[dom].sum() for dom in DOMAIN_ORDER},
            "row_total": rest["row_total"].sum()
        }])
        top = pd.concat([top, other_row], ignore_index=True)
    
    total_row = pd.DataFrame([{
        country_col: "總計",
        **{dom: top[dom].sum() for dom in DOMAIN_ORDER},
        "row_total": top["row_total"].sum()
    }])
    final_df = pd.concat([top, total_row], ignore_index=True)
    
    # 格式化顯示
    final_df.columns = ["國別"] + DOMAIN_ORDER + ["總計"]
    return final_df

# =========================================================
# 4) ✨ Streamlit 專屬渲染函式 (給 app.py 呼叫的入口)
# =========================================================
def render_streamlit(data_dir: Path):
    st.subheader("📊 表6：有效就業金卡十大國別 (按領域分)")
    
    try:
        src = pick_latest_report_file(data_dir)
        df_raw = read_and_detect_best_table(str(src))
        
        # 執行原本強大的運算邏輯
        table6_df = build_table6_data(df_raw)
        
        m = re.search(r"(20\d{2})(0[1-9]|1[0-2])", src.stem)
        tag = f"{m.group(1)}年{int(m.group(2))}月" if m else "最新"
        st.markdown(f"**目前顯示資料時間：{tag} 統計結果**")
        
        # 為了美觀，顯示時加上千分位撇號
        df_display = table6_df.copy()
        for col in df_display.columns:
            if col != "國別":
                df_display[col] = df_display[col].apply(lambda x: f"{int(x):,}")

        # ✨ 渲染網頁表格
        st.dataframe(df_display, hide_index=True, width="stretch")
        
    except Exception as e:
        st.error(f"自動偵測或產表時發生錯誤：{e}")