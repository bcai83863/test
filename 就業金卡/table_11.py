import re
from pathlib import Path
from typing import Optional, Tuple, List

import pandas as pd
import streamlit as st # ✨ 新增 Streamlit 套件

# =========================================================
# 0) 參數設定
# =========================================================
TOP_N = 10  # 前 10 名 + 其他 + 總計

# =========================================================
# 1) 選檔與讀取邏輯 (加上快取)
# =========================================================
def is_candidate_file(p: Path) -> bool:
    n = p.name
    if p.suffix.lower() not in (".xlsx", ".xls"):
        return False
    # 關鍵字：有效 + 國別/國籍 + 境內外
    return ("有效" in n) and (("國別" in n) or ("國籍" in n)) and ("境內外" in n)

def pick_latest_file(base: Path) -> Tuple[Path, pd.Timestamp]:
    files = sorted(list(base.glob("*.xlsx")) + list(base.glob("*.xls")))
    cands = [p for p in files if is_candidate_file(p)]
    if not cands:
        raise RuntimeError("資料夾內找不到『有效 × 國別/國籍 × 境內外』相關 Excel 檔。")

    pool: List[Tuple[Path, pd.Timestamp]] = []
    for p in cands:
        # 嘗試從檔名抓日期
        name = p.stem
        m = re.search(r"(20\d{2})(0[1-9]|1[0-2])([0-3]\d)?", name)
        if m:
            pool.append((p, pd.Timestamp(int(m.group(1)), int(m.group(2)), 1)))

    if pool:
        p, ym = sorted(pool, key=lambda x: x[1])[-1]
        return p, ym

    # 若抓不到日期則依修改時間
    latest = max(cands, key=lambda x: x.stat().st_mtime)
    return latest, pd.Timestamp.now()

@st.cache_data # ✨ 智慧表頭偵測快取
def load_with_smart_header(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    engine = "openpyxl" if path.suffix.lower() == ".xlsx" else "xlrd"
    raw = pd.read_excel(path, sheet_name=0, engine=engine, header=None)
    
    # 偵測表頭列 (尋找關鍵字密度最高的一列)
    best_i, best_score = 0, -1
    for i in range(min(80, len(raw))):
        row = raw.iloc[i].astype(str).fillna("").tolist()
        joined = " ".join(row)
        score = sum(3 for k in ["國別", "國籍", "境外", "境內"] if k in joined)
        score += (1 if "總計" in joined else 0)
        if score > best_score:
            best_score, best_i = score, i
            
    header = raw.iloc[best_i].astype(str).str.replace("\n", "", regex=False).str.strip()
    df = raw.iloc[best_i + 1 :].copy()
    df.columns = header
    df = df.dropna(axis=1, how="all").reset_index(drop=True)
    return df

# =========================================================
# 2) 資料處理邏輯
# =========================================================
def to_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.strip().replace({"-": "0", "": "0"}), 
        errors="coerce"
    ).fillna(0)

def build_table11(df: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    d = df.copy()
    d.columns = [str(c).strip() for c in d.columns]
    
    # 定位欄位
    col_country = next((c for c in d.columns if any(k in c for k in ["國別", "國籍"])), d.columns[0])
    col_in = next((c for c in d.columns if "境內" in c), None)
    col_out = next((c for c in d.columns if "境外" in c), None)
    
    if not col_in or not col_out:
        raise RuntimeError("找不到『境內』或『境外』欄位。")

    d[col_country] = d[col_country].astype(str).str.strip()
    d = d[(d[col_country] != "") & (d[col_country] != "總計")].copy()

    d["in_v"] = to_num(d[col_in])
    d["out_v"] = to_num(d[col_out])
    d["total_v"] = d["in_v"] + d["out_v"]

    g = d.groupby(col_country, as_index=False)[["out_v", "in_v", "total_v"]].sum()
    g = g.sort_values("total_v", ascending=False).reset_index(drop=True)

    # 處理 Top N 與 其他
    top = g.head(top_n).copy()
    rest = g.iloc[top_n:].copy()
    
    if not rest.empty:
        other_row = pd.DataFrame({
            col_country: ["其他"],
            "out_v": [rest["out_v"].sum()],
            "in_v": [rest["in_v"].sum()],
            "total_v": [rest["total_v"].sum()],
        })
        top = pd.concat([top, other_row], ignore_index=True)

    # 總計列
    grand_total = top["total_v"].sum()
    grand_in = top["in_v"].sum()
    grand_out = top["out_v"].sum()
    
    # 格式化輸出
    rows = []
    for _, r in top.iterrows():
        rows.append({
            "國別": r[col_country],
            "境外": int(r["out_v"]),
            "境內": int(r["in_v"]),
            "總計": int(r["total_v"]),
            "占比": f"{r['total_v']/(grand_total if grand_total > 0 else 1)*100:.2f}%"
        })
        
    rows.append({
        "國別": "總計",
        "境外": int(grand_out),
        "境內": int(grand_in),
        "總計": int(grand_total),
        "占比": "100.00%"
    })

    return pd.DataFrame(rows)

# =========================================================
# 3) ✨ Streamlit 專屬渲染函式 (入口)
# =========================================================
def render_streamlit(data_dir: Path):
    st.subheader("📊 表11：有效許可人次 (按國別及境內外分)")

    try:
        source_file, used_ym = pick_latest_file(data_dir)
        st.markdown(f"**目前顯示資料時間：截至 {used_ym.year}年{used_ym.month}月底**")

        # 讀取與解析
        df_raw = load_with_smart_header(str(source_file))
        table11_df = build_table11(df_raw, top_n=TOP_N)

        # 格式化數值
        df_display = table11_df.copy()
        for col in ["境外", "境內", "總計"]:
            df_display[col] = df_display[col].apply(lambda x: f"{x:,}")

        # ✨ 渲染網頁表格
        st.dataframe(df_display, hide_index=True, width="stretch")
        
    except Exception as e:
        st.error(f"讀取或產表時發生錯誤：{e}")