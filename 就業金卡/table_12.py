import re
from pathlib import Path
from typing import Optional, Tuple, List

import pandas as pd
import streamlit as st # ✨ 新增 Streamlit 套件

# =========================================================
# 0) 參數設定
# =========================================================
TOP_N = 10  # 十大國別
DOMAIN_ORDER = [
    "科技", "經濟", "教育", "文化、藝術", "運動", "金融", "法律",
    "建築設計", "國防", "數位", "專案會商", "環境", "生技"
]

# =========================================================
# 1) 智慧選檔邏輯
# =========================================================
def is_candidate_file(p: Path) -> bool:
    n = p.name
    if p.suffix.lower() not in (".xlsx", ".xls"): return False
    # 關鍵字：有效、國別/國籍、領域
    if "有效" not in n: return False
    if not (("國別" in n) or ("國籍" in n)): return False
    if "領域" not in n: return False
    return True

def pick_latest_file(base: Path) -> Tuple[Path, pd.Timestamp]:
    files = [p for p in base.iterdir() if is_candidate_file(p)]
    if not files:
        raise FileNotFoundError("找不到符合規格的『有效持卡國別及領域』Excel 檔案")

    pool = []
    for p in files:
        m = re.search(r"(20\d{2})(0[1-9]|1[0-2])", p.stem)
        if m:
            pool.append((pd.Timestamp(int(m.group(1)), int(m.group(2)), 1), p))
    
    if pool:
        ym, p = sorted(pool, key=lambda x: x[0])[-1]
        return p, ym
    
    latest_p = max(files, key=lambda x: x.stat().st_mtime)
    return latest_p, pd.Timestamp.now()

# =========================================================
# 2) 數據解析工具 (加上快取)
# =========================================================
def to_num(series):
    return pd.to_numeric(series.astype(str).str.replace(",", "").replace({"-": "0", "nan": "0"}), errors='coerce').fillna(0)

@st.cache_data # ✨ 快取機制：包含表頭偵測
def load_and_detect_header(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    engine = "openpyxl" if path.suffix == ".xlsx" else "xlrd"
    raw = pd.read_excel(path, header=None, engine=engine)
    
    # 偵測表頭：尋找包含「國」且有領域關鍵字的列
    h_row = 0
    for i in range(min(60, len(raw))):
        row_str = " ".join(raw.iloc[i].astype(str).fillna(""))
        if ("國" in row_str) and any(d in row_str for d in ["科技", "經濟"]):
            h_row = i
            break
            
    df = raw.iloc[h_row+1:].copy()
    df.columns = [str(c).strip().replace("\n", "") for c in raw.iloc[h_row]]
    return df

# =========================================================
# 3) 資料處理邏輯
# =========================================================
def build_table12_data(df: pd.DataFrame) -> pd.DataFrame:
    # 定位國別欄位
    col_country = next((c for c in df.columns if any(k in c for k in ["國別", "國籍"])), None)
    if not col_country: raise RuntimeError("找不到國別欄位")

    # 清洗數據
    d = df[~df[col_country].isin(["", "總計", "合計", "nan"])].copy()
    
    # 確保所有領域欄位都存在並轉為數值
    for dom in DOMAIN_ORDER:
        target_col = next((c for c in d.columns if dom in c), None)
        if target_col:
            d[dom] = to_num(d[target_col])
        else:
            d[dom] = 0
    
    # 按國別彙總
    g = d.groupby(col_country)[DOMAIN_ORDER].sum()
    g["row_total"] = g.sum(axis=1)
    g = g.sort_values("row_total", ascending=False)
    
    # 處理 Top 10 與 其他
    top = g.head(TOP_N).copy()
    rest = g.iloc[TOP_N:].copy()
    if not rest.empty:
        other_row = pd.DataFrame([{
            **{dom: rest[dom].sum() for dom in DOMAIN_ORDER},
            "row_total": rest["row_total"].sum()
        }], index=["其他"])
        top = pd.concat([top, other_row])
    
    # 加入總計
    grand_total = top["row_total"].sum()
    total_row = pd.DataFrame([{
        **{dom: top[dom].sum() for dom in DOMAIN_ORDER},
        "row_total": grand_total
    }], index=["總計"])
    final_df = pd.concat([top, total_row])
    
    # 計算占比
    denom = grand_total if grand_total > 0 else 1
    final_df["占比(%)"] = (final_df["row_total"] / denom * 100).apply(lambda x: f"{x:.2f}%")
    
    final_df = final_df.reset_index().rename(columns={"index": "國別", "row_total": "總計"})
    return final_df

# =========================================================
# 4) ✨ Streamlit 專屬渲染函式 (入口)
# =========================================================
def render_streamlit(data_dir: Path):
    st.subheader("📊 表12：有效許可人次 (按國別及領域別分)")

    try:
        source_file, used_ym = pick_latest_file(data_dir)
        st.markdown(f"**目前顯示資料時間：截至 {used_ym.year}年{used_ym.month}月底**")

        # 1. 讀取並處理數據
        df_raw = load_and_detect_header(str(source_file))
        table12_df = build_table12_data(df_raw)

        # 2. 為了美觀，將數值欄位轉為千分位字串
        df_display = table12_df.copy()
        numeric_cols = DOMAIN_ORDER + ["總計"]
        for col in numeric_cols:
            df_display[col] = df_display[col].apply(lambda x: f"{int(x):,}")

        # ✨ 渲染網頁表格
        # 使用 stretch 寬度，並讓國別欄位固定在左側
        st.dataframe(df_display, hide_index=True, width="stretch")
        
        st.info("💡 註：本表列出前十大國別，其餘歸類為「其他」。")

    except Exception as e:
        st.error(f"讀取或產表時發生錯誤：{e}")