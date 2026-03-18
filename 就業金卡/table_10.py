import re
from pathlib import Path
from typing import Optional, Tuple

import pandas as pd
import streamlit as st # ✨ 新增 Streamlit 套件

# =========================================================
# 0) 常數設定
# =========================================================
DOMAIN_ORDER = ["經濟","教育","科技","數位","文化、藝術","金融","專案會商","建築設計","國防","運動","法律","環境","生技"]

# =========================================================
# 1) 智慧選檔邏輯
# =========================================================
def is_candidate_file(p: Path) -> bool:
    n = p.name
    if p.suffix.lower() not in (".xlsx", ".xls"): return False
    # 必須包含：有效、領域、且與性別/男女相關
    if "有效" not in n: return False
    if "領域" not in n: return False
    if not any(k in n for k in ["性別", "男", "女"]): return False
    return True

def pick_latest_file(base: Path) -> Tuple[Path, pd.Timestamp]:
    files = [p for p in base.iterdir() if is_candidate_file(p)]
    if not files:
        raise FileNotFoundError("找不到符合規格的『有效持卡領域及性別』Excel 檔案")

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
# 2) 數據解析工具
# =========================================================
def to_num(s): 
    return pd.to_numeric(s.astype(str).str.replace(",", "").replace({"-": "0"}), errors='coerce').fillna(0)

@st.cache_data # ✨ 加上快取機制
def load_and_detect_header(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    engine = "openpyxl" if path.suffix == ".xlsx" else "xlrd"
    raw = pd.read_excel(path, header=None, engine=engine)
    
    # 偵測表頭列
    h_row = 0
    for i in range(min(50, len(raw))):
        row_str = " ".join(raw.iloc[i].astype(str).fillna(""))
        if "領域" in row_str and any(k in row_str for k in ["男", "女性"]):
            h_row = i
            break
            
    df = raw.iloc[h_row+1:].copy()
    df.columns = [str(c).strip().replace("\n", "") for c in raw.iloc[h_row]]
    return df

# =========================================================
# 3) 資料處理：產生表 10 內容
# =========================================================
def build_table10_data(df: pd.DataFrame) -> pd.DataFrame:
    # 找欄位
    col_dom = next((c for c in df.columns if any(k in c for k in ["領域", "專長"])), None)
    col_m = next((c for c in df.columns if c == "男" or "男性" in c), None)
    col_f = next((c for c in df.columns if c == "女" or "女性" in c), None)

    if not all([col_dom, col_m, col_f]):
        raise RuntimeError("Excel 缺少必要欄位 (領域/男/女)")
    
    df[col_dom] = df[col_dom].astype(str).str.strip()
    d = df[~df[col_dom].isin(["", "總計", "合計", "nan", "性別占比"])].copy()
    
    d["male_v"] = to_num(d[col_m])
    d["female_v"] = to_num(d[col_f])
    d["total_v"] = d["male_v"] + d["female_v"]
    
    g = d.groupby(col_dom)[["male_v", "female_v", "total_v"]].sum()
    # 補齊所有領域
    g = g.reindex(DOMAIN_ORDER).fillna(0)
    # 按總計降冪排序
    g = g.sort_values("total_v", ascending=False)
    
    grand_total = g["total_v"].sum()
    grand_male = g["male_v"].sum()
    grand_female = g["female_v"].sum()
    
    rows = []
    for domain, data in g.iterrows():
        rows.append({
            "領域專長": domain,
            "男性": int(data['male_v']),
            "女性": int(data['female_v']),
            "總計": int(data['total_v'])
        })
        
    # 加入總計
    rows.append({
        "領域專長": "總計",
        "男性": int(grand_male),
        "女性": int(grand_female),
        "總計": int(grand_total)
    })
    
    # 加入性別占比
    denom = grand_total if grand_total > 0 else 1
    rows.append({
        "領域專長": "性別占比",
        "男性": f"{grand_male/denom*100:.1f}%",
        "女性": f"{grand_female/denom*100:.1f}%",
        "總計": "100.0%"
    })
    
    return pd.DataFrame(rows)

# =========================================================
# 4) ✨ Streamlit 專屬渲染函式 (入口)
# =========================================================
def render_streamlit(data_dir: Path):
    st.subheader("📊 表10：有效許可人次 (按領域別及性別分)")

    try:
        source_file, used_ym = pick_latest_file(data_dir)
        st.markdown(f"**目前顯示資料時間：截至 {used_ym.year}年{used_ym.month}月底**")

        # 讀取與解析
        df_raw = load_and_detect_header(str(source_file))
        table10_df = build_table10_data(df_raw)

        # 格式化數值 (加上千分位)
        df_display = table10_df.copy()
        for col in ["男性", "女性", "總計"]:
            df_display[col] = df_display[col].apply(lambda x: f"{x:,}" if isinstance(x, (int, float)) else x)

        # ✨ 渲染網頁表格
        st.dataframe(df_display, hide_index=True, width="stretch")
        
    except Exception as e:
        st.error(f"讀取或產表時發生錯誤：{e}")