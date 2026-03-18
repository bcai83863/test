import re
from pathlib import Path
from typing import Tuple, Optional

import pandas as pd
import streamlit as st # ✨ 新增 Streamlit 套件

# =========================================================
# 0) 常數設定
# =========================================================
DOMAIN_ORDER = [
    "經濟", "教育", "科技", "數位", "文化、藝術", "金融", "專案會商",
    "建築設計", "國防", "運動", "法律", "環境", "生技"
]

# =========================================================
# 1) 智慧選檔邏輯
# =========================================================
def pick_best_source_for_table9(base: Path) -> Path:
    excel_files = list(base.glob("*.xlsx")) + list(base.glob("*.xls"))
    if not excel_files:
        raise FileNotFoundError(f"在 {base} 找不到任何 Excel 檔案")

    best_file = None
    max_score = -1

    for p in excel_files:
        name = p.name
        score = 0
        if "有效" in name: score += 5
        if "領域" in name: score += 3
        if any(k in name for k in ["境內", "境外", "境內外"]): score += 5
        
        if score > max_score:
            max_score = score
            best_file = p
            
    if not best_file:
        raise RuntimeError("找不到結構符合『有效持卡領域及境內外』的來源檔")
    return best_file

# =========================================================
# 2) 數據解析工具
# =========================================================
def to_num(series):
    return pd.to_numeric(series.astype(str).str.replace(",", "").replace({"-": "0", "nan": "0"}), errors='coerce').fillna(0)

@st.cache_data # ✨ 加上快取機制
def load_and_detect_header(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    engine = "openpyxl" if path.suffix.lower() == ".xlsx" else "xlrd"
    raw = pd.read_excel(path, header=None, engine=engine)
    
    # 偵測表頭列
    h_row = 0
    for i in range(min(50, len(raw))):
        row_str = " ".join(raw.iloc[i].astype(str).fillna(""))
        if "領域" in row_str and "境內" in row_str:
            h_row = i
            break
            
    df = raw.iloc[h_row+1:].copy()
    df.columns = [str(c).strip().replace("\n", "") for c in raw.iloc[h_row]]
    return df

# =========================================================
# 3) 資料處理：產生表 9 內容
# =========================================================
def build_table9_data(df: pd.DataFrame, file_stem: str) -> Tuple[pd.DataFrame, str]:
    # 找欄位名稱
    col_domain = next((c for c in df.columns if "領域" in c), None)
    col_in = next((c for c in df.columns if "境內" in c), None)
    col_out = next((c for c in df.columns if "境外" in c), None)
    
    if not all([col_domain, col_in, col_out]):
        raise RuntimeError("檔案缺少必要欄位 (領域/境內/境外)")

    # 清洗資料
    df[col_domain] = df[col_domain].astype(str).str.strip()
    d = df[~df[col_domain].isin(["", "總計", "合計", "nan"])].copy()
    
    d["境內"] = to_num(d[col_in])
    d["境外"] = to_num(d[col_out])
    d["總計"] = d["境內"] + d["境外"]
    
    # 彙總並排序
    g = d.groupby(col_domain)[["境外", "境內", "總計"]].sum()
    # 確保 DOMAIN_ORDER 裡的項都存在並填 0
    g = g.reindex(DOMAIN_ORDER).fillna(0)
    # 按總計排序
    g = g.sort_values("總計", ascending=False)
    
    grand_total = g["總計"].sum()
    
    # 建立輸出格式
    rows = []
    for domain, data in g.iterrows():
        rows.append({
            "領域專長": domain,
            "境外": int(data['境外']),
            "境內": int(data['境內']),
            "總計": int(data['總計']),
            "占比": f"{data['總計']/(grand_total if grand_total>0 else 1)*100:.2f}%"
        })
        
    # 加入總計列
    rows.append({
        "領域專長": "總計",
        "境外": int(g['境外'].sum()),
        "境內": int(g['境內'].sum()),
        "總計": int(grand_total),
        "占比": "100.00%"
    })
    
    # 抓取月份標籤 (從檔名)
    m = re.search(r"(20\d{2})(0[1-9]|1[0-2])", file_stem)
    used_ym = f"{m.group(1)}/{m.group(2)}" if m else "最新"
    
    return pd.DataFrame(rows), used_ym

# =========================================================
# 4) ✨ Streamlit 專屬渲染函式 (入口)
# =========================================================
def render_streamlit(data_dir: Path):
    st.subheader("📊 表9：有效許可人次 (按領域別及境內外分)")

    try:
        # 1. 自動挑選最佳檔案
        source_file = pick_best_source_for_table9(data_dir)
        
        # 2. 讀取並處理數據
        df_raw = load_and_detect_header(str(source_file))
        table9_df, used_ym = build_table9_data(df_raw, source_file.stem)
        
        y, m = (used_ym.split("/") if "/" in used_ym else ("最新", ""))
        st.markdown(f"**目前顯示資料時間：截至 {y}年{int(m) if m else ''}月底**")

        # 為了網頁表格好看，將數值欄位轉為千分位字串
        df_display = table9_df.copy()
        for col in ["境外", "境內", "總計"]:
            df_display[col] = df_display[col].apply(lambda x: f"{x:,}")

        # ✨ 渲染網頁表格
        st.dataframe(df_display, hide_index=True, width="stretch")
        
        st.info("💡 註：境內外身分依申請時填報之居住地為準。")

    except Exception as e:
        st.error(f"讀取或處理資料時發生錯誤：{e}")