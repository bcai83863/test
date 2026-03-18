import re
from pathlib import Path
from datetime import datetime, date
from typing import Optional, Tuple, List

import pandas as pd
import streamlit as st # ✨ 新增 Streamlit 套件
from font_utils import apply_streamlit_cjk_css

# =========================================================
# 0) 參數設定
# =========================================================
AGE_ORDER_CANON = [
    "未滿20歲", "20歲-29歲", "30歲-39歲", "40歲-49歲",
    "50歲-59歲", "60歲-69歲", "70歲以上"
]

# =========================================================
# 1) 選檔與標準化邏輯
# =========================================================
def pick_latest_file(base: Path) -> Tuple[Path, pd.Timestamp]:
    files = [p for p in base.iterdir() if p.suffix.lower() in (".xlsx", ".xls") and "有效" in p.name and "年齡" in p.name]
    if not files:
        raise FileNotFoundError("找不到符合規格的『有效持卡年齡』Excel 檔案")

    pool = []
    for p in files:
        m = re.search(r"(20\d{2})(0[1-9]|1[0-2])", p.stem)
        if m:
            pool.append((pd.Timestamp(int(m.group(1)), int(m.group(2)), 1), p))
    
    if pool:
        ym, p = sorted(pool, key=lambda x: x[0])[-1]
        return p, ym
    return max(files, key=lambda x: x.stat().st_mtime), pd.Timestamp.now()

def normalize_age(x):
    s = str(x).replace(" ", "").replace("歲", "")
    if "未滿20" in s: return "未滿20歲"
    if "70以上" in s or "超過70" in s or "70+" in s: return "70歲以上"
    m = re.search(r"(\d{2}).*?(\d{2})", s)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        return f"{a}歲-{b}歲"
    return None

# =========================================================
# 2) 數據解析工具 (加上快取)
# =========================================================
@st.cache_data # ✨ 加上快取機制
def load_and_detect_table14(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    engine = "openpyxl" if path.suffix == ".xlsx" else "xlrd"
    raw = pd.read_excel(path, header=None, engine=engine)
    
    h_row = 0
    for i in range(min(50, len(raw))):
        row_str = " ".join(raw.iloc[i].astype(str).fillna(""))
        if "年齡" in row_str and any(k in row_str for k in ["男", "女性"]):
            h_row = i
            break
            
    df = raw.iloc[h_row+1:].copy()
    df.columns = [str(c).strip().replace("\n", "") for c in raw.iloc[h_row]]
    return df

# =========================================================
# 3) 資料處理邏輯
# =========================================================
def build_table14_data(df: pd.DataFrame) -> pd.DataFrame:
    col_age = next((c for c in df.columns if "年齡" in c), df.columns[0])
    col_m = next((c for c in df.columns if c == "男" or "男性" in c), None)
    col_f = next((c for c in df.columns if c == "女" or "女性" in c), None)

    def to_num(s): return pd.to_numeric(s.astype(str).str.replace(",", "").replace({"-": "0"}), errors='coerce').fillna(0)
    
    df["age_std"] = df[col_age].apply(normalize_age)
    d = df[df["age_std"].notna()].copy()
    
    d["male_v"] = to_num(d[col_m])
    d["female_v"] = to_num(d[col_f])
    d["total_v"] = d["male_v"] + d["female_v"]
    
    g = d.groupby("age_std")[["male_v", "female_v", "total_v"]].sum().reindex(AGE_ORDER_CANON).fillna(0)
    
    grand_total = g["total_v"].sum()
    grand_male = g["male_v"].sum()
    grand_female = g["female_v"].sum()
    
    rows = []
    for age, data in g.iterrows():
        rows.append({
            "年齡區間": age,
            "男性": int(data['male_v']),
            "女性": int(data['female_v']),
            "總計": int(data['total_v']),
            "年齡占比": f"{data['total_v']/(grand_total if grand_total>0 else 1)*100:.2f}%"
        })
        
    rows.append({
        "年齡區間": "總計",
        "男性": int(grand_male),
        "女性": int(grand_female),
        "總計": int(grand_total),
        "年齡占比": "100.00%"
    })
    
    denom = grand_total if grand_total > 0 else 1
    rows.append({
        "年齡區間": "性別占比",
        "男性": f"{grand_male/denom*100:.1f}%",
        "女性": f"{grand_female/denom*100:.1f}%",
        "總計": "100.0%",
        "年齡占比": "-"
    })
    
    return pd.DataFrame(rows)

# =========================================================
# 4) ✨ Streamlit 專屬渲染函式 (入口)
# =========================================================
def render_streamlit(data_dir: Path):
    apply_streamlit_cjk_css()
    st.subheader("📊 表14：有效許可人次 (按年齡別及性別分)")

    try:
        source_file, used_ym = pick_latest_file(data_dir)
        y, m = used_ym.year, used_ym.month
        st.markdown(f"**目前顯示資料時間：截至 {y}年{m}月底**")

        # 讀取並處理
        df_raw = load_and_detect_table14(str(source_file))
        table14_df = build_table14_data(df_raw)

        # 為了網頁表格顯示，多加一欄「統計時間」來取代 Word 的合併儲存格效果
        df_display = table14_df.copy()
        # 排除最後兩列不顯示年份，讓版面更乾淨
        df_display.insert(0, "統計年月", f"{y}年{m}月底")
        df_display.loc[df_display.index[-2:], "統計年月"] = "-"

        # 格式化數值 (千分位)
        for col in ["男性", "女性", "總計"]:
            df_display[col] = df_display[col].apply(lambda x: f"{x:,}" if isinstance(x, (int, float)) else x)

        # ✨ 渲染網頁表格
        st.dataframe(df_display, hide_index=True, width="stretch")
        
        st.info("💡 註：最後兩列為全體總計與性別佔比分析。")

    except Exception as e:
        st.error(f"讀取或產表時發生錯誤：{e}")