import re
from pathlib import Path
from datetime import datetime, date

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
import streamlit as st 

from font_utils import apply_cjk_font_settings, apply_streamlit_cjk_css

# =========================================================
# 0) 設定
# =========================================================
AGE_ORDER = [
    "未滿20歲", "20歲-24歲", "25歲-29歲", "30歲-34歲", "35歲-39歲",
    "40歲-44歲", "45歲-49歲", "50歲-54歲", "55歲-59歲", "60歲-64歲", "超過65歲"
]

AGE_LABELS_DISPLAY = [
    "未滿20", "20-24", "25-29", "30-34", "35-39",
    "40-44", "45-49", "50-54", "55-59", "60-64", "65以上"
]

# =========================================================
# 1) 來源檔自動搜尋
# =========================================================
FILE_PATTERNS = ["*累計核發-年齡.xlsx", "*累計核發-年齡.xls"]
FILENAME_YM_RE = re.compile(r"(\d{6})-(\d{6})累計核發-年齡\.(xlsx|xls)$", re.IGNORECASE)

def pick_latest_source_file(base: Path) -> Path:
    candidates = []
    for pat in FILE_PATTERNS:
        candidates.extend(base.glob(pat))

    if not candidates:
        raise FileNotFoundError("在資料夾找不到『累計核發-年齡』來源檔")

    with_ym = []
    for p in candidates:
        m = FILENAME_YM_RE.search(p.name)
        if m:
            with_ym.append((int(m.group(2)), p))

    if with_ym:
        with_ym.sort(key=lambda x: x[0], reverse=True)
        return with_ym[0][1]

    return max(candidates, key=lambda p: p.stat().st_mtime)

# =========================================================
# 2) matplotlib 字型設定 (解決 Linux/Windows 相容性)
# =========================================================
def apply_font_settings():
    """套用跨平台 CJK 字型設定，確保 Streamlit Linux 可正確顯示中文。"""
    apply_cjk_font_settings()

# =========================================================
# 3) 資料讀取與解析
# =========================================================
@st.cache_data
def read_and_clean_excel(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    engine = "openpyxl" if path.suffix.lower() == ".xlsx" else "xlrd"
    df_raw = pd.read_excel(path, header=None, engine=engine)
    
    header_row = 0
    for i in range(min(40, len(df_raw))):
        row_str = " ".join(df_raw.iloc[i].astype(str))
        if "統計年月" in row_str and "男" in row_str:
            header_row = i
            break
            
    df_processed = df_raw.iloc[header_row + 1:].copy()
    df_processed.columns = [str(c).strip() for c in df_raw.iloc[header_row]]
    return df_processed.dropna(axis=1, how="all").reset_index(drop=True)

def parse_ym_strict(x):
    YM_FULL_RE = re.compile(r"^\s*(\d{2,4})\s*[\/\-.年]\s*(\d{1,2})\s*(?:月)?\s*$")
    if pd.isna(x): return pd.NaT
    if isinstance(x, (datetime, pd.Timestamp, date)): return pd.Timestamp(x.year, x.month, 1)
    s = str(x).strip()
    if re.fullmatch(r"\d{6}", s): return pd.to_datetime(s, format="%Y%m", errors="coerce")
    m = YM_FULL_RE.match(s)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if y < 1900: y += 1911
        return pd.Timestamp(y, mo, 1)
    return pd.NaT

def to_num(series):
    return pd.to_numeric(series.astype(str).str.replace(",", "").replace({"-": "0", "nan": "0"}), errors='coerce').fillna(0)

# =========================================================
# 4) 資料計算邏輯
# =========================================================
def build_fig4_data(df: pd.DataFrame, requested_ym: str | None):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    
    ignore = {"統計年月", "男", "女", "總計", "date"}
    age_cols = [c for c in df.columns if c not in ignore]
    if not age_cols: raise ValueError("找不到年齡分類欄位")
    age_col = age_cols[0]
    
    df["統計年月"] = df["統計年月"].ffill()
    df["date"] = df["統計年月"].apply(parse_ym_strict)
    
    # 移除統計資料累計塊 (如果有)
    hit = df.astype(str).apply(lambda col: col.str.contains("統計資料累計", na=False)).any(axis=1)
    if hit.any():
        df = df.loc[: hit[hit].index[0] - 1].copy()

    valid_dates = df["date"].dropna()
    max_date = valid_dates.max()
    cutoff = max_date
    if requested_ym:
        req = parse_ym_strict(requested_ym)
        if not pd.isna(req): cutoff = min(req, max_date)

    df_filtered = df[(df["date"] <= cutoff) & (df[age_col].astype(str).str.strip() != "總計")].copy()
    df_filtered["male_v"] = to_num(df_filtered["男"])
    df_filtered["female_v"] = to_num(df_filtered["女"])
    
    g = df_filtered.groupby(age_col)[["male_v", "female_v"]].sum().reindex(AGE_ORDER).fillna(0).reset_index()
    g.columns = ["年齡", "male", "female"]
    g["total"] = g["male"] + g["female"]
    
    return g, cutoff.strftime("%Y/%m")

# =========================================================
# 5) 繪圖邏輯
# =========================================================
def plot_fig4(fig_data: pd.DataFrame):
    apply_font_settings() # ✨ 套用字型
    
    # 邏輯：未滿20歲若為0則排除
    plot_df = fig_data.copy()
    if plot_df.loc[0, "total"] == 0:
        plot_df = plot_df.iloc[1:].reset_index(drop=True)
        display_labels = AGE_LABELS_DISPLAY[1:]
    else:
        display_labels = AGE_LABELS_DISPLAY

    fig, ax = plt.subplots(figsize=(10, 6), dpi=150)
    
    x = range(len(display_labels))
    ax.bar(x, plot_df["male"], width=0.5, label="男", color="#095EDB")
    ax.bar(x, plot_df["female"], width=0.5, bottom=plot_df["male"], label="女", color="#FF3300")

    # 軸設定
    ax.set_xticks(x)
    ax.set_xticklabels(display_labels, fontsize=11, color="black")
    ax.set_xlabel("年齡(歲)", fontweight="bold", fontsize=12, color="black")
    ax.set_ylabel("人次", fontweight="bold", fontsize=12, color="black", rotation=0, labelpad=20)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{int(v):,}"))
    ax.tick_params(axis="y", labelcolor="black")
    
    # 標註總數
    for i, t in enumerate(plot_df["total"]):
        if t > 0:
            ax.text(i, t + (plot_df["total"].max() * 0.01), f"{int(t):,}", ha='center', fontsize=10)

    ax.legend(loc="upper right", frameon=False)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    
    plt.tight_layout()
    return fig

# =========================================================
# 6) ✨ Streamlit 專屬渲染函式
# =========================================================
def render_streamlit(data_dir: Path):
    apply_streamlit_cjk_css()
    st.subheader("📊 圖4：累計核發就業金卡年齡及性別")
    
    requested = st.text_input("您可以手動輸入截止月份 (例：2025/11)，若留白則自動抓取最新資料：", 
                              placeholder="例如：2025/11", key="fig4_input")
    requested = requested.strip() or None

    try:
        source_file = pick_latest_source_file(data_dir)
        df_raw = read_and_clean_excel(str(source_file))
        fig_data, used_ym = build_fig4_data(df_raw, requested)
        
        y, m = used_ym.split("/")
        st.info(f"📅 目前顯示資料時間：截至 {y}年{int(m)}月")

        # 1. 顯示堆疊長條圖
        fig = plot_fig4(fig_data)
        st.pyplot(fig)
        
        # 2. 顯示原始數據表格
        st.markdown("##### 📄 年齡及性別人次明細")
        df_display = fig_data.copy()
        df_display.columns = ["年齡級距", "男性 (人次)", "女性 (人次)", "總計 (人次)"]
        
        for col in ["男性 (人次)", "女性 (人次)", "總計 (人次)"]:
            df_display[col] = df_display[col].apply(lambda x: f"{int(x):,}")
        
        st.dataframe(df_display, hide_index=True, width="stretch")
        
    except Exception as e:
        st.error(f"讀取或繪圖時發生錯誤：{e}")
