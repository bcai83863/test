import re
from pathlib import Path
from datetime import datetime, date

import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
from matplotlib import ticker
import streamlit as st 

from font_utils import apply_cjk_font_settings, apply_streamlit_cjk_css

# =========================================================
# 1) 自動尋找最新來源檔案
# =========================================================
FILE_PATTERN_RE = re.compile(r"(?P<start>\d{6})-(?P<end>\d{6})累計核發-領域\.(xlsx|xls)$")

def find_latest_source_file(base_dir: Path) -> Path:
    candidates = []
    # 這裡搜尋 base_dir 下的所有檔案
    for p in base_dir.glob("*累計核發-領域.*"):
        m = FILE_PATTERN_RE.match(p.name)
        if m:
            candidates.append((m.group("end"), p))

    if not candidates:
        raise FileNotFoundError("找不到符合命名規則的 Excel 檔案 (累計核發-領域)")

    candidates.sort(key=lambda x: x[0])
    return candidates[-1][1]

# =========================================================
# 2) matplotlib 字型設定函式
# =========================================================
def apply_font_settings():
    """套用跨平台 CJK 字型設定，確保 Streamlit Linux 可正確顯示中文。"""
    apply_cjk_font_settings()

# =========================================================
# 3) 讀 Excel（自動找表頭）
# =========================================================
@st.cache_data
def read_excel_with_header_detection(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    engine = "openpyxl" if path.suffix == ".xlsx" else "xlrd"
    raw = pd.read_excel(path, header=None, engine=engine)

    header_row = 0
    for i in range(min(30, len(raw))):
        row = " ".join(raw.iloc[i].astype(str))
        if "統計年月" in row and "總計" in row:
            header_row = i
            break

    header = raw.iloc[header_row].astype(str).str.strip()
    df = raw.iloc[header_row + 1 :].copy()
    df.columns = header
    return df.reset_index(drop=True)

# =========================================================
# 4) 年月解析與數值清洗
# =========================================================
YM_RE = re.compile(r"(\d{2,4})[\/\-.年](\d{1,2})")

def parse_ym(x):
    if pd.isna(x): return pd.NaT
    if isinstance(x, (datetime, date, pd.Timestamp)): return pd.Timestamp(x.year, x.month, 1)
    s = str(x).strip()
    if re.fullmatch(r"\d{6}", s): return pd.to_datetime(s, format="%Y%m", errors="coerce")
    m = YM_RE.search(s)
    if m:
        y, mth = int(m.group(1)), int(m.group(2))
        if y < 1900: y += 1911
        return pd.Timestamp(y, mth, 1)
    return pd.NaT

def to_num(s: pd.Series):
    return s.astype(str).str.replace(",", "", regex=False).replace({"-": "0", "nan": "0"}).pipe(pd.to_numeric, errors="coerce").fillna(0)

# =========================================================
# 5) 建立累計資料
# =========================================================
def build_fig1_series(df: pd.DataFrame, requested_ym: str | None):
    df = df.copy()
    df["date"] = df["統計年月"].ffill().apply(parse_ym)
    df = df.dropna(subset=["date"])

    if requested_ym:
        cutoff_date = parse_ym(requested_ym)
        if not pd.isna(cutoff_date):
            df = df[df["date"] <= cutoff_date]

    # 尋找領域欄位（除了統計年月和男女總計以外的那個）
    dim_cols = [c for c in df.columns if c not in ["統計年月", "男", "女", "總計", "date"]]
    if not dim_cols:
        raise ValueError("Excel 格式異常：找不到領域分類欄位")
    
    dim_col = dim_cols[0]
    df = df[df[dim_col].astype(str).str.strip() == "總計"]

    df["monthly"] = to_num(df["男"]) + to_num(df["女"])
    m = df.groupby("date", as_index=False)["monthly"].sum().sort_values("date")
    m["cumulative"] = m["monthly"].cumsum()

    max_date = m["date"].max()
    used_ym = f"{max_date.year}/{max_date.month}"
    return m, used_ym, max_date

# =========================================================
# 6) 繪圖邏輯
# =========================================================
def plot_fig1(m: pd.DataFrame, cutoff_date: pd.Timestamp):
    fig, ax = plt.subplots(figsize=(9.5, 4.8))
    line, = ax.plot(m["date"].to_numpy(), m["cumulative"].to_numpy(), linewidth=1.8)
    line_color = line.get_color()

    tmp = m.copy()
    tmp["year"] = tmp["date"].dt.year
    tmp["month"] = tmp["date"].dt.month
    cutoff_year, cutoff_month = int(cutoff_date.year), int(cutoff_date.month)

    # 挑選每年 12 月以及最後一個月作為標籤點
    selected_frames = []
    years_all = sorted([int(y) for y in tmp["year"].unique()])

    for y in years_all:
        df_y = tmp[tmp["year"] == y]
        if y < cutoff_year:
            dec = df_y[df_y["month"] == 12]
            selected_frames.append(dec.tail(1) if not dec.empty else df_y.tail(1))
        elif y == cutoff_year:
            target = df_y[df_y["month"] == cutoff_month]
            selected_frames.append(target.tail(1) if not target.empty else df_y.tail(1))

    points = pd.concat(selected_frames, ignore_index=True).drop_duplicates(subset=["date"]).sort_values("date").reset_index(drop=True)

    # 座標軸與標籤設定
    ax.tick_params(axis="both", labelsize=12)
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda v, _: f"{int(v):,}"))
    ax.set_ylim(bottom=0)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.set_xlabel("西元(年)", fontweight="bold", fontsize=14)
    ax.set_ylabel("人次", fontweight="bold", fontsize=14, rotation=0, labelpad=20)

    # 繪製圓點
    ax.plot(points["date"].to_numpy(), points["cumulative"].to_numpy(), linestyle="None", marker="o", markersize=6, color=line_color, zorder=5)

    # 標註數值
    for i, r in points.iterrows():
        ax.annotate(f"{int(r['cumulative']):,}", 
                    xy=(r["date"], r["cumulative"]), 
                    xytext=(0, 10), 
                    textcoords="offset points", 
                    ha="center", fontsize=11)

    plt.tight_layout()
    return fig

# =========================================================
# 7) ✨ Streamlit 入口函式
# =========================================================
def render_streamlit(data_dir: Path):
    apply_streamlit_cjk_css()
    apply_font_settings() # 套用字型設定
    
    st.subheader("📊 圖1：就業金卡累計核發人次")
    
    # 側邊欄或上方輸入框
    requested = st.text_input("您可以手動輸入截止月份 (例：2025/12)，若留白則自動抓取最新資料：", 
                              placeholder="例如：2025/12", key="fig1_input")
    requested = requested.strip() or None

    try:
        source_file = find_latest_source_file(data_dir)
        df = read_excel_with_header_detection(str(source_file))
        m, used_ym, cutoff_date = build_fig1_series(df, requested)
        
        st.info(f"📅 目前顯示資料時間：截至 {used_ym}")

        fig = plot_fig1(m, cutoff_date)
        st.pyplot(fig)
        
    except Exception as e:
        st.error(f"讀取或繪圖時發生錯誤：{e}")
