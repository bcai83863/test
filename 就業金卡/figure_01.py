import re
from pathlib import Path
from datetime import datetime, date

import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
from matplotlib import ticker
import streamlit as st # ✨ 新增 Streamlit 套件

# =========================================================
# 1) 自動尋找最新來源檔案
# =========================================================
FILE_PATTERN_RE = re.compile(r"(?P<start>\d{6})-(?P<end>\d{6})累計核發-領域\.(xlsx|xls)$")

def find_latest_source_file(base_dir: Path) -> Path:
    candidates = []
    for p in base_dir.glob("*累計核發-領域.*"):
        m = FILE_PATTERN_RE.match(p.name)
        if m:
            candidates.append((m.group("end"), p))

    if not candidates:
        raise FileNotFoundError("找不到符合命名規則的 Excel 檔案")

    candidates.sort(key=lambda x: x[0])
    latest = candidates[-1][1]
    return latest

# =========================================================
# 2) matplotlib 設定 (✨ 已更新為 Linux 雲端字體)
# =========================================================
def setup_matplotlib():
    matplotlib.rcParams["axes.unicode_minus"] = False
    # 這裡直接指定我們稍早安裝的思源黑體，確保網頁顯示無誤
    matplotlib.rcParams["font.sans-serif"] = ["Noto Sans CJK JP", "sans-serif"]

# =========================================================
# 3) 讀 Excel（自動找表頭）
# =========================================================
@st.cache_data # ✨ 加入快取機制：讓網頁切換時不用重複讀取 Excel，速度飛快！
def read_excel_with_header_detection(path: Path) -> pd.DataFrame:
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

def read_source(base_dir: Path):
    return read_excel_with_header_detection(find_latest_source_file(base_dir))

# =========================================================
# 4) 年月解析與數值清洗 (維持原樣)
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
    return s.astype(str).str.replace(",", "", regex=False).replace({"-": "", "nan": ""}).pipe(pd.to_numeric, errors="coerce").fillna(0)

# =========================================================
# 5) 建立累計資料 (維持原樣)
# =========================================================
def build_fig1_series(df: pd.DataFrame, requested_ym: str | None):
    df = df.copy()
    df["date"] = df["統計年月"].ffill().apply(parse_ym)
    df = df.dropna(subset=["date"])

    if requested_ym:
        cutoff_date = parse_ym(requested_ym)
        df = df[df["date"] <= cutoff_date]

    dim_col = [c for c in df.columns if c not in ["統計年月", "男", "女", "總計"]][0]
    df = df[df[dim_col].astype(str).str.strip() == "總計"]

    df["monthly"] = to_num(df["男"]) + to_num(df["女"])
    m = df.groupby("date", as_index=False)["monthly"].sum().sort_values("date")
    m["cumulative"] = m["monthly"].cumsum()

    cutoff_date = pd.Timestamp(m["date"].max().year, m["date"].max().month, 1)
    used_ym = cutoff_date.strftime("%Y/%m")
    return m, used_ym, cutoff_date

# =========================================================
# 6) 繪圖 (✨ 改為回傳 fig，而不是存檔)
# =========================================================
def plot_fig1(m: pd.DataFrame, cutoff_date: pd.Timestamp):
    fig, ax = plt.subplots(figsize=(9.5, 4.8))
    line, = ax.plot(m["date"].to_numpy(), m["cumulative"].to_numpy(), linewidth=1.8)
    line_color = line.get_color()

    tmp = m.copy()
    tmp["year"] = tmp["date"].dt.year
    tmp["month"] = tmp["date"].dt.month
    cutoff_year, cutoff_month = int(cutoff_date.year), int(cutoff_date.month)

    def pick_year_dec_or_last(df_year: pd.DataFrame) -> pd.DataFrame:
        dec = df_year[df_year["month"] == 12]
        return dec.tail(1) if not dec.empty else df_year.tail(1)

    selected_frames = []
    years_all = sorted([int(y) for y in tmp["year"].unique()])

    if cutoff_month == 12:
        for y in [y for y in years_all if y <= cutoff_year]:
            df_y = tmp[tmp["year"] == y]
            if not df_y.empty: selected_frames.append(pick_year_dec_or_last(df_y))
    else:
        for y in [y for y in years_all if y < cutoff_year]:
            df_y = tmp[tmp["year"] == y]
            if not df_y.empty: selected_frames.append(pick_year_dec_or_last(df_y))
        cutoff_row = tmp[(tmp["year"] == cutoff_year) & (tmp["month"] == cutoff_month)]
        selected_frames.append(cutoff_row.tail(1) if not cutoff_row.empty else tmp.tail(1))

    points = pd.concat(selected_frames, ignore_index=True).drop_duplicates(subset=["date"]).sort_values("date").reset_index(drop=True)

    major_ticks = [points["date"].iloc[0] - pd.DateOffset(years=1)] + points["date"].tolist()
    major_labels = [f"{cutoff_year}\n{cutoff_month}月" if i == len(major_ticks) - 1 and cutoff_month != 12 else "" for i, _ in enumerate(major_ticks)]
            
    minor_ticks, minor_labels = [], []
    for i in range(len(major_ticks)-1):
        d1, d2 = major_ticks[i], major_ticks[i+1]
        if d2.month == 12:
            minor_ticks.append(d1 + (d2 - d1) / 2)
            minor_labels.append(str(d2.year))

    ax.set_xticks(major_ticks)
    ax.set_xticklabels(major_labels, fontsize=14, ha="center")
    ax.set_xticks(minor_ticks, minor=True)
    ax.set_xticklabels(minor_labels, minor=True, fontsize=14, ha="center")
    ax.tick_params(axis='x', which='minor', length=0) 
    
    delta_days = 999
    if len(points) >= 2:
        delta_days = (points["date"].iloc[-1] - points["date"].iloc[-2]).days
        if delta_days < 100 and cutoff_month != 12:
            ax.get_xticklabels()[-1].set_ha("left")

    ax.tick_params(axis="y", labelsize=14)
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda v, _: f"{int(v):,}"))
    ax.set_ylim(bottom=0)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.set_xlabel("西元(年)", fontweight="bold", fontsize=18)
    ax.set_ylabel("人次", fontweight="bold", fontsize=18, rotation=0, labelpad=40)
    ax.yaxis.label.set_horizontalalignment("right")
    ax.yaxis.label.set_verticalalignment("center")
    ax.yaxis.set_label_coords(-0.09, 0.5)

    ax.plot(points["date"].to_numpy(), points["cumulative"].to_numpy(), linestyle="None", marker="o", markersize=7, color=line_color, zorder=5)

    for i, r in points.iterrows():
        ha_val, xy_text_offset = "center", (0, 13) 
        if len(points) >= 2 and delta_days < 100:
            if i == len(points) - 2:
                ha_val, xy_text_offset = "right", (-8, 13)
            elif i == len(points) - 1:
                ha_val, xy_text_offset = "left", (8, 13)

        ax.annotate(f"{int(r['cumulative']):,}", xy=(r["date"], r["cumulative"]), xytext=xy_text_offset, textcoords="offset points", ha=ha_val, va="bottom", fontsize=14)

    ax.set_xlim(major_ticks[0] - pd.Timedelta(days=20), tmp["date"].max() + pd.Timedelta(days=60))
    plt.tight_layout()
    
    return fig # ✨ 核心改變：我們不存檔了，直接把畫好的畫布(fig)交出去！

# =========================================================
# 7) ✨ Streamlit 專屬渲染函式 (給 app.py 呼叫的入口)
# =========================================================
def render_streamlit(data_dir: Path):
    setup_matplotlib()
    
    # 用 Streamlit 的排版，讓標題和輸入框並排或好看一點
    st.subheader("📊 圖1：就業金卡累計核發人次")
    
    # 替換掉原本終端機的 input()，變成精緻的網頁輸入框
    requested = st.text_input("您可以手動輸入截止月份 (例：2025/12)，若留白則自動抓取最新資料：", placeholder="例如：2025/12")
    requested = requested.strip() or None

    try:
        # 執行原本的數據邏輯
        df = read_source(data_dir)
        m, used_ym, cutoff_date = build_fig1_series(df, requested)
        
        ym_tag = used_ym.replace("/", "年") + "月"
        st.markdown(f"**目前顯示資料時間：截至 {ym_tag}**")

        # 呼叫畫圖函式拿回 fig，並讓 Streamlit 顯示出來
        fig = plot_fig1(m, cutoff_date)
        st.pyplot(fig)
        
    except Exception as e:
        st.error(f"讀取或繪圖時發生錯誤：{e}")