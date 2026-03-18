import re
from pathlib import Path
from datetime import datetime, date

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib as mpl
from matplotlib import font_manager as fm
import streamlit as st # ✨ 新增 Streamlit 套件

# =========================================================
# 1) 來源檔自動帶入
# =========================================================
FILE_PATTERNS = ["*累計核發-國別.xlsx", "*累計核發-國別.xls"]
FILENAME_YM_RE = re.compile(r"(\d{6})-(\d{6})累計核發-國別\.(xlsx|xls)$", re.IGNORECASE)

def pick_latest_source_file(base: Path) -> Path:
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
# 2) Excel 與日期處理
# =========================================================
@st.cache_data # ✨ 加上快取機制
def read_excel_with_header_detection(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    engine = "openpyxl" if path.suffix.lower() == ".xlsx" else "xlrd"
    raw = pd.read_excel(path, header=None, engine=engine)
    header_row = 0
    for i in range(min(30, len(raw))):
        row_str = " ".join(raw.iloc[i].astype(str))
        if "統計年月" in row_str and "男" in row_str and "女" in row_str:
            header_row = i
            break
    header = raw.iloc[header_row].astype(str).str.strip()
    df = raw.iloc[header_row + 1 :].copy()
    df.columns = header
    return df.dropna(axis=1, how="all").reset_index(drop=True)

YM_FULL_RE = re.compile(r"^\s*(\d{2,4})\s*[\/\-.年]\s*(\d{1,2})\s*(?:月)?\s*$")

def parse_ym_strict(x):
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

def to_num(series):
    return pd.to_numeric(series.astype(str).str.replace(",", "").replace({"-": "0", "nan": "0"}), errors='coerce').fillna(0)

# =========================================================
# 3) 資料計算邏輯
# =========================================================
def build_top10_data(df: pd.DataFrame, requested_ym: str | None):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    
    ignore = {"統計年月", "男", "女", "總計"}
    country_col = [c for c in df.columns if c not in ignore][0]
    
    df["統計年月"] = df["統計年月"].ffill()
    df["date"] = df["統計年月"].apply(parse_ym_strict)
    
    valid_dates = df["date"].dropna()
    max_date = valid_dates.max()
    
    cutoff = max_date
    if requested_ym:
        req = parse_ym_strict(requested_ym)
        if not pd.isna(req): cutoff = min(req, max_date)

    df_filtered = df[df["date"] <= cutoff].copy()
    df_filtered = df_filtered[df_filtered[country_col].astype(str).str.strip() != "總計"]
    
    df_filtered["total_val"] = to_num(df_filtered["男"]) + to_num(df_filtered["女"])
    g = df_filtered.groupby(country_col)["total_val"].sum().reset_index()
    g.columns = ["國別", "total"]
    g = g.sort_values("total", ascending=False).reset_index(drop=True)
    
    top10 = g.head(10).copy()
    others_sum = g.iloc[10:]["total"].sum()
    if others_sum > 0:
        top10 = pd.concat([top10, pd.DataFrame([{"國別": "其他", "total": others_sum}])], ignore_index=True)
    
    total_all = top10["total"].sum() if top10["total"].sum() > 0 else 1
    top10["share"] = top10["total"] / total_all
    
    return top10, cutoff.strftime("%Y/%m")

# =========================================================
# 4) 繪圖工具 (✨ 拔除存檔，直接回傳 fig，並修正字體)
# =========================================================
def setup_matplotlib_chinese_font():
    # ✨ 把 Linux 的 Noto Sans 放在第一順位
    candidates = ["Noto Sans CJK JP", "Microsoft JhengHei", "DFKai-SB", "Arial Unicode MS", "SimHei", "sans-serif"]
    available = {f.name for f in fm.fontManager.ttflist}
    chosen = next((f for f in candidates if f in available), "sans-serif")
    mpl.rcParams["font.sans-serif"] = [chosen]
    mpl.rcParams["axes.unicode_minus"] = False

def _wrap_label(text_obj, fig, max_px):
    renderer = fig.canvas.get_renderer()
    s = text_obj.get_text()
    if not s or text_obj.get_window_extent(renderer).width <= max_px:
        return
    lines, cur = [], ""
    for char in s:
        test_text = cur + char
        text_obj.set_text(test_text)
        if text_obj.get_window_extent(renderer).width > max_px:
            lines.append(cur)
            cur = char
        else:
            cur = test_text
    lines.append(cur)
    text_obj.set_text("\n".join(lines[:3]))

def plot_fig3(df: pd.DataFrame):
    setup_matplotlib_chinese_font()
    fig, ax = plt.subplots(figsize=(10, 6), dpi=150)
    
    values = df["total"]
    outer_labels = [f"{n} {s*100:.0f}%" for n, s in zip(df["國別"], df["share"])]
    colors = plt.cm.tab20(np.linspace(0, 1, len(values)))

    wedges, texts = ax.pie(values, labels=outer_labels, startangle=90, counterclock=False, 
                           colors=colors, labeldistance=1.1,
                           wedgeprops=dict(edgecolor="white", linewidth=1))
    
    for i, p in enumerate(wedges):
        ang = (p.theta2 - p.theta1)/2. + p.theta1
        y = np.sin(np.deg2rad(ang))
        x = np.cos(np.deg2rad(ang))
        if values[i] > (values.sum() * 0.03):
            ax.text(0.75*x, 0.75*y, f"{int(values[i]):,}", color="white", 
                    weight="bold", ha='center', va='center', fontsize=9)

    fig.canvas.draw()
    for t in texts:
        _wrap_label(t, fig, 150)

    plt.tight_layout()
    return fig # ✨ 核心改變：我們把畫布交給 Streamlit，不存檔了

# =========================================================
# 5) ✨ Streamlit 專屬渲染函式 (給 app.py 呼叫的入口)
# =========================================================
def render_streamlit(data_dir: Path):
    st.subheader("📊 圖3：累計就業金卡前十大國別")
    
    requested = st.text_input("您可以手動輸入截止月份 (例：2025/11)，若留白則自動抓取最新資料：", placeholder="例如：2025/11", key="fig3_input")
    requested = requested.strip() or None

    try:
        source_file = pick_latest_source_file(data_dir)
        df_raw = read_excel_with_header_detection(str(source_file))
        df_plot, used_ym = build_top10_data(df_raw, requested)
        
        y, m = used_ym.split("/")
        st.markdown(f"**目前顯示資料時間：截至 {y}年{int(m)}月**")

        # 1. 顯示圓餅圖
        fig = plot_fig3(df_plot)
        st.pyplot(fig)
        
        # 2. ✨ 加碼顯示原始數據表格
        st.markdown("##### 📄 前十大國別明細")
        
        df_display = df_plot.copy()
        df_display["total"] = df_display["total"].apply(lambda x: f"{int(x):,}")
        df_display["share"] = df_display["share"].apply(lambda x: f"{x*100:.1f}%")
        df_display.columns = ["核發國別", "累計人次", "佔比"]
        
        st.dataframe(df_display, hide_index=True, width="stretch")
        
    except Exception as e:
        st.error(f"讀取或繪圖時發生錯誤：{e}")