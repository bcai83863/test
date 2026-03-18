import re
import platform # ✨ 補上漏掉的套件
from pathlib import Path
from datetime import datetime, date

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st 

# =========================================================
# 0) 領域設定
# =========================================================
DOMAIN_ORDER = ["經濟","科技","教育","數位","金融","文化、藝術","專案會商","建築設計","國防","運動","法律","環境","生技"]
OTHER_GROUP = ["建築設計", "法律", "國防", "運動","環境","生技"]  # 圖2合併為其他

# =========================================================
# 1) 來源檔自動搜尋
# =========================================================
FILE_PATTERNS = ["*累計核發-領域.xlsx", "*累計核發-領域.xls"]
FILENAME_YM_RE = re.compile(r"(\d{6})-(\d{6})累計核發-領域\.(xlsx|xls)$", re.IGNORECASE)

def pick_latest_source_file(base: Path) -> Path:
    candidates = []
    for pat in FILE_PATTERNS:
        candidates.extend(base.glob(pat))

    if not candidates:
        raise FileNotFoundError(f"在資料夾找不到來源檔 (累計核發-領域)")

    with_ym = []
    for p in candidates:
        m = FILENAME_YM_RE.search(p.name)
        if m:
            end_ym = int(m.group(2))
            with_ym.append((end_ym, p))

    if with_ym:
        with_ym.sort(key=lambda x: x[0], reverse=True)
        return with_ym[0][1]

    return max(candidates, key=lambda p: p.stat().st_mtime)

# =========================================================
# 2) matplotlib 字型設定 (解決 Linux/Windows 相容性)
# =========================================================
def apply_font_settings():
    """確保圓餅圖在雲端 Linux 環境下能正確顯示中文"""
    if platform.system() == 'Linux':
        # Streamlit Cloud 專用字型清單
        plt.rcParams['font.sans-serif'] = ['Noto Sans CJK TC', 'Noto Sans CJK JP', 'DejaVu Sans']
    else:
        # 本地 Windows 專用
        plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'DFKai-SB', 'sans-serif']
    
    plt.rcParams['axes.unicode_minus'] = False # 解決負號顯示問題

# =========================================================
# 3) 資料讀取與解析
# =========================================================
@st.cache_data
def read_excel_with_header_detection(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    engine = "openpyxl" if path.suffix.lower() == ".xlsx" else "xlrd"
    raw = pd.read_excel(path, header=None, engine=engine)
    header_row = 0
    for i in range(min(60, len(raw))):
        joined = " ".join(raw.iloc[i].astype(str).tolist())
        if ("統計年月" in joined) and ("男" in joined) and ("女" in joined):
            header_row = i
            break
    header = raw.iloc[header_row].astype(str).str.strip()
    df = raw.iloc[header_row + 1 :].copy()
    df.columns = header
    return df.dropna(axis=1, how="all").reset_index(drop=True)

YM_FULL_RE = re.compile(r"^\s*(\d{2,4})\s*[\/\-.年]\s*(\d{1,2})\s*(?:月)?\s*$")

def parse_ym_strict(x):
    if pd.isna(x): return pd.NaT
    if isinstance(x, (datetime, pd.Timestamp, date)):
        return pd.Timestamp(x.year, x.month, 1)
    s = str(x).strip()
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
# 4) 資料運算邏輯
# =========================================================
def build_fig2_data(df: pd.DataFrame, requested_ym: str | None):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    
    ignore = {"統計年月", "男", "女", "總計", "date"}
    domain_cols = [c for c in df.columns if c not in ignore]
    if not domain_cols: raise ValueError("找不到領域分類欄位")
    domain_col = domain_cols[0]
    
    df["統計年月"] = df["統計年月"].ffill()
    df["date"] = df["統計年月"].apply(parse_ym_strict)
    
    max_date = df["date"].dropna().max()
    cutoff = max_date
    if requested_ym:
        req = parse_ym_strict(requested_ym)
        if not pd.isna(req): cutoff = min(req, max_date)

    df_filtered = df[df["date"] <= cutoff].copy()
    df_filtered = df_filtered[df_filtered[domain_col].astype(str).str.strip() != "總計"]
    
    df_filtered["total_val"] = to_num(df_filtered["男"]) + to_num(df_filtered["女"])
    g = df_filtered.groupby(domain_col)["total_val"].sum().reset_index()
    g.columns = ["領域", "total"]

    base = pd.DataFrame({"領域": DOMAIN_ORDER}).merge(g, on="領域", how="left").fillna(0)
    other_sum = base[base["領域"].isin(OTHER_GROUP)]["total"].sum()
    main = base[~base["領域"].isin(OTHER_GROUP)].copy()
    
    other_label = "其他\n(建築、法律、國防、\n運動、環境、生技)" # ✨ 預先在標籤內加換行
    main = pd.concat([main, pd.DataFrame([{"領域": other_label, "total": other_sum}])], ignore_index=True)
    main = main.sort_values("total", ascending=False).reset_index(drop=True)
    
    total_all = main["total"].sum() if main["total"].sum() > 0 else 1
    main["pct"] = main["total"] / total_all
    
    return main, cutoff.strftime("%Y/%m")

# =========================================================
# 5) 繪圖邏輯 (✨ 加入標籤換行與字型修復)
# =========================================================
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
    text_obj.set_text("\n".join(lines[:4])) 

def plot_fig2(df_plot: pd.DataFrame):
    apply_font_settings() # ✨ 確保字型套用
    
    fig, ax = plt.subplots(figsize=(10, 6), dpi=150)
    
    values = df_plot["total"]
    labels = df_plot["領域"]
    colors = plt.cm.tab20(np.linspace(0, 1, len(values)))

    wedges, _ = ax.pie(values, startangle=90, counterclock=False, colors=colors, 
                       wedgeprops=dict(edgecolor="white", linewidth=1))
    
    legend_labels = [f"{l} {int(v):,} ({p*100:.0f}%)" for l, v, p in zip(labels, values, df_plot["pct"])]
    leg = ax.legend(wedges, legend_labels, loc="center left", bbox_to_anchor=(1, 0.5), frameon=False, fontsize=10)
    
    fig.canvas.draw()
    for t in leg.get_texts():
        _wrap_label(t, fig, 350) 

    plt.tight_layout()
    return fig

# =========================================================
# 6) ✨ Streamlit 專屬渲染函式
# =========================================================
def render_streamlit(data_dir: Path):
    st.subheader("📊 圖2：累計就業金卡持卡人次及比例 (按領域分)")
    
    requested = st.text_input("您可以手動輸入截止月份 (例：2025/11)，若留白則自動抓取最新資料：", 
                              placeholder="例如：2025/11", key="fig2_input")
    requested = requested.strip() or None

    try:
        source_file = pick_latest_source_file(data_dir)
        df_raw = read_excel_with_header_detection(str(source_file))
        df_plot, used_ym = build_fig2_data(df_raw, requested)
        
        y, m = used_ym.split("/")
        st.info(f"📅 目前顯示資料時間：截至 {y}年{int(m)}月")

        # 1. 顯示圓餅圖
        fig = plot_fig2(df_plot)
        st.pyplot(fig)
        
        # 2. 顯示原始數據表格
        st.markdown("##### 📄 領域數據明細")
        df_display = df_plot.copy()
        df_display["total"] = df_display["total"].apply(lambda x: f"{int(x):,}")
        df_display["pct"] = df_display["pct"].apply(lambda x: f"{x*100:.1f}%")
        df_display.columns = ["申請領域", "累計人次", "佔比"]
        st.dataframe(df_display, hide_index=True, width="stretch")
        
    except Exception as e:
        st.error(f"讀取或繪圖時發生錯誤：{e}")