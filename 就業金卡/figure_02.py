import re
from pathlib import Path
from datetime import datetime, date

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st # ✨ 新增 Streamlit 套件

# =========================================================
# 0) 領域設定
# =========================================================
# 領域順序（依你表4）
DOMAIN_ORDER = ["經濟","科技","教育","數位","金融","文化、藝術","專案會商","建築設計","國防","運動","法律","環境","生技"]
OTHER_GROUP = ["建築設計", "法律", "國防", "運動","環境","生技"]  # 圖2合併為其他

# =========================================================
# 1) 來源檔自動帶入
# =========================================================
FILE_PATTERNS = ["*累計核發-領域.xlsx", "*累計核發-領域.xls"]
FILENAME_YM_RE = re.compile(r"(\d{6})-(\d{6})累計核發-領域\.(xlsx|xls)$", re.IGNORECASE)

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
# 2) Excel 讀取與日期解析
# =========================================================
@st.cache_data # ✨ 加上快取機制，提升網頁載入速度
def read_excel_with_header_detection(path_str: str) -> pd.DataFrame:
    path = Path(path_str) # 為了讓快取順利運作，傳入字串再轉回 Path
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
# 3) 資料處理
# =========================================================
def build_fig2_data(df: pd.DataFrame, requested_ym: str | None):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    
    ignore = {"統計年月", "男", "女", "總計"}
    domain_col = [c for c in df.columns if c not in ignore][0]
    
    df["統計年月"] = df["統計年月"].ffill()
    df["date"] = df["統計年月"].apply(parse_ym_strict)
    
    valid_dates = df["date"].dropna()
    max_date = valid_dates.max()
    
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
    
    other_label = "其他（建築設計、法律、國防、運動、環境、生技）"
    main = pd.concat([main, pd.DataFrame([{"領域": other_label, "total": other_sum}])], ignore_index=True)
    main = main.sort_values("total", ascending=False).reset_index(drop=True)
    
    total_all = main["total"].sum() if main["total"].sum() > 0 else 1
    main["pct"] = main["total"] / total_all
    
    return main, cutoff.strftime("%Y/%m")

# =========================================================
# 4) 繪圖工具 (✨ 拔除存檔，直接回傳 fig，並修正字體)
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
    text_obj.set_text("\n".join(lines[:3])) 

def plot_fig2(df_plot: pd.DataFrame):
    # ✨ 修正為 Linux 專用的思源黑體，徹底消滅豆腐塊！
    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["font.sans-serif"] = ["Noto Sans CJK JP", "sans-serif"]
    
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
    
    return fig # ✨ 核心改變：我們把畫布交給 Streamlit，不存檔了

# =========================================================
# 5) ✨ Streamlit 專屬渲染函式 (給 app.py 呼叫的入口)
# =========================================================
def render_streamlit(data_dir: Path):
    st.subheader("📊 圖2：累計就業金卡持卡人次及比例 (按領域分)")
    
    # 精緻的網頁輸入框，加入 key 避免跟其他頁面衝突
    requested = st.text_input("您可以手動輸入截止月份 (例：2025/11)，若留白則自動抓取最新資料：", placeholder="例如：2025/11", key="fig2_input")
    requested = requested.strip() or None

    try:
        source_file = pick_latest_source_file(data_dir)
        df_raw = read_excel_with_header_detection(str(source_file)) # 轉字串給快取
        df_plot, used_ym = build_fig2_data(df_raw, requested)
        
        y, m = used_ym.split("/")
        st.markdown(f"**目前顯示資料時間：截至 {y}年{int(m)}月**")

        # 1. 顯示圓餅圖
        fig = plot_fig2(df_plot)
        st.pyplot(fig)
        
        # 2. ✨ 加碼顯示原始數據表格 (取代原本的匯出 CSV)
        st.markdown("##### 📄 領域數據明細")
        
        # 整理一下表格的顯示格式，讓它更美觀
        df_display = df_plot.copy()
        df_display["total"] = df_display["total"].apply(lambda x: f"{int(x):,}") # 加上千分位撇號
        df_display["pct"] = df_display["pct"].apply(lambda x: f"{x*100:.1f}%")    # 換成百分比格式
        df_display.columns = ["申請領域", "累計人次", "佔比"]
        
        st.dataframe(df_display, hide_index=True, width="stretch")
        
    except Exception as e:
        st.error(f"讀取或繪圖時發生錯誤：{e}")