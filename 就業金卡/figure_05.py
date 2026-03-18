import re
from pathlib import Path
from typing import Optional, Tuple

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st # ✨ 新增 Streamlit 套件

# =========================================================
# 0) 設定
# =========================================================
CAND_KEYWORDS = ("有效", "領域")  # 檔名需同時包含
DOMAIN_ORDER = ["經濟", "教育", "科技", "數位", "文化、藝術", "金融", "專案會商", "建築設計", "法律", "國防", "運動", "環境", "生技"]
OTHER_GROUP = ["建築設計", "法律", "國防", "運動", "環境", "生技"]  # 合併成「其他」

# =========================================================
# 1) 選檔邏輯
# =========================================================
def infer_ym_from_filename(p: Path) -> Optional[Tuple[int, int]]:
    stem = p.stem
    m1 = re.search(r"(20\d{2})(0[1-9]|1[0-2])([0-3]\d)", stem)
    if m1: return int(m1.group(1)), int(m1.group(2))
    m2 = re.search(r"(20\d{2})(0[1-9]|1[0-2])", stem)
    if m2: return int(m2.group(1)), int(m2.group(2))
    return None

def pick_latest_source_file(base: Path) -> Tuple[Path, Optional[Tuple[int, int]]]:
    files = sorted(list(base.glob("*.xlsx")) + list(base.glob("*.xls")))
    cands = [p for p in files if all(k in p.name for k in CAND_KEYWORDS)]
    if not cands:
        raise RuntimeError(f"找不到同時包含 {CAND_KEYWORDS} 的 Excel 檔")

    scored = []
    for p in cands:
        ym = infer_ym_from_filename(p)
        if ym: scored.append((ym[0], ym[1], p))
    
    if scored:
        scored.sort()
        return scored[-1][2], (scored[-1][0], scored[-1][1])
    
    # 若無日期格式，依修改時間
    latest = sorted(cands, key=lambda x: x.stat().st_mtime)[-1]
    return latest, None

# =========================================================
# 2) 讀取與資料整理
# =========================================================
@st.cache_data # ✨ 加上快取機制
def read_source_data(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    engine = "openpyxl" if path.suffix.lower() == ".xlsx" else "xlrd"
    raw_xls = pd.read_excel(path, header=None, engine=engine)
    
    header_row = 0
    for i in range(min(50, len(raw_xls))):
        row_str = " ".join(raw_xls.iloc[i].astype(str))
        if "領域" in row_str and "總計" in row_str:
            header_row = i
            break
            
    df = raw_xls.iloc[header_row+1:].copy()
    df.columns = raw_xls.iloc[header_row].astype(str).str.strip()
    return df

def to_num(series):
    return pd.to_numeric(series.astype(str).str.replace(",", "").replace({"-": "0", "nan": "0"}), errors='coerce').fillna(0)

def build_fig5_data(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip().replace("\n", "") for c in df.columns]
    
    # 找領域欄與數值欄
    domain_col = next((c for c in df.columns if any(k in c for k in ["領域", "專長"])), None)
    val_col = next((c for c in df.columns if c == "總計" or "總計" in c), None)
    
    if not domain_col:
        raise RuntimeError("找不到領域欄位")

    df[domain_col] = df[domain_col].astype(str).str.replace("\u3000", "", regex=False).str.strip()
    df = df[~df[domain_col].isin(["", "總計", "合計"])].copy()
    
    if val_col:
        df["total_val"] = to_num(df[val_col])
    else:
        c_in = next((c for c in df.columns if "境內" in c), None)
        c_out = next((c for c in df.columns if "境外" in c), None)
        df["total_val"] = to_num(df[c_in]) + to_num(df[c_out])

    g = df.groupby(domain_col)["total_val"].sum().reset_index()
    g.columns = ["領域", "total"]

    base = pd.DataFrame({"領域": DOMAIN_ORDER}).merge(g, on="領域", how="left").fillna(0)
    other_sum = base[base["領域"].isin(OTHER_GROUP)]["total"].sum()
    other_label = "其他（建築設計、法律、國防、運動、環境、生技）"
    
    main = base[~base["領域"].isin(OTHER_GROUP)].copy()
    main = pd.concat([main, pd.DataFrame([{"領域": other_label, "total": other_sum}])], ignore_index=True)
    main = main.sort_values("total", ascending=False).reset_index(drop=True)
    
    total_all = main["total"].sum() if main["total"].sum() > 0 else 1
    main["pct"] = main["total"] / total_all
    return main

# =========================================================
# 3) 繪圖與自動換行工具 (✨ 修正字體並回傳 fig)
# =========================================================
def _wrap_label(text_obj, fig, max_px):
    renderer = fig.canvas.get_renderer()
    s = text_obj.get_text()
    if not s or text_obj.get_window_extent(renderer).width <= max_px:
        return
    lines, cur = [], ""
    for char in s:
        test = cur + char
        text_obj.set_text(test)
        if text_obj.get_window_extent(renderer).width > max_px:
            lines.append(cur)
            cur = char
        else: cur = test
    lines.append(cur)
    text_obj.set_text("\n".join(lines[:3]))

def plot_fig5(df_plot: pd.DataFrame):
    # ✨ 修正為 Linux 專用的思源黑體
    plt.rcParams["font.sans-serif"] = ["Noto Sans CJK JP", "sans-serif"]
    plt.rcParams["axes.unicode_minus"] = False
    
    fig, ax = plt.subplots(figsize=(10.5, 7.2), dpi=150)
    fig.subplots_adjust(right=0.75)

    values = df_plot["total"]
    colors = plt.cm.tab20(np.linspace(0, 1, len(values)))
    
    wedges, _ = ax.pie(values, startangle=90, counterclock=False, colors=colors,
                       wedgeprops=dict(edgecolor="white", linewidth=1.2))
    ax.axis("equal")

    # 右上角單位 (如果不想在網頁上顯示，也可以註解掉這行)
    fig.text(0.95, 0.95, "單位：人次", ha="right", va="top", fontsize=11)

    legend_labels = [f"{l}  {int(v):,} ({p*100:.0f}%)" for l, v, p in zip(df_plot["領域"], values, df_plot["pct"])]
    leg = ax.legend(wedges, legend_labels, loc="center left", bbox_to_anchor=(1, 0.5), frameon=False, fontsize=10)
    
    fig.canvas.draw()
    for t in leg.get_texts():
        _wrap_label(t, fig, 380)

    plt.tight_layout()
    return fig # ✨ 回傳給 Streamlit

# =========================================================
# 4) ✨ Streamlit 專屬渲染函式 (給 app.py 呼叫的入口)
# =========================================================
def render_streamlit(data_dir: Path):
    st.subheader("📊 圖5：有效就業金卡 (按領域分)")
    
    try:
        src, ym = pick_latest_source_file(data_dir)
        df_raw = read_source_data(str(src))
        df_plot = build_fig5_data(df_raw)
        
        if ym:
            y, m = ym
            st.markdown(f"**目前顯示資料時間：截至 {y}年{m}月**")
        else:
            st.markdown("**目前顯示資料時間：最新統計結果**")

        # 1. 顯示圓餅圖
        fig = plot_fig5(df_plot)
        st.pyplot(fig)
        
        # 2. ✨ 加碼顯示原始數據表格
        st.markdown("##### 📄 有效持卡領域明細")
        
        df_display = df_plot.copy()
        df_display.columns = ["申請領域", "有效人次", "佔比"]
        df_display["有效人次"] = df_display["有效人次"].apply(lambda x: f"{int(x):,}")
        df_display["佔比"] = df_display["佔比"].apply(lambda x: f"{x*100:.1f}%")
        
        st.dataframe(df_display, hide_index=True, width="stretch")
        
    except Exception as e:
        st.error(f"讀取或繪圖時發生錯誤：{e}")