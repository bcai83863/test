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
TOP_N = 10  # 十大國別

# =========================================================
# 1) 檔案挑選邏輯
# =========================================================
def guess_ym_from_filename(p: Path) -> Optional[pd.Timestamp]:
    name = p.stem
    m1 = re.search(r"(20\d{2})(0[1-9]|1[0-2])([0-3]\d)", name)
    if m1: return pd.Timestamp(int(m1.group(1)), int(m1.group(2)), 1)
    m2 = re.search(r"(20\d{2})(0[1-9]|1[0-2])", name)
    if m2: return pd.Timestamp(int(m2.group(1)), int(m2.group(2)), 1)
    return None

def is_candidate_file(p: Path) -> bool:
    n = p.name
    if p.suffix.lower() not in (".xlsx", ".xls"): return False
    if "有效" not in n: return False
    if not any(k in n for k in ["國別", "國籍"]): return False
    if not ("境內外" in n or ("境內" in n and "境外" in n)): return False
    return True

def pick_latest_file(base: Path) -> Tuple[Path, pd.Timestamp]:
    files = [p for p in base.iterdir() if is_candidate_file(p)]
    if not files:
        raise FileNotFoundError("找不到符合條件的『有效持卡國別』Excel 檔案")
    pool = []
    for p in files:
        ym = guess_ym_from_filename(p)
        if ym: pool.append((p, ym))
    if not pool:
        latest_p = max(files, key=lambda x: x.stat().st_mtime)
        return latest_p, pd.Timestamp.now()
    pool.sort(key=lambda x: x[1])
    return pool[-1][0], pool[-1][1]

# =========================================================
# 2) 讀取與資料處理
# =========================================================
@st.cache_data # ✨ 加上快取機制
def read_and_clean_excel(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    engine = "openpyxl" if path.suffix.lower() == ".xlsx" else "xlrd"
    raw_xls = pd.read_excel(path, header=None, engine=engine)
    h_row = 0
    for i in range(min(60, len(raw_xls))):
        row_str = " ".join(raw_xls.iloc[i].astype(str).fillna(""))
        if "國別" in row_str or "國籍" in row_str:
            h_row = i; break
            
    df_raw = raw_xls.iloc[h_row+1:].copy()
    df_raw.columns = raw_xls.iloc[h_row].astype(str).str.strip()
    return df_raw

def to_num(series):
    return pd.to_numeric(series.astype(str).str.replace(",", "").replace({"-": "0"}), errors='coerce').fillna(0)

def build_fig6_source(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip() for c in df.columns]
    country_col = next((c for c in df.columns if any(k in c for k in ["國別", "國籍"])), None)
    if not country_col: raise RuntimeError("找不到國別欄位")

    df[country_col] = df[country_col].astype(str).str.replace("\u3000", "", regex=False).str.strip()
    df = df[~df[country_col].isin(["", "總計", "合計", "性別占比", "nan"])].copy()

    # 這裡原本的邏輯有潛在問題：如果 df 裡面沒有「總計」欄位，它會去抓「境內」跟「境外」相加
    # 但 if "總計" in df.columns 這個判斷式，其實是在檢查字串陣列裡面有沒有 "總計"
    if "總計" in df.columns:
        df["total_val"] = to_num(df["總計"])
    else:
        df["total_val"] = to_num(df["境內"]) + to_num(df["境外"])

    g = df.groupby(country_col)["total_val"].sum().reset_index()
    g = g.sort_values("total_val", ascending=False).reset_index(drop=True)
    
    top = g.head(TOP_N).copy()
    rest = g.iloc[TOP_N:].copy()
    
    if not rest.empty:
        other = pd.DataFrame([{country_col: "其他", "total_val": rest["total_val"].sum()}])
        top = pd.concat([top, other], ignore_index=True)
        
    top.columns = ["國別", "total"]
    
    # 計算佔比
    total_all = top["total"].sum() if top["total"].sum() > 0 else 1
    top["pct"] = top["total"] / total_all
    
    return top

# =========================================================
# 3) 繪圖與自動換行工具 (✨ 修正字體並回傳 fig)
# =========================================================
def _wrap_label(text_obj, fig, max_px):
    renderer = fig.canvas.get_renderer()
    s = text_obj.get_text()
    if not s or text_obj.get_window_extent(renderer).width <= max_px: return
    lines, cur = [], ""
    for char in s:
        test = cur + char
        text_obj.set_text(test)
        if text_obj.get_window_extent(renderer).width > max_px:
            lines.append(cur); cur = char
        else: cur = test
    lines.append(cur)
    text_obj.set_text("\n".join(lines[:3]))

def plot_fig6_pie(df_top: pd.DataFrame):
    # ✨ 修正為 Linux 專用的思源黑體
    plt.rcParams["font.sans-serif"] = ["Noto Sans CJK JP", "sans-serif"]
    plt.rcParams["axes.unicode_minus"] = False
    
    fig, ax = plt.subplots(figsize=(8, 6), dpi=150)
    values = df_top["total"].values
    total_sum = values.sum()
    outer_labels = [f"{n}\n{v/total_sum*100:.0f}%" for n, v in zip(df_top["國別"], values)]
    colors = plt.cm.tab20(np.linspace(0, 1, len(values)))
    
    wedges, texts = ax.pie(values, labels=outer_labels, startangle=90, counterclock=False,
                           colors=colors, labeldistance=1.15,
                           wedgeprops=dict(edgecolor="white", linewidth=1.2),
                           textprops=dict(fontsize=10))
    fig.canvas.draw()
    for t in texts: _wrap_label(t, fig, 140)
    
    for i, p in enumerate(wedges):
        if values[i] > total_sum * 0.02:
            ang = (p.theta2 - p.theta1)/2. + p.theta1
            y, x = np.sin(np.deg2rad(ang)), np.cos(np.deg2rad(ang))
            ax.text(0.72*x, 0.72*y, f"{int(values[i]):,}", color="white", ha='center', va='center', fontsize=9)
            
    ax.axis("equal")
    plt.tight_layout()
    return fig # ✨ 回傳給 Streamlit

# =========================================================
# 4) ✨ Streamlit 專屬渲染函式 (給 app.py 呼叫的入口)
# =========================================================
def render_streamlit(data_dir: Path):
    st.subheader("📊 圖6：有效就業金卡十大國別")
    
    try:
        src, used_ym = pick_latest_file(data_dir)
        df_raw = read_and_clean_excel(str(src))
        df_source = build_fig6_source(df_raw)
        
        st.markdown(f"**目前顯示資料時間：截至 {used_ym.year}年{used_ym.month}月**")

        # 1. 顯示圓餅圖
        fig = plot_fig6_pie(df_source)
        st.pyplot(fig)
        
        # 2. ✨ 加碼顯示原始數據表格
        st.markdown("##### 📄 有效持卡十大國別明細")
        
        df_display = df_source.copy()
        df_display.columns = ["核發國別", "有效人次", "佔比"]
        df_display["有效人次"] = df_display["有效人次"].apply(lambda x: f"{int(x):,}")
        df_display["佔比"] = df_display["佔比"].apply(lambda x: f"{x*100:.1f}%")
        
        st.dataframe(df_display, hide_index=True, width="stretch")
        
    except Exception as e:
        st.error(f"讀取或繪圖時發生錯誤：{e}")