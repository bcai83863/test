import re
from pathlib import Path
from typing import Optional, Tuple

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

from font_utils import apply_cjk_font_settings, apply_streamlit_cjk_css

# =========================================================
# 0) 參數設定
# =========================================================
TOP_N = 10


# =========================================================
# 1) 選檔邏輯
# =========================================================
def guess_ym_from_filename(p: Path) -> Optional[pd.Timestamp]:
    name = p.stem
    m = re.search(r"(20\d{2})(0[1-9]|1[0-2])([0-3]\d)?", name)
    if not m:
        return None
    return pd.Timestamp(int(m.group(1)), int(m.group(2)), 1)


def is_candidate_file(p: Path) -> bool:
    n = p.name
    if p.suffix.lower() not in (".xlsx", ".xls"):
        return False
    return ("有效" in n) and (("國別" in n) or ("國籍" in n)) and ("境內外" in n)


def pick_latest_source_file(base: Path) -> Tuple[Path, pd.Timestamp]:
    files = sorted(list(base.glob("*.xlsx")) + list(base.glob("*.xls")))
    cands = [p for p in files if is_candidate_file(p)]
    if not cands:
        raise FileNotFoundError("找不到『有效 × 國籍/國別 × 境內外』來源檔案。")

    dated = []
    for p in cands:
        ym = guess_ym_from_filename(p)
        if ym is not None:
            dated.append((p, ym))

    if dated:
        dated.sort(key=lambda x: x[1])
        return dated[-1][0], dated[-1][1]

    latest = max(cands, key=lambda x: x.stat().st_mtime)
    return latest, pd.Timestamp.now()


# =========================================================
# 2) 讀取與資料整理
# =========================================================
@st.cache_data
def read_and_detect_header(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    engine = "openpyxl" if path.suffix.lower() == ".xlsx" else "xlrd"
    raw = pd.read_excel(path, sheet_name=0, engine=engine, header=None)

    best_i, best_score = 0, -1
    for i in range(min(80, len(raw))):
        row = raw.iloc[i].astype(str).fillna("").tolist()
        joined = " ".join(row)
        score = 0
        if ("國別" in joined) or ("國籍" in joined):
            score += 4
        if "境外" in joined:
            score += 3
        if "境內" in joined:
            score += 3
        if "總計" in joined:
            score += 1
        if score > best_score:
            best_score, best_i = score, i

    header = raw.iloc[best_i].astype(str).str.replace("\n", "", regex=False).str.strip()
    df = raw.iloc[best_i + 1 :].copy()
    df.columns = header
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all").reset_index(drop=True)
    return df


def to_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("\u3000", "", regex=False)
        .str.strip()
        .replace({"-": "0", "": "0", "nan": "0"}),
        errors="coerce",
    ).fillna(0)


def build_fig6_data(df: pd.DataFrame, top_n: int = TOP_N) -> pd.DataFrame:
    d = df.copy()
    d.columns = [str(c).strip().replace("\n", "") for c in d.columns]

    col_country = next((c for c in d.columns if any(k in c for k in ["國別", "國籍", "國家"])), d.columns[0])
    col_out = next((c for c in d.columns if "境外" in c), None)
    col_in = next((c for c in d.columns if "境內" in c), None)

    if not col_out or not col_in:
        raise RuntimeError("找不到『境外』或『境內』欄位。")

    d[col_country] = d[col_country].astype(str).str.replace("\u3000", "", regex=False).str.strip()
    d = d[~d[col_country].isin(["", "總計", "合計", "nan"])].copy()

    d["out_v"] = to_num(d[col_out])
    d["in_v"] = to_num(d[col_in])
    d["total_v"] = d["out_v"] + d["in_v"]

    g = d.groupby(col_country, as_index=False)[["out_v", "in_v", "total_v"]].sum()
    g = g.sort_values("total_v", ascending=False).reset_index(drop=True)

    top = g.head(top_n).copy()
    rest = g.iloc[top_n:].copy()

    if not rest.empty:
        top = pd.concat(
            [
                top,
                pd.DataFrame(
                    [
                        {
                            col_country: "其他",
                            "out_v": rest["out_v"].sum(),
                            "in_v": rest["in_v"].sum(),
                            "total_v": rest["total_v"].sum(),
                        }
                    ]
                ),
            ],
            ignore_index=True,
        )

    denom = top["total_v"].sum()
    denom = denom if denom > 0 else 1
    top["share"] = top["total_v"] / denom
    top = top.rename(columns={col_country: "國籍"}).reset_index(drop=True)
    return top[["國籍", "out_v", "in_v", "total_v", "share"]]


# =========================================================
# 3) 繪圖
# =========================================================
def apply_font_settings():
    apply_cjk_font_settings()


def _wrap_label(text_obj, fig, max_px: int):
    renderer = fig.canvas.get_renderer()
    label = text_obj.get_text()
    if not label or text_obj.get_window_extent(renderer).width <= max_px:
        return

    lines, cur = [], ""
    for char in label:
        test_text = cur + char
        text_obj.set_text(test_text)
        if text_obj.get_window_extent(renderer).width > max_px:
            lines.append(cur)
            cur = char
        else:
            cur = test_text
    lines.append(cur)
    text_obj.set_text("\n".join(lines[:3]))


def plot_fig6(df_plot: pd.DataFrame):
    apply_font_settings()

    fig, ax = plt.subplots(figsize=(10, 6), dpi=150)
    values = df_plot["total_v"]
    labels = df_plot["國籍"]
    colors = plt.cm.tab20(np.linspace(0, 1, len(values)))

    wedges, _ = ax.pie(
        values,
        startangle=90,
        counterclock=False,
        colors=colors,
        wedgeprops=dict(edgecolor="white", linewidth=1),
    )
    ax.axis("equal")

    legend_labels = [f"{c} {int(v):,} ({s*100:.1f}%)" for c, v, s in zip(labels, values, df_plot["share"])]
    leg = ax.legend(
        wedges,
        legend_labels,
        loc="center left",
        bbox_to_anchor=(1, 0.5),
        frameon=False,
        fontsize=10,
    )

    fig.canvas.draw()
    for text in leg.get_texts():
        _wrap_label(text, fig, 330)

    plt.tight_layout()
    return fig


# =========================================================
# 4) Streamlit 入口
# =========================================================
def render_streamlit(data_dir: Path):
    apply_streamlit_cjk_css()
    st.subheader("📊 圖6：有效就業金卡前十大國籍")

    try:
        source_file, used_ym = pick_latest_source_file(data_dir)
        st.info(f"📅 目前顯示資料時間：截至 {used_ym.year}年{int(used_ym.month)}月底")

        df_raw = read_and_detect_header(str(source_file))
        df_plot = build_fig6_data(df_raw, top_n=TOP_N)

        fig = plot_fig6(df_plot)
        st.pyplot(fig)

        st.markdown("##### 📄 前十大國籍明細")
        df_display = df_plot.copy()
        df_display.columns = ["國籍", "境外人次", "境內人次", "總計", "占比"]
        for col in ["境外人次", "境內人次", "總計"]:
            df_display[col] = df_display[col].apply(lambda x: f"{int(x):,}")
        df_display["占比"] = df_display["占比"].apply(lambda x: f"{x*100:.2f}%")
        st.dataframe(df_display, hide_index=True, width="stretch")

    except Exception as e:
        st.error(f"讀取或繪圖時發生錯誤：{e}")
