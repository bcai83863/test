import re
from pathlib import Path
from typing import Optional, Tuple

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st 

from font_utils import apply_cjk_font_settings, apply_streamlit_cjk_css

# =========================================================
# 1) 自動尋找最新來源檔
# =========================================================
RE_YYYYMMDD = re.compile(r"(20\d{2})(0[1-9]|1[0-2])([0-3]\d)")
RE_YYYYMM = re.compile(r"(20\d{2})(0[1-9]|1[0-2])")

def _ym_from_name(p: Path) -> Optional[pd.Timestamp]:
    s = p.stem
    m = RE_YYYYMMDD.search(s)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        return pd.Timestamp(y, mo, 1)
    m2 = RE_YYYYMM.search(s)
    if m2:
        y, mo = int(m2.group(1)), int(m2.group(2))
        return pd.Timestamp(y, mo, 1)
    return None

def pick_latest_age_gender_file(base: Path) -> Tuple[Path, pd.Timestamp]:
    files = sorted(list(base.glob("*.xlsx")) + list(base.glob("*.xls")))
    cands: list[tuple[Path, pd.Timestamp]] = []
    for p in files:
        n = p.name
        if ("有效" in n) and ("年齡" in n):
            ym = _ym_from_name(p)
            if ym is not None:
                cands.append((p, ym))

    if not cands:
        raise FileNotFoundError(
            "找不到檔名同時包含「有效」與「年齡」的 Excel 檔案。\n"
            "請確認檔案已上傳至資料夾。"
        )

    cands.sort(key=lambda x: x[1])
    return cands[-1][0], cands[-1][1]

# =========================================================
# 2) matplotlib 字型設定 (解決 Linux/Windows 相容性)
# =========================================================
def apply_font_settings():
    """套用跨平台 CJK 字型設定，確保 Streamlit Linux 可正確顯示中文。"""
    apply_cjk_font_settings()

# =========================================================
# 3) 讀取 Excel 與表頭自動偵測
# =========================================================
@st.cache_data
def load_best_sheet(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    engine = "openpyxl" if path.suffix.lower() == ".xlsx" else "xlrd"
    xls = pd.ExcelFile(path, engine=engine)
    
    best_df = None
    best_score = -1

    for sh in xls.sheet_names:
        try:
            raw = pd.read_excel(path, sheet_name=sh, engine=engine, header=None)
            
            # 權重偵測表頭列
            hdr_i, score_max = 0, -1
            for i in range(min(150, len(raw))):
                row = raw.iloc[i].astype(str).fillna("").tolist()
                joined = " ".join(row)
                score = 0
                if "年齡" in joined: score += 3
                if ("年齡區間" in joined) or ("年齡層" in joined): score += 4
                if ("男性" in joined) or ("男" in joined): score += 2
                if ("女性" in joined) or ("女" in joined): score += 2
                if "總計" in joined: score += 1
                if score > score_max:
                    score_max, hdr_i = score, i
            
            header = raw.iloc[hdr_i].astype(str).str.replace("\n", "", regex=False).str.strip()
            df = raw.iloc[hdr_i + 1 :].copy()
            df.columns = header
            df = df.dropna(axis=1, how="all").reset_index(drop=True)
            df = df.dropna(axis=0, how="all").reset_index(drop=True)

            cols = [str(c).strip() for c in df.columns]
            sh_score = 0
            if any("年齡區間" in c or "年齡層" in c or c == "年齡" for c in cols): sh_score += 6
            if any(c == "男性" or c == "男" or "男性" in c for c in cols): sh_score += 4
            if any(c == "女性" or c == "女" or "女性" in c for c in cols): sh_score += 4
            if any("總計" in c for c in cols): sh_score += 1

            if sh_score > best_score:
                best_score, best_df = sh_score, df
        except Exception:
            continue

    if best_df is None:
        raise RuntimeError("Excel 讀取失敗：找不到可辨識的數據工作表。")

    return best_df

# =========================================================
# 4) 資料整理與歸一化
# =========================================================
def _find_col(cols, keywords):
    for c in cols:
        cc = str(c).strip()
        if any(k == cc or k in cc for k in keywords): return c
    return None

def to_num(s: pd.Series) -> pd.Series:
    x = s.astype(str).str.replace(",", "", regex=False).str.replace("\u3000", "", regex=False).str.strip()
    x = x.replace({"nan": "0", "NaN": "0", "-": "0", "－": "0", "—": "0"})
    return pd.to_numeric(x, errors="coerce").fillna(0)

def canon_age_bucket(x: str) -> Optional[str]:
    s = str(x or "").replace("\u3000", "").replace(" ", "").replace("\n", "")
    s = s.replace("～", "-").replace("〜", "-").replace("~", "-").replace("—", "-").replace("–", "-")
    if "未滿" in s or "<20" in s or "小於20" in s: return "20歲以下"
    if "70" in s and ("以上" in s or "↑" in s): return "70歲以上"
    
    m = re.search(r"(20|30|40|50|60)\s*歲?\s*-\s*(29|39|49|59|69)\s*歲?", s)
    if m: return f"{m.group(1)}-{m.group(2)}歲"
    return None

def build_age_gender_series(df: pd.DataFrame) -> Tuple[list[str], np.ndarray, np.ndarray]:
    d = df.copy()
    d.columns = [str(c).strip().replace("\n", "") for c in d.columns]
    cols = list(d.columns)

    col_age = _find_col(cols, ["年齡區間", "年齡層", "年齡"])
    col_male = _find_col(cols, ["男性", "男"])
    col_female = _find_col(cols, ["女性", "女"])

    if not col_age or not col_male or not col_female:
        raise RuntimeError(f"找不到必要欄位。現有欄位：{cols}")

    d[col_age] = d[col_age].astype(str).str.replace("\u3000", "", regex=False).str.strip()
    bad = {"總計", "性別占比", "年齡占比", "nan"}
    d = d[~d[col_age].isin(bad)].copy()

    d["male"] = to_num(d[col_male])
    d["female"] = to_num(d[col_female])
    d["bucket"] = d[col_age].map(canon_age_bucket)
    d = d.dropna(subset=["bucket"]).copy()

    order = ["20歲以下", "20-29歲", "30-39歲", "40-49歲", "50-59歲", "60-69歲", "70歲以上"]
    g = d.groupby("bucket", as_index=False)[["male", "female"]].sum()
    g = g.set_index("bucket").reindex(order).fillna(0)

    return order, g["male"].to_numpy(), g["female"].to_numpy()

# =========================================================
# 5) 繪圖邏輯
# =========================================================
def plot_figure7(labels: list[str], male: np.ndarray, female: np.ndarray):
    apply_font_settings() # ✨ 套用字型設定

    labels_display = [l.replace("歲", "") for l in labels]
    x = np.arange(len(labels_display))
    totals = male + female

    fig, ax = plt.subplots(figsize=(10, 6), dpi=150)

    ax.bar(x, male, width=0.6, label="男", color="#095EDB")
    ax.bar(x, female, width=0.6, bottom=male, label="女", color="#FF3300")

    ax.set_xlabel("年齡(歲)", fontweight="bold", fontsize=12)
    ax.set_ylabel("人次", fontweight="bold", fontsize=12, rotation=0, labelpad=20)

    ax.set_xticks(x)
    ax.set_xticklabels(labels_display, fontsize=10)
    ax.tick_params(axis="y", labelsize=10)

    # 標註總數值
    max_val = np.max(totals) if len(totals) > 0 else 1
    for xi, t in zip(x, totals):
        if t > 0:
            ax.text(xi, t + (max_val * 0.01), f"{int(t):,}", ha="center", va="bottom", fontsize=10)

    ax.legend(loc="upper right", frameon=False)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    plt.tight_layout()
    return fig

# =========================================================
# 6) ✨ Streamlit 入口函式
# =========================================================
def render_streamlit(data_dir: Path):
    apply_streamlit_cjk_css()
    st.subheader("📊 圖7：有效就業金卡年齡及性別統計")
    
    try:
        src, ym = pick_latest_age_gender_file(data_dir)
        df = load_best_sheet(str(src))
        labels, male, female = build_age_gender_series(df)
        
        # 顯示日期資訊
        st.info(f"📅 目前顯示資料時間：截至 {ym.year}年{int(ym.month)}月")

        # 1. 顯示堆疊長條圖
        fig = plot_figure7(labels, male, female)
        st.pyplot(fig)
        
        # 2. 顯示原始數據表格
        st.markdown("##### 📄 有效持卡年齡及性別人次明細")
        df_display = pd.DataFrame({
            "年齡級距": labels,
            "男性 (人次)": male,
            "女性 (人次)": female,
            "總計 (人次)": male + female
        })
        
        # 格式化數字
        for col in ["男性 (人次)", "女性 (人次)", "總計 (人次)"]:
            df_display[col] = df_display[col].apply(lambda x: f"{int(x):,}")
            
        st.dataframe(df_display, hide_index=True, width="stretch")
        
    except Exception as e:
        st.error(f"讀取或繪圖時發生錯誤：{e}")
