import re
from pathlib import Path
from typing import Optional, Tuple, List

import pandas as pd
import streamlit as st # ✨ 新增 Streamlit 套件

# =========================================================
# 0) 參數設定
# =========================================================
TOP_N = 10  # 十大 + 其他

# =========================================================
# 1) 選檔與讀取邏輯 (加上快取)
# =========================================================
def guess_ym_from_filename(p: Path) -> Optional[pd.Timestamp]:
    name = p.stem
    # 搜尋 8 碼 YYYYMMDD 或 6 碼 YYYYMM
    m = re.search(r"(20\d{2})(0[1-9]|1[0-2])([0-3]\d)?", name)
    if m:
        return pd.Timestamp(int(m.group(1)), int(m.group(2)), 1)
    return None

def is_candidate_file(p: Path) -> bool:
    n = p.name
    if p.suffix.lower() not in (".xlsx", ".xls"):
        return False
    # 有效 + 國別/國籍 + 性別
    if "有效" not in n: return False
    if not (("國別" in n) or ("國籍" in n)): return False
    if not (("性別" in n) or ("男女" in n)): return False
    return guess_ym_from_filename(p) is not None

def pick_latest_file(base: Path) -> Tuple[Path, pd.Timestamp]:
    files = sorted(list(base.glob("*.xlsx")) + list(base.glob("*.xls")))
    pool = [(p, guess_ym_from_filename(p)) for p in files if is_candidate_file(p)]
    pool = [(p, ym) for (p, ym) in pool if ym is not None]

    if not pool:
        raise RuntimeError("資料夾內找不到符合表 13 的來源檔 (需含：有效、國別/國籍、性別)。")

    pool.sort(key=lambda x: x[1])
    return pool[-1][0], pool[-1][1]

@st.cache_data # ✨ 智慧表頭偵測快取
def read_and_detect_header(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    engine = "openpyxl" if path.suffix.lower() == ".xlsx" else "xlrd"
    raw = pd.read_excel(path, sheet_name=0, engine=engine, header=None)
    
    # 偵測表頭列 (尋找關鍵字密度最高的一列)
    best_i, best_score = 0, -1
    for i in range(min(80, len(raw))):
        row = raw.iloc[i].astype(str).fillna("").tolist()
        joined = " ".join(row)
        score = 0
        if ("國別" in joined) or ("國籍" in joined): score += 4
        if ("男" in joined) or ("男性" in joined): score += 3
        if ("女" in joined) or ("女性" in joined): score += 3
        if ("總計" in joined): score += 2
        if score > best_score:
            best_score, best_i = score, i
            
    header = raw.iloc[best_i].astype(str).str.replace("\n", "", regex=False).str.strip()
    df = raw.iloc[best_i + 1:].copy()
    df.columns = header
    df = df.dropna(axis=1, how="all").reset_index(drop=True)
    df.columns = [str(c).strip().replace("\n", "") for c in df.columns]
    return df

# =========================================================
# 2) 資料處理邏輯
# =========================================================
def to_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.strip().replace({"-": "0", "": "0"}), 
        errors="coerce"
    ).fillna(0)

def build_table13(df: pd.DataFrame, top_n: int = TOP_N) -> pd.DataFrame:
    d = df.copy()
    cols = list(d.columns)
    
    # 定位欄位
    col_country = next((c for c in cols if any(k in c for k in ["國籍", "國別", "國家"])), cols[0])
    col_male = next((c for c in cols if c in ("男", "男性")), None)
    col_female = next((c for c in cols if c in ("女", "女性")), None)

    if not col_male or not col_female:
        raise RuntimeError("找不到『男/女』欄位。")

    # 清理與彙整
    d[col_country] = d[col_country].astype(str).str.replace("\u3000", "", regex=False).str.strip()
    d = d[~d[col_country].isin(["", "總計", "合計", "性別占比", "nan"])].copy()

    d["male_v"] = to_num(d[col_male])
    d["female_v"] = to_num(d[col_female])
    d["total_v"] = d["male_v"] + d["female_v"]

    g = d.groupby(col_country, as_index=False)[["male_v", "female_v", "total_v"]].sum()
    g = g.sort_values("total_v", ascending=False).reset_index(drop=True)

    # 十大 + 其他
    top = g.head(top_n).copy()
    rest = g.iloc[top_n:].copy()

    if not rest.empty:
        other_row = pd.DataFrame([{
            col_country: "其他",
            "male_v": float(rest["male_v"].sum()),
            "female_v": float(rest["female_v"].sum()),
            "total_v": float(rest["total_v"].sum()),
        }])
        top = pd.concat([top, other_row], ignore_index=True)

    # 總計與性別占比
    grand_male = top["male_v"].sum()
    grand_female = top["female_v"].sum()
    grand_total = top["total_v"].sum()
    denom = grand_total if grand_total > 0 else 1.0

    # 建立輸出列
    rows = []
    for _, r in top.iterrows():
        rows.append({
            "國籍": r[col_country],
            "男性": int(r["male_v"]),
            "女性": int(r["female_v"]),
            "總計": int(r["total_v"])
        })
    
    # 加入總計列
    rows.append({
        "國籍": "總計",
        "男性": int(grand_male),
        "女性": int(grand_female),
        "總計": int(grand_total)
    })
    
    # 加入性別占比列
    rows.append({
        "國籍": "性別占比",
        "男性": f"{grand_male/denom*100:.1f}%",
        "女性": f"{grand_female/denom*100:.1f}%",
        "總計": "100%"
    })

    return pd.DataFrame(rows)

# =========================================================
# 3) ✨ Streamlit 專屬渲染函式 (入口)
# =========================================================
def render_streamlit(data_dir: Path):
    st.subheader("📊 表13：有效許可人次 (按國別及性別分)")

    try:
        source_file, used_ym = pick_latest_file(data_dir)
        st.markdown(f"**目前顯示資料時間：截至 {used_ym.year}年{used_ym.month}月底**")

        # 讀取與處理
        df_raw = read_and_detect_header(str(source_file))
        table13_df = build_table13(df_raw, top_n=TOP_N)

        # 格式化數值 (加上千分位)
        df_display = table13_df.copy()
        for col in ["男性", "女性", "總計"]:
            df_display[col] = df_display[col].apply(lambda x: f"{x:,}" if isinstance(x, (int, float)) else x)

        # ✨ 渲染網頁表格
        st.dataframe(df_display, hide_index=True, width="stretch")
        
        st.info("💡 註：本表列出前十大國籍，其餘歸類為「其他」。最後一列為性別占比分析。")

    except Exception as e:
        st.error(f"讀取或產表時發生錯誤：{e}")