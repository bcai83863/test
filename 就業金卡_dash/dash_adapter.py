from __future__ import annotations

import base64
import importlib
import io
import logging
import sys
import warnings
import traceback
from pathlib import Path
from typing import Any

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from dash import dash_table, dcc, html
from matplotlib.figure import Figure

# =========================================================
# 1) 字型修復與繪圖環境初始化
# =========================================================
FONT_FILENAME = "NotoSansTC-VariableFont_wght.ttf"
FONT_PATH = Path(__file__).parent / "assets" / "fonts" / FONT_FILENAME

def _setup_matplotlib_fonts():
    """強制讓 Matplotlib 在 Linux 環境下正確顯示中文"""
    plt.switch_backend('Agg') # 伺服器無介面模式
    if FONT_PATH.exists():
        try:
            fe = fm.FontEntry(fname=str(FONT_PATH), name='Noto Sans TC')
            fm.fontManager.ttflist.insert(0, fe)
            plt.rcParams['font.family'] = 'sans-serif'
            plt.rcParams['font.sans-serif'] = [fe.name, 'DejaVu Sans']
            plt.rcParams['axes.unicode_minus'] = False
            print(f"✅ 已成功載入字型：{FONT_PATH}")
        except Exception as e:
            print(f"⚠️ 字型註冊失敗：{e}")
    else:
        print(f"❌ 找不到字型檔：{FONT_PATH}")

_setup_matplotlib_fonts()

# =========================================================
# 2) 報表選單設定 (Export for app.py)
# =========================================================
REPORT_MENU: list[tuple[str, str]] = [
    ("圖1: 累計核發人次趨勢", "figure_01"), ("圖2: 累計持卡人次-按領域分", "figure_02"),
    ("圖3: 累計持卡人次-十大國別", "figure_03"), ("圖4: 累計核發人次-年齡及性別", "figure_04"),
    ("圖5: 有效持卡人次-按領域分", "figure_05"), ("圖6: 有效持卡人次-十大國別", "figure_06"),
    ("圖7: 有效持卡人次-年齡及性別", "figure_07"), ("表1: 歷年累計許可人次 (總覽)", "table_01"),
    ("表2: 歷年累計許可人次 (初次申請)", "table_02"), ("表3: 歷年累計許可人次 (領域分)", "table_03"),
    ("表4: 累計許可人次 (領域x性別)", "table_04"), ("表5: 累計許可人次 (國別x性別)", "table_05"),
    ("表6: 有效持卡十大國別 (國別x領域)", "table_06"), ("表7: 累計許可人次 (年齡x性別)", "table_07"),
    ("表8: 歷年有效許可人次 (外籍x港澳)", "table_08"), ("表9: 有效許可人次 (領域x境內外)", "table_09"),
    ("表10: 有效許可人次 (有效領域x性別)", "table_10"), ("表11: 有效許可人次 (有效國別x境內外)", "table_11"),
    ("表12: 有效許可人次 (有效國別x領域)", "table_12"), ("表13: 有效許可人次 (有效國別x性別)", "table_13"),
    ("表14: 有效許可人次 (有效年齡x性別)", "table_14"),
]

REPORT_LABEL_BY_MODULE = {module: label for label, module in REPORT_MENU}
LOGGER = logging.getLogger(__name__)

# =========================================================
# 3) 虛擬容器與紀錄器 (核心邏輯：過濾 CSS 並攔截調用)
# =========================================================
class StreamlitRecorder:
    def __init__(self, requested_ym: str | None = None):
        self.requested_ym = (requested_ym or "").strip()
        self.events: list[tuple[str, Any]] = []
        self.sidebar = self # 支援 st.sidebar.write

    def _record(self, kind: str, payload: Any):
        self.events.append((kind, payload))

    # --- ✨ 樣式清洗與文字輸出 ---
    def markdown(self, text: Any, **_: Any):
        content = str(text).strip()
        # 🚫 過濾掉專門針對 Streamlit UI 的 CSS Hacks
        if content.startswith("<style") or "font-family" in content or "data-testid" in content:
            print("🧹 [Cleanup] 已過濾不相容的 CSS/HTML 片段")
            return
        self._record("markdown", content)

    def title(self, text: Any, **_: Any): self._record("title", str(text))
    def subheader(self, text: Any, **_: Any): self._record("subheader", str(text))
    def info(self, text: Any, **_: Any): self._record("info", str(text))
    
    def write(self, *args: Any, **_: Any):
        for arg in args:
            if isinstance(arg, pd.DataFrame): self.dataframe(arg)
            elif isinstance(arg, (Figure, plt.Figure)): self.pyplot(arg)
            else: self.markdown(str(arg))

    # --- ✨ 確保運算邏輯不中斷 ---
    def cache_data(self, *args, **kwargs): return lambda f: f
    def cache_resource(self, *args, **kwargs): return lambda f: f

    def text_input(self, label: str, value: str = "", **_: Any) -> str:
        return self.requested_ym if self.requested_ym else value

    def dataframe(self, data: Any, **_: Any):
        if hasattr(data, "data"): data = data.data # 處理 Styler
        if isinstance(data, pd.DataFrame): self._record("dataframe", data)

    def table(self, data: Any, **_: Any): self.dataframe(data)

    def pyplot(self, fig: Any = None, **_: Any):
        if fig is None or fig is plt or str(type(fig)) == "<class 'module'>":
            fig = plt.gcf()
        if not isinstance(fig, Figure): fig = plt.gcf()
        self._record("pyplot", fig)
        plt.clf() # 清理畫布，避免下一張圖重疊

    # --- 佈局組件模擬 (支援解構語法) ---
    def columns(self, spec, **_):
        n = spec if isinstance(spec, int) else len(spec)
        return [self] * n
    def container(self, **_): return self
    def expander(self, *args, **kwargs): return self
    def __enter__(self): return self
    def __exit__(self, *args): pass
    def __getattr__(self, name: str): return lambda *args, **kwargs: None

# =========================================================
# 4) 組件轉換邏輯 (Matplotlib To Base64 & DataTable)
# =========================================================
def _matplotlib_figure_to_data_uri(fig: Figure) -> str:
    if FONT_PATH.exists():
        prop = fm.FontProperties(fname=str(FONT_PATH))
        for ax in fig.get_axes():
            ax.title.set_fontproperties(prop)
            ax.xaxis.label.set_fontproperties(prop)
            ax.yaxis.label.set_fontproperties(prop)
            for label in ax.get_xticklabels() + ax.get_yticklabels():
                label.set_fontproperties(prop)
            if ax.get_legend():
                for text in ax.get_legend().get_texts(): text.set_fontproperties(prop)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=140, bbox_inches="tight", facecolor='white')
    buf.seek(0)
    img_b64 = base64.b64encode(buf.read()).decode("ascii")
    plt.close(fig)
    return f"data:image/png;base64,{img_b64}"

def _event_to_component(kind: str, val: Any, idx: int):
    if kind == "title": return html.H2(val, style={"marginTop": "10px"})
    if kind == "subheader": return html.H4(val, style={"color": "#2c3e50", "marginTop": "20px"})
    if kind == "markdown": return dcc.Markdown(val, dangerously_allow_html=True)
    if kind == "info": return html.Div(val, style={"padding": "12px", "backgroundColor": "#e7f3fe", "borderLeft": "5px solid #2196F3", "marginBottom": "10px"})
    if kind == "dataframe":
        df = val.copy().fillna("") if isinstance(val, pd.DataFrame) else pd.DataFrame(val)
        return dash_table.DataTable(
            data=df.to_dict("records"),
            columns=[{"name": str(c), "id": str(c)} for c in df.columns],
            page_size=15,
            style_table={"overflowX": "auto", "marginBottom": "20px"},
            style_cell={"textAlign": "left", "padding": "8px", "fontFamily": "Noto Sans TC, sans-serif"},
            style_header={"backgroundColor": "#f8f9fa", "fontWeight": "bold"}
        )
    if kind == "pyplot":
        src = _matplotlib_figure_to_data_uri(val)
        return html.Img(src=src, style={"maxWidth": "100%", "marginTop": "10px", "border": "1px solid #eee", "borderRadius": "8px"})
    return None

# =========================================================
# 5) 主入口：渲染報表至 Dash (劫持模組層級)
# =========================================================
def render_report_to_dash(module_name: str, data_dir: Path, requested_ym: str | None = None):
    # 確保搜尋路徑包含 Excel 資料夾
    if str(data_dir) not in sys.path: sys.path.insert(0, str(data_dir))
    
    recorder = StreamlitRecorder(requested_ym=requested_ym)
    
    # ✨ 核心劫持：強制所有 import streamlit 導向我們的 recorder
    original_st = sys.modules.get('streamlit')
    sys.modules['streamlit'] = recorder 

    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            # 重新載入報表模組，確保它吃到偽裝的 st
            if module_name in sys.modules: del sys.modules[module_name]
            module = importlib.import_module(module_name)
            
            # 🚀 執行報表主程式
            module.render_streamlit(data_dir)
            print(f"✅ 報表 {module_name} 執行成功，抓取到 {len(recorder.events)} 個事件")
    except Exception:
        print(f"❌ 執行報表 {module_name} 失敗：")
        traceback.print_exc() # 將詳細錯誤印在 Render 日誌中
        return [html.Div("報表執行中發生錯誤，請檢查系統日誌。", style={"color": "red"})]
    finally:
        # 恢復環境，避免污染其他地方
        if original_st: sys.modules['streamlit'] = original_st
        else: del sys.modules['streamlit']

    # 將事件轉化為 Dash 元件
    components = []
    for i, (kind, val) in enumerate(recorder.events):
        comp = _event_to_component(kind, val, i)
        if comp: components.append(comp)
    
    return components if components else [html.Div("⚠️ 執行完畢但無可顯示內容。")]