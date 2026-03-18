from __future__ import annotations

import base64
import importlib
import io
import logging
import sys
import warnings
from pathlib import Path
from typing import Any

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from dash import dash_table, dcc, html
from matplotlib.figure import Figure

# =========================================================
# 1) 字型修復與環境初始化
# =========================================================
FONT_FILENAME = "NotoSansTC-VariableFont_wght.ttf"
FONT_PATH = Path(__file__).parent / "assets" / "fonts" / FONT_FILENAME

def _setup_matplotlib_fonts():
    plt.switch_backend('Agg')
    if FONT_PATH.exists():
        try:
            fe = fm.FontEntry(fname=str(FONT_PATH), name='Noto Sans TC')
            fm.fontManager.ttflist.insert(0, fe)
            plt.rcParams['font.family'] = 'sans-serif'
            plt.rcParams['font.sans-serif'] = [fe.name, 'DejaVu Sans']
            plt.rcParams['axes.unicode_minus'] = False
            print(f"✅ 字型載入成功：{FONT_PATH}")
        except Exception as e:
            print(f"⚠️ 字型註冊失敗：{e}")

_setup_matplotlib_fonts()

# =========================================================
# 2) 虛擬容器類別 (處理 with st.expander... 語法)
# =========================================================
class DummyContext:
    """讓 with st.xxx(): 語法不報錯並能繼續執行"""
    def __enter__(self): return self
    def __exit__(self, *args): pass
    def __getitem__(self, item): return self # 處理 col1, col2 = st.columns(2)
    def __iter__(self): yield self; yield self; yield self; yield self # 支援解構

# =========================================================
# 3) 核心紀錄器 (StreamlitRecorder)
# =========================================================
class StreamlitRecorder:
    def __init__(self, requested_ym: str | None = None):
        self.requested_ym = (requested_ym or "").strip()
        self.events: list[tuple[str, Any]] = []
        self.sidebar = self # 支援 st.sidebar.write()

    def _record(self, kind: str, payload: Any):
        self.events.append((kind, payload))

    def title(self, text: Any, **_: Any): self._record("title", str(text))
    def subheader(self, text: Any, **_: Any): self._record("subheader", str(text))
    def markdown(self, text: Any, **_: Any): self._record("markdown", str(text))
    def info(self, text: Any, **_: Any): self._record("info", str(text))
    def error(self, text: Any, **_: Any): self._record("error", str(text))
    
    def write(self, *args: Any, **_: Any):
        for arg in args:
            if isinstance(arg, pd.DataFrame): self.dataframe(arg)
            elif isinstance(arg, (Figure, plt.Figure)): self.pyplot(arg)
            else: self._record("markdown", str(arg))

    def text_input(self, label: str, value: str = "", **_: Any) -> str:
        return self.requested_ym if self.requested_ym else value

    def dataframe(self, data: Any, **_: Any):
        if hasattr(data, "data"): data = data.data
        if isinstance(data, pd.DataFrame): self._record("dataframe", data)

    def table(self, data: Any, **_: Any): self.dataframe(data)

    def pyplot(self, fig: Any = None, **_: Any):
        if fig is None or fig is plt:
            fig = plt.gcf()
        # 確保傳入的是 Figure 物件
        if not isinstance(fig, Figure):
            # 嘗試抓取當前畫布
            fig = plt.gcf()
        
        # 複製一份，避免全域污染
        self._record("pyplot", fig)
        plt.clf() # 清理全域畫布

    # 處理佈局組件 (columns, tabs, expander, container)
    def columns(self, spec, **_): return DummyContext()
    def tabs(self, _): return [DummyContext()] * 10
    def expander(self, label, **_): return DummyContext()
    def container(self, **_): return DummyContext()

    def __getattr__(self, name: str):
        return lambda *args, **kwargs: None

# =========================================================
# 4) 轉換邏輯
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
            legend = ax.get_legend()
            if legend:
                for text in legend.get_texts(): text.set_fontproperties(prop)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=140, bbox_inches="tight", facecolor='white')
    buf.seek(0)
    encoded = base64.b64encode(buf.read()).decode("ascii")
    plt.close(fig)
    return f"data:image/png;base64,{encoded}"

def _event_to_component(kind: str, val: Any, idx: int):
    if kind == "title": return html.H2(val, style={"marginBottom": "10px"})
    if kind == "subheader": return html.H4(val, style={"color": "#2c3e50", "marginTop": "20px"})
    if kind == "markdown": return dcc.Markdown(val)
    if kind == "info": return html.Div(val, style={"padding": "12px", "backgroundColor": "#e7f3fe", "borderLeft": "5px solid #2196F3", "marginBottom": "15px"})
    if kind == "dataframe":
        df = val.copy().fillna("")
        return dash_table.DataTable(
            data=df.to_dict("records"),
            columns=[{"name": str(c), "id": str(c)} for c in df.columns],
            page_size=15,
            style_table={"overflowX": "auto", "marginBottom": "20px"},
            style_cell={"textAlign": "left", "padding": "10px", "fontFamily": "Noto Sans TC, sans-serif"},
            style_header={"backgroundColor": "#f8f9fa", "fontWeight": "bold"}
        )
    if kind == "pyplot":
        src = _matplotlib_figure_to_data_uri(val)
        return html.Img(src=src, style={"maxWidth": "100%", "marginTop": "15px", "border": "1px solid #eee", "borderRadius": "8px"})
    return None

# =========================================================
# 5) 主入口
# =========================================================
def render_report_to_dash(module_name: str, data_dir: Path, requested_ym: str | None = None):
    if str(data_dir) not in sys.path: sys.path.insert(0, str(data_dir))
    recorder = StreamlitRecorder(requested_ym=requested_ym)
    
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            module = importlib.import_module(module_name)
            module = importlib.reload(module)
            module.st = recorder # 💉 注入紀錄器
            
            # 確保報表模組能找到數據
            module.render_streamlit(data_dir)
            
            print(f"📊 報表模組 {module_name} 執行完成，收集到 {len(recorder.events)} 個事件")
        except Exception as e:
            LOGGER.exception("模組執行失敗")
            return [html.Div(f"❌ 報表邏輯執行失敗：{e}", style={"color": "red"})]

    components = []
    for i, (kind, val) in enumerate(recorder.events):
        comp = _event_to_component(kind, val, i)
        if comp: components.append(comp)
    
    if not components:
        return [html.Div("⚠️ 報表載入成功，但未偵測到任何可顯示的圖表或表格。請檢查報表是否使用了不支援的 st 指令。", style={"color": "orange", "padding": "20px"})]
        
    return components