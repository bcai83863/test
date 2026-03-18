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
import plotly.graph_objects as go
from matplotlib.figure import Figure

# 配置日誌
logging.basicConfig(level=logging.INFO)
LOGGER = logging.getLogger(__name__)

# =========================================================
# 1) 字型修復 (Render 環境必備)
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
            print(f"✅ 已成功載入字型：{FONT_PATH}")
        except Exception as e: print(f"⚠️ 字型註冊失敗：{e}")

_setup_matplotlib_fonts()

# =========================================================
# 2) 報表選單設定 (Export for app.py)
# =========================================================
REPORT_MENU = [
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

# =========================================================
# 3) 萬能紀錄器 (支援所有常見 Streamlit 指令)
# =========================================================
class StreamlitRecorder:
    def __init__(self, requested_ym: str | None = None):
        self.requested_ym = (requested_ym or "").strip()
        self.events: list[tuple[str, Any]] = []
        self.sidebar = self

    def _record(self, kind: str, payload: Any):
        print(f"🎬 [Recorder] 捕捉到 {kind}")
        self.events.append((kind, payload))

    # ✨ 文字與樣式處理
    def markdown(self, text: Any, **_: Any):
        content = str(text).strip()
        if content.startswith("<style") or "font-family" in content:
            print("🧹 [Cleanup] 已過濾不相容的 CSS/HTML 片段")
            return
        self._record("markdown", content)

    def title(self, t: Any, **_: Any): self._record("title", str(t))
    def subheader(self, t: Any, **_: Any): self._record("subheader", str(t))
    def info(self, t: Any, **_: Any): self._record("info", str(t))
    def success(self, t: Any, **_: Any): self._record("info", f"✅ {t}")
    def warning(self, t: Any, **_: Any): self._record("info", f"⚠️ {t}")

    # ✨ 核心輸出攔截
    def write(self, *args: Any, **_: Any):
        for arg in args:
            if isinstance(arg, pd.DataFrame) or hasattr(arg, 'to_dict'): self.dataframe(arg)
            elif isinstance(arg, (Figure, plt.Figure)): self.pyplot(arg)
            elif isinstance(arg, go.Figure): self.plotly_chart(arg)
            else: self.markdown(str(arg))

    def dataframe(self, data: Any, **_: Any):
        if hasattr(data, "data"): data = data.data
        self._record("dataframe", data)

    def table(self, data: Any, **_: Any): self.dataframe(data)

    def pyplot(self, fig: Any = None, **_: Any):
        if fig is None or fig is plt or str(type(fig)) == "<class 'module'>":
            fig = plt.gcf()
        if not isinstance(fig, Figure): fig = plt.gcf()
        self._record("pyplot", fig)
        plt.clf()

    def plotly_chart(self, fig: Any, **_: Any):
        self._record("plotly", fig)

    def image(self, image: Any, **_: Any):
        self._record("markdown", "🖼️ (圖片顯示暫不支援，請改用 pyplot 或 plotly)")

    # ✨ 運算與佈局模擬
    def cache_data(self, *args, **kwargs): return lambda f: f
    def cache_resource(self, *args, **kwargs): return lambda f: f
    def columns(self, spec, **_):
        n = spec if isinstance(spec, int) else len(spec)
        return [self] * n
    def container(self, **_): return self
    def expander(self, *args, **kwargs): return self
    def __enter__(self): return self
    def __exit__(self, *args): pass
    def text_input(self, label: str, value: str = "", **_: Any) -> str:
        return self.requested_ym if self.requested_ym else value
    def __getattr__(self, name: str): return lambda *args, **kwargs: None

# =========================================================
# 4) 組件轉換邏輯
# =========================================================
def _matplotlib_figure_to_data_uri(fig: Figure) -> str:
    if FONT_PATH.exists():
        prop = fm.FontProperties(fname=str(FONT_PATH))
        for ax in fig.get_axes():
            ax.title.set_fontproperties(prop)
            ax.xaxis.label.set_fontproperties(prop)
            ax.yaxis.label.set_fontproperties(prop)
            for label in ax.get_xticklabels() + ax.get_yticklabels(): label.set_fontproperties(prop)
            if ax.get_legend():
                for t in ax.get_legend().get_texts(): t.set_fontproperties(prop)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=140, bbox_inches="tight", facecolor='white')
    buf.seek(0)
    img_b64 = base64.b64encode(buf.read()).decode("ascii")
    plt.close(fig)
    return f"data:image/png;base64,{img_b64}"

def _event_to_component(kind: str, val: Any, idx: int):
    if kind == "title": return html.H2(val)
    if kind == "subheader": return html.H4(val, style={"color": "#2c3e50", "marginTop": "20px"})
    if kind == "markdown": return dcc.Markdown(val, dangerously_allow_html=True)
    if kind == "info": return html.Div(val, style={"padding": "12px", "backgroundColor": "#e7f3fe", "borderLeft": "5px solid #2196F3", "marginBottom": "10px"})
    if kind == "dataframe":
        df = val.copy().fillna("") if isinstance(val, pd.DataFrame) else pd.DataFrame(val).fillna("")
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
    if kind == "plotly":
        return dcc.Graph(figure=val)
    return None

# =========================================================
# 5) 主入口
# =========================================================
def render_report_to_dash(module_name: str, data_dir: Path, requested_ym: str | None = None):
    if str(data_dir) not in sys.path: sys.path.insert(0, str(data_dir))
    recorder = StreamlitRecorder(requested_ym=requested_ym)
    original_st = sys.modules.get('streamlit')
    sys.modules['streamlit'] = recorder 
    try:
        if module_name in sys.modules: del sys.modules[module_name]
        module = importlib.import_module(module_name)
        module.render_streamlit(data_dir)
        print(f"✅ 報表 {module_name} 執行完畢，抓取到 {len(recorder.events)} 個事件")
    except Exception:
        print(f"❌ 執行報表 {module_name} 失敗：")
        traceback.print_exc()
        return [html.Div("報表執行中發生錯誤，請檢查系統日誌。", style={"color": "red"})]
    finally:
        if original_st: sys.modules['streamlit'] = original_st
        else: del sys.modules['streamlit']

    components = []
    for i, (kind, val) in enumerate(recorder.events):
        comp = _event_to_component(kind, val, i)
        if comp: components.append(comp)
    return components if components else [html.Div("⚠️ 執行成功但無內容。")]