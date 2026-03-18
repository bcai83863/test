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
# 1) 字型修復與繪圖環境初始化
# =========================================================
FONT_FILENAME = "NotoSansTC-VariableFont_wght.ttf"
FONT_PATH = Path(__file__).parent / "assets" / "fonts" / FONT_FILENAME

def _setup_matplotlib_fonts():
    plt.switch_backend('Agg') # 伺服器環境必備
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

_setup_matplotlib_fonts()

# =========================================================
# 2) 報表選單
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
# 3) Streamlit 指令模擬器 (強化版)
# =========================================================
class StreamlitRecorder:
    def __init__(self, requested_ym: str | None = None):
        self.requested_ym = (requested_ym or "").strip()
        self.events: list[tuple[str, Any]] = []

    def _record(self, kind: str, payload: Any):
        self.events.append((kind, payload))

    # 基礎文字輸出
    def title(self, text: Any, **_: Any): self._record("title", str(text))
    def subheader(self, text: Any, **_: Any): self._record("subheader", str(text))
    def markdown(self, text: Any, **_: Any): self._record("markdown", str(text))
    def info(self, text: Any, **_: Any): self._record("info", str(text))
    def error(self, text: Any, **_: Any): self._record("error", str(text))
    def write(self, *args: Any, **_: Any):
        for arg in args:
            if isinstance(arg, pd.DataFrame): self.dataframe(arg)
            else: self._record("markdown", str(arg))

    # 輸入攔截
    def text_input(self, label: str, value: str = "", **_: Any) -> str:
        return self.requested_ym if self.requested_ym else value

    # ✨ 數據輸出支援 (同時支援 dataframe 與 table)
    def dataframe(self, data: Any, **_: Any):
        if hasattr(data, "data"): data = data.data # 處理 Styler 物件
        if isinstance(data, pd.DataFrame): self._record("dataframe", data)

    def table(self, data: Any, **_: Any):
        self.dataframe(data)

    # ✨ 關鍵修復：st.pyplot
    def pyplot(self, fig: Figure | None = None, **_: Any):
        if fig is None:
            # 如果報表裡只寫 st.pyplot()，主動抓取當前畫布
            fig = plt.gcf()
        
        # 深度拷貝一份 Figure，避免後續 plt.close() 影響到顯示
        new_fig = Figure(figsize=fig.get_size_inches(), dpi=fig.get_dpi())
        new_fig.canvas.draw()
        for ax in fig.get_axes():
            # 這是一個簡化做法，直接紀錄當前的 fig 內容
            pass
        
        self._record("pyplot", fig)
        # 紀錄完後清理全域畫布，避免下一張圖重疊
        plt.clf()

    def __getattr__(self, name: str):
        return lambda *args, **kwargs: None

# =========================================================
# 4) 圖片與表格轉換邏輯 (加入字型強制套用)
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
    fig.savefig(buf, format="png", dpi=160, bbox_inches="tight", facecolor='white')
    buf.seek(0)
    encoded = base64.b64encode(buf.read()).decode("ascii")
    plt.close(fig) # 釋放記憶體
    return f"data:image/png;base64,{encoded}"

def _event_to_component(kind: str, val: Any, idx: int):
    if kind == "title": return html.H2(val)
    if kind == "subheader": return html.H4(val, style={"color": "#2c3e50", "marginTop": "20px"})
    if kind == "markdown": return dcc.Markdown(val)
    if kind == "info": return html.Div(val, style={"padding": "10px", "backgroundColor": "#e7f3fe", "borderLeft": "5px solid #2196F3", "marginBottom": "10px"})
    if kind == "dataframe":
        df_display = val.copy().fillna("")
        return dash_table.DataTable(
            data=df_display.to_dict("records"),
            columns=[{"name": str(c), "id": str(c)} for c in df_display.columns],
            page_size=20,
            style_table={"overflowX": "auto", "marginBottom": "20px"},
            style_cell={"textAlign": "left", "padding": "8px", "fontFamily": "Noto Sans TC, sans-serif"},
            style_header={"backgroundColor": "#f8f9fa", "fontWeight": "bold"}
        )
    if kind == "pyplot":
        src = _matplotlib_figure_to_data_uri(val)
        return html.Img(src=src, style={"maxWidth": "100%", "marginTop": "10px", "border": "1px solid #eee"})
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
            module.st = recorder
            module.render_streamlit(data_dir)
        except Exception as e:
            return [html.Div(f"❌ 模組執行錯誤：{e}", style={"color": "red"})]

    components = []
    for i, (kind, val) in enumerate(recorder.events):
        comp = _event_to_component(kind, val, i)
        if comp: components.append(comp)
    
    return components if components else [html.Div("⚠️ 此報表執行成功但未產出任何內容。")]