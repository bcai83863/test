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
# 1) 字型修復：強制載入與路徑定義
# =========================================================
# 請確保字型檔位於 就業金卡_dash/assets/fonts/ 下
FONT_FILENAME = "NotoSansTC-VariableFont_wght.ttf"
FONT_PATH = Path(__file__).parent / "assets" / "fonts" / FONT_FILENAME

def _setup_matplotlib_fonts():
    """初始化繪圖環境與字型"""
    plt.switch_backend('Agg') # 伺服器環境必備
    
    if FONT_PATH.exists():
        try:
            # 註冊字型到 Matplotlib 的管理系統
            fe = fm.FontEntry(fname=str(FONT_PATH), name='Noto Sans TC')
            fm.fontManager.ttflist.insert(0, fe)
            
            # 設定全局字型優先順序
            plt.rcParams['font.family'] = 'sans-serif'
            plt.rcParams['font.sans-serif'] = [fe.name, 'DejaVu Sans']
            plt.rcParams['axes.unicode_minus'] = False # 修復負號顯示
            print(f"✅ 字型載入成功，路徑：{FONT_PATH}")
        except Exception as e:
            print(f"⚠️ 字型註冊失敗：{e}")
    else:
        print(f"❌ 找不到字型檔：{FONT_PATH}")
        # 列出目錄內容供除錯
        if FONT_PATH.parent.exists():
            print(f"🔍 目錄內容：{[f.name for f in FONT_PATH.parent.glob('*')]}")

# 執行初始化
_setup_matplotlib_fonts()

# =========================================================
# 2) 報表清單設定
# =========================================================
REPORT_MENU: list[tuple[str, str]] = [
    ("圖1: 累計核發人次趨勢", "figure_01"),
    ("圖2: 累計持卡人次-按領域分", "figure_02"),
    ("圖3: 累計持卡人次-十大國別", "figure_03"),
    ("圖4: 累計核發人次-年齡及性別", "figure_04"),
    ("圖5: 有效持卡人次-按領域分", "figure_05"),
    ("圖6: 有效持卡人次-十大國別", "figure_06"),
    ("圖7: 有效持卡人次-年齡及性別", "figure_07"),
    ("表1: 歷年累計許可人次 (總覽)", "table_01"),
    ("表2: 歷年累計許可人次 (初次申請)", "table_02"),
    ("表3: 歷年累計許可人次 (領域分)", "table_03"),
    ("表4: 累計許可人次 (領域x性別)", "table_04"),
    ("表5: 累計許可人次 (國別x性別)", "table_05"),
    ("表6: 有效持卡十大國別 (國別x領域)", "table_06"),
    ("表7: 累計許可人次 (年齡x性別)", "table_07"),
    ("表8: 歷年有效許可人次 (外籍x港澳)", "table_08"),
    ("表9: 有效許可人次 (領域x境內外)", "table_09"),
    ("表10: 有效許可人次 (有效領域x性別)", "table_10"),
    ("表11: 有效許可人次 (有效國別x境內外)", "table_11"),
    ("表12: 有效許可人次 (有效國別x領域)", "table_12"),
    ("表13: 有效許可人次 (有效國別x性別)", "table_13"),
    ("表14: 有效許可人次 (有效年齡x性別)", "table_14"),
]

REPORT_LABEL_BY_MODULE = {module: label for label, module in REPORT_MENU}
LOGGER = logging.getLogger(__name__)

# =========================================================
# 3) Streamlit 指令攔截器 (Recorder)
# =========================================================
class StreamlitRecorder:
    def __init__(self, requested_ym: str | None = None):
        self.requested_ym = (requested_ym or "").strip()
        self.events: list[tuple[str, Any]] = []

    def _record(self, kind: str, payload: Any) -> None:
        self.events.append((kind, payload))

    def subheader(self, text: Any, **_: Any): self._record("subheader", str(text))
    def title(self, text: Any, **_: Any): self._record("title", str(text))
    def markdown(self, text: Any, **_: Any): self._record("markdown", str(text))
    def info(self, text: Any, **_: Any): self._record("info", str(text))
    def error(self, text: Any, **_: Any): self._record("error", str(text))

    def text_input(self, label: str, value: str = "", **_: Any) -> str:
        return self.requested_ym if self.requested_ym else value

    def dataframe(self, data: Any, **_: Any):
        if hasattr(data, "data") and isinstance(data.data, pd.DataFrame): data = data.data
        if isinstance(data, pd.DataFrame): self._record("dataframe", data)

    def pyplot(self, fig: Figure | None = None, **_: Any):
        if fig is not None: self._record("pyplot", fig)

    def __getattr__(self, name: str):
        return lambda *args, **kwargs: None

# =========================================================
# 4) 組件轉換工具與圖片處理
# =========================================================
def _dataframe_to_dash_table(df: pd.DataFrame, table_id: str):
    display_df = df.copy().fillna("")
    return dash_table.DataTable(
        id=table_id,
        data=display_df.to_dict("records"),
        columns=[{"name": str(c), "id": str(c)} for c in display_df.columns],
        page_size=20,
        style_table={"overflowX": "auto"},
        style_cell={
            "textAlign": "left", "padding": "10px", "fontSize": "14px",
            "fontFamily": "Noto Sans TC, sans-serif"
        },
        style_header={"backgroundColor": "#f8f9fa", "fontWeight": "bold"}
    )

def _matplotlib_figure_to_data_uri(fig: Figure) -> str:
    """將畫好的圖表轉成 Base64 圖片，並強制套用字型"""
    # ✨ 終極修復：遍歷圖中所有文字物件並強制指定字型檔
    if FONT_PATH.exists():
        prop = fm.FontProperties(fname=str(FONT_PATH))
        # 處理標題
        if fig.texts:
            for t in fig.texts: t.set_fontproperties(prop)
        
        # 處理各個坐標軸
        for ax in fig.get_axes():
            # 標題與標籤
            ax.title.set_fontproperties(prop)
            ax.xaxis.label.set_fontproperties(prop)
            ax.yaxis.label.set_fontproperties(prop)
            # 刻度文字
            for label in ax.get_xticklabels() + ax.get_yticklabels():
                label.set_fontproperties(prop)
            # 圖例
            legend = ax.get_legend()
            if legend:
                for text in legend.get_texts():
                    text.set_fontproperties(prop)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=160, bbox_inches="tight", facecolor='white')
    buf.seek(0)
    encoded = base64.b64encode(buf.read()).decode("ascii")
    return f"data:image/png;base64,{encoded}"

def _event_to_component(event_type: str, payload: Any, idx: int):
    if event_type == "title": return html.H2(payload)
    if event_type == "subheader": return html.H4(payload, style={"color": "#2c3e50"})
    if event_type == "markdown": return dcc.Markdown(payload)
    if event_type == "info": return html.Div(payload, style={"padding": "10px", "backgroundColor": "#e7f3fe", "borderLeft": "5px solid #2196F3"})
    if event_type == "dataframe": return _dataframe_to_dash_table(payload, f"table-{idx}")
    if event_type == "pyplot":
        src = _matplotlib_figure_to_data_uri(payload)
        return html.Img(src=src, style={"maxWidth": "100%", "marginTop": "20px", "border": "1px solid #eee"})
    return None

# =========================================================
# 5) 主渲染入口
# =========================================================
def render_report_to_dash(module_name: str, data_dir: Path, requested_ym: str | None = None):
    if str(data_dir) not in sys.path: sys.path.insert(0, str(data_dir))

    recorder = StreamlitRecorder(requested_ym=requested_ym)
    
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        module = importlib.import_module(module_name)
        module = importlib.reload(module)
        module.st = recorder
        
        try:
            module.render_streamlit(data_dir)
        except Exception as e:
            return [html.Div(f"❌ 執行報表模組出錯：{e}", style={"color": "red"})]

    components = []
    for i, (kind, val) in enumerate(recorder.events):
        comp = _event_to_component(kind, val, i)
        if comp: components.append(comp)
    
    return components if components else [html.Div("此報表目前無可顯示內容。")]