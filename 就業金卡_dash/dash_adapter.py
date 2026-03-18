from __future__ import annotations

import base64
import importlib
import io
import logging
import sys
import warnings
import platform
from pathlib import Path
from typing import Any

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from dash import dash_table, dcc, html
from matplotlib.figure import Figure

# =========================================================
# 1) 字型修復邏輯 (防止 Linux 環境出現口口字)
# =========================================================
def _setup_matplotlib_fonts():
    """在伺服器環境手動載入字型檔案"""
    # 設定繪圖後台為 Agg (非交互式，適合伺服器)
    plt.switch_backend('Agg')
    
    # 定義字型路徑
    font_path = Path(__file__).parent / "assets" / "fonts" / "NotoSansTC-VariableFont_wght.ttf"
    
    if font_path.exists():
        # 強制插入字型到 Matplotlib
        fe = fm.FontEntry(fname=str(font_path), name='Noto Sans TC')
        fm.fontManager.ttflist.insert(0, fe)
        plt.rcParams['font.family'] = fe.name
        plt.rcParams['axes.unicode_minus'] = False
        print(f"✅ 已成功載入字型：{font_path}")
    else:
        # 備援方案
        plt.rcParams['font.sans-serif'] = ['Noto Sans CJK TC', 'WenQuanYi Zen Hei', 'sans-serif']
        plt.rcParams['axes.unicode_minus'] = False
        print(f"⚠️ 找不到字型檔 {font_path}，使用系統備援字型。")

# 執行字型初始化
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
    """模擬 st.xxx 指令，將輸出存入 events 串列"""
    def __init__(self, requested_ym: str | None = None):
        self.requested_ym = (requested_ym or "").strip()
        self.events: list[tuple[str, Any]] = []

    def _record(self, kind: str, payload: Any) -> None:
        self.events.append((kind, payload))

    def subheader(self, text: Any, **_: Any) -> None: self._record("subheader", str(text))
    def title(self, text: Any, **_: Any) -> None: self._record("title", str(text))
    def markdown(self, text: Any, **_: Any) -> None: self._record("markdown", str(text))
    def info(self, text: Any, **_: Any) -> None: self._record("info", str(text))
    def error(self, text: Any, **_: Any) -> None: self._record("error", str(text))

    def text_input(self, label: str, value: str = "", **_: Any) -> str:
        # 攔截輸入框，優先回傳 Dash 介面傳進來的月份
        return self.requested_ym if self.requested_ym else value

    def dataframe(self, data: Any, **_: Any) -> None:
        if hasattr(data, "data") and isinstance(data.data, pd.DataFrame): data = data.data
        if isinstance(data, pd.DataFrame): self._record("dataframe", data)

    def pyplot(self, fig: Figure | None = None, **_: Any) -> None:
        if fig is not None: self._record("pyplot", fig)

    def __getattr__(self, name: str):
        # 忽略所有其他不支援的 Streamlit 指令 (如 st.cache_data)
        return lambda *args, **kwargs: None

# =========================================================
# 4) 組件轉換工具
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
    """將畫好的 Matplotlib 圖表轉成 Base64 圖片網址"""
    buf = io.BytesIO()
    # 存檔時強制指定 dpi 與 背景色
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor='white')
    buf.seek(0)
    encoded = base64.b64encode(buf.read()).decode("ascii")
    return f"data:image/png;base64,{encoded}"

def _event_to_component(event_type: str, payload: Any, idx: int):
    if event_type == "title": return html.H2(payload)
    if event_type == "subheader": return html.H4(payload, style={"color": "#2c3e50"})
    if event_type == "markdown": return dcc.Markdown(payload)
    if event_type == "info": return html.Div(payload, className="alert alert-info", style={"padding": "10px", "backgroundColor": "#e7f3fe", "borderLeft": "5px solid #2196F3"})
    if event_type == "dataframe": return _dataframe_to_dash_table(payload, f"table-{idx}")
    if event_type == "pyplot":
        src = _matplotlib_figure_to_data_uri(payload)
        return html.Img(src=src, style={"maxWidth": "100%", "marginTop": "20px", "border": "1px solid #eee"})
    return None

# =========================================================
# 5) 主入口：渲染報表至 Dash
# =========================================================
def render_report_to_dash(module_name: str, data_dir: Path, requested_ym: str | None = None):
    # 確保模組搜尋路徑正確
    if str(data_dir) not in sys.path: sys.path.insert(0, str(data_dir))

    recorder = StreamlitRecorder(requested_ym=requested_ym)
    
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        # 動態載入報表檔案 (例如 figure_01.py)
        module = importlib.import_module(module_name)
        module = importlib.reload(module)
        
        # 💉 關鍵注入：將報表內的 st 替換成我們的 recorder
        module.st = recorder
        
        # 執行報表主程式
        try:
            module.render_streamlit(data_dir)
        except Exception as e:
            return [html.Div(f"❌ 執行模組出錯：{e}", style={"color": "red"})]

    # 將紀錄的事件轉為 Dash 組件
    components = []
    for i, (kind, val) in enumerate(recorder.events):
        comp = _event_to_component(kind, val, i)
        if comp: components.append(comp)
    
    return components if components else [html.Div("⚠️ 此報表目前無可顯示內容")]