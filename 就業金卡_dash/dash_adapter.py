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
from dash import dash_table, dcc, html
from matplotlib.figure import Figure

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


def _quiet_streamlit_logs() -> None:
    for logger_name in (
        "streamlit",
        "streamlit.runtime",
        "streamlit.runtime.caching",
        "streamlit.runtime.caching.cache_data_api",
        "streamlit.runtime.scriptrunner_utils.script_run_context",
    ):
        logging.getLogger(logger_name).setLevel(logging.ERROR)


class StreamlitRecorder:
    """記錄 Streamlit 輸出，以便轉成 Dash 元件。"""

    def __init__(self, requested_ym: str | None = None):
        self.requested_ym = (requested_ym or "").strip()
        self.events: list[tuple[str, Any]] = []

    def _record(self, kind: str, payload: Any) -> None:
        self.events.append((kind, payload))

    def subheader(self, text: Any, **_: Any) -> None:
        self._record("subheader", str(text))

    def title(self, text: Any, **_: Any) -> None:
        self._record("title", str(text))

    def markdown(self, text: Any, **_: Any) -> None:
        self._record("markdown", str(text))

    def info(self, text: Any, **_: Any) -> None:
        self._record("info", str(text))

    def error(self, text: Any, **_: Any) -> None:
        self._record("error", str(text))

    def text_input(self, _label: str, value: str = "", **_: Any) -> str:
        if self.requested_ym:
            return self.requested_ym
        return value or ""

    def dataframe(self, data: Any, **_: Any) -> None:
        if hasattr(data, "data") and isinstance(getattr(data, "data"), pd.DataFrame):
            data = data.data
        if isinstance(data, pd.DataFrame):
            self._record("dataframe", data)
        else:
            self._record("markdown", f"⚠️ 非 DataFrame 輸出：{type(data).__name__}")

    def pyplot(self, fig: Figure | None = None, **_: Any) -> None:
        if fig is not None:
            self._record("pyplot", fig)

    def download_button(self, *_: Any, **__: Any) -> None:
        return None

    def __getattr__(self, _name: str):  # pragma: no cover
        def _noop(*_args: Any, **_kwargs: Any):
            return None

        return _noop


def _dataframe_to_dash_table(df: pd.DataFrame, table_id: str):
    display_df = df.copy().fillna("")
    return dash_table.DataTable(
        id=table_id,
        data=display_df.to_dict("records"),
        columns=[{"name": str(c), "id": str(c)} for c in display_df.columns],
        page_size=min(25, max(8, len(display_df))),
        style_table={"overflowX": "auto", "marginBottom": "16px"},
        style_cell={
            "textAlign": "left",
            "padding": "6px",
            "fontSize": "14px",
            "fontFamily": "Noto Sans CJK TC, WenQuanYi Zen Hei, sans-serif",
            "whiteSpace": "normal",
            "height": "auto",
            "minWidth": "80px",
        },
        style_header={
            "backgroundColor": "#f5f5f5",
            "fontWeight": "bold",
        },
    )


def _matplotlib_figure_to_data_uri(fig: Figure) -> str:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=170, bbox_inches="tight")
    buf.seek(0)
    encoded = base64.b64encode(buf.read()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def _event_to_component(event_type: str, payload: Any, idx: int):
    if event_type == "title":
        return html.H2(payload, style={"marginTop": "8px"})
    if event_type == "subheader":
        return html.H3(payload, style={"marginTop": "8px"})
    if event_type == "markdown":
        return dcc.Markdown(payload)
    if event_type == "info":
        return html.Div(
            payload,
            style={
                "backgroundColor": "#e8f4fd",
                "borderLeft": "4px solid #1f77b4",
                "padding": "8px 12px",
                "margin": "8px 0",
            },
        )
    if event_type == "error":
        return html.Div(
            payload,
            style={
                "backgroundColor": "#fdecec",
                "borderLeft": "4px solid #d93025",
                "padding": "8px 12px",
                "margin": "8px 0",
                "color": "#8a1f11",
            },
        )
    if event_type == "dataframe":
        return _dataframe_to_dash_table(payload, table_id=f"table-{idx}")
    if event_type == "pyplot":
        src = _matplotlib_figure_to_data_uri(payload)
        return html.Img(
            src=src,
            alt="report-figure",
            style={
                "maxWidth": "100%",
                "height": "auto",
                "display": "block",
                "marginBottom": "16px",
                "border": "1px solid #ddd",
                "borderRadius": "6px",
            },
        )
    return None


def render_report_to_dash(
    module_name: str,
    data_dir: Path,
    requested_ym: str | None = None,
):
    if str(data_dir) not in sys.path:
        sys.path.insert(0, str(data_dir))

    recorder = StreamlitRecorder(requested_ym=requested_ym)
    _quiet_streamlit_logs()

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        module = importlib.import_module(module_name)
        module = importlib.reload(module)

        module.st = recorder
        if hasattr(module, "apply_streamlit_cjk_css"):
            module.apply_streamlit_cjk_css = lambda: None

        module.render_streamlit(data_dir)

    components: list[Any] = []
    for idx, (event_type, payload) in enumerate(recorder.events):
        try:
            comp = _event_to_component(event_type, payload, idx)
            if comp is not None:
                components.append(comp)
        except Exception as exc:  # 轉換錯誤需明確呈現
            LOGGER.exception("轉換事件失敗: %s", event_type)
            components.append(
                html.Div(
                    f"轉換事件 `{event_type}` 失敗：{exc}",
                    style={"color": "#d93025"},
                )
            )

    if not components:
        components.append(
            html.Div("此報表沒有可顯示內容。", style={"color": "#666"})
        )
    return components
