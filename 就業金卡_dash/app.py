from __future__ import annotations

import logging
from pathlib import Path

from dash import Dash, Input, Output, dcc, html

try:
    from .dash_adapter import REPORT_LABEL_BY_MODULE, REPORT_MENU, render_report_to_dash
except ImportError:
    from dash_adapter import REPORT_LABEL_BY_MODULE, REPORT_MENU, render_report_to_dash

DATA_DIR = Path(__file__).resolve().parent.parent / "就業金卡"
LOGGER = logging.getLogger(__name__)

app = Dash(__name__)
server = app.server  # ✨ 關鍵：補上這行，gunicorn 才能抓到它
app.title = "就業金卡報表系統 (Plotly Dash)"

app.title = "就業金卡報表系統 (Plotly Dash)"

dropdown_options = [{"label": label, "value": module} for label, module in REPORT_MENU]
default_module = REPORT_MENU[0][1]

app.layout = html.Div(
    [
        html.H1("🏆 就業金卡數據儀表板（Plotly Dash）"),
        html.P("已將原本 21 份報表改為在 Dash 介面呈現。"),
        html.Div(
            [
                html.Div(
                    [
                        html.Label("請選擇報表"),
                        dcc.Dropdown(
                            id="report-select",
                            options=dropdown_options,
                            value=default_module,
                            clearable=False,
                        ),
                    ],
                    style={"flex": "2"},
                ),
                html.Div(
                    [
                        html.Label("截止月份（可留白）"),
                        dcc.Input(
                            id="cutoff-input",
                            type="text",
                            placeholder="例如：2025/11",
                            debounce=True,
                            style={"width": "100%"},
                        ),
                    ],
                    style={"flex": "1"},
                ),
            ],
            style={"display": "flex", "gap": "12px", "alignItems": "end"},
        ),
        html.Div(id="report-meta", style={"marginTop": "14px", "color": "#444"}),
        html.Hr(),
        dcc.Loading(html.Div(id="report-content"), type="default"),
    ],
    style={"maxWidth": "1400px", "margin": "0 auto", "padding": "18px"},
)


@app.callback(
    Output("report-content", "children"),
    Output("report-meta", "children"),
    Input("report-select", "value"),
    Input("cutoff-input", "value"),
)
def update_report(module_name: str | None, requested_ym: str | None):
    safe_module = (
        module_name
        if isinstance(module_name, str) and module_name in REPORT_LABEL_BY_MODULE
        else default_module
    )
    label = REPORT_LABEL_BY_MODULE.get(safe_module, safe_module)

    try:
        requested_ym_text = (
            requested_ym.strip()
            if isinstance(requested_ym, str)
            else (str(requested_ym).strip() if requested_ym is not None else "")
        )
        components = render_report_to_dash(
            module_name=safe_module,
            data_dir=DATA_DIR,
            requested_ym=requested_ym_text,
        )
        if requested_ym_text:
            meta = f"目前報表：{label}｜截止月份輸入：{requested_ym_text}"
        else:
            meta = f"目前報表：{label}｜截止月份：自動使用最新資料"
        return components, meta
    except Exception as exc:
        LOGGER.exception(
            "update_report callback failed (module=%r, cutoff=%r)",
            module_name,
            requested_ym,
        )
        return (
            [html.Div(f"載入報表失敗：{type(exc).__name__}: {exc}", style={"color": "#d93025"})],
            f"目前報表：{label}",
        )


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=8050)
