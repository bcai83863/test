from __future__ import annotations

from functools import lru_cache
from pathlib import Path
from typing import Iterator

import matplotlib.pyplot as plt
from matplotlib import font_manager as fm

_CJK_FONT_FAMILY_CANDIDATES = [
    "Noto Sans CJK TC",
    "Noto Sans CJK SC",
    "Noto Sans CJK JP",
    "Source Han Sans TW",
    "Source Han Sans CN",
    "WenQuanYi Zen Hei",
    "Microsoft JhengHei",
    "PingFang TC",
    "Heiti TC",
    "Arial Unicode MS",
]

_FONT_FILE_KEYWORDS = ("noto", "cjk", "sourcehan", "wqy", "wenquanyi")
_FONT_FILE_PATTERNS = ("*.ttf", "*.ttc", "*.otf")
_FONT_DIRS = (
    Path("/usr/share/fonts"),
    Path("/usr/local/share/fonts"),
    Path.home() / ".fonts",
)


def _iter_cjk_font_files() -> Iterator[Path]:
    for base in _FONT_DIRS:
        if not base.exists():
            continue
        for pattern in _FONT_FILE_PATTERNS:
            for font_path in base.rglob(pattern):
                name = font_path.name.lower()
                if any(keyword in name for keyword in _FONT_FILE_KEYWORDS):
                    yield font_path


@lru_cache(maxsize=1)
def _register_extra_cjk_fonts() -> None:
    for font_path in _iter_cjk_font_files():
        try:
            fm.fontManager.addfont(str(font_path))
        except (OSError, RuntimeError, ValueError):
            continue


@lru_cache(maxsize=1)
def pick_available_cjk_font() -> str:
    _register_extra_cjk_fonts()
    available = {font.name for font in fm.fontManager.ttflist}
    for family in _CJK_FONT_FAMILY_CANDIDATES:
        if family in available:
            return family
    return "DejaVu Sans"


def apply_cjk_font_settings() -> str:
    selected = pick_available_cjk_font()
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = [
        selected,
        *_CJK_FONT_FAMILY_CANDIDATES,
        "DejaVu Sans",
        "sans-serif",
    ]
    plt.rcParams["axes.unicode_minus"] = False
    return selected


def apply_streamlit_cjk_css() -> None:
    import streamlit as st

    st.markdown(
        """
        <style>
        html, body, [class*="css"], [data-testid="stAppViewContainer"],
        [data-testid="stMarkdownContainer"], [data-testid="stDataFrame"], [data-testid="stTable"] {
            font-family: "Noto Sans CJK TC", "Noto Sans CJK SC", "Noto Sans CJK JP",
                         "WenQuanYi Zen Hei", "Microsoft JhengHei", sans-serif;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
