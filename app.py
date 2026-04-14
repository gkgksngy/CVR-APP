import io
import os
import re
import sys
import time
import math
import subprocess
import importlib.util
import tempfile
import shutil
import traceback
from datetime import datetime

import pandas as pd
import requests
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image as RLImage,
    PageBreak,
    LongTable,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

import matplotlib.pyplot as plt
import matplotlib.font_manager as fm


st.set_page_config(page_title="CVR 운영형 계산기", layout="wide")
st.markdown(
    """
    <meta name="google" content="notranslate">
    """,
    unsafe_allow_html=True,
)


# =========================================================
# 세션 초기화
# =========================================================
DEFAULT_STATE = {
    "pp_loaded": False,
    "pp_last_message": "",
    "pp_log_text": "",
    "pp_contract_kind": "",
    "pp_basic_charge_unit": 0.0,
    "pp_power_bill_kw": 0.0,
    "pp_max_demand_kw": 0.0,
    "pp_annual_usage_kwh": 0.0,
    "pp_off_peak_rate": 0.0,
    "pp_mid_peak_rate": 0.0,
    "pp_peak_rate": 0.0,
    "pp_primary_voltage_kv": 154.0,
    "pp_contract_power_kw": 0.0,
    "pp_supply_voltage_text": "",
    "pp_meter_read_day": 0,
    "pp_yearly_bill_won": 0.0,
    "pp_voltage_class": "",
    "pp_auto_avg_base_kw": 0.0,
    "pp_auto_off_peak_kw": 0.0,
    "pp_auto_mid_peak_kw": 0.0,
    "pp_auto_peak_kw": 0.0,
    "pp_auto_off_ratio": 0.0,
    "pp_auto_mid_ratio": 0.0,
    "pp_auto_peak_ratio": 0.0,
}
for k, v in DEFAULT_STATE.items():
    if k not in st.session_state:
        st.session_state[k] = v


# =========================================================
# 기본 데이터
# =========================================================
LOAD_TYPES = {
    "사무/상업 혼합": {"cvrf": 0.85, "z": 0.35, "i": 0.25, "p": 0.40},
    "공장 혼합": {"cvrf": 0.72, "z": 0.25, "i": 0.25, "p": 0.50},
    "모터 부하 많음": {"cvrf": 0.60, "z": 0.20, "i": 0.35, "p": 0.45},
    "조명/히터 부하 많음": {"cvrf": 0.95, "z": 0.55, "i": 0.20, "p": 0.25},
    "인버터/SMPS 부하 많음": {"cvrf": 0.40, "z": 0.10, "i": 0.15, "p": 0.75},
}

SEASON_SCHEDULE = {
    "봄·가을": {
        "경부하": list(range(23, 24)) + list(range(0, 9)),
        "중간부하": [9, 10, 11, 13, 14, 15, 16, 20, 21, 22],
        "최대부하": [12, 17, 18, 19],
    },
    "여름": {
        "경부하": list(range(23, 24)) + list(range(0, 9)),
        "중간부하": [9, 12, 17, 18, 19, 20, 21, 22],
        "최대부하": [10, 11, 13, 14, 15, 16],
    },
    "겨울": {
        "경부하": list(range(23, 24)) + list(range(0, 9)),
        "중간부하": [9, 12, 13, 14, 15, 19, 20, 21, 22],
        "최대부하": [10, 11, 16, 17, 18],
    },
}

BASE_TARIFFS = {
    "산업용": {
        "을": {
            "고압A": {
                "선택 II": {
                    "여름": {"경부하": 124.0, "중간부하": 178.0, "최대부하": 261.0},
                    "봄·가을": {"경부하": 124.0, "중간부하": 147.0, "최대부하": 179.0},
                    "겨울": {"경부하": 131.0, "중간부하": 178.0, "최대부하": 236.0},
                }
            },
            "고압B": {
                "선택 II": {
                    "여름": {"경부하": 117.0, "중간부하": 170.0, "최대부하": 250.0},
                    "봄·가을": {"경부하": 117.0, "중간부하": 141.0, "최대부하": 173.0},
                    "겨울": {"경부하": 124.0, "중간부하": 170.0, "최대부하": 227.0},
                }
            },
            "고압C": {
                "선택 II": {
                    "여름": {"경부하": 111.0, "중간부하": 162.0, "최대부하": 239.0},
                    "봄·가을": {"경부하": 111.0, "중간부하": 135.0, "최대부하": 167.0},
                    "겨울": {"경부하": 117.0, "중간부하": 162.0, "최대부하": 218.0},
                }
            },
        }
    }
}


# =========================================================
# UI 커스텀 함수
# =========================================================
def colored_input(label, widget_func, color_type, **kwargs):
    colors_map = {
        "auto": "#d4edda",
        "verify": "#fff3cd",
        "manual": "#d1ecf1",
    }
    bg_color = colors_map.get(color_type, "#ffffff")
    st.markdown(
        f"""
        <div style="
            background-color: {bg_color};
            padding: 4px 10px;
            border-radius: 5px;
            margin-bottom: 2px;
            font-size: 14px;
            font-weight: bold;
            color: #333;">
            {label}
        </div>
        """,
        unsafe_allow_html=True,
    )
    kwargs["label_visibility"] = "collapsed"
    return widget_func(label, **kwargs)


# =========================================================
# 유틸 함수
# =========================================================
def has_module(module_name: str) -> bool:
    return importlib.util.find_spec(module_name) is not None


def run_command_capture(cmd):
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=10)
        return result.returncode, (result.stdout or "").strip(), (result.stderr or "").strip()
    except Exception as e:
        return -1, "", str(e)


def first_existing_path(candidates):
    for path in candidates:
        if path and os.path.exists(path):
            return path
    return None


def get_chrome_binary_candidates():
    return [
        os.environ.get("CHROME_BINARY"),
        os.environ.get("GOOGLE_CHROME_BIN"),
        shutil.which("chromium"),
        shutil.which("chromium-browser"),
        shutil.which("google-chrome"),
        shutil.which("google-chrome-stable"),
        "/usr/bin/chromium",
        "/usr/bin/chromium-browser",
        "/usr/bin/google-chrome",
        "/usr/bin/google-chrome-stable",
    ]


def get_chromedriver_candidates():
    return [
        os.environ.get("CHROMEDRIVER_PATH"),
        shutil.which("chromedriver"),
        "/usr/bin/chromedriver",
        "/usr/local/bin/chromedriver",
    ]


def get_runtime_environment_summary():
    chrome_path = first_existing_path(get_chrome_binary_candidates())
    chromedriver_path = first_existing_path(get_chromedriver_candidates())

    chrome_ver = "-"
    chromedriver_ver = "-"

    if chrome_path:
        rc, out, err = run_command_capture([chrome_path, "--version"])
        chrome_ver = out if rc == 0 and out else (err or "확인 실패")

    if chromedriver_path:
        rc, out, err = run_command_capture([chromedriver_path, "--version"])
        chromedriver_ver = out if rc == 0 and out else (err or "확인 실패")

    return {
        "python": sys.executable,
        "platform": sys.platform,
        "chrome_path": chrome_path or "-",
        "chrome_version": chrome_ver,
        "chromedriver_path": chromedriver_path or "-",
        "chromedriver_version": chromedriver_ver,
    }


def build_chrome_driver(logs):
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options

    chrome_binary = first_existing_path(get_chrome_binary_candidates())
    chromedriver_path = first_existing_path(get_chromedriver_candidates())

    if not chrome_binary:
        raise RuntimeError(
            "크롬/크로미움 실행 파일을 찾지 못했습니다. "
            "배포 서버에 chromium 또는 google-chrome 설치가 필요합니다."
        )

    options = Options()
    options.binary_location = chrome_binary

    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--remote-debugging-port=9222")
    options.add_argument("--disable-software-rasterizer")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--lang=ko-KR")
    options.add_argument("--user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

    tmp_profile = tempfile.mkdtemp(prefix="chrome-profile-")
    tmp_data = tempfile.mkdtemp(prefix="chrome-data-")
    options.add_argument(f"--user-data-dir={tmp_profile}")
    options.add_argument(f"--data-path={tmp_data}")

    add_log(logs, f"chrome binary: {chrome_binary}")
    add_log(logs, f"chromedriver path: {chromedriver_path or 'selenium manager 사용'}")

    try:
        if chromedriver_path:
            service = Service(executable_path=chromedriver_path)
            driver = webdriver.Chrome(service=service, options=options)
        else:
            driver = webdriver.Chrome(options=options)

        driver.set_page_load_timeout(40)
        driver.implicitly_wait(2)
        return driver
    except Exception:
        shutil.rmtree(tmp_profile, ignore_errors=True)
        shutil.rmtree(tmp_data, ignore_errors=True)
        raise


def install_package_for_current_python(packages):
    cmd = [sys.executable, "-m", "pip", "install"] + packages
    result = subprocess.run(cmd, capture_output=True, text=True)
    return result.returncode, result.stdout, result.stderr


def parse_number(text):
    if text is None:
        return 0.0
    cleaned = re.sub(r"[^0-9.\-]", "", str(text))
    if cleaned in ["", ".", "-", "-."]:
        return 0.0
    try:
        return float(cleaned)
    except Exception:
        return 0.0


def parse_voltage_from_text(text):
    if not text:
        return 0.0
    raw = str(text)
    low_voltage_match = re.search(r"(220/380|380/220|380V|220V)", raw, re.IGNORECASE)
    if low_voltage_match:
        return 0.38
    m = re.search(r"(\d+(?:\.\d+)?)\s*kV", raw, re.IGNORECASE)
    if m:
        return float(m.group(1))
    return 0.0


def classify_voltage_from_contract_kind(contract_kind, supply_text=""):
    raw = f"{contract_kind or ''} {supply_text or ''}".replace(" ", "")
    if "저압" in raw or "220/380" in raw or "380V" in raw or "220V" in raw:
        return "저압"
    if "고압C" in raw or "345" in raw:
        return "고압C"
    if "고압B" in raw or "154" in raw:
        return "고압B"
    if "고압A" in raw or "22.9" in raw:
        return "고압A"
    return ""


def primary_voltage_from_class(voltage_class, supply_text=""):
    parsed = parse_voltage_from_text(supply_text)
    if parsed > 0:
        return parsed
    if voltage_class == "저압":
        return 0.38
    if voltage_class == "고압A":
        return 22.9
    if voltage_class == "고압B":
        return 154.0
    if voltage_class == "고압C":
        return 345.0
    return 22.9


def get_voltage_class(primary_kv):
    if primary_kv < 1.0:
        return "저압"
    if 3.3 <= primary_kv <= 66.0:
        return "고압A"
    if 100.0 <= primary_kv < 300.0:
        return "고압B"
    if primary_kv >= 300.0:
        return "고압C"
    return "고압A"


def describe_voltage_class(voltage_class, primary_kv):
    if voltage_class == "저압":
        return "저압"
    if voltage_class == "고압A":
        return f"고압A ({primary_kv:.1f}kV)"
    if voltage_class == "고압B":
        return f"고압B ({primary_kv:.1f}kV)"
    if voltage_class == "고압C":
        return f"고압C ({primary_kv:.1f}kV)"
    return "저압"


def tariff_description(choice_name="선택 II"):
    if choice_name == "선택 II":
        return "일반적인 평균형 요금제"
    return "-"


def calc_tap_voltage_change(step_percent, tap_change_steps):
    return step_percent * tap_change_steps


def calc_cvrf(base_kw, voltage_drop_pct, cvrf, hours):
    saving_rate_pct = voltage_drop_pct * cvrf
    kw_saved = base_kw * (saving_rate_pct / 100.0)
    kwh_saved = kw_saved * hours
    return saving_rate_pct, kw_saved, kwh_saved


def calc_zip(base_kw, voltage_drop_pct, z, i, p, hours):
    if base_kw <= 0:
        return 0.0, 0.0, 0.0
    v_pu = 1.0 - voltage_drop_pct / 100.0
    new_kw = base_kw * (z * (v_pu ** 2) + i * v_pu + p)
    kw_saved = base_kw - new_kw
    saving_rate_pct = (kw_saved / base_kw) * 100.0 if base_kw else 0.0
    kwh_saved = kw_saved * hours
    return saving_rate_pct, kw_saved, kwh_saved


def calc_average_result(load_kw, voltage_drop_pct, cvrf, z, i, p, hours):
    cvrf_rate, cvrf_kw, cvrf_kwh = calc_cvrf(load_kw, voltage_drop_pct, cvrf, hours)
    zip_rate, zip_kw, zip_kwh = calc_zip(load_kw, voltage_drop_pct, z, i, p, hours)
    return (
        (cvrf_rate + zip_rate) / 2.0,
        (cvrf_kw + zip_kw) / 2.0,
        (cvrf_kwh + zip_kwh) / 2.0,
    )


def annual_kwh_to_avg_kw(annual_kwh):
    if annual_kwh <= 0:
        return 0.0
    return annual_kwh / 8760.0


def avg_kw_to_timeband_loads(avg_kw, off_ratio, mid_ratio, peak_ratio):
    return {
        "경부하": avg_kw * off_ratio,
        "중간부하": avg_kw * mid_ratio,
        "최대부하": avg_kw * peak_ratio,
    }


def hour_to_label(hour, season):
    schedule = SEASON_SCHEDULE[season]
    for label, hours in schedule.items():
        if hour in hours:
            return label
    return "중간부하"


def get_operating_hours_by_label(season, operating_start, operating_end):
    if operating_start == operating_end:
        active_hours = list(range(24))
    elif operating_start < operating_end:
        active_hours = list(range(operating_start, operating_end))
    else:
        active_hours = list(range(operating_start, 24)) + list(range(0, operating_end))

    counts = {"경부하": 0, "중간부하": 0, "최대부하": 0}
    for hour in active_hours:
        counts[hour_to_label(hour, season)] += 1
    return active_hours, counts


def get_active_days_per_year(operation_mode, holiday_count, include_holidays_as_shutdown):
    weekday_map = {
        "월~금 가동": 5,
        "월~토 가동": 6,
        "주7일 가동": 7,
    }
    weekly_days = weekday_map[operation_mode]
    base_days = int(round(365.0 * weekly_days / 7.0))
    if include_holidays_as_shutdown:
        base_days -= holiday_count
    return max(base_days, 0)


def safe_rate_table(voltage_class):
    try:
        return BASE_TARIFFS["산업용"]["을"][voltage_class]["선택 II"]
    except Exception:
        return {
            "여름": {"경부하": 120.0, "중간부하": 170.0, "최대부하": 250.0},
            "봄·가을": {"경부하": 120.0, "중간부하": 145.0, "최대부하": 175.0},
            "겨울": {"경부하": 127.0, "중간부하": 170.0, "최대부하": 230.0},
        }


def add_log(logs, text):
    now = datetime.now().strftime("%H:%M:%S")
    logs.append("[{0}] {1}".format(now, text))


def choose_latest_full_year(rows):
    if not rows:
        return None
    current_year = datetime.now().year
    filtered = [r for r in rows if int(r["year"]) < current_year]
    if filtered:
        return sorted(filtered, key=lambda x: int(x["year"]), reverse=True)[0]
    return sorted(rows, key=lambda x: int(x["year"]), reverse=True)[0]


CURRENT_PDF_FONT_NAME = "Helvetica"


def get_korean_font_info():
    candidates = [
        {
            "pdf_name": "Malgun",
            "mpl_fallback_name": "Malgun Gothic",
            "path": "C:/Windows/Fonts/malgun.ttf",
        },
        {
            "pdf_name": "NanumGothic",
            "mpl_fallback_name": "NanumGothic",
            "path": "C:/Windows/Fonts/NanumGothic.ttf",
        },
        {
            "pdf_name": "NanumBarunGothic",
            "mpl_fallback_name": "NanumBarunGothic",
            "path": "C:/Windows/Fonts/NanumBarunGothic.ttf",
        },
        {
            "pdf_name": "AppleGothic",
            "mpl_fallback_name": "AppleGothic",
            "path": "/System/Library/Fonts/Supplemental/AppleGothic.ttf",
        },
        {
            "pdf_name": "NanumGothicLinux",
            "mpl_fallback_name": "NanumGothic",
            "path": "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        },
        {
            "pdf_name": "NotoSansKR",
            "mpl_fallback_name": "Noto Sans CJK KR",
            "path": "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        },
        {
            "pdf_name": "NotoSansKR",
            "mpl_fallback_name": "Noto Sans CJK KR",
            "path": "/usr/share/fonts/opentype/noto/NotoSansCJKkr-Regular.otf",
        },
    ]

    for item in candidates:
        font_path = item["path"]
        if os.path.exists(font_path):
            pdf_name = item["pdf_name"]
            registered_ok = False
            try:
                registered = pdfmetrics.getRegisteredFontNames()
                if pdf_name not in registered:
                    pdfmetrics.registerFont(TTFont(pdf_name, font_path))
                registered_ok = True
            except Exception:
                registered_ok = False

            try:
                mpl_name = fm.FontProperties(fname=font_path).get_name()
            except Exception:
                mpl_name = item["mpl_fallback_name"]

            if registered_ok:
                return {
                    "pdf_name": pdf_name,
                    "font_path": font_path,
                    "mpl_name": mpl_name,
                }

            return {
                "pdf_name": "HYGothic-Medium",
                "font_path": font_path,
                "mpl_name": mpl_name,
            }

    preferred_font_files = []
    try:
        for font_path in fm.findSystemFonts(fontpaths=None, fontext="ttf") + fm.findSystemFonts(fontpaths=None, fontext="ttc") + fm.findSystemFonts(fontpaths=None, fontext="otf"):
            name = os.path.basename(font_path).lower()
            if any(key in name for key in ["notosanscjk", "notosanskr", "nanumgothic", "nanumbarungothic", "malgun", "applegothic"]):
                preferred_font_files.append(font_path)
    except Exception:
        preferred_font_files = []

    for font_path in preferred_font_files:
        try:
            pdf_name = os.path.splitext(os.path.basename(font_path))[0][:30] or "KoreanFont"
            registered = pdfmetrics.getRegisteredFontNames()
            if pdf_name not in registered:
                pdfmetrics.registerFont(TTFont(pdf_name, font_path))
            mpl_name = fm.FontProperties(fname=font_path).get_name()
            return {
                "pdf_name": pdf_name,
                "font_path": font_path,
                "mpl_name": mpl_name,
            }
        except Exception:
            continue

    try:
        registered = pdfmetrics.getRegisteredFontNames()
        if "HYGothic-Medium" not in registered:
            pdfmetrics.registerFont(UnicodeCIDFont("HYGothic-Medium"))
        return {
            "pdf_name": "HYGothic-Medium",
            "font_path": None,
            "mpl_name": "DejaVu Sans",
        }
    except Exception:
        return {
            "pdf_name": "Helvetica",
            "font_path": None,
            "mpl_name": "DejaVu Sans",
        }


def _wrap_table_data(data, font_name="Helvetica", font_size=9):
    styles = getSampleStyleSheet()
    cell_style = ParagraphStyle(
        "TableCellWrap",
        parent=styles["Normal"],
        fontName=font_name,
        fontSize=font_size,
        leading=max(font_size + 1.5, font_size * 1.15),
        alignment=1,
        wordWrap="CJK",
    )
    wrapped = []
    for row in data:
        wrapped_row = []
        for cell in row:
            if isinstance(cell, Paragraph):
                wrapped_row.append(cell)
            elif isinstance(cell, str):
                safe = cell.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")
                wrapped_row.append(Paragraph(safe, cell_style))
            else:
                wrapped_row.append(cell)
        wrapped.append(wrapped_row)
    return wrapped


def make_two_col_table(data, col_widths=None, font_name="Helvetica", font_size=9):
    wrapped = _wrap_table_data(data, font_name=font_name, font_size=font_size)
    tbl = Table(wrapped, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("FONTSIZE", (0, 0), (-1, -1), font_size),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9eaf7")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    return tbl


def make_long_table(data, col_widths=None, font_name="Helvetica", font_size=8):
    wrapped = _wrap_table_data(data, font_name=font_name, font_size=font_size)
    tbl = LongTable(wrapped, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("FONTSIZE", (0, 0), (-1, -1), font_size),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9eaf7")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return tbl


def format_hour_range(hours):
    if not hours:
        return "-"
    hours = sorted(hours)
    ranges = []
    start = hours[0]
    prev = hours[0]
    for h in hours[1:]:
        if h == prev + 1:
            prev = h
        else:
            ranges.append((start, prev + 1))
            start = h
            prev = h
    ranges.append((start, prev + 1))
    return ", ".join([f"{s:02d}:00 ~ {e:02d}:00" for s, e in ranges])


def _apply_font_to_axis(ax, font_prop):
    ax.title.set_fontproperties(font_prop)
    ax.xaxis.label.set_fontproperties(font_prop)
    ax.yaxis.label.set_fontproperties(font_prop)

    for label in ax.get_xticklabels():
        label.set_fontproperties(font_prop)
    for label in ax.get_yticklabels():
        label.set_fontproperties(font_prop)

    leg = ax.get_legend()
    if leg:
        for text in leg.get_texts():
            text.set_fontproperties(font_prop)


def create_matplotlib_line_chart(hourly_df, season, font_info):
    with plt.rc_context({"axes.unicode_minus": False}):
        if font_info.get("mpl_name"):
            plt.rcParams["font.family"] = font_info["mpl_name"]
        fig, ax = plt.subplots(figsize=(10, 4.8))
        ax.plot(hourly_df["시간번호"], hourly_df["전력사용량(kW)"], marker="o", linewidth=1.8)
        ax.set_xticks(list(range(24)))
        ax.set_xticklabels([f"{i:02d}" for i in range(24)], fontsize=8)
        ax.grid(True, alpha=0.3)

        font_prop = None
        if font_info.get("font_path") and os.path.exists(font_info["font_path"]):
            font_prop = fm.FontProperties(fname=font_info["font_path"])
        elif font_info.get("mpl_name"):
            font_prop = fm.FontProperties(family=font_info["mpl_name"])

        # PDF에서는 matplotlib 한글이 환경에 따라 깨질 수 있어서
        # 그래프 내부 제목/축 라벨은 제거하고, 한국어 제목은 PDF 본문 Paragraph로 표시한다.
        ax.set_title("")
        ax.set_xlabel("")
        ax.set_ylabel("")
        if font_prop is not None:
            _apply_font_to_axis(ax, font_prop)

        buf = io.BytesIO()
        fig.tight_layout()
        fig.savefig(buf, format="png", dpi=180, bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        return buf


def create_matplotlib_bar_chart(tap_compare_df, site_name, font_info):
    with plt.rc_context({"axes.unicode_minus": False}):
        if font_info.get("mpl_name"):
            plt.rcParams["font.family"] = font_info["mpl_name"]
        plot_df = tap_compare_df.copy()
        plot_df["탭"] = pd.to_numeric(plot_df["탭"], errors="coerce")
        plot_df = plot_df.sort_values(by="탭", ascending=True).reset_index(drop=True)
        fig, ax = plt.subplots(figsize=(10, 4.8))
        x_labels = plot_df["탭"].astype(int).astype(str)
        y_vals = plot_df["평균 절감전력(kW)"]
        bars = ax.bar(x_labels, y_vals)
        ax.grid(True, axis="y", alpha=0.3)
        ax.invert_xaxis()

        font_prop = None
        if font_info.get("font_path") and os.path.exists(font_info["font_path"]):
            font_prop = fm.FontProperties(fname=font_info["font_path"])
        elif font_info.get("mpl_name"):
            font_prop = fm.FontProperties(family=font_info["mpl_name"])

        # PDF에서는 matplotlib 한글이 환경에 따라 깨질 수 있어서
        # 그래프 내부 제목/축 라벨은 제거하고, 한국어 제목은 PDF 본문 Paragraph로 표시한다.
        ax.set_title("")
        ax.set_xlabel("")
        ax.set_ylabel("")
        if font_prop is not None:
            _apply_font_to_axis(ax, font_prop)

        for rect, val in zip(bars, y_vals):
            label = "0" if float(val) == 0 else f"{float(val):.3f}"
            ax.text(rect.get_x() + rect.get_width() / 2, rect.get_height(), label, ha="center", va="bottom", fontsize=8)

        buf = io.BytesIO()
        fig.tight_layout()
        fig.savefig(buf, format="png", dpi=180, bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        return buf


def split_even_rows(data):
    if len(data) <= 1:
        return data, [["시간", "시간번호", "구분", "운영여부", "전력사용량(kW)", "절감률(%)", "절감전력(kW)", "절감전력량(kWh)", "요금단가(원/kWh)", "절감요금(원)"]]
    header = data[0]
    body = data[1:]
    mid = math.ceil(len(body) / 2)
    return [header] + body[:mid], [header] + body[mid:]


def add_page_number(canvas, doc):
    canvas.saveState()
    try:
        canvas.setFont(CURRENT_PDF_FONT_NAME, 8)
    except Exception:
        canvas.setFont("Helvetica", 8)
    page_num = canvas.getPageNumber()
    canvas.drawRightString(doc.pagesize[0] - 15 * mm, 10 * mm, f"{page_num}")
    canvas.restoreState()



# =========================================================
# API 해석 보조 함수
# =========================================================
def create_requests_session_from_driver(driver):
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": "https://pp.kepco.co.kr/",
    })
    for cookie in driver.get_cookies():
        session.cookies.set(
            cookie.get("name"),
            cookie.get("value"),
            domain=cookie.get("domain"),
            path=cookie.get("path", "/"),
        )
    return session


def fetch_json(session, url, logs, label):
    try:
        response = session.get(url, timeout=20)
        response.raise_for_status()
        add_log(logs, f"{label} 호출 성공")
        return response.json()
    except Exception as e:
        add_log(logs, f"{label} 호출 실패: {e}")
        return None


def parse_hour_label_to_index(label):
    m = re.search(r"(\d+)", str(label))
    if not m:
        return None
    hour = int(m.group(1))
    if hour == 24:
        return 0
    if 1 <= hour <= 23:
        return hour
    if 0 <= hour <= 23:
        return hour
    return None


def build_hourly_map_from_rows(rows):
    hourly = {}
    if not isinstance(rows, list):
        return hourly
    for row in rows:
        hour_idx = parse_hour_label_to_index(row.get("MR_HHMI") or row.get("MR_HHMI2"))
        if hour_idx is None:
            continue
        hourly[hour_idx] = parse_number(row.get("F_AP_QT"))
    return hourly


def summarize_band_loads_from_hourly(hourly_map, season):
    if not hourly_map:
        return None
    full_values = [hourly_map[h] for h in sorted(hourly_map) if hourly_map[h] > 0]
    if not full_values:
        return None
    base_avg = sum(full_values) / len(full_values)

    band_values = {"경부하": [], "중간부하": [], "최대부하": []}
    for hour, value in hourly_map.items():
        if value <= 0:
            continue
        label = hour_to_label(hour, season)
        band_values[label].append(value)

    band_avg = {}
    for label in ["경부하", "중간부하", "최대부하"]:
        vals = band_values[label]
        band_avg[label] = sum(vals) / len(vals) if vals else base_avg

    return {
        "base_avg_kw": base_avg,
        "off_peak_kw": band_avg["경부하"],
        "mid_peak_kw": band_avg["중간부하"],
        "peak_kw": band_avg["최대부하"],
        "off_ratio": band_avg["경부하"] / base_avg if base_avg else 1.0,
        "mid_ratio": band_avg["중간부하"] / base_avg if base_avg else 1.0,
        "peak_ratio": band_avg["최대부하"] / base_avg if base_avg else 1.0,
    }


def sum_latest_12_months(rows):
    parsed = []
    for row in rows or []:
        year = int(parse_number(row.get("MR_YEAR")))
        month = int(parse_number(row.get("MR_MONTH")))
        usage = parse_number(row.get("F_AP_QT"))
        if year and month and usage > 0:
            parsed.append({"year": year, "month": month, "usage": usage})
    if not parsed:
        return 0.0
    parsed.sort(key=lambda x: (x["year"], x["month"]), reverse=True)
    return sum(r["usage"] for r in parsed[:12])


def extract_rates_from_contract_info(contract_info):
    off_peak = 0.0
    mid_peak = 0.0
    peak = 0.0
    basic_charge = 0.0
    details = contract_info.get("details") if isinstance(contract_info, dict) else []
    for item in details or []:
        time_cd = str(item.get("TIME_CD", "")).strip()
        if not basic_charge:
            basic_charge = parse_number(item.get("BASE_BILL_UCOST"))
        spring_rate = parse_number(item.get("SEASEN_1"))
        if time_cd == "1":
            off_peak = spring_rate
        elif time_cd == "2":
            mid_peak = spring_rate
        elif time_cd == "3":
            peak = spring_rate
    return basic_charge, off_peak, mid_peak, peak

# =========================================================
# Selenium 보조 함수
# =========================================================
def find_first(driver, selectors):
    for by, selector in selectors:
        try:
            elems = driver.find_elements(by, selector)
            if elems:
                return elems[0]
        except Exception:
            continue
    return None


def click_first(driver, selectors, logs, label):
    for by, selector in selectors:
        try:
            elems = driver.find_elements(by, selector)
            if elems:
                driver.execute_script("arguments[0].click();", elems[0])
                add_log(logs, "{0} 클릭 성공: {1}".format(label, selector))
                return True
        except Exception:
            continue
    add_log(logs, "{0} 클릭 실패".format(label))
    return False


# =========================================================
# 파워플래너 추출 함수
# =========================================================

def scrape_kepco_power_planner(user_id, user_pw):
    logs = []

    if not has_module("selenium"):
        return {
            "status": "error",
            "message": "현재 실행 중인 Python에 selenium이 설치되어 있지 않습니다.",
            "logs": "실행 파이썬: {0}\nselenium 모듈 없음".format(sys.executable),
        }

    try:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
    except Exception as e:
        return {
            "status": "error",
            "message": "selenium import 중 오류가 발생했습니다.",
            "logs": "실행 파이썬: {0}\nimport 오류: {1}".format(sys.executable, str(e)),
        }

    driver = None
    try:
        env_info = get_runtime_environment_summary()
        add_log(logs, "실행 파이썬: {0}".format(env_info["python"]))
        add_log(logs, "플랫폼: {0}".format(env_info["platform"]))
        add_log(logs, "chrome path: {0}".format(env_info["chrome_path"]))
        add_log(logs, "chrome version: {0}".format(env_info["chrome_version"]))
        add_log(logs, "chromedriver path: {0}".format(env_info["chromedriver_path"]))
        add_log(logs, "chromedriver version: {0}".format(env_info["chromedriver_version"]))

        driver = build_chrome_driver(logs)
        wait = WebDriverWait(driver, 25)

        login_url = "https://pp.kepco.co.kr/intro.do"
        driver.get(login_url)
        add_log(logs, "로그인 페이지 접속")
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']")))

        id_selectors = [
            (By.CSS_SELECTOR, "input[type='text']"),
            (By.CSS_SELECTOR, "input[placeholder*='아이디']"),
            (By.NAME, "id"),
            (By.ID, "id"),
        ]
        pw_selectors = [
            (By.CSS_SELECTOR, "input[type='password']"),
            (By.CSS_SELECTOR, "input[placeholder*='비밀번호']"),
            (By.NAME, "pw"),
            (By.ID, "pw"),
        ]
        login_btn_selectors = [
            (By.XPATH, "//button[contains(., '로그인')]"),
            (By.XPATH, "//a[contains(., '로그인')]"),
            (By.XPATH, "//input[@value='로그인']"),
            (By.CSS_SELECTOR, "button"),
            (By.CSS_SELECTOR, "a.btn_login"),
            (By.CSS_SELECTOR, "input[type='submit']"),
        ]

        id_box = find_first(driver, id_selectors)
        pw_box = find_first(driver, pw_selectors)
        if id_box is None or pw_box is None:
            return {"status": "error", "message": "아이디 또는 비밀번호 입력창을 찾지 못했습니다.", "logs": "\n".join(logs)}

        id_box.clear()
        id_box.send_keys(user_id)
        pw_box.clear()
        pw_box.send_keys(user_pw)
        add_log(logs, "아이디/비밀번호 입력 완료")

        if not click_first(driver, login_btn_selectors, logs, "로그인 버튼"):
            return {"status": "error", "message": "로그인 버튼 selector를 찾지 못했습니다.", "logs": "\n".join(logs)}

        time.sleep(3)
        session = create_requests_session_from_driver(driver)

        for page_url in [
            "https://pp.kepco.co.kr/rm/rm0101.do?menu_id=O010101",
            "https://pp.kepco.co.kr/rs/rs0106.do?menu_id=O010206",
            "https://pp.kepco.co.kr/rp/rp0103.do?menu_id=O010303",
        ]:
            try:
                session.get(page_url, timeout=15)
            except Exception:
                pass

        result = {
            "contract_kind": "",
            "basic_charge_unit": 0.0,
            "power_bill_kw": 0.0,
            "max_demand_kw": 0.0,
            "annual_usage_kwh": 0.0,
            "off_peak_rate": 0.0,
            "mid_peak_rate": 0.0,
            "peak_rate": 0.0,
            "primary_voltage_kv": 22.9,
            "contract_power_kw": 0.0,
            "supply_voltage_text": "",
            "yearly_bill_won": 0.0,
            "voltage_class": "",
            "auto_avg_base_kw": 0.0,
            "auto_off_peak_kw": 0.0,
            "auto_mid_peak_kw": 0.0,
            "auto_peak_kw": 0.0,
            "auto_off_ratio": 0.0,
            "auto_mid_ratio": 0.0,
            "auto_peak_ratio": 0.0,
        }

        contract_info = fetch_json(session, "https://pp.kepco.co.kr/rm/rm0101_contract_info.do", logs, "rm0101_contract_info.do") or {}
        smart_summary = fetch_json(session, "https://pp.kepco.co.kr/rm/getRM0101.do", logs, "getRM0101.do") or {}
        selected_period_rows = fetch_json(session, "https://pp.kepco.co.kr/rs/rs0106_chart.do", logs, "rs0106_chart.do")
        selected_hour_rows = fetch_json(session, "https://pp.kepco.co.kr/rs/rs0101N_hour.do", logs, "rs0101N_hour.do")
        monthly_pf_rows = fetch_json(session, "https://pp.kepco.co.kr/rp/rp0103_usage_pf_list.do", logs, "rp0103_usage_pf_list.do")
        if monthly_pf_rows is None:
            monthly_pf_rows = fetch_json(session, "https://pp.kepco.co.kr/rp/p0103_usage_pf_list.do", logs, "p0103_usage_pf_list.do")
        if monthly_pf_rows is None:
            monthly_pf_rows = []

        contract_kind = (contract_info.get("CNTR_KND_NM") or smart_summary.get("CNTR_KND_NM") or "").strip()
        result["contract_kind"] = contract_kind
        result["contract_power_kw"] = parse_number(contract_info.get("CNTR_PWR")) or parse_number(smart_summary.get("CNTR_PWR"))
        result["power_bill_kw"] = parse_number(smart_summary.get("JOJ_KW")) or parse_number(smart_summary.get("CNTR_PWR"))
        result["max_demand_kw"] = parse_number(smart_summary.get("REAL_MAX_PWR")) or parse_number(smart_summary.get("MAX_PWR"))

        basic_charge_unit, off_peak_rate, mid_peak_rate, peak_rate = extract_rates_from_contract_info(contract_info)
        result["basic_charge_unit"] = basic_charge_unit or parse_number(smart_summary.get("BASE_BILL_UCOST"))
        result["off_peak_rate"] = off_peak_rate
        result["mid_peak_rate"] = mid_peak_rate
        result["peak_rate"] = peak_rate

        voltage_class = classify_voltage_from_contract_kind(contract_kind, contract_info.get("SUPPLY_TEXT", ""))
        if not voltage_class:
            lhv = str(contract_info.get("LHV_CLCD") or smart_summary.get("LHV_CLCD") or "")
            voltage_class = {"0": "저압", "1": "고압A", "2": "고압B", "3": "고압C"}.get(lhv, "")
        result["voltage_class"] = voltage_class or get_voltage_class(result["primary_voltage_kv"])
        result["primary_voltage_kv"] = primary_voltage_from_class(result["voltage_class"], contract_info.get("SUPPLY_TEXT", ""))
        result["supply_voltage_text"] = describe_voltage_class(result["voltage_class"], result["primary_voltage_kv"])

        result["annual_usage_kwh"] = sum_latest_12_months(monthly_pf_rows)
        result["yearly_bill_won"] = parse_number(smart_summary.get("PREDICT_TOTAL_CHARGE")) or parse_number(smart_summary.get("TOTAL_CHARGE"))

        hourly_map = build_hourly_map_from_rows(selected_period_rows or [])
        if not hourly_map:
            hourly_map = build_hourly_map_from_rows(selected_hour_rows or [])
        band_summary = summarize_band_loads_from_hourly(hourly_map, "봄·가을")
        if band_summary:
            result["auto_avg_base_kw"] = band_summary["base_avg_kw"]
            result["auto_off_peak_kw"] = band_summary["off_peak_kw"]
            result["auto_mid_peak_kw"] = band_summary["mid_peak_kw"]
            result["auto_peak_kw"] = band_summary["peak_kw"]
            result["auto_off_ratio"] = band_summary["off_ratio"]
            result["auto_mid_ratio"] = band_summary["mid_ratio"]
            result["auto_peak_ratio"] = band_summary["peak_ratio"]
            add_log(logs, "시간대별 평균부하 자동 산출 성공")

        if result["annual_usage_kwh"] <= 0 and result["auto_avg_base_kw"] > 0:
            result["annual_usage_kwh"] = result["auto_avg_base_kw"] * 24 * 365

        if result["contract_power_kw"] <= 0 and result["power_bill_kw"] > 0:
            result["contract_power_kw"] = result["power_bill_kw"]
        if result["power_bill_kw"] <= 0 and result["max_demand_kw"] > 0:
            result["power_bill_kw"] = result["max_demand_kw"]

        return {
            "status": "success",
            "message": "파워플래너 값 추출을 완료했습니다.",
            "logs": "\n".join(logs),
            **result,
        }

    except Exception as e:
        add_log(logs, "예외 발생: {0}".format(str(e)))
        add_log(logs, traceback.format_exc())
        return {
            "status": "error",
            "message": "자동화 중 오류 발생: {0}".format(str(e)),
            "logs": "\n".join(logs),
        }
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass


# =========================================================
# PDF 생성 함수
# =========================================================
def build_pdf_bytes_report(
    site_name,
    season,
    created_at_text,
    calc_voltage_basis,
    primary_voltage_kv,
    secondary_voltage_v,
    new_voltage,
    current_voltage_for_calc,
    voltage_drop_pct,
    voltage_drop_v,
    primary_voltage_class,
    tariff_voltage_class,
    contract_kind,
    tariff_choice_desc,
    cvrf_value,
    z,
    i,
    p,
    correction_factor,
    reliability_text,
    current_tap,
    target_tap,
    tap_step_percent,
    avg_saved_kw,
    saving_rate_total,
    daily_saved_kwh,
    monthly_saved_kwh,
    yearly_saved_kwh,
    daily_cost_saving,
    monthly_cost_saving,
    yearly_cost_saving,
    operation_mode,
    operating_start,
    operating_end,
    day_operation_hours,
    active_days_per_year,
    holiday_reflect,
    holiday_count,
    load_unit,
    load_summary_df,
    period_df,
    hourly_pdf_df,
    rate_summary_df,
    season_time_band_df,
    tap_compare_df,
    fig_line_buf,
    fig_bar_buf,
):
    buffer = io.BytesIO()
    global CURRENT_PDF_FONT_NAME
    font_info = get_korean_font_info()
    font_name = font_info["pdf_name"]
    CURRENT_PDF_FONT_NAME = font_name

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=14 * mm,
        bottomMargin=14 * mm,
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="KTitle",
        fontName=font_name,
        fontSize=18,
        leading=22,
        alignment=TA_CENTER,
        spaceAfter=8,
    ))
    styles.add(ParagraphStyle(
        name="KSubTitle",
        fontName=font_name,
        fontSize=10,
        leading=13,
        alignment=TA_CENTER,
        textColor=colors.grey,
        spaceAfter=10,
    ))
    styles.add(ParagraphStyle(
        name="KHeading",
        fontName=font_name,
        fontSize=12,
        leading=15,
        alignment=TA_LEFT,
        spaceBefore=4,
        spaceAfter=6,
    ))
    styles.add(ParagraphStyle(
        name="KBody",
        fontName=font_name,
        fontSize=9,
        leading=12,
        alignment=TA_LEFT,
    ))

    elements = []

    # 1페이지
    elements.append(Paragraph("CVR 운영형 계산 결과 보고서", styles["KTitle"]))
    elements.append(
        Paragraph(
            f"{site_name} / {season} / 생성일시 {created_at_text}",
            styles["KSubTitle"],
        )
    )

    elements.append(Paragraph("1. 핵심 결과", styles["KHeading"]))
    core_data = [
        ["항목", "값", "항목", "값", "항목", "값"],
        ["예상 전압 저감률", f"{voltage_drop_pct:.2f}%", "예상 절감 전력", f"{avg_saved_kw:,.1f} kW", "예상 절감률", f"{saving_rate_total:.2f}%"],
        ["예상 일 절감량", f"{daily_saved_kwh:,.1f} kWh", "예상 월 절감량", f"{monthly_saved_kwh:,.1f} kWh", "예상 연 절감량", f"{yearly_saved_kwh:,.1f} kWh"],
        ["예상 일 절감요금", f"{daily_cost_saving:,.0f} 원", "예상 월 절감요금", f"{monthly_cost_saving:,.0f} 원", "예상 연 절감요금", f"{yearly_cost_saving:,.0f} 원"],
    ]
    elements.append(make_two_col_table(core_data, col_widths=[32*mm, 28*mm, 32*mm, 28*mm, 32*mm, 28*mm], font_name=font_name))
    elements.append(Spacer(1, 8))

    elements.append(Paragraph("2. 계산 조건 요약", styles["KHeading"]))
    cond_left = [
        ["항목", "값"],
        ["사업장명", site_name],
        ["계절", season],
        ["계산 기준 전압", "2차측" if calc_voltage_basis == "2차측 전압 사용" else "1차측"],
        ["1차측 수전전압", f"{primary_voltage_kv:,.1f} kV"],
        ["2차측 평균전압", f"{secondary_voltage_v:,.1f} V"],
        ["변경 후 계산 기준 전압", f"{new_voltage:,.1f} V"],
        ["전압 감소 예상치", f"{voltage_drop_v:,.1f} V"],
        ["자동 판정 전압등급", describe_voltage_class(primary_voltage_class, primary_voltage_kv)],
        ["요금 적용 전압등급", describe_voltage_class(tariff_voltage_class, primary_voltage_kv)],
    ]
    cond_right = [
        ["항목", "값"],
        ["요금 종별", contract_kind if contract_kind else "산업용 을 / 선택 II"],
        ["선택 요금제 설명", tariff_choice_desc],
        ["적용 CVRf", f"{cvrf_value:.2f}"],
        ["적용 ZIP 비율", f"Z {z:.2f} / I {i:.2f} / P {p:.2f}"],
        ["보정계수", f"{correction_factor:.2f}"],
        ["예상 신뢰도", reliability_text],
        ["현재 탭", str(current_tap)],
        ["변경 탭", str(target_tap)],
        ["탭 1스텝당 전압 변화율", f"{tap_step_percent:.2f} %"],
    ]
    cond_tbl = Table(
        [
            [
                make_two_col_table(cond_left, col_widths=[38*mm, 50*mm], font_name=font_name, font_size=8.5),
                make_two_col_table(cond_right, col_widths=[38*mm, 50*mm], font_name=font_name, font_size=8.5),
            ]
        ],
        colWidths=[88*mm, 88*mm],
    )
    cond_tbl.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    elements.append(cond_tbl)
    elements.append(Spacer(1, 8))

    elements.append(Paragraph("3. 현재안 / 제안안 비교", styles["KHeading"]))
    current_voltage_display_kv = current_voltage_for_calc / 1000.0
    proposed_voltage_display_kv = new_voltage / 1000.0

    current_vs_proposed = [
        ["구분", "현재", "제안", "기대 효과"],
        [
            "계산 기준 전압",
            f"{current_voltage_display_kv:,.2f} kV",
            f"{proposed_voltage_display_kv:,.2f} kV",
            f"{voltage_drop_v:,.1f} V 저감",
        ],
        [
            "평균 절감전력",
            "0.0 kW",
            f"{avg_saved_kw:,.1f} kW",
            f"{avg_saved_kw:,.1f} kW 증가",
        ],
        [
            "연 절감 비용",
            "0 원",
            f"{yearly_cost_saving:,.0f} 원",
            f"{yearly_cost_saving:,.0f} 원 증가",
        ],
    ]
    elements.append(make_two_col_table(current_vs_proposed, col_widths=[42*mm, 42*mm, 42*mm, 50*mm], font_name=font_name))
    elements.append(PageBreak())

    # 2페이지
    elements.append(Paragraph("4. 현재 적용 요금 기준", styles["KHeading"]))
    tariff_data = [
        ["항목", "값"],
        ["현재 계절", season],
        ["1차측 수전전압", f"{primary_voltage_kv:,.1f} kV"],
        ["자동 판정 전압등급", describe_voltage_class(primary_voltage_class, primary_voltage_kv)],
        ["요금 적용 전압등급", describe_voltage_class(tariff_voltage_class, primary_voltage_kv)],
        ["요금 종별", "산업용"],
        ["세부 구분", "을"],
        ["선택 요금제", "선택 II"],
        ["선택 요금제 설명", "경부하, 중간부하, 최대부하 단가가 균형적으로 구성된 일반적인 요금 구조입니다."],
        ["요금 적용 방식", "직접 입력"],
        ["경부하 단가", f"{rate_summary_df.loc[0, '요금(원/kWh)']:.1f} 원/kWh"],
        ["중간부하 단가", f"{rate_summary_df.loc[1, '요금(원/kWh)']:.1f} 원/kWh"],
        ["최대부하 단가", f"{rate_summary_df.loc[2, '요금(원/kWh)']:.1f} 원/kWh"],
    ]
    elements.append(make_two_col_table(tariff_data, col_widths=[55*mm, 115*mm], font_name=font_name, font_size=8.8))
    elements.append(Spacer(1, 8))

    elements.append(Paragraph("5. 그래프", styles["KHeading"]))
    elements.append(Paragraph("5-1. 시간대별 전력사용량 그래프", styles["KBody"]))
    elements.append(Spacer(1, 4))
    elements.append(RLImage(fig_line_buf, width=180*mm, height=86*mm))
    elements.append(PageBreak())

    # 3페이지
    elements.append(Paragraph("5-2. 탭별 예상 절감전력 비교", styles["KBody"]))
    elements.append(Spacer(1, 4))
    elements.append(RLImage(fig_bar_buf, width=180*mm, height=86*mm))
    elements.append(Spacer(1, 10))

    elements.append(Paragraph("6. 계산 결과 요약표", styles["KHeading"]))
    summary_result_data = [
        ["모델", "예상 절감률(%)", "예상 평균 절감전력(kW)", "예상 일 절감량(kWh)", "예상 연 절감량(kWh)", "예상 일 절감요금(원)", "예상 연 절감요금(원)"],
        [
            "운영시간/요일 반영 결과",
            f"{saving_rate_total:.3f}",
            f"{avg_saved_kw:.3f}",
            f"{daily_saved_kwh:.3f}",
            f"{yearly_saved_kwh:,.1f}",
            f"{daily_cost_saving:,.1f}",
            f"{yearly_cost_saving:,.0f}",
        ],
    ]
    elements.append(make_two_col_table(
        summary_result_data,
        col_widths=[36*mm, 24*mm, 30*mm, 28*mm, 28*mm, 28*mm, 28*mm],
        font_name=font_name,
        font_size=7.6
    ))
    elements.append(Spacer(1, 8))

    elements.append(Paragraph("7. 운영 조건 요약", styles["KHeading"]))
    op_data = [["가동 요일", "가동 시작", "가동 종료", "일 운영시간(h)", "연 가동일수(일)", "공휴일 비가동 반영", "연 공휴일 수"]]
    op_data.append([
        operation_mode,
        f"{operating_start:02d}:00",
        f"{operating_end:02d}:00",
        f"{day_operation_hours}",
        f"{active_days_per_year}",
        "예" if holiday_reflect else "아니오",
        f"{holiday_count}",
    ])
    elements.append(make_two_col_table(op_data, col_widths=[30*mm, 22*mm, 22*mm, 22*mm, 25*mm, 28*mm, 20*mm], font_name=font_name, font_size=8))
    elements.append(Spacer(1, 8))

    elements.append(Paragraph("8. 적용 요금표", styles["KHeading"]))
    rate_data = [["구분", "요금(원/kWh)"]]
    for _, row in rate_summary_df.iterrows():
        rate_data.append([row["구분"], f"{row['요금(원/kWh)']:.1f}"])
    elements.append(make_two_col_table(rate_data, col_widths=[60*mm, 40*mm], font_name=font_name))
    elements.append(Spacer(1, 8))

    elements.append(Paragraph("9. 계절별 시간대 구분표", styles["KHeading"]))
    season_band_data = [["계절", "경부하 시간", "중간부하 시간", "최대부하 시간"]]
    for _, row in season_time_band_df.iterrows():
        season_band_data.append([
            row["계절"],
            row["경부하 시간"],
            row["중간부하 시간"],
            row["최대부하 시간"],
        ])
    elements.append(make_two_col_table(season_band_data, col_widths=[22*mm, 52*mm, 62*mm, 44*mm], font_name=font_name, font_size=8))
    elements.append(Spacer(1, 8))

    elements.append(Paragraph("10. 시간대별 평균 부하 입력 요약", styles["KHeading"]))
    load_data = [["구분", f"입력 평균부하({load_unit})", "적용 운영시간(h/day)"]]
    for _, row in load_summary_df.iterrows():
        load_data.append([
            row["구분"],
            f"{row[f'입력 평균부하({load_unit})']}",
            f"{row['적용 운영시간(h/day)']}",
        ])
    elements.append(make_two_col_table(load_data, col_widths=[40*mm, 50*mm, 45*mm], font_name=font_name))
    elements.append(PageBreak())

    # 4페이지
    elements.append(Paragraph("11. 시간대 구간별 계산 결과", styles["KHeading"]))
    period_data = [[
        "구분", "운영시간수(h/day)", "평균 부하(kW)", "사용전력량(kWh/day)",
        "절감률(%)", "절감전력(kW)", "절감전력량(kWh/day)", "요금단가(원/kWh)", "절감요금(원/day)"
    ]]
    for _, row in period_df.iterrows():
        period_data.append([
            row["구분"],
            f"{int(row['운영시간(h/day)'])}",
            f"{row['평균부하(kW)']:,.0f}",
            f"{row['사용전력량(kWh/day)']:,.0f}",
            f"{row['절감률(%)']:.3f}",
            f"{row['절감전력(kW)']:,.3f}",
            f"{row['절감전력량(kWh/day)']:,.3f}",
            f"{row['적용단가(원/kWh)']:,.0f}",
            f"{row['절감요금(원/day)']:,.1f}",
        ])
    elements.append(make_two_col_table(
        period_data,
        col_widths=[18*mm, 22*mm, 23*mm, 28*mm, 18*mm, 20*mm, 26*mm, 20*mm, 26*mm],
        font_name=font_name,
        font_size=7.6
    ))
    elements.append(Spacer(1, 8))

    elements.append(Paragraph("12. 24시간 시간대별 상세 결과", styles["KHeading"]))
    hourly_data = [[
        "시간", "시간번호", "구분", "운영여부", "전력사용량(kW)",
        "절감률(%)", "절감전력(kW)", "절감전력량(kWh)",
        "요금단가(원/kWh)", "절감요금(원)"
    ]]
    for _, row in hourly_pdf_df.iterrows():
        hourly_data.append([
            row["시간"],
            f"{int(row['시간번호'])}",
            row["구분"],
            row["운영여부"],
            f"{row['전력사용량(kW)']:,.0f}",
            f"{row['절감률(%)']:.3f}",
            f"{row['절감전력(kW)']:,.3f}",
            f"{row['절감전력량(kWh)']:,.3f}",
            f"{row['요금단가(원/kWh)']:,.0f}",
            f"{row['절감요금(원)']:,.1f}",
        ])
    hourly_first, hourly_second = split_even_rows(hourly_data)

    elements.append(make_long_table(
        hourly_first,
        col_widths=[16*mm, 13*mm, 15*mm, 16*mm, 22*mm, 15*mm, 18*mm, 19*mm, 18*mm, 20*mm],
        font_name=font_name,
        font_size=6.9
    ))
    elements.append(PageBreak())

    # 5페이지
    elements.append(make_long_table(
        hourly_second,
        col_widths=[16*mm, 13*mm, 15*mm, 16*mm, 22*mm, 15*mm, 18*mm, 19*mm, 18*mm, 20*mm],
        font_name=font_name,
        font_size=6.9
    ))
    elements.append(Spacer(1, 8))

    elements.append(Paragraph("13. 탭별 비교표", styles["KHeading"]))
    tap_data = [[
        "탭", "계산 기준 전압(V)", "전압 저감률(%)", "절감률(%)",
        "평균 절감전력(kW)", "일 절감량(kWh)", "연 절감량(kWh)",
        "일 절감요금(원)", "연 절감요금(원)"
    ]]
    for _, row in tap_compare_df.iterrows():
        tap_data.append([
            f"{int(row['탭'])}",
            f"{row['계산 기준 전압(V)']:,.1f}",
            f"{row['전압 저감률(%)']:.2f}",
            f"{row['절감률(%)']:.3f}",
            f"{row['평균 절감전력(kW)']:,.3f}",
            f"{row['일 절감량(kWh)']:,.3f}",
            f"{row['연 절감량(kWh)']:,.1f}",
            f"{row['일 절감요금(원)']:,.0f}",
            f"{row['연 절감요금(원)']:,.0f}",
        ])
    elements.append(make_long_table(
        tap_data,
        col_widths=[10*mm, 28*mm, 18*mm, 15*mm, 22*mm, 20*mm, 23*mm, 20*mm, 23*mm],
        font_name=font_name,
        font_size=6.8
    ))

    doc.build(elements, onFirstPage=add_page_number, onLaterPages=add_page_number)
    pdf_value = buffer.getvalue()
    buffer.close()
    return pdf_value


# =========================================================
# 화면 레이아웃 시작
# =========================================================
st.title("CVR 운영형 계산기")
st.caption("실사용 및 미팅용 추정 화면입니다. 최종 적용 전에는 최신 계약요금표 및 실측 데이터 검토가 필요합니다.")

with st.expander("실행 환경 확인", expanded=False):
    font_info_debug = get_korean_font_info()
    st.write("현재 실행 중인 Python:", sys.executable)
    st.write("Python 버전:", sys.version)
    st.write("selenium 설치 여부:", "설치됨" if has_module("selenium") else "없음")
    st.write("webdriver_manager 설치 여부:", "설치됨" if has_module("webdriver_manager") else "없음")
    st.write("reportlab 설치 여부:", "설치됨" if has_module("reportlab") else "없음")
    st.write("matplotlib 설치 여부:", "설치됨" if has_module("matplotlib") else "없음")
    st.write("PDF 폰트:", font_info_debug["pdf_name"])
    st.write("그래프 폰트:", font_info_debug["mpl_name"])
    st.write("폰트 경로:", font_info_debug["font_path"] or "기본 폰트 사용")
    env_info_debug = get_runtime_environment_summary()
    st.write("Chrome 경로:", env_info_debug["chrome_path"])
    st.write("Chrome 버전:", env_info_debug["chrome_version"])
    st.write("ChromeDriver 경로:", env_info_debug["chromedriver_path"])
    st.write("ChromeDriver 버전:", env_info_debug["chromedriver_version"])

    col_install_1, col_install_2, col_install_3, col_install_4 = st.columns(4)

    with col_install_1:
        if st.button("selenium / webdriver-manager 설치"):
            with st.spinner("설치 중입니다..."):
                rc, out, err = install_package_for_current_python(["selenium", "webdriver-manager"])
                if rc == 0:
                    st.success("설치 완료. 앱 재실행 권장")
                else:
                    st.error("설치 실패")
                st.code("STDOUT:\n{0}\n\nSTDERR:\n{1}".format(out, err))

    with col_install_2:
        if st.button("pip 업그레이드"):
            with st.spinner("업그레이드 중입니다..."):
                rc, out, err = install_package_for_current_python(["--upgrade", "pip"])
                if rc == 0:
                    st.success("pip 업그레이드 완료")
                else:
                    st.error("pip 업그레이드 실패")
                st.code("STDOUT:\n{0}\n\nSTDERR:\n{1}".format(out, err))

    with col_install_3:
        if st.button("reportlab 설치"):
            with st.spinner("설치 중입니다..."):
                rc, out, err = install_package_for_current_python(["reportlab"])
                if rc == 0:
                    st.success("reportlab 설치 완료")
                else:
                    st.error("reportlab 설치 실패")
                st.code("STDOUT:\n{0}\n\nSTDERR:\n{1}".format(out, err))

    with col_install_4:
        if st.button("matplotlib 설치"):
            with st.spinner("설치 중입니다..."):
                rc, out, err = install_package_for_current_python(["matplotlib"])
                if rc == 0:
                    st.success("matplotlib 설치 완료")
                else:
                    st.error("matplotlib 설치 실패")
                st.code("STDOUT:\n{0}\n\nSTDERR:\n{1}".format(out, err))


st.markdown(
    """
    <style>
    .pp-right-sticky {
        position: -webkit-sticky;
        position: sticky;
        top: 72px;
        align-self: flex-start;
        height: fit-content;
        max-height: calc(100vh - 90px);
        overflow-y: auto;
        padding-right: 4px;
    }
    @media (max-width: 980px) {
        .pp-right-sticky {
            position: static;
            top: auto;
            max-height: none;
            overflow-y: visible;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


left, right = st.columns([1.7, 1.0])


# =========================================================
# 좌측 입력 섹션
# =========================================================
with left:
    st.subheader("입력 조건")

    site_name = colored_input("사업장명", st.text_input, "manual", value="예시 공장")
    season = colored_input("계절 선택", st.selectbox, "manual", options=["봄·가을", "여름", "겨울"], index=0)

    st.markdown("### 파워플래너 기반 빠른 입력")
    pp_mode = colored_input(
        "입력 방식",
        st.radio,
        "manual",
        options=["수동 입력", "파워플래너 값 수동 반영", "파워플래너 자동반영"],
        horizontal=True,
        index=1,
    )

    if pp_mode == "파워플래너 자동반영":
        with st.expander("파워플래너 계정 정보 입력", expanded=True):
            st.warning("보안정책, 사이트 구조 변경, 로딩 지연에 따라 실패할 수 있습니다. 배포 서버에서는 chromium/chromedriver가 설치되어 있어야 자동화가 동작합니다.")
            col1, col2 = st.columns(2)

            with col1:
                kepco_id = colored_input("파워플래너 아이디", st.text_input, "manual", key="pp_id")
            with col2:
                kepco_pw = colored_input("비밀번호", st.text_input, "manual", type="password", key="pp_pw")

            if st.button("파워플래너 값 자동 불러오기", type="primary"):
                if not kepco_id or not kepco_pw:
                    st.warning("아이디와 비밀번호를 모두 입력하세요.")
                else:
                    with st.spinner("파워플래너 값 추출 중입니다..."):
                        result = scrape_kepco_power_planner(kepco_id, kepco_pw)
                        st.session_state["pp_last_message"] = result.get("message", "")
                        st.session_state["pp_log_text"] = result.get("logs", "")

                        if result["status"] == "success":
                            st.session_state["pp_loaded"] = True
                            st.session_state["pp_contract_kind"] = result.get("contract_kind", "")
                            st.session_state["pp_basic_charge_unit"] = result.get("basic_charge_unit", 0.0)
                            st.session_state["pp_power_bill_kw"] = result.get("power_bill_kw", 0.0)
                            st.session_state["pp_max_demand_kw"] = result.get("max_demand_kw", 0.0)
                            st.session_state["pp_annual_usage_kwh"] = result.get("annual_usage_kwh", 0.0)
                            st.session_state["pp_off_peak_rate"] = result.get("off_peak_rate", 0.0)
                            st.session_state["pp_mid_peak_rate"] = result.get("mid_peak_rate", 0.0)
                            st.session_state["pp_peak_rate"] = result.get("peak_rate", 0.0)
                            st.session_state["pp_primary_voltage_kv"] = result.get("primary_voltage_kv", 154.0)
                            st.session_state["pp_contract_power_kw"] = result.get("contract_power_kw", 0.0)
                            st.session_state["pp_supply_voltage_text"] = result.get("supply_voltage_text", "")
                            st.session_state["pp_meter_read_day"] = result.get("meter_read_day", 0)
                            st.session_state["pp_yearly_bill_won"] = result.get("yearly_bill_won", 0.0)
                            st.session_state["pp_voltage_class"] = result.get("voltage_class", "")
                            st.session_state["pp_auto_avg_base_kw"] = result.get("auto_avg_base_kw", 0.0)
                            st.session_state["pp_auto_off_peak_kw"] = result.get("auto_off_peak_kw", 0.0)
                            st.session_state["pp_auto_mid_peak_kw"] = result.get("auto_mid_peak_kw", 0.0)
                            st.session_state["pp_auto_peak_kw"] = result.get("auto_peak_kw", 0.0)
                            st.session_state["pp_auto_off_ratio"] = result.get("auto_off_ratio", 0.0)
                            st.session_state["pp_auto_mid_ratio"] = result.get("auto_mid_ratio", 0.0)
                            st.session_state["pp_auto_peak_ratio"] = result.get("auto_peak_ratio", 0.0)
                            st.success(result["message"])
                        else:
                            st.error(result["message"])

            if st.session_state["pp_last_message"]:
                st.markdown("#### 자동화 결과 메시지")
                st.code(st.session_state["pp_last_message"])

            if st.session_state["pp_log_text"]:
                with st.expander("자동화 로그 보기", expanded=False):
                    st.code(st.session_state["pp_log_text"])

    if pp_mode == "파워플래너 값 수동 반영":
        st.info("파워플래너 화면에 보이는 값을 그대로 넣으면 계산에 반영됩니다.")
        c1, c2 = st.columns(2)

        with c1:
            pp_contract_kind_input = colored_input(
                "계약종별",
                st.text_input,
                "auto",
                value=st.session_state["pp_contract_kind"] or "산업용 을 / 선택 II",
            )
            pp_annual_usage_kwh_input = colored_input(
                "전년도 사용량(kWh)",
                st.number_input,
                "auto",
                min_value=0.0,
                value=float(st.session_state["pp_annual_usage_kwh"]),
                step=1000.0,
            )
            pp_power_bill_kw_input = colored_input(
                "요금적용전력(kW)",
                st.number_input,
                "auto",
                min_value=0.0,
                value=float(st.session_state["pp_power_bill_kw"]),
                step=1.0,
            )
            pp_max_demand_kw_input = colored_input(
                "최대수요전력(kW)",
                st.number_input,
                "auto",
                min_value=0.0,
                value=float(st.session_state["pp_max_demand_kw"]),
                step=1.0,
            )

        with c2:
            pp_primary_voltage_kv_input = colored_input(
                "1차측 수전전압(kV)",
                st.number_input,
                "verify",
                min_value=0.0,
                value=float(st.session_state["pp_primary_voltage_kv"]),
                step=0.1,
            )
            pp_basic_charge_unit_input = colored_input(
                "기본요금단가(원/kW)",
                st.number_input,
                "auto",
                min_value=0.0,
                value=float(st.session_state["pp_basic_charge_unit"]),
                step=1.0,
            )
            pp_off_peak_rate_input = colored_input(
                "경부하 단가(원/kWh)",
                st.number_input,
                "verify",
                min_value=0.0,
                value=float(st.session_state["pp_off_peak_rate"]),
                step=0.1,
            )
            pp_mid_peak_rate_input = colored_input(
                "중간부하 단가(원/kWh)",
                st.number_input,
                "verify",
                min_value=0.0,
                value=float(st.session_state["pp_mid_peak_rate"]),
                step=0.1,
            )
            pp_peak_rate_input = colored_input(
                "최대부하 단가(원/kWh)",
                st.number_input,
                "verify",
                min_value=0.0,
                value=float(st.session_state["pp_peak_rate"]),
                step=0.1,
            )

        if st.button("수동 입력값 적용", type="primary"):
            st.session_state["pp_loaded"] = True
            st.session_state["pp_contract_kind"] = pp_contract_kind_input
            st.session_state["pp_annual_usage_kwh"] = pp_annual_usage_kwh_input
            st.session_state["pp_power_bill_kw"] = pp_power_bill_kw_input
            st.session_state["pp_max_demand_kw"] = pp_max_demand_kw_input
            st.session_state["pp_primary_voltage_kv"] = pp_primary_voltage_kv_input
            st.session_state["pp_basic_charge_unit"] = pp_basic_charge_unit_input
            st.session_state["pp_off_peak_rate"] = pp_off_peak_rate_input
            st.session_state["pp_mid_peak_rate"] = pp_mid_peak_rate_input
            st.session_state["pp_peak_rate"] = pp_peak_rate_input
            st.success("파워플래너 값 적용 완료")

    pp_loaded = st.session_state["pp_loaded"]
    avg_kw_from_annual = annual_kwh_to_avg_kw(st.session_state["pp_annual_usage_kwh"])

    if pp_loaded:
        summary_avg_kw = st.session_state.get("pp_auto_avg_base_kw", 0.0) or avg_kw_from_annual
        st.success(
            "평균부하: {0:,.1f} kW / 계약종별: {1} / 공급방식: {2}".format(
                summary_avg_kw,
                st.session_state["pp_contract_kind"] or "-",
                st.session_state.get("pp_supply_voltage_text", "-") or "-",
            )
        )

    st.markdown("### 전압 조건")
    v1, v2, v3 = st.columns(3)

    with v1:
        primary_voltage_kv = colored_input(
            "1차측 수전전압(kV)",
            st.number_input,
            "verify",
            min_value=0.1,
            value=float(st.session_state["pp_primary_voltage_kv"] if pp_loaded else 154.0),
            step=0.1,
        )

    with v2:
        secondary_voltage_v = colored_input(
            "2차측 평균 전압(V)",
            st.number_input,
            "manual",
            min_value=1.0,
            value=22900.0,
            step=100.0,
        )

    with v3:
        calc_voltage_basis = colored_input(
            "CVR 계산 기준",
            st.selectbox,
            "manual",
            options=["2차측 전압 사용", "1차측 전압 사용"],
            index=0,
        )

    primary_voltage_class = get_voltage_class(primary_voltage_kv)
    st.info("자동 판정 전압등급: {0}".format(primary_voltage_class))

    current_voltage_for_calc = secondary_voltage_v if calc_voltage_basis == "2차측 전압 사용" else primary_voltage_kv * 1000.0

    t1, t2, t3 = st.columns(3)

    with t1:
        current_tap = colored_input("현재 탭", st.number_input, "manual", value=4, step=1)

    with t2:
        target_tap = colored_input("변경 탭", st.number_input, "manual", value=3, step=1)

    with t3:
        tap_step_percent = colored_input(
            "탭 1스텝당 전압 변화율(%)",
            st.number_input,
            "manual",
            min_value=0.1,
            value=1.25,
            step=0.05,
        )

    st.markdown("### 부하 조건")
    load_type = colored_input("부하 유형", st.selectbox, "manual", options=list(LOAD_TYPES.keys()), index=1)
    defaults = LOAD_TYPES[load_type]

    zip_custom = st.checkbox("ZIP 직접 입력", value=False)
    if zip_custom:
        z = colored_input("Z", st.slider, "manual", min_value=0.0, max_value=1.0, value=float(defaults["z"]), step=0.01)
        i = colored_input("I", st.slider, "manual", min_value=0.0, max_value=1.0, value=float(defaults["i"]), step=0.01)
        p = colored_input("P", st.slider, "manual", min_value=0.0, max_value=1.0, value=float(defaults["p"]), step=0.01)
        s = z + i + p
        if s == 0:
            z, i, p = defaults["z"], defaults["i"], defaults["p"]
        else:
            z, i, p = z / s, i / s, p / s
    else:
        z, i, p = defaults["z"], defaults["i"], defaults["p"]

    st.caption("현재 ZIP 합계: {0:.2f}".format(z + i + p))

    st.markdown("### 운영 조건")
    operation_mode = colored_input(
        "가동 요일",
        st.selectbox,
        "manual",
        options=["월~금 가동", "월~토 가동", "주7일 가동"],
        index=2,
    )

    o1, o2 = st.columns(2)
    with o1:
        operating_start = colored_input(
            "가동 시작",
            st.selectbox,
            "manual",
            options=list(range(24)),
            index=9,
            format_func=lambda x: "{0:02d}:00".format(x),
        )
    with o2:
        operating_end = colored_input(
            "가동 종료",
            st.selectbox,
            "manual",
            options=list(range(24)),
            index=9,
            format_func=lambda x: "{0:02d}:00".format(x),
        )

    holiday_reflect = st.checkbox("공휴일 비가동 반영", value=False)
    holiday_count = colored_input("연 공휴일 수", st.number_input, "manual", min_value=0, value=15, step=1)

    st.markdown("### 평균 부하 입력")
    auto_ratio_ready = (
        pp_mode == "파워플래너 자동반영"
        and pp_loaded
        and st.session_state.get("pp_auto_avg_base_kw", 0.0) > 0
    )

    input_mode = colored_input(
        "입력 방식",
        st.radio,
        "manual",
        options=["직접 입력", "파워플래너 자동 산출"],
        horizontal=True,
        index=1 if auto_ratio_ready else 0,
    )
    load_unit = colored_input("입력 단위", st.radio, "manual", options=["kW", "MW"], horizontal=True, index=1)
    unit_mul = 1000.0 if load_unit == "MW" else 1.0

    if input_mode == "파워플래너 자동 산출":
        default_avg_kw = (
            st.session_state.get("pp_auto_avg_base_kw", 0.0)
            if auto_ratio_ready and st.session_state.get("pp_auto_avg_base_kw", 0.0) > 0
            else (avg_kw_from_annual if pp_loaded and avg_kw_from_annual > 0 else 18000.0)
        )
        ratio_color = "auto" if auto_ratio_ready else "manual"

        if auto_ratio_ready:
            st.caption("파워플래너 시간대 평균부하를 기준으로 자동 산출되며, 필요 시 직접 수정할 수 있습니다.")

        avg_base_input = colored_input(
            f"기준 평균부하({load_unit})",
            st.number_input,
            ratio_color,
            min_value=0.0,
            value=float(default_avg_kw / unit_mul),
            step=1.0 if load_unit == "MW" else 1000.0,
        )
        avg_base_kw = avg_base_input * unit_mul

        r1, r2, r3 = st.columns(3)
        default_off_ratio = st.session_state.get("pp_auto_off_ratio", 0.92) if auto_ratio_ready else 0.92
        default_mid_ratio = st.session_state.get("pp_auto_mid_ratio", 1.00) if auto_ratio_ready else 1.00
        default_peak_ratio = st.session_state.get("pp_auto_peak_ratio", 1.10) if auto_ratio_ready else 1.10

        with r1:
            off_ratio = colored_input(
                "경부하 비율",
                st.number_input,
                ratio_color,
                min_value=0.10,
                value=float(default_off_ratio or 0.92),
                step=0.01,
            )
        with r2:
            mid_ratio = colored_input(
                "중간부하 비율",
                st.number_input,
                ratio_color,
                min_value=0.10,
                value=float(default_mid_ratio or 1.00),
                step=0.01,
            )
        with r3:
            peak_ratio = colored_input(
                "최대부하 비율",
                st.number_input,
                ratio_color,
                min_value=0.10,
                value=float(default_peak_ratio or 1.10),
                step=0.01,
            )

        auto_loads = avg_kw_to_timeband_loads(avg_base_kw, off_ratio, mid_ratio, peak_ratio)
        off_peak_kw = auto_loads["경부하"]
        mid_peak_kw = auto_loads["중간부하"]
        peak_kw = auto_loads["최대부하"]
    else:
        default_off_kw = st.session_state.get("pp_auto_off_peak_kw", 16000.0) if pp_loaded else 16000.0
        default_mid_kw = st.session_state.get("pp_auto_mid_peak_kw", 18000.0) if pp_loaded else 18000.0
        default_peak_kw = st.session_state.get("pp_auto_peak_kw", 20000.0) if pp_loaded else 20000.0
        off_peak_kw = colored_input(
            f"경부하 평균부하({load_unit})",
            st.number_input,
            "manual",
            min_value=0.0,
            value=float(default_off_kw / unit_mul),
            step=1.0 if load_unit == "MW" else 1000.0,
        ) * unit_mul
        mid_peak_kw = colored_input(
            f"중간부하 평균부하({load_unit})",
            st.number_input,
            "manual",
            min_value=0.0,
            value=float(default_mid_kw / unit_mul),
            step=1.0 if load_unit == "MW" else 1000.0,
        ) * unit_mul
        peak_kw = colored_input(
            f"최대부하 평균부하({load_unit})",
            st.number_input,
            "manual",
            min_value=0.0,
            value=float(default_peak_kw / unit_mul),
            step=1.0 if load_unit == "MW" else 1000.0,
        ) * unit_mul


# =========================================================
# 계산 로직
# =========================================================
tap_delta_steps = max(int(current_tap - target_tap), 0)
voltage_drop_pct = calc_tap_voltage_change(tap_step_percent, tap_delta_steps)
new_voltage = current_voltage_for_calc * (1 - voltage_drop_pct / 100.0)
voltage_drop_v = current_voltage_for_calc - new_voltage

active_hours, operating_hour_count = get_operating_hours_by_label(
    season=season,
    operating_start=operating_start,
    operating_end=operating_end,
)
active_days_per_year = get_active_days_per_year(
    operation_mode=operation_mode,
    holiday_count=int(holiday_count),
    include_holidays_as_shutdown=holiday_reflect,
)

loads_by_label = {
    "경부하": off_peak_kw,
    "중간부하": mid_peak_kw,
    "최대부하": peak_kw,
}
if (
    st.session_state["pp_loaded"]
    and st.session_state["pp_off_peak_rate"] > 0
    and st.session_state["pp_mid_peak_rate"] > 0
    and st.session_state["pp_peak_rate"] > 0
):
    rate_table = {
        season: {
            "경부하": float(st.session_state["pp_off_peak_rate"]),
            "중간부하": float(st.session_state["pp_mid_peak_rate"]),
            "최대부하": float(st.session_state["pp_peak_rate"]),
        }
    }
else:
    rate_table = safe_rate_table(primary_voltage_class)
tariff_voltage_class = primary_voltage_class

rows = []
daily_base_kwh = 0.0
daily_saved_kwh = 0.0
daily_cost_saving = 0.0

for label in ["경부하", "중간부하", "최대부하"]:
    hours = operating_hour_count[label]
    load_kw = loads_by_label[label]
    rate = rate_table[season][label]

    saving_rate_pct, saved_kw, saved_kwh = calc_average_result(
        load_kw, voltage_drop_pct, defaults["cvrf"], z, i, p, hours
    )
    base_kwh = load_kw * hours
    cost_save = saved_kwh * rate

    rows.append({
        "구분": label,
        "운영시간(h/day)": hours,
        "평균부하(kW)": round(load_kw, 2),
        "사용전력량(kWh/day)": round(base_kwh, 2),
        "절감률(%)": round(saving_rate_pct, 3),
        "절감전력(kW)": round(saved_kw, 2),
        "절감전력량(kWh/day)": round(saved_kwh, 2),
        "적용단가(원/kWh)": round(rate, 1),
        "절감요금(원/day)": round(cost_save, 1),
    })

    daily_base_kwh += base_kwh
    daily_saved_kwh += saved_kwh
    daily_cost_saving += cost_save

period_df = pd.DataFrame(rows)

saving_rate_total = (daily_saved_kwh / daily_base_kwh * 100.0) if daily_base_kwh else 0.0
day_operation_hours = sum(operating_hour_count.values())
avg_saved_kw = (daily_saved_kwh / day_operation_hours) if day_operation_hours else 0.0
monthly_saved_kwh = daily_saved_kwh * (active_days_per_year / 12.0 if active_days_per_year else 0.0)
yearly_saved_kwh = daily_saved_kwh * active_days_per_year
monthly_cost_saving = daily_cost_saving * (active_days_per_year / 12.0 if active_days_per_year else 0.0)
yearly_cost_saving = daily_cost_saving * active_days_per_year

hour_rows = []
for h in range(24):
    label = hour_to_label(h, season)
    operating = h in active_hours
    kw = loads_by_label[label] if operating else 0.0
    rate = rate_table[season][label]
    saving_rate_pct, saved_kw, saved_kwh = calc_average_result(kw, voltage_drop_pct, defaults["cvrf"], z, i, p, 1.0)
    hour_rows.append({
        "시간": "{0:02d}:00".format(h),
        "시간번호": h,
        "구분": label,
        "운영여부": "가동" if operating else "비가동",
        "전력사용량(kW)": round(kw, 2),
        "절감률(%)": round(saving_rate_pct, 3),
        "절감전력(kW)": round(saved_kw, 3),
        "절감전력량(kWh)": round(saved_kwh, 3),
        "요금단가(원/kWh)": round(rate, 1),
        "절감요금(원)": round(saved_kwh * rate, 1),
    })

hourly_df = pd.DataFrame(hour_rows)

# 탭별 비교표
max_compare_drop_pct = 7.5
tap_compare_rows = []
current_tap_int = int(current_tap)

# 현재 탭(기준점)도 함께 표시
base_compare_voltage = current_voltage_for_calc
base_compare_row = {
    "탭": current_tap_int,
    "계산 기준 전압(V)": round(base_compare_voltage, 1),
    "전압 저감률(%)": 0.0,
    "절감률(%)": 0.0,
    "평균 절감전력(kW)": 0.0,
    "일 절감량(kWh)": 0.0,
    "연 절감량(kWh)": 0.0,
    "일 절감요금(원)": 0.0,
    "연 절감요금(원)": 0.0,
}
tap_compare_rows.append(base_compare_row)

for tap in range(max(current_tap_int - 1, 1), 0, -1):
    delta_steps = max(current_tap_int - int(tap), 0)
    tap_voltage_drop_pct = calc_tap_voltage_change(tap_step_percent, delta_steps)
    if tap_voltage_drop_pct < 1.25 or tap_voltage_drop_pct > max_compare_drop_pct:
        continue
    tap_new_voltage = current_voltage_for_calc * (1 - tap_voltage_drop_pct / 100.0)

    tap_daily_saved_kwh = 0.0
    tap_daily_cost = 0.0

    for label in ["경부하", "중간부하", "최대부하"]:
        hours = operating_hour_count[label]
        load_kw = loads_by_label[label]
        rate = rate_table[season][label]
        tap_saving_rate_pct, tap_saved_kw, tap_saved_kwh = calc_average_result(
            load_kw,
            tap_voltage_drop_pct,
            defaults["cvrf"],
            z,
            i,
            p,
            hours
        )
        tap_daily_saved_kwh += tap_saved_kwh
        tap_daily_cost += tap_saved_kwh * rate

    tap_avg_saved_kw = (tap_daily_saved_kwh / day_operation_hours) if day_operation_hours else 0.0
    tap_saving_rate_total = (tap_daily_saved_kwh / daily_base_kwh * 100.0) if daily_base_kwh else 0.0
    tap_yearly_saved_kwh = tap_daily_saved_kwh * active_days_per_year
    tap_yearly_cost = tap_daily_cost * active_days_per_year

    tap_compare_rows.append({
        "탭": tap,
        "계산 기준 전압(V)": round(tap_new_voltage, 1),
        "전압 저감률(%)": round(tap_voltage_drop_pct, 2),
        "절감률(%)": round(tap_saving_rate_total, 3),
        "평균 절감전력(kW)": round(tap_avg_saved_kw, 3),
        "일 절감량(kWh)": round(tap_daily_saved_kwh, 3),
        "연 절감량(kWh)": round(tap_yearly_saved_kwh, 1),
        "일 절감요금(원)": round(tap_daily_cost, 0),
        "연 절감요금(원)": round(tap_yearly_cost, 0),
    })

tap_compare_df = pd.DataFrame(tap_compare_rows)
if not tap_compare_df.empty:
    tap_compare_df = tap_compare_df[
        (tap_compare_df["탭"] == current_tap_int) |
        (tap_compare_df["전압 저감률(%)"] == 0.0) |
        (
            (tap_compare_df["전압 저감률(%)"] >= 1.25) &
            (tap_compare_df["전압 저감률(%)"] <= max_compare_drop_pct)
        )
    ].copy()
    tap_compare_df = tap_compare_df.drop_duplicates(subset=["탭"], keep="first")
    tap_compare_df = tap_compare_df.sort_values(
        by=["탭"],
        ascending=[False]
    ).reset_index(drop=True)

# 요약 데이터프레임
rate_summary_df = pd.DataFrame({
    "구분": ["경부하", "중간부하", "최대부하"],
    "요금(원/kWh)": [
        rate_table[season]["경부하"],
        rate_table[season]["중간부하"],
        rate_table[season]["최대부하"],
    ]
})

season_time_band_df = pd.DataFrame([{
    "계절": season,
    "경부하 시간": format_hour_range(SEASON_SCHEDULE[season]["경부하"]),
    "중간부하 시간": format_hour_range(SEASON_SCHEDULE[season]["중간부하"]),
    "최대부하 시간": format_hour_range(SEASON_SCHEDULE[season]["최대부하"]),
}])

load_summary_df = pd.DataFrame([
    {"구분": "경부하", f"입력 평균부하({load_unit})": round(off_peak_kw / unit_mul, 3), "적용 운영시간(h/day)": operating_hour_count["경부하"]},
    {"구분": "중간부하", f"입력 평균부하({load_unit})": round(mid_peak_kw / unit_mul, 3), "적용 운영시간(h/day)": operating_hour_count["중간부하"]},
    {"구분": "최대부하", f"입력 평균부하({load_unit})": round(peak_kw / unit_mul, 3), "적용 운영시간(h/day)": operating_hour_count["최대부하"]},
])

correction_factor = round(saving_rate_total / voltage_drop_pct, 2) if voltage_drop_pct > 0 else 0.0
if correction_factor >= 0.9:
    reliability_text = "높음"
elif correction_factor >= 0.7:
    reliability_text = "보통"
else:
    reliability_text = "낮음"

contract_kind_display = st.session_state["pp_contract_kind"] if st.session_state["pp_contract_kind"] else "산업용 을 / 선택 II"
tariff_choice_desc = tariff_description("선택 II")
created_at_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# =========================================================
# 우측 결과 섹션
# =========================================================
with right:
    st.markdown('<div class="pp-right-sticky">', unsafe_allow_html=True)
    logo_col1, logo_col2, logo_col3 = st.columns([1, 1, 1])
    with logo_col2:
        try:
            st.image("logo.png", use_container_width=True)
        except Exception:
            st.caption("회사 로고 이미지를 표시하려면 작업 폴더에 'logo.png' 파일을 추가하세요.")

    st.markdown("### 🎨 입력항목 색상 범례")
    st.markdown(
        '<div style="background-color: #d4edda; padding: 6px; margin-bottom: 5px; border-radius: 5px; color: #333; font-weight: bold;">🟩 연녹색: 파워플래너 자동반영</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div style="background-color: #fff3cd; padding: 6px; margin-bottom: 5px; border-radius: 5px; color: #333; font-weight: bold;">🟨 연노랑: 파워플래너 자동반영 후 수동 검증 권장</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div style="background-color: #d1ecf1; padding: 6px; margin-bottom: 20px; border-radius: 5px; color: #333; font-weight: bold;">🟦 하늘색: 사용자가 수동으로 입력할 정보</div>',
        unsafe_allow_html=True,
    )

    st.subheader("계산 결과")
    m1, m2, m3 = st.columns(3)
    m1.metric("전압 저감률", "{0:.2f}%".format(voltage_drop_pct))
    m2.metric("평균 절감전력", "{0:,.1f} kW".format(avg_saved_kw))
    m3.metric("절감률", "{0:.2f}%".format(saving_rate_total))

    m4, m5, m6 = st.columns(3)
    m4.metric("일 절감량", "{0:,.1f} kWh".format(daily_saved_kwh))
    m5.metric("월 절감량", "{0:,.1f} kWh".format(monthly_saved_kwh))
    m6.metric("연 절감량", "{0:,.1f} kWh".format(yearly_saved_kwh))

    m7, m8, m9 = st.columns(3)
    m7.metric("일 절감요금", "{0:,.0f} 원".format(daily_cost_saving))
    m8.metric("월 절감요금", "{0:,.0f} 원".format(monthly_cost_saving))
    m9.metric("연 절감요금", "{0:,.0f} 원".format(yearly_cost_saving))

    st.caption("변경 후 계산 기준 전압: {0:,.1f}".format(new_voltage))

    if st.session_state["pp_loaded"]:
        st.markdown("### 파워플래너 반영값")
        st.write("- 계약종별: **{0}**".format(st.session_state["pp_contract_kind"] or "-"))
        st.write("- 전압등급: **{0}**".format(st.session_state.get("pp_voltage_class", "-") or "-"))
        st.write("- 기본요금단가: **{0:,.1f} 원/kW**".format(st.session_state["pp_basic_charge_unit"]))
        st.write("- 요금적용전력: **{0:,.1f} kW**".format(st.session_state["pp_power_bill_kw"]))
        st.write("- 최대수요전력: **{0:,.1f} kW**".format(st.session_state["pp_max_demand_kw"]))
        st.write("- 최근 12개월 사용량: **{0:,.1f} kWh**".format(st.session_state["pp_annual_usage_kwh"]))
        st.write("- 공급방식: **{0}**".format(st.session_state["pp_supply_voltage_text"] or "-"))
        st.write("- 검침일: **{0}일**".format(int(st.session_state.get("pp_meter_read_day", 0)) if st.session_state.get("pp_meter_read_day", 0) else "-"))
        if st.session_state.get("pp_auto_avg_base_kw", 0.0) > 0:
            st.write("- 자동 산출 기준 평균부하: **{0:,.1f} kW**".format(st.session_state["pp_auto_avg_base_kw"]))
            st.write("- 자동 산출 경/중/최 평균부하: **{0:,.1f} / {1:,.1f} / {2:,.1f} kW**".format(
                st.session_state.get("pp_auto_off_peak_kw", 0.0),
                st.session_state.get("pp_auto_mid_peak_kw", 0.0),
                st.session_state.get("pp_auto_peak_kw", 0.0),
            ))
            st.write("- 자동 산출 경/중/최 비율: **{0:.2f} / {1:.2f} / {2:.2f}**".format(
                st.session_state.get("pp_auto_off_ratio", 0.0),
                st.session_state.get("pp_auto_mid_ratio", 0.0),
                st.session_state.get("pp_auto_peak_ratio", 0.0),
            ))

    st.markdown("</div>", unsafe_allow_html=True)

st.divider()
st.subheader("그래프")
g1, g2 = st.columns(2)

with g1:
    fig_load = px.line(
        hourly_df,
        x="시간번호",
        y="전력사용량(kW)",
        markers=True,
        hover_data=["시간", "구분", "운영여부", "절감전력(kW)", "절감전력량(kWh)"],
        title="{0} 시간대별 전력사용량".format(season),
    )
    fig_load.update_xaxes(
        tickmode="array",
        tickvals=list(range(24)),
        ticktext=["{0:02d}".format(i) for i in range(24)],
    )
    st.plotly_chart(fig_load, use_container_width=True)

with g2:
    tap_chart_df = tap_compare_df.copy()
    if not tap_chart_df.empty:
        tap_chart_df["탭"] = pd.to_numeric(tap_chart_df["탭"], errors="coerce")
        tap_chart_df = tap_chart_df.sort_values(by="탭", ascending=True).reset_index(drop=True)
        fig_bar = go.Figure()
        fig_bar.add_bar(
            x=tap_chart_df["탭"].tolist(),
            y=tap_chart_df["평균 절감전력(kW)"].tolist(),
            text=[f"{v:.3f}" if float(v) != 0 else "0" for v in tap_chart_df["평균 절감전력(kW)"].tolist()],
            textposition="outside",
        )
        fig_bar.update_layout(title="탭별 예상 절감전력 비교")
        fig_bar.update_xaxes(
            title_text="탭",
            autorange="reversed",
            tickmode="array",
            tickvals=tap_chart_df["탭"].tolist(),
            ticktext=[str(int(v)) for v in tap_chart_df["탭"].tolist()],
        )
        fig_bar.update_yaxes(title_text="평균 절감전력(kW)")
        st.plotly_chart(fig_bar, use_container_width=True)
    else:
        st.info("표시할 탭별 비교 데이터가 없습니다.")

st.divider()
st.subheader("상세 데이터")
st.dataframe(period_df, use_container_width=True, hide_index=True)

with st.expander("24시간 상세 결과", expanded=False):
    st.dataframe(hourly_df, use_container_width=True, hide_index=True)

with st.expander("탭별 비교표", expanded=False):
    tap_table_df = tap_compare_df.copy()
    if not tap_table_df.empty:
        tap_table_df["탭"] = pd.to_numeric(tap_table_df["탭"], errors="coerce")
        tap_table_df = tap_table_df.sort_values(by="탭", ascending=False).reset_index(drop=True)
    st.dataframe(tap_table_df, use_container_width=True, hide_index=True)


# =========================================================
# 결과 저장
# =========================================================
st.divider()
st.subheader("결과 저장")

excel_output = io.BytesIO()
with pd.ExcelWriter(excel_output, engine="openpyxl") as writer:
    period_df.to_excel(writer, index=False, sheet_name="시간대구간결과")
    hourly_df.to_excel(writer, index=False, sheet_name="24시간상세")
    tap_compare_df.to_excel(writer, index=False, sheet_name="탭별비교")
    rate_summary_df.to_excel(writer, index=False, sheet_name="적용요금표")
excel_output.seek(0)

pdf_error = None
pdf_bytes = None

try:
    pdf_font_info = get_korean_font_info()
    fig_line_buf = create_matplotlib_line_chart(hourly_df, season, font_info=pdf_font_info)
    fig_bar_buf = create_matplotlib_bar_chart(tap_compare_df, site_name, font_info=pdf_font_info)

    pdf_bytes = build_pdf_bytes_report(
        site_name=site_name,
        season=season,
        created_at_text=created_at_text,
        calc_voltage_basis=calc_voltage_basis,
        primary_voltage_kv=primary_voltage_kv,
        secondary_voltage_v=secondary_voltage_v,
        new_voltage=new_voltage,
        current_voltage_for_calc=current_voltage_for_calc,
        voltage_drop_pct=voltage_drop_pct,
        voltage_drop_v=voltage_drop_v,
        primary_voltage_class=primary_voltage_class,
        tariff_voltage_class=tariff_voltage_class,
        contract_kind=contract_kind_display,
        tariff_choice_desc=tariff_choice_desc,
        cvrf_value=defaults["cvrf"],
        z=z,
        i=i,
        p=p,
        correction_factor=correction_factor,
        reliability_text=reliability_text,
        current_tap=current_tap,
        target_tap=target_tap,
        tap_step_percent=tap_step_percent,
        avg_saved_kw=avg_saved_kw,
        saving_rate_total=saving_rate_total,
        daily_saved_kwh=daily_saved_kwh,
        monthly_saved_kwh=monthly_saved_kwh,
        yearly_saved_kwh=yearly_saved_kwh,
        daily_cost_saving=daily_cost_saving,
        monthly_cost_saving=monthly_cost_saving,
        yearly_cost_saving=yearly_cost_saving,
        operation_mode=operation_mode,
        operating_start=operating_start,
        operating_end=operating_end,
        day_operation_hours=day_operation_hours,
        active_days_per_year=active_days_per_year,
        holiday_reflect=holiday_reflect,
        holiday_count=holiday_count,
        load_unit=load_unit,
        load_summary_df=load_summary_df,
        period_df=period_df,
        hourly_pdf_df=hourly_df,
        rate_summary_df=rate_summary_df,
        season_time_band_df=season_time_band_df,
        tap_compare_df=tap_compare_df,
        fig_line_buf=fig_line_buf,
        fig_bar_buf=fig_bar_buf,
    )
except Exception as e:
    pdf_error = str(e)

col_excel, col_pdf = st.columns(2)

with col_excel:
    st.download_button(
        "📊 엑셀 저장",
        data=excel_output.getvalue(),
        file_name="CVR결과_{0}_{1}.xlsx".format(site_name, datetime.now().strftime("%Y%m%d_%H%M%S")),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with col_pdf:
    if pdf_bytes is not None:
        st.download_button(
            "📄 PDF 저장",
            data=pdf_bytes,
            file_name="CVR보고서_{0}_{1}.pdf".format(site_name, datetime.now().strftime("%Y%m%d_%H%M%S")),
            mime="application/pdf",
        )
    else:
        st.button("📄 PDF 저장", disabled=True)
        st.error("PDF 생성 실패: {0}".format(pdf_error if pdf_error else "알 수 없는 오류"))
