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
from reportlab.graphics.shapes import Drawing, String
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
    "pp_hourly_profile_kw": {},
    "pp_daily_band_kwh": {},
    "pp_monthly_band_kwh": {},
    "pp_auto_source": "",
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
        # 파워플래너 시간대별 구분표 기준
        "경부하": list(range(23, 24)) + list(range(0, 8)),
        "중간부하": list(range(8, 11)) + list(range(13, 18)) + [22],
        "최대부하": list(range(11, 13)) + list(range(18, 22)),
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
    # 빠른 로딩: 이미지/폰트까지 모두 기다리지 않고 DOM이 뜨면 진행
    options.page_load_strategy = "eager"

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
    # 파워플래너 값 추출에는 이미지가 필요 없으므로 네트워크 로딩량 감소
    options.add_experimental_option("prefs", {
        "profile.managed_default_content_settings.images": 2,
        "profile.default_content_setting_values.notifications": 2,
    })

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

        driver.set_page_load_timeout(18)
        driver.implicitly_wait(0.5)
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


def create_matplotlib_line_chart(graph_df, font_info):
    with plt.rc_context({"axes.unicode_minus": False}):
        if font_info.get("mpl_name"):
            plt.rcParams["font.family"] = font_info["mpl_name"]

        fig, ax = plt.subplots(figsize=(10, 4.8))
        ax.plot(
            graph_df["시간번호"],
            graph_df["전력사용량(kW)"],
            marker="o",
            linewidth=1.8,
            color="#1f77b4",
            markerfacecolor="#1f77b4",
            markeredgecolor="#1f77b4",
        )
        ax.set_xticks(list(range(24)))
        ax.set_xticklabels([f"{i:02d}" for i in range(24)], fontsize=8)
        ax.grid(True, alpha=0.3)
        ax.set_title("")
        ax.set_xlabel("")
        ax.set_ylabel("")
        fig.subplots_adjust(left=0.08, bottom=0.08, right=0.98, top=0.98)

        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=180, bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        return buf




def create_matplotlib_bar_chart(tap_compare_df, font_info):
    with plt.rc_context({"axes.unicode_minus": False}):
        if font_info.get("mpl_name"):
            plt.rcParams["font.family"] = font_info["mpl_name"]

        plot_df = tap_compare_df.copy()
        plot_df["탭"] = pd.to_numeric(plot_df["탭"], errors="coerce")
        plot_df = plot_df.sort_values(by="탭", ascending=True).reset_index(drop=True)

        fig, ax = plt.subplots(figsize=(10, 4.8))
        x_labels = plot_df["탭"].astype(int).astype(str)
        y_vals = plot_df["평균 절감전력(kW)"]
        bars = ax.bar(
            x_labels,
            y_vals,
            color="#1f77b4",
            edgecolor="#1f77b4",
        )
        ax.grid(True, axis="y", alpha=0.3)
        ax.invert_xaxis()
        ax.set_title("")
        ax.set_xlabel("")
        ax.set_ylabel("")

        for rect, val in zip(bars, y_vals):
            label = "0" if float(val) == 0 else f"{float(val):.3f}"
            ax.text(
                rect.get_x() + rect.get_width() / 2,
                rect.get_height(),
                label,
                ha="center",
                va="bottom",
                fontsize=8,
                color="#333333",
            )

        fig.subplots_adjust(left=0.08, bottom=0.08, right=0.98, top=0.98)

        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=180, bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        return buf




def make_pdf_graph_block(image_buf, title_text, x_label_text, y_label_text, font_name, body_style):
    y_drawing = Drawing(22 * mm, 86 * mm)
    y_drawing.add(
        String(
            10 * mm,
            43 * mm,
            y_label_text,
            fontName=font_name,
            fontSize=8,
            fillColor=colors.black,
            textAnchor="middle",
            angle=90,
        )
    )

    graph_table = Table(
        [[y_drawing, RLImage(image_buf, width=170 * mm, height=80 * mm)]],
        colWidths=[22 * mm, 170 * mm],
    )
    graph_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))

    x_style = ParagraphStyle(
        "GraphXAxisStyle",
        parent=body_style,
        alignment=TA_CENTER,
        fontName=font_name,
        fontSize=9,
        leading=11,
        spaceBefore=3,
        spaceAfter=0,
    )

    return [
        Paragraph(title_text, body_style),
        Spacer(1, 4),
        graph_table,
        Paragraph(x_label_text, x_style),
    ]

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
# 파워플래너 화면 해석 보조 함수
# =========================================================
def normalize_space(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


def normalize_key(text):
    return re.sub(r"\s+", "", str(text or "")).strip()


def safe_read_html_tables(html, logs=None, label=""):
    try:
        return pd.read_html(io.StringIO(html))
    except Exception as e:
        if logs is not None and label:
            add_log(logs, f"{label} 표 해석 실패: {e}")
        return []


def table_signature(df):
    try:
        parts = [" ".join([normalize_space(str(c)) for c in df.columns.tolist()])]
        preview_rows = min(len(df), 8)
        for i in range(preview_rows):
            row = df.iloc[i].tolist()
            parts.append(" ".join([normalize_space(str(v)) for v in row]))
        return normalize_space(" ".join(parts))
    except Exception:
        return ""


def find_table_by_keywords(tables, keywords):
    normalized_keywords = [normalize_space(k) for k in keywords if k]
    for df in tables:
        signature = table_signature(df)
        if all(k in signature for k in normalized_keywords):
            return df.copy()
    return None


def flatten_table_pairs(df):
    pairs = {}
    if df is None or df.empty:
        return pairs

    try:
        working = df.copy().fillna("")
        working.columns = [normalize_space(c) for c in working.columns]
        values = working.astype(str).values.tolist()
        for row in values:
            row = [normalize_space(v) for v in row]
            if not any(row):
                continue
            for i in range(0, len(row) - 1, 2):
                key = row[i]
                val = row[i + 1]
                if not key or not val:
                    continue
                if normalize_key(key) == normalize_key(val):
                    continue
                pairs[key] = val
    except Exception:
        return pairs

    return pairs


def extract_pairs_from_tables(tables):
    pairs = {}
    for df in tables:
        pairs.update(flatten_table_pairs(df))
    return pairs


def pick_pair_value(pairs, patterns):
    if not pairs:
        return ""
    items = list(pairs.items())
    for pattern in patterns:
        if isinstance(pattern, (list, tuple)):
            need = [normalize_key(x) for x in pattern]
            for key, value in items:
                nkey = normalize_key(key)
                if all(x in nkey for x in need):
                    return str(value)
        else:
            need = normalize_key(pattern)
            for key, value in items:
                if need and need in normalize_key(key):
                    return str(value)
    return ""


def extract_first_regex(text, patterns):
    raw = str(text or "")
    for pattern in patterns:
        m = re.search(pattern, raw, re.IGNORECASE | re.MULTILINE | re.DOTALL)
        if m:
            return normalize_space(m.group(1))
    return ""


def month_to_season(month):
    if month in [6, 7, 8]:
        return "여름"
    if month in [11, 12, 1, 2]:
        return "겨울"
    return "봄·가을"



def aggregate_hourly_profile_kw(hourly_map):
    """
    hourly_map:
      - 15분 자료면 key가 0,0.25,0.5... 형식이고 value는 해당 구간 kWh
      - 시간 자료면 key가 0~23 형식이고 value는 해당 시간 kWh(=평균 kW로 간주 가능)

    return:
      {0: hour_avg_kw, ..., 23: hour_avg_kw}
    """
    if not hourly_map:
        return {}

    buckets = {}
    for raw_key, raw_val in hourly_map.items():
        try:
            t = float(raw_key)
            v = float(raw_val)
        except Exception:
            continue
        if v <= 0:
            continue

        hour = int(t) % 24
        frac = round(abs(t - int(t)), 2)

        # 15분 자료는 kWh / 0.25h = 평균 kW
        if frac in (0.25, 0.5, 0.75):
            kw = v / 0.25
        else:
            kw = v

        buckets.setdefault(hour, []).append(kw)

    return {h: round(sum(vals) / len(vals), 3) for h, vals in sorted(buckets.items()) if vals}


def _summarize_by_actual_load_level(hourly_profile_kw):
    """
    시간표가 아니라 실제 24시간 부하 크기 기준으로 경/중/최대부하 평균을 나눈다.
    최대부하 시간대가 0.00으로 무너지는 경우를 방지하기 위한 보정 로직이다.
    """
    values = [float(v) for v in hourly_profile_kw.values() if float(v) > 0]
    if not values:
        return None

    base_avg_kw = sum(values) / len(values)
    max_kw = max(values)
    if base_avg_kw <= 0 or max_kw <= 0:
        return None

    peak_vals = [v for v in values if v >= max_kw * 0.85]
    mid_vals = [v for v in values if max_kw * 0.45 <= v < max_kw * 0.85]
    off_vals = [v for v in values if v < max_kw * 0.45]

    if not peak_vals or not mid_vals or not off_vals:
        ordered = sorted(values)
        n = len(ordered)
        if n >= 3:
            low_end = max(1, int(round(n * 0.33)))
            high_start = min(n - 1, int(round(n * 0.67)))
            off_vals = ordered[:low_end]
            mid_vals = ordered[low_end:high_start]
            peak_vals = ordered[high_start:]
        else:
            off_vals = ordered[:1]
            mid_vals = ordered
            peak_vals = ordered[-1:]

    def avg(vals, fallback):
        return sum(vals) / len(vals) if vals else fallback

    band_avg = {
        "경부하": avg(off_vals, base_avg_kw),
        "중간부하": avg(mid_vals, base_avg_kw),
        "최대부하": avg(peak_vals, max_kw),
    }

    return {
        "base_avg_kw": round(base_avg_kw, 3),
        "off_peak_kw": round(band_avg["경부하"], 3),
        "mid_peak_kw": round(band_avg["중간부하"], 3),
        "peak_kw": round(band_avg["최대부하"], 3),
        "off_ratio": round(band_avg["경부하"] / base_avg_kw, 4),
        "mid_ratio": round(band_avg["중간부하"] / base_avg_kw, 4),
        "peak_ratio": round(band_avg["최대부하"] / base_avg_kw, 4),
        "hourly_profile_kw": hourly_profile_kw,
        "source": "actual_load_level",
    }


def summarize_band_loads_from_hourly(hourly_map, season):
    """
    파워플래너 시간대별 사용량 값을 그대로 사용하되,
    경/중/최대부하는 파워플래너 계절별 시간대 구분표 기준으로만 나눈다.
    조회 당일 자료처럼 미래 시간이 0으로 채워진 경우가 많아서 계산용 평균부하/비율은 양수 데이터 기준으로 산출한다.
    그래프 표시용 hourly_profile_kw는 원본 24시간 형태를 유지한다.
    """
    hourly_profile_kw = aggregate_hourly_profile_kw(hourly_map)
    if not hourly_profile_kw:
        return None

    target_season = season if season in SEASON_SCHEDULE else month_to_season(datetime.now().month)
    valid_items = {}
    for h in range(24):
        try:
            kw = float(hourly_profile_kw.get(h, 0.0))
        except Exception:
            kw = 0.0
        if kw > 0:
            valid_items[h] = kw

    if not valid_items:
        return None

    base_avg_kw = sum(valid_items.values()) / len(valid_items)
    if base_avg_kw <= 0:
        return None

    band_buckets = {"경부하": [], "중간부하": [], "최대부하": []}
    band_energy = {"경부하": 0.0, "중간부하": 0.0, "최대부하": 0.0}

    for hour, kw in sorted(valid_items.items()):
        band = hour_to_label(hour, target_season)
        band_buckets[band].append(kw)
        band_energy[band] += kw

    band_avg = {}
    for label in ["경부하", "중간부하", "최대부하"]:
        vals = band_buckets[label]
        band_avg[label] = (sum(vals) / len(vals)) if vals else 0.0

    total_energy = sum(band_energy.values())
    energy_ratio = {k: (band_energy[k] / total_energy if total_energy else 0.0) for k in band_energy}

    return {
        "base_avg_kw": round(base_avg_kw, 3),
        "off_peak_kw": round(band_avg["경부하"], 3),
        "mid_peak_kw": round(band_avg["중간부하"], 3),
        "peak_kw": round(band_avg["최대부하"], 3),
        "off_ratio": round(band_avg["경부하"] / base_avg_kw, 4) if base_avg_kw else 0.0,
        "mid_ratio": round(band_avg["중간부하"] / base_avg_kw, 4) if base_avg_kw else 0.0,
        "peak_ratio": round(band_avg["최대부하"] / base_avg_kw, 4) if base_avg_kw else 0.0,
        "energy_off_ratio": round(energy_ratio["경부하"], 4),
        "energy_mid_ratio": round(energy_ratio["중간부하"], 4),
        "energy_peak_ratio": round(energy_ratio["최대부하"], 4),
        "hourly_profile_kw": hourly_profile_kw,
        "valid_hour_count": len(valid_items),
        "source": "powerplanner_tou_hourly_positive",
    }

def select_15min_view_if_available(driver, by, logs, label):
    try:
        radio_selectors = [
            (by.XPATH, "//label[contains(normalize-space(.),'15분')]"),
            (by.XPATH, "//span[contains(normalize-space(.),'15분')]"),
            (by.XPATH, "//input[@type='radio' and contains(@onclick,'15') or @value='15']"),
        ]
        clicked = click_first(driver, radio_selectors, logs, f'{label} 15분 선택')
        if clicked:
            time.sleep(0.25)
            search_selectors = [
                (by.XPATH, "//button[contains(normalize-space(.),'조회') or contains(normalize-space(.),'검색') ]"),
                (by.XPATH, "//a[contains(normalize-space(.),'조회') or contains(normalize-space(.),'검색') ]"),
                (by.CSS_SELECTOR, "button"),
            ]
            click_first(driver, search_selectors, logs, f'{label} 조회 버튼')
            time.sleep(0.5)
    except Exception as e:
        add_log(logs, f'{label} 15분 기준 설정 실패(무시): {e}')


def extract_tariff_rates_from_tables(tables, season=None):
    if not season:
        season = month_to_season(datetime.now().month)

    season_alias = {
        "여름": ["여름", "여름철"],
        "봄·가을": ["봄·가을", "봄/가을", "봄가을", "봄·가을철"],
        "겨울": ["겨울", "겨울철"],
    }

    target_df = None
    for df in tables:
        sig = table_signature(df)
        if all(x in sig for x in ["경부하", "중간부하", "최대부하"]):
            target_df = df.copy()
            break
    if target_df is None:
        return {}

    target_df = target_df.copy().fillna("")
    target_df.columns = [normalize_space(c) for c in target_df.columns]

    basic_charge = 0.0
    for col in target_df.columns:
        if "기본요금" in col:
            for v in target_df[col].tolist():
                num = parse_number(v)
                if num > 0:
                    basic_charge = num
                    break

    season_col = None
    for alias in season_alias.get(season, []):
        for col in target_df.columns:
            if alias in normalize_space(col):
                season_col = col
                break
        if season_col:
            break

    if season_col is None:
        for col in target_df.columns:
            ncol = normalize_space(col)
            if "봄" in ncol and "가을" in ncol:
                season_col = col
                break

    time_col = None
    for col in target_df.columns:
        if "시간대" in normalize_space(col):
            time_col = col
            break
    if time_col is None:
        time_col = target_df.columns[1] if len(target_df.columns) > 1 else target_df.columns[0]

    rates = {"basic_charge_unit": basic_charge, "off_peak_rate": 0.0, "mid_peak_rate": 0.0, "peak_rate": 0.0}
    if season_col is None:
        return rates

    for _, row in target_df.iterrows():
        time_name = normalize_space(row.get(time_col, ""))
        rate = parse_number(row.get(season_col, 0))
        if "경부하" in time_name:
            rates["off_peak_rate"] = rate
        elif "중간부하" in time_name:
            rates["mid_peak_rate"] = rate
        elif "최대부하" in time_name:
            rates["peak_rate"] = rate

    return rates


def parse_hour_value(hour_text):
    m = re.search(r"(\d{1,2})", str(hour_text or ""))
    if not m:
        return None
    hour = int(m.group(1))
    if hour == 24:
        return 0
    if 0 <= hour <= 23:
        return hour
    return None


def extract_hourly_usage_map_from_tables(tables):
    hourly = {}
    for df in tables:
        sig = table_signature(df)
        if "시" not in sig:
            continue

        work = df.copy().fillna("")
        try:
            if isinstance(work.columns, pd.MultiIndex):
                work.columns = [normalize_space(" ".join([str(x) for x in col if str(x) != "nan"])) for col in work.columns]
            else:
                work.columns = [normalize_space(c) for c in work.columns]
        except Exception:
            work.columns = [normalize_space(c) for c in work.columns]

        for _, row in work.iterrows():
            values = [normalize_space(v) for v in row.tolist()]
            for idx, cell in enumerate(values):
                hour = parse_hour_value(cell)
                if hour is None:
                    continue
                for j in range(idx + 1, min(idx + 4, len(values))):
                    num = parse_number(values[j])
                    if num > 0:
                        hourly[hour] = max(hourly.get(hour, 0.0), num)
                        break
    return hourly


def extract_hourly_usage_map_from_text(text):
    hourly = {}
    raw = str(text or "")
    for line in raw.splitlines():
        line = normalize_space(line)
        if not line:
            continue
        pairs = re.findall(r"(?:(\d{1,2})[:시]?(?:00)?)\s+([0-9,]{2,}(?:\.\d+)?)", line)
        for hour_txt, usage_txt in pairs:
            hour = parse_hour_value(hour_txt)
            usage = parse_number(usage_txt)
            if hour is None or usage <= 0:
                continue
            hourly[hour] = max(hourly.get(hour, 0.0), usage)
    return hourly


def extract_pattern_hourly_map_from_text(text):
    weekday = {}
    holiday = {}
    raw = str(text or "")
    for line in raw.splitlines():
        line = normalize_space(line)
        m = re.match(r"^(\d{1,2})[:시]00\s+([0-9,]{2,}(?:\.\d+)?)\s+[0-9,\.]+\s+[0-9,\.]+\s+([0-9,]{2,}(?:\.\d+)?)", line)
        if m:
            hour = parse_hour_value(m.group(1))
            if hour is None:
                continue
            weekday[hour] = parse_number(m.group(2))
            holiday[hour] = parse_number(m.group(3))
            continue
        m2 = re.match(r"^(\d{1,2})\s+([0-9,]{2,}(?:\.\d+)?)\s+[0-9,\.]+\s+[0-9,\.]+\s+([0-9,]{2,}(?:\.\d+)?)", line)
        if m2:
            hour = parse_hour_value(m2.group(1))
            if hour is None:
                continue
            weekday[hour] = parse_number(m2.group(2))
            holiday[hour] = parse_number(m2.group(3))
    mixed = {}
    for h in sorted(set(weekday) | set(holiday)):
        wd = weekday.get(h, 0.0)
        hd = holiday.get(h, 0.0)
        if wd > 0 and hd > 0:
            mixed[h] = wd * 5.0 / 7.0 + hd * 2.0 / 7.0
        elif wd > 0:
            mixed[h] = wd
        elif hd > 0:
            mixed[h] = hd
    return mixed


def extract_latest_12_months_usage_from_tables(tables):
    monthly_rows = []

    for df in tables:
        sig = table_signature(df)
        if "당해(kWh)" in sig and "월" in sig:
            work = df.copy().fillna("")
            work.columns = [normalize_space(c) for c in work.columns]
            month_col = None
            usage_col = None
            for col in work.columns:
                ncol = normalize_space(col)
                if ncol == "월" or "월" in ncol:
                    month_col = col
                if "당해" in ncol and "kWh" in ncol:
                    usage_col = col
            if month_col and usage_col:
                for _, row in work.iterrows():
                    month_text = normalize_space(row.get(month_col, ""))
                    m = re.search(r"(\d{1,2})월", month_text)
                    if not m:
                        continue
                    month = int(m.group(1))
                    usage = parse_number(row.get(usage_col, 0))
                    if usage > 1000:
                        monthly_rows.append((month, usage))
                if monthly_rows:
                    monthly_rows.sort(key=lambda x: x[0], reverse=True)
                    return sum(v for _, v in monthly_rows[:12])

    for df in tables:
        sig = table_signature(df)
        if "사용량합계" in sig and "kWh" in sig:
            work = df.copy().fillna("")
            work.columns = [normalize_space(c) for c in work.columns]
            for col in work.columns:
                if "사용량합계" in normalize_space(col):
                    for v in work[col].tolist():
                        num = parse_number(v)
                        if num > 1000:
                            return num
    return 0.0


def extract_latest_12_months_usage_from_text(text):
    raw = str(text or "")
    monthly = []
    for m in re.finditer(r"(\d{1,2})월(?:\([^\)]*\))?\s+([0-9,]+(?:\.\d+)?)", raw):
        month = int(m.group(1))
        usage = parse_number(m.group(2))
        if usage > 1000:
            monthly.append((month, usage))
    if monthly:
        dedup = {}
        for month, usage in monthly:
            dedup[month] = max(dedup.get(month, 0.0), usage)
        rows = sorted(dedup.items(), key=lambda x: x[0], reverse=True)
        return sum(v for _, v in rows[:12])

    m = re.search(r"사용량합계\s*(?:\(kWh\))?\s*([0-9,]+(?:\.\d+)?)", raw)
    if m:
        return parse_number(m.group(1))
    return 0.0


def extract_yearly_usage_from_tables(tables):
    rows = []
    for df in tables:
        sig = table_signature(df)
        if "연도" in sig and "사용량(kWh)" in sig:
            work = df.copy().fillna("")
            work.columns = [normalize_space(c) for c in work.columns]
            year_col = None
            usage_col = None
            for col in work.columns:
                ncol = normalize_space(col)
                if "연도" in ncol:
                    year_col = col
                if "사용량" in ncol and "kWh" in ncol:
                    usage_col = col
            if year_col and usage_col:
                for _, row in work.iterrows():
                    year_text = normalize_space(row.get(year_col, ""))
                    m = re.search(r"(20\d{2})", year_text)
                    if not m:
                        continue
                    year = int(m.group(1))
                    usage = parse_number(row.get(usage_col, 0))
                    if usage > 0:
                        rows.append((year, usage))
    if rows:
        now_year = datetime.now().year
        full_years = [(y, v) for y, v in rows if y < now_year]
        target = sorted(full_years if full_years else rows, key=lambda x: x[0], reverse=True)[0]
        return target[1]
    return 0.0


def extract_yearly_usage_from_text(text):
    rows = []
    raw = str(text or "")
    for m in re.finditer(r"(20\d{2})년\s+([0-9,]+(?:\.\d+)?)", raw):
        year = int(m.group(1))
        usage = parse_number(m.group(2))
        if usage > 0:
            rows.append((year, usage))
    if rows:
        now_year = datetime.now().year
        full_years = [(y, v) for y, v in rows if y < now_year]
        target = sorted(full_years if full_years else rows, key=lambda x: x[0], reverse=True)[0]
        return target[1]
    return 0.0


def extract_max_demand_from_tables(tables):
    candidates = []
    for df in tables:
        sig = table_signature(df)
        if "최대수요" not in sig:
            continue
        work = df.copy().fillna("")
        work.columns = [normalize_space(c) for c in work.columns]
        md_cols = [col for col in work.columns if "최대수요" in normalize_space(col) and "kW" in normalize_space(col)]
        for md_col in md_cols:
            for v in work[md_col].tolist():
                num = parse_number(v)
                if num > 0:
                    candidates.append(num)
    return max(candidates) if candidates else 0.0


def extract_max_demand_from_text(text):
    raw = str(text or "")
    patterns = [
        r"최대수요전력\s*[:：]?\s*([0-9,]+(?:\.\d+)?)\s*kW",
        r"최대수요\s*\(kW\)\s*([0-9,]+(?:\.\d+)?)",
        r"최대\s*\(kW\)\s*([0-9,]+(?:\.\d+)?)",
        r"최대수요\s*([0-9,]+(?:\.\d+)?)\s*kW",
    ]
    values = []
    for pattern in patterns:
        for m in re.finditer(pattern, raw, re.IGNORECASE):
            num = parse_number(m.group(1))
            if num > 0:
                values.append(num)
    return max(values) if values else 0.0


def fetch_and_parse_powerplanner_page(driver, wait, by, page_meta, result, logs):
    page = visit_powerplanner_page(
        driver=driver,
        wait=wait,
        by=by,
        url=page_meta["url"],
        logs=logs,
        label=page_meta["label"],
        ready_patterns=page_meta.get("ready"),
        post_load=page_meta.get("post_load"),
    )
    parser = page_meta.get("parser")
    if parser is not None:
        parser(page, result, logs)


def visit_powerplanner_page(driver, wait, by, url, logs, label, ready_patterns=None, post_load=None):
    driver.get(url)
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    wait.until(lambda d: len(d.find_elements(by.TAG_NAME, "body")) > 0)
    time.sleep(0.35)

    if callable(post_load):
        try:
            post_load(driver, by, logs, label)
            wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
            time.sleep(0.3)
        except Exception as e:
            add_log(logs, f"{label} 후처리 실패(무시): {e}")

    if ready_patterns:
        body_text = driver.find_element(by.TAG_NAME, "body").text
        for pattern in ready_patterns:
            if pattern and pattern not in body_text:
                add_log(logs, f"{label} 화면 로드 후 확인문구 미일치: {pattern}")
                break

    html = driver.page_source
    text = driver.find_element(by.TAG_NAME, "body").text
    add_log(logs, f"{label} 화면 접속 완료")
    return {"url": url, "html": html, "text": text}


def scrape_smartview_page(page, result, logs):
    text = page["text"]
    tables = safe_read_html_tables(page["html"], logs=logs, label="스마트뷰")
    pairs = extract_pairs_from_tables(tables)

    contract_kind = pick_pair_value(pairs, ["적용전기요금", "계약종별"])
    if not contract_kind:
        contract_kind = extract_first_regex(
            text,
            [r"적용전기요금\s*([^\n]+)", r"적용전기요금\s*[:：]?\s*([^\n]+)"],
        )
    if contract_kind:
        result["contract_kind"] = contract_kind

    basic_charge = pick_pair_value(pairs, ["기본요금단가"])
    if not basic_charge:
        basic_charge = extract_first_regex(text, [r"기본요금단가\s*[:：]?\s*([0-9,\.]+)\s*원"])
    if basic_charge:
        result["basic_charge_unit"] = max(result["basic_charge_unit"], parse_number(basic_charge))

    power_bill = pick_pair_value(pairs, ["요금적용전력"])
    if not power_bill:
        power_bill = extract_first_regex(text, [r"요금적용전력\s*[:：]?\s*([0-9,\.]+)\s*kW"])
    if power_bill:
        result["power_bill_kw"] = max(result["power_bill_kw"], parse_number(power_bill))

    max_demand = pick_pair_value(pairs, ["최대수요전력"])
    if not max_demand:
        max_demand = extract_first_regex(text, [r"최대수요전력\s*[:：]?\s*([0-9,\.]+)\s*kW"])
    if max_demand:
        result["max_demand_kw"] = max(result["max_demand_kw"], parse_number(max_demand))

    realtime_usage = extract_first_regex(
        text,
        [r"실시간사용량\s*([0-9,\.]+)\s*kWh", r"실시간 사용량\s*([0-9,\.]+)\s*kWh"],
    )
    if realtime_usage:
        result["realtime_usage_kwh"] = parse_number(realtime_usage)

    realtime_fee = extract_first_regex(
        text,
        [r"실시간요금\s*([0-9,\.]+)\s*원", r"실시간 요금은\s*([0-9,\.]+)\s*원"],
    )
    if realtime_fee:
        result["yearly_bill_won"] = max(result.get("yearly_bill_won", 0.0), parse_number(realtime_fee))

    current_season = month_to_season(datetime.now().month)
    rates = extract_tariff_rates_from_tables(tables, season=current_season)
    if rates:
        if rates.get("basic_charge_unit", 0) > 0 and result["basic_charge_unit"] <= 0:
            result["basic_charge_unit"] = rates["basic_charge_unit"]
        result["off_peak_rate"] = max(result["off_peak_rate"], rates.get("off_peak_rate", 0.0))
        result["mid_peak_rate"] = max(result["mid_peak_rate"], rates.get("mid_peak_rate", 0.0))
        result["peak_rate"] = max(result["peak_rate"], rates.get("peak_rate", 0.0))

    add_log(logs, "스마트뷰 요약값 해석 완료")


def scrape_customer_info_page(page, result, logs):
    text = page["text"]
    tables = safe_read_html_tables(page["html"], logs=logs, label="고객정보")
    pairs = extract_pairs_from_tables(tables)

    contract_kind = pick_pair_value(pairs, ["계약종별"])
    if contract_kind:
        result["contract_kind"] = contract_kind

    contract_power = pick_pair_value(pairs, ["계약전력"])
    if contract_power:
        result["contract_power_kw"] = parse_number(contract_power)

    supply_text = pick_pair_value(pairs, ["공급방식"])
    if supply_text:
        result["supply_voltage_text"] = supply_text

    meter_read_text = pick_pair_value(pairs, ["검침일"])
    if meter_read_text:
        m = re.search(r"(\d{1,2})", meter_read_text)
        if m:
            result["meter_read_day"] = int(m.group(1))

    if not result["contract_kind"]:
        result["contract_kind"] = extract_first_regex(text, [r"계약종별\s*([^\n]+)"])
    if result["contract_power_kw"] <= 0:
        result["contract_power_kw"] = parse_number(extract_first_regex(text, [r"계약전력\s*([0-9,\.]+\s*kw)"]))
    if not result["supply_voltage_text"]:
        result["supply_voltage_text"] = extract_first_regex(text, [r"공급방식\s*([^\n]+)"])

    voltage_class = classify_voltage_from_contract_kind(result.get("contract_kind", ""), result.get("supply_voltage_text", ""))
    result["voltage_class"] = voltage_class or result.get("voltage_class", "")
    result["primary_voltage_kv"] = primary_voltage_from_class(result["voltage_class"], result.get("supply_voltage_text", ""))
    add_log(logs, "고객정보 해석 완료")



def scrape_hourly_usage_page(page, result, logs):
    tables = safe_read_html_tables(page["html"], logs=logs, label="시간대별 사용량")
    hourly_map = extract_hourly_usage_map_from_tables(tables)
    if len(hourly_map) < 6:
        hourly_map = extract_hourly_usage_map_from_text(page["text"])
    if hourly_map:
        band_summary = summarize_band_loads_from_hourly(hourly_map, month_to_season(datetime.now().month))
        if not band_summary:
            band_summary = summarize_band_loads_from_hourly(hourly_map, "봄·가을")
        if band_summary:
            result["auto_avg_base_kw"] = band_summary["base_avg_kw"]
            result["auto_off_peak_kw"] = band_summary["off_peak_kw"]
            result["auto_mid_peak_kw"] = band_summary["mid_peak_kw"]
            result["auto_peak_kw"] = band_summary["peak_kw"]
            result["auto_off_ratio"] = band_summary["off_ratio"]
            result["auto_mid_ratio"] = band_summary["mid_ratio"]
            result["auto_peak_ratio"] = band_summary["peak_ratio"]
            result["hourly_profile_kw"] = band_summary.get("hourly_profile_kw", {})
            result["auto_source"] = "usage_hourly"
            add_log(logs, f"시간대별 사용량 기준 평균부하 자동 산출 성공: 평균 {band_summary['base_avg_kw']:.1f} kW")

    md = max(extract_max_demand_from_tables(tables), extract_max_demand_from_text(page["text"]))
    if md > 0:
        result["max_demand_kw"] = max(result["max_demand_kw"], md)


def flatten_columns_for_powerplanner(df):
    work = df.copy().fillna(0)
    try:
        if isinstance(work.columns, pd.MultiIndex):
            cols = []
            for col in work.columns:
                parts = [normalize_space(str(x)) for x in col if normalize_space(str(x)) and str(x).lower() != "nan"]
                cols.append(" ".join(parts))
            work.columns = cols
        else:
            work.columns = [normalize_space(c) for c in work.columns]
    except Exception:
        work.columns = [normalize_space(c) for c in work.columns]
    return work


def extract_daily_timeband_energy_summary(tables, season=None):
    """
    파워플래너 일별요금 표의 전력사용량(kWh) 경/중/최대부하 컬럼을 직접 합산한다.
    시간대별 순간 패턴보다 파워플래너 요금표에 쓰이는 구간별 kWh가 우선이다.
    """
    target_season = season if season in SEASON_SCHEDULE else month_to_season(datetime.now().month)
    schedule = SEASON_SCHEDULE[target_season]
    band_hours = {label: max(len(hours), 1) for label, hours in schedule.items()}

    best = None
    for df in tables:
        sig = table_signature(df)
        if not ("전력사용량" in sig and "중간부하" in sig and "경부하" in sig and "최대부하" in sig):
            continue

        work = flatten_columns_for_powerplanner(df)
        cols = list(work.columns)
        usage_cols = {}

        # MultiIndex가 제대로 유지된 경우: "전력사용량(kWh) 중간부하" 형태를 우선 사용
        for band in ["경부하", "중간부하", "최대부하"]:
            candidates = [c for c in cols if "전력사용량" in c and band in c]
            if candidates:
                usage_cols[band] = candidates[0]

        # read_html이 헤더를 단순화해서 "중간부하/최대부하/경부하"가 반복된 경우 위치 기반 보정
        if len(usage_cols) < 3:
            band_only = [i for i, c in enumerate(cols) if any(b in c for b in ["경부하", "중간부하", "최대부하"])]
            # 일별요금 표는 보통 [최대수요, 전력사용량 3개, 전력량요금 4개] 순서다.
            # 첫 번째로 등장하는 경/중/최대부하 묶음을 전력사용량(kWh)로 본다.
            first_group = []
            for i in band_only:
                c = cols[i]
                if any(b in c for b in ["경부하", "중간부하", "최대부하"]):
                    first_group.append((i, c))
                if len(first_group) >= 3:
                    break
            for i, c in first_group:
                for band in ["경부하", "중간부하", "최대부하"]:
                    if band in c and band not in usage_cols:
                        usage_cols[band] = c

        if len(usage_cols) < 3:
            continue

        date_col = None
        for c in cols:
            if "일자" in c or normalize_key(c) in ["일", "날짜"]:
                date_col = c
                break

        totals = {"경부하": 0.0, "중간부하": 0.0, "최대부하": 0.0}
        valid_days = 0
        rows_used = 0

        for _, row in work.iterrows():
            if date_col:
                date_text = normalize_space(row.get(date_col, ""))
                # 04.24처럼 조회 당일 이후 0행도 있으므로 총합 0인 행은 제외한다.
                if date_text and not re.search(r"\d", date_text):
                    continue

            vals = {band: parse_number(row.get(col, 0)) for band, col in usage_cols.items()}
            row_total = sum(vals.values())
            if row_total <= 0:
                continue

            for band in totals:
                totals[band] += vals[band]
            valid_days += 1
            rows_used += 1

        if rows_used <= 0 or sum(totals.values()) <= 0:
            continue

        # 일별요금 월 조회 표이므로 유효 일수로 나눠 1일 평균 kWh → 시간대별 평균 kW로 변환
        band_avg_kw = {}
        for band in ["경부하", "중간부하", "최대부하"]:
            daily_kwh = totals[band] / valid_days
            band_avg_kw[band] = daily_kwh / band_hours[band]

        base_avg_kw = (sum(totals.values()) / valid_days) / 24.0
        if base_avg_kw <= 0:
            continue

        summary = {
            "base_avg_kw": round(base_avg_kw, 3),
            "off_peak_kw": round(band_avg_kw["경부하"], 3),
            "mid_peak_kw": round(band_avg_kw["중간부하"], 3),
            "peak_kw": round(band_avg_kw["최대부하"], 3),
            "off_ratio": round(band_avg_kw["경부하"] / base_avg_kw, 4),
            "mid_ratio": round(band_avg_kw["중간부하"] / base_avg_kw, 4),
            "peak_ratio": round(band_avg_kw["최대부하"] / base_avg_kw, 4),
            "daily_band_kwh": {band: round(totals[band] / valid_days, 3) for band in totals},
            "monthly_band_kwh": {band: round(totals[band], 3) for band in totals},
            "valid_days": valid_days,
            "source": "powerplanner_daily_bill_timeband",
        }
        if best is None or valid_days > best.get("valid_days", 0):
            best = summary

    return best


def scrape_daily_usage_page(page, result, logs):
    tables = safe_read_html_tables(page["html"], logs=logs, label="일별 사용량")

    # 파워플래너 일별요금 표의 경/중/최대부하 kWh를 최우선으로 반영한다.
    # 이 값은 파워플래너 요금계산에 직접 쓰이는 시간대 구분값이므로, CVR 계산에서도 가장 신뢰도가 높다.
    daily_band_summary = extract_daily_timeband_energy_summary(tables, season=month_to_season(datetime.now().month))
    if daily_band_summary:
        # 정확도 우선: 파워플래너 요금표에 실제로 쓰이는 일별요금 경/중/최대 kWh를 항상 최우선 반영한다.
        # 시간대별 그래프는 화면 표시용 보조자료로만 두고, CVR 계산 기준은 이 값으로 고정한다.
        result["auto_avg_base_kw"] = daily_band_summary["base_avg_kw"]
        result["auto_off_peak_kw"] = daily_band_summary["off_peak_kw"]
        result["auto_mid_peak_kw"] = daily_band_summary["mid_peak_kw"]
        result["auto_peak_kw"] = daily_band_summary["peak_kw"]
        result["auto_off_ratio"] = daily_band_summary["off_ratio"]
        result["auto_mid_ratio"] = daily_band_summary["mid_ratio"]
        result["auto_peak_ratio"] = daily_band_summary["peak_ratio"]
        result["daily_band_kwh"] = daily_band_summary.get("daily_band_kwh", {})
        result["monthly_band_kwh"] = daily_band_summary.get("monthly_band_kwh", {})
        result["auto_source"] = "daily_bill_timeband"
        result["hourly_profile_kw"] = {}
        add_log(
            logs,
            f"정확도 우선 적용: 일별요금 경/중/최대 kWh 기준으로 평균부하를 최종 반영: "
            f"평균 {daily_band_summary['base_avg_kw']:,.1f} kW / "
            f"경 {daily_band_summary['off_peak_kw']:,.1f} / "
            f"중 {daily_band_summary['mid_peak_kw']:,.1f} / "
            f"최대 {daily_band_summary['peak_kw']:,.1f} kW "
            f"({daily_band_summary['valid_days']}일 기준)"
        )


    md = max(extract_max_demand_from_tables(tables), extract_max_demand_from_text(page["text"]))
    if md > 0:
        result["max_demand_kw"] = max(result["max_demand_kw"], md)
        add_log(logs, f"일별 사용량에서 최대수요 후보 반영: {md:,.1f} kW")


def scrape_monthly_usage_page(page, result, logs):
    tables = safe_read_html_tables(page["html"], logs=logs, label="월별 사용량")
    annual_usage = extract_latest_12_months_usage_from_tables(tables)
    if annual_usage <= 0:
        annual_usage = extract_latest_12_months_usage_from_text(page["text"])
    if annual_usage > 1000:
        result["annual_usage_kwh"] = max(result["annual_usage_kwh"], annual_usage)
        add_log(logs, f"월별 사용량에서 최근 합계 반영: {annual_usage:,.1f} kWh")

    md = max(extract_max_demand_from_tables(tables), extract_max_demand_from_text(page["text"]))
    if md > 0:
        result["max_demand_kw"] = max(result["max_demand_kw"], md)


def scrape_yearly_usage_page(page, result, logs):
    tables = safe_read_html_tables(page["html"], logs=logs, label="연도별 사용량")
    annual_usage = extract_yearly_usage_from_tables(tables)
    if annual_usage <= 0:
        annual_usage = extract_yearly_usage_from_text(page["text"])
    if annual_usage > 0 and result["annual_usage_kwh"] <= 0:
        result["annual_usage_kwh"] = annual_usage
        add_log(logs, f"연도별 사용량에서 최근 연간 사용량 반영: {annual_usage:,.1f} kWh")

    md = max(extract_max_demand_from_tables(tables), extract_max_demand_from_text(page["text"]))
    if md > 0:
        result["max_demand_kw"] = max(result["max_demand_kw"], md)




def scrape_pattern_hourly_page(page, result, logs):
    hourly_map = extract_pattern_hourly_map_from_text(page.get("text", ""))
    if not hourly_map:
        hourly_map = extract_hourly_usage_map_from_text(page.get("text", ""))
    if not hourly_map:
        tables = safe_read_html_tables(page["html"], logs=logs, label="시간대별 패턴")
        hourly_map = extract_hourly_usage_map_from_tables(tables)

    # 실제 시간대별 사용량에서 이미 자동 산출을 성공했다면 패턴 페이지는 보조 정보로만 둔다.
    if result.get("auto_source") == "usage_hourly" and result.get("auto_avg_base_kw", 0) > 0:
        add_log(logs, "시간대별 패턴은 보조 정보로만 확인하고, 기준 평균부하는 시간대별 사용량 값을 유지합니다.")
        return

    if hourly_map:
        season = month_to_season(datetime.now().month)
        band_summary = summarize_band_loads_from_hourly(hourly_map, season)
        if not band_summary:
            band_summary = summarize_band_loads_from_hourly(hourly_map, "봄·가을")
        if band_summary:
            result["auto_avg_base_kw"] = band_summary["base_avg_kw"]
            result["auto_off_peak_kw"] = band_summary["off_peak_kw"]
            result["auto_mid_peak_kw"] = band_summary["mid_peak_kw"]
            result["auto_peak_kw"] = band_summary["peak_kw"]
            result["auto_off_ratio"] = band_summary["off_ratio"]
            result["auto_mid_ratio"] = band_summary["mid_ratio"]
            result["auto_peak_ratio"] = band_summary["peak_ratio"]
            result["hourly_profile_kw"] = band_summary.get("hourly_profile_kw", {})
            result["auto_source"] = "pattern_hourly"
            add_log(logs, f"시간대별 패턴 기준 평균부하 자동 산출 성공: 평균 {band_summary['base_avg_kw']:.1f} kW")

def scrape_realtime_charge_page(page, result, logs):
    text = page["text"]
    tables = safe_read_html_tables(page["html"], logs=logs, label="실시간·예상요금")
    pairs = extract_pairs_from_tables(tables)

    contract_kind = pick_pair_value(pairs, ["적용전기요금"])
    if not contract_kind:
        contract_kind = extract_first_regex(text, [r"적용전기요금\s*[:：]?\s*([^\n]+)"])
    if contract_kind:
        result["contract_kind"] = contract_kind

    basic_charge = pick_pair_value(pairs, ["기본요금"])
    if not basic_charge:
        basic_charge = extract_first_regex(text, [r"기본요금\s*[:：]?\s*([0-9,\.]+)\s*원"])
    if basic_charge and result["basic_charge_unit"] <= 0:
        result["basic_charge_unit"] = parse_number(basic_charge)

    realtime_fee = extract_first_regex(text, [r"실시간 요금은\s*([0-9,\.]+)\s*원", r"실시간요금\s*([0-9,\.]+)\s*원"])
    if realtime_fee:
        result["yearly_bill_won"] = max(result["yearly_bill_won"], parse_number(realtime_fee))

    current_season = month_to_season(datetime.now().month)
    rates = extract_tariff_rates_from_tables(tables, season=current_season)
    if rates:
        result["off_peak_rate"] = max(result["off_peak_rate"], rates.get("off_peak_rate", 0.0))
        result["mid_peak_rate"] = max(result["mid_peak_rate"], rates.get("mid_peak_rate", 0.0))
        result["peak_rate"] = max(result["peak_rate"], rates.get("peak_rate", 0.0))
        if result["basic_charge_unit"] <= 0:
            result["basic_charge_unit"] = rates.get("basic_charge_unit", 0.0)

    add_log(logs, "실시간·예상요금 해석 완료")


def scrape_timeband_charge_page(page, result, logs):
    tables = safe_read_html_tables(page["html"], logs=logs, label="시간대별요금")
    current_season = month_to_season(datetime.now().month)
    rates = extract_tariff_rates_from_tables(tables, season=current_season)
    if rates:
        result["off_peak_rate"] = max(result["off_peak_rate"], rates.get("off_peak_rate", 0.0))
        result["mid_peak_rate"] = max(result["mid_peak_rate"], rates.get("mid_peak_rate", 0.0))
        result["peak_rate"] = max(result["peak_rate"], rates.get("peak_rate", 0.0))
        add_log(logs, "시간대별 요금 단가 해석 완료")


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
        wait = WebDriverWait(driver, 15)

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
            (By.XPATH, "//button[contains(., '로그인')]") ,
            (By.XPATH, "//a[contains(., '로그인')]") ,
            (By.XPATH, "//input[@value='로그인']") ,
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

        wait.until(lambda d: "Logout" in d.page_source or "스마트뷰" in d.page_source or "고객정보" in d.page_source)
        time.sleep(0.8)
        add_log(logs, "로그인 성공 및 메인 진입 확인")

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
            "meter_read_day": 0,
            "yearly_bill_won": 0.0,
            "voltage_class": "",
            "auto_avg_base_kw": 0.0,
            "auto_off_peak_kw": 0.0,
            "auto_mid_peak_kw": 0.0,
            "auto_peak_kw": 0.0,
            "auto_off_ratio": 0.0,
            "auto_mid_ratio": 0.0,
            "auto_peak_ratio": 0.0,
            "hourly_profile_kw": {},
            "daily_band_kwh": {},
            "monthly_band_kwh": {},
            "auto_source": "",
        }

        # 속도 최적화 핵심:
        # 1) 필수 화면만 먼저 접속
        # 2) 값이 부족할 때만 보조 화면을 추가 접속
        # 기존처럼 모든 화면을 무조건 순회하면 사업장/서버 상태에 따라 1~2분까지 늘어날 수 있음
        essential_pages = [
            {
                "label": "스마트뷰",
                "url": "https://pp.kepco.co.kr/rm/rm0101.do?menu_id=O010101",
                "ready": ["스마트뷰"],
                "parser": scrape_smartview_page,
            },
            {
                "label": "고객정보",
                "url": "https://pp.kepco.co.kr/mb/mb0101.do?menu_id=O010601",
                "ready": ["고객정보"],
                "parser": scrape_customer_info_page,
            },
            {
                "label": "시간대별 사용량",
                "url": "https://pp.kepco.co.kr/rs/rs0101N.do?menu_id=O010201",
                "ready": ["시간대별", "사용량"],
                "post_load": select_15min_view_if_available,
                "parser": scrape_hourly_usage_page,
            },
            {
                "label": "월별 사용량",
                "url": "https://pp.kepco.co.kr/rs/rs0103.do?menu_id=O010203",
                "ready": ["월별", "사용량"],
                "parser": scrape_monthly_usage_page,
            },
            {
                "label": "시간대별요금",
                "url": "https://pp.kepco.co.kr/re/re0102.do?menu_id=O010402",
                "ready": ["시간대별", "요금"],
                "parser": scrape_timeband_charge_page,
            },
        ]

        fallback_pages = [
            {
                "label": "시간대별 패턴",
                "url": "https://pp.kepco.co.kr/rp/rp0101.do?menu_id=O010301",
                "ready": ["시간대별", "패턴기간"],
                "post_load": select_15min_view_if_available,
                "parser": scrape_pattern_hourly_page,
                "need": lambda r: r.get("auto_avg_base_kw", 0.0) <= 0,
            },
            {
                "label": "연도별 사용량",
                "url": "https://pp.kepco.co.kr/rs/rs0104.do?menu_id=O010204",
                "ready": ["연도별", "사용량"],
                "parser": scrape_yearly_usage_page,
                "need": lambda r: r.get("annual_usage_kwh", 0.0) <= 0,
            },
            {
                "label": "일별 사용량",
                "url": "https://pp.kepco.co.kr/rs/rs0102.do?menu_id=O010202",
                "ready": ["일별", "사용량"],
                "parser": scrape_daily_usage_page,
                "need": lambda r: r.get("max_demand_kw", 0.0) <= 0,
            },
            {
                "label": "실시간·예상요금",
                "url": "https://pp.kepco.co.kr/pr/pr0101.do?menu_id=O010401",
                "ready": ["실시간", "예상"],
                "parser": scrape_realtime_charge_page,
                "need": lambda r: (
                    r.get("off_peak_rate", 0.0) <= 0
                    or r.get("mid_peak_rate", 0.0) <= 0
                    or r.get("peak_rate", 0.0) <= 0
                    or r.get("basic_charge_unit", 0.0) <= 0
                ),
            },
        ]

        for page_meta in essential_pages:
            try:
                fetch_and_parse_powerplanner_page(driver, wait, By, page_meta, result, logs)
            except Exception as page_error:
                add_log(logs, f"{page_meta['label']} 처리 실패: {page_error}")
                add_log(logs, traceback.format_exc())

        for page_meta in fallback_pages:
            try:
                if not page_meta.get("need", lambda r: True)(result):
                    add_log(logs, f"{page_meta['label']} 생략: 이미 필요한 값 확보")
                    continue
                fetch_and_parse_powerplanner_page(driver, wait, By, page_meta, result, logs)
            except Exception as page_error:
                add_log(logs, f"{page_meta['label']} 처리 실패: {page_error}")
                add_log(logs, traceback.format_exc())

        if not result["voltage_class"]:
            result["voltage_class"] = classify_voltage_from_contract_kind(result.get("contract_kind", ""), result.get("supply_voltage_text", ""))
        if result["primary_voltage_kv"] <= 0:
            result["primary_voltage_kv"] = primary_voltage_from_class(result["voltage_class"], result.get("supply_voltage_text", ""))
        if not result["supply_voltage_text"]:
            result["supply_voltage_text"] = describe_voltage_class(result["voltage_class"], result["primary_voltage_kv"])

        if result["annual_usage_kwh"] <= 0 and result["auto_avg_base_kw"] > 0:
            result["annual_usage_kwh"] = result["auto_avg_base_kw"] * 24 * 365
            add_log(logs, "연간 사용량이 없어 시간대 평균부하 기준으로 보정했습니다.")

        if result["contract_power_kw"] <= 0 and result["power_bill_kw"] > 0:
            result["contract_power_kw"] = result["power_bill_kw"]
        if result["power_bill_kw"] <= 0 and result["max_demand_kw"] > 0:
            result["power_bill_kw"] = result["max_demand_kw"]

        success_score = 0
        for key in [
            "contract_kind",
            "basic_charge_unit",
            "power_bill_kw",
            "max_demand_kw",
            "annual_usage_kwh",
            "off_peak_rate",
            "mid_peak_rate",
            "peak_rate",
            "contract_power_kw",
            "auto_avg_base_kw",
        ]:
            value = result.get(key)
            if isinstance(value, str) and value.strip():
                success_score += 1
            elif isinstance(value, (int, float)) and value > 0:
                success_score += 1

        if result.get("auto_avg_base_kw", 0.0) <= 0:
            add_log(logs, "평균부하 자동 산출 실패: 시간대별 사용량/패턴 표에서 사업장별 값을 확정하지 못했습니다.")
        if success_score < 5:
            return {
                "status": "error",
                "message": "파워플래너 화면은 열렸지만 필요한 값 추출이 충분하지 않습니다. 자동화 로그를 확인해 주세요.",
                "logs": "\n".join(logs),
                **result,
            }

        add_log(logs, f"최종 추출 성공 점수: {success_score}/10")
        return {
            "status": "success",
            "message": "파워플래너 화면 기준 값 추출을 완료했습니다.",
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



def find_first(driver, selectors):
    for by, selector in selectors:
        try:
            elems = driver.find_elements(by, selector)
            for elem in elems:
                try:
                    if elem.is_displayed():
                        return elem
                except Exception:
                    return elem
        except Exception:
            continue
    return None


def click_first(driver, selectors, logs=None, label="요소"):
    for by, selector in selectors:
        try:
            elems = driver.find_elements(by, selector)
            for elem in elems:
                try:
                    driver.execute_script("arguments[0].click();", elem)
                    if logs is not None:
                        add_log(logs, f"{label} 클릭 성공: {selector}")
                    return True
                except Exception:
                    continue
        except Exception:
            continue
    if logs is not None:
        add_log(logs, f"{label} 클릭 실패")
    return False

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
    elements.extend(
        make_pdf_graph_block(
            image_buf=fig_line_buf,
            title_text="5-1. 시간대별 전력사용량 그래프",
            x_label_text="시간",
            y_label_text="전력사용량(kW)",
            font_name=font_name,
            body_style=styles["KBody"],
        )
    )
    elements.append(PageBreak())

    # 3페이지
    elements.extend(
        make_pdf_graph_block(
            image_buf=fig_bar_buf,
            title_text="5-2. 탭별 예상 절감전력 비교",
            x_label_text="탭",
            y_label_text="평균 절감전력(kW)",
            font_name=font_name,
            body_style=styles["KBody"],
        )
    )
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
        options=["파워플래너 자동반영", "파워플래너 값 수동 반영", "수동 입력"],
        horizontal=True,
        index=0,
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
                            st.session_state["pp_hourly_profile_kw"] = result.get("hourly_profile_kw", {})
                            st.session_state["pp_daily_band_kwh"] = result.get("daily_band_kwh", {})
                            st.session_state["pp_monthly_band_kwh"] = result.get("monthly_band_kwh", {})
                            st.session_state["pp_auto_source"] = result.get("auto_source", "")
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
        ratio_color = "auto" if auto_ratio_ready else "verify"

        if auto_ratio_ready:
            st.caption("파워플래너 일별요금 경/중/최대 kWh를 우선 기준으로 자동 산출되며, 필요 시 직접 수정할 수 있습니다.")
            default_avg_kw = st.session_state.get("pp_auto_avg_base_kw", 0.0)
            default_off_ratio = st.session_state.get("pp_auto_off_ratio", 1.0)
            default_mid_ratio = st.session_state.get("pp_auto_mid_ratio", 1.0)
            default_peak_ratio = st.session_state.get("pp_auto_peak_ratio", 1.0)
        else:
            st.warning("파워플래너 시간대별 평균부하 자동 산출값이 없어 기본값을 넣지 않았습니다. 자동반영을 다시 실행하거나 직접 입력으로 전환해 주세요.")
            default_avg_kw = avg_kw_from_annual if pp_loaded and avg_kw_from_annual > 0 else 0.0
            default_off_ratio = 1.0
            default_mid_ratio = 1.0
            default_peak_ratio = 1.0

        avg_base_input = colored_input(
            f"기준 평균부하({load_unit})",
            st.number_input,
            ratio_color,
            min_value=0.0,
            value=float(default_avg_kw / unit_mul),
            step=0.1 if load_unit == "MW" else 100.0,
        )
        avg_base_kw = avg_base_input * unit_mul

        r1, r2, r3 = st.columns(3)
        with r1:
            off_ratio = colored_input(
                "경부하 부하계수",
                st.number_input,
                ratio_color,
                min_value=0.0,
                value=max(float(default_off_ratio or 0.0), 0.0),
                step=0.01,
                format="%.2f",
            )
        with r2:
            mid_ratio = colored_input(
                "중간부하 부하계수",
                st.number_input,
                ratio_color,
                min_value=0.0,
                value=max(float(default_mid_ratio or 0.0), 0.0),
                step=0.01,
                format="%.2f",
            )
        with r3:
            peak_ratio = colored_input(
                "최대부하 부하계수",
                st.number_input,
                ratio_color,
                min_value=0.0,
                value=max(float(default_peak_ratio or 0.0), 0.0),
                step=0.01,
                format="%.2f",
            )

        auto_loads = avg_kw_to_timeband_loads(avg_base_kw, off_ratio, mid_ratio, peak_ratio)
        off_peak_kw = auto_loads["경부하"]
        mid_peak_kw = auto_loads["중간부하"]
        peak_kw = auto_loads["최대부하"]
    else:
        default_off_kw = st.session_state.get("pp_auto_off_peak_kw", 0.0) if pp_loaded else 0.0
        default_mid_kw = st.session_state.get("pp_auto_mid_peak_kw", 0.0) if pp_loaded else 0.0
        default_peak_kw = st.session_state.get("pp_auto_peak_kw", 0.0) if pp_loaded else 0.0
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

# =========================================================
# 파워플래너 시간대별 값 우선 계산 엔진
# =========================================================
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

pp_profile_raw = st.session_state.get("pp_hourly_profile_kw", {})
use_pp_hourly_engine = (
    st.session_state.get("pp_loaded", False)
    and input_mode == "파워플래너 자동 산출"
    and st.session_state.get("pp_auto_source", "") != "daily_bill_timeband"
    and isinstance(pp_profile_raw, dict)
    and len(pp_profile_raw) >= 6
)

if use_pp_hourly_engine:
    kw_by_hour_for_calc = {}
    for h in range(24):
        try:
            val = float(pp_profile_raw.get(h, pp_profile_raw.get(str(h), 0.0)))
        except Exception:
            val = 0.0
        kw_by_hour_for_calc[h] = max(val, 0.0)
else:
    loads_by_label = {
        "경부하": off_peak_kw,
        "중간부하": mid_peak_kw,
        "최대부하": peak_kw,
    }
    kw_by_hour_for_calc = {
        h: (loads_by_label[hour_to_label(h, season)] if h in active_hours else 0.0)
        for h in range(24)
    }

if use_pp_hourly_engine:
    band_values_positive = {"경부하": [], "중간부하": [], "최대부하": []}
    band_energy_positive = {"경부하": 0.0, "중간부하": 0.0, "최대부하": 0.0}
    positive_values = []
    for h, kw in kw_by_hour_for_calc.items():
        if h not in active_hours or kw <= 0:
            continue
        band = hour_to_label(h, season)
        band_values_positive[band].append(kw)
        band_energy_positive[band] += kw
        positive_values.append(kw)

    base_avg_for_ratio = (sum(positive_values) / len(positive_values)) if positive_values else 0.0
    loads_by_label = {}
    for label in ["경부하", "중간부하", "최대부하"]:
        vals = band_values_positive[label]
        loads_by_label[label] = (sum(vals) / len(vals)) if vals else 0.0

    if base_avg_for_ratio > 0:
        avg_base_kw = base_avg_for_ratio
        off_peak_kw = loads_by_label["경부하"]
        mid_peak_kw = loads_by_label["중간부하"]
        peak_kw = loads_by_label["최대부하"]
        off_ratio = off_peak_kw / avg_base_kw
        mid_ratio = mid_peak_kw / avg_base_kw
        peak_ratio = peak_kw / avg_base_kw

    total_positive_energy = sum(band_energy_positive.values())
    energy_share_by_label = {
        label: (band_energy_positive[label] / total_positive_energy if total_positive_energy else 0.0)
        for label in ["경부하", "중간부하", "최대부하"]
    }
else:
    loads_by_label = {
        "경부하": off_peak_kw,
        "중간부하": mid_peak_kw,
        "최대부하": peak_kw,
    }
    daily_energy_for_share = {
        label: loads_by_label[label] * operating_hour_count[label]
        for label in ["경부하", "중간부하", "최대부하"]
    }
    total_energy_for_share = sum(daily_energy_for_share.values())
    energy_share_by_label = {
        label: (daily_energy_for_share[label] / total_energy_for_share if total_energy_for_share else 0.0)
        for label in ["경부하", "중간부하", "최대부하"]
    }

hour_rows = []
daily_base_kwh = 0.0
daily_saved_kwh = 0.0
daily_cost_saving = 0.0
band_calc = {
    label: {"hours": 0, "base_kwh": 0.0, "saved_kwh": 0.0, "cost": 0.0, "kw_sum": 0.0, "kw_count": 0}
    for label in ["경부하", "중간부하", "최대부하"]
}

for h in range(24):
    label = hour_to_label(h, season)
    operating = h in active_hours
    kw = kw_by_hour_for_calc.get(h, 0.0) if operating else 0.0
    rate = rate_table[season][label]
    saving_rate_pct, saved_kw, saved_kwh = calc_average_result(kw, voltage_drop_pct, defaults["cvrf"], z, i, p, 1.0)
    cost_save = saved_kwh * rate

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
        "절감요금(원)": round(cost_save, 1),
    })

    include_for_summary = operating and (kw > 0 if use_pp_hourly_engine else True)
    if include_for_summary:
        band_calc[label]["hours"] += 1
        band_calc[label]["base_kwh"] += kw
        band_calc[label]["saved_kwh"] += saved_kwh
        band_calc[label]["cost"] += cost_save
        band_calc[label]["kw_sum"] += kw
        band_calc[label]["kw_count"] += 1
        daily_base_kwh += kw
        daily_saved_kwh += saved_kwh
        daily_cost_saving += cost_save

hourly_df = pd.DataFrame(hour_rows)

rows = []
for label in ["경부하", "중간부하", "최대부하"]:
    summary_hours = band_calc[label]["hours"]
    avg_kw_label = (
        band_calc[label]["kw_sum"] / band_calc[label]["kw_count"]
        if band_calc[label]["kw_count"] > 0
        else loads_by_label.get(label, 0.0)
    )
    base_kwh = band_calc[label]["base_kwh"]
    saved_kwh = band_calc[label]["saved_kwh"]
    cost_save = band_calc[label]["cost"]
    saving_rate_pct = (saved_kwh / base_kwh * 100.0) if base_kwh else 0.0
    saved_kw = (saved_kwh / summary_hours) if summary_hours else 0.0

    rows.append({
        "구분": label,
        "운영시간(h/day)": summary_hours,
        "평균부하(kW)": round(avg_kw_label, 2),
        "사용전력량(kWh/day)": round(base_kwh, 2),
        "절감률(%)": round(saving_rate_pct, 3),
        "절감전력(kW)": round(saved_kw, 2),
        "절감전력량(kWh/day)": round(saved_kwh, 2),
        "적용단가(원/kWh)": round(rate_table[season][label], 1),
        "절감요금(원/day)": round(cost_save, 1),
        "kWh비중(%)": round(energy_share_by_label.get(label, 0.0) * 100.0, 2),
    })

period_df = pd.DataFrame(rows)

saving_rate_total = (daily_saved_kwh / daily_base_kwh * 100.0) if daily_base_kwh else 0.0
day_operation_hours = sum(band_calc[label]["hours"] for label in ["경부하", "중간부하", "최대부하"])
avg_saved_kw = (daily_saved_kwh / day_operation_hours) if day_operation_hours else 0.0
monthly_saved_kwh = daily_saved_kwh * (active_days_per_year / 12.0 if active_days_per_year else 0.0)
yearly_saved_kwh = daily_saved_kwh * active_days_per_year
monthly_cost_saving = daily_cost_saving * (active_days_per_year / 12.0 if active_days_per_year else 0.0)
yearly_cost_saving = daily_cost_saving * active_days_per_year

max_compare_drop_pct = 7.5
tap_compare_rows = []
current_tap_int = int(current_tap)
base_compare_voltage = current_voltage_for_calc
tap_compare_rows.append({
    "탭": current_tap_int,
    "계산 기준 전압(V)": round(base_compare_voltage, 1),
    "전압 저감률(%)": 0.0,
    "절감률(%)": 0.0,
    "평균 절감전력(kW)": 0.0,
    "일 절감량(kWh)": 0.0,
    "연 절감량(kWh)": 0.0,
    "일 절감요금(원)": 0.0,
    "연 절감요금(원)": 0.0,
})

for tap in range(max(current_tap_int - 1, 1), 0, -1):
    delta_steps = max(current_tap_int - int(tap), 0)
    tap_voltage_drop_pct = calc_tap_voltage_change(tap_step_percent, delta_steps)
    if tap_voltage_drop_pct < 1.25 or tap_voltage_drop_pct > max_compare_drop_pct:
        continue
    tap_new_voltage = current_voltage_for_calc * (1 - tap_voltage_drop_pct / 100.0)
    tap_daily_saved_kwh = 0.0
    tap_daily_cost = 0.0
    tap_base_kwh = 0.0
    tap_hours = 0

    for h in range(24):
        if h not in active_hours:
            continue
        kw = kw_by_hour_for_calc.get(h, 0.0)
        if use_pp_hourly_engine and kw <= 0:
            continue
        label = hour_to_label(h, season)
        rate = rate_table[season][label]
        _, _, tap_saved_kwh = calc_average_result(kw, tap_voltage_drop_pct, defaults["cvrf"], z, i, p, 1.0)
        tap_daily_saved_kwh += tap_saved_kwh
        tap_daily_cost += tap_saved_kwh * rate
        tap_base_kwh += kw
        tap_hours += 1

    tap_avg_saved_kw = (tap_daily_saved_kwh / tap_hours) if tap_hours else 0.0
    tap_saving_rate_total = (tap_daily_saved_kwh / tap_base_kwh * 100.0) if tap_base_kwh else 0.0
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
        ((tap_compare_df["전압 저감률(%)"] >= 1.25) & (tap_compare_df["전압 저감률(%)"] <= max_compare_drop_pct))
    ].copy()
    tap_compare_df = tap_compare_df.drop_duplicates(subset=["탭"], keep="first")
    tap_compare_df = tap_compare_df.sort_values(by=["탭"], ascending=[False]).reset_index(drop=True)

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
            st.write("- 자동 산출 경/중/최 부하계수: **{0:.2f} / {1:.2f} / {2:.2f}**".format(
                st.session_state.get("pp_auto_off_ratio", 0.0),
                st.session_state.get("pp_auto_mid_ratio", 0.0),
                st.session_state.get("pp_auto_peak_ratio", 0.0),
            ))
            if st.session_state.get("pp_auto_source", "") == "daily_bill_timeband":
                st.write("- 계산 기준: **파워플래너 일별요금 경/중/최대 kWh 우선**")
                monthly_band = st.session_state.get("pp_monthly_band_kwh", {}) or {}
                total_band = sum(float(v) for v in monthly_band.values()) if isinstance(monthly_band, dict) else 0.0
                if total_band > 0:
                    st.write("- 월 경/중/최 kWh 비중: **{0:.2f}% / {1:.2f}% / {2:.2f}%**".format(
                        float(monthly_band.get("경부하", 0.0)) / total_band * 100.0,
                        float(monthly_band.get("중간부하", 0.0)) / total_band * 100.0,
                        float(monthly_band.get("최대부하", 0.0)) / total_band * 100.0,
                    ))

    st.markdown("</div>", unsafe_allow_html=True)

st.divider()
st.subheader("그래프")
g1, g2 = st.columns(2)


with g1:
    graph_df = hourly_df.copy()
    use_pp_profile = (
        st.session_state.get("pp_loaded", False)
        and input_mode == "파워플래너 자동 산출"
        and isinstance(st.session_state.get("pp_hourly_profile_kw", {}), dict)
        and len(st.session_state.get("pp_hourly_profile_kw", {})) >= 6
    )
    if use_pp_profile:
        pp_graph_rows = []
        pp_profile = st.session_state.get("pp_hourly_profile_kw", {})
        for h in range(24):
            val = float(pp_profile.get(h, 0.0))
            pp_graph_rows.append({
                "시간번호": h,
                "시간": f"{h:02d}:00",
                "전력사용량(kW)": val,
                "구분": hour_to_label(h, season),
                "운영여부": "가동" if h in active_hours else "비가동",
                "절감전력(kW)": 0.0,
                "절감전력량(kWh)": 0.0,
            })
        graph_df = pd.DataFrame(pp_graph_rows)

    fig_load = px.line(
        graph_df,
        x="시간번호",
        y="전력사용량(kW)",
        markers=True,
        hover_data=["시간", "구분", "운영여부"],
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
            customdata=tap_chart_df[[
                "계산 기준 전압(V)",
                "전압 저감률(%)",
                "절감률(%)",
                "일 절감량(kWh)",
                "연 절감량(kWh)",
                "일 절감요금(원)",
                "연 절감요금(원)",
            ]].values,
            hovertemplate=(
                "탭=%{x}<br>"
                "평균 절감전력=%{y:,.3f} kW<br>"
                "계산 기준 전압=%{customdata[0]:,.1f} V<br>"
                "전압 저감률=%{customdata[1]:.2f}%<br>"
                "절감률=%{customdata[2]:.3f}%<br>"
                "일 절감량=%{customdata[3]:,.3f} kWh<br>"
                "연 절감량=%{customdata[4]:,.1f} kWh<br>"
                "일 절감요금=%{customdata[5]:,.0f} 원<br>"
                "연 절감요금=%{customdata[6]:,.0f} 원"
                "<extra></extra>"
            ),
        )
        fig_bar.update_layout(title="탭별 예상 절감전력 비교", hovermode="closest")
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

# PDF는 생성 시간이 길기 때문에 화면 재계산 때마다 만들지 않고,
# 아래 PDF 저장 버튼을 누른 경우에만 생성한다.

col_excel, col_pdf = st.columns(2)

with col_excel:
    st.download_button(
        "📊 엑셀 저장",
        data=excel_output.getvalue(),
        file_name="CVR결과_{0}_{1}.xlsx".format(site_name, datetime.now().strftime("%Y%m%d_%H%M%S")),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with col_pdf:
    if st.button("📄 PDF 생성"):
        try:
            pdf_font_info = get_korean_font_info()
            fig_line_buf = create_matplotlib_line_chart(graph_df, font_info=pdf_font_info)
            fig_bar_buf = create_matplotlib_bar_chart(tap_compare_df, font_info=pdf_font_info)

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
            st.download_button(
                "📄 PDF 다운로드",
                data=pdf_bytes,
                file_name="CVR보고서_{0}_{1}.pdf".format(site_name, datetime.now().strftime("%Y%m%d_%H%M%S")),
                mime="application/pdf",
            )
        except Exception as e:
            st.error("PDF 생성 실패: {0}".format(str(e)))
