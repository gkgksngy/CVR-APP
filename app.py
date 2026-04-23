import os
import io
import tempfile
from datetime import datetime

import pandas as pd
import streamlit as st
import plotly.express as px

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Image,
    Table,
    TableStyle,
    PageBreak,
)


# -------------------------------------------------
# 페이지 설정
# -------------------------------------------------
st.set_page_config(page_title="CVR 운영형 계산기", layout="wide")

st.markdown(
    """
    <meta name="google" content="notranslate">
    <style>
    .notranslate { translate: no; }

    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 1.5rem;
    }

    div[data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e9edf3;
        border-radius: 12px;
        padding: 10px 14px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# -------------------------------------------------
# 기본 데이터
# -------------------------------------------------
LOAD_TYPES = {
    "사무/상업 혼합": {"cvrf": 0.85, "z": 0.35, "i": 0.25, "p": 0.40},
    "공장 혼합": {"cvrf": 0.72, "z": 0.25, "i": 0.25, "p": 0.50},
    "모터 부하 많음": {"cvrf": 0.60, "z": 0.20, "i": 0.35, "p": 0.45},
    "조명/히터 부하 많음": {"cvrf": 0.95, "z": 0.55, "i": 0.20, "p": 0.25},
    "인버터/SMPS 부하 많음": {"cvrf": 0.40, "z": 0.10, "i": 0.15, "p": 0.75},
}

LOAD_ORDER = ["경부하", "중간부하", "최대부하"]

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


# -------------------------------------------------
# 중요
# 아래 BASE_TARIFFS는 네가 기존에 쓰던 긴 요금표 블록 그대로 붙여넣으면 됨
# -------------------------------------------------
BASE_TARIFFS = {
    "산업용": {
        "갑": {
            "고압A": {
                "선택 I": {
                    "여름": {"경부하": 118.0, "중간부하": 170.0, "최대부하": 248.0},
                    "봄·가을": {"경부하": 118.0, "중간부하": 140.0, "최대부하": 170.0},
                    "겨울": {"경부하": 125.0, "중간부하": 170.0, "최대부하": 225.0},
                },
                "선택 II": {
                    "여름": {"경부하": 121.0, "중간부하": 175.0, "최대부하": 255.0},
                    "봄·가을": {"경부하": 121.0, "중간부하": 144.0, "최대부하": 175.0},
                    "겨울": {"경부하": 128.0, "중간부하": 175.0, "최대부하": 231.0},
                },
                "선택 III": {
                    "여름": {"경부하": 125.0, "중간부하": 181.0, "최대부하": 262.0},
                    "봄·가을": {"경부하": 125.0, "중간부하": 149.0, "최대부하": 182.0},
                    "겨울": {"경부하": 132.0, "중간부하": 181.0, "최대부하": 238.0},
                },
            },
            "고압B": {
                "선택 I": {
                    "여름": {"경부하": 112.0, "중간부하": 163.0, "최대부하": 240.0},
                    "봄·가을": {"경부하": 112.0, "중간부하": 135.0, "최대부하": 166.0},
                    "겨울": {"경부하": 119.0, "중간부하": 163.0, "최대부하": 218.0},
                },
                "선택 II": {
                    "여름": {"경부하": 115.0, "중간부하": 168.0, "최대부하": 246.0},
                    "봄·가을": {"경부하": 115.0, "중간부하": 139.0, "최대부하": 171.0},
                    "겨울": {"경부하": 122.0, "중간부하": 168.0, "최대부하": 224.0},
                },
                "선택 III": {
                    "여름": {"경부하": 119.0, "중간부하": 174.0, "최대부하": 253.0},
                    "봄·가을": {"경부하": 119.0, "중간부하": 144.0, "최대부하": 178.0},
                    "겨울": {"경부하": 126.0, "중간부하": 174.0, "최대부하": 231.0},
                },
            },
            "고압C": {
                "선택 I": {
                    "여름": {"경부하": 106.0, "중간부하": 155.0, "최대부하": 231.0},
                    "봄·가을": {"경부하": 106.0, "중간부하": 129.0, "최대부하": 160.0},
                    "겨울": {"경부하": 112.0, "중간부하": 155.0, "최대부하": 210.0},
                },
                "선택 II": {
                    "여름": {"경부하": 109.0, "중간부하": 160.0, "최대부하": 236.0},
                    "봄·가을": {"경부하": 109.0, "중간부하": 133.0, "최대부하": 165.0},
                    "겨울": {"경부하": 115.0, "중간부하": 160.0, "최대부하": 215.0},
                },
                "선택 III": {
                    "여름": {"경부하": 113.0, "중간부하": 166.0, "최대부하": 243.0},
                    "봄·가을": {"경부하": 113.0, "중간부하": 138.0, "최대부하": 172.0},
                    "겨울": {"경부하": 119.0, "중간부하": 166.0, "최대부하": 222.0},
                },
            },
        },
        "을": {
            "고압A": {
                "선택 I": {
                    "여름": {"경부하": 120.8, "중간부하": 173.1, "최대부하": 254.4},
                    "봄·가을": {"경부하": 120.8, "중간부하": 143.2, "최대부하": 173.5},
                    "겨울": {"경부하": 127.9, "중간부하": 173.1, "최대부하": 229.3},
                },
                "선택 II": {
                    "여름": {"경부하": 124.0, "중간부하": 178.0, "최대부하": 261.0},
                    "봄·가을": {"경부하": 124.0, "중간부하": 147.0, "최대부하": 179.0},
                    "겨울": {"경부하": 131.0, "중간부하": 178.0, "최대부하": 236.0},
                },
                "선택 III": {
                    "여름": {"경부하": 128.0, "중간부하": 184.0, "최대부하": 269.0},
                    "봄·가을": {"경부하": 128.0, "중간부하": 152.0, "최대부하": 186.0},
                    "겨울": {"경부하": 135.0, "중간부하": 184.0, "최대부하": 243.0},
                },
            }
        },
    },
    "일반용": {
        "갑": {
            "고압A": {
                "선택 I": {
                    "여름": {"경부하": 60.9, "중간부하": 111.4, "최대부하": 132.7},
                    "봄·가을": {"경부하": 60.9, "중간부하": 68.5, "최대부하": 79.2},
                    "겨울": {"경부하": 68.6, "중간부하": 98.7, "최대부하": 112.7},
                }
            }
        }
    },
    "농사용": {
        "갑": {
            "저압": {
                "선택 II": {
                    "여름": {"경부하": 48.0, "중간부하": 55.0, "최대부하": 63.0},
                    "봄·가을": {"경부하": 46.0, "중간부하": 52.0, "최대부하": 58.0},
                    "겨울": {"경부하": 50.0, "중간부하": 57.0, "최대부하": 64.0},
                }
            }
        }
    },
    "주택용": {
        "갑": {
            "저압": {
                "선택 II": {
                    "여름": {"경부하": 85.0, "중간부하": 95.0, "최대부하": 110.0},
                    "봄·가을": {"경부하": 80.0, "중간부하": 90.0, "최대부하": 100.0},
                    "겨울": {"경부하": 88.0, "중간부하": 98.0, "최대부하": 112.0},
                }
            }
        }
    },
}


# -------------------------------------------------
# 계산 함수
# -------------------------------------------------
def get_voltage_class(primary_kv: float) -> str:
    if primary_kv < 1.0:
        return "저압"
    if primary_kv < 154.0:
        return "고압A"
    if primary_kv < 345.0:
        return "고압B"
    return "고압C"


def calc_tap_voltage_change(step_percent: float, tap_change_steps: int) -> float:
    return step_percent * tap_change_steps


def calc_cvrf(base_kw: float, voltage_drop_pct: float, cvrf: float, hours: float):
    saving_rate_pct = voltage_drop_pct * cvrf
    kw_saved = base_kw * (saving_rate_pct / 100.0)
    kwh_saved = kw_saved * hours
    return saving_rate_pct, kw_saved, kwh_saved


def calc_zip(base_kw: float, voltage_drop_pct: float, z: float, i: float, p: float, hours: float):
    if base_kw <= 0:
        return 0.0, 0.0, 0.0
    v_pu = 1.0 - voltage_drop_pct / 100.0
    new_kw = base_kw * (z * (v_pu ** 2) + i * v_pu + p)
    kw_saved = base_kw - new_kw
    saving_rate_pct = (kw_saved / base_kw) * 100.0 if base_kw else 0.0
    kwh_saved = kw_saved * hours
    return saving_rate_pct, kw_saved, kwh_saved


def calc_average_result(load_kw: float, voltage_drop_pct: float, cvrf: float, z: float, i: float, p: float, hours: float):
    cvrf_rate, cvrf_kw, cvrf_kwh = calc_cvrf(load_kw, voltage_drop_pct, cvrf, hours)
    zip_rate, zip_kw, zip_kwh = calc_zip(load_kw, voltage_drop_pct, z, i, p, hours)

    avg_rate = (cvrf_rate + zip_rate) / 2.0
    avg_kw = (cvrf_kw + zip_kw) / 2.0
    avg_kwh = (cvrf_kwh + zip_kwh) / 2.0
    return avg_rate, avg_kw, avg_kwh


def confidence_text(voltage_drop_pct: float, p_share: float, calibration: float):
    if calibration >= 0.97 and voltage_drop_pct <= 2.0 and p_share <= 0.55:
        return "높음"
    if calibration >= 0.90 and voltage_drop_pct <= 3.0 and p_share <= 0.70:
        return "보통"
    return "낮음"


def hour_to_label(hour: int, season: str) -> str:
    schedule = SEASON_SCHEDULE[season]
    for label, hours in schedule.items():
        if hour in hours:
            return label
    return "중간부하"


def format_hour_range(hours):
    if not hours:
        return "-"
    ordered = sorted(hours)
    segments = []
    start = ordered[0]
    prev = ordered[0]
    for h in ordered[1:]:
        if h == prev + 1:
            prev = h
        else:
            segments.append((start, prev + 1))
            start = h
            prev = h
    segments.append((start, prev + 1))

    text_list = []
    for s, e in segments:
        s_txt = f"{s:02d}:00"
        e_txt = "24:00" if e == 24 else f"{e:02d}:00"
        text_list.append(f"{s_txt} ~ {e_txt}")
    return ", ".join(text_list)


def get_recommended_calibration(data_quality: str) -> float:
    mapping = {
        "실측 15분/1시간 데이터 보유 (보정계수 1.00)": 1.00,
        "실측 일부 + 일부 추정 (보정계수 0.95)": 0.95,
        "운영패턴 명확, ZIP/CVRF 일반값 사용 (보정계수 0.90)": 0.90,
        "시간대 평균값만 보유 (보정계수 0.85)": 0.85,
        "개략 검토 수준 (보정계수 0.80)": 0.80,
    }
    return mapping.get(data_quality, 0.85)


def get_active_days_per_year(operation_mode: str, custom_days, holiday_count: int, include_holidays_as_shutdown: bool) -> int:
    weekday_map = {
        "월~금 가동": 5,
        "월~토 가동": 6,
        "주7일 가동": 7,
    }

    if operation_mode in weekday_map:
        weekly_days = weekday_map[operation_mode]
    else:
        weekly_days = len(custom_days)

    base_days = int(round(365.0 * weekly_days / 7.0))
    if include_holidays_as_shutdown:
        base_days -= holiday_count
    return max(base_days, 0)


def get_operating_hours_by_label(season: str, operating_start: int, operating_end: int):
    if operating_start == operating_end:
        active_hours = list(range(24))
    elif operating_start < operating_end:
        active_hours = list(range(operating_start, operating_end))
    else:
        active_hours = list(range(operating_start, 24)) + list(range(0, operating_end))

    counts = {"경부하": 0, "중간부하": 0, "최대부하": 0}
    for hour in active_hours:
        label = hour_to_label(hour, season)
        counts[label] += 1
    return active_hours, counts


def make_selected_season_table_df(season: str):
    schedule = SEASON_SCHEDULE[season]
    return pd.DataFrame([
        {
            "계절": season,
            "경부하 시간": format_hour_range(schedule["경부하"]),
            "중간부하 시간": format_hour_range(schedule["중간부하"]),
            "최대부하 시간": format_hour_range(schedule["최대부하"]),
        }
    ])


def make_rate_table_df(rate_dict: dict, season: str):
    return pd.DataFrame([
        {"구분": "경부하", "요금(원/kWh)": rate_dict[season]["경부하"]},
        {"구분": "중간부하", "요금(원/kWh)": rate_dict[season]["중간부하"]},
        {"구분": "최대부하", "요금(원/kWh)": rate_dict[season]["최대부하"]},
    ])


def calc_profile_based_day(season: str, loads_by_label: dict, voltage_drop_pct: float, applied_cvrf: float,
                           z: float, i: float, p: float, rate_dict: dict, operating_hour_count: dict):
    period_rows = []
    total_base_kwh = 0.0
    total_saved_kwh = 0.0
    total_cost_saving = 0.0

    for label in LOAD_ORDER:
        hours = operating_hour_count[label]
        load_kw = loads_by_label[label]
        rate = rate_dict[season][label]

        avg_rate, avg_kw, avg_kwh = calc_average_result(
            load_kw, voltage_drop_pct, applied_cvrf, z, i, p, hours
        )

        base_kwh = load_kw * hours
        cost_saving = avg_kwh * rate

        total_base_kwh += base_kwh
        total_saved_kwh += avg_kwh
        total_cost_saving += cost_saving

        period_rows.append({
            "구분": label,
            "운영시간수(h/day)": hours,
            "평균 부하(kW)": round(load_kw, 3),
            "사용전력량(kWh/day)": round(base_kwh, 3),
            "절감률(%)": round(avg_rate, 3),
            "절감전력(kW)": round(avg_kw, 3),
            "절감전력량(kWh/day)": round(avg_kwh, 3),
            "요금단가(원/kWh)": round(rate, 3),
            "절감요금(원/day)": round(cost_saving, 1),
        })

    total_operating_hours = sum(operating_hour_count.values())
    daily_avg_saved_kw = total_saved_kwh / total_operating_hours if total_operating_hours else 0.0
    saving_rate_total = (total_saved_kwh / total_base_kwh * 100.0) if total_base_kwh else 0.0

    return {
        "daily_base_kwh": total_base_kwh,
        "daily_saved_kwh": total_saved_kwh,
        "daily_avg_saved_kw": daily_avg_saved_kw,
        "daily_cost_saving": total_cost_saving,
        "saving_rate_total": saving_rate_total,
        "period_df": pd.DataFrame(period_rows),
        "total_operating_hours": total_operating_hours,
    }


def make_hourly_profile_df(season: str, loads_by_label: dict, voltage_drop_pct: float, applied_cvrf: float,
                           z: float, i: float, p: float, rate_dict: dict, active_hours):
    rows = []

    for hour in range(24):
        label = hour_to_label(hour, season)
        load_kw = loads_by_label[label] if hour in active_hours else 0.0
        rate = rate_dict[season][label]

        avg_rate, avg_kw, avg_kwh = calc_average_result(
            load_kw, voltage_drop_pct, applied_cvrf, z, i, p, 1.0
        )

        rows.append({
            "시간": f"{hour:02d}:00",
            "시간번호": hour,
            "구분": label,
            "운영여부": "가동" if hour in active_hours else "비가동",
            "전력사용량(kW)": round(load_kw, 3),
            "절감률(%)": round(avg_rate, 3),
            "절감전력(kW)": round(avg_kw, 3),
            "절감전력량(kWh)": round(avg_kwh, 3),
            "요금단가(원/kWh)": round(rate, 3),
            "절감요금(원)": round(avg_kwh * rate, 1),
        })

    return pd.DataFrame(rows)


def fallback_rate_table():
    return {
        "여름": {"경부하": 120.0, "중간부하": 170.0, "최대부하": 250.0},
        "봄·가을": {"경부하": 120.0, "중간부하": 145.0, "최대부하": 175.0},
        "겨울": {"경부하": 127.0, "중간부하": 170.0, "최대부하": 230.0},
    }


def get_valid_voltage_options(tariff_category: str):
    if tariff_category in ["농사용", "주택용"]:
        return ["저압"]
    return ["자동 판정값 사용", "저압", "고압A", "고압B", "고압C"]


def get_valid_type_options(tariff_category: str):
    if tariff_category in ["산업용", "일반용"]:
        return ["갑", "을"]
    return ["갑"]


def get_valid_option_options(tariff_category: str):
    if tariff_category in ["농사용", "주택용"]:
        return ["선택 II"]
    return ["선택 I", "선택 II", "선택 III"]


def safe_get_rate_table(category: str, tariff_type: str, voltage_final: str, option: str):
    try:
        return BASE_TARIFFS[category][tariff_type][voltage_final][option]
    except KeyError:
        return None


def safe_str(v):
    if pd.isna(v):
        return ""
    if isinstance(v, float):
        return f"{v:,.3f}".rstrip("0").rstrip(".")
    return str(v)


# -------------------------------------------------
# 엑셀 저장
# -------------------------------------------------
def build_excel_file(
    compare_df,
    input_profile_df,
    operation_summary_df,
    selected_rate_df,
    selected_season_df,
    period_detail_df,
    hourly_detail_df,
    scenario_df,
    compare_summary_df,
):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        compare_df.to_excel(writer, index=False, sheet_name="결과요약")
        compare_summary_df.to_excel(writer, index=False, sheet_name="현재안제안안비교")
        input_profile_df.to_excel(writer, index=False, sheet_name="부하입력")
        operation_summary_df.to_excel(writer, index=False, sheet_name="운영조건")
        selected_rate_df.to_excel(writer, index=False, sheet_name="적용요금표")
        selected_season_df.to_excel(writer, index=False, sheet_name="계절시간대")
        period_detail_df.to_excel(writer, index=False, sheet_name="시간대구간결과")
        hourly_detail_df.to_excel(writer, index=False, sheet_name="24시간상세")
        scenario_df.to_excel(writer, index=False, sheet_name="탭별비교")
    output.seek(0)
    return output


# -------------------------------------------------
# PDF 스타일
# -------------------------------------------------
def register_korean_font():
    candidate_fonts = [
        "C:/Windows/Fonts/malgun.ttf",
        "C:/Windows/Fonts/NanumGothic.ttf",
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
    ]
    for font_path in candidate_fonts:
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(TTFont("KoreanFont", font_path))
                return "KoreanFont"
            except Exception:
                pass
    return "Helvetica"


def make_pdf_styles(font_name: str):
    return {
        "title": ParagraphStyle(
            "title",
            fontName=font_name,
            fontSize=20,
            leading=24,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#1b2d55"),
            spaceAfter=8,
        ),
        "subtitle": ParagraphStyle(
            "subtitle",
            fontName=font_name,
            fontSize=9,
            leading=12,
            alignment=TA_CENTER,
            textColor=colors.grey,
            spaceAfter=12,
        ),
        "section": ParagraphStyle(
            "section",
            fontName=font_name,
            fontSize=12,
            leading=16,
            alignment=TA_LEFT,
            textColor=colors.HexColor("#1f3b73"),
            spaceBefore=10,
            spaceAfter=6,
        ),
        "body": ParagraphStyle(
            "body",
            fontName=font_name,
            fontSize=8.5,
            leading=11,
            alignment=TA_LEFT,
        ),
    }


def make_small_table(data, font_name, col_widths=None, header_bg="#dce8f7", font_size=8):
    table = Table(data, colWidths=col_widths)
    table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("FONTSIZE", (0, 0), (-1, -1), font_size),
        ("LEADING", (0, 0), (-1, -1), font_size + 1),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(header_bg)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#b9c2cf")),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    return table


def dataframe_to_story_tables(df: pd.DataFrame, font_name: str, rows_per_chunk: int = 24):
    story = []
    if df.empty:
        return story

    usable_width = landscape(A4)[0] - 24 * mm
    col_count = len(df.columns)
    col_widths = [usable_width / max(col_count, 1)] * col_count

    for start in range(0, len(df), rows_per_chunk):
        chunk = df.iloc[start:start + rows_per_chunk].copy()
        data = [list(chunk.columns)] + chunk.astype(object).fillna("").values.tolist()
        data = [[safe_str(c) for c in row] for row in data]
        table = make_small_table(data, font_name, col_widths=col_widths, font_size=6.4)
        story.append(table)
        story.append(Spacer(1, 6))

    return story


# -------------------------------------------------
# PDF 생성
# -------------------------------------------------
def build_pdf_file(
    site_name,
    season,
    voltage_basis_text,
    primary_voltage_kv,
    secondary_voltage_v,
    est_new_voltage,
    voltage_drop_v,
    primary_voltage_class,
    tariff_voltage_final,
    tariff_category,
    tariff_type,
    tariff_option,
    applied_cvrf,
    z,
    i,
    p,
    calibration,
    confidence,
    voltage_drop_pct,
    final_saved_kw,
    final_saving_rate,
    final_saved_kwh_day,
    final_saved_kwh_month,
    final_saved_kwh_year,
    daily_cost_saving,
    monthly_cost_saving,
    yearly_cost_saving,
    compare_df,
    compare_summary_df,
    operation_summary_df,
    selected_rate_df,
    selected_season_df,
    input_profile_df,
    period_detail_df,
    hourly_detail_df,
    scenario_df,
    fig_load,
    fig_bar,
):
    pdf_buffer = io.BytesIO()
    font_name = register_korean_font()
    styles = make_pdf_styles(font_name)

    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=landscape(A4),
        rightMargin=12 * mm,
        leftMargin=12 * mm,
        topMargin=10 * mm,
        bottomMargin=10 * mm,
    )

    story = []

    # 첫 페이지 제목
    story.append(Paragraph("CVR 운영형 계산 결과 보고서", styles["title"]))
    story.append(Paragraph(f"{site_name} / {season} / 생성일시 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles["subtitle"]))

    # KPI 표
    story.append(Paragraph("1. 핵심 결과", styles["section"]))
    kpi_data = [
        ["항목", "값", "항목", "값", "항목", "값"],
        ["예상 전압 저감률", f"{voltage_drop_pct:.2f}%", "예상 절감 전력", f"{final_saved_kw:,.1f} kW", "예상 절감률", f"{final_saving_rate:.2f}%"],
        ["예상 일 절감량", f"{final_saved_kwh_day:,.1f} kWh", "예상 월 절감량", f"{final_saved_kwh_month:,.1f} kWh", "예상 연 절감량", f"{final_saved_kwh_year:,.1f} kWh"],
        ["예상 일 절감요금", f"{daily_cost_saving:,.0f} 원", "예상 월 절감요금", f"{monthly_cost_saving:,.0f} 원", "예상 연 절감요금", f"{yearly_cost_saving:,.0f} 원"],
    ]
    story.append(make_small_table(kpi_data, font_name, col_widths=[38*mm, 42*mm, 38*mm, 42*mm, 38*mm, 42*mm], font_size=8))
    story.append(Spacer(1, 8))

    # 계산 조건 / 적용 조건 2단 표
    story.append(Paragraph("2. 계산 조건 요약", styles["section"]))

    left_data = [
        ["항목", "값"],
        ["사업장명", site_name],
        ["계절", season],
        ["계산 기준 전압", voltage_basis_text],
        ["1차측 수전전압", f"{primary_voltage_kv:,.1f} kV"],
        ["2차측 평균전압", f"{secondary_voltage_v:,.1f} V"],
        ["변경 후 계산 기준 전압", f"{est_new_voltage:,.1f} V"],
        ["전압 감소 예상치", f"{voltage_drop_v:,.1f} V"],
        ["자동 판정 전압등급", primary_voltage_class],
        ["요금 적용 전압등급", tariff_voltage_final],
    ]

    right_data = [
        ["항목", "값"],
        ["요금 종별", f"{tariff_category} {tariff_type} / {tariff_option}"],
        ["적용 CVRf", f"{applied_cvrf:.2f}"],
        ["적용 ZIP 비율", f"Z {z:.2f} / I {i:.2f} / P {p:.2f}"],
        ["보정계수", f"{calibration:.2f}"],
        ["예상 신뢰도", confidence],
        ["현재 탭", safe_str(current_tap_global)],
        ["변경 탭", safe_str(target_tap_global)],
        ["탭 1스텝당 전압 변화율", f"{tap_step_percent_global:.2f} %"],
        ["비고", "운영시간/요일 반영 결과"],
    ]

    summary_two_col = Table(
        [[
            make_small_table(left_data, font_name, col_widths=[45*mm, 80*mm], font_size=7.8),
            make_small_table(right_data, font_name, col_widths=[45*mm, 80*mm], font_size=7.8),
        ]],
        colWidths=[130*mm, 130*mm]
    )
    story.append(summary_two_col)
    story.append(Spacer(1, 8))

    # 현재안 / 제안안 비교
    story.append(Paragraph("3. 현재안 / 제안안 비교", styles["section"]))
    story.extend(dataframe_to_story_tables(compare_summary_df, font_name, rows_per_chunk=20))

    # 그래프
    temp_files = []
    try:
        load_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        bar_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        load_temp.close()
        bar_temp.close()

        fig_load.write_image(load_temp.name, width=1400, height=650, scale=2)
        fig_bar.write_image(bar_temp.name, width=1400, height=650, scale=2)

        temp_files.extend([load_temp.name, bar_temp.name])

        story.append(Paragraph("4. 그래프", styles["section"]))
        story.append(Paragraph("4-1. 시간대별 전력사용량 그래프", styles["body"]))
        story.append(Spacer(1, 4))
        story.append(Image(load_temp.name, width=250 * mm, height=105 * mm))
        story.append(Spacer(1, 8))
        story.append(Paragraph("4-2. 탭별 예상 절감전력 비교", styles["body"]))
        story.append(Spacer(1, 4))
        story.append(Image(bar_temp.name, width=250 * mm, height=105 * mm))

        story.append(PageBreak())

        # 상세 표들
        story.append(Paragraph("5. 계산 결과 요약표", styles["section"]))
        story.extend(dataframe_to_story_tables(compare_df, font_name, rows_per_chunk=22))

        story.append(Paragraph("6. 운영 조건 요약", styles["section"]))
        story.extend(dataframe_to_story_tables(operation_summary_df, font_name, rows_per_chunk=22))

        story.append(Paragraph("7. 적용 요금표", styles["section"]))
        story.extend(dataframe_to_story_tables(selected_rate_df, font_name, rows_per_chunk=22))

        story.append(Paragraph("8. 계절별 시간대 구분표", styles["section"]))
        story.extend(dataframe_to_story_tables(selected_season_df, font_name, rows_per_chunk=22))

        story.append(Paragraph("9. 시간대별 평균 부하 입력 요약", styles["section"]))
        story.extend(dataframe_to_story_tables(input_profile_df, font_name, rows_per_chunk=22))

        story.append(PageBreak())

        story.append(Paragraph("10. 시간대 구간별 계산 결과", styles["section"]))
        story.extend(dataframe_to_story_tables(period_detail_df, font_name, rows_per_chunk=20))

        story.append(Paragraph("11. 24시간 시간대별 상세 결과", styles["section"]))
        story.extend(dataframe_to_story_tables(hourly_detail_df, font_name, rows_per_chunk=18))

        story.append(Paragraph("12. 탭별 비교표", styles["section"]))
        story.extend(dataframe_to_story_tables(scenario_df, font_name, rows_per_chunk=20))

        doc.build(story)
        pdf_buffer.seek(0)
        return pdf_buffer

    finally:
        for path in temp_files:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except Exception:
                    pass


# -------------------------------------------------
# 화면
# -------------------------------------------------
st.title("CVR 운영형 계산기")
st.caption("실사용 및 미팅용 추정 화면입니다. 최종 적용 전에는 최신 계약요금표, ETAP 상세해석, 말단 전압 검토가 필요합니다.")

left, right = st.columns([1.75, 1.0])

with left:
    st.subheader("입력 조건")

    site_name = st.text_input("사업장명", "예시 공장")
    season = st.selectbox("계절 선택", ["봄·가을", "여름", "겨울"], index=0)

    st.markdown("### 전압 조건")
    col1, col2, col3 = st.columns(3)
    with col1:
        primary_voltage_kv = st.number_input("1차측 수전전압 (kV)", min_value=0.1, value=22.9, step=0.1)
    with col2:
        secondary_voltage_v = st.number_input("2차측 평균 전압 (V)", min_value=1.0, value=22900.0, step=100.0)
    with col3:
        calc_voltage_basis = st.selectbox("CVR 계산 기준 전압", ["2차측 전압 사용", "1차측 전압 사용"], index=0)

    primary_voltage_class = get_voltage_class(primary_voltage_kv)
    st.info(f"자동 판정 전압등급: {primary_voltage_class}")

    if calc_voltage_basis == "2차측 전압 사용":
        current_voltage_for_calc = secondary_voltage_v
        voltage_basis_text = "2차측"
    else:
        current_voltage_for_calc = primary_voltage_kv * 1000.0
        voltage_basis_text = "1차측"

    tap1, tap2, tap3 = st.columns(3)
    with tap1:
        current_tap = st.number_input("현재 탭", value=4, step=1)
    with tap2:
        target_tap = st.number_input("변경 탭", value=2, step=1)
    with tap3:
        tap_step_percent = st.number_input("탭 1스텝당 전압 변화율 (%)", min_value=0.1, value=1.25, step=0.05)

    st.markdown("### 부하 조건")
    load_type = st.selectbox("부하 유형", list(LOAD_TYPES.keys()), index=1)
    defaults = LOAD_TYPES[load_type]

    st.markdown("### ZIP 설정")
    use_custom_zip = st.checkbox("ZIP 비율 직접 입력", value=False)

    if use_custom_zip:
        z_raw = st.slider("Z 비율", 0.0, 1.0, float(defaults["z"]), 0.01)
        i_raw = st.slider("I 비율", 0.0, 1.0, float(defaults["i"]), 0.01)
        p_raw = st.slider("P 비율", 0.0, 1.0, float(defaults["p"]), 0.01)

        total_zip = z_raw + i_raw + p_raw
        if total_zip == 0:
            z = defaults["z"]
            i = defaults["i"]
            p = defaults["p"]
            st.warning("ZIP 비율 합계가 0이라 기본값으로 복원했습니다.")
        else:
            z = z_raw / total_zip
            i = i_raw / total_zip
            p = p_raw / total_zip
    else:
        z = defaults["z"]
        i = defaults["i"]
        p = defaults["p"]

    st.info(f"현재 ZIP 합계: {z + i + p:.2f} (정상값 1.00)")

    st.markdown("### 보정계수")
    calibration_quality = st.selectbox(
        "입력 데이터 신뢰도",
        [
            "실측 15분/1시간 데이터 보유 (보정계수 1.00)",
            "실측 일부 + 일부 추정 (보정계수 0.95)",
            "운영패턴 명확, ZIP/CVRF 일반값 사용 (보정계수 0.90)",
            "시간대 평균값만 보유 (보정계수 0.85)",
            "개략 검토 수준 (보정계수 0.80)",
        ],
        index=3,
    )

    recommended_calibration = get_recommended_calibration(calibration_quality)
    calibration_mode = st.radio("보정계수 적용 방식", ["자동 추천값 사용", "직접 입력"], horizontal=True)

    if calibration_mode == "자동 추천값 사용":
        calibration = recommended_calibration
        st.success(f"권장 보정계수: {calibration:.2f}")
    else:
        calibration = st.number_input("보정계수 직접 입력", min_value=0.50, max_value=1.10, value=float(recommended_calibration), step=0.01)

    st.caption("권장 기준: 1.00 / 0.95 / 0.90 / 0.85 / 0.80")

    st.markdown("### 요금 조건")
    tariff_category = st.selectbox("요금 종별", ["산업용", "일반용", "농사용", "주택용"], index=0)

    type_options = get_valid_type_options(tariff_category)
    tariff_type = st.selectbox("세부 구분", type_options, index=min(1, len(type_options) - 1))

    voltage_options = get_valid_voltage_options(tariff_category)
    tariff_voltage = st.selectbox("요금 전압 구분", voltage_options, index=0)

    option_options = get_valid_option_options(tariff_category)
    tariff_option = st.selectbox("선택 요금제", option_options, index=min(1, len(option_options) - 1))

    if tariff_voltage == "자동 판정값 사용":
        tariff_voltage_final = primary_voltage_class
    else:
        tariff_voltage_final = tariff_voltage

    rate_mode = st.radio("전력량 요금 적용 방식", ["자동 요금표 사용", "직접 입력"], horizontal=True)

    auto_rate_table = safe_get_rate_table(
        category=tariff_category,
        tariff_type=tariff_type,
        voltage_final=tariff_voltage_final,
        option=tariff_option,
    )

    rate_warning = None
    if auto_rate_table is None:
        auto_rate_table = fallback_rate_table()
        rate_warning = f"선택 조합({tariff_category}/{tariff_type}/{tariff_voltage_final}/{tariff_option})의 요금표가 없어 기본 예시 단가로 계산합니다."

    if rate_mode == "자동 요금표 사용":
        active_rate_table = auto_rate_table
        if rate_warning:
            st.warning(rate_warning)
        else:
            st.success("현재 선택한 요금 조건이 계산 단가에 반영됩니다.")
    else:
        st.caption("현재 선택 계절 기준으로 경부하 / 중간부하 / 최대부하 단가를 직접 입력합니다.")
        r1, r2, r3 = st.columns(3)
        default_rates = auto_rate_table[season]

        with r1:
            custom_off = st.number_input("경부하 요금(원/kWh)", min_value=0.0, value=float(default_rates["경부하"]), step=0.1)
        with r2:
            custom_mid = st.number_input("중간부하 요금(원/kWh)", min_value=0.0, value=float(default_rates["중간부하"]), step=0.1)
        with r3:
            custom_peak = st.number_input("최대부하 요금(원/kWh)", min_value=0.0, value=float(default_rates["최대부하"]), step=0.1)

        active_rate_table = {
            "여름": {"경부하": custom_off, "중간부하": custom_mid, "최대부하": custom_peak},
            "봄·가을": {"경부하": custom_off, "중간부하": custom_mid, "최대부하": custom_peak},
            "겨울": {"경부하": custom_off, "중간부하": custom_mid, "최대부하": custom_peak},
        }

    st.markdown("### 운영 조건")
    operation_mode = st.selectbox("가동 요일", ["월~금 가동", "월~토 가동", "주7일 가동", "사용자 지정"], index=0)

    custom_days = []
    if operation_mode == "사용자 지정":
        custom_days = st.multiselect("가동 요일 선택", ["월", "화", "수", "목", "금", "토", "일"], default=["월", "화", "수", "목", "금"])

    hh1, hh2 = st.columns(2)
    with hh1:
        operating_start = st.selectbox("가동 시작 시간", list(range(24)), index=9, format_func=lambda x: f"{x:02d}:00")
    with hh2:
        operating_end = st.selectbox("가동 종료 시간", list(range(24)), index=18, format_func=lambda x: f"{x:02d}:00")

    st.caption("시작시간과 종료시간이 같으면 24시간 가동으로 처리합니다.")

    holiday_reflect = st.checkbox("공휴일 비가동 반영", value=True)
    holiday_count = st.number_input("연간 공휴일 수", min_value=0, value=15, step=1)

    st.markdown("### 시간대별 평균 부하 입력")
    load_unit = st.radio("부하 입력 단위", ["kW", "MW"], horizontal=True, index=1)

    unit_multiplier = 1000.0 if load_unit == "MW" else 1.0
    default_off = 40.0 if load_unit == "MW" else 40000.0
    default_mid = 50.0 if load_unit == "MW" else 50000.0
    default_peak = 60.0 if load_unit == "MW" else 60000.0

    off_peak_input = st.number_input(f"경부하 평균 부하 ({load_unit})", min_value=0.0, value=default_off, step=1.0 if load_unit == "MW" else 1000.0)
    mid_peak_input = st.number_input(f"중간부하 평균 부하 ({load_unit})", min_value=0.0, value=default_mid, step=1.0 if load_unit == "MW" else 1000.0)
    peak_input = st.number_input(f"최대부하 평균 부하 ({load_unit})", min_value=0.0, value=default_peak, step=1.0 if load_unit == "MW" else 1000.0)

    off_peak_kw = off_peak_input * unit_multiplier
    mid_peak_kw = mid_peak_input * unit_multiplier
    peak_kw = peak_input * unit_multiplier


# -------------------------------------------------
# 계산
# -------------------------------------------------
tap_delta_steps = max(int(current_tap - target_tap), 0)
voltage_drop_pct = calc_tap_voltage_change(tap_step_percent, tap_delta_steps)
est_new_voltage = current_voltage_for_calc * (1 - voltage_drop_pct / 100.0)
voltage_drop_v = current_voltage_for_calc - est_new_voltage

applied_cvrf = defaults["cvrf"] * calibration
confidence = confidence_text(voltage_drop_pct, p, calibration)

loads_by_label = {
    "경부하": off_peak_kw,
    "중간부하": mid_peak_kw,
    "최대부하": peak_kw,
}

active_hours, operating_hour_count = get_operating_hours_by_label(
    season=season,
    operating_start=operating_start,
    operating_end=operating_end,
)

active_days_per_year = get_active_days_per_year(
    operation_mode=operation_mode,
    custom_days=custom_days,
    holiday_count=int(holiday_count),
    include_holidays_as_shutdown=holiday_reflect,
)

selected_season_df = make_selected_season_table_df(season)
selected_rate_df = make_rate_table_df(active_rate_table, season)

profile_result = calc_profile_based_day(
    season=season,
    loads_by_label=loads_by_label,
    voltage_drop_pct=voltage_drop_pct,
    applied_cvrf=applied_cvrf,
    z=z,
    i=i,
    p=p,
    rate_dict=active_rate_table,
    operating_hour_count=operating_hour_count,
)

hourly_detail_df = make_hourly_profile_df(
    season=season,
    loads_by_label=loads_by_label,
    voltage_drop_pct=voltage_drop_pct,
    applied_cvrf=applied_cvrf,
    z=z,
    i=i,
    p=p,
    rate_dict=active_rate_table,
    active_hours=active_hours,
)

final_saving_rate = profile_result["saving_rate_total"]
final_saved_kw = profile_result["daily_avg_saved_kw"]
final_saved_kwh_day = profile_result["daily_saved_kwh"]
daily_cost_saving = profile_result["daily_cost_saving"]
period_detail_df = profile_result["period_df"].copy()

monthly_cost_saving = daily_cost_saving * (active_days_per_year / 12.0 if active_days_per_year else 0.0)
yearly_cost_saving = daily_cost_saving * active_days_per_year

final_saved_kwh_month = final_saved_kwh_day * (active_days_per_year / 12.0 if active_days_per_year else 0.0)
final_saved_kwh_year = final_saved_kwh_day * active_days_per_year

compare_df = pd.DataFrame([
    {
        "모델": "운영시간/요일 반영 결과",
        "예상 절감률(%)": round(final_saving_rate, 3),
        "예상 평균 절감전력(kW)": round(final_saved_kw, 3),
        "예상 일 절감량(kWh)": round(final_saved_kwh_day, 3),
        "예상 연 절감량(kWh)": round(final_saved_kwh_year, 1),
        "예상 일 절감요금(원)": round(daily_cost_saving, 1),
        "예상 연 절감요금(원)": round(yearly_cost_saving, 0),
    }
])

input_profile_df = pd.DataFrame([
    {"구분": "경부하", f"입력 평균부하({load_unit})": off_peak_input, "적용 운영시간(h/day)": operating_hour_count["경부하"]},
    {"구분": "중간부하", f"입력 평균부하({load_unit})": mid_peak_input, "적용 운영시간(h/day)": operating_hour_count["중간부하"]},
    {"구분": "최대부하", f"입력 평균부하({load_unit})": peak_input, "적용 운영시간(h/day)": operating_hour_count["최대부하"]},
])

operation_summary_df = pd.DataFrame([
    {
        "가동 요일": operation_mode if operation_mode != "사용자 지정" else ", ".join(custom_days) if custom_days else "선택 없음",
        "가동 시작": f"{operating_start:02d}:00",
        "가동 종료": f"{operating_end:02d}:00",
        "일 운영시간(h)": len(active_hours),
        "연 가동일수(일)": active_days_per_year,
        "공휴일 비가동 반영": "예" if holiday_reflect else "아니오",
        "연 공휴일 수": int(holiday_count),
    }
])

scenario_rows = []
for tap in range(1, 6):
    delta = max(int(current_tap - tap), 0)
    vdrop = calc_tap_voltage_change(tap_step_percent, delta)
    new_voltage = current_voltage_for_calc * (1 - vdrop / 100.0)

    scenario_profile = calc_profile_based_day(
        season=season,
        loads_by_label=loads_by_label,
        voltage_drop_pct=vdrop,
        applied_cvrf=applied_cvrf,
        z=z,
        i=i,
        p=p,
        rate_dict=active_rate_table,
        operating_hour_count=operating_hour_count,
    )

    scenario_rows.append({
        "탭": tap,
        "계산 기준 전압(V)": round(new_voltage, 1),
        "전압 저감률(%)": round(vdrop, 3),
        "절감률(%)": round(scenario_profile["saving_rate_total"], 3),
        "평균 절감전력(kW)": round(scenario_profile["daily_avg_saved_kw"], 3),
        "일 절감량(kWh)": round(scenario_profile["daily_saved_kwh"], 3),
        "연 절감량(kWh)": round(scenario_profile["daily_saved_kwh"] * active_days_per_year, 1),
        "일 절감요금(원)": round(scenario_profile["daily_cost_saving"], 0),
        "연 절감요금(원)": round(scenario_profile["daily_cost_saving"] * active_days_per_year, 0),
    })

scenario_df = pd.DataFrame(scenario_rows)

current_row = scenario_df[scenario_df["탭"] == current_tap]
target_row = scenario_df[scenario_df["탭"] == target_tap]

if not current_row.empty and not target_row.empty:
    current_voltage_val = float(current_row["계산 기준 전압(V)"].iloc[0])
    target_voltage_val = float(target_row["계산 기준 전압(V)"].iloc[0])
    current_kw_save = float(current_row["평균 절감전력(kW)"].iloc[0])
    target_kw_save = float(target_row["평균 절감전력(kW)"].iloc[0])
    current_cost_year = float(current_row["연 절감요금(원)"].iloc[0])
    target_cost_year = float(target_row["연 절감요금(원)"].iloc[0])

    compare_summary_df = pd.DataFrame([
        {
            "구분": "계산 기준 전압",
            "현재": f"{current_voltage_val / 1000:.2f} kV",
            "제안": f"{target_voltage_val / 1000:.2f} kV",
            "기대 효과": f"{(current_voltage_val - target_voltage_val) / 1000:.2f} kV 저감",
        },
        {
            "구분": "평균 절감전력",
            "현재": f"{current_kw_save:,.1f} kW",
            "제안": f"{target_kw_save:,.1f} kW",
            "기대 효과": f"{(target_kw_save - current_kw_save):,.1f} kW 증가",
        },
        {
            "구분": "연 절감 비용",
            "현재": f"{current_cost_year:,.0f} 원",
            "제안": f"{target_cost_year:,.0f} 원",
            "기대 효과": f"{(target_cost_year - current_cost_year):,.0f} 원 증가",
        },
    ])
else:
    compare_summary_df = pd.DataFrame([
        {"구분": "안내", "현재": "-", "제안": "-", "기대 효과": "현재 탭과 변경 탭이 1~5 범위 내에 있어야 비교됩니다."}
    ])


# PDF에서 사용할 전역값
current_tap_global = current_tap
target_tap_global = target_tap
tap_step_percent_global = tap_step_percent


# -------------------------------------------------
# 그래프
# -------------------------------------------------
chart_df = hourly_detail_df.copy()
chart_df["시간표시"] = chart_df["시간"]

fig_load = px.line(
    chart_df,
    x="시간번호",
    y="전력사용량(kW)",
    markers=True,
    custom_data=["시간표시", "구분", "운영여부", "전력사용량(kW)", "절감전력(kW)", "절감요금(원)"],
)

fig_load.update_traces(
    mode="lines+markers",
    hovertemplate=(
        "<b>%{customdata[0]}</b><br>"
        "구분: %{customdata[1]}<br>"
        "운영여부: %{customdata[2]}<br>"
        "전력사용량: %{customdata[3]:,.3f} kW<br>"
        "절감전력: %{customdata[4]:,.3f} kW<br>"
        "절감요금: %{customdata[5]:,.1f} 원"
        "<extra></extra>"
    )
)
fig_load.update_layout(
    title=f"{season} 시간대별 전력사용량",
    xaxis_title="시간",
    yaxis_title="전력사용량 (kW)",
    hovermode="x unified",
)
fig_load.update_xaxes(
    tickmode="array",
    tickvals=list(range(24)),
    ticktext=[f"{i:02d}" for i in range(24)],
)

scenario_chart_df = scenario_df.copy()
scenario_chart_df["탭문자"] = scenario_chart_df["탭"].astype(str)

fig_bar = px.bar(
    scenario_chart_df,
    x="탭문자",
    y="평균 절감전력(kW)",
    custom_data=["탭", "계산 기준 전압(V)", "전압 저감률(%)", "절감률(%)", "일 절감량(kWh)", "연 절감요금(원)"],
)

fig_bar.update_traces(
    hovertemplate=(
        "<b>탭 %{customdata[0]}</b><br>"
        "계산 기준 전압: %{customdata[1]:,.1f} V<br>"
        "전압 저감률: %{customdata[2]:,.3f} %<br>"
        "절감률: %{customdata[3]:,.3f} %<br>"
        "일 절감량: %{customdata[4]:,.3f} kWh<br>"
        "연 절감요금: %{customdata[5]:,.0f} 원"
        "<extra></extra>"
    )
)
fig_bar.update_layout(
    title=f"{site_name} 탭 변화별 예상 절감전력",
    xaxis_title="탭",
    yaxis_title="평균 절감전력 (kW)",
)


# -------------------------------------------------
# 저장 파일 생성
# -------------------------------------------------
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
safe_site_name = site_name.replace(" ", "_") if site_name else "site"

excel_file = build_excel_file(
    compare_df=compare_df,
    input_profile_df=input_profile_df,
    operation_summary_df=operation_summary_df,
    selected_rate_df=selected_rate_df,
    selected_season_df=selected_season_df,
    period_detail_df=period_detail_df,
    hourly_detail_df=hourly_detail_df,
    scenario_df=scenario_df,
    compare_summary_df=compare_summary_df,
)

pdf_file = build_pdf_file(
    site_name=site_name,
    season=season,
    voltage_basis_text=voltage_basis_text,
    primary_voltage_kv=primary_voltage_kv,
    secondary_voltage_v=secondary_voltage_v,
    est_new_voltage=est_new_voltage,
    voltage_drop_v=voltage_drop_v,
    primary_voltage_class=primary_voltage_class,
    tariff_voltage_final=tariff_voltage_final,
    tariff_category=tariff_category,
    tariff_type=tariff_type,
    tariff_option=tariff_option,
    applied_cvrf=applied_cvrf,
    z=z,
    i=i,
    p=p,
    calibration=calibration,
    confidence=confidence,
    voltage_drop_pct=voltage_drop_pct,
    final_saved_kw=final_saved_kw,
    final_saving_rate=final_saving_rate,
    final_saved_kwh_day=final_saved_kwh_day,
    final_saved_kwh_month=final_saved_kwh_month,
    final_saved_kwh_year=final_saved_kwh_year,
    daily_cost_saving=daily_cost_saving,
    monthly_cost_saving=monthly_cost_saving,
    yearly_cost_saving=yearly_cost_saving,
    compare_df=compare_df,
    compare_summary_df=compare_summary_df,
    operation_summary_df=operation_summary_df,
    selected_rate_df=selected_rate_df,
    selected_season_df=selected_season_df,
    input_profile_df=input_profile_df,
    period_detail_df=period_detail_df,
    hourly_detail_df=hourly_detail_df,
    scenario_df=scenario_df,
    fig_load=fig_load,
    fig_bar=fig_bar,
)


# -------------------------------------------------
# 오른쪽 결과
# -------------------------------------------------
with right:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=220)

    st.subheader("계산 결과")

    c1, c2, c3 = st.columns(3)
    c1.metric("예상 전압 저감률", f"{voltage_drop_pct:.2f}%")
    c2.metric("예상 절감 전력", f"{final_saved_kw:,.1f} kW")
    c3.metric("예상 절감률", f"{final_saving_rate:.2f}%")

    c4, c5, c6 = st.columns(3)
    c4.metric("예상 일 절감량", f"{final_saved_kwh_day:,.1f} kWh")
    c5.metric("예상 월 절감량", f"{final_saved_kwh_month:,.1f} kWh")
    c6.metric("예상 연 절감량", f"{final_saved_kwh_year:,.1f} kWh")

    st.subheader("결과 요약")
    st.markdown(
        f"""
- 선택 계절: **{season}**
- 계산 기준 전압: **{voltage_basis_text}**
- 1차측 수전전압: **{primary_voltage_kv:,.1f} kV**
- 2차측 평균전압: **{secondary_voltage_v:,.1f} V**
- 변경 후 계산 기준 전압: **{est_new_voltage:,.1f} V**
- 자동 판정 전압등급: **{primary_voltage_class}**
- 요금 적용 전압등급: **{tariff_voltage_final}**
- 요금 종별: **{tariff_category} {tariff_type} / {tariff_option}**
- 적용 CVRf: **{applied_cvrf:.2f}**
- 적용 ZIP 비율: **Z {z:.2f} / I {i:.2f} / P {p:.2f}**
- 보정계수: **{calibration:.2f}**
- 예상 신뢰도: **{confidence}**
"""
    )

    st.subheader("전기요금 절감 효과")
    d1, d2, d3 = st.columns(3)
    d1.metric("예상 일 절감 요금", f"{daily_cost_saving:,.0f} 원")
    d2.metric("예상 월 절감 요금", f"{monthly_cost_saving:,.0f} 원")
    d3.metric("예상 연 절감 요금", f"{yearly_cost_saving:,.0f} 원")

    st.markdown("### 현재안 / 제안안 비교")
    st.dataframe(compare_summary_df, use_container_width=True, hide_index=True)


# -------------------------------------------------
# 결과 저장
# -------------------------------------------------
st.divider()
st.subheader("결과 저장")

save_col1, save_col2 = st.columns(2)

with save_col1:
    st.download_button(
        label="PDF 저장",
        data=pdf_file,
        file_name=f"CVR보고서_{safe_site_name}_{timestamp}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

with save_col2:
    st.download_button(
        label="엑셀 저장",
        data=excel_file,
        file_name=f"CVR결과_{safe_site_name}_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


# -------------------------------------------------
# 그래프
# -------------------------------------------------
st.divider()
st.subheader("1. 부하 및 탭별 절감전력 시각화")

chart_left, chart_right = st.columns(2)
with chart_left:
    st.markdown("#### 시간대별 전력사용량 그래프")
    st.plotly_chart(fig_load, use_container_width=True)

with chart_right:
    st.markdown("#### 탭별 예상 절감전력 비교")
    st.plotly_chart(fig_bar, use_container_width=True)


# -------------------------------------------------
# 결과 요약표
# -------------------------------------------------
st.divider()
st.subheader("2. 계산 결과 요약표")
st.dataframe(compare_df, use_container_width=True, hide_index=True)


# -------------------------------------------------
# 상세 데이터
# -------------------------------------------------
st.divider()
st.subheader("3. 상세 데이터")

with st.expander("3-1. 계절별 시간대 구분표", expanded=False):
    st.dataframe(selected_season_df, use_container_width=True, hide_index=True)

with st.expander("3-2. 적용 요금표", expanded=False):
    st.dataframe(selected_rate_df, use_container_width=True, hide_index=True)

with st.expander("3-3. 운영 조건 요약", expanded=False):
    st.dataframe(operation_summary_df, use_container_width=True, hide_index=True)

with st.expander("3-4. 시간대별 평균 부하 입력 요약", expanded=False):
    st.dataframe(input_profile_df, use_container_width=True, hide_index=True)

with st.expander("3-5. 시간대 구간별 계산 결과", expanded=False):
    st.dataframe(period_detail_df, use_container_width=True, hide_index=True)

with st.expander("3-6. 24시간 시간대별 상세 결과", expanded=False):
    st.dataframe(hourly_detail_df, use_container_width=True, hide_index=True)

with st.expander("3-7. 탭별 비교표", expanded=False):
    st.dataframe(scenario_df, use_container_width=True, hide_index=True)


# -------------------------------------------------
# 활용 가이드
# -------------------------------------------------
st.divider()
st.subheader("4. 활용 가이드")
st.markdown(
    """
- 본 화면은 **실사용 / 사전 검토 / 미팅용 운영형 계산기**입니다.
- **절감전력량(kWh)** 은 전압 저감률, 부하 특성, 운영시간에 의해 결정됩니다.
- **절감요금(원)** 은 절감전력량에 **적용 단가(원/kWh)** 를 곱해서 계산합니다.
- 요금 절감액은 현재 선택한 **요금 종별 / 갑·을 / 전압구분 / 선택요금제**에 따라 달라집니다.
- 자동 요금표는 **예시 구조용**입니다. 실제 적용 시에는 **최신 계약요금표** 기준으로 수정해야 합니다.
- 최종 적용 전에는 **ETAP 상세 모델링**, **말단 전압 검토**, **실측 데이터 검증**이 필요합니다.
"""
)
