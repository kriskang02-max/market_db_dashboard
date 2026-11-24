# -*- coding: utf-8 -*-
"""매칭형 레포 주간 제안서 PPTX 생성.

matching_data.csv + 수기 입력(JSON)으로 템플릿을 채운다.
"""
from __future__ import annotations

import csv
import io
import math
import os
import re
from collections import defaultdict
from datetime import datetime, timedelta
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.util import Pt

FONT_DATA = "원신한 Light"
# 01·02 변동폭 텍스트: 상승=레드, 하락=블루 (배경 히트맵 대신)
COLOR_CHG_UP = RGBColor(0xDC, 0x26, 0x26)
COLOR_CHG_DOWN = RGBColor(0x25, 0x63, 0xEB)
COLOR_CHG_FLAT = RGBColor(0x64, 0x74, 0x8B)
COLOR_LEVEL = RGBColor(0x1E, 0x29, 0x3B)  # 절대금리/본문
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_HEADER_GRAY = RGBColor(0xF1, 0xF5, 0xF9)
COLOR_SLATE_BG = RGBColor(0xF8, 0xFA, 0xFC)
COLOR_SHINHAN_BG = RGBColor(0xEF, 0xF6, 0xFF)  # 연한 신한 블루
COLOR_SHINHAN_NAVY = RGBColor(0x00, 0x38, 0x76)
COLOR_TEAL = RGBColor(0x0F, 0x76, 0x6E)
COLOR_SKY = RGBColor(0x38, 0xBD, 0xF8)
COLOR_GOLD = RGBColor(0xFD, 0xE0, 0x47)
COLOR_DEFAULT = COLOR_LEVEL
# 하위 호환 별칭
COLOR_UP = COLOR_CHG_UP
COLOR_DOWN = COLOR_CHG_DOWN

ROOT = os.path.dirname(os.path.abspath(__file__))
DEFAULT_CSV = os.path.join(ROOT, "matching_data.csv")
DEFAULT_TEMPLATE = os.path.join(ROOT, "신한 매칭형 레포 채권펀드 주간 투자 제안서.pptx")

# PPTX 표 인덱스 (슬라이드 0 shapes) — 템플릿 편집 시 resolve_slide_shapes()가 내용으로 재탐색
IDX_REPO_SCENARIO = 71
IDX_PORTFOLIO_DETAIL = 72
IDX_MINPYEONG = 73
IDX_COND_100 = 127  # 레버리지 100% (05)
IDX_COND_200 = 121  # 레버리지 200% (05)
IDX_EXPECTED = 122
IDX_CREDIT_SP = 123
IDX_CURVE_SP = 124
IDX_SCENARIO_RET = 125
IDX_RATE_SENS = 126

# 텍스트 박스
IDX_DATE_LINE = 84
IDX_FUND_STRATEGY = 94
IDX_FUND_SECTOR = 96
IDX_FUND_LEV = 98
IDX_BASE_RATE = 100
IDX_REPO_SPOT = 102
IDX_SUPPLY_BANK = 79
IDX_SUPPLY_NBFC = 80
IDX_SUPPLY_SPECIAL = 128  # 특수채 수급 (템플릿에 존재)
# RV 코멘트: 제목 아래 넓은 상자. 라벨 옆 짧은 칸은 비움
IDX_RV_CURVE = 103
IDX_RV_CREDIT = 104
IDX_RV_CASH = 105
IDX_RV_CURVE_TAG = 108
IDX_RV_CREDIT_TAG = 110
IDX_RV_CASH_TAG = 112
SUPPLY_SPECIAL_SHAPE_NAME = "MatchingSupplySpecial"
RV_CREDIT_MIN_HEIGHT_EMU = 210000  # 크레딧 RV 2줄
COLOR_SUPPLY = RGBColor(0xCB, 0xD5, 0xE1)  # 사이드바 수급 코멘트
SUPPLY_FONT_PT = 6.75
IDX_TITLE_03 = 81
IDX_FOOTNOTE_GROUP = 85
IDX_REC_TITLE = 87
IDX_REC_COMPOSITION = 82
IDX_REC_RETURN = 88
IDX_REC_BEP = 89
IDX_PORT_SECTION_TITLE = 106  # 06 금주 추천 포트폴리오 상세
IDX_PORT_SECTION_CARD = 8  # 06 영역 카드 배경
IDX_NOTE_MINPYEONG = 114
IDX_NOTE_EXPECTED = 116
IDX_NOTE_CREDIT = 118
IDX_NOTE_CURVE = 120
FOOTNOTE_NUDGE_EMU = 70000  # 01·02 각주
FOOTNOTE_NUDGE_EMU_BOTTOM = 55000  # 03·04 각주
COLOR_ACCENT = COLOR_SHINHAN_NAVY
COLOR_PORT_TOTAL_FILL = COLOR_SHINHAN_BG
COLOR_PORT_RATE_FILL = COLOR_SHINHAN_BG
COLOR_SCEN_SELECT_FILL = COLOR_SHINHAN_BG
# 05 Repo 표: 템플릿이 밝은 글씨 → 어두운 배경 유지
COLOR_REPO_HEADER = RGBColor(0x1E, 0x29, 0x3B)
COLOR_REPO_ROW = RGBColor(0x0F, 0x17, 0x2A)
COLOR_REPO_ROW_ALT = RGBColor(0x1E, 0x29, 0x3B)
COLOR_REPO_AVG = RGBColor(0x0C, 0x4A, 0x6E)  # 1Y평균 열
IDX_SIDE_TITLES = (77, 78, 81, 86, 90)  # 01~05 사이드바 소제목
IDX_REC_CARD = 2
IDX_REC_GOLD_BAR = 6  # 04 카드 좌측 포인트 바

SECTOR_ROWS = [
    ("국고채", "국고"),
    ("특은채 AAA", "특은"),
    ("특수/시은채 AAA", "특수시은"),
    ("카드채 AA", "카드"),
    ("캐피탈채 AA-", "캐피탈"),
]
TENORS = ["1년", "1.5년", "2년", "2.5년", "3년"]
TENOR_FIELDS = {
    "6개월": "6월이하(당일)",
    "1년": "1년이하(당일)",
    "1.5년": "1.5년이하(당일)",
    "2년": "2년이하(당일)",
    "2.5년": "2.5년이하(당일)",
    "3년": "3년이하(당일)",
}
PREV_TENOR = {
    "1년": "6개월",
    "1.5년": "1년",
    "2년": "1.5년",
    "2.5년": "2년",
    "3년": "2.5년",
}
# 기대수익률용: 1년 전 만기 커브 SP, 롤다운 듀레이션 = 만기−1년
PREV_TENOR_1Y = {
    "1년": None,
    "1.5년": "6개월",
    "2년": "1년",
    "2.5년": "1.5년",
    "3년": "2년",
}
NOTE_EXPECTED = "* 기대수익률 = YTM + 커브SP x (듀레이션-1년) (1년 보유 기준)"
NOTE_MINPYEONG = "* 시가평가 3사 평균 민평금리 기준 (괄호 없이 bp는 전주 대비 변동폭)"
NOTE_CREDIT = "* 동일 만기 국고채 대비 스프레드 및 최근 10개년 백분위"
NOTE_CURVE = "* 커브 SP = 당해 만기 YTM - 직전(-0.5Y) 만기 YTM (1년은 1Y-6M 기준)"
INST_MAP = {
    "국고": "시가평가 3사평균 국고채권",
    "특은": "시가평가 3사평균 산업은행",
    "한전": "시가평가 3사평균 한국전력공사",
    "시은": "시가평가 3사평균 신한은행",
    "카드": "시가평가 3사평균 신한카드",
    "캐피탈": "시가평가 3사평균 신한캐피탈",
    "CD": "CD 3사평균 AAA",
    "REPO": "REPO 종합 1일물",
}

COMBOS_100 = [
    {"tenor": "1년", "label": "특은 100% + 특은 80%", "legs": [("특은", 1.0), ("특은", 0.8)], "rwa": "0%"},
    {"tenor": "1년", "label": "특은 100% + 특수/시은 80%", "legs": [("특은", 1.0), ("특수시은", 0.8)], "rwa": "~20%"},
    {"tenor": "1년", "label": "특수/시은 100% + 여전 80%", "legs": [("특수시은", 1.0), ("여전", 0.8)], "rwa": "20~40%"},
    {"tenor": "1.5년", "label": "특은 100% + 특은 80%", "legs": [("특은", 1.0), ("특은", 0.8)], "rwa": "0%"},
    {"tenor": "1.5년", "label": "특은 100% + 특수/시은 80%", "legs": [("특은", 1.0), ("특수시은", 0.8)], "rwa": "~20%"},
    {"tenor": "1.5년", "label": "특수/시은 100% + 여전 80%", "legs": [("특수시은", 1.0), ("여전", 0.8)], "rwa": "20~40%"},
    {"tenor": "2년", "label": "특은 100% + 특은 80%", "legs": [("특은", 1.0), ("특은", 0.8)], "rwa": "0%"},
    {"tenor": "2년", "label": "특은 100% + 특수/시은 80%", "legs": [("특은", 1.0), ("특수시은", 0.8)], "rwa": "~20%"},
    {"tenor": "2년", "label": "특수/시은 100% + 여전 80%", "legs": [("특수시은", 1.0), ("여전", 0.8)], "rwa": "20~40%"},
]
# 200%: 06 추천과 동일하게 채권 레그 + CD 40% + Repo -200%
COMBOS_200 = [
    {"tenor": "1년", "label": "특은 160% + 특은 100%", "legs": [("특은", 1.6), ("특은", 1.0), ("CD", 0.4)], "rwa": "0%"},
    {"tenor": "1년", "label": "특은 160% + 특수/시은 100%", "legs": [("특은", 1.6), ("특수시은", 1.0), ("CD", 0.4)], "rwa": "~20%"},
    {"tenor": "1년", "label": "특수/시은 160% + 여전 100%", "legs": [("특수시은", 1.6), ("여전", 1.0), ("CD", 0.4)], "rwa": "20~40%"},
    {"tenor": "1.5년", "label": "특은 160% + 특은 100%", "legs": [("특은", 1.6), ("특은", 1.0), ("CD", 0.4)], "rwa": "0%"},
    {"tenor": "1.5년", "label": "특은 160% + 특수/시은 100%", "legs": [("특은", 1.6), ("특수시은", 1.0), ("CD", 0.4)], "rwa": "~20%"},
    {"tenor": "1.5년", "label": "특수/시은 160% + 여전 100%", "legs": [("특수시은", 1.6), ("여전", 1.0), ("CD", 0.4)], "rwa": "20~40%"},
    {"tenor": "2년", "label": "특은 160% + 특은 100%", "legs": [("특은", 1.6), ("특은", 1.0), ("CD", 0.4)], "rwa": "0%"},
    {"tenor": "2년", "label": "특은 160% + 특수/시은 100%", "legs": [("특은", 1.6), ("특수시은", 1.0), ("CD", 0.4)], "rwa": "~20%"},
    {"tenor": "2년", "label": "특수/시은 160% + 여전 100%", "legs": [("특수시은", 1.6), ("여전", 1.0), ("CD", 0.4)], "rwa": "20~40%"},
]


def _num(v: Any) -> float | None:
    if v is None:
        return None
    s = str(v).strip().replace(",", "")
    if not s or s in ("—", "-", "NA", "nan"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _ymd(s: str) -> datetime | None:
    try:
        return datetime.strptime(s, "%Y-%m-%d")
    except Exception:
        return None


def load_series(csv_path: str) -> dict[str, dict[str, dict[str, float]]]:
    """instrument -> field -> date -> value"""
    out: dict[str, dict[str, dict[str, float]]] = defaultdict(lambda: defaultdict(dict))
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            inst = (row.get("instrument") or "").strip()
            field = (row.get("field") or "").strip()
            date = (row.get("date") or "").strip()
            val = _num(row.get("value"))
            if not inst or not field or not date or val is None:
                continue
            out[inst][field][date] = val
    return out


def available_asof_dates(series: dict) -> list[str]:
    dates = set()
    ktb = series.get(INST_MAP["국고"], {})
    for field in TENOR_FIELDS.values():
        dates.update(ktb.get(field, {}).keys())
    return sorted(dates, reverse=True)


def pick_prev_week_date(dates_sorted_asc: list[str], asof: str) -> str | None:
    """전주 동일 요일에 가깝고 asof 이전인 최근 관측일."""
    target = _ymd(asof)
    if not target:
        return None
    want = target - timedelta(days=7)
    best = None
    best_diff = None
    for d in dates_sorted_asc:
        dt = _ymd(d)
        if not dt or dt >= target:
            continue
        diff = abs((dt - want).days)
        if best is None or diff < best_diff or (diff == best_diff and dt > _ymd(best)):
            best = d
            best_diff = diff
    return best


def sector_rate(series: dict, sector_key: str, tenor: str, date: str) -> float | None:
    field = TENOR_FIELDS.get(tenor)
    if not field:
        return None
    if sector_key == "특수시은":
        a = series.get(INST_MAP["한전"], {}).get(field, {}).get(date)
        b = series.get(INST_MAP["시은"], {}).get(field, {}).get(date)
        if a is None or b is None:
            return None
        return (a + b) / 2.0
    if sector_key == "여전":
        a = series.get(INST_MAP["카드"], {}).get(field, {}).get(date)
        b = series.get(INST_MAP["캐피탈"], {}).get(field, {}).get(date)
        if a is None and b is None:
            return None
        if a is None:
            return b
        if b is None:
            return a
        return (a + b) / 2.0
    inst = INST_MAP.get(sector_key)
    if not inst:
        return None
    return series.get(inst, {}).get(field, {}).get(date)


def curve_sp(series: dict, sector_key: str, tenor: str, date: str) -> float | None:
    prev = PREV_TENOR.get(tenor)
    if not prev:
        return None
    a = sector_rate(series, sector_key, tenor, date)
    b = sector_rate(series, sector_key, prev, date)
    if a is None or b is None:
        return None
    return a - b


def expected_return(series: dict, sector_key: str, tenor: str, date: str) -> float | None:
    """YTM + (YTM−YTM_1Y전) × (만기년−1). 1년 이하는 YTM만."""
    ytm = sector_rate(series, sector_key, tenor, date)
    ty = tenor_years(tenor)
    if ytm is None or ty is None:
        return None
    roll_dur = ty - 1.0
    if roll_dur <= 0:
        return ytm
    prev_t = PREV_TENOR_1Y.get(tenor)
    if not prev_t:
        return ytm
    prev_ytm = sector_rate(series, sector_key, prev_t, date)
    if prev_ytm is None:
        return None
    return ytm + (ytm - prev_ytm) * roll_dur


def credit_sp(series: dict, sector_key: str, tenor: str, date: str) -> float | None:
    if sector_key == "국고":
        return 0.0
    a = sector_rate(series, sector_key, tenor, date)
    b = sector_rate(series, "국고", tenor, date)
    if a is None or b is None:
        return None
    return (a - b) * 100.0  # bp


def percentile_rank(values: list[float], current: float) -> float | None:
    if not values or current is None or any(v is None for v in values):
        return None
    clean = [v for v in values if v is not None and math.isfinite(v)]
    if not clean:
        return None
    n = len(clean)
    below = sum(1 for v in clean if v < current)
    equal = sum(1 for v in clean if v == current)
    return (below + 0.5 * equal) / n * 100.0


def history_dates(asof: str, years: float = 10.0) -> tuple[datetime, datetime]:
    end = _ymd(asof) or datetime.today()
    start = end - timedelta(days=int(365.25 * years))
    return start, end


def hist_metric(
    series: dict,
    sector_key: str,
    tenor: str,
    asof: str,
    kind: str,
) -> list[float]:
    start, end = history_dates(asof, 10.0)
    field = TENOR_FIELDS[tenor]
    # gather union of dates from underlying instruments
    dates = set()
    keys = []
    if sector_key == "특수시은":
        keys = [INST_MAP["한전"], INST_MAP["시은"]]
    elif sector_key == "여전":
        keys = [INST_MAP["카드"], INST_MAP["캐피탈"]]
    else:
        keys = [INST_MAP.get(sector_key, "")]
    for inst in keys:
        dates.update(series.get(inst, {}).get(field, {}).keys())
    out = []
    for d in sorted(dates):
        dt = _ymd(d)
        if not dt or dt < start or dt > end:
            continue
        if kind == "credit":
            v = credit_sp(series, sector_key, tenor, d)
        else:
            v = curve_sp(series, sector_key, tenor, d)
            if v is not None:
                v = v * 100.0  # bp
        if v is not None and math.isfinite(v):
            out.append(v)
    return out


def fmt_pct(v: float | None, digits: int = 2) -> str:
    if v is None or not math.isfinite(v):
        return ""
    return f"{v:.{digits}f}%"


def fmt_chg_bp(chg_pct: float | None) -> str:
    """수익률 변동(%p) → 정수 bp 문자열 (+/- 기호, 괄호 없음)."""
    if chg_pct is None or not math.isfinite(chg_pct):
        return ""
    bp = int(round(chg_pct * 100.0))
    if bp > 0:
        return f"+{bp}bp"
    if bp < 0:
        return f"{bp}bp"
    return "0bp"


def fmt_signed_pct(v: float | None, digits: int = 2) -> str:
    """부호 포함 % (동결·BEP 열과 동일 자릿수)."""
    if v is None or not math.isfinite(v):
        return ""
    eps = 0.5 * (10 ** (-digits))
    if abs(v) < eps:
        return f"{0:.{digits}f}%"
    return f"{v:+.{digits}f}%"


def fmt_curve_bp(bp: float | None) -> str:
    """커브 SP bp → 반올림 정수 xbp."""
    if bp is None or not math.isfinite(bp):
        return ""
    n = int(round(bp))
    if n > 0:
        return f"+{n}bp"
    if n < 0:
        return f"{n}bp"
    return "0bp"


def fmt_bp_pct(bp: float | None, pct: float | None) -> str:
    if bp is None or not math.isfinite(bp):
        return ""
    sign = "+" if bp > 0 else ""
    left = f"{sign}{bp:.0f}bp" if abs(bp) >= 10 else f"{sign}{bp:.1f}bp"
    if pct is None or not math.isfinite(pct):
        return left
    return f"{left}\n{pct:.0f}%"


def fmt_curve_cell(bp: float | None, pct: float | None) -> str:
    left = fmt_curve_bp(bp)
    if not left:
        return ""
    if pct is None or not math.isfinite(pct):
        return left
    return f"{left}\n{pct:.0f}%"


def fmt_minpyeong(level: float | None, chg: float | None) -> str:
    if level is None:
        return ""
    if chg is None:
        return fmt_pct(level)
    return f"{fmt_pct(level)}\n{fmt_chg_bp(chg)}"


def sign_color(v: float | None) -> RGBColor | None:
    """레거시: 양수=업컬러, 음수=다운컬러."""
    if v is None or not math.isfinite(v) or v == 0:
        return None
    return COLOR_CHG_UP if v > 0 else COLOR_CHG_DOWN


def chg_text_color(chg_pct: float | None) -> RGBColor:
    """01·02 둘째줄 변동폭: +레드 / -블루 / 0 그레이."""
    if chg_pct is None or not math.isfinite(chg_pct):
        return COLOR_CHG_FLAT
    bp = int(round(chg_pct * 100.0))
    if bp > 0:
        return COLOR_CHG_UP
    if bp < 0:
        return COLOR_CHG_DOWN
    return COLOR_CHG_FLAT


def _lerp_byte(a: int, b: int, t: float) -> int:
    return int(round(a + (b - a) * t))


def heatmap_rgb(pct: float | None) -> RGBColor | None:
    """03·04 백분위 히트맵: 낮음=페일코랄, 중간=흰, 높음=소프트블루."""
    if pct is None or not math.isfinite(pct):
        return None
    t = max(0.0, min(100.0, float(pct))) / 100.0
    coral = (255, 241, 242)  # #FFF1F2
    mid = (255, 255, 255)
    soft_lo = (239, 246, 255)  # #EFF6FF
    soft_hi = (191, 219, 254)  # #BFDBFE
    if t <= 0.5:
        u = t * 2.0
        r = _lerp_byte(coral[0], mid[0], u)
        g = _lerp_byte(coral[1], mid[1], u)
        b = _lerp_byte(coral[2], mid[2], u)
    else:
        u = (t - 0.5) * 2.0
        # 50→70: 흰색→EFF6FF, 70→100: EFF6FF→BFDBFE
        if u <= 0.4:
            uu = u / 0.4
            r = _lerp_byte(mid[0], soft_lo[0], uu)
            g = _lerp_byte(mid[1], soft_lo[1], uu)
            b = _lerp_byte(mid[2], soft_lo[2], uu)
        else:
            uu = (u - 0.4) / 0.6
            r = _lerp_byte(soft_lo[0], soft_hi[0], uu)
            g = _lerp_byte(soft_lo[1], soft_hi[1], uu)
            b = _lerp_byte(soft_lo[2], soft_hi[2], uu)
    return RGBColor(r, g, b)


def rgb_to_hex(c: RGBColor | None, with_hash: bool = True) -> str | None:
    if c is None:
        return None
    try:
        hx = f"{int(c[0]):02X}{int(c[1]):02X}{int(c[2]):02X}"
    except Exception:
        return None
    return f"#{hx}" if with_hash else hx


def set_cell_fill(cell, fill: RGBColor | None) -> None:
    """표 셀 배경색 설정 (히트맵).

    PowerPoint는 parse_xml로 새로 붙인 solidFill을 손상으로 보는 경우가 있어
    기존 solidFill/srgbClr 값을 직접 수정하고, 없을 때만 안전하게 삽입한다.
    """
    if fill is None:
        return
    # hMerge 셀은 건드리면 표가 깨질 수 있음
    try:
        if cell._tc.get("hMerge") in ("1", "true", "True"):
            return
    except Exception:
        pass
    hx = rgb_to_hex(fill, with_hash=False)
    if not hx:
        return
    tcPr = cell._tc.get_or_add_tcPr()
    solid = tcPr.find(qn("a:solidFill"))
    if solid is not None:
        srgb = solid.find(qn("a:srgbClr"))
        if srgb is not None:
            srgb.set("val", hx)
            # schemeClr 등 다른 색 자식 제거
            for child in list(solid):
                if child is not srgb:
                    solid.remove(child)
            return
        for child in list(solid):
            solid.remove(child)
        from lxml.etree import SubElement

        srgb = SubElement(solid, qn("a:srgbClr"))
        srgb.set("val", hx)
        return

    # solidFill이 없으면 ln* 뒤에 삽입 (스키마 순서 유지)
    from lxml.etree import SubElement, Element

    solid = Element(qn("a:solidFill"))
    srgb = SubElement(solid, qn("a:srgbClr"))
    srgb.set("val", hx)
    # remove competing fills
    for tag in ("a:noFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill"):
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
    insert_at = 0
    for i, child in enumerate(list(tcPr)):
        local = child.tag.split("}")[-1]
        if local in ("lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr", "cell3D"):
            insert_at = i + 1
    tcPr.insert(insert_at, solid)


def tenor_years(tenor: str) -> float | None:
    t = (tenor or "").strip()
    if t == "3개월":
        return 0.25
    if t == "6개월":
        return 0.5
    if t.endswith("년"):
        try:
            return float(t.replace("년", ""))
        except ValueError:
            return None
    return None


def matching_duration(tenor: str, weight_sum: float) -> float | None:
    """(만기년 - 0.5) × 비중합(소수). 예: 2년·260% → 1.5×2.6=3.9"""
    ty = tenor_years(tenor)
    if ty is None or weight_sum is None:
        return None
    return (ty - 0.5) * weight_sum


def portfolio_bond_duration(rows_in: list[dict]) -> float | None:
    """채권/CD 레그의 (만기-0.5)×비중 합산. Repo 제외."""
    total = 0.0
    any_leg = False
    for r in rows_in:
        kind = (r.get("kind") or "").lower()
        sector = (r.get("sector") or "").upper()
        if kind == "repo" or sector == "REPO":
            continue
        w = (_num(r.get("weight")) or 0.0) / 100.0
        tenor = r.get("tenor") or ""
        d = matching_duration(tenor, w)
        if d is None:
            continue
        total += d
        any_leg = True
    return total if any_leg else None


def build_grids(series: dict, asof: str) -> dict[str, Any]:
    dates_asc = sorted(
        set().union(
            *[
                series.get(INST_MAP["국고"], {}).get(TENOR_FIELDS[t], {}).keys()
                for t in TENORS
            ]
        )
    )
    prev = pick_prev_week_date(dates_asc, asof)
    minpyeong = {}
    expected = {}
    credit = {}
    curve = {}
    for _label, key in SECTOR_ROWS:
        for tenor in TENORS:
            cur = sector_rate(series, key, tenor, asof)
            prv = sector_rate(series, key, tenor, prev) if prev else None
            chg = (cur - prv) if (cur is not None and prv is not None) else None
            minpyeong[(key, tenor)] = {
                "level": cur,
                "chg": chg,
                "text": fmt_minpyeong(cur, chg),
                # 01: 배경 흰색, 1줄 차콜 볼드 / 2줄 변동폭 색
                "line_colors": [COLOR_LEVEL, chg_text_color(chg)],
                "line_bolds": [True, True],
                "fill": COLOR_WHITE,
            }
            exp = expected_return(series, key, tenor, asof)
            exp_prv = expected_return(series, key, tenor, prev) if prev else None
            exp_chg = (exp - exp_prv) if (exp is not None and exp_prv is not None) else None
            expected[(key, tenor)] = {
                "level": exp,
                "chg": exp_chg,
                "text": fmt_minpyeong(exp, exp_chg),
                "line_colors": [COLOR_LEVEL, chg_text_color(exp_chg)],
                "line_bolds": [True, True],
                "fill": COLOR_WHITE,
            }
            csp = credit_sp(series, key, tenor, asof)
            csp_pct = percentile_rank(hist_metric(series, key, tenor, asof, "credit"), csp) if csp is not None else None
            credit_fill = heatmap_rgb(csp_pct) if key != "국고" else COLOR_WHITE
            credit[(key, tenor)] = {
                "bp": csp,
                "pct": csp_pct,
                "text": fmt_bp_pct(csp, csp_pct) if key != "국고" else "",
                # 03: 히트맵 유지, 텍스트는 전부 차콜
                "line_colors": [COLOR_LEVEL, COLOR_LEVEL] if key != "국고" else [None, None],
                "line_bolds": [True, False],
                "fill": credit_fill,
            }
            cv = curve_sp(series, key, tenor, asof)
            cv_bp = cv * 100.0 if cv is not None else None
            cv_pct = percentile_rank(hist_metric(series, key, tenor, asof, "curve"), cv_bp) if cv_bp is not None else None
            curve[(key, tenor)] = {
                "bp": cv_bp,
                "pct": cv_pct,
                "text": fmt_curve_cell(cv_bp, cv_pct),
                "line_colors": [COLOR_LEVEL, COLOR_LEVEL],
                "line_bolds": [True, False],
                "fill": heatmap_rgb(cv_pct) if cv_pct is not None else COLOR_WHITE,
            }
    return {
        "asof": asof,
        "prev_week": prev,
        "minpyeong": minpyeong,
        "expected": expected,
        "credit": credit,
        "curve": curve,
    }


def default_inputs() -> dict[str, Any]:
    return {
        "asof": "",
        "date_line": "",
        "fund": {
            "strategy": "1~3년 매칭 형 + Repo 전략",
            "sector": "은행 · 카드 · 캐피탈 · CD",
            "leverage": "Repo 매도 0 ~ 200%",
            "base_rate": "3.00%",
            "repo_spot": "3.03% (+3bp)",
        },
        "supply": {
            "bank": "• 은행채: 시중 유동성 및 수요 확대로 수급 양호",
            "special": "• 특수채: 공사채 발행 안정, 스프레드 축소재 우위",
            "nbfc": "• 여전채: 2~3년 카드·캐피탈 매수 우위 기조 지속",
        },
        "rv": {
            "curve": "2.0Y 롤다운 극대화",
            "credit": "특은AAA(+42bp, 99%) 및 캐피탈AA- 10년 분포 상단",
            "cash": "6M AAA CD 40%",
        },
        "repo_scenarios": [
            {"label": "동결", "final_rate": 3.00, "repo": 3.10, "avg_1y": 3.10},
            {"label": "1회 인상", "final_rate": 3.25, "repo": 3.35, "avg_1y": 3.35},
            {"label": "2회 인상", "final_rate": 3.50, "repo": 3.60, "avg_1y": 3.54},
            {"label": "3회 인상", "final_rate": 3.75, "repo": 3.85, "avg_1y": 3.68},
        ],
        # 현재월~+12개월(13칸) 기준금리 경로 → 평균(+10bp)이 Repo 1Y평균에 연동
        "rate_path": {
            "titles": ["동결", "1회 인상", "2회 인상", "3회 인상"],
            "rates": [
                [3.00] * 13,
                [3.25] * 13,
                [3.25, 3.25, 3.25, 3.50, 3.50, 3.50, 3.50, 3.50, 3.50, 3.50, 3.50, 3.50, 3.50],
                [3.25, 3.25, 3.25, 3.50, 3.50, 3.75, 3.75, 3.75, 3.75, 3.75, 3.75, 3.75, 3.75],
            ],
            "repo_spread": 0.10,
        },
        "portfolio": {
            "title": "2.0년 · 차입 200%",
            "composition": "2Y특은 160% + 2Y캐피탈 100% + 6M CD 40%",
            "repo_scenario_idx": 0,
            "rows": [
                {"kind": "채권", "name": "2년 특수은행채 AAA", "sector": "특은", "tenor": "2년", "weight": 160},
                {"kind": "채권", "name": "2년 캐피탈채 AA-", "sector": "캐피탈", "tenor": "2년", "weight": 100},
                {"kind": "CD", "name": "6개월 AAA 은행 CD", "sector": "CD", "tenor": "6개월", "weight": 40, "rate": None},
                {"kind": "Repo", "name": "1일물 Repo (조달)", "sector": "REPO", "tenor": "", "weight": -200, "rate": None},
            ],
        },
        "rate_shocks": [-0.6, -0.4, -0.2, 0.0, 0.2, 0.4, 0.6],
    }


def merge_inputs(user: dict | None) -> dict[str, Any]:
    base = default_inputs()
    if not user:
        return base
    for k, v in user.items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            merged = dict(base[k])
            merged.update(v)
            base[k] = merged
        else:
            base[k] = v
    apply_rate_path_to_scenarios(base)
    return base


REPO_SPREAD_DEFAULT = 0.10


def apply_rate_path_to_scenarios(inputs: dict) -> None:
    """기준금리 경로 평균(+스프레드) → repo_scenarios.avg_1y / final / repo 동기화."""
    path = inputs.get("rate_path") or {}
    titles = path.get("titles") or []
    rates = path.get("rates") or []
    spread = _num(path.get("repo_spread"))
    if spread is None:
        spread = REPO_SPREAD_DEFAULT
    scenarios = list(inputs.get("repo_scenarios") or [])
    while len(scenarios) < 4:
        scenarios.append({"label": "", "final_rate": None, "repo": None, "avg_1y": None})
    for i in range(4):
        sc = dict(scenarios[i] or {})
        if i < len(titles) and titles[i]:
            sc["label"] = str(titles[i]).strip()
        row = rates[i] if i < len(rates) else []
        nums = [_num(x) for x in (row or [])[:13]]
        nums = [x for x in nums if x is not None]
        if nums:
            avg_policy = sum(nums) / len(nums)
            sc["avg_1y"] = round(avg_policy + spread, 4)
            sc["final_rate"] = nums[-1]
            sc["repo"] = round(nums[-1] + spread, 4)
        scenarios[i] = sc
    inputs["repo_scenarios"] = scenarios


def scenario_header_label(label: str) -> str:
    s = (label or "").strip()
    if not s:
        return ""
    if "동결" in s:
        return "동결"
    m = re.search(r"(\d+)\s*회", s)
    if m:
        return f"{m.group(1)}회 인상"
    return s.split("(")[0].strip() or s


def cd_rate(series: dict, tenor_label: str, date: str) -> float | None:
    # CD fields are like 6월(적용일)
    field_map = {
        "1개월": "1월(적용일)",
        "2개월": "2월(적용일)",
        "3개월": "3월(적용일)",
        "4개월": "4월(적용일)",
        "6개월": "6월(적용일)",
        "9개월": "9월(적용일)",
        "1년": "1년(적용일)",
    }
    field = field_map.get(tenor_label)
    if not field:
        return None
    return series.get(INST_MAP["CD"], {}).get(field, {}).get(date)


def repo_spot(series: dict, date: str) -> float | None:
    # prefer 가중평균수익률, fallback 시가; walk back if missing
    dates = sorted(series.get(INST_MAP["REPO"], {}).get("가중평균수익률", {}).keys())
    candidates = [d for d in dates if d <= date]
    for d in reversed(candidates[-10:]):
        v = series.get(INST_MAP["REPO"], {}).get("가중평균수익률", {}).get(d)
        if v is not None:
            return v
        v = series.get(INST_MAP["REPO"], {}).get("시가", {}).get(d)
        if v is not None:
            return v
    return None


def _combo_asset_er(
    grids: dict,
    combo: dict,
    series: dict | None,
    asof: str,
) -> tuple[float | None, float, float]:
    """조합의 (자산기대ER, asset비중합, bond비중합). 실패 시 (None, …)."""
    minpyeong = grids.get("minpyeong") or {}
    tenor = combo["tenor"]
    asset = 0.0
    bond_w = 0.0
    er = 0.0
    for sector, w in combo["legs"]:
        if sector == "CD":
            rate = cd_rate(series or {}, "6개월", asof) if series else None
            if rate is None:
                return None, 0.0, 0.0
            asset += w
            er += w * rate
            continue
        key = "캐피탈" if sector == "여전" else sector
        cell = minpyeong.get((key, tenor))
        if not cell or cell["level"] is None:
            return None, 0.0, 0.0
        asset += w
        bond_w += w
        er += w * cell["level"]
    return er, asset, bond_w


def combo_returns(
    grids: dict,
    combos: list[dict],
    scenarios: list[dict],
    series: dict | None = None,
    asof: str | None = None,
    prev_grids: dict | None = None,
) -> list[dict]:
    """05 조건별 수익. 채권은 01 민평, CD는 시장 CD (06과 동일 기준).

    wow: 동결 시나리오 기대수익률의 전주 대비 변동(부호 포함 %, 예: -0.04%).
    """
    asof = asof or grids.get("asof") or ""
    prev_asof = (prev_grids or {}).get("asof") or ""
    freeze_repo = None
    if scenarios:
        freeze_repo = _num(scenarios[0].get("avg_1y"))
        if freeze_repo is None:
            freeze_repo = _num(scenarios[0].get("repo"))
    rows = []
    for combo in combos:
        tenor = combo["tenor"]
        er, asset, bond_w = _combo_asset_er(grids, combo, series, asof)
        if er is None:
            rows.append(
                {
                    "tenor": tenor,
                    "label": combo["label"],
                    "rwa": combo["rwa"],
                    "vals": ["", "", ""],
                    "bep": "",
                    "wow": "",
                }
            )
            continue
        repo_w = 1.0 - asset  # negative
        vals = []
        freeze_ret = None
        for i, sc in enumerate(scenarios[:3]):
            repo = _num(sc.get("avg_1y"))
            if repo is None:
                repo = _num(sc.get("repo"))
            if repo is None:
                vals.append("")
                continue
            ret = er + repo_w * repo
            vals.append(fmt_pct(ret))
            if i == 0:
                freeze_ret = ret
                if freeze_repo is None:
                    freeze_repo = repo
        bep = ""
        dur = matching_duration(tenor, bond_w)
        if freeze_ret is not None and dur and dur != 0:
            bep = fmt_pct(freeze_ret / dur)
        wow = ""
        if freeze_ret is not None and prev_grids and prev_asof:
            er_p, asset_p, _bw = _combo_asset_er(prev_grids, combo, series, prev_asof)
            if er_p is not None:
                # 동일 동결 Repo 가정으로 전주 동결 기대수익률 재산출
                repo_for_wow = freeze_repo if freeze_repo is not None else 0.0
                prev_freeze = er_p + (1.0 - asset_p) * repo_for_wow
                wow = fmt_signed_pct(freeze_ret - prev_freeze)
        rows.append(
            {
                "tenor": tenor,
                "label": combo["label"],
                "rwa": combo["rwa"],
                "vals": vals,
                "bep": bep,
                "wow": wow,
            }
        )
    return rows


def portfolio_metrics(series: dict, grids: dict, inputs: dict) -> dict[str, Any]:
    pf = inputs.get("portfolio") or {}
    rows_in = pf.get("rows") or []
    scenarios = inputs.get("repo_scenarios") or []
    try:
        sel_idx = int(pf.get("repo_scenario_idx") if pf.get("repo_scenario_idx") is not None else 0)
    except (TypeError, ValueError):
        sel_idx = 0
    if scenarios:
        sel_idx = max(0, min(sel_idx, len(scenarios) - 1))
    else:
        sel_idx = 0
    sel_sc = scenarios[sel_idx] if scenarios else {}
    # 06 카드/상세 Repo: 선택한 시나리오 1Y평균 (없으면 Repo)
    sel_repo = _num(sel_sc.get("avg_1y"))
    if sel_repo is None:
        sel_repo = _num(sel_sc.get("repo"))
    sel_label = scenario_header_label(sel_sc.get("label") or "") or "동결"
    asof = grids["asof"]
    detail_rows = []
    contrib_sum = 0.0
    weight_sum = 0.0
    for r in rows_in:
        kind = r.get("kind") or ""
        name = r.get("name") or ""
        weight = _num(r.get("weight")) or 0.0
        rate = _num(r.get("rate"))
        sector = r.get("sector") or ""
        tenor = r.get("tenor") or ""
        if rate is None:
            if kind == "Repo" or sector == "REPO":
                rate = sel_repo
            elif kind == "CD" or sector == "CD":
                rate = cd_rate(series, tenor or "6개월", asof)
            else:
                # 06 포트폴리오는 01 민평수익률 기준
                key = sector if sector != "여전" else "캐피탈"
                cell = grids["minpyeong"].get((key, tenor))
                rate = cell["level"] if cell else None
        weight_pct = weight  # already in %
        contrib = (rate * weight_pct / 100.0) if rate is not None else None
        if contrib is not None:
            contrib_sum += contrib
        weight_sum += weight_pct
        detail_rows.append(
            {
                "kind": kind,
                "name": name,
                "sector": sector,
                "tenor": tenor,
                "rate": rate,
                "weight": weight_pct,
                "contrib": contrib,
                "rate_text": fmt_pct(rate) if rate is not None else "[     ]%",
                "weight_text": f"{weight_pct:.0f}%",
                "contrib_text": fmt_pct(contrib) if contrib is not None else "[     ]%",
            }
        )
    total = contrib_sum
    port_dur = portfolio_bond_duration(rows_in)
    bep_pct = (total / port_dur) if (total is not None and port_dur) else None

    # scenario returns for selected tenor buckets using portfolio bonds only
    scen_headers = []
    for sc in scenarios[:4]:
        short = scenario_header_label(sc.get("label") or "")
        fr = _num(sc.get("final_rate"))
        # 헤더 Repo 줄 = 1Y 평균(포트폴리오 시나리오 수익 산출에 사용)
        avg = _num(sc.get("avg_1y"))
        if avg is None:
            avg = _num(sc.get("repo"))
        scen_headers.append(
            f"{short} ({fmt_pct(fr) if fr is not None else ''})\n"
            f"Repo {fmt_pct(avg) if avg is not None else ''}"
        )
    scen_vals = []
    for sc in scenarios[:4]:
        repo = _num(sc.get("avg_1y"))
        if repo is None:
            scen_vals.append("")
            continue
        s = 0.0
        ok = True
        for r in rows_in:
            w = (_num(r.get("weight")) or 0.0) / 100.0
            kind = r.get("kind") or ""
            sector = r.get("sector") or ""
            tenor = r.get("tenor") or ""
            if kind == "Repo" or sector == "REPO":
                s += w * repo
            elif kind == "CD" or sector == "CD":
                rate = cd_rate(series, tenor or "6개월", asof)
                if rate is None:
                    ok = False
                    break
                s += w * rate
            else:
                key = sector if sector != "여전" else "캐피탈"
                cell = grids["minpyeong"].get((key, tenor))
                if not cell or cell["level"] is None:
                    ok = False
                    break
                s += w * cell["level"]
        scen_vals.append(fmt_pct(s) if ok else "")

    # 시중금리 민감도: 합계 + (-shock)×듀레이션
    # 예) 듀레이션 3.9, shock -0.1% → 6.28 + 0.1×3.9 = 6.67
    shocks = inputs.get("rate_shocks") or [-0.6, -0.4, -0.2, 0.0, 0.2, 0.4, 0.6]
    sens = []
    base = total
    dur = port_dur or 0.0
    for sh in shocks:
        if base is None:
            sens.append("")
            continue
        sens.append(fmt_pct(base - sh * dur))

    if total is not None:
        return_text = f"예상수익률 {fmt_pct(total)} ({sel_label} 기준)"
    else:
        return_text = f"예상수익률 [X.XX]% ({sel_label} 기준)"

    return {
        "detail_rows": detail_rows,
        "total": total,
        "total_text": fmt_pct(total),
        "weight_sum": weight_sum,
        "duration": port_dur,
        "bep_pct": bep_pct,
        "bep_text": fmt_pct(bep_pct) if bep_pct is not None else "[X.XX]%",
        "return_text": return_text,
        "repo_scenario_idx": sel_idx,
        "repo_scenario_label": sel_label,
        "scen_headers": scen_headers,
        "scen_vals": scen_vals,
        "sens_vals": sens,
        "shocks": shocks,
    }


def set_shape_text(shape, text: str, align: str | None = None) -> None:
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    lines = str(text).split("\n")
    # reuse existing paragraphs/runs to keep style
    while len(tf.paragraphs) < len(lines):
        tf.add_paragraph()
    align_map = {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT,
    }
    para_align = align_map.get((align or "").lower())
    for i, para in enumerate(tf.paragraphs):
        line = lines[i] if i < len(lines) else ""
        if para_align is not None:
            para.alignment = para_align
        if para.runs:
            para.runs[0].text = line
            for r in para.runs[1:]:
                r.text = ""
        else:
            run = para.add_run()
            run.text = line


def _ensure_para_runs(para, n: int):
    while len(para.runs) < n:
        para.add_run()
    return list(para.runs)


def _set_rec_metric_text(
    shape,
    prefix: str,
    value: str,
    suffix: str,
    value_color: RGBColor,
    *,
    suffix_size_pt: float | None = 6.8,
) -> None:
    """04 카드 지표: 라벨은 연회색, 핵심 숫자는 강조색, 기준 문구는 소형."""
    if not shape.has_text_frame:
        return
    para = shape.text_frame.paragraphs[0]
    runs = _ensure_para_runs(para, 3)
    parts = [
        (prefix, RGBColor(0xCB, 0xD5, 0xE1), False, None),
        (value, value_color, True, None),
        (suffix, RGBColor(0xCB, 0xD5, 0xE1), False, suffix_size_pt),
    ]
    for i, (txt, col, bold, size) in enumerate(parts):
        runs[i].text = txt
        _apply_run_style(runs[i], col, FONT_DATA, font_size_pt=size, bold=bold)
    for r in runs[3:]:
        r.text = ""


def _style_sidebar_title(shape) -> None:
    """01~05 소제목: 하늘색 볼드, 밑줄 제거."""
    if not shape.has_text_frame:
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            try:
                run.font.underline = False
            except Exception:
                pass
            try:
                rPr = run._r.get_or_add_rPr()
                for child in list(rPr):
                    if child.tag == qn("a:u"):
                        rPr.remove(child)
                if "u" in rPr.attrib:
                    del rPr.attrib["u"]
            except Exception:
                pass
            _apply_run_style(run, COLOR_SKY, FONT_DATA, font_size_pt=None, bold=True)


def _set_shape_solid_fill(shape, color: RGBColor) -> None:
    hx = rgb_to_hex(color, with_hash=False)
    if not hx:
        return
    spPr = shape._element.spPr
    solid = spPr.find(qn("a:solidFill"))
    if solid is None:
        from lxml.etree import Element, SubElement

        solid = Element(qn("a:solidFill"))
        srgb = SubElement(solid, qn("a:srgbClr"))
        srgb.set("val", hx)
        spPr.insert(0, solid)
        return
    srgb = solid.find(qn("a:srgbClr"))
    if srgb is None:
        from lxml.etree import SubElement

        for child in list(solid):
            solid.remove(child)
        srgb = SubElement(solid, qn("a:srgbClr"))
    srgb.set("val", hx)


def _apply_run_style(
    run,
    color: RGBColor | None = None,
    font_name: str = FONT_DATA,
    font_size_pt: float | None = 6,
    bold: bool | None = None,
) -> None:
    try:
        run.font.name = font_name
    except Exception:
        pass
    # 한글 렌더링용 EA/CS 타입페이스도 동일 지정
    try:
        rPr = run._r.get_or_add_rPr()
        for tag in ("a:latin", "a:ea", "a:cs"):
            el = rPr.find(qn(tag))
            if el is None:
                from lxml.etree import SubElement

                el = SubElement(rPr, qn(tag))
            el.set("typeface", font_name)
    except Exception:
        pass
    if font_size_pt is not None:
        try:
            run.font.size = Pt(font_size_pt)
        except Exception:
            pass
    if bold is not None:
        try:
            run.font.bold = bold
        except Exception:
            pass
    if color is not None:
        try:
            run.font.color.rgb = color
        except Exception:
            pass


def set_cell_text(
    cell,
    text: str,
    center: bool = True,
    line_colors: list | None = None,
    font_name: str = FONT_DATA,
    font_size_pt: float | None = 6,
    bold: bool | None = None,
    color: RGBColor | None = None,
    line_bolds: list | None = None,
) -> None:
    lines = str(text).split("\n") if text is not None else [""]
    tf = cell.text_frame
    while len(tf.paragraphs) < len(lines):
        tf.add_paragraph()
    for i, para in enumerate(tf.paragraphs):
        line = lines[i] if i < len(lines) else ""
        line_color = color
        if line_colors and i < len(line_colors) and line_colors[i] is not None:
            line_color = line_colors[i]
        line_bold = bold
        if line_bolds and i < len(line_bolds) and line_bolds[i] is not None:
            line_bold = line_bolds[i]
        if center:
            para.alignment = PP_ALIGN.CENTER
        if para.runs:
            para.runs[0].text = line
            _apply_run_style(para.runs[0], line_color, font_name, font_size_pt, line_bold)
            for r in para.runs[1:]:
                r.text = ""
        else:
            run = para.add_run()
            run.text = line
            _apply_run_style(run, line_color, font_name, font_size_pt, line_bold)


def emphasize_shape_text(
    shape,
    *,
    font_size_pt: float | None = None,
    bold: bool | None = True,
    color: RGBColor | None = None,
) -> None:
    if not shape.has_text_frame:
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if font_size_pt is not None:
                try:
                    run.font.size = Pt(font_size_pt)
                except Exception:
                    pass
            if bold is not None:
                try:
                    run.font.bold = bold
                except Exception:
                    pass
            if color is not None:
                try:
                    run.font.color.rgb = color
                except Exception:
                    pass


def emphasize_shape_line(shape, color: RGBColor, width_pt: float = 2.25) -> None:
    """카드/도형 테두리 강조."""
    try:
        shape.line.color.rgb = color
        shape.line.width = Pt(width_pt)
    except Exception:
        pass


def fill_grid_table(table, grid: dict, value_key: str = "text") -> None:
    # 헤더 행: 쿨그레이
    try:
        for ci in range(len(table.columns)):
            set_cell_fill(table.cell(0, ci), COLOR_HEADER_GRAY)
    except Exception:
        pass
    for ri, (_label, key) in enumerate(SECTOR_ROWS, start=1):
        # 섹터 라벨 열도 흰 유지
        try:
            set_cell_fill(table.cell(ri, 0), COLOR_WHITE)
        except Exception:
            pass
        for ci, tenor in enumerate(TENORS, start=1):
            cell = grid.get((key, tenor), {})
            tc = table.cell(ri, ci)
            set_cell_text(
                tc,
                cell.get(value_key, "") or "",
                line_colors=cell.get("line_colors"),
                line_bolds=cell.get("line_bolds"),
            )
            set_cell_fill(tc, cell.get("fill") if cell.get("fill") is not None else COLOR_WHITE)


def match_combo_row_index(combos: list[dict], portfolio_rows: list[dict]) -> int | None:
    """06 포트폴리오 채권 레그와 동일한 05 조합 행 인덱스."""
    legs: list[tuple[str, float]] = []
    tenor = None
    for r in portfolio_rows or []:
        sector = r.get("sector") or ""
        kind = r.get("kind") or ""
        if sector == "REPO" or kind == "Repo" or sector == "CD" or kind == "CD":
            continue
        key = "캐피탈" if sector == "여전" else sector
        w = (_num(r.get("weight")) or 0.0) / 100.0
        legs.append((key, round(w, 4)))
        if r.get("tenor"):
            tenor = r.get("tenor")
    if not legs or not tenor:
        return None
    legs_s = sorted(legs)
    for i, combo in enumerate(combos):
        if combo.get("tenor") != tenor:
            continue
        c_legs = sorted(
            [
                ("캐피탈" if s == "여전" else s, round(float(w), 4))
                for s, w in (combo.get("legs") or [])
                if s != "CD"
            ]
        )
        if c_legs == legs_s:
            return i
    return None


def _cond_col_ratios(n: int = 8) -> list[float]:
    """만기,구성,RWA,동결,인상1,인상2,BEP,전주대비 열 비율.

    구성·RWA는 줄바꿈 방지로 넓히고, BEP·전주대비에서 확보.
    """
    # 구성(레버리지) 열을 넓혀 200% 표 라벨 줄바꿈 방지
    ratios = [0.08, 0.30, 0.10, 0.10, 0.10, 0.10, 0.10, 0.12]
    return ratios[:n]


def _wow_value_color(wow_text: str) -> RGBColor:
    """전주대비: +파란색, −빨간색, 그 외 본문색."""
    s = (wow_text or "").strip().replace(",", "").replace("%", "").replace("−", "-").replace("–", "-")
    if not s or s in ("—", "-", "–"):
        return COLOR_LEVEL
    try:
        v = float(s)
    except ValueError:
        if s.startswith("+"):
            return COLOR_CHG_DOWN  # 파랑
        if s.startswith("-"):
            return COLOR_CHG_UP  # 빨강
        return COLOR_LEVEL
    if v > 0:
        return COLOR_CHG_DOWN
    if v < 0:
        return COLOR_CHG_UP
    return COLOR_LEVEL


def ensure_cond_wow_column(table) -> None:
    """05 표에 BEP 옆 '전주대비' 열이 없으면 추가하고, 모든 행 셀 수를 8로 맞춘다."""
    from copy import deepcopy

    from lxml.etree import QName

    tbl = table._tbl
    tblGrid = tbl.tblGrid
    grid_cols = list(tblGrid.gridCol_lst)
    if len(grid_cols) < 7:
        return
    if len(grid_cols) < 8:
        new_gc = deepcopy(grid_cols[-1])
        tblGrid.append(new_gc)
        grid_cols = list(tblGrid.gridCol_lst)
    ratios = _cond_col_ratios(8)
    total_w = sum(int(gc.get("w") or 0) for gc in grid_cols) or 4500000
    widths = [max(200000, int(total_w * r)) for r in ratios]
    widths[-1] += total_w - sum(widths)
    for gc, w in zip(grid_cols[:8], widths):
        gc.set("w", str(w))
    for tr in tbl.tr_lst:
        tcs = list(tr.tc_lst)
        if not tcs:
            continue
        # 중간 열(인상2회) 테두리를 마지막 열에 복제해 noFill 우측선 방지
        border_src = tcs[min(5, len(tcs) - 1)]
        while len(tcs) < 8:
            new_tc = deepcopy(tcs[-1])
            for attr in ("gridSpan", "hMerge", "vMerge"):
                if attr in new_tc.attrib:
                    del new_tc.attrib[attr]
            for child in list(new_tc):
                tag = QName(child.tag).localname if isinstance(child.tag, str) else ""
                if tag == "txBody":
                    for p in child.findall(qn("a:p")):
                        for r in p.findall(qn("a:r")):
                            t = r.find(qn("a:t"))
                            if t is not None:
                                t.text = ""
            try:
                src_tcPr = border_src.find(qn("a:tcPr"))
                dst_tcPr = new_tc.find(qn("a:tcPr"))
                if src_tcPr is not None and dst_tcPr is not None:
                    src_lnR = src_tcPr.find(qn("a:lnR"))
                    dst_lnR = dst_tcPr.find(qn("a:lnR"))
                    if src_lnR is not None:
                        if dst_lnR is not None:
                            dst_tcPr.remove(dst_lnR)
                        # lnR는 lnL 다음
                        dst_tcPr.insert(1, deepcopy(src_lnR))
            except Exception:
                pass
            tr.append(new_tc)
            tcs = list(tr.tc_lst)


def _shape_text(shape) -> str:
    try:
        if shape.has_text_frame:
            return "".join(r.text for p in shape.text_frame.paragraphs for r in p.runs)
    except Exception:
        pass
    return ""


def _table_header(shape, col: int = 0) -> str:
    try:
        if shape.has_table:
            return (shape.table.cell(0, col).text_frame.text or "").strip()
    except Exception:
        pass
    return ""


def resolve_slide_shapes(slide) -> dict:
    """템플릿 도형 인덱스가 바뀌어도 내용으로 핵심 도형을 찾는다."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    shapes = list(slide.shapes)
    out: dict[str, Any] = {"shapes": shapes}

    def pick(idx: int | None, fallback_idx: int):
        if idx is not None and 0 <= idx < len(shapes):
            return shapes[idx]
        return shapes[fallback_idx]

    # tables by header
    for i, s in enumerate(shapes):
        if not s.has_table:
            continue
        h0 = _table_header(s, 0)
        h1 = _table_header(s, 1)
        if "레버리지 100%" in h1:
            out["cond100"] = s
        elif "레버리지 200%" in h1:
            out["cond200"] = s
        elif h0.startswith("시나리오"):
            out["repo_scenario"] = s
        elif "포트폴리오 편입" in h1 or h0 == "구분":
            out["portfolio_detail"] = s
        elif h0.startswith("섹터") and s.top < 1200000 and s.left < 4000000:
            out["minpyeong"] = s
        elif h0.startswith("섹터") and s.top < 1200000 and s.left >= 4000000:
            out["expected"] = s
        elif h0.startswith("섹터") and 2000000 < s.top < 3200000 and s.left < 4000000:
            out["credit_sp"] = s
        elif h0.startswith("섹터") and 2000000 < s.top < 3200000 and s.left >= 4000000:
            out["curve_sp"] = s
        elif "기준금리 시나리오" in h0:
            out["scenario_ret"] = s
        elif "시중금리 변동" in h0:
            out["rate_sens"] = s

    # text / groups by content
    for i, s in enumerate(shapes):
        if s.shape_type == MSO_SHAPE_TYPE.GROUP:
            for c in s.shapes:
                ct = _shape_text(c)
                if "민평금리" in ct or "BEP 정의" in ct:
                    out["footnote"] = s
                    break
            continue
        t = _shape_text(s)
        if t.startswith("• 은행채") or (t.startswith("은행채") and s.left < 500000):
            out["supply_bank"] = s
        elif "특수채:" in t and s.left < 500000:
            out["supply_special"] = s
        elif t.startswith("• 여전채") or (t.startswith("여전채") and s.left < 500000):
            out["supply_nbfc"] = s
        elif "신한자산운용" in t and "|" in t:
            out["date_line"] = s
        elif t.startswith("03 ") and "RV" in t:
            out["title_03"] = s

    # named special
    for s in shapes:
        if getattr(s, "name", "") == SUPPLY_SPECIAL_SHAPE_NAME:
            out["supply_special"] = s
            break

    # fallbacks to IDX_*
    out.setdefault("cond100", pick(None, IDX_COND_100))
    out.setdefault("cond200", pick(None, IDX_COND_200))
    out.setdefault("repo_scenario", pick(None, IDX_REPO_SCENARIO))
    out.setdefault("portfolio_detail", pick(None, IDX_PORTFOLIO_DETAIL))
    out.setdefault("minpyeong", pick(None, IDX_MINPYEONG))
    out.setdefault("expected", pick(None, IDX_EXPECTED))
    out.setdefault("credit_sp", pick(None, IDX_CREDIT_SP))
    out.setdefault("curve_sp", pick(None, IDX_CURVE_SP))
    out.setdefault("scenario_ret", pick(None, IDX_SCENARIO_RET))
    out.setdefault("rate_sens", pick(None, IDX_RATE_SENS))
    out.setdefault("footnote", pick(None, IDX_FOOTNOTE_GROUP))
    out.setdefault("supply_bank", pick(None, IDX_SUPPLY_BANK))
    out.setdefault("supply_nbfc", pick(None, IDX_SUPPLY_NBFC))
    out.setdefault("date_line", pick(None, IDX_DATE_LINE))
    out.setdefault("title_03", pick(None, IDX_TITLE_03))
    if "supply_special" not in out:
        try:
            out["supply_special"] = shapes[IDX_SUPPLY_SPECIAL]
        except Exception:
            pass
    return out


def _parse_supply_body(raw: str, label: str) -> str:
    """'• 은행채: 본문' / '은행채: 본문' / '본문' → 본문만."""
    s = (raw or "").strip()
    s = re.sub(r"^[•\-\*]+\s*", "", s)
    s = re.sub(rf"^{re.escape(label)}\s*[:：]\s*", "", s)
    return s.strip()


def _copy_textbox_body_pr(src_shape, dst_shape) -> None:
    """텍스트박스 bodyPr(여백·wrap·anchor)을 원본과 동일하게 맞춤."""
    try:
        src_body = src_shape._element.find(".//" + qn("a:bodyPr"))
        dst_tf = dst_shape._element.find(".//" + qn("p:txBody"))
        if src_body is None or dst_tf is None:
            return
        dst_body = dst_tf.find(qn("a:bodyPr"))
        if dst_body is None:
            dst_body = dst_tf.makeelement(qn("a:bodyPr"), {})
            dst_tf.insert(0, dst_body)
        for key in ("wrap", "lIns", "tIns", "rIns", "bIns", "anchor", "anchorCtr", "upright", "compatLnSpc"):
            if key in src_body.attrib:
                dst_body.set(key, src_body.get(key))
            elif key in dst_body.attrib:
                del dst_body.attrib[key]
        # python-pptx margin 속성도 동기화
        try:
            dst_shape.text_frame.margin_left = src_shape.text_frame.margin_left
            dst_shape.text_frame.margin_right = src_shape.text_frame.margin_right
            dst_shape.text_frame.margin_top = src_shape.text_frame.margin_top
            dst_shape.text_frame.margin_bottom = src_shape.text_frame.margin_bottom
        except Exception:
            pass
    except Exception:
        pass


def _set_supply_comment(shape, label: str, body: str) -> None:
    """사이드바 수급 1줄: • 라벨: 본문 (은행채/여전채와 동일 3-run)."""
    if not shape.has_text_frame:
        return
    body = (body or "").strip()
    tf = shape.text_frame
    try:
        tf.word_wrap = False
    except Exception:
        pass
    para = tf.paragraphs[0]
    try:
        para.alignment = PP_ALIGN.LEFT
    except Exception:
        pass
    for p in tf.paragraphs[1:]:
        for r in p.runs:
            r.text = ""
    runs = _ensure_para_runs(para, 3)
    parts = [
        ("• ", COLOR_SUPPLY, False),
        (f"{label}:", COLOR_SUPPLY, True),
        (f" {body}" if body else "", COLOR_SUPPLY, False),
    ]
    for i, (txt, col, bold) in enumerate(parts):
        runs[i].text = txt
        _apply_run_style(runs[i], col, FONT_DATA, font_size_pt=SUPPLY_FONT_PT, bold=bold)
    for r in runs[3:]:
        r.text = ""


def _find_or_add_supply_special(slide, bank_shape, nbfc_shape, resolved: dict | None = None):
    if resolved and resolved.get("supply_special") is not None:
        return resolved["supply_special"]
    for s in slide.shapes:
        if getattr(s, "name", "") == SUPPLY_SPECIAL_SHAPE_NAME:
            return s
        t = _shape_text(s)
        if "특수채:" in t and getattr(s, "left", 0) < 500000:
            return s
    box = slide.shapes.add_textbox(
        bank_shape.left,
        bank_shape.top + bank_shape.height + 18000,
        max(bank_shape.width, nbfc_shape.width),
        bank_shape.height,
    )
    box.name = SUPPLY_SPECIAL_SHAPE_NAME
    _copy_textbox_body_pr(bank_shape, box)
    return box


def _layout_supply_comments(slide, shapes, supply: dict, resolved: dict | None = None) -> None:
    """은행채 → 특수채 → 여전채 순으로 동일 서식·간격 배치."""
    resolved = resolved or {}
    bank = resolved.get("supply_bank") or shapes[IDX_SUPPLY_BANK]
    nbfc = resolved.get("supply_nbfc") or shapes[IDX_SUPPLY_NBFC]
    special = _find_or_add_supply_special(slide, bank, nbfc, resolved)
    gap = 18000
    h = bank.height
    left = bank.left
    width = max(bank.width, nbfc.width)

    bank_body = _parse_supply_body(supply.get("bank") or "", "은행채")
    special_body = _parse_supply_body(supply.get("special") or "", "특수채")
    nbfc_body = _parse_supply_body(supply.get("nbfc") or "", "여전채")

    # 위치: 은행채 유지 → 특수채 → 여전채 (왼쪽·너비·여백 동일)
    try:
        special.left = left
        special.width = width
        special.height = h
        special.top = bank.top + h + gap
        nbfc.left = left
        nbfc.width = width
        nbfc.height = h
        nbfc.top = special.top + h + gap
    except Exception:
        pass
    _copy_textbox_body_pr(bank, special)

    _set_supply_comment(bank, "은행채", bank_body)
    _set_supply_comment(special, "특수채", special_body)
    _set_supply_comment(nbfc, "여전채", nbfc_body)

    # 03 제목과 겹치면 RV 블록을 아래로 밀기 (특수채 박스 포함)
    try:
        title03 = (resolved or {}).get("title_03") or shapes[IDX_TITLE_03]
        need_top = nbfc.top + nbfc.height + 50000
        if title03.top < need_top:
            delta = need_top - title03.top
            for s in slide.shapes:
                try:
                    if s.left < 2300000 and s.top >= title03.top - 5000:
                        s.top = s.top + delta
                except Exception:
                    pass
    except Exception:
        pass


def _bring_shape_to_front(slide, shape) -> None:
    try:
        spTree = slide.shapes._spTree
        el = shape._element
        spTree.remove(el)
        spTree.append(el)
    except Exception:
        pass


def _send_shape_backward(slide, shape) -> None:
    """도형을 spTree 앞쪽(뒤 z-order)으로 보낸다."""
    try:
        spTree = slide.shapes._spTree
        el = shape._element
        spTree.remove(el)
        # [0]=nvGrpSpPr 등 예약 노드 뒤를 Prefer
        spTree.insert(2, el)
    except Exception:
        pass


def _group_display_top(group) -> int:
    """그룹 콘텐츠의 슬라이드 기준 최상단 Y (자식 좌표는 child-space)."""
    try:
        xfrm = group._element.grpSpPr.xfrm
        off_y = int(xfrm.off.get("y"))
        ch_y = int(xfrm.chOff.get("y"))
        child_tops = [c.top for c in group.shapes]
        if not child_tops:
            return off_y
        return off_y + (min(child_tops) - ch_y)
    except Exception:
        return int(group.top)


def _nudge_group_absolute(group, delta_y: int) -> None:
    """그룹 off·chOff·자식 y를 같은 만큼 이동 (표시·절대좌표 모두 반영)."""
    if not delta_y:
        return
    try:
        xfrm = group._element.grpSpPr.xfrm
        off = xfrm.off
        ch_off = xfrm.chOff
        off.set("y", str(int(off.get("y")) + delta_y))
        ch_off.set("y", str(int(ch_off.get("y")) + delta_y))
    except Exception:
        try:
            group.top = int(group.top) + delta_y
        except Exception:
            return
    try:
        for c in group.shapes:
            c.top = int(c.top) + delta_y
    except Exception:
        pass


def _find_footnote_group(slide, shapes):
    """05 하단 민평/백분위/BEP 각주 그룹을 찾는다."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    def _is_foot(s) -> bool:
        try:
            if s.shape_type != MSO_SHAPE_TYPE.GROUP:
                return False
            for c in s.shapes:
                if not c.has_text_frame:
                    continue
                tx = "".join(r.text for p in c.text_frame.paragraphs for r in p.runs)
                if "민평금리" in tx or "BEP 정의" in tx:
                    return True
        except Exception:
            return False
        return False

    try:
        cand = shapes[IDX_FOOTNOTE_GROUP]
        if _is_foot(cand):
            return cand
    except Exception:
        pass
    try:
        for s in slide.shapes:
            if _is_foot(s):
                return s
    except Exception:
        pass
    return None


def _ensure_footnote_below_tables(slide, shapes, shape100, shape200, resolved: dict | None = None) -> None:
    """05 각주 z-order만 조정. 위치는 템플릿(사용자 수정)을 존중한다."""
    foot = None
    if resolved:
        foot = resolved.get("footnote")
    if foot is None:
        foot = _find_footnote_group(slide, shapes)
    if foot is None:
        return
    # 표가 각주보다 위에 그려지도록 (템플릿 좌표는 유지)
    _bring_shape_to_front(slide, shape100)
    _bring_shape_to_front(slide, shape200)


def layout_cond_tables(shape100, shape200) -> None:
    """05 두 표에 동일 열너비·동일 행높이를 적용한다."""
    try:
        ensure_cond_wow_column(shape100.table)
        ensure_cond_wow_column(shape200.table)
    except Exception:
        pass
    # 두 표 외곽 폭·높이를 맞춘 뒤, 같은 절대 열너비 적용
    w = max(int(shape100.width), int(shape200.width))
    h = max(int(shape100.height), int(shape200.height))
    try:
        shape100.width = w
        shape200.width = w
        shape100.height = h
        shape200.height = h
    except Exception:
        pass
    ratios = _cond_col_ratios(8)
    widths = [max(180000, int(w * r)) for r in ratios]
    widths[-1] += w - sum(widths)
    n_rows = max(len(shape100.table.rows), len(shape200.table.rows), 1)
    row_h = max(160000, h // n_rows)
    # 합이 표 높이와 맞도록 마지막 행에서 보정
    row_heights = [row_h] * n_rows
    row_heights[-1] += h - sum(row_heights)
    for shp in (shape100, shape200):
        try:
            tbl = shp.table
            for ci, ww in enumerate(widths):
                if ci < len(tbl.columns):
                    tbl.columns[ci].width = ww
            for ri, tr in enumerate(tbl._tbl.tr_lst):
                if ri < len(row_heights):
                    tr.set("h", str(int(row_heights[ri])))
        except Exception:
            pass


def fill_cond_table(
    table,
    rows: list[dict],
    font_size_pt: float = 7,
    highlight_row: int | None = None,
    highlight_col: int | None = None,
) -> None:
    """05 조건별 표: 흰 배경, BEP·전주대비 슬레이트+틸 볼드(동일 서식)."""
    ensure_cond_wow_column(table)
    # highlight_col: 시나리오 인덱스 0=동결→열3, 1→열4, 2→열5
    scen_ci = (3 + int(highlight_col)) if highlight_col is not None and 0 <= int(highlight_col) <= 2 else None
    ncols = len(table.columns)
    # 구성 열 헤더는 템플릿의 레버리지 표기 유지
    try:
        leverage_hdr = (table.cell(0, 1).text_frame.text or "").strip() or "구성"
    except Exception:
        leverage_hdr = "구성"
    headers = ["만기", leverage_hdr, "RWA", "동결", "인상1회", "인상2회", "BEP", "전주대비"]
    for ci in range(min(ncols, 8)):
        try:
            is_scen_hdr = scen_ci is not None and ci == scen_ci
            if ci == 6:  # BEP
                hdr_color = COLOR_TEAL
            elif ci == 7:  # 전주대비 헤더는 검정
                hdr_color = COLOR_LEVEL
            elif is_scen_hdr:
                hdr_color = COLOR_SHINHAN_NAVY
            else:
                hdr_color = COLOR_LEVEL
            set_cell_fill(
                table.cell(0, ci),
                COLOR_SHINHAN_BG if is_scen_hdr else COLOR_HEADER_GRAY,
            )
            set_cell_text(
                table.cell(0, ci),
                headers[ci],
                font_size_pt=font_size_pt,
                bold=True,
                color=hdr_color,
            )
        except Exception:
            pass
    for i, row in enumerate(rows):
        ri = i + 1
        if ri >= len(table.rows):
            break
        is_hi = highlight_row is not None and i == highlight_row
        for ci in range(min(ncols, 8)):
            if is_hi or (scen_ci is not None and ci == scen_ci):
                fill = COLOR_SHINHAN_BG
            elif ci in (6, 7):
                fill = COLOR_SLATE_BG
            else:
                fill = COLOR_WHITE
            try:
                set_cell_fill(table.cell(ri, ci), fill)
            except Exception:
                pass
        vals = row.get("vals") or []
        txt_color = COLOR_SHINHAN_NAVY if is_hi else COLOR_LEVEL
        set_cell_text(
            table.cell(ri, 0),
            row.get("tenor") or "",
            font_size_pt=font_size_pt,
            bold=True,
            color=txt_color,
        )
        set_cell_text(
            table.cell(ri, 1),
            row.get("label") or "",
            center=False,
            font_size_pt=font_size_pt,
            bold=True,
            color=txt_color,
        )
        set_cell_text(
            table.cell(ri, 2),
            row.get("rwa") or "",
            font_size_pt=font_size_pt,
            bold=True,
            color=txt_color,
        )
        for j, v in enumerate(vals[:3]):
            ci = 3 + j
            is_scen = scen_ci is not None and ci == scen_ci
            set_cell_text(
                table.cell(ri, ci),
                v,
                font_size_pt=font_size_pt,
                bold=True,
                color=COLOR_SHINHAN_NAVY if is_scen else txt_color,
            )
        bep_color = COLOR_SHINHAN_NAVY if is_hi else COLOR_TEAL
        set_cell_text(
            table.cell(ri, 6),
            row.get("bep") or "",
            font_size_pt=font_size_pt,
            bold=True,
            color=bep_color,
        )
        # 전주대비: 헤더는 검정, 값 +파랑/−빨강
        try:
            if len(table.columns) < 8:
                ensure_cond_wow_column(table)
            wow = (row.get("wow") or "").strip() or "—"
            set_cell_fill(
                table.cell(ri, 7),
                COLOR_SHINHAN_BG if is_hi else COLOR_SLATE_BG,
            )
            set_cell_text(
                table.cell(ri, 7),
                wow,
                font_size_pt=font_size_pt,
                bold=True,
                color=_wow_value_color(wow),
            )
        except Exception:
            pass


def build_preview(csv_path: str, inputs: dict | None = None) -> dict[str, Any]:
    series = load_series(csv_path)
    dates = available_asof_dates(series)
    inp = merge_inputs(inputs)
    asof = (inp.get("asof") or "").strip() or (dates[0] if dates else "")
    if asof not in dates and dates:
        # pick nearest earlier
        earlier = [d for d in dates if d <= asof]
        asof = earlier[0] if earlier else dates[0]
    inp["asof"] = asof
    grids = build_grids(series, asof)
    prev_asof = grids.get("prev_week") or ""
    prev_grids = build_grids(series, prev_asof) if prev_asof else None
    scen = inp.get("repo_scenarios") or []
    cond100 = combo_returns(
        grids, COMBOS_100, scen, series=series, asof=asof, prev_grids=prev_grids
    )
    cond200 = combo_returns(
        grids, COMBOS_200, scen, series=series, asof=asof, prev_grids=prev_grids
    )
    port = portfolio_metrics(series, grids, inp)
    pf_rows = (inp.get("portfolio") or {}).get("rows") or []
    hi100 = match_combo_row_index(COMBOS_100, pf_rows)
    hi200 = match_combo_row_index(COMBOS_200, pf_rows)
    # serialize grids for UI (RGBColor → #RRGGBB)
    def _color_hex(c):
        if c is None:
            return None
        try:
            return f"#{int(c[0]):02X}{int(c[1]):02X}{int(c[2]):02X}"
        except Exception:
            return None

    def ser(g):
        out = {}
        for k, v in g.items():
            item = dict(v)
            if "line_colors" in item:
                item["line_colors"] = [_color_hex(c) for c in (item.get("line_colors") or [])]
            if "fill" in item:
                item["fill"] = _color_hex(item.get("fill"))
            out[f"{k[0]}|{k[1]}"] = item
        return out

    return {
        "asof": asof,
        "dates": dates[:120],
        "prev_week": grids["prev_week"],
        "minpyeong": ser(grids["minpyeong"]),
        "expected": ser(grids["expected"]),
        "credit": ser(grids["credit"]),
        "curve": ser(grids["curve"]),
        "cond100": cond100,
        "cond200": cond200,
        "cond100_hi": hi100,
        "cond200_hi": hi200,
        "portfolio": port,
        "inputs": inp,
        "sectors": [x[0] for x in SECTOR_ROWS],
        "sector_keys": [x[1] for x in SECTOR_ROWS],
        "tenors": TENORS,
    }


def fill_pptx(template_path: str, csv_path: str, inputs: dict | None = None) -> bytes:
    series = load_series(csv_path)
    dates = available_asof_dates(series)
    inp = merge_inputs(inputs)
    asof = (inp.get("asof") or "").strip() or (dates[0] if dates else "")
    if asof not in dates and dates:
        earlier = [d for d in dates if d <= asof]
        asof = earlier[0] if earlier else dates[0]
    inp["asof"] = asof
    grids = build_grids(series, asof)
    port = portfolio_metrics(series, grids, inp)
    prev_asof = grids.get("prev_week") or ""
    prev_grids = build_grids(series, prev_asof) if prev_asof else None
    scen = inp.get("repo_scenarios") or []
    cond100 = combo_returns(
        grids, COMBOS_100, scen, series=series, asof=asof, prev_grids=prev_grids
    )
    cond200 = combo_returns(
        grids, COMBOS_200, scen, series=series, asof=asof, prev_grids=prev_grids
    )

    prs = Presentation(template_path)
    slide = prs.slides[0]
    shapes = list(slide.shapes)
    resolved = resolve_slide_shapes(slide)

    # date line: 작성일(오늘) 기준
    weekdays = "월화수목금토일"
    today = datetime.now()
    auto_date = (
        f"{today.year}.{today.month:02d}.{today.day:02d} "
        f"({weekdays[today.weekday()]}) | 신한자산운용 채권운용2팀 강찬홍"
    )
    set_shape_text(resolved.get("date_line") or shapes[IDX_DATE_LINE], inp.get("date_line") or auto_date)

    fund = inp.get("fund") or {}
    set_shape_text(shapes[IDX_FUND_STRATEGY], fund.get("strategy") or "")
    set_shape_text(shapes[IDX_FUND_SECTOR], fund.get("sector") or "")
    set_shape_text(shapes[IDX_FUND_LEV], fund.get("leverage") or "")
    set_shape_text(shapes[IDX_BASE_RATE], fund.get("base_rate") or "")
    # auto-fill repo spot if blank-ish
    spot = fund.get("repo_spot") or ""
    if not spot or spot.startswith("["):
        rv = repo_spot(series, asof)
        br = _num(str(fund.get("base_rate") or "").replace("%", ""))
        if rv is not None and br is not None:
            spot = f"{rv:.2f}% ({'+' if rv - br >= 0 else ''}{round((rv - br) * 100)}bp)"
        elif rv is not None:
            spot = f"{rv:.2f}%"
    set_shape_text(shapes[IDX_REPO_SPOT], spot)

    supply = inp.get("supply") or {}
    _layout_supply_comments(slide, shapes, supply, resolved)
    rv = inp.get("rv") or {}
    curve_txt = (rv.get("curve") or "").strip()
    credit_txt = (rv.get("credit") or "").strip()
    cash_txt = (rv.get("cash") or "").strip()

    def _rv_lines(text: str, max_lines: int = 1) -> str:
        lines = [ln.strip() for ln in str(text or "").split("\n") if ln.strip()]
        return "\n".join(lines[:max_lines])

    # 크레딧 RV: 2줄 가능하도록 높이 확보
    try:
        credit_box = shapes[IDX_RV_CREDIT]
        if credit_box.height < RV_CREDIT_MIN_HEIGHT_EMU:
            credit_box.height = RV_CREDIT_MIN_HEIGHT_EMU
        # 현금성 박스와 겹치면 소폭 하향
        cash_box = shapes[IDX_RV_CASH]
        credit_bottom = credit_box.top + credit_box.height
        if cash_box.top < credit_bottom + 20000:
            cash_box.top = credit_bottom + 40000
            try:
                shapes[IDX_RV_CASH_TAG].top = cash_box.top - 120000
            except Exception:
                pass
    except Exception:
        pass
    set_shape_text(shapes[IDX_RV_CURVE], _rv_lines(curve_txt, 1), align="left")
    set_shape_text(shapes[IDX_RV_CREDIT], _rv_lines(credit_txt, 2), align="left")
    set_shape_text(shapes[IDX_RV_CASH], _rv_lines(cash_txt, 1), align="left")
    # 라벨 옆 짧은 칸은 비워 중복 표시 방지
    try:
        set_shape_text(shapes[IDX_RV_CURVE_TAG], "", align="left")
        set_shape_text(shapes[IDX_RV_CREDIT_TAG], "", align="left")
        set_shape_text(shapes[IDX_RV_CASH_TAG], "", align="left")
    except Exception:
        pass

    pf = inp.get("portfolio") or {}
    set_shape_text(shapes[IDX_REC_TITLE], pf.get("title") or "")
    set_shape_text(shapes[IDX_REC_COMPOSITION], pf.get("composition") or "")
    # 04 카드: 예상수익률 숫자는 골드, BEP 값은 시안
    _set_rec_metric_text(
        shapes[IDX_REC_RETURN],
        prefix="예상수익률 ",
        value=port.get("total_text") or "[X.XX]%",
        suffix=f" ({port.get('repo_scenario_label') or '동결'} 기준)",
        value_color=COLOR_GOLD,
    )
    _set_rec_metric_text(
        shapes[IDX_REC_BEP],
        prefix="BEP ",
        value=port.get("bep_text") or "[X.XX]%",
        suffix="",
        value_color=COLOR_SKY,
    )
    # 01~05 사이드바 소제목: 밑줄 제거 + 하늘색 볼드
    for tidx in IDX_SIDE_TITLES:
        try:
            _style_sidebar_title(shapes[tidx])
        except Exception:
            pass
    # 04 카드 좌측 골드 포인트 바 (약 3px)
    try:
        bar = shapes[IDX_REC_GOLD_BAR]
        bar.width = 38100
        _set_shape_solid_fill(bar, COLOR_GOLD)
    except Exception:
        pass

    # grids
    fill_grid_table(resolved["minpyeong"].table, grids["minpyeong"])
    fill_grid_table(resolved["expected"].table, grids["expected"])
    fill_grid_table(resolved["credit_sp"].table, grids["credit"])
    fill_grid_table(resolved["curve_sp"].table, grids["curve"])
    hi100 = match_combo_row_index(COMBOS_100, pf.get("rows") or [])
    hi200 = match_combo_row_index(COMBOS_200, pf.get("rows") or [])
    try:
        scen_hi = int(port.get("repo_scenario_idx") if port.get("repo_scenario_idx") is not None else 0)
    except (TypeError, ValueError):
        scen_hi = 0
    cond100_shp = resolved["cond100"]
    cond200_shp = resolved["cond200"]
    layout_cond_tables(cond100_shp, cond200_shp)
    fill_cond_table(
        cond100_shp.table, cond100, font_size_pt=7, highlight_row=hi100, highlight_col=scen_hi
    )
    fill_cond_table(
        cond200_shp.table, cond200, font_size_pt=7, highlight_row=hi200, highlight_col=scen_hi
    )
    _ensure_footnote_below_tables(slide, shapes, cond100_shp, cond200_shp, resolved)

    # 01~04 각주: 표와 겹치지 않게 소폭 하향, 기대수익률 산식 문구 갱신
    try:
        shapes[IDX_NOTE_MINPYEONG].top = shapes[IDX_NOTE_MINPYEONG].top + FOOTNOTE_NUDGE_EMU
        shapes[IDX_NOTE_EXPECTED].top = shapes[IDX_NOTE_EXPECTED].top + FOOTNOTE_NUDGE_EMU
        shapes[IDX_NOTE_CREDIT].top = shapes[IDX_NOTE_CREDIT].top + FOOTNOTE_NUDGE_EMU_BOTTOM
        shapes[IDX_NOTE_CURVE].top = shapes[IDX_NOTE_CURVE].top + FOOTNOTE_NUDGE_EMU_BOTTOM
        set_shape_text(shapes[IDX_NOTE_EXPECTED], NOTE_EXPECTED)
    except Exception:
        pass

    # repo scenario table (사이드바 05) — 전체 6.4pt
    scen_tbl = resolved["repo_scenario"].table
    repo_fs = 6.4
    for ci in range(len(scen_tbl.columns)):
        try:
            set_cell_fill(scen_tbl.cell(0, ci), COLOR_REPO_HEADER)
            hdr = scen_tbl.cell(0, ci)
            set_cell_text(
                hdr,
                (hdr.text_frame.text or "").strip(),
                font_size_pt=repo_fs,
                bold=True,
            )
        except Exception:
            pass
    for i, sc in enumerate(inp.get("repo_scenarios") or []):
        ri = i + 1
        if ri >= len(scen_tbl.rows):
            break
        row_fill = COLOR_REPO_ROW_ALT if i % 2 else COLOR_REPO_ROW
        for ci in range(len(scen_tbl.columns)):
            try:
                set_cell_fill(scen_tbl.cell(ri, ci), COLOR_REPO_AVG if ci == 3 else row_fill)
            except Exception:
                pass
        set_cell_text(scen_tbl.cell(ri, 0), sc.get("label") or "", center=False, font_size_pt=repo_fs)
        set_cell_text(scen_tbl.cell(ri, 1), fmt_pct(_num(sc.get("final_rate"))), font_size_pt=repo_fs)
        set_cell_text(scen_tbl.cell(ri, 2), fmt_pct(_num(sc.get("repo"))), font_size_pt=repo_fs)
        set_cell_text(scen_tbl.cell(ri, 3), fmt_pct(_num(sc.get("avg_1y"))), font_size_pt=repo_fs)

    # 06 섹션 강조 (제목·카드 테두리)
    try:
        emphasize_shape_text(
            shapes[IDX_PORT_SECTION_TITLE],
            font_size_pt=10,
            bold=True,
            color=COLOR_ACCENT,
        )
        emphasize_shape_line(shapes[IDX_PORT_SECTION_CARD], COLOR_ACCENT, width_pt=2.5)
    except Exception:
        pass

    # portfolio detail (01 민평수익률 기준)
    det = resolved["portfolio_detail"].table
    set_cell_text(det.cell(0, 2), "민평수익률", font_size_pt=7, bold=True, color=COLOR_SHINHAN_NAVY)
    set_cell_text(det.cell(0, 4), "기여수익률", font_size_pt=7, bold=True, color=COLOR_SHINHAN_NAVY)
    for ci in range(len(det.columns)):
        try:
            set_cell_fill(det.cell(0, ci), COLOR_HEADER_GRAY)
        except Exception:
            pass
    detail_rows = port.get("detail_rows") or []
    sel_label = port.get("repo_scenario_label") or "동결"
    for i, r in enumerate(detail_rows[:4]):
        ri = i + 1
        kind = r.get("kind") or ""
        name = r.get("name") or ""
        sector = r.get("sector") or ""
        # Repo 라벨에 선택 시나리오 명시
        if kind == "Repo" or sector == "REPO":
            kind = "Repo"
            name = f"Repo ({sel_label})"
        set_cell_text(det.cell(ri, 0), kind, center=False, font_size_pt=7, bold=True, color=COLOR_LEVEL)
        set_cell_text(det.cell(ri, 1), name, center=False, font_size_pt=7, bold=True, color=COLOR_LEVEL)
        set_cell_text(det.cell(ri, 2), r.get("rate_text") or "", font_size_pt=7, bold=True, color=COLOR_LEVEL)
        set_cell_text(det.cell(ri, 3), r.get("weight_text") or "", font_size_pt=7, bold=True, color=COLOR_LEVEL)
        set_cell_text(det.cell(ri, 4), r.get("contrib_text") or "", font_size_pt=7, bold=True, color=COLOR_LEVEL)
        for ci in range(5):
            set_cell_fill(det.cell(ri, ci), COLOR_WHITE)
    # total row: 신한 블루 배경 + 네이비 볼드
    if len(det.rows) > 5:
        total_txt = port.get("total_text") or ""
        set_cell_text(
            det.cell(5, 0), "합계 (보수 차감 전)", center=False, font_size_pt=7, bold=True, color=COLOR_SHINHAN_NAVY
        )
        set_cell_text(det.cell(5, 2), total_txt, font_size_pt=7, bold=True, color=COLOR_SHINHAN_NAVY)
        set_cell_text(det.cell(5, 3), "100%", font_size_pt=7, bold=True, color=COLOR_SHINHAN_NAVY)
        set_cell_text(det.cell(5, 4), total_txt, font_size_pt=7, bold=True, color=COLOR_SHINHAN_NAVY)
        for ci in range(5):
            set_cell_fill(det.cell(5, ci), COLOR_SHINHAN_BG)

    # scenario return table — 상단·헤더 쿨그레이, Base(선택 시나리오) 열만 신한 블루
    scen_ret = resolved["scenario_ret"].table
    headers = port.get("scen_headers") or []
    try:
        sel_i = int(port.get("repo_scenario_idx") if port.get("repo_scenario_idx") is not None else 0)
    except (TypeError, ValueError):
        sel_i = 0
    for ci in range(len(scen_ret.columns)):
        try:
            set_cell_fill(scen_ret.cell(0, ci), COLOR_HEADER_GRAY)
        except Exception:
            pass
    for i, h in enumerate(headers[:4]):
        is_base = i == sel_i
        set_cell_text(
            scen_ret.cell(1, i),
            h,
            font_size_pt=7,
            bold=True if is_base else None,
            color=COLOR_SHINHAN_NAVY if is_base else COLOR_LEVEL,
        )
        set_cell_fill(scen_ret.cell(1, i), COLOR_SHINHAN_BG if is_base else COLOR_HEADER_GRAY)
    vals = port.get("scen_vals") or []
    for i, v in enumerate(vals[:4]):
        if len(scen_ret.rows) <= 2:
            break
        is_base = i == sel_i
        set_cell_text(
            scen_ret.cell(2, i),
            v,
            font_size_pt=7,
            bold=True,
            color=COLOR_SHINHAN_NAVY if is_base else COLOR_LEVEL,
        )
        set_cell_fill(scen_ret.cell(2, i), COLOR_SHINHAN_BG if is_base else COLOR_WHITE)

    # sensitivity — 헤더 쿨그레이, +0.0% Base 열만 신한 블루
    sens_tbl = resolved["rate_sens"].table
    shocks = port.get("shocks") or []
    if len(sens_tbl.rows) > 0:
        for ci in range(len(sens_tbl.columns)):
            try:
                set_cell_fill(sens_tbl.cell(0, ci), COLOR_HEADER_GRAY)
            except Exception:
                pass
    for i, sh in enumerate(shocks[:7]):
        sign = "+" if sh > 0 else ""
        is_base = abs(float(sh)) < 1e-12
        set_cell_text(
            sens_tbl.cell(1, i),
            f"{sign}{sh:.1f}%",
            font_size_pt=7,
            bold=True if is_base else None,
            color=COLOR_SHINHAN_NAVY if is_base else COLOR_LEVEL,
        )
        set_cell_fill(sens_tbl.cell(1, i), COLOR_SHINHAN_BG if is_base else COLOR_HEADER_GRAY)
    for i, v in enumerate((port.get("sens_vals") or [])[:7]):
        if len(sens_tbl.rows) <= 2:
            break
        sh = shocks[i] if i < len(shocks) else None
        is_base = sh is not None and abs(float(sh)) < 1e-12
        set_cell_text(
            sens_tbl.cell(2, i),
            v,
            font_size_pt=7,
            bold=True,
            color=COLOR_SHINHAN_NAVY if is_base else COLOR_LEVEL,
        )
        set_cell_fill(sens_tbl.cell(2, i), COLOR_SHINHAN_BG if is_base else COLOR_WHITE)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def output_filename(asof: str | None = None) -> str:
    """다운로드 파일명 날짜는 작성일(오늘) 기준."""
    tag = datetime.now().strftime("%Y%m%d")
    return f"신한 매칭형 레포펀드 주간 제안서_{tag}.pptx"
