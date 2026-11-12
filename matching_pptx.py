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
COLOR_UP = RGBColor(0x3B, 0x82, 0xF6)  # 플러스 = 파란색
COLOR_DOWN = RGBColor(0xEF, 0x44, 0x44)  # 마이너스 = 빨간색
COLOR_DEFAULT = RGBColor(0x1E, 0x29, 0x3B)

ROOT = os.path.dirname(os.path.abspath(__file__))
DEFAULT_CSV = os.path.join(ROOT, "matching_data.csv")
DEFAULT_TEMPLATE = os.path.join(ROOT, "신한 매칭형 레포 채권펀드 주간 투자 제안서.pptx")

# PPTX 표 인덱스 (슬라이드 0 shapes)
IDX_REPO_SCENARIO = 71
IDX_PORTFOLIO_DETAIL = 72
IDX_COND_100 = 73
IDX_MINPYEONG = 74
IDX_EXPECTED = 123
IDX_CREDIT_SP = 124
IDX_CURVE_SP = 125
IDX_COND_200 = 122
IDX_SCENARIO_RET = 126
IDX_RATE_SENS = 127

# 텍스트 박스
IDX_DATE_LINE = 85
IDX_FUND_STRATEGY = 95
IDX_FUND_SECTOR = 97
IDX_FUND_LEV = 99
IDX_BASE_RATE = 101
IDX_REPO_SPOT = 103
IDX_SUPPLY_BANK = 80
IDX_SUPPLY_NBFC = 81
# RV 코멘트: 제목 아래 한 줄(넓은 상자). 라벨 옆 짧은 칸은 비움
IDX_RV_CURVE = 104
IDX_RV_CREDIT = 105
IDX_RV_CASH = 106
IDX_RV_CURVE_TAG = 109
IDX_RV_CREDIT_TAG = 111
IDX_RV_CASH_TAG = 113
IDX_REC_TITLE = 88
IDX_REC_COMPOSITION = 83
IDX_REC_RETURN = 89
IDX_REC_BEP = 90
IDX_PORT_SECTION_TITLE = 107  # 06 금주 추천 포트폴리오 상세
IDX_PORT_SECTION_CARD = 8  # 06 영역 카드 배경
IDX_NOTE_MINPYEONG = 115
IDX_NOTE_EXPECTED = 117
IDX_NOTE_CREDIT = 119
IDX_NOTE_CURVE = 121
FOOTNOTE_NUDGE_EMU = 70000  # 01·02 각주
FOOTNOTE_NUDGE_EMU_BOTTOM = 55000  # 03·04 각주
COLOR_ACCENT = RGBColor(0x1D, 0x4E, 0xD8)  # 06 강조 블루
COLOR_PORT_TOTAL_FILL = RGBColor(0xFE, 0xF3, 0xC7)  # 합계행 하이라이트
COLOR_PORT_RATE_FILL = RGBColor(0xEF, 0xF6, 0xFF)  # 수익률 열 연한 강조
COLOR_SCEN_SELECT_FILL = RGBColor(0xFE, 0xF3, 0xC7)  # 선택 시나리오 열 음영

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
    if v is None or not math.isfinite(v) or v == 0:
        return None
    return COLOR_UP if v > 0 else COLOR_DOWN


def _lerp_byte(a: int, b: int, t: float) -> int:
    return int(round(a + (b - a) * t))


def heatmap_rgb(pct: float | None) -> RGBColor | None:
    """백분위 0→붉은색, 50→흰색, 100→파란색 (셀 배경용 연한 톤)."""
    if pct is None or not math.isfinite(pct):
        return None
    t = max(0.0, min(100.0, float(pct))) / 100.0
    red = (252, 165, 165)  # #FCA5A5
    mid = (255, 255, 255)
    blue = (147, 197, 253)  # #93C5FD
    if t <= 0.5:
        u = t * 2.0
        r = _lerp_byte(red[0], mid[0], u)
        g = _lerp_byte(red[1], mid[1], u)
        b = _lerp_byte(red[2], mid[2], u)
    else:
        u = (t - 0.5) * 2.0
        r = _lerp_byte(mid[0], blue[0], u)
        g = _lerp_byte(mid[1], blue[1], u)
        b = _lerp_byte(mid[2], blue[2], u)
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
                "line_colors": [None, sign_color(chg)],
                "fill": None,
            }
            exp = expected_return(series, key, tenor, asof)
            exp_prv = expected_return(series, key, tenor, prev) if prev else None
            exp_chg = (exp - exp_prv) if (exp is not None and exp_prv is not None) else None
            expected[(key, tenor)] = {
                "level": exp,
                "chg": exp_chg,
                "text": fmt_minpyeong(exp, exp_chg),
                "line_colors": [None, sign_color(exp_chg)],
                "fill": None,
            }
            csp = credit_sp(series, key, tenor, asof)
            csp_pct = percentile_rank(hist_metric(series, key, tenor, asof, "credit"), csp) if csp is not None else None
            credit_fill = heatmap_rgb(csp_pct) if key != "국고" else None
            credit[(key, tenor)] = {
                "bp": csp,
                "pct": csp_pct,
                "text": fmt_bp_pct(csp, csp_pct) if key != "국고" else "",
                "line_colors": [sign_color(csp), None] if key != "국고" else [None, None],
                "fill": credit_fill,
            }
            cv = curve_sp(series, key, tenor, asof)
            cv_bp = cv * 100.0 if cv is not None else None
            cv_pct = percentile_rank(hist_metric(series, key, tenor, asof, "curve"), cv_bp) if cv_bp is not None else None
            curve[(key, tenor)] = {
                "bp": cv_bp,
                "pct": cv_pct,
                "text": fmt_curve_cell(cv_bp, cv_pct),
                "line_colors": [sign_color(cv_bp), None],
                "fill": heatmap_rgb(cv_pct),
            }
    apply_chg_heatmap(minpyeong)
    apply_chg_heatmap(expected)
    return {
        "asof": asof,
        "prev_week": prev,
        "minpyeong": minpyeong,
        "expected": expected,
        "credit": credit,
        "curve": curve,
    }


def apply_chg_heatmap(grid: dict) -> None:
    """01·02: 전주 대비 변동폭 상대 순위로 히트맵 (하락=붉, 상승=파랑)."""
    chgs = [v.get("chg") for v in grid.values() if v.get("chg") is not None and math.isfinite(v["chg"])]
    if not chgs:
        return
    lo, hi = min(chgs), max(chgs)
    for cell in grid.values():
        chg = cell.get("chg")
        if chg is None or not math.isfinite(chg):
            cell["fill"] = None
            continue
        if hi == lo:
            pct = 50.0
        else:
            pct = (chg - lo) / (hi - lo) * 100.0
        cell["fill"] = heatmap_rgb(pct)


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


def combo_returns(
    grids: dict,
    combos: list[dict],
    scenarios: list[dict],
    series: dict | None = None,
    asof: str | None = None,
) -> list[dict]:
    """05 조건별 수익. 채권은 01 민평, CD는 시장 CD (06과 동일 기준)."""
    minpyeong = grids["minpyeong"]
    asof = asof or grids.get("asof") or ""
    rows = []
    for combo in combos:
        tenor = combo["tenor"]
        asset = 0.0
        bond_w = 0.0
        er = 0.0
        ok = True
        for sector, w in combo["legs"]:
            if sector == "CD":
                rate = cd_rate(series or {}, "6개월", asof) if series else None
                if rate is None:
                    ok = False
                    break
                asset += w
                er += w * rate
                continue
            # 여전 → 캐피탈 민평; 06과 동일
            key = "캐피탈" if sector == "여전" else sector
            cell = minpyeong.get((key, tenor))
            if not cell or cell["level"] is None:
                ok = False
                break
            asset += w
            bond_w += w
            er += w * cell["level"]
        if not ok:
            rows.append({"tenor": tenor, "label": combo["label"], "rwa": combo["rwa"], "vals": ["", "", ""], "bep": ""})
            continue
        repo_w = 1.0 - asset  # negative
        vals = []
        freeze_ret = None
        for i, sc in enumerate(scenarios[:3]):
            repo = _num(sc.get("avg_1y"))
            if repo is None:
                vals.append("")
                continue
            ret = er + repo_w * repo
            vals.append(fmt_pct(ret))
            if i == 0:
                freeze_ret = ret
        # BEP: 동결수익 / 매칭듀레이션(채권 레그만, CD 제외)
        bep = ""
        dur = matching_duration(tenor, bond_w)
        if freeze_ret is not None and dur and dur != 0:
            bep_pct = freeze_ret / dur
            bep = fmt_pct(bep_pct)
        rows.append({"tenor": tenor, "label": combo["label"], "rwa": combo["rwa"], "vals": vals, "bep": bep})
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
        if center:
            para.alignment = PP_ALIGN.CENTER
        if para.runs:
            para.runs[0].text = line
            _apply_run_style(para.runs[0], line_color, font_name, font_size_pt, bold)
            for r in para.runs[1:]:
                r.text = ""
        else:
            run = para.add_run()
            run.text = line
            _apply_run_style(run, line_color, font_name, font_size_pt, bold)


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
    for ri, (_label, key) in enumerate(SECTOR_ROWS, start=1):
        for ci, tenor in enumerate(TENORS, start=1):
            cell = grid.get((key, tenor), {})
            tc = table.cell(ri, ci)
            set_cell_text(
                tc,
                cell.get(value_key, "") or "",
                line_colors=cell.get("line_colors"),
            )
            set_cell_fill(tc, cell.get("fill"))


def fill_cond_table(table, rows: list[dict], font_size_pt: float = 7) -> None:
    # rows 1..9 map to combo rows; col0/1/2 labels already in template
    for i, row in enumerate(rows):
        ri = i + 1
        if ri >= len(table.rows):
            break
        vals = row.get("vals") or []
        for j, v in enumerate(vals[:3]):
            set_cell_text(table.cell(ri, 3 + j), v, font_size_pt=font_size_pt)
        set_cell_text(table.cell(ri, 6), row.get("bep") or "", font_size_pt=font_size_pt)


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
    cond100 = combo_returns(
        grids, COMBOS_100, inp.get("repo_scenarios") or [], series=series, asof=asof
    )
    cond200 = combo_returns(
        grids, COMBOS_200, inp.get("repo_scenarios") or [], series=series, asof=asof
    )
    port = portfolio_metrics(series, grids, inp)
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
    cond100 = combo_returns(
        grids, COMBOS_100, inp.get("repo_scenarios") or [], series=series, asof=asof
    )
    cond200 = combo_returns(
        grids, COMBOS_200, inp.get("repo_scenarios") or [], series=series, asof=asof
    )

    prs = Presentation(template_path)
    slide = prs.slides[0]
    shapes = list(slide.shapes)

    # date line
    dt = _ymd(asof)
    weekdays = "월화수목금토일"
    auto_date = ""
    if dt:
        auto_date = f"{dt.year}.{dt.month:02d}.{dt.day:02d} ({weekdays[dt.weekday()]}) | 신한자산운용 채권운용2팀 강찬홍"
    set_shape_text(shapes[IDX_DATE_LINE], inp.get("date_line") or auto_date)

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
    set_shape_text(shapes[IDX_SUPPLY_BANK], supply.get("bank") or "")
    set_shape_text(shapes[IDX_SUPPLY_NBFC], supply.get("nbfc") or "")
    rv = inp.get("rv") or {}
    curve_txt = (rv.get("curve") or "").strip()
    credit_txt = (rv.get("credit") or "").strip()
    cash_txt = (rv.get("cash") or "").strip()
    # 제목 아래 넓은 칸에만 1줄 반영 (여러 줄이면 첫 줄)
    set_shape_text(shapes[IDX_RV_CURVE], curve_txt.split("\n")[0] if curve_txt else "", align="left")
    set_shape_text(shapes[IDX_RV_CREDIT], credit_txt.split("\n")[0] if credit_txt else "", align="left")
    set_shape_text(shapes[IDX_RV_CASH], cash_txt.split("\n")[0] if cash_txt else "", align="left")
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
    set_shape_text(shapes[IDX_REC_RETURN], port.get("return_text") or "")
    set_shape_text(shapes[IDX_REC_BEP], f"BEP {port.get('bep_text')}" if port.get("bep_text") else "BEP [X.XX]%")

    # grids
    fill_grid_table(shapes[IDX_MINPYEONG].table, grids["minpyeong"])
    fill_grid_table(shapes[IDX_EXPECTED].table, grids["expected"])
    fill_grid_table(shapes[IDX_CREDIT_SP].table, grids["credit"])
    fill_grid_table(shapes[IDX_CURVE_SP].table, grids["curve"])
    fill_cond_table(shapes[IDX_COND_100].table, cond100, font_size_pt=7)
    fill_cond_table(shapes[IDX_COND_200].table, cond200, font_size_pt=7)

    # 01~04 각주: 표와 겹치지 않게 소폭 하향, 기대수익률 산식 문구 갱신
    try:
        shapes[IDX_NOTE_MINPYEONG].top = shapes[IDX_NOTE_MINPYEONG].top + FOOTNOTE_NUDGE_EMU
        shapes[IDX_NOTE_EXPECTED].top = shapes[IDX_NOTE_EXPECTED].top + FOOTNOTE_NUDGE_EMU
        shapes[IDX_NOTE_CREDIT].top = shapes[IDX_NOTE_CREDIT].top + FOOTNOTE_NUDGE_EMU_BOTTOM
        shapes[IDX_NOTE_CURVE].top = shapes[IDX_NOTE_CURVE].top + FOOTNOTE_NUDGE_EMU_BOTTOM
        set_shape_text(shapes[IDX_NOTE_EXPECTED], NOTE_EXPECTED)
    except Exception:
        pass

    # repo scenario table
    scen_tbl = shapes[IDX_REPO_SCENARIO].table
    for i, sc in enumerate(inp.get("repo_scenarios") or []):
        ri = i + 1
        if ri >= len(scen_tbl.rows):
            break
        set_cell_text(scen_tbl.cell(ri, 0), sc.get("label") or "", center=False, font_size_pt=7)
        set_cell_text(scen_tbl.cell(ri, 1), fmt_pct(_num(sc.get("final_rate"))), font_size_pt=7)
        set_cell_text(scen_tbl.cell(ri, 2), fmt_pct(_num(sc.get("repo"))), font_size_pt=7)
        set_cell_text(scen_tbl.cell(ri, 3), fmt_pct(_num(sc.get("avg_1y"))), font_size_pt=7)

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
    det = shapes[IDX_PORTFOLIO_DETAIL].table
    set_cell_text(det.cell(0, 2), "민평수익률", font_size_pt=7, bold=True, color=COLOR_ACCENT)
    set_cell_text(det.cell(0, 4), "기여수익률", font_size_pt=7, bold=True, color=COLOR_ACCENT)
    detail_rows = port.get("detail_rows") or []
    for i, r in enumerate(detail_rows[:4]):
        ri = i + 1
        set_cell_text(det.cell(ri, 0), r.get("kind") or "", center=False, font_size_pt=7, bold=True)
        set_cell_text(det.cell(ri, 1), r.get("name") or "", center=False, font_size_pt=7, bold=True)
        set_cell_text(det.cell(ri, 2), r.get("rate_text") or "", font_size_pt=7, bold=True, color=COLOR_ACCENT)
        set_cell_text(det.cell(ri, 3), r.get("weight_text") or "", font_size_pt=7, bold=True)
        set_cell_text(det.cell(ri, 4), r.get("contrib_text") or "", font_size_pt=7, bold=True, color=COLOR_ACCENT)
        set_cell_fill(det.cell(ri, 2), COLOR_PORT_RATE_FILL)
        set_cell_fill(det.cell(ri, 4), COLOR_PORT_RATE_FILL)
    # total row: 민평·기여 합계, 투자비중 100%
    if len(det.rows) > 5:
        total_txt = port.get("total_text") or ""
        set_cell_text(det.cell(5, 0), "합계 (보수 차감 전)", center=False, font_size_pt=7, bold=True)
        set_cell_text(det.cell(5, 2), total_txt, font_size_pt=7, bold=True, color=COLOR_ACCENT)
        set_cell_text(det.cell(5, 3), "100%", font_size_pt=7, bold=True)
        set_cell_text(det.cell(5, 4), total_txt, font_size_pt=7, bold=True, color=COLOR_ACCENT)
        for ci in range(5):
            set_cell_fill(det.cell(5, ci), COLOR_PORT_TOTAL_FILL)

    # scenario return table (3x4) — header row1 already has scenario labels; fill row2
    scen_ret = shapes[IDX_SCENARIO_RET].table
    headers = port.get("scen_headers") or []
    for i, h in enumerate(headers[:4]):
        set_cell_text(scen_ret.cell(1, i), h, font_size_pt=7)
    vals = port.get("scen_vals") or []
    for i, v in enumerate(vals[:4]):
        if len(scen_ret.rows) > 2:
            set_cell_text(scen_ret.cell(2, i), v, font_size_pt=7, bold=True)
    # 06에서 고른 Repo 시나리오 열(헤더+값) 음영
    try:
        sel_i = int(port.get("repo_scenario_idx") if port.get("repo_scenario_idx") is not None else 0)
    except (TypeError, ValueError):
        sel_i = 0
    if 0 <= sel_i < 4:
        set_cell_fill(scen_ret.cell(1, sel_i), COLOR_SCEN_SELECT_FILL)
        if len(scen_ret.rows) > 2:
            set_cell_fill(scen_ret.cell(2, sel_i), COLOR_SCEN_SELECT_FILL)

    # sensitivity
    sens_tbl = shapes[IDX_RATE_SENS].table
    shocks = port.get("shocks") or []
    for i, sh in enumerate(shocks[:7]):
        sign = "+" if sh > 0 else ""
        set_cell_text(sens_tbl.cell(1, i), f"{sign}{sh:.1f}%", font_size_pt=7)
    for i, v in enumerate((port.get("sens_vals") or [])[:7]):
        if len(sens_tbl.rows) > 2:
            set_cell_text(sens_tbl.cell(2, i), v, font_size_pt=7, bold=True)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def output_filename(asof: str) -> str:
    tag = (asof or "").replace("-", "") or datetime.now().strftime("%Y%m%d")
    return f"신한 매칭형 레포펀드 주간 제안서_{tag}.pptx"
