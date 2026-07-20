# -*- coding: utf-8 -*-
"""
고정자산 감가상각비 재계산 스크립트
- sample_asset_ledger.xlsx (또는 회사 고정자산대장)를 읽어 각 자산의 당기(기준일 회계연도)
  감가상각비를 월할상각(정액법) / 국세청 상각률표(정률법) 기준으로 재계산한다.
- 아래 3가지 특수 케이스를 처리한다.
  1) 당기 중 처분(매각) 자산: 처분일까지만 상각
  2) 내용연수 재추정 자산: 재추정 시점의 장부가액을 새 내용연수로 재상각
  3) 내용연수 종료 자산: 더 이상 상각하지 않음
- 회사반영금액과 비교하여 차이나는 자산만 별도 시트에 정리하고 recalc_result.xlsx 로 저장한다.
- 취득원가/잔존가치/내용연수가 계산 불가능한 값(내용연수 0/음수, 잔존가치≥취득원가,
  취득원가 0 이하)인 자산은 재계산에서 제외하고 "데이터오류" 시트에 사유와 함께 모아
  보여준다(다른 정상 자산의 계산에는 영향을 주지 않는다).
"""
import datetime as dt
import os
import sys
from decimal import Decimal, ROUND_HALF_UP

import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    import anthropic
except ImportError:
    anthropic = None

# Windows 콘솔(PowerShell/cmd)이 시스템 코드페이지(cp949 등)로 표준출력을 열어
# 한글 안내 메시지가 깨지는 것을 막기 위해 표준출력 인코딩을 UTF-8로 고정한다.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

IN_PATH = r"C:\Users\wh981\재무검증도구\depreciation-recalc-tool\sample_asset_ledger.xlsx"
OUT_PATH = r"C:\Users\wh981\재무검증도구\depreciation-recalc-tool\recalc_result.xlsx"

# 기준일: 이 날짜가 속한 회계연도(1/1~12/31)의 당기 감가상각비를 재계산한다.
REF_DATE = dt.date(2025, 12, 31)
FY_YEAR = REF_DATE.year

# 중요성 기준 금액(원): 회사반영 대비 재계산 차이의 절대값이 이 금액 미만이면
# "경미한 차이", 이상이면 "유의한 차이"로 구분한다.
MATERIALITY_THRESHOLD = 1_000_000

# "유의한 차이" 자산의 AI 추정원인 코멘트 생성에 사용할 모델.
ANTHROPIC_MODEL = "claude-sonnet-4-6"

# 다기간(연도별) 비교 대상 회계연도 목록. 기본값처럼 [FY_YEAR] 하나만 두면
# 기존과 완전히 동일하게 단일 연도만 계산한다. 여러 연도 추이를 보려면
# 예: COMPARISON_YEARS = [2023, 2024, 2025]
COMPARISON_YEARS = [FY_YEAR]

# 다기간 비교에서 전년대비 이 비율(%) 이상 증감하면 "경고"로 표시한다.
YOY_ANOMALY_THRESHOLD_PCT = 20.0

# ---------------------------------------------------------------------------
# 컬럼명 매핑
# 회사마다 고정자산대장의 컬럼명이 다를 수 있으므로, 실제 파일의 컬럼명을
# 여기서만 바꿔 지정하면 스크립트 본문은 수정할 필요가 없다.
# 예: 회사 파일에서 취득원가 컬럼이 "취득가액"이라는 이름이면
#     "취득원가": "취득가액" 으로 바꾸면 된다.
#
# 처분일 / 내용연수재추정일 / 재추정후내용연수 는 선택 항목이다.
# 회사 파일에 해당 컬럼 자체가 없으면(매핑된 이름이 파일에 없으면) 모든 자산에
# 처분/재추정이 없는 것으로 간주하고 계속 진행한다.
# ---------------------------------------------------------------------------
COLUMN_MAP = {
    "자산명": "자산명",
    "자산분류": "자산분류(유형자산/무형자산)",
    "취득일": "취득일",
    "취득원가": "취득원가",
    "잔존가치": "잔존가치",
    "내용연수": "내용연수(년)",
    "상각방법": "상각방법(정액법/정률법)",
    "회사반영상각비": "회사반영_당기감가상각비",
    "처분일": "처분일",
    "재추정일": "내용연수재추정일",
    "재추정내용연수": "재추정후내용연수(년)",
    "재추정후상각방법": "재추정후상각방법(정액법/정률법)",
    "총예정생산량": "총예정생산량",
    "당기실제생산량": "당기실제생산량",
    "자본적지출일": "자본적지출일",
    "자본적지출액": "자본적지출액",
    "상각중단시작일": "상각중단시작일",
    "상각중단종료일": "상각중단종료일",
    "회사반영누계상각액": "회사반영_전기말감가상각누계액",
}

REQUIRED_KEYS = [
    "자산명", "자산분류", "취득일", "취득원가", "잔존가치",
    "내용연수", "상각방법", "회사반영상각비",
]
OPTIONAL_KEYS = ["처분일", "재추정일", "재추정내용연수", "재추정후상각방법",
                 "총예정생산량", "당기실제생산량",
                 "자본적지출일", "자본적지출액", "상각중단시작일", "상각중단종료일",
                 "회사반영누계상각액"]

# 국세청 고시 내용연수별 상각률표(정률법). 표에 없는 내용연수는
# 잔존가치 5% 가정(r = 1-0.05**(1/n))으로 근사한다.
RATE_TABLE = {
    2: 0.684, 3: 0.536, 4: 0.528, 5: 0.451, 6: 0.394,
    7: 0.349, 8: 0.313, 9: 0.284, 10: 0.259,
    11: 0.239, 12: 0.221, 13: 0.206, 14: 0.193, 15: 0.183,
    16: 0.173, 17: 0.164, 18: 0.157, 19: 0.150, 20: 0.144,
}


def classify_materiality(diff: int, threshold: int = MATERIALITY_THRESHOLD) -> str:
    return "유의한 차이" if abs(diff) >= threshold else "경미한 차이"


def validate_asset_inputs(cost: float, salvage: float, life: int,
                           method: str = None, total_units: float = None,
                           reest_method: str = None,
                           capex_date=None, capex_amount: float = None,
                           susp_start=None, susp_end=None,
                           accum_reported: float = None) -> list:
    """
    재계산에 들어가기 전 취득원가/잔존가치/내용연수가 계산 가능한 값인지 검증한다.
    문제가 있으면 오류 사유 문자열 리스트를, 문제가 없으면 빈 리스트를 반환한다.
    여기서 걸러진 자산은 recalc_asset을 호출하지 않으므로(0/음수 내용연수로 인한
    ZeroDivisionError 등), 다른 자산의 계산에 영향을 주지 않고 안전하게 제외된다.

    생산량비례법(method="생산량비례법")은 내용연수를 계산에 쓰지 않고 대신
    총예정생산량을 쓰므로, 이 경우엔 내용연수 검증 대신 총예정생산량을 검증한다.

    자본적지출/상각중단/재추정후상각방법은 자산에 이벤트가 있을 때만 값이 채워지므로
    (선택 컬럼), 각각 "짝이 맞는지"와 "값 자체가 유효한지"만 검증한다.
    """
    errors = []
    if method != "생산량비례법" and life < 1:
        errors.append(f"내용연수 오류(내용연수={life}년, 1년 이상의 정수여야 함)")
    if cost <= 0:
        errors.append(f"취득원가 오류(취득원가={cost:,.0f}원, 0보다 커야 함)")
    if salvage >= cost:
        errors.append(f"잔존가치 오류(잔존가치={salvage:,.0f}원, 취득원가={cost:,.0f}원 미만이어야 함)")
    if method == "생산량비례법" and (total_units is None or total_units <= 0):
        errors.append(f"총예정생산량 오류(총예정생산량={total_units}, 0보다 커야 함)")
    if reest_method is not None and reest_method not in ("정액법", "정률법"):
        errors.append(f"재추정후상각방법 오류(값={reest_method}, 정액법/정률법 중 하나여야 함)")
    if (capex_date is None) != (capex_amount is None):
        errors.append("자본적지출 오류(자본적지출일과 자본적지출액은 함께 입력되어야 함)")
    elif capex_amount is not None and capex_amount <= 0:
        errors.append(f"자본적지출액 오류(자본적지출액={capex_amount:,.0f}원, 0보다 커야 함)")
    if (susp_start is None) != (susp_end is None):
        errors.append("상각중단기간 오류(상각중단시작일과 상각중단종료일은 함께 입력되어야 함)")
    elif susp_start is not None and susp_end is not None and susp_end < susp_start:
        errors.append(f"상각중단기간 오류(상각중단종료일={susp_end}이 상각중단시작일={susp_start}보다 빠름)")
    if accum_reported is not None and accum_reported < 0:
        errors.append(f"회사반영_전기말감가상각누계액 오류(값={accum_reported:,.0f}원, 0 이상이어야 함)")
    return errors


def get_ai_estimated_cause(diff: int, elapsed_months: int, method: str,
                            is_reestimated: bool, is_disposed: bool) -> str:
    """
    "유의한 차이" 자산에 대해 Claude API로 차이 발생 원인을 한 줄로 추정한다.
    API 키가 없거나(.env 미설정) 호출 중 오류가 발생하면 이 기능만 건너뛰고
    "-"를 반환한다(나머지 재계산 로직에는 영향을 주지 않는다).
    """
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key or anthropic is None:
        return "-"

    try:
        client = anthropic.Anthropic(api_key=api_key)
        prompt = (
            "고정자산 감가상각비 재계산 검증에서 아래 자산의 회사반영 감가상각비와 "
            "재계산 감가상각비 사이에 '유의한 차이'가 발생했습니다.\n"
            f"- 차이금액(재계산-회사반영): {diff:,}원\n"
            f"- 경과개월수(취득일~기준일): {elapsed_months}개월\n"
            f"- 상각방법: {method}\n"
            f"- 내용연수 재추정 여부: {'있음' if is_reestimated else '없음'}\n"
            f"- 당기중 처분 여부: {'있음' if is_disposed else '없음'}\n\n"
            "회계/세무 실무자 관점에서 이런 차이가 발생했을 가능성이 가장 높은 원인을 "
            "한국어 한 줄(50자 이내)로 추정해서 답변하세요. 결론 한 줄 외 다른 설명은 출력하지 마세요."
        )
        response = client.messages.create(
            model=ANTHROPIC_MODEL,
            max_tokens=200,
            messages=[{"role": "user", "content": prompt}],
        )
        text = next((b.text for b in response.content if b.type == "text"), "").strip()
        return text if text else "-"
    except Exception as e:
        print(f"  [경고] AI 추정원인 호출 실패, 이 자산은 건너뜁니다: {e}")
        return "-"


def get_rule_based_cause(is_reestimated: bool, is_disposed: bool, note: str,
                          is_capex: bool = False, is_suspended: bool = False) -> str:
    """
    이미 계산되어 있는 처분/재추정/자본적지출/상각중단/내용연수종료 정보만으로 원인이
    명확히 설명되는 케이스를 규칙으로 1차 분류한다. 분류 가능하면 원인 문자열을,
    애매하면 None을 반환해서 호출부가 AI 호출(get_ai_estimated_cause)로 폴백하도록
    한다(모든 API 호출을 규칙으로 대체할 수 있어 비용/속도를 아낀다).
    """
    if is_disposed and "당기중처분" in note:
        return "처분일 반영 오류 추정(처분일까지의 월할상각 미반영 가능성)"
    if is_disposed and "전기이전처분" in note:
        return "전기 이전 처분자산에 당기 상각비를 계속 반영했을 가능성"
    if is_reestimated:
        return "내용연수 재추정 반영 오류 추정(재추정시점 장부가액/신규 내용연수 미반영 가능성)"
    if is_capex and "자본적지출" in note:
        return "자본적지출 반영 오류 추정(지출액 가산 또는 잔여내용연수 재상각 미반영 가능성)"
    if is_suspended and "상각중단" in note:
        return "감가상각중단 반영 오류 추정(중단기간 0원 처리 또는 종료시점 연장 미반영 가능성)"
    if "내용연수종료" in note:
        return "내용연수 종료 자산에 상각비를 계속 반영했을 가능성"
    return None


def get_rate(life: int) -> float:
    if life in RATE_TABLE:
        return RATE_TABLE[life]
    return round(1 - 0.05 ** (1 / life), 3)


def month_index(year: int, month: int) -> int:
    return year * 12 + month


def idx_to_year(idx: int) -> int:
    """month_index(year, month)의 역함수 중 연도만 구한다(month는 1~12 범위 가정)."""
    return (idx - 1) // 12


def round_won(x: float) -> int:
    return int(Decimal(str(x)).quantize(0, rounding=ROUND_HALF_UP))


def to_date_or_none(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or pd.isna(v):
        return None
    return pd.to_datetime(v).date()


def elapsed_months_to_ref(acq: dt.date, ref: dt.date) -> int:
    """취득일부터 기준일까지 경과 개월수(취득월 포함, 월 단위 카운트)."""
    return month_index(ref.year, ref.month) - month_index(acq.year, acq.month) + 1


def fy_ref_date(year: int) -> dt.date:
    """
    다기간 비교에서 각 회계연도의 기준일을 반환한다. REF_DATE의 연도와 같으면
    REF_DATE를 그대로 쓰고(12/31이 아닌 기준일 설정도 존중), 그 외 연도는
    해당 연도의 12/31을 기준일로 본다.
    """
    return REF_DATE if year == FY_YEAR else dt.date(year, 12, 31)


# ---------------------------------------------------------------------------
# 정액법 재계산
# ---------------------------------------------------------------------------
def build_straight_line_segments(acq, cost, salvage, life_years, reest_date, reest_life_years):
    """
    정액법 상각 스케줄을 '구간(segment)' 목록으로 만든다.
    - 재추정이 없으면 구간은 1개: [취득월 ~ 취득월+내용연수*12-1], 기준가액=취득원가.
    - 내용연수 재추정이 있으면, 재추정월 직전까지가 1구간(기존 내용연수 기준),
      재추정월부터는 그 시점의 장부가액을 새 기준가액으로 하는 2구간(재추정후내용연수 기준)이
      새로 시작된다. 즉 "재추정 시점부터 남은 장부가액을 새 내용연수로 재상각"을 그대로 구현한다.
    """
    start_idx = month_index(acq.year, acq.month)
    life_months = life_years * 12

    if reest_date is None:
        return [dict(start_idx=start_idx, end_idx=start_idx + life_months - 1,
                      basis=cost, salvage=salvage, life_months=life_months)]

    reest_idx = month_index(reest_date.year, reest_date.month)
    monthly_dep1 = (cost - salvage) / life_months
    # 재추정 시점까지 원래 스케줄로 상각된 개월수(생애 종료를 넘지 않도록 방어)
    months_before_reest = max(0, min(reest_idx - start_idx, life_months))
    bv_at_reest = cost - monthly_dep1 * months_before_reest

    seg1 = dict(start_idx=start_idx, end_idx=reest_idx - 1,
                basis=cost, salvage=salvage, life_months=life_months)

    new_life_months = reest_life_years * 12
    seg2 = dict(start_idx=reest_idx, end_idx=reest_idx + new_life_months - 1,
                basis=bv_at_reest, salvage=salvage, life_months=new_life_months)
    return [seg1, seg2]


def straight_line_current_period_dep(segments, disposal_idx, fy_year):
    """
    당기 회계연도와 각 구간이 겹치는 개월수만큼 정액상각비를 계산해 합산한다.
    - 처분일이 있으면 당기 말(fy_end_idx)을 처분월로 앞당겨, 처분월 이후 개월은
      상각 개월수 계산에서 자동으로 제외된다(처분일까지만 상각).
    - 내용연수(구간의 end_idx)를 넘어서는 달은 겹침이 없어 자동으로 0개월 처리되므로
      내용연수가 끝난 자산은 별도 분기 없이도 더 이상 상각되지 않는다.
    """
    fy_start_idx = month_index(fy_year, 1)
    fy_end_idx = month_index(fy_year, 12)
    if disposal_idx is not None:
        fy_end_idx = min(fy_end_idx, disposal_idx)

    total_dep = 0.0
    total_months = 0
    for seg in segments:
        overlap_start = max(seg["start_idx"], fy_start_idx)
        overlap_end = min(seg["end_idx"], fy_end_idx)
        if overlap_start > overlap_end:
            continue
        months = overlap_end - overlap_start + 1
        monthly_dep = (seg["basis"] - seg["salvage"]) / seg["life_months"]
        total_dep += monthly_dep * months
        total_months += months
    return total_dep, total_months


# ---------------------------------------------------------------------------
# 정률법 재계산
# ---------------------------------------------------------------------------
def declining_balance_current_period_dep(acq, cost, salvage, life_years,
                                          reest_date, reest_life_years,
                                          disposal_date, fy_year):
    """
    정률법은 매년 '기초 장부가액 x 상각률'을 연 단위로 누적 적용하고,
    취득 첫 해/처분되는 해만 월할 비례한다(국세청 상각률표를 적용하는 일반적인 실무 방식과 동일).

    - 내용연수 재추정: 재추정일이 속한 회계연도의 1월 1일부터 새 내용연수/새 상각률을
      적용하는 것으로 단순화한다(재추정은 통상 회계연도 초부터 적용하는 것이 일반적이며,
      정액법과 달리 정률법은 상각률이 연 단위 장부가액에 곱해지는 구조라 재추정일을
      회계연도 중간의 특정 월로 정밀하게 나누는 것이 실무적 의미가 크지 않기 때문).
    - 처분: 처분월까지만 그 해의 상각 개월수에 포함한다.
    - 내용연수 종료: 그 해의 상각 종료월(active_end_idx)을 넘어서면 개월수가 0이 되어
      더 이상 상각되지 않는다.
    """
    rate = get_rate(life_years)
    start_idx = month_index(acq.year, acq.month)
    active_end_idx = start_idx + life_years * 12 - 1  # 상각 종료월(생애 종료월)
    active_rate = rate

    reest_year = reest_date.year if reest_date is not None else None
    disposal_idx = month_index(disposal_date.year, disposal_date.month) if disposal_date is not None else None

    book_value = float(cost)
    current_period_dep = 0.0
    current_period_months = 0

    y = acq.year
    while y <= fy_year:
        if reest_year is not None and y == reest_year:
            active_rate = get_rate(reest_life_years)
            active_end_idx = month_index(y, 1) + reest_life_years * 12 - 1

        year_start_idx = month_index(y, 1)
        year_end_idx = month_index(y, 12)
        eff_start = max(start_idx, year_start_idx)
        eff_end = min(active_end_idx, year_end_idx)
        if disposal_idx is not None:
            eff_end = min(eff_end, disposal_idx)

        months = eff_end - eff_start + 1 if eff_end >= eff_start else 0
        dep = 0.0
        if months > 0:
            dep = book_value * active_rate * months / 12
            if book_value - dep < salvage:
                dep = book_value - salvage
            book_value -= dep

        if y == fy_year:
            current_period_dep = dep
            current_period_months = months
        y += 1

    return max(current_period_dep, 0.0), current_period_months


# ---------------------------------------------------------------------------
# 이벤트 기반 통합 스케줄 (자본적지출 / 재추정+방법변경 / 상각중단)
# 정액법 전용 함수(build_straight_line_segments) 하나로는 "재추정으로 상각방법이
# 정률법으로 바뀌는" 케이스를 표현할 수 없으므로, 여기서는 "구간(segment)마다
# method가 다를 수 있는" 통합 스케줄을 만든다. 이 경로는 자본적지출/상각중단/
# 방법이 실제로 바뀌는 재추정이 하나라도 있는 자산에 대해서만 사용되고
# (recalc_asset의 has_extended_events 분기), 그 외의 기존 자산은 기존
# build_straight_line_segments/declining_balance_current_period_dep를 그대로 탄다.
# ---------------------------------------------------------------------------
def nominal_month_index(calendar_idx: int, susp_start_idx, susp_end_idx) -> int:
    """
    상각중단을 반영해 달력월 인덱스를 '명목 경과월 인덱스'로 변환한다.
    - 중단 이전: 그대로.
    - 중단 구간 안: 경과가 멈춘 것으로 보고 중단 시작 직전월로 고정.
    - 중단 이후: 중단 길이만큼 당겨서(상각이 그만큼 덜 진행된 것으로) 계산한다.
    장부가액(경과개월 기반) 계산에만 쓰인다 — 구간 종료월(end_idx) 연장은
    apply_suspension_extension이 별도로 처리한다.
    """
    if susp_start_idx is None:
        return calendar_idx
    if calendar_idx < susp_start_idx:
        return calendar_idx
    if calendar_idx <= susp_end_idx:
        return susp_start_idx - 1
    return calendar_idx - (susp_end_idx - susp_start_idx + 1)


def _declining_balance_year_loop(start_idx, basis, salvage, rate, active_end_idx,
                                  disposal_idx, susp_start_idx, susp_end_idx, up_to_idx):
    """
    declining_balance_current_period_dep(위)의 연단위 누적 로직을, 특정 구간이
    start_idx부터 up_to_idx(포함)까지 상각됐을 때의 (최종 장부가액, up_to_idx가
    속한 회계연도의 그 해 상각비, 그 해 상각개월수)를 구하도록 일반화한 버전이다.
    상각중단과 겹치는 개월을 그 해의 상각개월수에서 제외하는 점만 원본과 다르다.
    """
    if up_to_idx < start_idx:
        return float(basis), 0.0, 0

    start_year = idx_to_year(start_idx)
    up_to_year = idx_to_year(up_to_idx)

    book_value = float(basis)
    dep_in_target_year, months_in_target_year = 0.0, 0

    y = start_year
    while y <= up_to_year:
        year_start_idx = month_index(y, 1)
        year_end_idx = month_index(y, 12)
        eff_start = max(start_idx, year_start_idx)
        eff_end = min(active_end_idx, year_end_idx)
        if disposal_idx is not None:
            eff_end = min(eff_end, disposal_idx)
        if y == up_to_year:
            eff_end = min(eff_end, up_to_idx)

        months = eff_end - eff_start + 1 if eff_end >= eff_start else 0
        if susp_start_idx is not None and months > 0:
            s, e = max(eff_start, susp_start_idx), min(eff_end, susp_end_idx)
            if s <= e:
                months -= (e - s + 1)

        dep = 0.0
        if months > 0:
            dep = book_value * rate * months / 12
            if book_value - dep < salvage:
                dep = book_value - salvage
            book_value -= dep

        if y == up_to_year:
            dep_in_target_year, months_in_target_year = dep, months
        y += 1

    return book_value, dep_in_target_year, months_in_target_year


def _book_value_at(seg: dict, up_to_idx: int, susp_start_idx, susp_end_idx) -> float:
    """구간(seg) 시작부터 up_to_idx(달력월, 포함)까지 상각했을 때의 장부가액."""
    if up_to_idx < seg["start_idx"]:
        return seg["basis"]
    if seg["method"] == "정액법":
        nom_up_to = nominal_month_index(up_to_idx, susp_start_idx, susp_end_idx)
        nom_start = nominal_month_index(seg["start_idx"], susp_start_idx, susp_end_idx)
        elapsed = max(0, min(nom_up_to - nom_start + 1, seg["life_months"]))
        monthly = (seg["basis"] - seg["salvage"]) / seg["life_months"]
        return seg["basis"] - monthly * elapsed
    else:
        book_value, _, _ = _declining_balance_year_loop(
            seg["start_idx"], seg["basis"], seg["salvage"], seg["rate"], seg["end_idx"],
            None, susp_start_idx, susp_end_idx, up_to_idx)
        return book_value


def build_depreciation_schedule(acq, cost, salvage, life_years, method,
                                 reest_date, reest_life_years, reest_method,
                                 capex_date, capex_amount,
                                 susp_start_idx=None, susp_end_idx=None) -> list:
    """
    자본적지출(최대 1건)과 재추정(+방법변경, 최대 1건)을 시간순으로 반영해
    구간(segment) 리스트를 만든다. 각 구간(dict)은 start_idx, end_idx(명목 종료월),
    method, basis, salvage, life_months, rate를 담고, method는 구간마다 다를 수 있다.

    - 같은 날짜에 자본적지출과 재추정이 겹치면 자본적지출을 먼저 반영한다(가산된
      장부가액을 재추정의 기준가로 삼는 것이 실무적으로 자연스럽기 때문).
    - 자본적지출: end_idx/method 불변, basis = 그 시점 장부가액 + 자본적지출액
      ("잔여 내용연수로 계속 상각").
    - 재추정: basis = 그 시점 장부가액(가산 없음), method = reest_method or 기존 method,
      end_idx = 재추정월 + reest_life_years*12 - 1 (새 내용연수로 완전히 재설정).
    """
    start_idx0 = month_index(acq.year, acq.month)
    cur = dict(start_idx=start_idx0, end_idx=start_idx0 + life_years * 12 - 1,
               method=method, basis=float(cost), salvage=float(salvage),
               life_months=life_years * 12,
               rate=get_rate(life_years) if method == "정률법" else None)

    breakpoints = []
    if capex_date is not None:
        breakpoints.append(dict(kind="capex", idx=month_index(capex_date.year, capex_date.month),
                                 amount=capex_amount))
    if reest_date is not None:
        breakpoints.append(dict(kind="reest", idx=month_index(reest_date.year, reest_date.month),
                                 new_life_years=reest_life_years, new_method=reest_method))
    breakpoints.sort(key=lambda b: (b["idx"], 0 if b["kind"] == "capex" else 1))

    segments = []
    for bp in breakpoints:
        prev_end_idx = cur["end_idx"]  # capex는 이 종료월을 그대로 이어받는다
        book_value = _book_value_at(cur, bp["idx"] - 1, susp_start_idx, susp_end_idx)
        cur = dict(cur, end_idx=min(cur["end_idx"], bp["idx"] - 1))
        segments.append(cur)

        if bp["kind"] == "capex":
            cur = dict(start_idx=bp["idx"], end_idx=prev_end_idx, method=segments[-1]["method"],
                       basis=book_value + bp["amount"], salvage=segments[-1]["salvage"],
                       life_months=prev_end_idx - bp["idx"] + 1, rate=segments[-1]["rate"])
        else:
            new_method = bp["new_method"] or segments[-1]["method"]
            new_life_months = bp["new_life_years"] * 12
            cur = dict(start_idx=bp["idx"], end_idx=bp["idx"] + new_life_months - 1,
                       method=new_method, basis=book_value, salvage=segments[-1]["salvage"],
                       life_months=new_life_months,
                       rate=get_rate(bp["new_life_years"]) if new_method == "정률법" else None)

    segments.append(cur)
    return segments


def apply_suspension_extension(segments: list, susp_start_idx, susp_end_idx) -> list:
    """
    상각중단이 마지막 구간 안에서 발생한 경우, 그 구간의 종료월(end_idx)을 중단
    기간만큼 늘린다. 상각중단이 마지막이 아닌 중간 구간에서 발생하면(그 구간은 어차피
    다음 이벤트가 먼저 도래해 명목 내용연수가 끝나기 전에 끝나므로) 연장하지 않는다
    (v1 스코프 제한 — README에 명시).
    """
    if susp_start_idx is None:
        return segments
    last = segments[-1]
    if last["start_idx"] <= susp_start_idx <= last["end_idx"]:
        last["end_idx"] += (susp_end_idx - susp_start_idx + 1)
    return segments


def segments_current_period_dep(segments: list, disposal_idx, fy_year: int,
                                 susp_start_idx=None, susp_end_idx=None) -> tuple:
    """method가 섞인 구간 리스트를 받아 당기(fy_year)와 겹치는 상각비를 합산한다."""
    fy_start_idx = month_index(fy_year, 1)
    fy_end_idx = month_index(fy_year, 12)
    if disposal_idx is not None:
        fy_end_idx = min(fy_end_idx, disposal_idx)

    total_dep, total_months = 0.0, 0
    for seg in segments:
        overlap_start = max(seg["start_idx"], fy_start_idx)
        overlap_end = min(seg["end_idx"], fy_end_idx)
        if overlap_start > overlap_end:
            continue

        if seg["method"] == "정액법":
            susp_months = 0
            if susp_start_idx is not None:
                s, e = max(overlap_start, susp_start_idx), min(overlap_end, susp_end_idx)
                if s <= e:
                    susp_months = e - s + 1
            months = (overlap_end - overlap_start + 1) - susp_months
            if months <= 0:
                continue
            monthly_dep = (seg["basis"] - seg["salvage"]) / seg["life_months"]
            total_dep += monthly_dep * months
            total_months += months
        else:
            _, dep, months = _declining_balance_year_loop(
                seg["start_idx"], seg["basis"], seg["salvage"], seg["rate"], seg["end_idx"],
                None, susp_start_idx, susp_end_idx, overlap_end)
            total_dep += dep
            total_months += months
    return total_dep, total_months


# ---------------------------------------------------------------------------
# 생산량비례법 재계산
# ---------------------------------------------------------------------------
def units_of_production_current_period_dep(cost, salvage, total_units, period_units,
                                             disposal, fy_year) -> tuple:
    """
    생산량비례법 당기상각비 = (취득원가-잔존가치) x (당기실제생산량/총예정생산량).
    - 처분일이 당기 이전 회계연도면 이미 상각이 종료된 것으로 보고 0원을 반환한다.
    - 취득 이후 누적생산량 이력은 추적하지 않으므로(당기 값만 입력받는 v1 단순화),
      상각누계액이 상각대상금액(취득원가-잔존가치)을 넘지 않도록만 방어적으로 캡을 건다.
    """
    if disposal is not None and disposal.year < fy_year:
        return 0.0, 0
    dep = (cost - salvage) * (period_units / total_units)
    dep = max(0.0, min(dep, cost - salvage))
    months = 12 if period_units and period_units > 0 else 0
    return dep, months


# ---------------------------------------------------------------------------
# 자산 1건 재계산
# ---------------------------------------------------------------------------
def recalc_asset(acq, cost, salvage, life, method, disposal, reest_date, reest_life, ref_date, fy_year,
                  total_units=None, period_units=None,
                  reest_method=None, capex_date=None, capex_amount=None,
                  susp_start=None, susp_end=None):
    # 자본적지출/상각중단/방법이 실제로 바뀌는 재추정이 하나라도 있으면 새 통합
    # 이벤트 엔진으로 넘어간다. 그 외(기존 52개 테스트가 타는 경로)는 기존
    # 정액법/정률법/생산량비례법 분기를 문자 그대로 그대로 사용한다(회귀 방지).
    has_extended_events = (
        capex_date is not None
        or susp_start is not None
        or (reest_method is not None and reest_method != method)
    )

    active_end_idx = None
    if method in ("정액법", "정률법") and has_extended_events:
        susp_s = month_index(susp_start.year, susp_start.month) if susp_start is not None else None
        susp_e = month_index(susp_end.year, susp_end.month) if susp_end is not None else None
        segments = build_depreciation_schedule(
            acq, cost, salvage, life, method, reest_date, reest_life, reest_method,
            capex_date, capex_amount, susp_s, susp_e)
        segments = apply_suspension_extension(segments, susp_s, susp_e)
        disposal_idx = month_index(disposal.year, disposal.month) if disposal is not None else None
        dep_raw, months = segments_current_period_dep(segments, disposal_idx, fy_year, susp_s, susp_e)
        active_end_idx = segments[-1]["end_idx"]
    elif method == "정액법":
        segments = build_straight_line_segments(acq, cost, salvage, life, reest_date, reest_life)
        disposal_idx = month_index(disposal.year, disposal.month) if disposal is not None else None
        dep_raw, months = straight_line_current_period_dep(segments, disposal_idx, fy_year)
        active_end_idx = segments[-1]["end_idx"]
    elif method == "정률법":
        dep_raw, months = declining_balance_current_period_dep(
            acq, cost, salvage, life, reest_date, reest_life, disposal, fy_year)
        if reest_date is not None:
            active_end_idx = month_index(reest_date.year, 1) + reest_life * 12 - 1
        else:
            active_end_idx = month_index(acq.year, acq.month) + life * 12 - 1
    elif method == "생산량비례법":
        dep_raw, months = units_of_production_current_period_dep(
            cost, salvage, total_units, period_units, disposal, fy_year)
    else:
        raise ValueError(f"알 수 없는 상각방법: {method}")

    if active_end_idx is not None:
        ref_idx = month_index(ref_date.year, ref_date.month)
        life_ended = ref_idx > active_end_idx
    else:
        # 생산량비례법은 누적생산량 이력을 추적하지 않아(v1 단순화) 시간 기준
        # 생애종료 판정을 할 수 없으므로 항상 "종료 아님"으로 본다.
        life_ended = False

    notes = []
    if disposal is not None:
        if disposal.year < fy_year:
            notes.append(f"전기이전처분({disposal.isoformat()})")
        else:
            notes.append(f"당기중처분({disposal.isoformat()})")
    if reest_date is not None:
        method_note = f"→{reest_method}" if (reest_method is not None and reest_method != method) else ""
        notes.append(f"내용연수재추정({reest_date.isoformat()}→{reest_life}년{method_note})")
    if capex_date is not None:
        notes.append(f"자본적지출({capex_date.isoformat()}→{capex_amount:,.0f}원)")
    if susp_start is not None:
        notes.append(f"상각중단({susp_start.isoformat()}~{susp_end.isoformat()})")
    if life_ended and disposal is None:
        notes.append("내용연수종료")
    note = "; ".join(notes) if notes else "-"

    return round_won(dep_raw), months, life_ended, note


def recalc_accumulated_dep(acq, cost, salvage, life, method, disposal, reest_date, reest_life,
                            fy_year, total_units=None, period_units=None,
                            reest_method=None, capex_date=None, capex_amount=None,
                            susp_start=None, susp_end=None):
    """
    취득연도부터 전기(fy_year-1년)까지 매 연도의 재계산 상각비를 합산해 "전기말
    재계산 감가상각누계액"을 구한다. 회사가 제시한 전기말 누계액과 비교해서, 당기에
    갑자기 발생한 차이가 아니라 전기 이전부터 존재하던 차이인지 확인하는 참고용
    지표다(회사반영 당기 감가상각비와의 일치 여부 판정에는 영향을 주지 않는다).

    생산량비례법은 당기 실제생산량만 입력받고 과거 연도별 생산량 이력을 추적하지
    않으므로(v1 한계), 과거 연도별 상각액을 재구성할 수 없어 None을 반환한다.
    """
    if method == "생산량비례법":
        return None
    total = 0
    for y in range(acq.year, fy_year):
        dep, _, _, _ = recalc_asset(
            acq, cost, salvage, life, method, disposal, reest_date, reest_life,
            fy_ref_date(y), y, total_units=total_units, period_units=period_units,
            reest_method=reest_method, capex_date=capex_date, capex_amount=capex_amount,
            susp_start=susp_start, susp_end=susp_end)
        total += dep
    return total


# ---------------------------------------------------------------------------
# 컬럼 인식
# ---------------------------------------------------------------------------
def resolve_columns(df):
    print("=== 컬럼 인식 결과 ===")
    resolved = {}
    missing_required = []
    for key in REQUIRED_KEYS:
        actual = COLUMN_MAP.get(key)
        found = actual in df.columns if actual else False
        print(f"  {key} = {actual}{'' if found else '  [!! 파일에서 찾을 수 없음]'}")
        if not found:
            missing_required.append(key)
        resolved[key] = actual if found else None

    for key in OPTIONAL_KEYS:
        actual = COLUMN_MAP.get(key)
        found = actual in df.columns if actual else False
        tag = "(선택)" if found else "(선택, 파일에 없어 미적용)"
        print(f"  {key} = {actual} {tag}")
        resolved[key] = actual if found else None

    if missing_required:
        raise KeyError(
            "다음 필수 컬럼을 파일에서 찾을 수 없습니다: "
            + ", ".join(missing_required)
            + "  → 스크립트 상단 COLUMN_MAP 에서 실제 컬럼명으로 수정하세요."
        )
    print()
    return resolved


def build_category_summary(result_df: pd.DataFrame, error_df: pd.DataFrame) -> pd.DataFrame:
    """자산분류별 전체/정상/불일치/유의한차이/오류 건수와 불일치율을 집계한다."""
    if len(result_df) > 0:
        normal = result_df.groupby("자산분류").agg(
            정상계산건수=("자산명", "count"),
            불일치건수=("일치여부", lambda s: (s == "불일치").sum()),
            유의한차이건수=("중요성구분", lambda s: (s == "유의한 차이").sum()),
        )
    else:
        normal = pd.DataFrame(columns=["정상계산건수", "불일치건수", "유의한차이건수"])
        normal.index.name = "자산분류"

    if len(error_df) > 0:
        error_counts = error_df.groupby("자산분류").agg(오류건수=("자산명", "count"))
    else:
        error_counts = pd.DataFrame(columns=["오류건수"])
        error_counts.index.name = "자산분류"

    summary = normal.join(error_counts, how="outer").fillna(0)
    for col in ("정상계산건수", "불일치건수", "유의한차이건수", "오류건수"):
        summary[col] = summary[col].astype(int)
    summary["전체건수"] = summary["정상계산건수"] + summary["오류건수"]
    summary["불일치율(%)"] = summary.apply(
        lambda row: round(row["불일치건수"] / row["정상계산건수"] * 100, 1) if row["정상계산건수"] > 0 else 0.0,
        axis=1,
    )
    summary = summary.reset_index().rename(columns={"자산분류": "자산분류"})
    return summary[["자산분류", "전체건수", "정상계산건수", "불일치건수", "불일치율(%)", "유의한차이건수", "오류건수"]]


def parse_asset_row(r, cols) -> dict:
    """엑셀 한 행에서 계산에 필요한 원시값을 파싱해 dict로 반환한다."""
    total_units_raw = r[cols["총예정생산량"]] if cols["총예정생산량"] else None
    period_units_raw = r[cols["당기실제생산량"]] if cols["당기실제생산량"] else None
    reest_life_raw = r[cols["재추정내용연수"]] if cols["재추정내용연수"] else None
    reest_method_raw = r[cols["재추정후상각방법"]] if cols["재추정후상각방법"] else None
    capex_amount_raw = r[cols["자본적지출액"]] if cols["자본적지출액"] else None
    accum_reported_raw = r[cols["회사반영누계상각액"]] if cols["회사반영누계상각액"] else None
    return dict(
        자산명=r[cols["자산명"]],
        자산분류=r[cols["자산분류"]],
        acq=to_date_or_none(r[cols["취득일"]]),
        cost=float(r[cols["취득원가"]]),
        salvage=float(r[cols["잔존가치"]]),
        life=int(r[cols["내용연수"]]),
        method=str(r[cols["상각방법"]]).strip(),
        reported=round_won(r[cols["회사반영상각비"]]),
        disposal=to_date_or_none(r[cols["처분일"]]) if cols["처분일"] else None,
        reest_date=to_date_or_none(r[cols["재추정일"]]) if cols["재추정일"] else None,
        reest_life=int(reest_life_raw) if (reest_life_raw is not None and not pd.isna(reest_life_raw)) else None,
        reest_method=(str(reest_method_raw).strip()
                      if (reest_method_raw is not None and not pd.isna(reest_method_raw)) else None),
        total_units=float(total_units_raw) if (total_units_raw is not None and not pd.isna(total_units_raw)) else None,
        period_units=float(period_units_raw) if (period_units_raw is not None and not pd.isna(period_units_raw)) else None,
        capex_date=to_date_or_none(r[cols["자본적지출일"]]) if cols["자본적지출일"] else None,
        capex_amount=(float(capex_amount_raw)
                      if (capex_amount_raw is not None and not pd.isna(capex_amount_raw)) else None),
        susp_start=to_date_or_none(r[cols["상각중단시작일"]]) if cols["상각중단시작일"] else None,
        susp_end=to_date_or_none(r[cols["상각중단종료일"]]) if cols["상각중단종료일"] else None,
        accum_reported=(float(accum_reported_raw)
                        if (accum_reported_raw is not None and not pd.isna(accum_reported_raw)) else None),
    )


def build_multi_year_trend_df(df: pd.DataFrame, cols: dict, years: list) -> pd.DataFrame:
    """
    자산별로 years에 지정된 각 회계연도의 재계산 상각비를 구해 long format으로 쌓는다.
    자산명이 중복될 수 있으므로 그룹핑용 키는 자산명이 아니라 원본 행 순번(자산ID)을 쓴다.
    데이터 오류(계산 불가능한 값)로 걸러지는 자산은 단일연도 처리와 동일하게 제외한다.
    """
    rows = []
    for asset_id, (_, r) in enumerate(df.iterrows()):
        p = parse_asset_row(r, cols)
        errors = validate_asset_inputs(
            p["cost"], p["salvage"], p["life"], method=p["method"], total_units=p["total_units"],
            reest_method=p["reest_method"], capex_date=p["capex_date"], capex_amount=p["capex_amount"],
            susp_start=p["susp_start"], susp_end=p["susp_end"])
        if errors:
            continue
        for y in years:
            recalced, months, life_ended, note = recalc_asset(
                p["acq"], p["cost"], p["salvage"], p["life"], p["method"], p["disposal"],
                p["reest_date"], p["reest_life"], fy_ref_date(y), y,
                total_units=p["total_units"], period_units=p["period_units"],
                reest_method=p["reest_method"], capex_date=p["capex_date"], capex_amount=p["capex_amount"],
                susp_start=p["susp_start"], susp_end=p["susp_end"])
            rows.append({
                "자산ID": asset_id, "자산명": p["자산명"], "자산분류": p["자산분류"],
                "상각방법": p["method"], "회계연도": y,
                "재계산_당기감가상각비": recalced, "당기해당월수": months, "비고": note,
            })
    return pd.DataFrame(rows)


def detect_yoy_anomalies(trend_df: pd.DataFrame, threshold_pct: float = YOY_ANOMALY_THRESHOLD_PCT) -> pd.DataFrame:
    """
    자산ID별로 회계연도 순 정렬 후 전년대비 증감률을 계산해 이상탐지 컬럼을 추가한다.
    전년도 상각비가 0원이면 %증감이 무의미하므로(0으로 나누기) '신규발생'/'-'로 표기한다.
    """
    df = trend_df.sort_values(["자산ID", "회계연도"]).reset_index(drop=True).copy()
    prev = df.groupby("자산ID")["재계산_당기감가상각비"].shift(1)
    pct = (df["재계산_당기감가상각비"] - prev) / prev.replace(0, float("nan")) * 100
    df["전년대비증감률(%)"] = pct.round(1)

    def _flag(row_prev, row_cur, row_pct):
        if pd.isna(row_prev):
            return "-"
        if row_prev == 0:
            return "신규발생" if row_cur > 0 else "-"
        return "경고" if abs(row_pct) >= threshold_pct else "-"

    df["이상탐지"] = [
        _flag(p, c, pc) for p, c, pc in zip(prev, df["재계산_당기감가상각비"], pct)
    ]
    return df


def main():
    df = pd.read_excel(IN_PATH, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]

    cols = resolve_columns(df)

    error_cols = ["자산명", "자산분류", "취득일", "취득원가", "잔존가치", "내용연수(년)",
                  "상각방법", "회사반영_당기감가상각비", "총예정생산량", "당기실제생산량",
                  "자본적지출일", "자본적지출액", "상각중단시작일", "상각중단종료일",
                  "회사반영_전기말감가상각누계액", "오류사유"]
    rows = []
    error_rows = []
    for _, r in df.iterrows():
        p = parse_asset_row(r, cols)
        acq, cost, salvage, life, method, reported = (
            p["acq"], p["cost"], p["salvage"], p["life"], p["method"], p["reported"])
        disposal, reest_date, reest_life, reest_method = (
            p["disposal"], p["reest_date"], p["reest_life"], p["reest_method"])
        total_units, period_units = p["total_units"], p["period_units"]
        capex_date, capex_amount = p["capex_date"], p["capex_amount"]
        susp_start, susp_end = p["susp_start"], p["susp_end"]
        accum_reported = p["accum_reported"]

        validation_errors = validate_asset_inputs(
            cost, salvage, life, method=method, total_units=total_units,
            reest_method=reest_method, capex_date=capex_date, capex_amount=capex_amount,
            susp_start=susp_start, susp_end=susp_end, accum_reported=accum_reported)
        if validation_errors:
            error_rows.append({
                "자산명": p["자산명"],
                "자산분류": p["자산분류"],
                "취득일": acq,
                "취득원가": cost,
                "잔존가치": salvage,
                "내용연수(년)": life,
                "상각방법": method,
                "회사반영_당기감가상각비": reported,
                "총예정생산량": total_units,
                "당기실제생산량": period_units,
                "자본적지출일": capex_date,
                "자본적지출액": capex_amount,
                "상각중단시작일": susp_start,
                "상각중단종료일": susp_end,
                "회사반영_전기말감가상각누계액": accum_reported,
                "오류사유": "; ".join(validation_errors),
            })
            continue

        elapsed = elapsed_months_to_ref(acq, REF_DATE)

        recalced, cur_months, life_ended, note = recalc_asset(
            acq, cost, salvage, life, method, disposal, reest_date, reest_life, REF_DATE, FY_YEAR,
            total_units=total_units, period_units=period_units,
            reest_method=reest_method, capex_date=capex_date, capex_amount=capex_amount,
            susp_start=susp_start, susp_end=susp_end)

        # 전기말 감가상각누계액 비교(참고용): 회사반영 당기 감가상각비의 일치/중요성
        # 판정에는 영향을 주지 않는다 — 당기 차이(또는 누계액 차이)가 당기에 갑자기
        # 발생한 게 아니라 전기 이전부터 있었는지 확인하는 별도 절차이기 때문이다.
        accum_recalc = recalc_accumulated_dep(
            acq, cost, salvage, life, method, disposal, reest_date, reest_life, FY_YEAR,
            total_units=total_units, period_units=period_units,
            reest_method=reest_method, capex_date=capex_date, capex_amount=capex_amount,
            susp_start=susp_start, susp_end=susp_end)
        accum_diff = (accum_recalc - accum_reported) if (accum_recalc is not None and accum_reported is not None) else None
        if accum_diff is not None:
            accum_match = "일치" if accum_diff == 0 else "불일치"
        else:
            accum_match = "-"

        diff = recalced - reported
        match = "일치" if diff == 0 else "불일치"
        materiality = classify_materiality(diff)

        # 규칙 기반 분류는 비용이 들지 않으므로 "불일치"인 모든 자산에 시도하고,
        # 규칙으로 설명이 안 되는 애매한 경우에 한해서만(그리고 유의한 차이일 때만)
        # 비용이 드는 AI 호출로 폴백한다.
        cause = "-"
        cause_source = "-"
        if match == "불일치":
            rule_cause = get_rule_based_cause(
                reest_date is not None, disposal is not None, note,
                is_capex=capex_date is not None, is_suspended=susp_start is not None)
            if rule_cause is not None:
                cause, cause_source = rule_cause, "규칙"
            elif materiality == "유의한 차이":
                cause = get_ai_estimated_cause(
                    diff, elapsed, method, reest_date is not None, disposal is not None)
                cause_source = "AI" if cause != "-" else "-"

        rows.append({
            "자산명": p["자산명"],
            "자산분류": p["자산분류"],
            "취득일": acq,
            "취득원가": cost,
            "잔존가치": salvage,
            "내용연수(년)": life,
            "상각방법": method,
            "처분일": disposal,
            "내용연수재추정일": reest_date,
            "재추정후내용연수(년)": reest_life,
            "재추정후상각방법": reest_method,
            "총예정생산량": total_units,
            "당기실제생산량": period_units,
            "자본적지출일": capex_date,
            "자본적지출액": capex_amount,
            "상각중단시작일": susp_start,
            "상각중단종료일": susp_end,
            "경과개월수(취득일~기준일)": elapsed,
            "당기해당월수": cur_months,
            "내용연수종료여부": "종료" if life_ended else "-",
            "비고": note,
            "회사반영_당기감가상각비": reported,
            "재계산_당기감가상각비": recalced,
            "차이(재계산-회사반영)": diff,
            "회사반영_전기말감가상각누계액": accum_reported,
            "재계산_전기말감가상각누계액": accum_recalc,
            "누계액차이(재계산-회사반영)": accum_diff,
            "누계액일치여부": accum_match,
            "일치여부": match,
            "중요성구분": materiality,
            "추정원인": cause,
            "추정원인출처": cause_source,
        })

    result_df = pd.DataFrame(rows)
    diff_df = result_df[result_df["일치여부"] == "불일치"].copy()
    diff_df = diff_df.reindex(diff_df["차이(재계산-회사반영)"].abs().sort_values(ascending=False).index)
    # 감사자가 직접 채워 넣는 검토란. diff_df에서 파생되는 material_diff_df에도
    # 그대로 이어지므로("차이자산"/"유의한차이자산" 두 시트 모두 대상) 여기 한 곳에만 추가한다.
    diff_df["검토자"] = ""
    diff_df["검토의견"] = ""
    diff_df["검토일"] = pd.NaT
    material_diff_df = diff_df[diff_df["중요성구분"] == "유의한 차이"].copy()
    error_df = pd.DataFrame(error_rows, columns=error_cols)
    category_summary_df = build_category_summary(result_df, error_df)

    # 다기간(연도별) 비교: COMPARISON_YEARS가 [FY_YEAR] 하나뿐이면(기본값) 이 블록은
    # 건너뛰고, 기존 4개+통계 시트 출력이 이전과 완전히 동일하게 유지된다.
    trend_df = None
    pivot_df = None
    if len(COMPARISON_YEARS) > 1:
        trend_df = detect_yoy_anomalies(build_multi_year_trend_df(df, cols, COMPARISON_YEARS))
        pivot_df = trend_df.pivot_table(
            index=["자산ID", "자산명", "자산분류", "상각방법"],
            columns="회계연도", values="재계산_당기감가상각비",
        ).reset_index()

    with pd.ExcelWriter(OUT_PATH, engine="openpyxl") as writer:
        result_df.to_excel(writer, sheet_name="재계산결과", index=False)
        diff_df.to_excel(writer, sheet_name="차이자산", index=False)
        material_diff_df.to_excel(writer, sheet_name="유의한차이자산", index=False)
        error_df.to_excel(writer, sheet_name="데이터오류", index=False)
        category_summary_df.to_excel(writer, sheet_name="자산군별통계", index=False)
        if trend_df is not None:
            trend_df.to_excel(writer, sheet_name="연도별추이", index=False)
            pivot_df.to_excel(writer, sheet_name="연도별요약", index=False)
        _format_workbook(writer, result_df, diff_df, material_diff_df, error_df, category_summary_df,
                          trend_df, pivot_df)

    print("DONE:", OUT_PATH)
    print(f"기준일: {REF_DATE}, 당기: {FY_YEAR}년, 중요성기준: {MATERIALITY_THRESHOLD:,}원")
    print(f"총 {len(result_df) + len(error_df)}건 중 정상 계산 {len(result_df)}건, "
          f"데이터 오류로 계산 제외 {len(error_df)}건")
    print(f"정상 계산 {len(result_df)}건 중 불일치 {len(diff_df)}건 "
          f"(유의한 차이 {len(material_diff_df)}건 / 경미한 차이 {len(diff_df) - len(material_diff_df)}건)")
    if len(error_df) > 0:
        print(f"[!!] 데이터 오류로 제외된 자산 {len(error_df)}건 → '데이터오류' 시트에서 사유를 확인하세요.")
    print("=== 자산군별 통계 ===")
    print(category_summary_df.to_string(index=False))
    if trend_df is not None:
        warn_count = len(trend_df[trend_df["이상탐지"] == "경고"])
        print(f"=== 다기간 비교({COMPARISON_YEARS}) === "
              f"전년대비 {YOY_ANOMALY_THRESHOLD_PCT}% 이상 증감 경고 {warn_count}건")


def _format_workbook(writer, result_df, diff_df, material_diff_df, error_df, category_summary_df,
                      trend_df=None, pivot_df=None):
    wb = writer.book
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center")
    mismatch_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    error_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    review_input_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    date_cols = ["취득일", "처분일", "내용연수재추정일", "검토일",
                 "자본적지출일", "상각중단시작일", "상각중단종료일"]
    money_cols = ["취득원가", "잔존가치", "회사반영_당기감가상각비", "재계산_당기감가상각비",
                  "차이(재계산-회사반영)", "자본적지출액",
                  "회사반영_전기말감가상각누계액", "재계산_전기말감가상각누계액",
                  "누계액차이(재계산-회사반영)"]

    sheets = [
        ("재계산결과", result_df), ("차이자산", diff_df),
        ("유의한차이자산", material_diff_df), ("데이터오류", error_df),
        ("자산군별통계", category_summary_df),
    ]
    if trend_df is not None:
        sheets += [("연도별추이", trend_df), ("연도별요약", pivot_df)]

    for sheet_name, df in sheets:
        ws = wb[sheet_name]
        n_rows, n_cols = df.shape

        for col_idx in range(1, n_cols + 1):
            c = ws.cell(row=1, column=col_idx)
            c.font = header_font
            c.fill = header_fill
            c.alignment = header_align

        col_names = list(df.columns)
        for col_idx, col_name in enumerate(col_names, start=1):
            letter = get_column_letter(col_idx)
            # "연도별요약"은 회계연도(정수)가 컬럼 헤더가 되는 pivot 결과라 money_cols
            # 이름 목록으로는 못 잡으므로, 이 시트에서는 정수 헤더를 금액으로 취급한다.
            is_pivot_year_col = sheet_name == "연도별요약" and isinstance(col_name, int)
            if col_name in date_cols:
                for row_idx in range(2, n_rows + 2):
                    ws.cell(row=row_idx, column=col_idx).number_format = "yyyy-mm-dd"
            elif col_name in money_cols or is_pivot_year_col:
                for row_idx in range(2, n_rows + 2):
                    ws.cell(row=row_idx, column=col_idx).number_format = "#,##0"
            if col_name in ("추정원인", "오류사유", "검토의견"):
                ws.column_dimensions[letter].width = 60
            else:
                ws.column_dimensions[letter].width = max(14, len(str(col_name)) + 4)

        if "일치여부" in col_names and n_rows > 0:
            match_col_idx = col_names.index("일치여부") + 1
            for row_idx in range(2, n_rows + 2):
                cell = ws.cell(row=row_idx, column=match_col_idx)
                if cell.value == "불일치":
                    for c_idx in range(1, n_cols + 1):
                        ws.cell(row=row_idx, column=c_idx).fill = mismatch_fill

        if sheet_name == "데이터오류" and n_rows > 0:
            for row_idx in range(2, n_rows + 2):
                for c_idx in range(1, n_cols + 1):
                    ws.cell(row=row_idx, column=c_idx).fill = error_fill

        # 검토란(검토자/검토의견/검토일)은 감사자가 직접 입력하는 영역이므로,
        # 위의 "불일치 행 전체 빨간색" 칠보다 뒤에 별도 색으로 덧칠해 구분한다.
        if sheet_name in ("차이자산", "유의한차이자산") and n_rows > 0:
            for col_name in ("검토자", "검토의견", "검토일"):
                if col_name in col_names:
                    col_idx = col_names.index(col_name) + 1
                    for row_idx in range(2, n_rows + 2):
                        ws.cell(row=row_idx, column=col_idx).fill = review_input_fill

        ws.freeze_panes = "A2"
        if n_rows > 0:
            ws.auto_filter.ref = f"A1:{get_column_letter(n_cols)}{n_rows + 1}"


if __name__ == "__main__":
    main()
