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
"""
import datetime as dt
import sys
from decimal import Decimal, ROUND_HALF_UP

import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# Windows 콘솔(PowerShell/cmd)이 시스템 코드페이지(cp949 등)로 표준출력을 열어
# 한글 안내 메시지가 깨지는 것을 막기 위해 표준출력 인코딩을 UTF-8로 고정한다.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

IN_PATH = r"C:\Users\wh981\재무검증도구\depreciation-recalc-tool\sample_asset_ledger.xlsx"
OUT_PATH = r"C:\Users\wh981\재무검증도구\depreciation-recalc-tool\recalc_result.xlsx"

# 기준일: 이 날짜가 속한 회계연도(1/1~12/31)의 당기 감가상각비를 재계산한다.
REF_DATE = dt.date(2025, 12, 31)
FY_YEAR = REF_DATE.year

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
}

REQUIRED_KEYS = [
    "자산명", "자산분류", "취득일", "취득원가", "잔존가치",
    "내용연수", "상각방법", "회사반영상각비",
]
OPTIONAL_KEYS = ["처분일", "재추정일", "재추정내용연수"]

# 국세청 고시 내용연수별 상각률표(정률법). 표에 없는 내용연수는
# 잔존가치 5% 가정(r = 1-0.05**(1/n))으로 근사한다.
RATE_TABLE = {
    2: 0.684, 3: 0.536, 4: 0.528, 5: 0.451, 6: 0.394,
    7: 0.349, 8: 0.313, 9: 0.284, 10: 0.259,
    11: 0.239, 12: 0.221, 13: 0.206, 14: 0.193, 15: 0.183,
    16: 0.173, 17: 0.164, 18: 0.157, 19: 0.150, 20: 0.144,
}


def get_rate(life: int) -> float:
    if life in RATE_TABLE:
        return RATE_TABLE[life]
    return round(1 - 0.05 ** (1 / life), 3)


def month_index(year: int, month: int) -> int:
    return year * 12 + month


def round_won(x: float) -> int:
    return int(Decimal(str(x)).quantize(0, rounding=ROUND_HALF_UP))


def to_date_or_none(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or pd.isna(v):
        return None
    return pd.to_datetime(v).date()


def elapsed_months_to_ref(acq: dt.date, ref: dt.date) -> int:
    """취득일부터 기준일까지 경과 개월수(취득월 포함, 월 단위 카운트)."""
    return month_index(ref.year, ref.month) - month_index(acq.year, acq.month) + 1


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
# 자산 1건 재계산
# ---------------------------------------------------------------------------
def recalc_asset(acq, cost, salvage, life, method, disposal, reest_date, reest_life, ref_date, fy_year):
    if method == "정액법":
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
    else:
        raise ValueError(f"알 수 없는 상각방법: {method}")

    ref_idx = month_index(ref_date.year, ref_date.month)
    life_ended = ref_idx > active_end_idx

    notes = []
    if disposal is not None:
        if disposal.year < fy_year:
            notes.append(f"전기이전처분({disposal.isoformat()})")
        else:
            notes.append(f"당기중처분({disposal.isoformat()})")
    if reest_date is not None:
        notes.append(f"내용연수재추정({reest_date.isoformat()}→{reest_life}년)")
    if life_ended and disposal is None:
        notes.append("내용연수종료")
    note = "; ".join(notes) if notes else "-"

    return round_won(dep_raw), months, life_ended, note


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


def main():
    df = pd.read_excel(IN_PATH, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]

    cols = resolve_columns(df)

    rows = []
    for _, r in df.iterrows():
        acq = to_date_or_none(r[cols["취득일"]])
        cost = float(r[cols["취득원가"]])
        salvage = float(r[cols["잔존가치"]])
        life = int(r[cols["내용연수"]])
        method = str(r[cols["상각방법"]]).strip()
        reported = round_won(r[cols["회사반영상각비"]])

        disposal = to_date_or_none(r[cols["처분일"]]) if cols["처분일"] else None
        reest_date = to_date_or_none(r[cols["재추정일"]]) if cols["재추정일"] else None
        reest_life_raw = r[cols["재추정내용연수"]] if cols["재추정내용연수"] else None
        reest_life = int(reest_life_raw) if (reest_life_raw is not None and not pd.isna(reest_life_raw)) else None

        elapsed = elapsed_months_to_ref(acq, REF_DATE)

        recalced, cur_months, life_ended, note = recalc_asset(
            acq, cost, salvage, life, method, disposal, reest_date, reest_life, REF_DATE, FY_YEAR)

        diff = recalced - reported
        match = "일치" if diff == 0 else "불일치"

        rows.append({
            "자산명": r[cols["자산명"]],
            "자산분류": r[cols["자산분류"]],
            "취득일": acq,
            "취득원가": cost,
            "잔존가치": salvage,
            "내용연수(년)": life,
            "상각방법": method,
            "처분일": disposal,
            "내용연수재추정일": reest_date,
            "재추정후내용연수(년)": reest_life,
            "경과개월수(취득일~기준일)": elapsed,
            "당기해당월수": cur_months,
            "내용연수종료여부": "종료" if life_ended else "-",
            "비고": note,
            "회사반영_당기감가상각비": reported,
            "재계산_당기감가상각비": recalced,
            "차이(재계산-회사반영)": diff,
            "일치여부": match,
        })

    result_df = pd.DataFrame(rows)
    diff_df = result_df[result_df["일치여부"] == "불일치"].copy()
    diff_df = diff_df.reindex(diff_df["차이(재계산-회사반영)"].abs().sort_values(ascending=False).index)

    with pd.ExcelWriter(OUT_PATH, engine="openpyxl") as writer:
        result_df.to_excel(writer, sheet_name="재계산결과", index=False)
        diff_df.to_excel(writer, sheet_name="차이자산", index=False)
        _format_workbook(writer, result_df, diff_df)

    print("DONE:", OUT_PATH)
    print(f"기준일: {REF_DATE}, 당기: {FY_YEAR}년, 총 {len(result_df)}건 중 불일치 {len(diff_df)}건")


def _format_workbook(writer, result_df, diff_df):
    wb = writer.book
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center")
    mismatch_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    date_cols = ["취득일", "처분일", "내용연수재추정일"]
    money_cols = ["취득원가", "잔존가치", "회사반영_당기감가상각비", "재계산_당기감가상각비", "차이(재계산-회사반영)"]

    for sheet_name, df in (("재계산결과", result_df), ("차이자산", diff_df)):
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
            if col_name in date_cols:
                for row_idx in range(2, n_rows + 2):
                    ws.cell(row=row_idx, column=col_idx).number_format = "yyyy-mm-dd"
            elif col_name in money_cols:
                for row_idx in range(2, n_rows + 2):
                    ws.cell(row=row_idx, column=col_idx).number_format = "#,##0"
            ws.column_dimensions[letter].width = max(14, len(col_name) + 4)

        if "일치여부" in col_names and n_rows > 0:
            match_col_idx = col_names.index("일치여부") + 1
            for row_idx in range(2, n_rows + 2):
                cell = ws.cell(row=row_idx, column=match_col_idx)
                if cell.value == "불일치":
                    for c_idx in range(1, n_cols + 1):
                        ws.cell(row=row_idx, column=c_idx).fill = mismatch_fill

        ws.freeze_panes = "A2"
        if n_rows > 0:
            ws.auto_filter.ref = f"A1:{get_column_letter(n_cols)}{n_rows + 1}"


if __name__ == "__main__":
    main()
