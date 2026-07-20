# -*- coding: utf-8 -*-
"""
recalc.py의 계산 함수들에 대한 pytest 테스트.
각 케이스의 정답값은 손계산(또는 독립적인 수식)으로 먼저 도출한 뒤,
recalc.py의 실제 함수 출력과 비교한다.
"""
import datetime as dt

import pandas as pd

import recalc as R


# ---------------------------------------------------------------------------
# 1) 월할상각 경과개월수 계산 (elapsed_months_to_ref / month_index)
# ---------------------------------------------------------------------------
class TestElapsedMonths:
    def test_month_index_basic(self):
        # 2025년 12월 = 2025*12 + 12
        assert R.month_index(2025, 12) == 2025 * 12 + 12

    def test_elapsed_months_same_year(self):
        # 2025-03-01 ~ 2025-12-31: 3월~12월 취득월 포함 10개월
        assert R.elapsed_months_to_ref(dt.date(2025, 3, 1), dt.date(2025, 12, 31)) == 10

    def test_elapsed_months_multi_year(self):
        # 2023-01-15 ~ 2025-12-31: 2023-01부터 2025-12까지 취득월 포함 36개월(만 3년)
        assert R.elapsed_months_to_ref(dt.date(2023, 1, 15), dt.date(2025, 12, 31)) == 36

    def test_elapsed_months_acquired_in_ref_month(self):
        # 취득월 = 기준월이면 경과개월수는 1개월
        assert R.elapsed_months_to_ref(dt.date(2025, 12, 1), dt.date(2025, 12, 31)) == 1


# ---------------------------------------------------------------------------
# 2) 정액법 재계산
# ---------------------------------------------------------------------------
class TestStraightLine:
    def test_basic_full_year(self):
        # 취득 2023-01-01, 취득원가 12,000,000, 잔존가치 0, 내용연수 5년
        # 월상각액 = 12,000,000 / 60 = 200,000, 당기(2025) 12개월 전액 해당
        # 재계산 당기상각비 = 200,000 * 12 = 2,400,000
        dep, months, life_ended, note = R.recalc_asset(
            acq=dt.date(2023, 1, 1), cost=12_000_000, salvage=0, life=5, method="정액법",
            disposal=None, reest_date=None, reest_life=None,
            ref_date=dt.date(2025, 12, 31), fy_year=2025,
        )
        assert dep == 2_400_000
        assert months == 12
        assert life_ended is False
        assert note == "-"

    def test_partial_first_year(self):
        # 취득 2025-07-01, 취득원가 6,000,000, 잔존가치 0, 내용연수 5년(60개월)
        # 월상각액 = 100,000, 당기(2025)는 취득월(7월)부터 12월까지 6개월만 해당
        # 재계산 당기상각비 = 100,000 * 6 = 600,000
        dep, months, life_ended, note = R.recalc_asset(
            acq=dt.date(2025, 7, 1), cost=6_000_000, salvage=0, life=5, method="정액법",
            disposal=None, reest_date=None, reest_life=None,
            ref_date=dt.date(2025, 12, 31), fy_year=2025,
        )
        assert dep == 600_000
        assert months == 6
        assert life_ended is False


# ---------------------------------------------------------------------------
# 3) 정률법 재계산
# ---------------------------------------------------------------------------
class TestDecliningBalance:
    def test_third_year_of_depreciation(self):
        # 취득 2023-01-01, 취득원가 10,000,000, 잔존가치 0, 내용연수 5년 (상각률표 0.451)
        # 2023년: 10,000,000 * 0.451 = 4,510,000 -> 기말 장부가액 5,490,000
        # 2024년: 5,490,000 * 0.451 = 2,475,990 -> 기말 장부가액 3,014,010
        # 2025년(당기): 3,014,010 * 0.451 = 1,359,318.51 -> 반올림 1,359,319
        dep, months, life_ended, note = R.recalc_asset(
            acq=dt.date(2023, 1, 1), cost=10_000_000, salvage=0, life=5, method="정률법",
            disposal=None, reest_date=None, reest_life=None,
            ref_date=dt.date(2025, 12, 31), fy_year=2025,
        )
        assert dep == 1_359_319
        assert months == 12
        assert life_ended is False

    def test_get_rate_table_lookup(self):
        assert R.get_rate(5) == 0.451
        assert R.get_rate(10) == 0.259

    def test_get_rate_fallback_formula(self):
        # 상각률표에 없는 내용연수(예: 25년)는 잔존가치 5% 가정 근사식 사용
        life = 25
        expected = round(1 - 0.05 ** (1 / life), 3)
        assert R.get_rate(life) == expected


# ---------------------------------------------------------------------------
# 4) 처분자산 처리 (당기 중 처분 시 처분일까지만 상각)
# ---------------------------------------------------------------------------
class TestDisposal:
    def test_straight_line_disposal_mid_year(self):
        # 취득 2022-06-01, 취득원가 24,000,000, 잔존가치 0, 내용연수 5년
        # 월상각액 = 400,000, 2025-06-30 처분 -> 당기 1~6월 6개월만 상각
        # 재계산 당기상각비 = 400,000 * 6 = 2,400,000
        dep, months, life_ended, note = R.recalc_asset(
            acq=dt.date(2022, 6, 1), cost=24_000_000, salvage=0, life=5, method="정액법",
            disposal=dt.date(2025, 6, 30), reest_date=None, reest_life=None,
            ref_date=dt.date(2025, 12, 31), fy_year=2025,
        )
        assert dep == 2_400_000
        assert months == 6
        assert "당기중처분(2025-06-30)" in note

    def test_disposal_before_current_year(self):
        # 처분일이 당기 이전이면 "전기이전처분" 비고 + 당기 상각비 0
        dep, months, life_ended, note = R.recalc_asset(
            acq=dt.date(2020, 1, 1), cost=12_000_000, salvage=0, life=10, method="정액법",
            disposal=dt.date(2024, 3, 31), reest_date=None, reest_life=None,
            ref_date=dt.date(2025, 12, 31), fy_year=2025,
        )
        assert dep == 0
        assert months == 0
        assert "전기이전처분(2024-03-31)" in note


# ---------------------------------------------------------------------------
# 5) 내용연수 재추정 처리
# ---------------------------------------------------------------------------
class TestReestimation:
    def test_straight_line_reestimation(self):
        # 취득 2021-01-01, 취득원가 10,000,000, 잔존가치 0, 원내용연수 10년
        # 2025-01-01에 잔여내용연수를 3년으로 재추정
        # 재추정 시점까지 경과 48개월 상각: 10,000,000/120*48 = 4,000,000 -> 재추정시 장부가 6,000,000
        # 재추정 후 월상각액 = 6,000,000 / 36 = 166,666.67, 당기(2025) 12개월 전액 해당
        # 재계산 당기상각비 = 6,000,000 / 36 * 12 = 2,000,000
        dep, months, life_ended, note = R.recalc_asset(
            acq=dt.date(2021, 1, 1), cost=10_000_000, salvage=0, life=10, method="정액법",
            disposal=None, reest_date=dt.date(2025, 1, 1), reest_life=3,
            ref_date=dt.date(2025, 12, 31), fy_year=2025,
        )
        assert dep == 2_000_000
        assert months == 12
        assert life_ended is False
        assert "내용연수재추정(2025-01-01→3년)" in note


# ---------------------------------------------------------------------------
# 6) 내용연수 종료 처리
# ---------------------------------------------------------------------------
class TestLifeEnded:
    def test_straight_line_life_ended(self):
        # 취득 2018-01-01, 내용연수 5년 -> 2022-12에 상각 종료, 당기(2025)는 상각 대상 아님
        dep, months, life_ended, note = R.recalc_asset(
            acq=dt.date(2018, 1, 1), cost=5_000_000, salvage=0, life=5, method="정액법",
            disposal=None, reest_date=None, reest_life=None,
            ref_date=dt.date(2025, 12, 31), fy_year=2025,
        )
        assert dep == 0
        assert months == 0
        assert life_ended is True
        assert note == "내용연수종료"

    def test_declining_balance_life_ended(self):
        # 취득 2015-01-01, 내용연수 5년 -> 2019-12에 상각 종료, 당기(2025)는 상각 대상 아님
        dep, months, life_ended, note = R.recalc_asset(
            acq=dt.date(2015, 1, 1), cost=3_000_000, salvage=0, life=5, method="정률법",
            disposal=None, reest_date=None, reest_life=None,
            ref_date=dt.date(2025, 12, 31), fy_year=2025,
        )
        assert dep == 0
        assert months == 0
        assert life_ended is True
        assert note == "내용연수종료"


# ---------------------------------------------------------------------------
# 7) 중요성 기준 판정
# ---------------------------------------------------------------------------
class TestMateriality:
    def test_diff_zero_is_minor(self):
        assert R.classify_materiality(0) == "경미한 차이"

    def test_diff_just_below_threshold_is_minor(self):
        assert R.classify_materiality(999_999) == "경미한 차이"

    def test_diff_at_threshold_is_material(self):
        # 정확히 임계값(기본 1,000,000원)이면 "유의한 차이" (>= 판정)
        assert R.classify_materiality(1_000_000) == "유의한 차이"

    def test_diff_above_threshold_is_material(self):
        assert R.classify_materiality(1_500_000) == "유의한 차이"

    def test_negative_diff_uses_absolute_value(self):
        assert R.classify_materiality(-1_500_000) == "유의한 차이"
        assert R.classify_materiality(-500_000) == "경미한 차이"

    def test_custom_threshold(self):
        assert R.classify_materiality(300_000, threshold=500_000) == "경미한 차이"
        assert R.classify_materiality(500_000, threshold=500_000) == "유의한 차이"


# ---------------------------------------------------------------------------
# 8) 입력값 검증 (validate_asset_inputs) - 계산 불가능한 값은 "데이터 오류"로 분리
# ---------------------------------------------------------------------------
class TestValidateAssetInputs:
    def test_life_zero_is_invalid(self):
        errors = R.validate_asset_inputs(cost=5_000_000, salvage=0, life=0)
        assert any("내용연수" in e for e in errors)

    def test_life_negative_is_invalid(self):
        errors = R.validate_asset_inputs(cost=8_000_000, salvage=0, life=-5)
        assert any("내용연수" in e for e in errors)

    def test_salvage_greater_than_cost_is_invalid(self):
        errors = R.validate_asset_inputs(cost=10_000_000, salvage=15_000_000, life=5)
        assert any("잔존가치" in e for e in errors)

    def test_salvage_equal_to_cost_is_invalid(self):
        # 잔존가치 == 취득원가면 상각대상금액이 0이 되어 이상하므로 오류로 취급(>= 판정)
        errors = R.validate_asset_inputs(cost=10_000_000, salvage=10_000_000, life=5)
        assert any("잔존가치" in e for e in errors)

    def test_cost_zero_is_invalid(self):
        errors = R.validate_asset_inputs(cost=0, salvage=0, life=5)
        assert any("취득원가" in e for e in errors)

    def test_cost_negative_is_invalid(self):
        errors = R.validate_asset_inputs(cost=-1_000_000, salvage=0, life=5)
        assert any("취득원가" in e for e in errors)

    def test_valid_inputs_return_no_errors(self):
        assert R.validate_asset_inputs(cost=10_000_000, salvage=1_000_000, life=5) == []


# ---------------------------------------------------------------------------
# 보조 함수: round_won (원 단위 반올림, ROUND_HALF_UP)
# ---------------------------------------------------------------------------
class TestRoundWon:
    def test_round_half_up(self):
        assert R.round_won(1_234_567.5) == 1_234_568

    def test_round_down_below_half(self):
        assert R.round_won(1_234_567.4) == 1_234_567

    def test_round_down_just_below_half(self):
        assert R.round_won(1_234_567.49) == 1_234_567


# ---------------------------------------------------------------------------
# 9) 생산량비례법 (units_of_production_current_period_dep)
# ---------------------------------------------------------------------------
class TestUnitsOfProduction:
    def test_basic_ratio(self):
        # 상각대상금액 9,000,000 x (2,000/10,000) = 1,800,000
        dep, months = R.units_of_production_current_period_dep(
            cost=10_000_000, salvage=1_000_000, total_units=10_000, period_units=2_000,
            disposal=None, fy_year=2025)
        assert dep == 1_800_000
        assert months == 12

    def test_disposed_before_current_year_returns_zero(self):
        dep, months = R.units_of_production_current_period_dep(
            cost=10_000_000, salvage=0, total_units=10_000, period_units=2_000,
            disposal=dt.date(2024, 6, 1), fy_year=2025)
        assert dep == 0.0
        assert months == 0

    def test_zero_period_units_returns_zero_months(self):
        dep, months = R.units_of_production_current_period_dep(
            cost=10_000_000, salvage=0, total_units=10_000, period_units=0,
            disposal=None, fy_year=2025)
        assert dep == 0.0
        assert months == 0

    def test_dep_capped_at_depreciable_amount(self):
        # 당기실제생산량이 총예정생산량을 초과해도 상각대상금액(9,000,000)을 넘지 않는다
        dep, _ = R.units_of_production_current_period_dep(
            cost=10_000_000, salvage=1_000_000, total_units=10_000, period_units=15_000,
            disposal=None, fy_year=2025)
        assert dep == 9_000_000


class TestRecalcAssetUnitsOfProduction:
    def test_recalc_asset_units_of_production(self):
        dep, months, life_ended, note = R.recalc_asset(
            acq=dt.date(2023, 1, 1), cost=10_000_000, salvage=1_000_000, life=5,
            method="생산량비례법", disposal=None, reest_date=None, reest_life=None,
            ref_date=dt.date(2025, 12, 31), fy_year=2025,
            total_units=10_000, period_units=2_000,
        )
        assert dep == 1_800_000
        assert months == 12
        assert life_ended is False  # v1: 생산량비례법은 누적이력 미추적으로 항상 종료 아님


class TestValidateAssetInputsUnitsOfProduction:
    def test_existing_calls_without_method_unaffected(self):
        # method 인자를 안 넘기는 기존 8개 케이스와 동일하게 동작해야 한다(회귀 확인).
        assert R.validate_asset_inputs(cost=10_000_000, salvage=1_000_000, life=5) == []
        errors = R.validate_asset_inputs(cost=5_000_000, salvage=0, life=0)
        assert any("내용연수" in e for e in errors)

    def test_life_ignored_for_units_of_production(self):
        errors = R.validate_asset_inputs(
            cost=10_000_000, salvage=1_000_000, life=0,
            method="생산량비례법", total_units=10_000)
        assert not any("내용연수" in e for e in errors)

    def test_total_units_missing_is_invalid(self):
        errors = R.validate_asset_inputs(
            cost=10_000_000, salvage=1_000_000, life=5,
            method="생산량비례법", total_units=None)
        assert any("총예정생산량" in e for e in errors)

    def test_total_units_zero_is_invalid(self):
        errors = R.validate_asset_inputs(
            cost=10_000_000, salvage=1_000_000, life=5,
            method="생산량비례법", total_units=0)
        assert any("총예정생산량" in e for e in errors)

    def test_valid_units_of_production_inputs(self):
        errors = R.validate_asset_inputs(
            cost=10_000_000, salvage=1_000_000, life=5,
            method="생산량비례법", total_units=10_000)
        assert errors == []


# ---------------------------------------------------------------------------
# 10) 자산군별 통계 (build_category_summary)
# ---------------------------------------------------------------------------
class TestCategorySummary:
    def test_summary_counts_and_mismatch_rate(self):
        result_df = pd.DataFrame([
            {"자산명": "A", "자산분류": "유형자산", "일치여부": "일치", "중요성구분": "경미한 차이"},
            {"자산명": "B", "자산분류": "유형자산", "일치여부": "불일치", "중요성구분": "유의한 차이"},
            {"자산명": "C", "자산분류": "무형자산", "일치여부": "일치", "중요성구분": "경미한 차이"},
        ])
        error_df = pd.DataFrame([
            {"자산명": "D", "자산분류": "유형자산"},
        ])
        summary = R.build_category_summary(result_df, error_df).set_index("자산분류")

        assert summary.loc["유형자산", "전체건수"] == 3
        assert summary.loc["유형자산", "정상계산건수"] == 2
        assert summary.loc["유형자산", "불일치건수"] == 1
        assert summary.loc["유형자산", "불일치율(%)"] == 50.0
        assert summary.loc["유형자산", "유의한차이건수"] == 1
        assert summary.loc["유형자산", "오류건수"] == 1

        assert summary.loc["무형자산", "전체건수"] == 1
        assert summary.loc["무형자산", "불일치율(%)"] == 0.0
        assert summary.loc["무형자산", "오류건수"] == 0

    def test_empty_error_df(self):
        result_df = pd.DataFrame([
            {"자산명": "A", "자산분류": "유형자산", "일치여부": "일치", "중요성구분": "경미한 차이"},
        ])
        error_df = pd.DataFrame(columns=["자산명", "자산분류"])
        summary = R.build_category_summary(result_df, error_df).set_index("자산분류")
        assert summary.loc["유형자산", "오류건수"] == 0
        assert summary.loc["유형자산", "전체건수"] == 1


# ---------------------------------------------------------------------------
# 11) 룰 기반 원인 분류 (get_rule_based_cause)
# ---------------------------------------------------------------------------
class TestRuleBasedCause:
    def test_disposed_current_year(self):
        cause = R.get_rule_based_cause(
            is_reestimated=False, is_disposed=True, note="당기중처분(2025-06-30)")
        assert cause is not None and "처분일" in cause

    def test_disposed_prior_year(self):
        cause = R.get_rule_based_cause(
            is_reestimated=False, is_disposed=True, note="전기이전처분(2024-03-31)")
        assert cause is not None and "전기" in cause

    def test_reestimated(self):
        cause = R.get_rule_based_cause(
            is_reestimated=True, is_disposed=False, note="내용연수재추정(2025-01-01→8년)")
        assert cause is not None and "재추정" in cause

    def test_life_ended(self):
        cause = R.get_rule_based_cause(
            is_reestimated=False, is_disposed=False, note="내용연수종료")
        assert cause is not None and "종료" in cause

    def test_ambiguous_case_returns_none(self):
        # 처분/재추정/내용연수종료 어느 쪽에도 해당하지 않으면 규칙으로 설명할 수 없으므로
        # None을 반환해 호출부가 AI 호출로 폴백하도록 한다.
        assert R.get_rule_based_cause(is_reestimated=False, is_disposed=False, note="-") is None


# ---------------------------------------------------------------------------
# 12) 다기간(연도별) 비교 (fy_ref_date / build_multi_year_trend_df / detect_yoy_anomalies)
# ---------------------------------------------------------------------------
def _make_cols(df):
    required = ["자산명", "자산분류", "취득일", "취득원가", "잔존가치", "내용연수", "상각방법", "회사반영상각비"]
    optional = ["처분일", "재추정일", "재추정내용연수", "총예정생산량", "당기실제생산량"]
    cols = {k: k for k in required}
    cols.update({k: (k if k in df.columns else None) for k in optional})
    return cols


class TestFyRefDate:
    def test_current_fy_year_uses_ref_date(self):
        assert R.fy_ref_date(R.FY_YEAR) == R.REF_DATE

    def test_other_year_uses_dec_31(self):
        assert R.fy_ref_date(2023) == dt.date(2023, 12, 31)


class TestMultiYearTrend:
    def _sample_df(self):
        # 자산A: 정액법, 취득 2023-01-01, 연간상각액 2,400,000(취득~기준일까지 매년 동일)
        # 자산B/자산C: 이름이 둘 다 "차량"으로 중복되지만 서로 다른 자산 - 그룹핑 키가
        # 자산명이 아니라 자산ID(행 순번)여야 서로 합쳐지지 않는지 검증한다.
        # 자산B: 취득 2024-01-01(2023년엔 미취득 → 0원), 연간상각액 2,000,000
        # 자산C: 취득 2023-01-01, 내용연수 2년 → 2025년엔 내용연수 종료로 0원
        return pd.DataFrame([
            {"자산명": "본사건물", "자산분류": "유형자산", "취득일": dt.date(2023, 1, 1),
             "취득원가": 12_000_000, "잔존가치": 0, "내용연수": 5, "상각방법": "정액법",
             "회사반영상각비": 0},
            {"자산명": "차량", "자산분류": "유형자산", "취득일": dt.date(2024, 1, 1),
             "취득원가": 6_000_000, "잔존가치": 0, "내용연수": 3, "상각방법": "정액법",
             "회사반영상각비": 0},
            {"자산명": "차량", "자산분류": "유형자산", "취득일": dt.date(2023, 1, 1),
             "취득원가": 4_000_000, "잔존가치": 0, "내용연수": 2, "상각방법": "정액법",
             "회사반영상각비": 0},
        ])

    def test_matches_recalc_asset_per_year(self):
        df = self._sample_df()
        cols = _make_cols(df)
        trend_df = R.build_multi_year_trend_df(df, cols, [2023, 2024, 2025])

        asset_a = trend_df[trend_df["자산명"] == "본사건물"].set_index("회계연도")
        assert asset_a.loc[2023, "재계산_당기감가상각비"] == 2_400_000
        assert asset_a.loc[2024, "재계산_당기감가상각비"] == 2_400_000
        assert asset_a.loc[2025, "재계산_당기감가상각비"] == 2_400_000

    def test_duplicate_names_kept_separate_by_asset_id(self):
        df = self._sample_df()
        cols = _make_cols(df)
        trend_df = R.build_multi_year_trend_df(df, cols, [2023, 2024, 2025])

        cha_rows = trend_df[trend_df["자산명"] == "차량"]
        assert cha_rows["자산ID"].nunique() == 2  # 이름은 같아도 별개 자산으로 유지

        by_id = cha_rows.set_index(["자산ID", "회계연도"])["재계산_당기감가상각비"]
        asset_b_id = cha_rows[cha_rows["회계연도"] == 2023]
        asset_b_id = asset_b_id[asset_b_id["재계산_당기감가상각비"] == 0]["자산ID"].iloc[0]
        asset_c_id = cha_rows[cha_rows["회계연도"] == 2023]
        asset_c_id = asset_c_id[asset_c_id["재계산_당기감가상각비"] == 2_000_000]["자산ID"].iloc[0]

        assert by_id[(asset_b_id, 2023)] == 0            # 자산B: 2024년 취득 전이라 0원
        assert by_id[(asset_b_id, 2024)] == 2_000_000
        assert by_id[(asset_c_id, 2024)] == 2_000_000
        assert by_id[(asset_c_id, 2025)] == 0             # 자산C: 내용연수(2년) 종료로 0원

    def test_detect_yoy_anomalies_flags(self):
        df = self._sample_df()
        cols = _make_cols(df)
        trend_df = R.build_multi_year_trend_df(df, cols, [2023, 2024, 2025])
        flagged = R.detect_yoy_anomalies(trend_df, threshold_pct=20.0)

        # 자산A: 매년 동일 금액 → 첫 연도는 "-"(비교대상 없음), 이후는 변동 없어 "-"
        asset_a = flagged[flagged["자산명"] == "본사건물"].set_index("회계연도")
        assert asset_a.loc[2023, "이상탐지"] == "-"
        assert asset_a.loc[2024, "이상탐지"] == "-"

        cha_rows = flagged[flagged["자산명"] == "차량"]
        # 자산B: 0원 → 2,000,000원으로 "신규발생"
        new_asset = cha_rows[(cha_rows["회계연도"] == 2024) & (cha_rows["재계산_당기감가상각비"] == 2_000_000)]
        prev_year_of_new_asset = cha_rows[
            (cha_rows["자산ID"] == new_asset["자산ID"].iloc[0]) & (cha_rows["회계연도"] == 2023)]
        assert prev_year_of_new_asset["재계산_당기감가상각비"].iloc[0] == 0
        assert new_asset["이상탐지"].iloc[0] == "신규발생"

        # 자산C: 2,000,000원 → 0원으로 급감 → "경고"(20% 임계치 초과)
        declined = cha_rows[(cha_rows["회계연도"] == 2025) & (cha_rows["재계산_당기감가상각비"] == 0)]
        assert declined["이상탐지"].iloc[0] == "경고"
