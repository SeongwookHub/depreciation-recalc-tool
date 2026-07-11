# -*- coding: utf-8 -*-
"""
recalc.py의 계산 함수들에 대한 pytest 테스트.
각 케이스의 정답값은 손계산(또는 독립적인 수식)으로 먼저 도출한 뒤,
recalc.py의 실제 함수 출력과 비교한다.
"""
import datetime as dt

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
# 보조 함수: round_won (원 단위 반올림, ROUND_HALF_UP)
# ---------------------------------------------------------------------------
class TestRoundWon:
    def test_round_half_up(self):
        assert R.round_won(1_234_567.5) == 1_234_568

    def test_round_down_below_half(self):
        assert R.round_won(1_234_567.4) == 1_234_567

    def test_round_down_just_below_half(self):
        assert R.round_won(1_234_567.49) == 1_234_567
