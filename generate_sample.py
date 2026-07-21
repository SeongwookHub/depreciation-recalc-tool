# -*- coding: utf-8 -*-
"""
sample_asset_ledger.xlsx 생성 스크립트(재현용 — 코드가 바뀌어 재계산값이 달라지면
다시 실행해 정답값을 새로 채워 넣는 용도로 리포지토리에 남겨둔다).
recalc.py의 recalc_asset을 그대로 호출해서 "정상"으로 분류될 행의 회사반영금액을
재계산값과 동일하게 채우고(=일치), 일부 행만 의도적으로 틀리게 만들어 불일치를 만든다.
데이터오류 행은 recalc_asset을 호출하지 않고 값만 채운다(어차피 검증단계에서 걸러짐).
"""
import datetime as dt
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pandas as pd
import recalc as R

REF_DATE = dt.date(2025, 12, 31)
FY_YEAR = 2025

COLS = ["자산명", "자산분류(유형자산/무형자산)", "취득일", "취득원가", "잔존가치", "내용연수(년)",
        "상각방법(정액법/정률법)", "회사반영_당기감가상각비", "처분일", "내용연수재추정일",
        "재추정후내용연수(년)", "재추정후상각방법(정액법/정률법)", "총예정생산량", "당기실제생산량",
        "전기말누적생산량", "자본적지출일", "자본적지출액", "상각중단시작일", "상각중단종료일",
        "회사반영_전기말감가상각누계액"]

rows = []
mismatch_names = set()
accum_mismatch_names = set()


def add_normal(name, category, cost, salvage, life, method, acq,
                disposal=None, reest_date=None, reest_life=None, reest_method=None,
                total_units=None, period_units=None, prior_period_units=None,
                capex_date=None, capex_amount=None, susp_start=None, susp_end=None,
                mismatch_offset=None, accum_mismatch_offset=None):
    dep, months, life_ended, note = R.recalc_asset(
        acq, cost, salvage, life, method, disposal, reest_date, reest_life, REF_DATE, FY_YEAR,
        total_units=total_units, period_units=period_units,
        reest_method=reest_method, capex_date=capex_date, capex_amount=capex_amount,
        susp_start=susp_start, susp_end=susp_end,
    )
    reported = dep if mismatch_offset is None else dep + mismatch_offset
    if mismatch_offset is not None:
        mismatch_names.add(name)

    accum_recalc = R.recalc_accumulated_dep(
        acq, cost, salvage, life, method, disposal, reest_date, reest_life, FY_YEAR,
        total_units=total_units, period_units=period_units,
        reest_method=reest_method, capex_date=capex_date, capex_amount=capex_amount,
        susp_start=susp_start, susp_end=susp_end,
    )
    if accum_recalc is None:
        accum_reported = None  # 생산량비례법: 재계산 자체가 안 되므로 비교 대상에서 제외
    else:
        accum_reported = accum_recalc if accum_mismatch_offset is None else accum_recalc + accum_mismatch_offset
        if accum_mismatch_offset is not None:
            accum_mismatch_names.add(name)

    rows.append([
        name, category, acq, cost, salvage, life, method, reported,
        disposal, reest_date, reest_life, reest_method, total_units, period_units, prior_period_units,
        capex_date, capex_amount, susp_start, susp_end, accum_reported,
    ])


def add_error(name, category, cost, salvage, life, method, acq, reported,
              disposal=None, reest_date=None, reest_life=None, reest_method=None,
              total_units=None, period_units=None, prior_period_units=None,
              capex_date=None, capex_amount=None, susp_start=None, susp_end=None):
    rows.append([
        name, category, acq, cost, salvage, life, method, reported,
        disposal, reest_date, reest_life, reest_method, total_units, period_units, prior_period_units,
        capex_date, capex_amount, susp_start, susp_end, None,
    ])


# ---------------------------------------------------------------------------
# A. 기본자산(이벤트 없음) - 정액법: 유형 8 + 무형 4 = 12
# ---------------------------------------------------------------------------
acq_variety = [dt.date(2020, 1, 1), dt.date(2022, 6, 30), dt.date(2023, 7, 1),
               dt.date(2024, 12, 1), dt.date(2021, 3, 15), dt.date(2019, 1, 1),
               dt.date(2025, 1, 1), dt.date(2018, 5, 1)]
for i, acq in enumerate(acq_variety, start=1):
    add_normal(f"정액법기본자산_유형{i:02d}", "유형자산", 10_000_000 + i * 500_000, 0,
               [5, 6, 8, 10, 4, 7, 5, 12][i - 1], "정액법", acq)
for i, acq in enumerate(acq_variety[:4], start=1):
    add_normal(f"정액법기본자산_무형{i:02d}", "무형자산", 8_000_000 + i * 300_000, 0,
               [5, 10, 4, 7][i - 1], "정액법", acq)

# ---------------------------------------------------------------------------
# B. 기본자산(이벤트 없음) - 정률법: 유형 8 + 무형 4 = 12 (내용연수 표범위 + 근사식 케이스)
# ---------------------------------------------------------------------------
db_lives = [5, 6, 8, 10, 15, 3, 20, 25]  # 마지막(25)은 상각률표 밖 -> 근사식
for i, (acq, life) in enumerate(zip(acq_variety, db_lives), start=1):
    add_normal(f"정률법기본자산_유형{i:02d}", "유형자산", 12_000_000 + i * 400_000, 0,
               life, "정률법", acq)
for i, (acq, life) in enumerate(zip(acq_variety[:4], [5, 9, 12, 22]), start=1):
    add_normal(f"정률법기본자산_무형{i:02d}", "무형자산", 9_000_000 + i * 350_000, 0,
               life, "정률법", acq)

# ---------------------------------------------------------------------------
# C. 기본자산(이벤트 없음) - 생산량비례법: 유형 4 + 무형 2 = 6
# ---------------------------------------------------------------------------
for i in range(1, 5):
    # 유형01~03은 전기말누적생산량을 입력해 전기말/당기말 누계액 수식화를 시연하고,
    # 유형04는 미입력 상태(기존처럼 값 없이 공란)로 남겨 하위호환도 함께 검증한다.
    add_normal(f"생산량비례법기본자산_유형{i:02d}", "유형자산", 20_000_000, 2_000_000, 1,
               "생산량비례법", dt.date(2022, 1, 1),
               total_units=100_000, period_units=8_000 + i * 1_000,
               prior_period_units=(40_000 + i * 2_000) if i <= 3 else None)
for i in range(1, 3):
    add_normal(f"생산량비례법기본자산_무형{i:02d}", "무형자산", 15_000_000, 0, 1,
               "생산량비례법", dt.date(2023, 1, 1),
               total_units=50_000, period_units=6_000 + i * 500,
               prior_period_units=20_000 + i * 1_000)

# ---------------------------------------------------------------------------
# D. 처분만: 정액/정률 x 당기중/전기이전 x 유형/무형 대표 10건
# ---------------------------------------------------------------------------
add_normal("처분_정액_당기중_유형01", "유형자산", 10_000_000, 0, 5, "정액법",
           dt.date(2021, 1, 1), disposal=dt.date(2025, 6, 30))
add_normal("처분_정액_당기중_무형01", "무형자산", 8_000_000, 0, 4, "정액법",
           dt.date(2022, 1, 1), disposal=dt.date(2025, 9, 30))
add_normal("처분_정액_전기이전_유형01", "유형자산", 9_000_000, 0, 5, "정액법",
           dt.date(2019, 1, 1), disposal=dt.date(2024, 12, 31))
add_normal("처분_정률_당기중_유형01", "유형자산", 15_000_000, 0, 8, "정률법",
           dt.date(2021, 1, 1), disposal=dt.date(2025, 4, 30))
add_normal("처분_정률_당기중_무형01", "무형자산", 11_000_000, 0, 6, "정률법",
           dt.date(2022, 1, 1), disposal=dt.date(2025, 11, 30))
add_normal("처분_정률_전기이전_유형01", "유형자산", 13_000_000, 0, 5, "정률법",
           dt.date(2018, 1, 1), disposal=dt.date(2023, 12, 31))
add_normal("처분_정액_당기중_유형02", "유형자산", 7_000_000, 0, 5, "정액법",
           dt.date(2023, 6, 30), disposal=dt.date(2025, 1, 31), mismatch_offset=300_000)
add_normal("처분_정률_당기중_유형02", "유형자산", 16_000_000, 0, 10, "정률법",
           dt.date(2020, 1, 1), disposal=dt.date(2025, 8, 15))
add_normal("처분_정액_전기이전_무형01", "무형자산", 6_000_000, 0, 4, "정액법",
           dt.date(2018, 1, 1), disposal=dt.date(2022, 12, 31))
add_normal("처분_정률_전기이전_무형01", "무형자산", 10_000_000, 0, 6, "정률법",
           dt.date(2017, 1, 1), disposal=dt.date(2023, 6, 30))

# ---------------------------------------------------------------------------
# E. 재추정만(방법 유지): 정액/정률 x 유형/무형 = 6
# ---------------------------------------------------------------------------
add_normal("재추정_정액_유형01", "유형자산", 12_000_000, 0, 5, "정액법",
           dt.date(2021, 1, 1), reest_date=dt.date(2024, 1, 1), reest_life=8,
           accum_mismatch_offset=800_000)
add_normal("재추정_정액_무형01", "무형자산", 9_000_000, 0, 4, "정액법",
           dt.date(2022, 1, 1), reest_date=dt.date(2024, 7, 1), reest_life=6)
add_normal("재추정_정률_유형01", "유형자산", 14_000_000, 0, 6, "정률법",
           dt.date(2021, 1, 1), reest_date=dt.date(2024, 1, 1), reest_life=10,
           accum_mismatch_offset=350_000)
add_normal("재추정_정률_무형01", "무형자산", 10_000_000, 0, 5, "정률법",
           dt.date(2022, 1, 1), reest_date=dt.date(2024, 1, 1), reest_life=8)
add_normal("재추정_정액_유형02", "유형자산", 11_000_000, 0, 5, "정액법",
           dt.date(2020, 1, 1), reest_date=dt.date(2023, 1, 1), reest_life=7, mismatch_offset=-200_000)
add_normal("재추정_정률_유형02", "유형자산", 13_000_000, 0, 5, "정률법",
           dt.date(2020, 1, 1), reest_date=dt.date(2023, 1, 1), reest_life=9)

# ---------------------------------------------------------------------------
# F. 재추정+방법변경: 정액→정률 2, 정률→정액 2 = 4
# ---------------------------------------------------------------------------
add_normal("방법변경_정액to정률_01", "유형자산", 12_000_000, 0, 10, "정액법",
           dt.date(2021, 1, 1), reest_date=dt.date(2024, 1, 1), reest_life=6, reest_method="정률법")
add_normal("방법변경_정액to정률_02", "무형자산", 9_000_000, 0, 8, "정액법",
           dt.date(2020, 1, 1), reest_date=dt.date(2025, 4, 1), reest_life=5, reest_method="정률법")
add_normal("방법변경_정률to정액_01", "유형자산", 15_000_000, 0, 10, "정률법",
           dt.date(2020, 1, 1), reest_date=dt.date(2024, 1, 1), reest_life=6, reest_method="정액법")
add_normal("방법변경_정률to정액_02", "유형자산", 10_000_000, 0, 8, "정률법",
           dt.date(2021, 1, 1), reest_date=dt.date(2025, 7, 1), reest_life=4, reest_method="정액법",
           mismatch_offset=400_000)

# ---------------------------------------------------------------------------
# G. 자본적지출만: 정액/정률 x 지출시점(당기초/중/말/전기) = 8
# ---------------------------------------------------------------------------
add_normal("자본적지출_정액_전기_01", "유형자산", 10_000_000, 0, 10, "정액법",
           dt.date(2020, 1, 1), capex_date=dt.date(2023, 1, 1), capex_amount=2_000_000,
           accum_mismatch_offset=-450_000)
add_normal("자본적지출_정액_당기초_01", "유형자산", 8_000_000, 0, 8, "정액법",
           dt.date(2019, 1, 1), capex_date=dt.date(2025, 1, 1), capex_amount=1_500_000)
add_normal("자본적지출_정액_당기중_01", "무형자산", 9_000_000, 0, 6, "정액법",
           dt.date(2021, 1, 1), capex_date=dt.date(2025, 6, 1), capex_amount=1_000_000)
add_normal("자본적지출_정액_당기말_01", "유형자산", 7_000_000, 0, 5, "정액법",
           dt.date(2022, 1, 1), capex_date=dt.date(2025, 12, 1), capex_amount=800_000)
add_normal("자본적지출_정률_전기_01", "유형자산", 14_000_000, 0, 10, "정률법",
           dt.date(2020, 1, 1), capex_date=dt.date(2023, 1, 1), capex_amount=2_500_000)
add_normal("자본적지출_정률_당기초_01", "유형자산", 11_000_000, 0, 8, "정률법",
           dt.date(2019, 1, 1), capex_date=dt.date(2025, 1, 1), capex_amount=1_800_000)
add_normal("자본적지출_정률_당기중_01", "무형자산", 12_000_000, 0, 9, "정률법",
           dt.date(2021, 1, 1), capex_date=dt.date(2025, 7, 1), capex_amount=1_200_000,
           mismatch_offset=250_000)
add_normal("자본적지출_정률_당기말_01", "유형자산", 9_500_000, 0, 6, "정률법",
           dt.date(2022, 1, 1), capex_date=dt.date(2025, 11, 1), capex_amount=900_000)

# ---------------------------------------------------------------------------
# H. 상각중단만: (당기전체덮음/당기일부만걸침/이미종료돼재개됨) x 정액/정률 = 8
# ---------------------------------------------------------------------------
add_normal("상각중단_정액_당기전체_01", "유형자산", 6_000_000, 0, 5, "정액법",
           dt.date(2020, 1, 1), susp_start=dt.date(2025, 1, 1), susp_end=dt.date(2025, 12, 31),
           accum_mismatch_offset=600_000)
add_normal("상각중단_정률_당기전체_01", "유형자산", 10_000_000, 0, 6, "정률법",
           dt.date(2020, 1, 1), susp_start=dt.date(2025, 1, 1), susp_end=dt.date(2025, 12, 31))
add_normal("상각중단_정액_당기일부_01", "무형자산", 6_000_000, 0, 5, "정액법",
           dt.date(2020, 1, 1), susp_start=dt.date(2022, 6, 1), susp_end=dt.date(2022, 12, 31))
add_normal("상각중단_정률_당기일부_01", "유형자산", 9_000_000, 0, 5, "정률법",
           dt.date(2021, 1, 1), susp_start=dt.date(2023, 3, 1), susp_end=dt.date(2023, 9, 30))
add_normal("상각중단_정액_재개후_01", "유형자산", 6_000_000, 0, 5, "정액법",
           dt.date(2020, 1, 1), susp_start=dt.date(2021, 1, 1), susp_end=dt.date(2021, 6, 30))
add_normal("상각중단_정률_재개후_01", "유형자산", 8_000_000, 0, 5, "정률법",
           dt.date(2021, 1, 1), susp_start=dt.date(2022, 1, 1), susp_end=dt.date(2022, 6, 30))
add_normal("상각중단_정액_당기일부_02", "유형자산", 7_500_000, 0, 5, "정액법",
           dt.date(2022, 1, 1), susp_start=dt.date(2024, 4, 1), susp_end=dt.date(2024, 10, 31),
           mismatch_offset=150_000)
add_normal("상각중단_정률_재개후_02", "무형자산", 5_500_000, 0, 4, "정률법",
           dt.date(2021, 6, 1), susp_start=dt.date(2022, 1, 1), susp_end=dt.date(2022, 3, 31))

# ---------------------------------------------------------------------------
# I. 추가취득(=기중 신규취득, 코드검증용): 6/30,7/1,1/1,12/1 x 정액/정률 = 8
# ---------------------------------------------------------------------------
new_acq_dates = [dt.date(2025, 6, 30), dt.date(2025, 7, 1), dt.date(2025, 1, 1), dt.date(2025, 12, 1)]
for i, acq in enumerate(new_acq_dates, start=1):
    add_normal(f"추가취득_정액_{i:02d}", "유형자산", 5_000_000, 0, 5, "정액법", acq)
for i, acq in enumerate(new_acq_dates, start=1):
    add_normal(f"추가취득_정률_{i:02d}", "유형자산", 6_000_000, 0, 6, "정률법", acq)

# ---------------------------------------------------------------------------
# J. 2개 이벤트 조합: capex+reest(순서다른2), capex+중단, reest+중단,
#    처분+capex, 처분+중단, 처분+reest = 12
# ---------------------------------------------------------------------------
add_normal("조합_capex먼저reest나중", "유형자산", 10_000_000, 0, 8, "정액법",
           dt.date(2021, 1, 1), capex_date=dt.date(2023, 1, 1), capex_amount=1_000_000,
           reest_date=dt.date(2024, 6, 1), reest_life=6)
add_normal("조합_reest먼저capex나중", "유형자산", 10_000_000, 0, 8, "정액법",
           dt.date(2021, 1, 1), reest_date=dt.date(2023, 1, 1), reest_life=6,
           capex_date=dt.date(2024, 6, 1), capex_amount=1_000_000)
add_normal("조합_capex중단_01", "유형자산", 9_000_000, 0, 6, "정액법",
           dt.date(2021, 1, 1), capex_date=dt.date(2023, 1, 1), capex_amount=1_500_000,
           susp_start=dt.date(2024, 3, 1), susp_end=dt.date(2024, 8, 31))
add_normal("조합_capex중단_02", "무형자산", 8_500_000, 0, 6, "정률법",
           dt.date(2021, 1, 1), capex_date=dt.date(2023, 6, 1), capex_amount=1_200_000,
           susp_start=dt.date(2024, 1, 1), susp_end=dt.date(2024, 4, 30))
add_normal("조합_reest중단_01", "유형자산", 11_000_000, 0, 7, "정액법",
           dt.date(2021, 1, 1), reest_date=dt.date(2023, 1, 1), reest_life=6,
           susp_start=dt.date(2024, 3, 1), susp_end=dt.date(2024, 9, 30))
add_normal("조합_reest중단_02", "유형자산", 12_000_000, 0, 7, "정률법",
           dt.date(2020, 1, 1), reest_date=dt.date(2023, 1, 1), reest_life=6,
           susp_start=dt.date(2024, 1, 1), susp_end=dt.date(2024, 5, 31))
add_normal("조합_처분capex_01", "유형자산", 8_000_000, 0, 5, "정액법",
           dt.date(2022, 1, 1), disposal=dt.date(2025, 9, 30),
           capex_date=dt.date(2024, 1, 1), capex_amount=1_000_000)
add_normal("조합_처분capex_02", "유형자산", 9_000_000, 0, 6, "정률법",
           dt.date(2021, 1, 1), disposal=dt.date(2025, 6, 30),
           capex_date=dt.date(2023, 1, 1), capex_amount=1_100_000)
add_normal("조합_처분중단_01", "무형자산", 7_000_000, 0, 5, "정액법",
           dt.date(2021, 1, 1), disposal=dt.date(2025, 10, 31),
           susp_start=dt.date(2023, 1, 1), susp_end=dt.date(2023, 6, 30))
add_normal("조합_처분중단_02", "유형자산", 10_500_000, 0, 6, "정률법",
           dt.date(2020, 1, 1), disposal=dt.date(2025, 5, 31),
           susp_start=dt.date(2022, 1, 1), susp_end=dt.date(2022, 8, 31))
add_normal("조합_처분reest_01", "유형자산", 9_500_000, 0, 6, "정액법",
           dt.date(2021, 1, 1), disposal=dt.date(2025, 8, 31),
           reest_date=dt.date(2023, 1, 1), reest_life=5)
add_normal("조합_처분reest_02", "유형자산", 11_500_000, 0, 7, "정률법",
           dt.date(2020, 1, 1), disposal=dt.date(2025, 3, 31),
           reest_date=dt.date(2022, 1, 1), reest_life=5, mismatch_offset=200_000)

# ---------------------------------------------------------------------------
# K. 3개 이벤트 조합: capex+reest+중단, capex+reest+처분, capex+중단+처분,
#    reest+중단+처분 = 8
# ---------------------------------------------------------------------------
add_normal("3조합_capex_reest_중단_01", "유형자산", 12_000_000, 0, 8, "정액법",
           dt.date(2020, 1, 1), capex_date=dt.date(2022, 1, 1), capex_amount=1_000_000,
           reest_date=dt.date(2024, 1, 1), reest_life=6,
           susp_start=dt.date(2025, 2, 1), susp_end=dt.date(2025, 7, 31))
add_normal("3조합_capex_reest_중단_02", "무형자산", 13_000_000, 0, 8, "정률법",
           dt.date(2020, 1, 1), capex_date=dt.date(2022, 6, 1), capex_amount=1_100_000,
           reest_date=dt.date(2024, 1, 1), reest_life=6,
           susp_start=dt.date(2025, 1, 1), susp_end=dt.date(2025, 6, 30))
add_normal("3조합_capex_reest_처분_01", "유형자산", 10_000_000, 0, 8, "정액법",
           dt.date(2020, 1, 1), capex_date=dt.date(2022, 1, 1), capex_amount=1_000_000,
           reest_date=dt.date(2023, 1, 1), reest_life=6, disposal=dt.date(2025, 9, 30),
           accum_mismatch_offset=-700_000)
add_normal("3조합_capex_reest_처분_02", "유형자산", 14_000_000, 0, 9, "정률법",
           dt.date(2019, 1, 1), capex_date=dt.date(2021, 1, 1), capex_amount=1_500_000,
           reest_date=dt.date(2023, 1, 1), reest_life=7, disposal=dt.date(2025, 6, 30))
add_normal("3조합_capex_중단_처분_01", "유형자산", 9_000_000, 0, 6, "정액법",
           dt.date(2021, 1, 1), capex_date=dt.date(2023, 1, 1), capex_amount=800_000,
           susp_start=dt.date(2024, 1, 1), susp_end=dt.date(2024, 6, 30), disposal=dt.date(2025, 10, 31))
add_normal("3조합_capex_중단_처분_02", "무형자산", 8_800_000, 0, 6, "정률법",
           dt.date(2021, 1, 1), capex_date=dt.date(2023, 6, 1), capex_amount=900_000,
           susp_start=dt.date(2024, 1, 1), susp_end=dt.date(2024, 5, 31), disposal=dt.date(2025, 7, 31))
add_normal("3조합_reest_중단_처분_01", "유형자산", 11_000_000, 0, 7, "정액법",
           dt.date(2020, 1, 1), reest_date=dt.date(2022, 1, 1), reest_life=6,
           susp_start=dt.date(2023, 1, 1), susp_end=dt.date(2023, 6, 30), disposal=dt.date(2025, 11, 30))
add_normal("3조합_reest_중단_처분_02", "유형자산", 12_500_000, 0, 7, "정률법",
           dt.date(2019, 1, 1), reest_date=dt.date(2022, 1, 1), reest_life=6,
           susp_start=dt.date(2023, 1, 1), susp_end=dt.date(2023, 4, 30), disposal=dt.date(2025, 12, 31))

# ---------------------------------------------------------------------------
# L. 경계 케이스 6건
# ---------------------------------------------------------------------------
add_normal("경계_capex reest동일날짜", "유형자산", 10_000_000, 0, 5, "정액법",
           dt.date(2022, 1, 1), capex_date=dt.date(2024, 1, 1), capex_amount=2_000_000,
           reest_date=dt.date(2024, 1, 1), reest_life=4)
add_normal("경계_reest가중단구간안", "유형자산", 9_000_000, 0, 6, "정액법",
           dt.date(2021, 1, 1), susp_start=dt.date(2023, 3, 1), susp_end=dt.date(2023, 9, 30),
           reest_date=dt.date(2023, 6, 1), reest_life=5)
add_normal("경계_처분이중단구간안", "유형자산", 8_000_000, 0, 5, "정액법",
           dt.date(2022, 1, 1), susp_start=dt.date(2025, 3, 1), susp_end=dt.date(2025, 9, 30),
           disposal=dt.date(2025, 6, 30))
add_normal("경계_capex직후방법변경", "유형자산", 10_000_000, 0, 10, "정액법",
           dt.date(2022, 1, 1), capex_date=dt.date(2024, 1, 1), capex_amount=1_000_000,
           reest_date=dt.date(2025, 7, 1), reest_life=5, reest_method="정률법")
add_normal("경계_중단이중간구간", "유형자산", 6_000_000, 0, 5, "정액법",
           dt.date(2020, 1, 1), susp_start=dt.date(2021, 6, 1), susp_end=dt.date(2021, 12, 31),
           reest_date=dt.date(2022, 1, 1), reest_life=3)
add_normal("경계_중단후재개후처분", "유형자산", 12_000_000, 0, 8, "정액법",
           dt.date(2020, 1, 1), susp_start=dt.date(2022, 1, 1), susp_end=dt.date(2022, 6, 30),
           disposal=dt.date(2025, 9, 30))

# ---------------------------------------------------------------------------
# M. 데이터오류 9건 (recalc_asset 호출 없이 직접 값 기재)
# ---------------------------------------------------------------------------
add_error("오류_내용연수0", "유형자산", 10_000_000, 0, 0, "정액법", dt.date(2022, 1, 1), 1_000_000)
add_error("오류_내용연수음수", "유형자산", 10_000_000, 0, -5, "정액법", dt.date(2022, 1, 1), 1_000_000)
add_error("오류_취득원가0이하", "유형자산", 0, 0, 5, "정액법", dt.date(2022, 1, 1), 500_000)
add_error("오류_잔존가치초과", "유형자산", 10_000_000, 15_000_000, 5, "정액법", dt.date(2022, 1, 1), 500_000)
add_error("오류_생산량비례법생산량누락", "유형자산", 10_000_000, 0, 5, "생산량비례법",
          dt.date(2022, 1, 1), 500_000, total_units=None, period_units=5_000)
add_error("오류_자본적지출날짜금액불일치", "유형자산", 10_000_000, 0, 5, "정액법",
          dt.date(2022, 1, 1), 500_000, capex_date=dt.date(2024, 1, 1), capex_amount=None)
add_error("오류_상각중단시작종료역전", "유형자산", 10_000_000, 0, 5, "정액법",
          dt.date(2022, 1, 1), 500_000, susp_start=dt.date(2024, 6, 1), susp_end=dt.date(2024, 1, 1))
add_error("오류_재추정후상각방법오타", "유형자산", 10_000_000, 0, 5, "정액법",
          dt.date(2022, 1, 1), 500_000, reest_date=dt.date(2024, 1, 1), reest_life=4, reest_method="정율법")
add_error("오류_상각중단짝안맞음", "유형자산", 10_000_000, 0, 5, "정액법",
          dt.date(2022, 1, 1), 500_000, susp_start=dt.date(2024, 1, 1), susp_end=None)

# ---------------------------------------------------------------------------
# 저장
# ---------------------------------------------------------------------------
out_df = pd.DataFrame(rows, columns=COLS)
out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sample_asset_ledger.xlsx")
out_df.to_excel(out_path, index=False)
print(f"생성 완료: {len(out_df)}건 -> {out_path}")
print(f"의도적 당기상각비 불일치 자산({len(mismatch_names)}건): {sorted(mismatch_names)}")
print(f"의도적 누계액 불일치 자산({len(accum_mismatch_names)}건): {sorted(accum_mismatch_names)}")
