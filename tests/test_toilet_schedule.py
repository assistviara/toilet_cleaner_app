import sys
from pathlib import Path

import pandas as pd

# 親フォルダを import 対象に追加
sys.path.append(str(Path(__file__).resolve().parents[1]))

from toilet_logic import (
    is_eligible_for_cleaning,
    parse_skip_words,
    build_schedule_from_row_staff_table,
)


def test_is_eligible_for_cleaning_x_is_off():
    assert is_eligible_for_cleaning("×") is False
    assert is_eligible_for_cleaning("x") is False
    assert is_eligible_for_cleaning("X") is False
    assert is_eligible_for_cleaning("ｘ") is False


def test_is_eligible_for_cleaning_blank_is_working():
    assert is_eligible_for_cleaning("") is True
    assert is_eligible_for_cleaning(" ") is True


def test_is_eligible_for_cleaning_time_text_is_working():
    assert is_eligible_for_cleaning("7:30-16:30") is True
    assert is_eligible_for_cleaning("9-18") is True


def test_parse_skip_words():
    result = parse_skip_words("研修, 出張, 会議")
    assert result == {"研修", "出張", "会議"}


def test_is_eligible_for_cleaning_skip_word_exact_or_contained():
    assert is_eligible_for_cleaning("研修", skip_duty_words={"研修"}) is False
    assert is_eligible_for_cleaning("研修 9:00-17:00", skip_duty_words={"研修"}) is False
    assert is_eligible_for_cleaning("出張(福岡)", skip_duty_words={"出張"}) is False
    assert is_eligible_for_cleaning("通常勤務", skip_duty_words={"研修", "出張"}) is True


def test_build_schedule_basic_rotation():
    df = pd.DataFrame({
        "氏名": ["A", "B", "C", "D"],
        "男性チェック": ["△", "△", "", ""],
        "社員チェック": ["", "", "", "x"],   # D は社員除外
        "1": ["", "", "", ""],
        "2": ["×", "", "", ""],              # A は休み
        "3": ["", "", "研修", ""],           # C は研修で除外
        "4": ["", "×", "", ""],              # B は休み
    })

    result_df, summary_df = build_schedule_from_row_staff_table(
        df=df,
        name_col="氏名",
        male_col="男性チェック",
        employee_col="社員チェック",
        day_cols=["1", "2", "3", "4"],
        skip_duty_words={"研修"},
    )

    expected_male = ["A", "B", "A", "A"]
    expected_female = ["C", "C", None, "C"]

    assert result_df["男性便所担当"].tolist() == expected_male

    actual_female = result_df["女性便所担当"].tolist()
    for actual, expected in zip(actual_female, expected_female):
        if expected is None:
            assert pd.isna(actual)
        else:
            assert actual == expected

    d_row = summary_df[summary_df["職員名"] == "D"].iloc[0]
    assert d_row["社員除外"] == "はい"
    assert d_row["担当回数"] == 0

def test_skip_word_changes_assignment():
    df = pd.DataFrame({
        "氏名": ["井上", "大久保", "竹畠"],
        "男性チェック": ["△", "△", ""],
        "社員チェック": ["", "", ""],
        "1": ["", "", ""],
        "2": ["", "", ""],
        "3": ["出張", "", ""],   # 井上を除外して差が出るようにする
    })

    result_df, _ = build_schedule_from_row_staff_table(
        df=df,
        name_col="氏名",
        male_col="男性チェック",
        employee_col="社員チェック",
        day_cols=["1", "2", "3"],
        skip_duty_words={"出張"},
    )

    # 1日: 井上
    # 2日: 大久保
    # 3日: 本来は井上に戻りうるが、出張なので除外され大久保
    assert result_df["男性便所担当"].tolist() == ["井上", "大久保", "大久保"]


    def test_fairness_counts_not_extreme_when_all_work_evenly():
        df = pd.DataFrame({
            "氏名": ["A", "B", "C", "D"],
            "男性チェック": ["△", "△", "", ""],
            "社員チェック": ["", "", "", ""],
            "1": ["", "", "", ""],
            "2": ["", "", "", ""],
            "3": ["", "", "", ""],
            "4": ["", "", "", ""],
            "5": ["", "", "", ""],
            "6": ["", "", "", ""],
        })

        result_df, summary_df = build_schedule_from_row_staff_table(
            df=df,
            name_col="氏名",
            male_col="男性チェック",
            employee_col="社員チェック",
            day_cols=["1", "2", "3", "4", "5", "6"],
            skip_duty_words=set(),
        )

        # 男性は A, B の2人で6日を回す → 3回ずつ
        male_summary = summary_df[
            (summary_df["性別区分"] == "男性") & (summary_df["社員除外"] == "いいえ")
        ].sort_values("職員名")

        male_counts = male_summary["担当回数"].tolist()
        assert male_counts == [3, 3]

        # 女性は C, D の2人で6日を回す → 3回ずつ
        female_summary = summary_df[
            (summary_df["性別区分"] == "女性") & (summary_df["社員除外"] == "いいえ")
        ].sort_values("職員名")

        female_counts = female_summary["担当回数"].tolist()
        assert female_counts == [3, 3]

def test_fairness_by_ratio_protects_low_attendance_staff():
    df = pd.DataFrame({
        "氏名": ["A", "B", "C"],
        "男性チェック": ["△", "△", "△"],
        "社員チェック": ["", "", ""],
        "1": ["", "", ""],
        "2": ["", "", ""],
        "3": ["", "", "×"],
        "4": ["", "", "×"],
        "5": ["", "", "×"],
        "6": ["", "", "×"],
    })

    result_df, summary_df = build_schedule_from_row_staff_table(
        df=df,
        name_col="氏名",
        male_col="男性チェック",
        employee_col="社員チェック",
        day_cols=["1", "2", "3", "4", "5", "6"],
        skip_duty_words=set(),
    )

    # C は出勤可能日数が 2 日しかないので、
    # 回数だけで均等化しすぎないことを確認する
    c_row = summary_df[summary_df["職員名"] == "C"].iloc[0]
    a_row = summary_df[summary_df["職員名"] == "A"].iloc[0]
    b_row = summary_df[summary_df["職員名"] == "B"].iloc[0]

    assert c_row["出勤可能日数"] == 2
    assert a_row["出勤可能日数"] == 6
    assert b_row["出勤可能日数"] == 6

    # C の担当率が、A/B より極端に高くなっていないこと
    assert c_row["担当率"] <= max(a_row["担当率"], b_row["担当率"]) + 0.01


def test_blank_name_rows_are_ignored():
    df = pd.DataFrame({
        "氏名": ["A", "", "B"],
        "男性チェック": ["△", "", ""],
        "社員チェック": ["", "", ""],
        "1": ["", "", "×"],   # B に勤務情報を入れる
        "2": ["", "", ""],
        "3": ["", "", ""],
    })

    result_df, summary_df = build_schedule_from_row_staff_table(
        df=df,
        name_col="氏名",
        male_col="男性チェック",
        employee_col="社員チェック",
        day_cols=["1", "2", "3"],
        skip_duty_words=set(),
    )

    # 名前空欄の行が集計対象に入らない
    assert "" not in summary_df["職員名"].fillna("").tolist()

    # 男性担当に空欄名が出ない
    male_names = [x for x in result_df["男性便所担当"].tolist() if x is not None]
    assert "" not in male_names

    # 女性担当に空欄名が出ない
    female_names = [x for x in result_df["女性便所担当"].tolist() if x is not None]
    assert "" not in female_names

def test_clean_staff_rows_ignores_blank_and_formula_like_rows():
    df = pd.DataFrame({
        "氏名": ["A", "", "", "B"],
        "男性チェック": ["△", "", "", ""],
        "社員チェック": ["", "", "", ""],
        "1": ["", "", "", "×"],   # B に勤務表らしい情報を入れる
        "2": ["×", "", "", ""],
        "3": ["", "", "", ""],
    })

    result_df, summary_df = build_schedule_from_row_staff_table(
        df=df,
        name_col="氏名",
        male_col="男性チェック",
        employee_col="社員チェック",
        day_cols=["1", "2", "3"],
        skip_duty_words=set(),
    )

    names = summary_df["職員名"].fillna("").tolist()

    assert "A" in names
    assert "B" in names
    assert "" not in names

def test_formula_like_rows_are_ignored():
    df = pd.DataFrame({
        "氏名": ["A", "", "平均"],
        "男性チェック": ["△", "", ""],
        "社員チェック": ["", "", ""],
        "1": ["", "", ""],
        "2": ["×", "", ""],
        "3": ["", "", ""],
    })

    result_df, summary_df = build_schedule_from_row_staff_table(
        df=df,
        name_col="氏名",
        male_col="男性チェック",
        employee_col="社員チェック",
        day_cols=["1", "2", "3"],
        skip_duty_words=set(),
    )

    names = summary_df["職員名"].fillna("").tolist()
    assert "A" in names
    assert "平均" not in names