import sys
from pathlib import Path

import pandas as pd

# 親フォルダを import 対象に追加
sys.path.append(str(Path(__file__).resolve().parents[1]))

from toilet_logic import (
    is_eligible_for_cleaning,
    parse_skip_words,
    build_schedule_from_row_staff_table,
    rotate_queue_from_last,
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


def test_rotate_queue_from_last():
    queue = ["A", "B", "C", "D"]
    rotated = rotate_queue_from_last(queue, "B")
    assert rotated == ["C", "D", "A", "B"]


def test_rotate_queue_from_last_when_name_not_found():
    queue = ["A", "B", "C"]
    rotated = rotate_queue_from_last(queue, "Z")
    assert rotated == ["A", "B", "C"]


def test_build_schedule_basic_rotation():
    df = pd.DataFrame({
        "氏名": ["A", "B", "C", "D"],
        "男性チェック": ["△", "△", "", ""],
        "社員チェック": ["", "", "", "〇"],   # D は社員除外
        "1": ["", "", "", "×"],
        "2": ["×", "", "×", ""],
        "3": ["", "", "研修", ""],
        "4": ["", "×", "", ""],
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
    expected_female = ["C", None, None, "C"]

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
        "3": ["出張", "", ""],   # 井上はこの日除外
        "4": ["", "", ""],
        "5": ["×", "×", "×"],   # 全員 clean_staff_rows を通すため
    })

    result_df, _ = build_schedule_from_row_staff_table(
        df=df,
        name_col="氏名",
        male_col="男性チェック",
        employee_col="社員チェック",
        day_cols=["1", "2", "3", "4", "5"],
        skip_duty_words={"出張"},
    )

    # 1日: 井上
    # 2日: 大久保
    # 3日: 井上は出張なので除外され、大久保
    # 4日: この時点では負担率の都合で井上に戻る
    assert result_df["男性便所担当"].tolist()[:4] == ["井上", "大久保", "大久保", "井上"]




def test_blank_name_rows_are_ignored():
    df = pd.DataFrame({
        "氏名": ["A", "", "B"],
        "男性チェック": ["△", "", ""],
        "社員チェック": ["", "", ""],
        "1": ["", "", "×"],
        "2": ["×", "", ""],   # A にも × を入れる
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

    assert "" not in summary_df["職員名"].fillna("").tolist()

    male_names = [x for x in result_df["男性便所担当"].tolist() if pd.notna(x)]
    assert "" not in male_names

    female_names = [x for x in result_df["女性便所担当"].tolist() if pd.notna(x)]
    assert "" not in female_names


def test_clean_staff_rows_ignores_blank_and_formula_like_rows():
    df = pd.DataFrame({
        "氏名": ["A", "", "=COUNTA(B:B)", "B"],
        "男性チェック": ["△", "", "", ""],
        "社員チェック": ["", "", "", ""],
        "1": ["", "", "", "×"],
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
    assert "=COUNTA(B:B)" not in names


def test_formula_like_rows_are_ignored():
    df = pd.DataFrame({
        "氏名": ["A", "=COUNTA(B:B)", "18"],
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
    assert "=COUNTA(B:B)" not in names
    assert "18" not in names


def test_previous_last_person_changes_start_position():
    df = pd.DataFrame({
        "氏名": ["A", "B", "C", "D"],
        "男性チェック": ["△", "△", "", ""],
        "社員チェック": ["", "", "", ""],
        "1": ["", "", "", ""],
        "2": ["", "", "", ""],
        "3": ["", "", "", ""],
        "4": ["×", "×", "×", "×"],   # 全員 clean を通すために × を持たせる
    })

    result_df, summary_df = build_schedule_from_row_staff_table(
        df=df,
        name_col="氏名",
        male_col="男性チェック",
        employee_col="社員チェック",
        day_cols=["1", "2", "3", "4"],
        skip_duty_words=set(),
        previous_male_last="A",
        previous_female_last="C",
    )

    # 男性は A の次なので B から開始
    assert result_df["男性便所担当"].tolist()[0] == "B"

    # 女性は C の次なので D から開始
    assert result_df["女性便所担当"].tolist()[0] == "D"

def test_keep_order_rotation_basic():
    df = pd.DataFrame({
        "氏名": ["A", "B", "C", "D"],
        "男性チェック": ["△", "△", "", ""],
        "社員チェック": ["", "", "", "〇"],   # D は社員
        "1": ["", "", "", "×"],
        "2": ["×", "", "×", ""],
        "3": ["", "", "研修", ""],
        "4": ["", "×", "", ""],
    })

    result_df, summary_df = build_schedule_from_row_staff_table(
        df=df,
        name_col="氏名",
        male_col="男性チェック",
        employee_col="社員チェック",
        day_cols=["1", "2", "3", "4"],
        skip_duty_words={"研修"},
    )

    # 男性は順番どおり
    assert result_df["男性便所担当"].tolist() == ["A", "B", "A", "A"]

    # 女性は C のみ（D は社員）
    actual_female = result_df["女性便所担当"].tolist()
    expected_female = ["C", None, None, "C"]

    for actual, expected in zip(actual_female, expected_female):
        if expected is None:
            assert pd.isna(actual)
        else:
            assert actual == expected

def test_skip_keeps_order_and_skips_absent_person():
    df = pd.DataFrame({
        "氏名": ["井上", "大久保", "竹畠"],
        "男性チェック": ["△", "△", ""],
        "社員チェック": ["", "", ""],
        "1": ["", "", ""],
        "2": ["", "", ""],
        "3": ["出張", "", ""],
        "4": ["", "", ""],
        "5": ["×", "×", "×"],
    })

    result_df, _ = build_schedule_from_row_staff_table(
        df=df,
        name_col="氏名",
        male_col="男性チェック",
        employee_col="社員チェック",
        day_cols=["1", "2", "3", "4", "5"],
        skip_duty_words={"出張"},
    )

    # 順番制:
    # 1日 井上
    # 2日 大久保
    # 3日 井上は出張なのでスキップ、大久保
    # 4日 次は井上
    assert result_df["男性便所担当"].tolist()[:4] == ["井上", "大久保", "大久保", "井上"]