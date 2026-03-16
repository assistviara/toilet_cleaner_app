import re
from collections import deque
from typing import Iterable, Optional

import pandas as pd


# =========================
# 設定値
# =========================
# 本番仕様:
# ×印だけ休み
# それ以外（空欄、時間表記など）はすべて出勤扱い
DEFAULT_OFF_WORDS = {"×", "x", "X", "ｘ"}

DEFAULT_SKIP_DUTY_WORDS = {
    # 初期値が必要ならここに書ける
    # "研修",
    # "出張",
}


# =========================
# 共通関数
# =========================
def normalize_cell(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def is_checked(value) -> bool:
    """
    男性チェック列 / 社員チェック列 用
    △ / ○ / 〇 / ✓ / 1 / True などをチェックありとして扱う
    """
    if pd.isna(value):
        return False

    if isinstance(value, bool):
        return value

    if isinstance(value, (int, float)):
        return value == 1

    s = str(value).strip()

    return s in {
        "1", "true", "True", "TRUE", "yes", "Yes", "YES", "y", "Y",
        "on", "ON", "checked", "CHECKED",
        "○", "〇", "✓", "✔", "レ", "△",
        "男", "男性", "社員",
    }


def parse_skip_words(text: str) -> set[str]:
    """
    UI入力:
    研修, 出張, 会議
    ↓
    {"研修", "出張", "会議"}
    """
    if not text.strip():
        return set()

    return {
        word.strip()
        for word in text.split(",")
        if word.strip()
    }


def contains_skip_word(value, skip_duty_words: Optional[Iterable[str]] = None) -> bool:
    """
    セル内に、除外ワードが含まれていたら True
    """
    s = normalize_cell(value)
    if not s:
        return False

    skip_words = set(DEFAULT_SKIP_DUTY_WORDS if skip_duty_words is None else skip_duty_words)
    if not skip_words:
        return False

    return any(word in s for word in skip_words)


def is_eligible_for_cleaning(
    value,
    off_words: Optional[Iterable[str]] = None,
    skip_duty_words: Optional[Iterable[str]] = None,
) -> bool:
    """
    本番仕様:
    - ×系だけ休み
    - それ以外は出勤
    - ただし skip_duty_words に含まれる語がセル内にあれば担当から除外
    """
    s = normalize_cell(value)
    off_words = set(DEFAULT_OFF_WORDS if off_words is None else off_words)

    if s in off_words:
        return False

    if contains_skip_word(s, skip_duty_words=skip_duty_words):
        return False

    return True


def sort_day_columns(day_cols: list) -> list:
    def day_key(col):
        if isinstance(col, int):
            return col
        m = re.search(r"\d+", str(col))
        return int(m.group()) if m else 10**9

    return sorted(day_cols, key=day_key)


def assign_one_keep_order(queue: deque, eligible_staff_today: set[str]) -> Optional[str]:
    """
    上から順に見て、その日担当できる最初の人を選ぶ。
    選ばれた人だけ最後尾へ回す。
    休みだった人は順番を維持する。
    """
    original = list(queue)
    chosen = None
    idx = None

    for i, person in enumerate(original):
        if person in eligible_staff_today:
            chosen = person
            idx = i
            break

    if chosen is None:
        return None

    new_order = original[:idx] + original[idx + 1:] + [chosen]
    queue.clear()
    queue.extend(new_order)
    return chosen


def rotate_queue_from_last(queue: list[str], last_person: Optional[str]) -> list[str]:
    """
    last_person の次の人が先頭になるように queue を回転させる
    例:
    queue = [A, B, C, D], last_person = B
    -> [C, D, A, B]
    """
    if not queue or not last_person:
        return queue.copy()

    if last_person not in queue:
        return queue.copy()

    idx = queue.index(last_person)
    return queue[idx + 1:] + queue[:idx + 1]


def read_last_assignees(previous_schedule_path: str) -> tuple[Optional[str], Optional[str]]:
    """
    前回作成済みの当番表Excelから、
    最後に担当した男性・女性を読む
    対象シート: 日別担当表
    """
    try:
        df = pd.read_excel(previous_schedule_path, sheet_name="日別担当表")

        male_last = None
        female_last = None

        if "男性便所担当" in df.columns:
            male_series = df["男性便所担当"].dropna().astype(str).str.strip()
            if not male_series.empty:
                male_last = male_series.iloc[-1]

        if "女性便所担当" in df.columns:
            female_series = df["女性便所担当"].dropna().astype(str).str.strip()
            if not female_series.empty:
                female_last = female_series.iloc[-1]

        return male_last, female_last

    except Exception:
        return None, None


# =========================
# メインロジック
# =========================
def clean_staff_rows(
    df: pd.DataFrame,
    name_col,
    male_col,
    employee_col,
    day_cols: list,
) -> pd.DataFrame:
    """
    今回の『月間稼働計画表.xlsx』専用寄りの職員抽出

    ルール:
    - B列の名前が実在職員名らしい行だけ残す
    - 社員もここでは残す（後で除外判定するため）
    - ただし集計行や見出し行は除外
    - 勤務欄に × が一つもない名前は除外
    """
    work_df = df.copy()

    work_df[name_col] = (
        work_df[name_col]
        .fillna("")
        .astype(str)
        .str.replace("\u3000", " ", regex=False)
        .str.strip()
    )

    excluded_names = {
        "",
        "日付",
        "曜日",
        "月間稼働計画表",
        "行事予定",
        "８時からの人員",
        "20時までの人員",
        "２０時までの人員",
    }

    def is_real_name(value: str) -> bool:
        s = normalize_cell(value)

        if not s:
            return False

        if s in excluded_names:
            return False

        # 数字だけは名前ではない（例: 18）
        if re.fullmatch(r"\d+", s):
            return False

        # 数式・集計ワード除外
        upper_s = s.upper()
        if s.startswith("="):
            return False
        if "COUNTA" in upper_s:
            return False
        if "合計" in s:
            return False
        if "人員" in s:
            return False

        return True

    def has_schedule_signal(row) -> bool:
        if normalize_cell(row.get(male_col, "")) != "":
            return True
        if normalize_cell(row.get(employee_col, "")) != "":
            return True

        for day in day_cols:
            if normalize_cell(row.get(day, "")) != "":
                return True

        return False

    def has_off_mark(row) -> bool:
        """
        勤務表行に × / x / X / ｘ が一つもない名前は除外
        """
        for day in day_cols:
            if normalize_cell(row.get(day, "")) in DEFAULT_OFF_WORDS:
                return True
        return False

    valid_name_mask = work_df[name_col].apply(is_real_name)
    signal_mask = work_df.apply(has_schedule_signal, axis=1)
    off_mark_mask = work_df.apply(has_off_mark, axis=1)

    work_df = work_df[valid_name_mask & signal_mask & off_mark_mask].copy()

    return work_df


def build_schedule_from_row_staff_table(
    df: pd.DataFrame,
    name_col,
    male_col,
    employee_col,
    day_cols: list,
    off_words: Optional[Iterable[str]] = None,
    skip_duty_words: Optional[Iterable[str]] = None,
    previous_male_last: Optional[str] = None,
    previous_female_last: Optional[str] = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:

    work_df = clean_staff_rows(
        df=df,
        name_col=name_col,
        male_col=male_col,
        employee_col=employee_col,
        day_cols=day_cols,
    )

    work_df["__is_male"] = work_df[male_col].apply(is_checked)
    work_df["__is_employee"] = work_df[employee_col].apply(is_checked)

    male_candidates = work_df[
        (work_df["__is_male"]) & (~work_df["__is_employee"])
    ][name_col].tolist()

    female_candidates = work_df[
        (~work_df["__is_male"]) & (~work_df["__is_employee"])
    ][name_col].tolist()

    excluded_staff = work_df[
        work_df["__is_employee"]
    ][name_col].tolist()

    male_queue = deque(rotate_queue_from_last(male_candidates, previous_male_last))
    female_queue = deque(rotate_queue_from_last(female_candidates, previous_female_last))

    male_counts = {name: 0 for name in male_candidates}
    female_counts = {name: 0 for name in female_candidates}

    day_records = []

    for day in day_cols:
        male_eligible = set()
        female_eligible = set()

        for _, row in work_df.iterrows():
            name = row[name_col]
            status = row[day]

            # 社員は担当候補から除外
            if row["__is_employee"]:
                continue

            if not is_eligible_for_cleaning(
                status,
                off_words=off_words,
                skip_duty_words=skip_duty_words,
            ):
                continue

            if row["__is_male"]:
                male_eligible.add(name)
            else:
                female_eligible.add(name)

        male_person = assign_one_keep_order(
            queue=male_queue,
            eligible_staff_today=male_eligible,
        ) if male_candidates else None

        female_person = assign_one_keep_order(
            queue=female_queue,
            eligible_staff_today=female_eligible,
        ) if female_candidates else None

        if male_person:
            male_counts[male_person] += 1
        if female_person:
            female_counts[female_person] += 1

        remarks = []
        if male_candidates and not male_person:
            remarks.append("男性担当者なし")
        if female_candidates and not female_person:
            remarks.append("女性担当者なし")

        # E列=4 を 1日に見せる
        display_day = day - 3 if isinstance(day, int) else str(day)

        day_records.append({
            "日付": display_day,
            "男性便所担当": male_person,
            "女性便所担当": female_person,
            "備考": " / ".join(remarks),
        })

    result_df = pd.DataFrame(day_records)

    summary_rows = []

    for name in male_candidates:
        summary_rows.append({
            "職員名": name,
            "性別区分": "男性",
            "社員除外": "いいえ",
            "担当回数": male_counts[name],
        })

    for name in female_candidates:
        summary_rows.append({
            "職員名": name,
            "性別区分": "女性",
            "社員除外": "いいえ",
            "担当回数": female_counts[name],
        })

    for name in excluded_staff:
        gender = "男性" if bool(
            work_df.loc[work_df[name_col] == name, "__is_male"].iloc[0]
        ) else "女性"

        summary_rows.append({
            "職員名": name,
            "性別区分": gender,
            "社員除外": "はい",
            "担当回数": 0,
        })

    summary_df = pd.DataFrame(summary_rows).sort_values(
        ["社員除外", "性別区分", "担当回数", "職員名"],
        ascending=[True, True, False, True]
    ).reset_index(drop=True)

    return result_df, summary_df


def export_schedule_excel(
    input_path: str,
    output_path: str,
    sheet_name: str = "2026.3",
    off_words: Optional[Iterable[str]] = None,
    skip_duty_words: Optional[Iterable[str]] = None,
    previous_schedule_path: Optional[str] = None,
) -> str:
    """
    今回の添付Excel専用設定
    B列 = 氏名
    C列 = 社員
    D列 = 男性
    E列〜AI列 = 1日〜31日
    """
    df = pd.read_excel(input_path, sheet_name=sheet_name, header=None)

    previous_male_last = None
    previous_female_last = None

    if previous_schedule_path:
        previous_male_last, previous_female_last = read_last_assignees(previous_schedule_path)

    name_col = 1
    employee_col = 2
    male_col = 3
    day_cols = list(range(4, 35))

    result_df, summary_df = build_schedule_from_row_staff_table(
        df=df,
        name_col=name_col,
        male_col=male_col,
        employee_col=employee_col,
        day_cols=day_cols,
        off_words=off_words,
        skip_duty_words=skip_duty_words,
        previous_male_last=previous_male_last,
        previous_female_last=previous_female_last,
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        result_df.to_excel(writer, sheet_name="日別担当表", index=False)
        summary_df.to_excel(writer, sheet_name="担当回数集計", index=False)

    return output_path