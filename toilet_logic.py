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
    △ / x / ○ / ✓ / 1 / True などをチェックありとして扱う
    """
    if pd.isna(value):
        return False

    if isinstance(value, bool):
        return value

    if isinstance(value, (int, float)):
        return value == 1

    s = str(value).strip().lower()
    return s in {
        "1", "true", "yes", "y", "on", "checked",
        "○", "✓", "✔", "レ", "△", "x", "ｘ",
        "男", "男性", "社員"
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


def sort_day_columns(day_cols: list[str]) -> list[str]:
    def day_key(col):
        m = re.search(r"\d+", str(col))
        return int(m.group()) if m else 10**9
    return sorted(day_cols, key=day_key)


def find_first_existing(columns: list[str], candidates: list[str]) -> Optional[str]:
    for c in candidates:
        if c in columns:
            return c
    return None


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

def assign_one_fair(
    eligible_staff_today: set[str],
    assigned_counts: dict[str, int],
    last_assigned_day: dict[str, int],
    base_order: dict[str, int],
    current_day_index: int,
) -> Optional[str]:
    """
    公平性優先で1人選ぶ

    優先順位:
    1. 担当回数が少ない
    2. 最後に担当した日が古い（長く担当していない）
    3. 元の並び順が早い
    """
    if not eligible_staff_today:
        return None



    def sort_key(name: str):
        return (
            assigned_counts[name],         # 少ない人を優先
            last_assigned_day[name],       # 小さいほど古い（未担当は -1）
            base_order[name],              # 元の並び順
        )

    chosen = min(eligible_staff_today, key=sort_key)

    assigned_counts[chosen] += 1
    last_assigned_day[chosen] = current_day_index

    return chosen

def assign_one_fair_by_ratio(
    eligible_staff_today: set[str],
    assigned_counts: dict[str, int],
    eligible_days: dict[str, int],
    last_assigned_day: dict[str, int],
    base_order: dict[str, int],
    current_day_index: int,
) -> Optional[str]:
    """
    負担率ベースで1人選ぶ

    優先順位:
    1. 担当率（担当回数 / 出勤可能日数）が低い
    2. 担当回数が少ない
    3. 最後に担当した日が古い
    4. 元の並び順が早い
    """
    if not eligible_staff_today:
        return None

    def load_ratio(name: str) -> float:
        days = eligible_days.get(name, 0)
        if days <= 0:
            return float("inf")
        return assigned_counts[name] / days

    def sort_key(name: str):
        return (
            load_ratio(name),          # 負担率が低い人
            assigned_counts[name],     # 同率なら担当回数が少ない人
            last_assigned_day[name],   # 同率なら長く担当していない人
            base_order[name],          # 最後は元の順番
        )

    chosen = min(eligible_staff_today, key=sort_key)

    assigned_counts[chosen] += 1
    last_assigned_day[chosen] = current_day_index

    return chosen

# =========================
# メインロジック
# =========================

def clean_staff_rows(
    df: pd.DataFrame,
    name_col: str,
    male_col: str,
    employee_col: str,
    day_cols: list[str],
) -> pd.DataFrame:
    """
    本当に職員行だけを残す
    条件:
    - 氏名が空でない
    - かつ、チェック列または日付列に何らかの値がある
    """
    work_df = df.copy()

    # 氏名を正規化
    work_df[name_col] = (
        work_df[name_col]
        .fillna("")
        .astype(str)
        .str.replace("\u3000", " ", regex=False)  # 全角スペース対策
        .str.strip()
    )

    # 氏名が有効か
    valid_name_mask = work_df[name_col] != ""

    # その行に勤務表らしい情報があるか
    def has_schedule_signal(row) -> bool:
        # 男性チェック・社員チェックがある
        if normalize_cell(row.get(male_col, "")) != "":
            return True
        if normalize_cell(row.get(employee_col, "")) != "":
            return True

        # 1〜31日のどこかに値がある
        for day in day_cols:
            if normalize_cell(row.get(day, "")) != "":
                return True

        return False

    signal_mask = work_df.apply(has_schedule_signal, axis=1)

    # 両方満たす行だけ残す
    work_df = work_df[valid_name_mask & signal_mask].copy()

    return work_df

def build_schedule_from_row_staff_table(
    df: pd.DataFrame,
    name_col: str,
    male_col: str,
    employee_col: str,
    day_cols: list[str],
    off_words: Optional[Iterable[str]] = None,
    skip_duty_words: Optional[Iterable[str]] = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    # work_df = df.copy()
    # work_df[name_col] = work_df[name_col].astype(str).str.strip()

    # # 名前が空欄の行は職員として扱わない
    # work_df = work_df[work_df[name_col] != ""].copy()

    work_df = clean_staff_rows(
        df=df,
        name_col=name_col,
        male_col=male_col,
        employee_col=employee_col,
        day_cols=day_cols,
    )

    work_df["__is_male"] = work_df[male_col].apply(is_checked)
    work_df["__is_employee"] = work_df[employee_col].apply(is_checked)

    eligible_days_all = count_eligible_days(
        work_df=work_df,
        name_col=name_col,
        day_cols=day_cols,
        off_words=off_words,
        skip_duty_words=skip_duty_words,
    )

    male_candidates = work_df[
        (work_df["__is_male"]) & (~work_df["__is_employee"])
    ][name_col].tolist()

    male_eligible_days = {name: eligible_days_all[name] for name in male_candidates}

    female_candidates = work_df[
        (~work_df["__is_male"]) & (~work_df["__is_employee"])
    ][name_col].tolist()

    female_eligible_days = {name: eligible_days_all[name] for name in female_candidates}

    excluded_staff = work_df[work_df["__is_employee"]][name_col].tolist()

    male_counts = {name: 0 for name in male_candidates}
    female_counts = {name: 0 for name in female_candidates}

    male_last_assigned = {name: -1 for name in male_candidates}
    female_last_assigned = {name: -1 for name in female_candidates}

    male_base_order = {name: i for i, name in enumerate(male_candidates)}
    female_base_order = {name: i for i, name in enumerate(female_candidates)}

    day_records = []

    for day_index, day in enumerate(day_cols):
        male_eligible = set()
        female_eligible = set()

        for _, row in work_df.iterrows():
            name = row[name_col]
            status = row[day]

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

        male_queue = male_candidates.copy()
        female_queue = female_candidates.copy()
        
        male_person = assign_one_fair_by_ratio_with_order(
            queue=male_queue,
            eligible_staff_today=male_eligible,
            assigned_counts=male_counts,
            eligible_days=male_eligible_days,
            last_assigned_day=male_last_assigned,
            current_day_index=day_index,
        ) if male_candidates else None

        female_person = assign_one_fair_by_ratio_with_order(
            queue=female_queue,
            eligible_staff_today=female_eligible,
            assigned_counts=female_counts,
            eligible_days=female_eligible_days,
            last_assigned_day=female_last_assigned,
            current_day_index=day_index,
        ) if female_candidates else None
        

        remarks = []
        if male_candidates and not male_person:
            remarks.append("男性担当者なし")
        if female_candidates and not female_person:
            remarks.append("女性担当者なし")

        day_records.append({
            "日付": day,
            "男性便所担当": male_person,
            "女性便所担当": female_person,
            "備考": " / ".join(remarks)
        })

    result_df = pd.DataFrame(day_records)

    summary_rows = []
    for name in male_candidates:
        summary_rows.append({
            "職員名": name,
            "性別区分": "男性",
            "社員除外": "いいえ",
            "出勤可能日数": male_eligible_days[name],
            "担当回数": male_counts[name],
            "担当率": round(male_counts[name] / male_eligible_days[name], 4) if male_eligible_days[name] > 0 else 0,
        })

    for name in female_candidates:
        summary_rows.append({
            "職員名": name,
            "性別区分": "女性",
            "社員除外": "いいえ",
            "出勤可能日数": female_eligible_days[name],
            "担当回数": female_counts[name],
            "担当率": round(female_counts[name] / female_eligible_days[name], 4) if female_eligible_days[name] > 0 else 0,
        })

    for name in excluded_staff:
        gender = "男性" if bool(
            work_df.loc[work_df[name_col] == name, "__is_male"].iloc[0]
        ) else "女性"

        summary_rows.append({
            "職員名": name,
            "性別区分": gender,
            "社員除外": "はい",
            "出勤可能日数": 0,
            "担当回数": 0,
            "担当率": 0,
        })
    summary_df = pd.DataFrame(summary_rows).sort_values(
        ["社員除外", "性別区分", "担当率", "担当回数", "職員名"],
        ascending=[True, True, False, False, True]
    ).reset_index(drop=True)

    return result_df, summary_df


def export_schedule_excel(
    input_path: str,
    output_path: str,
    sheet_name=0,
    name_col: Optional[str] = None,
    male_col: Optional[str] = None,
    employee_col: Optional[str] = None,
    day_cols: Optional[list[str]] = None,
    off_words: Optional[Iterable[str]] = None,
    skip_duty_words: Optional[Iterable[str]] = None,
) -> str:
    df = pd.read_excel(input_path, sheet_name=sheet_name)
    cols = list(df.columns)

    if name_col is None:
        name_col = find_first_existing(cols, ["氏名", "名前", "職員名"])
    if male_col is None:
        male_col = find_first_existing(cols, ["男性チェック", "男性", "男"])
    if employee_col is None:
        employee_col = find_first_existing(cols, ["社員チェック", "社員", "正社員"])

    if not name_col or not male_col or not employee_col:
        raise ValueError(
            "氏名列 / 男性チェック列 / 社員チェック列 が見つかりません。\n"
            "Excelにこの3列を追加してください。"
        )

    if day_cols is None:
        candidates = []
        for c in cols:
            s = str(c).strip()
            if s in {name_col, male_col, employee_col}:
                continue
            if (
                re.fullmatch(r"\d{1,2}", s)
                or re.fullmatch(r"day_?\d{1,2}", s, re.I)
                or re.fullmatch(r"\d{1,2}日", s)
            ):
                candidates.append(c)

        day_cols = sort_day_columns(candidates)

    if not day_cols:
        raise ValueError("日付列（1〜31）が見つかりません。")

    result_df, summary_df = build_schedule_from_row_staff_table(
        df=df,
        name_col=name_col,
        male_col=male_col,
        employee_col=employee_col,
        day_cols=day_cols,
        off_words=off_words,
        skip_duty_words=skip_duty_words,
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        result_df.to_excel(writer, sheet_name="日別担当表", index=False)
        summary_df.to_excel(writer, sheet_name="担当回数集計", index=False)

        sheet = writer.book["担当回数集計"]

        # 担当率列を％表示にする
        for cell in sheet["F"][1:]:  # F列 = 担当率（1行目はヘッダ）
            cell.number_format = "0.00%"

    return output_path

def count_eligible_days(
    work_df: pd.DataFrame,
    name_col: str,
    day_cols: list[str],
    off_words: Optional[Iterable[str]] = None,
    skip_duty_words: Optional[Iterable[str]] = None,
) -> dict[str, int]:
    """
    各職員の「掃除担当候補になれる日数」を数える
    """
    eligible_days = {}

    for _, row in work_df.iterrows():
        name = row[name_col]

        if row["__is_employee"]:
            eligible_days[name] = 0
            continue

        count = 0
        for day in day_cols:
            status = row[day]
            if is_eligible_for_cleaning(
                status,
                off_words=off_words,
                skip_duty_words=skip_duty_words,
            ):
                count += 1

        eligible_days[name] = count

    return eligible_days

def assign_one_fair_by_ratio_with_order(
    queue: list[str],
    eligible_staff_today: set[str],
    assigned_counts: dict[str, int],
    eligible_days: dict[str, int],
    last_assigned_day: dict[str, int],
    current_day_index: int,
) -> Optional[str]:
    """
    負担率を優先しつつ、同率なら元の順番を優先する版
    """
    if not eligible_staff_today:
        return None

    def load_ratio(name: str) -> float:
        days = eligible_days.get(name, 0)
        if days <= 0:
            return float("inf")
        return assigned_counts[name] / days

    # その日の候補者の中で最小の担当率
    min_ratio = min(load_ratio(name) for name in eligible_staff_today)

    # 最小担当率の人たちだけ残す
    ratio_candidates = {
        name for name in eligible_staff_today
        if load_ratio(name) == min_ratio
    }

    # その中から queue 上で最も前にいる人を選ぶ
    chosen = None
    for name in queue:
        if name in ratio_candidates:
            chosen = name
            break

    if chosen is None:
        return None

    # 担当回数更新
    assigned_counts[chosen] += 1
    last_assigned_day[chosen] = current_day_index

    # queue の最後に回す
    queue.remove(chosen)
    queue.append(chosen)

    return chosen