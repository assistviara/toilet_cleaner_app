from pathlib import Path

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from toilet_logic import export_schedule_excel, parse_skip_words


def format_summary_text(summary_df):
    """
    担当回数集計 DataFrame を、画面表示用の文字列に変換する
    非社員のみを対象に表示
    """
    visible_df = summary_df[summary_df["社員除外"] == "いいえ"].copy()

    male_df = visible_df[visible_df["性別区分"] == "男性"]
    female_df = visible_df[visible_df["性別区分"] == "女性"]

    lines = []
    lines.append("担当回数（非社員のみ）")

    lines.append("\n男性:")
    if male_df.empty:
        lines.append("  該当なし")
    else:
        for _, row in male_df.iterrows():
            lines.append(f"  {row['職員名']}  {row['担当回数']}回")

    lines.append("\n女性:")
    if female_df.empty:
        lines.append("  該当なし")
    else:
        for _, row in female_df.iterrows():
            lines.append(f"  {row['職員名']}  {row['担当回数']}回")

    return "\n".join(lines)


class ToiletScheduleApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("トイレ掃除表 作成ツール")
        self.root.geometry("820x620")

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.previous_schedule_var = tk.StringVar()
        self.skip_words_var = tk.StringVar()
        self.status_var = tk.StringVar(value="勤務表Excelを選択してください。")

        self.sheet_names = []

        self._build_ui()

    def _build_ui(self):
        frame = tk.Frame(self.root, padx=16, pady=16)
        frame.pack(fill="both", expand=True)

        title = tk.Label(
            frame,
            text="トイレ掃除表 作成ツール",
            font=("Meiryo", 15, "bold")
        )
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))

        desc = tk.Label(
            frame,
            text=(
                "入力Excelは次の列構成を想定しています。\n"
                "B列=氏名 / C列=社員 / D列=男性 / E〜AI列=1〜31日\n"
                "勤務表の中では × 印だけ休み扱いです。\n"
                "また、掃除対象外ワードを含むセルは担当から外れます。\n"
                "前回の当番表を指定すると、最後の担当者の次から開始できます。"
            ),
            justify="left",
            anchor="w"
        )
        desc.grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 14))

        tk.Label(frame, text="勤務表Excel").grid(row=2, column=0, sticky="w")
        entry_input = tk.Entry(frame, textvariable=self.input_var, width=72)
        entry_input.grid(row=3, column=0, columnspan=2, sticky="we", padx=(0, 8), pady=(4, 10))
        tk.Button(frame, text="参照", width=10, command=self.select_input).grid(row=3, column=2, sticky="e", pady=(4, 10))

        tk.Label(frame, text="対象シート").grid(row=4, column=0, sticky="w")
        self.sheet_combo = ttk.Combobox(
            frame,
            textvariable=self.sheet_var,
            state="readonly",
            width=30
        )
        self.sheet_combo.grid(row=5, column=0, columnspan=3, sticky="we", pady=(4, 10))

        tk.Label(frame, text="出力Excel").grid(row=6, column=0, sticky="w")
        entry_output = tk.Entry(frame, textvariable=self.output_var, width=72)
        entry_output.grid(row=7, column=0, columnspan=2, sticky="we", padx=(0, 8), pady=(4, 10))
        tk.Button(frame, text="保存先", width=10, command=self.select_output).grid(row=7, column=2, sticky="e", pady=(4, 10))

        tk.Label(frame, text="前回の当番表Excel（任意）").grid(row=8, column=0, sticky="w")
        entry_prev = tk.Entry(frame, textvariable=self.previous_schedule_var, width=72)
        entry_prev.grid(row=9, column=0, columnspan=2, sticky="we", padx=(0, 8), pady=(4, 10))
        tk.Button(frame, text="参照", width=10, command=self.select_previous_schedule).grid(row=9, column=2, sticky="e", pady=(4, 10))

        tk.Label(frame, text="掃除対象外ワード（カンマ区切り）").grid(row=10, column=0, sticky="w")
        entry_skip = tk.Entry(frame, textvariable=self.skip_words_var, width=72)
        entry_skip.grid(row=11, column=0, columnspan=3, sticky="we", pady=(4, 16))
        entry_skip.insert(0, "")

        button_frame = tk.Frame(frame)
        button_frame.grid(row=12, column=0, columnspan=3, sticky="w", pady=(0, 16))

        tk.Button(
            button_frame,
            text="掃除表を作成",
            width=18,
            height=2,
            command=self.run_generation
        ).pack(side="left")

        tk.Button(
            button_frame,
            text="終了",
            width=10,
            command=self.root.destroy
        ).pack(side="left", padx=(12, 0))

        status_title = tk.Label(frame, text="状態")
        status_title.grid(row=13, column=0, sticky="w")

        status_box = tk.Label(
            frame,
            textvariable=self.status_var,
            justify="left",
            anchor="nw",
            relief="sunken",
            bg="white",
            width=96,
            height=12
        )
        status_box.grid(row=14, column=0, columnspan=3, sticky="nsew", pady=(4, 0))

        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(14, weight=1)

    def select_input(self):
        path = filedialog.askopenfilename(
            title="勤務表Excelを選択",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if path:
            self.input_var.set(path)

            input_path = Path(path)
            default_output = input_path.with_name(f"{input_path.stem}_掃除表.xlsx")
            if not self.output_var.get().strip():
                self.output_var.set(str(default_output))

            try:
                excel_file = pd.ExcelFile(path)
                self.sheet_names = excel_file.sheet_names
                self.sheet_combo["values"] = self.sheet_names

                if self.sheet_names:
                    # 2026.3 があれば優先、なければ先頭
                    if "2026.3" in self.sheet_names:
                        self.sheet_var.set("2026.3")
                    else:
                        self.sheet_var.set(self.sheet_names[0])

            except Exception as e:
                messagebox.showerror("シート取得エラー", f"シート一覧の取得に失敗しました。\n{e}")
                self.sheet_combo["values"] = []
                self.sheet_var.set("")

            self.status_var.set(f"入力ファイルを選択しました。\n{path}")

    def select_output(self):
        initial = self.output_var.get().strip()
        initialfile = Path(initial).name if initial else "トイレ掃除表.xlsx"

        path = filedialog.asksaveasfilename(
            title="出力先を選択",
            defaultextension=".xlsx",
            initialfile=initialfile,
            filetypes=[("Excel files", "*.xlsx")]
        )
        if path:
            self.output_var.set(path)
            self.status_var.set(f"出力先を設定しました。\n{path}")

    def select_previous_schedule(self):
        path = filedialog.askopenfilename(
            title="前回の当番表Excelを選択",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if path:
            self.previous_schedule_var.set(path)
            self.status_var.set(f"前回当番表を選択しました。\n{path}")

    def run_generation(self):
        input_file = self.input_var.get().strip()
        output_file = self.output_var.get().strip()
        sheet_name = self.sheet_var.get().strip()
        previous_schedule_file = self.previous_schedule_var.get().strip()
        skip_words = parse_skip_words(self.skip_words_var.get())

        if not input_file:
            messagebox.showwarning("注意", "勤務表Excelを選択してください。")
            return

        if not output_file:
            messagebox.showwarning("注意", "出力Excelの保存先を指定してください。")
            return

        if not sheet_name:
            messagebox.showwarning("注意", "対象シートを選択してください。")
            return

        try:
            self.status_var.set("処理中です。少し待ってください...")
            self.root.update_idletasks()

            created = export_schedule_excel(
                input_path=input_file,
                output_path=output_file,
                sheet_name=sheet_name,
                skip_duty_words=skip_words,
                previous_schedule_path=previous_schedule_file if previous_schedule_file else None,
            )

            skip_info = "なし" if not skip_words else ", ".join(sorted(skip_words))
            prev_info = previous_schedule_file if previous_schedule_file else "なし"

            summary_df = pd.read_excel(created, sheet_name="担当回数集計")
            summary_text = format_summary_text(summary_df)

            self.status_var.set(
                "作成が完了しました。\n\n"
                f"対象シート:\n{sheet_name}\n\n"
                f"前回当番表:\n{prev_info}\n\n"
                f"出力ファイル:\n{created}\n\n"
                f"掃除対象外ワード:\n{skip_info}\n\n"
                f"{summary_text}"
            )
            messagebox.showinfo("完了", f"掃除表を作成しました。\n{created}")

        except PermissionError:
            msg = (
                "出力先のExcelファイルに書き込めません。\n\n"
                "次の原因が考えられます。\n"
                "・出力先のExcelファイルが開いたまま\n"
                "・同名ファイルがロックされている\n\n"
                "Excelを閉じてから、もう一度実行してください。"
            )
            self.status_var.set(msg)
            messagebox.showerror("保存エラー", msg)

        except FileNotFoundError:
            msg = (
                "入力ファイルが見つかりません。\n\n"
                "ファイルが移動または削除されている可能性があります。\n"
                "もう一度、勤務表Excelを選択してください。"
            )
            self.status_var.set(msg)
            messagebox.showerror("ファイルエラー", msg)

        except Exception as e:
            msg = f"予期しないエラーが発生しました。\n\n{e}"
            self.status_var.set(msg)
            messagebox.showerror("エラー", msg)


def main():
    root = tk.Tk()
    app = ToiletScheduleApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()