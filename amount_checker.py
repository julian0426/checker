import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from decimal import Decimal, InvalidOperation
import pandas as pd

# Excel columns (0-based): AB=27, AQ=42, AV=47, BG=58
COL_TAX_FREE = 27
COL_AMOUNT = 42
COL_CUSTOMER = 47
COL_TYPE = 58

CATEGORIES = ("과세", "면세", "인보이스")


def to_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def to_amount(value):
    if pd.isna(value) or value == "":
        return 0.0
    if isinstance(value, str):
        value = value.replace(",", "").replace("원", "").strip()
        if not value:
            return 0.0
    try:
        return float(Decimal(str(value)))
    except (InvalidOperation, ValueError):
        raise ValueError(f"금액으로 변환할 수 없는 값입니다: {value!r}")


def normalize_code(value):
    s = to_text(value)
    if s.endswith(".0"):
        s = s[:-2]
    return s


class AmountChecker:
    def __init__(self, root):
        self.root = root
        self.root.title("고객사별 매출 금액 대조 프로그램")
        self.root.geometry("1250x760")
        self.root.minsize(1000, 620)
        self.source_path = None
        self.expected = {}
        self.rows = {}
        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=12)
        top.pack(fill="x")
        ttk.Label(top, text="고객사별 매출 금액 대조", font=("맑은 고딕", 17, "bold")).pack(side="left")
        ttk.Button(top, text="엑셀 업로드", command=self.load_excel).pack(side="right")

        self.file_label = ttk.Label(self.root, text="업로드된 파일이 없습니다.", padding=(14, 0))
        self.file_label.pack(anchor="w")

        rules = (
            "집계 기준: AQ열 금액 / AV열 고객사 / AB열 과세·면세 / BG열 수주유형코드\n"
            "과세·면세: 801 주문(+), 811 취소(-)  |  인보이스: 851 주문(+), 861 취소(-)"
        )
        ttk.Label(self.root, text=rules, padding=14, foreground="#555").pack(anchor="w")

        action = ttk.Frame(self.root, padding=(12, 0, 12, 8))
        action.pack(fill="x")
        ttk.Button(action, text="고객사 행 추가", command=self.add_customer).pack(side="left")
        ttk.Button(action, text="입력 금액 초기화", command=self.clear_inputs).pack(side="left", padx=6)
        ttk.Button(action, text="비교 결과 엑셀 저장", command=self.export_result).pack(side="right")

        # Requested order: all Excel totals, then all submitted amounts, then all differences.
        columns = (
            "customer",
            "tax_expected", "exempt_expected", "invoice_expected",
            "tax_input", "exempt_input", "invoice_input",
            "tax_diff", "exempt_diff", "invoice_diff"
        )
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings", height=22)

        headers = {
            "customer": "고객사명",
            "tax_expected": "과세(엑셀)",
            "exempt_expected": "면세(엑셀)",
            "invoice_expected": "인보이스(엑셀)",
            "tax_input": "과세(입력)",
            "exempt_input": "면세(입력)",
            "invoice_input": "인보이스(입력)",
            "tax_diff": "과세 차액",
            "exempt_diff": "면세 차액",
            "invoice_diff": "인보이스 차액"
        }
        widths = {
            "customer": 180,
            "tax_expected": 115, "exempt_expected": 115, "invoice_expected": 125,
            "tax_input": 115, "exempt_input": 115, "invoice_input": 125,
            "tax_diff": 115, "exempt_diff": 115, "invoice_diff": 125
        }
        for c in columns:
            self.tree.heading(c, text=headers[c])
            self.tree.column(
                c, width=widths[c], minwidth=85,
                anchor="w" if c == "customer" else "e",
                stretch=False
            )

        # Put the vertical scrollbar directly beside the table; horizontal scrollbar below it.
        table_frame = ttk.Frame(self.root)
        table_frame.pack(fill="both", expand=True, padx=12, pady=(0, 0))
        self.tree.master = table_frame
        self.tree.pack(side="left", fill="both", expand=True)

        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        yscroll.pack(side="right", fill="y")
        xscroll = ttk.Scrollbar(self.root, orient="horizontal", command=self.tree.xview)
        xscroll.pack(fill="x", padx=12)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.tree.bind("<Double-1>", self.edit_cell)

        self.summary = ttk.Label(
            self.root,
            text="엑셀을 업로드하면 고객사별 집계 결과가 표시됩니다.",
            padding=12
        )
        self.summary.pack(anchor="w")

        ttk.Label(
            self.root,
            text="입력 금액 칸(과세/면세/인보이스)을 더블클릭해 직접 입력하세요. 차액 = 입력 금액 - 엑셀 집계 금액.",
            padding=(12, 0, 12, 12),
            foreground="#555"
        ).pack(anchor="w")

    def load_excel(self):
        path = filedialog.askopenfilename(
            title="매출 데이터 엑셀 선택",
            filetypes=[("Excel 파일", "*.xlsx *.xlsm *.xls"), ("모든 파일", "*.*")]
        )
        if not path:
            return
        try:
            suffix = Path(path).suffix.lower()
            engine = "xlrd" if suffix == ".xls" else "openpyxl"
            df = pd.read_excel(path, header=None, dtype=object, engine=engine)
            required = max(COL_TAX_FREE, COL_AMOUNT, COL_CUSTOMER, COL_TYPE)
            if df.shape[1] <= required:
                raise ValueError(
                    f"엑셀 열 수가 부족합니다. 최소 BG열(59열)까지 있어야 합니다. "
                    f"현재 열 수: {df.shape[1]}"
                )

            self.expected = {}
            skipped = 0
            for idx, row in df.iterrows():
                code = normalize_code(row.iloc[COL_TYPE])
                customer = to_text(row.iloc[COL_CUSTOMER])
                if not customer or code not in ("801", "811", "851", "861"):
                    continue

                try:
                    amount = abs(to_amount(row.iloc[COL_AMOUNT]))
                except ValueError as exc:
                    raise ValueError(f"{idx + 1}행 AQ열: {exc}")

                if code in ("801", "811"):
                    tax_type = to_text(row.iloc[COL_TAX_FREE])
                    if tax_type == "과세":
                        category = "과세"
                    elif tax_type == "면세":
                        category = "면세"
                    else:
                        skipped += 1
                        continue
                else:
                    category = "인보이스"

                signed = -amount if code in ("811", "861") else amount
                self.expected.setdefault(customer, {c: 0.0 for c in CATEGORIES})
                self.expected[customer][category] += signed

            if not self.expected:
                raise ValueError("집계 대상 데이터가 없습니다. 고객사명, AQ열 금액, AB열, BG열을 확인하세요.")

            self.source_path = path
            self.file_label.config(text=f"파일: {Path(path).name}")
            self.rows = {}
            self.tree.delete(*self.tree.get_children())
            for customer in sorted(self.expected):
                self._insert_customer(customer, {"과세": "", "면세": "", "인보이스": ""})
            self._refresh()
            messagebox.showinfo(
                "업로드 완료",
                f"고객사 {len(self.expected)}개 집계 완료\n"
                f"AB열 구분값이 과세/면세가 아니어서 제외된 행: {skipped}건"
            )
        except Exception as exc:
            messagebox.showerror("엑셀 처리 오류", str(exc))

    def _insert_customer(self, customer, inputs=None):
        inputs = inputs or {"과세": "", "면세": "", "인보이스": ""}
        self.rows[customer] = inputs
        self.tree.insert("", "end", iid=customer, values=self._values(customer))

    def add_customer(self):
        win = tk.Toplevel(self.root)
        win.title("고객사 추가")
        win.resizable(False, False)
        ttk.Label(win, text="고객사명을 입력하세요.").pack(padx=14, pady=(14, 5))
        entry = ttk.Entry(win, width=32)
        entry.pack(padx=14, pady=5)
        entry.focus_set()

        def add():
            name = entry.get().strip()
            if not name:
                return
            if name in self.rows:
                messagebox.showwarning("중복 고객사", "이미 등록된 고객사입니다.", parent=win)
                return
            self.expected.setdefault(name, {c: 0.0 for c in CATEGORIES})
            self._insert_customer(name)
            self._refresh()
            win.destroy()

        ttk.Button(win, text="추가", command=add).pack(pady=12)
        entry.bind("<Return>", lambda e: add())

    def clear_inputs(self):
        for customer in self.rows:
            for category in CATEGORIES:
                self.rows[customer][category] = ""
        self._refresh()

    def edit_cell(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item:
            return
        col_index = int(col[1:]) - 1
        # Editable columns: 과세/면세/인보이스 input fields (columns 5, 6, 7)
        mapping = {5: "과세", 6: "면세", 7: "인보이스"}
        if col_index not in mapping:
            return
        category = mapping[col_index]
        bbox = self.tree.bbox(item, col)
        if not bbox:
            return

        x, y, w, h = bbox
        current = self.rows[item][category]
        editor = ttk.Entry(self.tree)
        editor.place(x=x, y=y, width=w, height=h)
        editor.insert(0, str(current))
        editor.select_range(0, "end")
        editor.focus_set()

        def save(_=None):
            value = editor.get().strip().replace(",", "").replace("원", "")
            if value:
                try:
                    Decimal(value)
                except InvalidOperation:
                    messagebox.showerror("입력 오류", "금액은 숫자로 입력하세요.")
                    editor.focus_set()
                    return
            self.rows[item][category] = value
            editor.destroy()
            self._refresh()

        editor.bind("<Return>", save)
        editor.bind("<FocusOut>", save)
        editor.bind("<Escape>", lambda e: editor.destroy())

    def _values(self, customer):
        exp = self.expected.get(customer, {c: 0.0 for c in CATEGORIES})
        inp = self.rows.get(customer, {c: "" for c in CATEGORIES})
        vals = [customer]

        # Excel totals
        for category in CATEGORIES:
            vals.append(self._fmt(exp[category]))

        # User-entered amounts
        for category in CATEGORIES:
            raw = inp[category]
            entered = to_amount(raw) if raw != "" else None
            vals.append("" if entered is None else self._fmt(entered))

        # Differences
        for category in CATEGORIES:
            raw = inp[category]
            entered = to_amount(raw) if raw != "" else None
            diff = entered - exp[category] if entered is not None else None
            vals.append("" if diff is None else self._fmt(diff))

        return vals

    @staticmethod
    def _fmt(value):
        return f"{value:,.2f}".rstrip("0").rstrip(".")

    def _refresh(self):
        for customer in self.rows:
            vals = self._values(customer)
            if self.tree.exists(customer):
                self.tree.item(customer, values=vals)
            else:
                self.tree.insert("", "end", iid=customer, values=vals)

        count = len(self.rows)
        filled = sum(
            1 for c in self.rows
            if all(self.rows[c][category] != "" for category in CATEGORIES)
        )
        self.summary.config(text=f"고객사 {count}개  |  제출 금액 3개 항목 입력 완료 {filled}개")

    def export_result(self):
        if not self.rows:
            messagebox.showwarning("저장할 데이터 없음", "먼저 엑셀을 업로드하세요.")
            return

        path = filedialog.asksaveasfilename(
            title="비교 결과 저장",
            defaultextension=".xlsx",
            initialfile="고객사별_금액대조결과.xlsx",
            filetypes=[("Excel 파일", "*.xlsx")]
        )
        if not path:
            return

        records = []
        for customer in sorted(self.rows):
            exp = self.expected.get(customer, {c: 0.0 for c in CATEGORIES})
            inp = self.rows[customer]
            rec = {"고객사명": customer}

            # Export in the same grouped order as the on-screen table.
            for category in CATEGORIES:
                rec[f"{category}_엑셀"] = exp[category]
            for category in CATEGORIES:
                rec[f"{category}_입력"] = to_amount(inp[category]) if inp[category] != "" else None
            for category in CATEGORIES:
                entered = to_amount(inp[category]) if inp[category] != "" else None
                rec[f"{category}_차액"] = None if entered is None else entered - exp[category]

            records.append(rec)

        try:
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                pd.DataFrame(records).to_excel(writer, index=False, sheet_name="금액 대조")
                ws = writer.sheets["금액 대조"]
                ws.freeze_panes = "B2"
                ws.auto_filter.ref = ws.dimensions
                for col in ws.columns:
                    letter = col[0].column_letter
                    max_len = max(len(str(cell.value or "")) for cell in col)
                    ws.column_dimensions[letter].width = min(max(max_len + 3, 14), 28)
            messagebox.showinfo("저장 완료", f"비교 결과를 저장했습니다.\n{path}")
        except Exception as exc:
            messagebox.showerror("저장 오류", str(exc))


if __name__ == "__main__":
    root = tk.Tk()
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    app = AmountChecker(root)
    root.mainloop()
