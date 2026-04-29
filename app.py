import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import pandas as pd

from import_shipment import import_shipment_data
from import_sales import import_sales_data
from file_utils import ensure_data_dir, get_db_path, export_dataframe_to_excel


APP_TITLE = "出荷合計情報ツール（公開用サンプル）"
APP_VERSION = "v1.6"


class ShippingSummaryApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title(f"{APP_TITLE} {APP_VERSION}")
        self.geometry("620x260")
        self.resizable(False, False)

        ensure_data_dir()
        self.db_path = get_db_path()

        self._build_main_ui()

    # =========================================================
    # 共通：表示名 / 出力名変換
    # =========================================================
    def to_japanese_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        rename_map = {
            "customer_code": "届け先コード",
            "customer_name": "届け先名称",
            "row_count": "枚数",
            "item_code": "品目コード",
            "item_name": "品名",
            "order_no": "注文番号",
            "sheet_count": "枚数",
            "total_qty": "個数",
            "shipping_qty": "個数",
            "shelf_no": "棚番",
            "shipment_no": "売上№",
            "line_no": "行番号",
            "group_mark": "GRP",
        }
        return df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    def get_connection(self):
        return sqlite3.connect(self.db_path)

    # =========================================================
    # メイン画面
    # =========================================================
    def _build_main_ui(self):
        title = tk.Label(self, text="出荷合計情報ツール", font=("Yu Gothic UI", 18, "bold"))
        title.pack(pady=(15, 10))

        filter_frame = tk.Frame(self)
        filter_frame.pack(pady=5)

        today = pd.Timestamp.today()

        tk.Label(filter_frame, text="年月日", font=("Yu Gothic UI", 10)).grid(row=0, column=0, padx=5, pady=5)

        self.cmb_year = ttk.Combobox(filter_frame, width=8, state="readonly")
        self.cmb_year["values"] = [str(y) for y in range(today.year - 2, today.year + 2)]
        self.cmb_year.set(str(today.year))
        self.cmb_year.grid(row=0, column=1, padx=2)

        self.cmb_month = ttk.Combobox(filter_frame, width=5, state="readonly")
        self.cmb_month["values"] = [f"{m:02d}" for m in range(1, 13)]
        self.cmb_month.set(f"{today.month:02d}")
        self.cmb_month.grid(row=0, column=2, padx=2)

        self.cmb_day = ttk.Combobox(filter_frame, width=5, state="readonly")
        self.cmb_day["values"] = [f"{d:02d}" for d in range(1, 32)]
        self.cmb_day.set(f"{today.day:02d}")
        self.cmb_day.grid(row=0, column=3, padx=2)

        tk.Label(filter_frame, text="From", font=("Yu Gothic UI", 10)).grid(row=1, column=0, padx=5, pady=5)

        self.cmb_from_h = ttk.Combobox(filter_frame, width=5, state="readonly")
        self.cmb_from_h["values"] = [f"{h:02d}" for h in range(24)]
        self.cmb_from_h.set("00")
        self.cmb_from_h.grid(row=1, column=1, padx=2)

        self.cmb_from_m = ttk.Combobox(filter_frame, width=5, state="readonly")
        self.cmb_from_m["values"] = [f"{m:02d}" for m in range(0, 60, 5)]
        self.cmb_from_m.set("00")
        self.cmb_from_m.grid(row=1, column=2, padx=2)

        tk.Label(filter_frame, text="To", font=("Yu Gothic UI", 10)).grid(row=1, column=3, padx=10, pady=5)

        self.cmb_to_h = ttk.Combobox(filter_frame, width=5, state="readonly")
        self.cmb_to_h["values"] = [f"{h:02d}" for h in range(24)]
        self.cmb_to_h.set("23")
        self.cmb_to_h.grid(row=1, column=4, padx=2)

        self.cmb_to_m = ttk.Combobox(filter_frame, width=5, state="readonly")
        self.cmb_to_m["values"] = [f"{m:02d}" for m in range(0, 60, 5)]
        self.cmb_to_m.set("55")
        self.cmb_to_m.grid(row=1, column=5, padx=2)

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=20)

        tk.Button(
            btn_frame, text="出荷明細データ取込", width=18,
            command=self.load_shipment_data
        ).grid(row=0, column=0, padx=8, pady=5)

        tk.Button(
            btn_frame, text="売上データ取込", width=18,
            command=self.load_sales_data
        ).grid(row=0, column=1, padx=8, pady=5)

        tk.Button(
            btn_frame, text="届け先選択画面",
            width=18,
            command=self.open_customer_window
        ).grid(row=1, column=0, padx=8, pady=5)

        tk.Button(
            btn_frame, text="終了", width=18, fg="red",
            command=self.destroy
        ).grid(row=1, column=1, padx=8, pady=5)

    # =========================================================
    # 日付・時間条件
    # =========================================================
    def get_search_day(self):
        return int(self.cmb_year.get()) * 10000 + int(self.cmb_month.get()) * 100 + int(self.cmb_day.get())

    def get_from_time(self):
        return int(self.cmb_from_h.get()) * 10000 + int(self.cmb_from_m.get()) * 100

    def get_to_time(self):
        return int(self.cmb_to_h.get()) * 10000 + int(self.cmb_to_m.get()) * 100 + 59

    # =========================================================
    # データ取込
    # =========================================================
    def load_shipment_data(self):
        try:
            result = import_shipment_data(self.db_path)
            if result == 0:
                messagebox.showinfo("完了", "出荷明細データの取り込みが完了しました。")
        except Exception as e:
            messagebox.showerror("エラー", f"出荷明細データの取り込みに失敗しました。\n{e}")

    def load_sales_data(self):
        try:
            result = import_sales_data(self.db_path)
            if result == 0:
                messagebox.showinfo("完了", "売上データの取り込みが完了しました。")
        except Exception as e:
            messagebox.showerror("エラー", f"売上データの取り込みに失敗しました。\n{e}")

    # =========================================================
    # SQL
    # =========================================================
    def create_summary_sql(self, by_order=False):
        if by_order:
            # サンプルExcel完全一致：注番は「注番・品目コード・枚数・個数」のみ
            sql = """
                SELECT
                    order_no,
                    item_code,
                    COUNT(item_code) AS sheet_count,
                    SUM(shipping_qty) AS total_qty
                FROM shipment_detail_temp
                WHERE customer_code = ?
                  AND created_date = ?
                  AND created_time >= ?
                  AND created_time <= ?
                GROUP BY order_no, item_code
                ORDER BY order_no ASC, item_code ASC
            """
        else:
            # サンプルExcel完全一致：標準は「品目コード・枚数・個数」のみ
            sql = """
                SELECT
                    item_code,
                    COUNT(item_code) AS sheet_count,
                    SUM(shipping_qty) AS total_qty
                FROM shipment_detail_temp
                WHERE customer_code = ?
                  AND created_date = ?
                  AND created_time >= ?
                  AND created_time <= ?
                GROUP BY item_code
                ORDER BY item_code ASC
            """
        return sql

    def create_customer_list_sql(self):
        sql = """
            SELECT
                customer_code,
                customer_name,
                COUNT(*) AS row_count
            FROM shipment_detail_temp
            WHERE created_date = ?
              AND created_time >= ?
              AND created_time <= ?
            GROUP BY customer_code, customer_name
            ORDER BY row_count DESC, customer_name ASC
        """
        return sql

    def create_detail_sql(self):
        sql = """
            SELECT
                item_code,
                item_name,
                shipping_qty,
                order_no,
                shelf_no,
                shipment_no,
                line_no
            FROM shipment_detail_temp
            WHERE customer_code = ?
              AND created_date = ?
              AND created_time >= ?
              AND created_time <= ?
            ORDER BY item_code ASC, order_no ASC, shipment_no ASC, line_no ASC
        """
        return sql

    # =========================================================
    # 届け先選択画面
    # サンプル画面完全一致：届け先名称・枚数（届け先コードは表示しない）
    # =========================================================
    def open_customer_window(self):
        try:
            conn = self.get_connection()
            df = pd.read_sql_query(
                self.create_customer_list_sql(),
                conn,
                params=(
                    self.get_search_day(),
                    self.get_from_time(),
                    self.get_to_time(),
                ),
            )
            conn.close()
        except Exception as e:
            messagebox.showerror("エラー", f"届け先一覧の表示に失敗しました。\n{e}")
            return

        if df.empty:
            messagebox.showinfo("確認", "指定した年月日・時間帯のデータがありません。")
            return

        win = tk.Toplevel(self)
        win.title("届け先選択")
        win.geometry("760x540")

        lbl = tk.Label(win, text="届け先選択画面", font=("Yu Gothic UI", 16, "bold"))
        lbl.pack(pady=8)

        # 表示は2列のみ（サンプル画面完全一致）
        columns = ("customer_name", "row_count")
        tree = ttk.Treeview(win, columns=columns, show="headings", height=20)

        tree.heading("customer_name", text="届け先名称")
        tree.heading("row_count", text="枚数")

        tree.column("customer_name", width=620)
        tree.column("row_count", width=100, anchor=tk.E)

        # 内部処理用に customer_code は iid に保持
        for _, row in df.iterrows():
            tree.insert(
                "", "end",
                iid=str(row["customer_code"]),
                values=(row["customer_name"], row["row_count"])
            )

        vsb = ttk.Scrollbar(win, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)

        tree.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        vsb.pack(side="right", fill="y", pady=10)

        btn_frame = tk.Frame(win)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="選択（標準）", width=16, command=lambda: self.open_standard_summary(tree)).grid(row=0, column=0, padx=5)
        tk.Button(btn_frame, text="選択（注番）", width=16, command=lambda: self.open_order_summary(tree)).grid(row=0, column=1, padx=5)
        tk.Button(btn_frame, text="明細出力", width=16, command=lambda: self.export_detail(tree)).grid(row=0, column=2, padx=5)
        tk.Button(btn_frame, text="終了", width=16, fg="red", command=win.destroy).grid(row=0, column=3, padx=5)

    def get_selected_customer(self, tree):
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("確認", "届け先を選択してください。")
            return None, None

        item_id = selected[0]
        values = tree.item(item_id)["values"]

        customer_code = item_id
        customer_name = values[0]

        return customer_code, customer_name

    # =========================================================
    # 合計出荷情報（標準）
    # サンプルExcel完全一致：品目コード・枚数・個数
    # =========================================================
    def open_standard_summary(self, tree):
        customer_code, customer_name = self.get_selected_customer(tree)
        if customer_code is None:
            return

        try:
            conn = self.get_connection()
            df = pd.read_sql_query(
                self.create_summary_sql(by_order=False),
                conn,
                params=(customer_code, self.get_search_day(), self.get_from_time(), self.get_to_time()),
            )
            conn.close()
        except Exception as e:
            messagebox.showerror("エラー", f"合計出荷情報（標準）の表示に失敗しました。\n{e}")
            return

        if df.empty:
            messagebox.showinfo("確認", "合計データがありません。")
            return

        win = tk.Toplevel(self)
        win.title("合計出荷情報（標準）")
        win.geometry("950x540")

        lbl = tk.Label(win, text=str(customer_name), font=("Yu Gothic UI", 16, "bold"))
        lbl.pack(pady=8)

        columns = ("item_code", "sheet_count", "total_qty")
        tree2 = ttk.Treeview(win, columns=columns, show="headings", height=20)

        tree2.heading("item_code", text="品目コード")
        tree2.heading("sheet_count", text="枚数")
        tree2.heading("total_qty", text="個数")

        tree2.column("item_code", width=700)
        tree2.column("sheet_count", width=100, anchor=tk.E)
        tree2.column("total_qty", width=100, anchor=tk.E)

        for _, row in df.iterrows():
            tree2.insert("", "end", values=(row["item_code"], row["sheet_count"], row["total_qty"]))

        vsb = ttk.Scrollbar(win, orient="vertical", command=tree2.yview)
        tree2.configure(yscrollcommand=vsb.set)

        tree2.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        vsb.pack(side="right", fill="y", pady=10)

        btn_frame = tk.Frame(win)
        btn_frame.pack(pady=10)

        tk.Button(
            btn_frame, text="EXCEL出力", width=18,
            command=lambda: self.export_standard_summary(customer_code, customer_name)
        ).grid(row=0, column=0, padx=5)

        tk.Button(btn_frame, text="終了", width=18, fg="red", command=win.destroy).grid(row=0, column=1, padx=5)

    def export_standard_summary(self, customer_code, customer_name):
        try:
            conn = self.get_connection()
            df = pd.read_sql_query(
                self.create_summary_sql(by_order=False),
                conn,
                params=(customer_code, self.get_search_day(), self.get_from_time(), self.get_to_time()),
            )
            conn.close()

            if df.empty:
                messagebox.showinfo("確認", "出力対象データがありません。")
                return

            df = df[["item_code", "sheet_count", "total_qty"]]
            df = self.to_japanese_columns(df)

            output_path = export_dataframe_to_excel(df, f"{customer_name}_出荷合計数（標準）.xlsx")
            messagebox.showinfo("完了", f"EXCEL出力が完了しました。\n{output_path}")
        except Exception as e:
            messagebox.showerror("エラー", f"合計出荷情報（標準）のEXCEL出力に失敗しました。\n{e}")

    # =========================================================
    # 合計出荷情報（注番）
    # サンプルExcel完全一致：注番・品目コード・枚数・個数
    # =========================================================
    def open_order_summary(self, tree):
        customer_code, customer_name = self.get_selected_customer(tree)
        if customer_code is None:
            return

        try:
            conn = self.get_connection()
            df = pd.read_sql_query(
                self.create_summary_sql(by_order=True),
                conn,
                params=(customer_code, self.get_search_day(), self.get_from_time(), self.get_to_time()),
            )
            conn.close()
        except Exception as e:
            messagebox.showerror("エラー", f"合計出荷情報（注番）の表示に失敗しました。\n{e}")
            return

        if df.empty:
            messagebox.showinfo("確認", "合計データがありません。")
            return

        win = tk.Toplevel(self)
        win.title("合計出荷情報（注番）")
        win.geometry("950x540")

        lbl = tk.Label(win, text=str(customer_name), font=("Yu Gothic UI", 16, "bold"))
        lbl.pack(pady=8)

        columns = ("order_no", "item_code", "sheet_count", "total_qty")
        tree2 = ttk.Treeview(win, columns=columns, show="headings", height=20)

        tree2.heading("order_no", text="注番")
        tree2.heading("item_code", text="品目コード")
        tree2.heading("sheet_count", text="枚数")
        tree2.heading("total_qty", text="個数")

        tree2.column("order_no", width=200)
        tree2.column("item_code", width=550)
        tree2.column("sheet_count", width=100, anchor=tk.E)
        tree2.column("total_qty", width=100, anchor=tk.E)

        for _, row in df.iterrows():
            tree2.insert("", "end", values=(row["order_no"], row["item_code"], row["sheet_count"], row["total_qty"]))

        vsb = ttk.Scrollbar(win, orient="vertical", command=tree2.yview)
        tree2.configure(yscrollcommand=vsb.set)

        tree2.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        vsb.pack(side="right", fill="y", pady=10)

        btn_frame = tk.Frame(win)
        btn_frame.pack(pady=10)

        tk.Button(
            btn_frame, text="EXCEL出力", width=18,
            command=lambda: self.export_order_summary(customer_code, customer_name)
        ).grid(row=0, column=0, padx=5)

        tk.Button(btn_frame, text="終了", width=18, fg="red", command=win.destroy).grid(row=0, column=1, padx=5)

    def export_order_summary(self, customer_code, customer_name):
        try:
            conn = self.get_connection()
            df = pd.read_sql_query(
                self.create_summary_sql(by_order=True),
                conn,
                params=(customer_code, self.get_search_day(), self.get_from_time(), self.get_to_time()),
            )
            conn.close()

            if df.empty:
                messagebox.showinfo("確認", "出力対象データがありません。")
                return

            df = df[["order_no", "item_code", "sheet_count", "total_qty"]]
            df = df.rename(columns={"order_no": "注番"})
            df = self.to_japanese_columns(df)

            output_path = export_dataframe_to_excel(df, f"{customer_name}_出荷合計数（注番）.xlsx")
            messagebox.showinfo("完了", f"EXCEL出力が完了しました。\n{output_path}")
        except Exception as e:
            messagebox.showerror("エラー", f"合計出荷情報（注番）のEXCEL出力に失敗しました。\n{e}")

    # =========================================================
    # 明細出力
    # =========================================================
    def export_detail(self, tree):
        customer_code, customer_name = self.get_selected_customer(tree)
        if customer_code is None:
            return

        try:
            conn = self.get_connection()
            df = pd.read_sql_query(
                self.create_detail_sql(),
                conn,
                params=(customer_code, self.get_search_day(), self.get_from_time(), self.get_to_time()),
            )
            conn.close()

            if df.empty:
                messagebox.showinfo("確認", "明細データがありません。")
                return

            # 品目コード単位でGRPを付与（公開用サンプル）
            df["group_mark"] = ""
            current_mark = "■"

            if not df.empty:
                df.loc[0, "group_mark"] = current_mark
                for i in range(1, len(df)):
                    if df.loc[i, "item_code"] != df.loc[i - 1, "item_code"]:
                        current_mark = "□" if current_mark == "■" else "■"
                    df.loc[i, "group_mark"] = current_mark

            df = df[["group_mark", "item_code", "item_name", "shipping_qty", "order_no", "shelf_no", "shipment_no", "line_no"]]
            df = self.to_japanese_columns(df)

            output_path = export_dataframe_to_excel(df, f"{customer_name}_出荷明細.xlsx")
            messagebox.showinfo("完了", f"明細データのEXCEL出力が完了しました。\n{output_path}")

        except Exception as e:
            messagebox.showerror("エラー", f"明細データのEXCEL出力に失敗しました。\n{e}")


if __name__ == "__main__":
    app = ShippingSummaryApp()
    app.mainloop()
