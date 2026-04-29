import sqlite3
import pandas as pd

from file_utils import get_sample_file_path


SALES_FILE = "sales_data.xlsx"


def import_sales_data(db_path: str) -> int:
    sales_path = get_sample_file_path(SALES_FILE)

    if not sales_path.exists():
        raise FileNotFoundError(f"Sample file not found: {sales_path}")

    df = pd.read_excel(sales_path)

    rename_map = {
        "タイムスタンプ(日付)": "timestamp_date",
        "タイムスタンプ(時間)": "timestamp_time",
        "作成ﾀｲﾑｽﾀﾝﾌﾟ(日付)": "created_date",
        "作成ﾀｲﾑｽﾀﾝﾌﾟ(時間)": "created_time",
        "売上伝票番号": "shipment_no",
        "行番号": "line_no",
        "規格": "item_name",
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    required_defaults = {
        "created_date": 19000101,
        "created_time": 0,
        "shipment_no": "SHIP-000",
        "line_no": 1,
        "item_name": "Sample Item",
    }

    for col, default_value in required_defaults.items():
        if col not in df.columns:
            df[col] = default_value

    conn = sqlite3.connect(db_path)

    try:
        conn.execute("BEGIN TRANSACTION")

        conn.execute("DROP TABLE IF EXISTS sales_temp")
        df.to_sql("sales_temp", conn, if_exists="replace", index=False)

        table_check = conn.execute("""
            SELECT name
            FROM sqlite_master
            WHERE type='table' AND name='shipment_detail_temp'
        """).fetchone()

        if table_check:
            existing_cols = [row[1] for row in conn.execute("PRAGMA table_info(shipment_detail_temp)").fetchall()]

            if "created_date" not in existing_cols:
                conn.execute("ALTER TABLE shipment_detail_temp ADD COLUMN created_date NUMERIC")

            if "created_time" not in existing_cols:
                conn.execute("ALTER TABLE shipment_detail_temp ADD COLUMN created_time NUMERIC")

            if "item_name" not in existing_cols:
                conn.execute("ALTER TABLE shipment_detail_temp ADD COLUMN item_name TEXT")

            conn.execute("""
                UPDATE shipment_detail_temp
                SET
                    created_date = (
                        SELECT sales_temp.created_date
                        FROM sales_temp
                        WHERE sales_temp.shipment_no = shipment_detail_temp.shipment_no
                          AND sales_temp.line_no = shipment_detail_temp.line_no
                    ),
                    created_time = (
                        SELECT sales_temp.created_time
                        FROM sales_temp
                        WHERE sales_temp.shipment_no = shipment_detail_temp.shipment_no
                          AND sales_temp.line_no = shipment_detail_temp.line_no
                    ),
                    item_name = (
                        SELECT sales_temp.item_name
                        FROM sales_temp
                        WHERE sales_temp.shipment_no = shipment_detail_temp.shipment_no
                          AND sales_temp.line_no = shipment_detail_temp.line_no
                    )
                WHERE EXISTS (
                    SELECT 1
                    FROM sales_temp
                    WHERE sales_temp.shipment_no = shipment_detail_temp.shipment_no
                      AND sales_temp.line_no = shipment_detail_temp.line_no
                )
            """)

        conn.commit()
        return 0

    except Exception:
        conn.rollback()
        raise

    finally:
        conn.close()
