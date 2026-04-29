import sqlite3
import pandas as pd

from file_utils import get_sample_file_path


SHIPMENT_FILE = "shipment_detail.xlsx"
SHELF_MASTER_FILE = "shelf_master.xlsx"


def import_shipment_data(db_path: str) -> int:
    shipment_path = get_sample_file_path(SHIPMENT_FILE)

    if not shipment_path.exists():
        raise FileNotFoundError(f"Sample file not found: {shipment_path}")

    df = pd.read_excel(shipment_path)

    rename_map = {
        "納品先コード": "customer_code",
        "納品先名称": "customer_name",
        "品目コード": "item_code",
        "規格": "item_name",
        "品目略称": "item_name",
        "出荷数量": "shipping_qty",
        "相手先注文番号": "order_no",
        "出荷伝票番号": "shipment_no",
        "行番号": "line_no",
        "作成日付": "created_date",
        "作成時間": "created_time",
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    required_defaults = {
        "customer_code": 0,
        "customer_name": "Sample Customer",
        "item_code": "ITEM-000",
        "item_name": "Sample Item",
        "shipping_qty": 0,
        "order_no": "ORDER-000",
        "shipment_no": "SHIP-000",
        "line_no": 1,
        "created_date": 19000101,
        "created_time": 0,
    }

    for col, default_value in required_defaults.items():
        if col not in df.columns:
            df[col] = default_value

    conn = sqlite3.connect(db_path)

    try:
        conn.execute("BEGIN TRANSACTION")

        conn.execute("DROP TABLE IF EXISTS shipment_detail_temp")
        df.to_sql("shipment_detail_temp", conn, if_exists="replace", index=False)

        existing_cols = [row[1] for row in conn.execute("PRAGMA table_info(shipment_detail_temp)").fetchall()]
        if "shelf_no" not in existing_cols:
            conn.execute("ALTER TABLE shipment_detail_temp ADD COLUMN shelf_no TEXT")

        shelf_path = get_sample_file_path(SHELF_MASTER_FILE)
        if shelf_path.exists():
            shelf_df = pd.read_excel(shelf_path)

            shelf_rename_map = {
                "品目コード": "item_code",
                "棚番": "shelf_no",
            }
            shelf_df = shelf_df.rename(columns={k: v for k, v in shelf_rename_map.items() if k in shelf_df.columns})

            if "item_code" in shelf_df.columns and "shelf_no" in shelf_df.columns:
                conn.execute("DROP TABLE IF EXISTS shelf_master_temp")
                shelf_df.to_sql("shelf_master_temp", conn, if_exists="replace", index=False)

                conn.execute("""
                    UPDATE shipment_detail_temp
                    SET shelf_no = (
                        SELECT shelf_master_temp.shelf_no
                        FROM shelf_master_temp
                        WHERE shipment_detail_temp.item_code = shelf_master_temp.item_code
                    )
                    WHERE EXISTS (
                        SELECT 1
                        FROM shelf_master_temp
                        WHERE shipment_detail_temp.item_code = shelf_master_temp.item_code
                    )
                """)

        conn.commit()
        return 0

    except Exception:
        conn.rollback()
        raise

    finally:
        conn.close()
