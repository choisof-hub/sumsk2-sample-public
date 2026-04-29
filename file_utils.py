from pathlib import Path
import pandas as pd


DATA_DIR = Path("data")
OUTPUT_DIR = Path("output")
DB_FILE = DATA_DIR / "sample.db"


def ensure_data_dir():
    DATA_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)


def get_db_path() -> str:
    ensure_data_dir()
    return str(DB_FILE)


def get_sample_file_path(file_name: str) -> Path:
    ensure_data_dir()
    return DATA_DIR / file_name


def export_dataframe_to_excel(df: pd.DataFrame, file_name: str) -> str:
    ensure_data_dir()
    safe_name = file_name.replace("/", "_").replace("\\", "_")
    output_path = OUTPUT_DIR / safe_name
    df.to_excel(output_path, index=False)
    return str(output_path)
