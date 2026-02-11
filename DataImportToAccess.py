import re
import sys
from pathlib import Path
from typing import Dict, List, Tuple, Iterable

import pandas as pd

try:
    import pyodbc
except ImportError as exc:
    raise SystemExit(
        "pyodbc is required to write to Access via ODBC. "
        "Install it in your conda env, then re-run."
    ) from exc


ACCESS_DRIVER = r"{Microsoft Access Driver (*.mdb, *.accdb)}"

# === Configuration ===
MAX_COLNAME_LEN = 30
MAPPING_TABLE = "__column_map"

# - "short": sanitized + truncated to 30 chars + uniqueness suffixes
# - "letters": A, B, C, ... (Excel-style)
NAMING_MODE = "short"

# Insert tuning for stability
INSERT_BATCH_SIZE = 300
USE_FAST_EXECUTEMANY = False


def sanitize_identifier(name: str) -> str:
    s = str(name).strip()
    s = re.sub(r"[^\w]+", "_", s, flags=re.UNICODE)
    s = re.sub(r"_+", "_", s).strip("_")

    if not s:
        s = "col"
    if s[0].isdigit():
        s = f"c_{s}"
    return s


def excel_letters(n: int) -> str:
    n += 1
    out = []
    while n > 0:
        n, r = divmod(n - 1, 26)
        out.append(chr(ord("A") + r))
    return "".join(reversed(out))


def make_unique_truncated(names: List[str], max_len: int) -> Tuple[List[str], Dict[str, str]]:
    used = set()
    mapping: Dict[str, str] = {}
    shorts: List[str] = []

    for original in names:
        base = sanitize_identifier(original)
        base = base[:max_len] if len(base) > max_len else base
        if not base:
            base = "col"

        candidate = base
        if candidate not in used:
            used.add(candidate)
            shorts.append(candidate)
            mapping[str(original)] = candidate
            continue

        i = 2
        while True:
            suffix = f"_{i}"
            trimmed = base[: max_len - len(suffix)] if max_len > len(suffix) else ""
            candidate = f"{trimmed}{suffix}" if trimmed else f"c{suffix}"
            if candidate not in used:
                used.add(candidate)
                shorts.append(candidate)
                mapping[str(original)] = candidate
                break
            i += 1

    return shorts, mapping


def make_column_mapping(df: pd.DataFrame) -> Dict[str, str]:
    originals = [str(c) for c in df.columns]

    if NAMING_MODE == "letters":
        return {orig: excel_letters(i) for i, orig in enumerate(originals)}

    _, mapping = make_unique_truncated(originals, MAX_COLNAME_LEN)
    return mapping


def access_type_for_series(s: pd.Series) -> str:
    if pd.api.types.is_integer_dtype(s):
        return "LONG"
    if pd.api.types.is_float_dtype(s):
        return "DOUBLE"
    if pd.api.types.is_bool_dtype(s):
        return "YESNO"
    if pd.api.types.is_datetime64_any_dtype(s):
        return "DATETIME"
    max_len = int(s.dropna().astype(str).map(len).max() or 0)
    return "TEXT(255)" if max_len <= 255 else "LONGTEXT"


def table_exists(cursor, table: str) -> bool:
    return bool(cursor.tables(table=table, tableType="TABLE").fetchone())


def drop_table_if_exists(cursor, table: str) -> None:
    if table_exists(cursor, table):
        cursor.execute(f"DROP TABLE [{table}]")


def ensure_mapping_table(cursor) -> None:
    if table_exists(cursor, MAPPING_TABLE):
        return

    cursor.execute(
        f"""
        CREATE TABLE [{MAPPING_TABLE}] (
            [table_name] TEXT(64),
            [original_name] LONGTEXT,
            [short_name] TEXT(64)
        )
        """
    )


def write_mapping(cursor, target_table: str, mapping: Dict[str, str]) -> None:
    ensure_mapping_table(cursor)
    cursor.execute(f"DELETE FROM [{MAPPING_TABLE}] WHERE [table_name] = ?", (target_table,))
    rows = [(target_table, orig, short) for orig, short in mapping.items()]
    cursor.fast_executemany = False
    cursor.executemany(
        f"INSERT INTO [{MAPPING_TABLE}] ([table_name], [original_name], [short_name]) VALUES (?, ?, ?)",
        rows,
    )


def create_table(cursor, table: str, df: pd.DataFrame, mapping: Dict[str, str]) -> None:
    cols = []
    for orig in df.columns:
        access_col = mapping[str(orig)]
        col_type = access_type_for_series(df[orig])
        cols.append(f"[{access_col}] {col_type}")

    drop_table_if_exists(cursor, table)
    cursor.execute(f"CREATE TABLE [{table}] ({', '.join(cols)})")


def _batched(iterable: Iterable[tuple], batch_size: int) -> Iterable[List[tuple]]:
    batch = []
    for item in iterable:
        batch.append(item)
        if len(batch) >= batch_size:
            yield batch
            batch = []
    if batch:
        yield batch


def insert_rows(cursor, table: str, df: pd.DataFrame, mapping: Dict[str, str]) -> None:
    access_cols = [mapping[str(c)] for c in df.columns]
    placeholders = ", ".join(["?"] * len(access_cols))
    col_list = ", ".join(f"[{c}]" for c in access_cols)
    sql = f"INSERT INTO [{table}] ({col_list}) VALUES ({placeholders})"

    rows_iter = df.where(pd.notnull(df), None).itertuples(index=False, name=None)

    cursor.fast_executemany = USE_FAST_EXECUTEMANY
    for batch in _batched(rows_iter, INSERT_BATCH_SIZE):
        cursor.executemany(sql, batch)


def make_unique_table_name(desired: str, used: set) -> str:
    base = sanitize_identifier(desired)
    if not base:
        base = "Sheet"
    candidate = base
    if candidate not in used:
        used.add(candidate)
        return candidate

    i = 2
    while True:
        suffix = f"_{i}"
        candidate = f"{base}{suffix}"
        if candidate not in used:
            used.add(candidate)
            return candidate
        i += 1


def main():
    if len(sys.argv) < 3:
        print("Usage: python DataImportToAccess.py <input.xlsx> <output.accdb>")
        print("\nExample:")
        print("  python DataImportToAccess.py data.xlsx database.accdb")
        print("\nNote: The Access database file must exist before running this script.")
        print("      Create an empty .accdb file in Microsoft Access first.")
        sys.exit(1)

    excel_path = Path(sys.argv[1])
    access_path = Path(sys.argv[2])

    if not excel_path.exists():
        print(f"Error: Excel file not found: {excel_path}")
        sys.exit(1)

    if not access_path.exists():
        print(f"Error: Access database not found: {access_path}")
        print("\nPlease create an empty Access database first:")
        print("1. Open Microsoft Access")
        print("2. Create a new blank database")
        print(f"3. Save it as: {access_path}")
        print("4. Close Access and run this script again")
        sys.exit(1)

    print(f"Reading Excel file: {excel_path}")
    print(f"Target Access database: {access_path}")
    print()

    excel = pd.ExcelFile(excel_path)

    conn_str = (
        f"DRIVER={ACCESS_DRIVER};"
        f"DBQ={access_path};"
        "ExtendedAnsiSQL=1;"
    )

    used_table_names = set()

    try:
        with pyodbc.connect(conn_str, autocommit=False) as conn:
            cur = conn.cursor()

            for sheet_name in excel.sheet_names:
                print(f"Processing sheet: {sheet_name}")

                df = pd.read_excel(excel_path, sheet_name=sheet_name)

                table_name = make_unique_table_name(sheet_name, used_table_names)
                mapping = make_column_mapping(df)

                create_table(cur, table_name, df, mapping)
                write_mapping(cur, table_name, mapping)
                insert_rows(cur, table_name, df, mapping)

                conn.commit()
                print(f"✓ Sheet '{sheet_name}' exported -> Table [{table_name}] ({len(df)} rows)")

        print("\nAll sheets successfully exported!")
        print(f"Column mapping saved in table [{MAPPING_TABLE}] (per table).")

    except pyodbc.Error as e:
        print(f"\nError connecting to Access database: {e}")
        print("\nPossible causes:")
        print("- The database file is corrupted")
        print("- The database is already open in Access")
        print("- The Microsoft Access ODBC driver is not installed")
        sys.exit(1)


if __name__ == "__main__":
    main()
