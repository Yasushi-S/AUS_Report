"""
AUS.xlsx → aus_data.db インポートスクリプト

使い方:
    python import_excel.py [--xlsx AUS.xlsx] [--db aus_data.db]
"""

import sqlite3
import argparse
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("ERROR: openpyxl が未インストールです。'pip install openpyxl' を実行してください。")
    sys.exit(1)

CREATE_TABLE_SQL = """
CREATE TABLE IF NOT EXISTS cases (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    seq_no           INTEGER NOT NULL,
    age              INTEGER NOT NULL,
    address          TEXT    NOT NULL,
    gestational_age  INTEGER NOT NULL,
    date             TEXT    NOT NULL,
    patient_id       TEXT    NOT NULL,
    note             TEXT    NOT NULL DEFAULT ''
)
"""

INSERT_SQL = """
INSERT INTO cases (seq_no, age, address, gestational_age, date, patient_id, note)
VALUES (?, ?, ?, ?, ?, ?, ?)
"""


def import_data(xlsx_path: str, db_path: str) -> None:
    xlsx = Path(xlsx_path)
    if not xlsx.exists():
        print(f"ERROR: {xlsx_path} が見つかりません。")
        sys.exit(1)

    print(f"読み込み中: {xlsx_path}")
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active

    rows = list(ws.iter_rows(min_row=2, values_only=True))
    print(f"Excelのデータ行数: {len(rows)} 件")

    conn = sqlite3.connect(db_path)
    conn.execute(CREATE_TABLE_SQL)
    conn.commit()

    inserted = 0
    skipped = 0
    errors = []

    for i, row in enumerate(rows, start=2):
        seq_no, age, address, gestational_age, date, patient_id, note = row

        # 必須フィールドが欠けている行はスキップ
        if seq_no is None or age is None or gestational_age is None or date is None or patient_id is None:
            errors.append(f"行{i}: 必須フィールドが空のためスキップ {row}")
            skipped += 1
            continue

        address = address if address is not None else ''
        note = note if note is not None else ''
        patient_id = str(patient_id).zfill(5)
        date = str(date).strip()

        try:
            conn.execute(INSERT_SQL, (int(seq_no), int(age), address, int(gestational_age), date, patient_id, note))
            inserted += 1
        except sqlite3.IntegrityError:
            errors.append(f"行{i}: 重複スキップ (patient_id={patient_id}, date={date})")
            skipped += 1

    conn.commit()
    conn.close()

    print(f"\n完了: {inserted} 件インポート、{skipped} 件スキップ")
    if errors:
        print("\n--- スキップ詳細 ---")
        for msg in errors:
            print(f"  {msg}")


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='AUS.xlsx を aus_data.db にインポートします')
    parser.add_argument('--xlsx', default='AUS.xlsx', help='Excelファイルパス (デフォルト: AUS.xlsx)')
    parser.add_argument('--db',   default='aus_data.db', help='SQLiteファイルパス (デフォルト: aus_data.db)')
    args = parser.parse_args()

    import_data(args.xlsx, args.db)
