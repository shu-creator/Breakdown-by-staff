#!/usr/bin/env python3
"""
Excelファイルを読み込んで、全体の「日程表作成担当者」と「1本目」の担当者集計を行うスクリプト
Usage:
    python aggregate_tantousha.py path/to/excel_file.xlsx
"""
import sys
import pandas as pd
from collections import Counter

def is_valid_sheet(sheet_name: str) -> bool:
    """
    シート名が「済」で終わり、先頭部分が数字のみかを判定
    例: "1済", "10済"
    """
    sheet_name = sheet_name.strip()
    if not sheet_name.endswith("済"):
        return False
    num_part = sheet_name[:-1].strip()
    return num_part.isdigit()


def aggregate(file_path: str):
    """
    指定したExcelファイルを読み込んで、担当者集計を実行
    Returns:
        total_counter: Counter 全体の担当者集計
        first_course_counter: Counter 「1本目」の担当者集計
    """
    xls = pd.ExcelFile(file_path)
    valid_sheets = [s for s in xls.sheet_names if is_valid_sheet(s)]

    total_counter = Counter()
    first_course_counter = Counter()

    for sheet in valid_sheets:
        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=0)
            # 必須列のチェック（J列(9)とI列(8)）
            if df.shape[1] < 10:
                continue

            # 全体集計 (J列 -> インデックス9)
            col_tantou = df.iloc[:, 9]
            for val in col_tantou.dropna():
                if isinstance(val, str):
                    v = val.strip()
                    if v and v != "日程表作成担当者":
                        total_counter[v] += 1

            # 注意事項以降から「1本目」を含む行の担当者を集計
            col_notes = df.iloc[:, 8]
            header_idx = None
            for idx, note in col_notes.items():
                if isinstance(note, str) and note.strip() == "注意事項":
                    header_idx = idx
                    break
            if header_idx is None:
                continue

            notes = col_notes.loc[header_idx + 1:]
            tantou2 = col_tantou.loc[header_idx + 1:]
            for note, tantou in zip(notes, tantou2):
                if isinstance(note, str) and "1本目" in note and isinstance(tantou, str):
                    v2 = tantou.strip()
                    if v2 and v2 != "日程表作成担当者":
                        first_course_counter[v2] += 1

        except Exception as e:
            print(f"Error processing sheet '{sheet}': {e}")

    return total_counter, first_course_counter


def main():
    if len(sys.argv) < 2:
        print("Usage: python aggregate_tantousha.py <excel_file>")
        sys.exit(1)

    file_path = sys.argv[1]
    total, first = aggregate(file_path)

    print("【全体の日程表作成担当者の集計】")
    for name, count in total.most_common():
        print(f"{name}: {count}")

    print("\n【『1本目』のコースにおける担当者の集計】")
    for name, count in first.most_common():
        print(f"{name}: {count}")

if __name__ == '__main__':
    main()
