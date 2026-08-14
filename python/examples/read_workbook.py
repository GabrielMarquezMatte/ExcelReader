"""Print the first rows of a workbook.  Usage: python read_workbook.py <path> [max_rows]"""

import sys

from excelreader import open_workbook


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2

    path = sys.argv[1]
    max_rows = int(sys.argv[2]) if len(sys.argv) > 2 else 10

    with open_workbook(path) as workbook:
        print(f"sheets={workbook.sheet_count} current={workbook.sheet_name!r} date1904={workbook.is_date1904}")
        for index, row in enumerate(workbook.rows()):
            if index >= max_rows:
                break
            print(index, [(cell.column, cell.type.name, cell.value) for cell in row])

    return 0


if __name__ == "__main__":
    sys.exit(main())
