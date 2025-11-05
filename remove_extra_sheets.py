import argparse
from pathlib import Path

import openpyxl


def remove_sheets(workbook_path: Path) -> None:
    """Delete all worksheets except the first one in the workbook."""
    # 加载指定路径的工作簿
    wb = openpyxl.load_workbook(workbook_path)
    sheet_names = wb.sheetnames

    # Preserve the first sheet only
    # 从第二个工作表开始逐个删除
    for name in sheet_names[1:]:
        del wb[name]

    # 保存修改并关闭工作簿
    wb.save(workbook_path)
    wb.close()


def main() -> None:
    # 解析命令行参数，获取工作簿路径
    parser = argparse.ArgumentParser(
        description="Delete all worksheets except the first one in the given workbook."
    )
    parser.add_argument(
        "workbook",
        type=Path,
        help="Path to the target workbook (xlsx/xlsm).",
    )
    args = parser.parse_args()

    remove_sheets(args.workbook)
    print(f"Removed extra worksheets in {args.workbook}")


if __name__ == "__main__":
    main()
