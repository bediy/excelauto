import argparse
import math
import re
from collections import defaultdict
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Tuple

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string
from openpyxl.utils.units import pixels_to_EMU
from PIL import Image as PILImage

# 目标单元格
TARGET_CELL = "B19"
# 合并区域的列与行范围，决定图片的目标尺寸
TARGET_COLUMNS = ("B", "C", "D", "E")
TARGET_ROWS = (19, 20, 21)
# Excel 的默认列宽与行高
DEFAULT_COLUMN_WIDTH = 8.38
DEFAULT_ROW_HEIGHT = 15.0


def column_width_to_pixels(width: float | None) -> float:
    """将列宽（字符单位）转换为像素。"""
    if not width or width <= 0:
        width = DEFAULT_COLUMN_WIDTH
    if width < 1:
        pixels = math.floor(width * 12 + 0.5)
    else:
        pixels = math.floor(width * 7 + 5)
    return float(pixels)


def row_height_to_pixels(height: float | None) -> float:
    """将行高（磅）转换为像素。"""
    if not height or height <= 0:
        height = DEFAULT_ROW_HEIGHT
    return float(height * 4 / 3)


def merged_columns_width_in_pixels(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    columns: Tuple[str, ...],
) -> int:
    """统计合并区域内各列宽度的像素总和，作为图片宽度参考。"""
    total_width = 0.0
    for letter in columns:
        col_dim = ws.column_dimensions.get(letter)
        width = column_width_to_pixels(col_dim.width if col_dim and col_dim.width is not None else None)
        total_width += width
    return int(total_width)


def merged_rows_height_in_pixels(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    rows: Tuple[int, ...],
) -> int:
    """统计合并区域内各行高度的像素总和，确保图片高度匹配。"""
    total_height = 0.0
    for index in rows:
        row_dim = ws.row_dimensions.get(index)
        height = row_height_to_pixels(row_dim.height if row_dim and row_dim.height is not None else None)
        total_height += height
    return int(total_height)


def load_images_by_person(images_dir: Path) -> Dict[str, List[Tuple[int, Path]]]:
    """按姓名归集人员图片，并按照文件名中末尾的数字序号排序。"""
    image_map: Dict[str, List[Tuple[int, Path]]] = defaultdict(list)
    pattern = re.compile(r"^(?P<name>.+?)(?P<index>\d+)?$", re.UNICODE)

    for path in images_dir.iterdir():
        if not path.is_file():
            continue
        if path.suffix.lower() not in {".jpg", ".jpeg", ".png", ".bmp"}:
            continue
        match = pattern.match(path.stem)
        if not match:
            continue
        name = match.group("name").strip()
        if not name:
            continue
        index_str = match.group("index")
        order = int(index_str) if index_str and index_str.isdigit() else 0
        image_map[name].append((order, path))

    for name in image_map:
        image_map[name].sort(key=lambda item: item[0])
    return image_map


def resize_image(image_path: Path, width: int, height: int) -> PILImage.Image:
    """使用 Pillow 调整图片尺寸，保持简单缩放。"""
    with PILImage.open(image_path) as image:
        resized = image.resize((width, height))
    return resized


def image_to_stream(image: PILImage.Image, fmt: str) -> BytesIO:
    """将 PIL Image 转换为内存流，便于 openpyxl 消耗。"""
    stream = BytesIO()
    image.save(stream, format=fmt)
    stream.seek(0)
    return stream


def insert_images_to_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    image_paths: List[Path],
    cell_address: str,
    columns: Tuple[str, ...],
    rows: Tuple[int, ...],
) -> None:
    """将图片插入模板的合并单元格中，并横向平铺整个区域。"""
    if not image_paths:
        return

    column_letter, row_number = coordinate_from_string(cell_address)
    column_letter = column_letter.upper()
    col_idx = column_index_from_string(column_letter) - 1
    row_idx = row_number - 1

    total_width_px = merged_columns_width_in_pixels(ws, columns)
    total_height_px = merged_rows_height_in_pixels(ws, rows)
    if total_width_px <= 0 or total_height_px <= 0:
        return

    column_widths_px: List[int] = []
    for letter in columns:
        col_dim = ws.column_dimensions.get(letter)
        column_widths_px.append(int(column_width_to_pixels(col_dim.width if col_dim and col_dim.width is not None else None)))

    max_images = min(len(image_paths), 2)
    if max_images == 0:
        return

    base_width = total_width_px / max_images
    width_allocations: List[int] = []
    cumulative_width = 0.0
    previous_right = 0
    for idx in range(max_images):
        cumulative_width += base_width
        right = round(cumulative_width)
        width_allocations.append(max(right - previous_right, 1))
        previous_right = right

    offset_px = 0
    for idx, path in enumerate(image_paths[:max_images]):
        image_width_px = width_allocations[idx]
        resized = resize_image(path, image_width_px, total_height_px)
        fmt = (path.suffix or ".png").replace(".", "").upper()
        if fmt == "JPG":
            fmt = "JPEG"
        stream = image_to_stream(resized, fmt)
        resized.close()

        img = XLImage(stream)
        img.width = image_width_px
        img.height = total_height_px

        # 计算在合并区域内的列偏移，让图片依次横向铺满
        remaining_offset = offset_px
        col_offset = 0
        for width_px in column_widths_px:
            if remaining_offset < width_px:
                break
            remaining_offset -= width_px
            col_offset += 1
        if col_offset >= len(column_widths_px):
            col_offset = len(column_widths_px) - 1
            remaining_offset = max(column_widths_px[-1] - 1, 0)

        img.anchor = OneCellAnchor(
            _from=AnchorMarker(
                col=col_idx + col_offset,
                colOff=pixels_to_EMU(remaining_offset),
                row=row_idx,
                rowOff=0,
            ),
            ext=XDRPositiveSize2D(
                pixels_to_EMU(image_width_px),
                pixels_to_EMU(total_height_px),
            ),
        )

        ws.add_image(img)
        offset_px += image_width_px


def process_workbook(workbook_path: Path, images_dir: Path) -> None:
    """根据图片目录数据，将照片插入到对应人员的工作表。"""
    image_map = load_images_by_person(images_dir)
    wb = openpyxl.load_workbook(workbook_path)

    for name, ordered_paths in image_map.items():
        if name not in wb.sheetnames:
            continue
        if len(ordered_paths) < 2:
            continue
        ws = wb[name]
        first_two = [path for _, path in ordered_paths[:2]]
        insert_images_to_sheet(ws, first_two, TARGET_CELL, TARGET_COLUMNS, TARGET_ROWS)

    wb.save(workbook_path)
    wb.close()


def main() -> None:
    """命令行入口，负责参数解析与流程串联。"""
    parser = argparse.ArgumentParser(description="将人员照片批量插入到对应的 Excel 工作表。")
    parser.add_argument("workbook", type=Path, help="目标工作簿路径（xlsx/xlsm）。")
    parser.add_argument("images_dir", type=Path, help="人员照片所在目录。")
    args = parser.parse_args()

    if not args.workbook.exists():
        raise FileNotFoundError(f"未找到目标工作簿 {args.workbook}")
    if not args.images_dir.exists() or not args.images_dir.is_dir():
        raise NotADirectoryError(f"图片目录不存在或不是文件夹：{args.images_dir}")

    process_workbook(args.workbook, args.images_dir)
    print("图片插入完成。")


if __name__ == "__main__":
    main()
