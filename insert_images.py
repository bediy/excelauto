import argparse
import math
import re
from collections import defaultdict
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Tuple

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from PIL import Image as PILImage

# 目标单元格
TARGET_CELL = "B19"
# 合并区域包含的列与行范围，便于后续计算整体尺寸
# 像素到 EMU 的转换系数
TARGET_COLUMNS = ("B", "C", "D", "E")
TARGET_ROWS = (19, 20, 21)
EMU_PER_PIXEL = 9525
# Excel 默认列宽与行高
DEFAULT_COLUMN_WIDTH = 8.38
DEFAULT_ROW_HEIGHT = 15.0


def column_width_to_pixels(width: float | None) -> float:
    """将列宽单位转换为像素。"""
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


def merged_rows_height_in_pixels(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    rows: Tuple[int, ...],
) -> int:
    """统计合并区域内各行的像素高度总和，确保图片高度匹配。"""
    total_height = 0.0
    for index in rows:
        row_dim = ws.row_dimensions.get(index)
        height = row_height_to_pixels(row_dim.height if row_dim and row_dim.height is not None else None)
        total_height += height
    return int(total_height)


def load_images_by_person(images_dir: Path) -> Dict[str, List[Tuple[int, Path]]]:
    """按姓名整理图片路径，并根据文件名末尾数字排序。"""
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
    """使用 Pillow 调整图片尺寸，打破原始纵横比。"""
    with PILImage.open(image_path) as image:
        resized = image.resize((width, height))
    return resized


def image_to_stream(image: PILImage.Image, fmt: str) -> BytesIO:
    """将 PIL Image 转换为内存二进制流。"""
    stream = BytesIO()
    image.save(stream, format=fmt)
    stream.seek(0)
    return stream

def resolve_horizontal_anchor(
    columns: Tuple[str, ...],
    column_widths_px: List[int],
    offset_px: int,
) -> Tuple[int, int]:
    """Compute the column index and intra-column offset for a given pixel offset."""
    remaining = offset_px
    for letter, width_px in zip(columns, column_widths_px):
        col_idx = column_index_from_string(letter) - 1
        if remaining < width_px:
            return col_idx, remaining
        remaining -= width_px
    last_letter = columns[-1]
    last_idx = column_index_from_string(last_letter) - 1
    last_width = column_widths_px[-1] if column_widths_px else 0
    return last_idx, max(min(remaining, last_width), 0)


def insert_images_to_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    image_paths: List[Path],
    cell_address: str,
    columns: Tuple[str, ...],
    rows: Tuple[int, ...],
) -> None:
    """将两张人像横向铺满合并区域，保证不重叠且贴合边界。"""
    if not image_paths:
        return

    _, row_number = coordinate_from_string(cell_address)
    row_idx = row_number - 1

    column_widths_px: List[int] = []
    total_width_px = 0
    for letter in columns:
        col_dim = ws.column_dimensions.get(letter)
        width_px = int(round(column_width_to_pixels(col_dim.width if col_dim and col_dim.width is not None else None)))
        column_widths_px.append(width_px)
        total_width_px += width_px

    total_height_px = merged_rows_height_in_pixels(ws, rows)
    if total_width_px <= 0 or total_height_px <= 0:
        return

    print(
        f"[DEBUG] target cell {cell_address} merged area size: "
        f"{total_width_px}px x {total_height_px}px; columns={columns}; widths={column_widths_px}; rows={rows}"
    )

    max_images = min(len(image_paths), 2)
    if max_images == 0:
        return

    base_width = total_width_px / max_images
    width_allocations: List[int] = []
    previous_right = 0
    for idx in range(max_images):
        right = round(base_width * (idx + 1))
        width_allocations.append(max(right - previous_right, 1))
        previous_right = right

    offset_px = 0
    for idx, path in enumerate(image_paths[:max_images]):
        image_width_px = width_allocations[idx]
        print(
            f"[DEBUG] processing image {idx + 1}/{max_images} '{path.name}': "
            f"allocated size {image_width_px}px x {total_height_px}px, "
            f"offset {offset_px}px"
        )
        resized = resize_image(path, image_width_px, total_height_px)
        fmt = (path.suffix or ".png").replace(".", "").upper()
        if fmt == "JPG":
            fmt = "JPEG"
        stream = image_to_stream(resized, fmt)
        resized.close()

        img = XLImage(stream)
        img.width = image_width_px
        img.height = total_height_px
        anchor_col_idx, anchor_col_offset_px = resolve_horizontal_anchor(
            columns,
            column_widths_px,
            offset_px,
        )
        anchor_marker = AnchorMarker(
            col=anchor_col_idx,
            colOff=int(anchor_col_offset_px * EMU_PER_PIXEL),
            row=row_idx,
            rowOff=0,
        )
        anchor_ext = XDRPositiveSize2D(
            cx=int(image_width_px * EMU_PER_PIXEL),
            cy=int(total_height_px * EMU_PER_PIXEL),
        )
        img.anchor = OneCellAnchor(_from=anchor_marker, ext=anchor_ext)
        print(
            "[DEBUG] anchor info: "
            f"col={anchor_marker.col}, row={anchor_marker.row}, "
            f"colOff={anchor_marker.colOff}, rowOff={anchor_marker.rowOff}, "
            f"ext=({anchor_ext.cx}, {anchor_ext.cy})"
        )

        ws.add_image(img)

        offset_px += image_width_px


def process_workbook(workbook_path: Path, images_dir: Path) -> None:
    """根据图片目录内容，把照片插入到对应姓名的工作表。"""
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
    parser = argparse.ArgumentParser(description="将人员照片批量插入到对应姓名的工作表中。")
    parser.add_argument("workbook", type=Path, help="目标工作簿路径（xlsx/xlsm）。")
    parser.add_argument("images_dir", type=Path, help="存放人员照片的文件夹路径。")
    args = parser.parse_args()

    if not args.workbook.exists():
        raise FileNotFoundError(f"未找到工作簿：{args.workbook}")
    if not args.images_dir.exists() or not args.images_dir.is_dir():
        raise NotADirectoryError(f"照片目录不存在或不是文件夹：{args.images_dir}")

    process_workbook(args.workbook, args.images_dir)
    print("图片插入完成。")


if __name__ == "__main__":
    main()
