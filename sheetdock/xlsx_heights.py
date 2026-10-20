"""Worksheet row heights in points, converted to Tk pixels without JSON."""
from copy import copy
import math
from openpyxl.worksheet.dimensions import RowDimension

MAX_POINTS = 409.5


def height_to_pixels(points, scale=4/3):
    if not math.isfinite(points) or points < 0 or not math.isfinite(scale) or scale <= 0:
        raise ValueError('行の高さまたは表示倍率が不正です')
    return round(points * scale)


def pixels_to_height(pixels, scale=4/3):
    if type(pixels) not in (int, float) or not math.isfinite(pixels):
        raise ValueError('行の高さが不正です')
    maximum = height_to_pixels(MAX_POINTS, scale)
    if not 5 <= round(pixels) <= maximum:
        raise ValueError('行の高さは5ピクセル以上、Excelの409.5ポイント以下で指定してください')
    return min(MAX_POINTS, round(pixels) / scale)


def display_heights(sheet, count, scale=4/3, minimum=5):
    maximum = height_to_pixels(MAX_POINTS, scale)
    values = []
    for row in range(1, count + 1):
        dimension = sheet.row_dimensions.get(row)
        points = dimension.height if dimension is not None else None
        if points is None:
            points = sheet.sheet_format.defaultRowHeight
            if points is None:
                points = 15
        values.append(max(minimum, min(maximum, height_to_pixels(points, scale))))
    return values


def dimension_state(sheet):
    return [(row, copy(dimension)) for row, dimension in sheet.row_dimensions.items()]


def restore_dimensions(sheet, state):
    sheet.row_dimensions.clear()
    for row, dimension in state:
        restored = copy(dimension)
        restored.parent = sheet
        sheet.row_dimensions[row] = restored


def set_height(sheet, row, points):
    original = sheet.row_dimensions.get(row)
    if (original is not None and (original.hidden or original.height == 0)
            or (sheet.sheet_format.zeroHeight or sheet.sheet_format.defaultRowHeight == 0)
            and (original is None or original.height is None)):
        raise ValueError('非表示または高さ0の行は変更できません')
    dimension = copy(original) if original is not None else RowDimension(sheet, index=row)
    dimension.parent = sheet
    dimension.height = points
    sheet.row_dimensions[row] = dimension
