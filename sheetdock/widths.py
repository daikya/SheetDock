"""列幅の検証と列数への調整。"""

import math

DEFAULT_WIDTH = 120


def fit_widths(widths, count):
    values = widths if isinstance(widths, list) else []
    result = []
    for i in range(count):
        value = values[i] if i < len(values) else DEFAULT_WIDTH
        valid = type(value) in (int, float) and math.isfinite(value) and 5 <= value <= 5000
        result.append(round(value) if valid else DEFAULT_WIDTH)
    return result
