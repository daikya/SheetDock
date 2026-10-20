"""CSV view heights (pixels); None keeps the widget's normal height."""
import math

MIN_HEIGHT, MAX_HEIGHT = 5, 2000


def fit_heights(heights, count):
    values = heights if isinstance(heights, list) else []
    result = []
    for i in range(count):
        value = values[i] if i < len(values) else None
        valid = type(value) in (int, float) and math.isfinite(value) and MIN_HEIGHT <= value <= MAX_HEIGHT
        result.append(round(value) if valid else None)
    return result


def accumulate_row_resize(pending, event):
    for row, sizes in event.get('resized', {}).get('rows', {}).items():
        old = pending[row][0] if row in pending else sizes['old_size']
        pending[row] = (old, sizes['new_size'])
