"""Small, explicit XLSX display-format subset; never changes workbook values."""
from datetime import date, datetime, time
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext
import re


def raw_value(value):
    if value is None:
        return ''
    if isinstance(value, bool):
        return 'TRUE' if value else 'FALSE'
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    return str(value)


def format_value(value, number_format):
    """Return (display text, supported). Unknown patterns use the raw value."""
    raw = raw_value(value)
    if value is None or isinstance(value, (str, bool)):
        return raw, True
    pattern = (number_format or 'General').lower()
    if pattern in {'general', '@'}:
        return raw, True
    if isinstance(value, (date, datetime, time)):
        dates = {
            'yyyy-mm-dd': lambda d: f'{d.year:04}-{d.month:02}-{d.day:02}',
            'yyyy/mm/dd': lambda d: f'{d.year:04}/{d.month:02}/{d.day:02}',
            'yyyy/m/d': lambda d: f'{d.year:04}/{d.month}/{d.day}',
            'mm-dd-yy': lambda d: f'{d.month:02}-{d.day:02}-{d.year % 100:02}',
        }
        times = {
            'hh:mm:ss': lambda t: f'{t.hour:02}:{t.minute:02}:{t.second:02}',
            'h:mm:ss': lambda t: f'{t.hour}:{t.minute:02}:{t.second:02}',
            'hh:mm': lambda t: f'{t.hour:02}:{t.minute:02}',
            'h:mm': lambda t: f'{t.hour}:{t.minute:02}',
        }
        if isinstance(value, date) and pattern in dates:
            return dates[pattern](value), True
        if isinstance(value, (time, datetime)) and pattern in times:
            return times[pattern](value), True
        if isinstance(value, datetime):
            for d, date_text in dates.items():
                for t, time_text in times.items():
                    if pattern == d + ' ' + t:
                        return date_text(value) + ' ' + time_text(value), True
        return raw, False
    if type(value) not in (int, float):
        return raw, False
    match = re.fullmatch(r'(0|#,##0)(?:\.(0{1,6}))?(%)?', pattern)
    if match is None:
        return raw, False
    places = len(match[2] or '')
    try:
        with localcontext() as context:
            context.prec = 400
            number = Decimal(str(value))
            if not number.is_finite():
                return raw, False
            if match[3]:
                number *= 100
            number = number.quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)
            if not number:
                number = abs(number)
            text = format(number, f'{"," if match[1] == "#,##0" else ""}.{places}f')
            return text + ('%' if match[3] else ''), True
    except InvalidOperation:
        return raw, False


def cell_alignment(cell):
    horizontal = cell.alignment.horizontal
    if horizontal in {'left', 'center', 'right'}:
        return horizontal
    if isinstance(cell.value, bool) or cell.data_type == 'e':
        return 'center'
    if isinstance(cell.value, (int, float, date, time)):
        return 'right'
    return 'left'
