"""Fresh typed/style/protection workbooks for production integration tests."""
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation


def make_samples(folder):
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    names = ('basic', 'formula', 'merged', 'readonly', 'mixed', 'hidden', 'error')
    paths = [folder / f'sort_{name}.xlsx' for name in names]
    if any(path.exists() for path in paths):
        raise FileExistsError('同名サンプルが存在します。別フォルダーを指定してください')
    for name, path in zip(names, paths):
        book = Workbook()
        sheet = book.active
        sheet.title = 'Data'
        sheet.append(['コード', '数量', '日付', '有効', '区分', 'メモ'])
        for row, (code, number, place) in enumerate([
                ('00123', 10, '東京'), ('00002', 2, '大阪'), ('00003', 2, '東京'),
                ('00004', None, '京都'), ('00005', -1, '東京')], 2):
            sheet.append([code, number, datetime(2026, 10, row), row % 2 == 0, place, '改行\n引用符"と,カンマ'])
            sheet.row_dimensions[row].height = 18 + row * 3
            sheet.cell(row, 2).number_format = '#,##0.00'
            sheet.cell(row, 3).number_format = 'yyyy-mm-dd'
            sheet.cell(row, 1).font = Font(bold=row % 2 == 0, color='FF115599')
            sheet.cell(row, 5).fill = PatternFill('solid', fgColor='FFDDFFDD' if row % 2 else 'FFFFDDDD')
            sheet.cell(row, 6).alignment = Alignment(wrap_text=True)
        sheet.column_dimensions['A'].width = 18
        sheet.column_dimensions['F'].width = 30
        book.create_sheet('Other')['A1'] = '他シート不変'
        if name == 'formula':
            book['Other']['A2'] = '=Data!B2'
        elif name == 'merged':
            sheet.merge_cells('E2:F2')
        elif name == 'readonly':
            rule = DataValidation(type='list', formula1='"東京,大阪"')
            sheet.add_data_validation(rule)
            rule.add('E2')
        elif name == 'mixed':
            sheet['B3'] = '2'
        elif name == 'hidden':
            sheet.row_dimensions[3].hidden = True
        elif name == 'error':
            sheet['F3'] = '#N/A'
        try:
            with path.open('xb') as stream:
                book.save(stream)
        finally:
            book.close()
