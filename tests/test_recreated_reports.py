"""Default GUI/batch path must rebuild reports even from an old or absent layout."""
from pathlib import Path
from datetime import datetime
from copy import copy
import hashlib, subprocess, uuid, json
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side

ROOT = Path(__file__).resolve().parents[1]
B = ROOT / 'tests/.runs' / uuid.uuid4().hex
B.mkdir(parents=True)
REPORTS = {'СВОД', 'Стат. по типам', 'Стат. по зам.', 'Итоги работы'}
MAIN = 'Все загруженные файлы'

def run(output, source):
    subprocess.run([str(ROOT / 'Свод_проверки.exe'), '--batch', str(output),
                    '2026-10-04', MAIN, str(source)], check=True, timeout=90)

def validate(path):
    w = load_workbook(path)
    assert REPORTS <= set(w.sheetnames), w.sheetnames
    assert 'Свод' not in w.sheetnames
    assert w['СВОД']['D4'].value == 'Всего'
    assert w['СВОД']['H4'].value == 'Проверено'
    assert w['СВОД']['L4'].value == 'Осталось проверить'
    assert w['СВОД']['P4'].value == 'Проверено за сутки'
    assert w['Стат. по зам.']['A1'].value == 'СТАТИСТИКА ПО ЗАМЕЧАНИЯМ'
    assert w['Итоги работы']['A1'].value == 'СВОДНАЯ ТАБЛИЦА ПО ИТОГАМ РАБОТЫ ПРОВЕРЯЮЩИХ'
    assert w['Итоги работы']['C5'].value == 'коробов'
    assert w['Итоги работы']['D5'].value == 'томов'
    for tab in REPORTS:
        s = w[tab]
        assert s['A4'].value == 'Этап' and s['A4'].border.top.style == 'thick'
        assert s.print_area and s.page_setup.orientation == 'landscape'
        assert s['A1'].font.name == 'Times New Roman'
        for row in s:
            for c in row:
                if c.data_type == 'f':
                    assert '#REF!' not in c.value
                    assert 'IF(' in c.value or 'COUNT' in c.value or 'SUM' in c.value
    stats = w['Стат. по зам.']
    assert [stats[a].font.color.rgb for a in ['C6', 'E6', 'I6', 'M6']] == ['FF996600', 'FF00B050', 'FFFF0000', 'FF0070C0']
    assert w['СВОД']['D7'].border.right.style == 'dashed' and w['СВОД']['E7'].border.left.style == 'dashed'
    assert w[MAIN]['M5'].value == 'Номер тома' and w[MAIN]['M7'].number_format == '@'
    assert sum(bool(view.tabSelected) for s in w for view in s.views.sheetView) == 1
    assert any(stats.row_dimensions[r].outlineLevel == 1 for r in stats.row_dimensions)
    return w

# Reproduce the user's old workbook with a horizontal report and no required captions.
source = B / 'old-layout.xlsx'
w = load_workbook(ROOT / 'tests/fixtures/Анна.xlsx')
for s in list(w):
    if s.title.upper() == 'СВОД': w.remove(s)
s = w.create_sheet('Свод')
s['A1'] = 'Старое оформление'
s['A7'] = 'Получено коробов'; s['A8'] = 'Томов по акту'
s['B5'] = datetime(2026,10,2); s['C5'] = datetime(2026,10,4)
s['B7'] = "='Справка по томам'!B1"; s['B8'] = "='Справка по томам'!B2"
s['A17'] = 'Примечание: исходные сведения сохраняются'
if 'Справка по томам' not in w.sheetnames: w.create_sheet('Справка по томам')
w['Справка по томам']['B1'] = 89; w['Справка по томам']['B2'] = 1975
old_main = {c.coordinate:c.value for row in w[MAIN] for c in row if c.value is not None}
w.save(source); digest = hashlib.sha256(source.read_bytes()).hexdigest()
target = B / 'recreated.xlsx'; run(target, source); w = validate(target)
assert hashlib.sha256(source.read_bytes()).hexdigest() == digest
assert w['СВОД']['A1'].value != 'Старое оформление'
assert any('исходные сведения' in str(c.value) for row in w['СВОД'] for c in row)
assert w['Справка по томам']['B1'].value == 89 and w['Справка по томам']['B2'].value == 1975

# Existing report tabs must also be regenerated, including stale formatting and formulas.
for tab in REPORTS:
    w[tab]['A1'] = 'ПОВРЕЖДЕНО'
    w[tab]['A4'].border = Border()
    w[tab]['C6'].font = Font(color='FF00FF')
w.save(target); run(target, source); w = validate(target)
assert all(w[s]['A1'].value != 'ПОВРЕЖДЕНО' for s in REPORTS)
assert w['Справка по томам']['B2'].value == 1975
no_layout = B / 'no-layout.xlsx'
w = load_workbook(source); w.remove(w['Свод']); w.save(no_layout)
fresh = B / 'new.xlsx'; run(fresh, no_layout); validate(fresh)
# A custom incomplete template previously entered the silent "keep old report" branch.
w = load_workbook(source); s = w['Свод']; s['A4']='Этап'; s['B4']='Дата / проверяющий'; w.save(source)
custom = B / 'custom.xlsx'; run(custom, source); validate(custom)
print('Report recreation passed: old, absent and custom layouts; stale split tabs; repeat merge; styles, formulas, selection and source totals')
print(target)
