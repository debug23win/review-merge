"""Working documentation (--rd): the RD check sheet «Проверка комплектов» merged from reviewers' copies; artificial data only."""
import json, subprocess, uuid
from datetime import datetime
from pathlib import Path
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
EXE = ROOT / 'Свод_проверки.exe'
MAIN = 'Проверка комплектов'
B = ROOT / 'tests/.runs' / uuid.uuid4().hex
B.mkdir(parents=True)
FLAGS = ['П1', 'П2', 'П3', 'П4', 'П5', 'П6', 'Э1', 'Э2', 'Э3', 'Э4', 'Э5', 'Э6', 'ПР']
HEADS = ['Номер элемента', 'Наименование элемента', 'Шифр ПД', 'Шифр РД', 'Наименование комплекта', 'Изм.', 'Раздел', 'Сверка с ВЭ РД', 'Согласование ТЭ', 'Дата ТЭ', '№ письма ТЭ',
         'Проверил', 'Дата проверки'] + [f + '. вид замечания' for f in FLAGS] + ['Замечания (текст для письма подрядчику)', 'Примечание', 'Факт СМР', 'Видов замечаний', 'Замечания устранены',
         'Статус комплекта', 'Действующая редакция', 'Порядковый № элемента', 'Элемент строки']
COL = {h.split('.')[0] if h[:2] in FLAGS else h: i for i, h in enumerate(HEADS, 1)}
# The stage list: element, name, sets with the status of ТЭ.
STAGE = [('1.1', 'Мост', [('К1', 'Да'), ('К2', 'Да'), ('К3', 'Замечания')]), ('1.2', 'Труба', [('Т1', 'Да'), ('Т2', 'Да')]), ('1.3', 'Дорога', [('Д1', 'Да')])]
NAME = {el: name for el, name, _ in STAGE}
D = lambda day: datetime(2026, 10, day)


def form(path, rows, stage=True, acts=(), formula_rows=0, extra=None):
    """An RD check book: the check sheet (header in row 5, column numbers in row 6, data from row 7), «Ведомость» and «Проверка актов»."""
    wb = Workbook(); m = wb.active; m.title = MAIN
    m['A1'] = 'Форма проверки рабочей документации'; m['A2'] = '9 этап – «Испытательный участок»'
    for c, h in enumerate(HEADS, 1): m.cell(5, c, h); m.cell(6, c, c)
    for r, row in enumerate(rows, 7):
        for key, value in row.items(): m.cell(r, COL[key], value)
    for r in range(7, 7 + formula_rows): m.cell(r, COL['Статус комплекта'], '=""')  # the formula rows a form has prepared
    if stage:
        v = wb.create_sheet('Ведомость')
        for c, h in {1: 'Номер элемента', 2: 'Наименование элемента', 3: 'Шифр ПД', 4: 'Шифр РД', 5: 'Наименование документации', 9: 'Статус комплекта'}.items(): v.cell(5, c, h)
        for c in range(1, 31): v.cell(6, c, c)
        r = 7
        for el, name, sets in STAGE:
            for k, (code, te) in enumerate(sets):
                if k == 0: v.cell(r, 1, el); v.cell(r, 2, name)
                v.cell(r, 4, code); v.cell(r, 5, 'Комплект ' + code); v.cell(r, 9, te); r += 1
    a = wb.create_sheet('Проверка актов')
    for c, h in enumerate(['№ п/п', 'Номер элемента', 'Наименование элемента', '№ акта', 'Дата акта', 'Проверил', 'Дата проверки', 'А1. акт', 'А2. акт', 'Замечания (текст)', 'Замечания устранены', 'Результат проверки акта'], 1):
        a.cell(5, c, h); a.cell(6, c, c)
    for r, act in enumerate(acts, 7):
        for c, value in act.items(): a.cell(r, c, value)
    if extra: extra(wb)
    wb.save(path); return path


def row(code, el=None, **values):
    data = {'Шифр РД': code}
    if el: data.update({'Номер элемента': el, 'Наименование элемента': NAME[el]})
    data.update({{'who': 'Проверил', 'date': 'Дата проверки', 'text': 'Замечания (текст для письма подрядчику)', 'note': 'Примечание', 'fixed': 'Замечания устранены', 'rev': 'Изм.'}.get(k, k): v for k, v in values.items()})
    return data


def build(output, *sources, day='2026-10-08'):
    result = subprocess.run([str(EXE), '--batch', str(output), day, MAIN, *map(str, sources), '--rd'], timeout=300)
    error = Path(str(output) + '.error.txt')
    if error.exists(): return {'error': error.read_text(encoding='utf-8-sig')}
    assert result.returncode == 0
    return json.loads(Path(str(output) + '.run.json').read_text(encoding='utf-8-sig'))


target = form(B / 'Сводная РД.xlsx', [row('К1', '1.1', who='Анна', date=D(5)), row('К2', who='Анна', date=D(5), П1=1, text='Нет подписи'),
                                       row('К3', who='ТЭ', ПР=1, text='Не согласован ТЭ')],
              acts=[{2: '1.2', 4: 'А-1', 6: 'Борис', 7: D(8), 8: 1}])
boris = form(B / 'Борис.xlsx', [row('Т1', '1.2', who='Борис', date=D(6)), row('Т2', who='Борис', date=D(6), Э3=1), row('К2', '1.1', who='Анна', date=D(5), fixed=D(7), text='Нет подписи')])
vera = form(B / 'Вера.xlsx', [row('Д1', '1.3', who='Вера', date=D(7), note='проверить ВОР'), row('Т2', '1.2', who='Вера', date=D(7), П2=1, text='Нет печати')])
report = build(target, boris, vera)
assert 'error' not in report, report
assert (report['Rd'], report['Documents'], report['Elements'], report['Boxes'], report['Reviewed']) == (True, 6, 3, 2, 6), report
assert report['Pdfs'] and all(Path(p).exists() for p in report['Pdfs']) and report['PdfError'] is None, report
assert [Path(p).name for p in report['Pdfs']] == ['Сводная РД СВОД 2026.10.08.pdf', 'Сводная РД Итоги проверки 2026.10.08.pdf'], report['Pdfs']

wb, cached = load_workbook(target), load_workbook(target, data_only=True)
assert {'СВОД', 'Итоги проверки', 'Статистика', 'Элементы', 'Ведомость', 'Проверка актов', MAIN} <= set(wb.sheetnames), wb.sheetnames
assert wb.active.title == 'СВОД'
m, mv = wb[MAIN], cached[MAIN]
codes = {m.cell(r, COL['Шифр РД']).value: r for r in range(7, m.max_row + 1) if m.cell(r, COL['Шифр РД']).value}
assert list(codes) == ['К1', 'К2', 'К3', 'Т1', 'Т2', 'Д1'], codes
get = lambda code, key: m.cell(codes[code], COL[key]).value
# New sets go after the data in file order; an element number is written only where the element changes.
assert [get(c, 'Номер элемента') for c in ('Т1', 'Т2', 'Д1')] == ['1.2', None, '1.3'] and get('Т1', 'Наименование элемента') == 'Труба'
# Every «1» is kept, the reviewer and the date come from the latest dated check, «Замечания устранены» from the files.
assert (get('Т2', 'Проверил'), get('Т2', 'Дата проверки'), get('Т2', 'П2'), get('Т2', 'Э3'), get('Т2', 'Замечания (текст для письма подрядчику)')) == ('Вера', D(7), 1, 1, 'Нет печати')
assert get('К2', 'Замечания устранены') == D(7) and get('К2', 'Проверил') == 'Анна' and get('К2', 'П1') == 1
status = {c: mv.cell(r, COL['Статус комплекта']).value for c, r in codes.items()}
assert status == {'К1': 'готово к ВПР', 'К2': 'готово к ВПР', 'К3': 'замечания', 'Т1': 'готово к ВПР', 'Т2': 'замечания', 'Д1': 'замечания'}, status
assert all(str(m.cell(r, COL[k]).value).startswith('=') for r in range(7, 7 + 6 + 50) for k in ('Сверка с ВЭ РД', 'Статус комплекта', 'Действующая редакция', 'Элемент строки'))
assert '$D$7:$D$62' in m.cell(7, COL['Действующая редакция']).value, m.cell(7, COL['Действующая редакция']).value

# «Элементы»: the element of a remark of ТЭ has no check date; the others are finished on the day of their last set.
e, ev = wb['Элементы'], cached['Элементы']
val = lambda x: '' if x is None else x
elements = {ev.cell(r, 1).value: [val(ev.cell(r, c).value) for c in (4, 5, 8, 9, 11, 12, 13, 16, 17, 18)] for r in range(6, 9)}
assert elements == {'1.1': [3, 3, 2, 1, 2, 'замечания', 'не поступал', '', '', 1], '1.2': [2, 2, 2, 1, 1, 'замечания', 'замечания', D(7), 'Вера', 0],
                    '1.3': [1, 1, 1, 1, 0, 'замечания', 'не поступал', D(7), 'Вера', 0]}, elements
assert e['A6'].value.startswith('=IFERROR(INDEX(') and ev['A9'].value in (None, '')

# «СВОД»: days from the first check (working days and the report date), the reviewers under the last day; ТЭ under the table.
s, sv = wb['СВОД'], cached['СВОД']
assert sv['U2'].value == D(8) and sv['A8'].value == 'Этап 9', (sv['U2'].value, sv['A8'].value)
assert sv['D4'].value == 'Всего комплектов\nс замечаниями: 3' and sv['H4'].value == 'Устранено замечаний\nпо комплектам: 1' and sv['L4'].value == 'Устраняется комплектов\nс замечаниями: 2', (sv['D4'].value, sv['H4'].value, sv['L4'].value)
days = [r for r in range(8, s.max_row + 1) if isinstance(sv.cell(r, 2).value, datetime)]
assert [sv.cell(r, 2).value for r in days] == [D(5), D(6), D(7), D(8)], [sv.cell(r, 2).value for r in days]
assert [[sv.cell(r, c).value for c in (4, 6, 8, 10, 12, 14, 16, 18)] for r in days] == [[3, 6, 0, 2, 3, 4, 0, 2], [3, 6, 0, 3, 3, 3, 0, 1], [3, 6, 2, 5, 1, 1, 2, 2], [3, 6, 2, 5, 1, 1, 0, 0]]
people = {sv.cell(r, 2).value: [sv.cell(r, c).value for c in (8, 10, 16, 18)] for r in range(days[-1] + 1, days[-1] + 4)}
assert people == {'Анна': [0, 2, 0, 0], 'Борис': [0, 1, 0, 0], 'Вера': [2, 2, 0, 0]}, people
assert all(s.row_dimensions[r].hidden for r in range(days[0] + 1, days[0] + 4)) and not s.row_dimensions[days[-1] + 1].hidden
assert all(s.row_dimensions[r].outline_level == 1 for r in range(days[-1] + 1, days[-1] + 4))
te = days[-1] + 5
assert sv.cell(te, 1).value.startswith('Замечания ТЭ') and (sv.cell(te, 4).value, sv.cell(te, 6).value) == (1, 1), [sv.cell(te, c).value for c in range(1, 8)]
assert s['J8'].value.startswith('=IF($B8>$U$2,"",COUNTIFS(') and 'Элементы' in s['H8'].value

# «Итоги проверки»: elements finished and sets checked by reviewer; the bold of non-zero values; the alternative table with the acts.
t, tv = wb['Итоги проверки'], cached['Итоги проверки']
assert [tv.cell(4, c).value for c in (3, 5, 7, 9)] == ['Анна', 'Борис', 'Вера', 'ВСЕГО']
assert [tv.cell(6, c).value for c in range(3, 11)] == [0, 2, 0, 1, 2, 2, 2, 5], [tv.cell(6, c).value for c in range(3, 11)]
assert [tv.cell(9, c).value for c in range(3, 11)] == [0, 0, 0, 0, 2, 2, 2, 2] and t.cell(9, 7).font.b and not t.cell(9, 3).font.b
alt = next(r for r in range(1, t.max_row + 1) if tv.cell(r, 2).value == 'Борис' and r > 12)
assert [tv.cell(alt, c).value for c in (3, 4, 5, 6)] == [1, 0, 0, 1], [tv.cell(alt, c).value for c in (3, 4, 5, 6)]
assert '$A$1:$S$' in s.print_area and '$A$1:$J$10' in t.print_area, (s.print_area, t.print_area)

# «Статистика»: sets of the stage list, presented, checked, statuses.
st = cached['Статистика']
stats = {st.cell(r, 1).value: st.cell(r, 2).value for r in range(6, st.max_row + 1) if st.cell(r, 1).value}
assert (stats['В ведомости элементов РД (ВЭ РД)'], stats['Представлено на проверку (действующие редакции из ВЭ РД)'], stats['Проверено по состоянию на дату отчёта'], stats['Готово к ВПР'], stats['Замечания ТЭ («Проверил: ТЭ»)']) == (6, 6, 5, 3, 1), stats

# The same files again change nothing; a build without files recalculates the reports from the book.
before = {(n, c.coordinate): c.value for n in (MAIN, 'СВОД', 'Итоги проверки', 'Элементы') for r in load_workbook(target, data_only=True)[n].iter_rows() for c in r}
again = build(target, boris, vera)
after = {(n, c.coordinate): c.value for n in (MAIN, 'СВОД', 'Итоги проверки', 'Элементы') for r in load_workbook(target, data_only=True)[n].iter_rows() for c in r}
assert again['Documents'] == 6 and before == after, [k for k in before if before[k] != after.get(k)][:10]
alone = build(target)
assert alone['Files'] == 0 and alone['Documents'] == 6 and alone['Boxes'] == 2, alone

# A form with prepared formula rows: the ranges of the formulas, of other sheets and of the validation follow the data down.
def own_sheet(wb):
    wb.create_sheet('Мой лист')['A1'] = "=COUNTIF('Проверка комплектов'!$D$7:$D$12,\"?*\")"
grown = form(B / 'Рост.xlsx', [row('К1', '1.1', who='Анна', date=D(5))], formula_rows=6, extra=own_sheet)
many = form(B / 'Много.xlsx', [row('Х%d' % i, '1.2' if i == 0 else None, who='Борис', date=D(6)) for i in range(60)])
result = build(grown, many)
assert result['Documents'] == 61, result
g = load_workbook(grown)
assert g['Мой лист']['A1'].value == "=COUNTIF('Проверка комплектов'!$D$7:$D$117,\"?*\")", g['Мой лист']['A1'].value
assert '$D$7:$D$117' in g[MAIN].cell(7, COL['Действующая редакция']).value
sets_not_in_stage = [load_workbook(grown, data_only=True)[MAIN].cell(r, COL['Сверка с ВЭ РД']).value for r in range(8, 68)]
assert set(sets_not_in_stage) == {'нет в ВЭ РД'}, set(sets_not_in_stage)

# The row of column numbers moved among the data by a sort: skipped in a reviewer's copy, the consolidated book is not touched.
def numbers_at_end(path):
    wb = load_workbook(path); m = wb[MAIN]; values = [c.value for c in m[6]]; m.delete_rows(6); last = max(r for r in range(1, m.max_row + 1) if m.cell(r, 4).value) + 1
    for c, v in enumerate(values, 1): m.cell(last, c, v)
    wb.save(path); return path
copy = numbers_at_end(form(B / 'Копия с номерами.xlsx', [row('Т1', '1.2', who='Борис', date=D(6))]))
fine = form(B / 'Сводная 2.xlsx', [row('К1', '1.1', who='Анна', date=D(5))])
skipped = build(fine, copy)
assert 'error' not in skipped and skipped['Documents'] == 2 and any('номерами столбцов' in p for p in skipped['ProblemList']), skipped
sorted_target = numbers_at_end(form(B / 'Сводная 3.xlsx', [row('К1', '1.1', who='Анна', date=D(5))]))
before = Path(sorted_target).read_bytes()
refused = build(sorted_target, boris)
assert 'error' in refused and 'номерами столбцов' in refused['error'] and Path(sorted_target).read_bytes() == before, refused

# A book without «Ведомость» is not an RD form: nothing is written.
bare = form(B / 'Без ведомости.xlsx', [row('К1', '1.1', who='Анна', date=D(5))], stage=False)
failed = build(bare, boris)
assert 'error' in failed and 'Ведомость' in failed['error'], failed
print('RD tests passed: merge by set code, new sets after the data, statuses, elements, «СВОД», «Итоги проверки», «Статистика», PDF, repeated and file-less builds, growing ranges, missing stage list')
