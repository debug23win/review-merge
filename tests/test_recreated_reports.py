"""Rebuilt report layouts, reviewer histories and unique global totals; artificial data only."""
from pathlib import Path
from datetime import datetime
import hashlib, subprocess, uuid, math, posixpath, zipfile
import xml.etree.ElementTree as ET
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter as letter

ROOT = Path(__file__).resolve().parents[1]
B = ROOT / 'tests/.runs' / uuid.uuid4().hex
B.mkdir(parents=True)
REPORTS = {'СВОД', 'Стат. по типам', 'Стат. по зам.', 'Итоги работы'}
MAIN = 'Все загруженные файлы'
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
AS_OF = datetime(2026, 10, 4)
EDGES = ('left', 'right', 'top', 'bottom')


def assert_unframed(sheet, first=1, last=3, width=None):
    """Check every cell, including otherwise empty merged/title cells."""
    for row in sheet.iter_rows(min_row=first, max_row=last, max_col=width or sheet.max_column):
        for cell in row:
            assert not cell.alignment.wrap_text, (sheet.title, cell.coordinate, 'wrapped caption')
            assert all(getattr(cell.border, edge).style in (None, 'none') for edge in EDGES), (
                sheet.title, cell.coordinate, 'border outside table', cell.border)


def sheet_xml(path, name):
    """Resolve sheet parts instead of depending on workbook sheet ordering."""
    with zipfile.ZipFile(path) as archive:
        book = ET.fromstring(archive.read('xl/workbook.xml'))
        rels = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
        targets = {rel.get('Id'): rel.get('Target') for rel in rels}
        entry = next(sheet for sheet in book.find('s:sheets', NS) if sheet.get('name') == name)
        target = targets[entry.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')]
        member = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
        return ET.fromstring(archive.read(member))


def column_width(doc, column):
    matches = [node for node in doc.findall('s:cols/s:col', NS)
               if int(node.get('min')) <= column <= int(node.get('max'))]
    assert len(matches) == 1, (column, matches)
    return float(matches[0].get('width'))


def assert_cached(formulas, cached, row, column, expected):
    cell = formulas.cell(row, column)
    actual = cached.cell(row, column).value
    assert cell.data_type == 'f', (formulas.title, cell.coordinate, cell.value)
    assert '#REF!' not in cell.value, cell.value
    if expected is None:
        assert actual in (None, ''), (formulas.title, cell.coordinate, actual, 'future should be blank')
    else:
        assert isinstance(actual, (int, float)) and math.isclose(actual, expected, abs_tol=1e-9), (
            formulas.title, cell.coordinate, actual, expected, cell.value)
    return cell.value


def assert_day_formula(formula, anchor, cumulative):
    assert 'COUNTIFS(' in formula and "'Справка по томам'!" in formula, formula
    assert anchor in formula, formula
    if cumulative:
        assert '\">0\"' in formula and '\"<=\"&' + anchor in formula, formula
    else:
        assert '\"<=\"' not in formula, formula


def report_snapshot(path):
    """Compare semantic styles and values, not regenerated style IDs or ZIP bytes."""
    formulas, cached = load_workbook(path), load_workbook(path, data_only=True)
    result = {}
    for name in sorted(REPORTS):
        sheet = formulas[name]
        cells = []
        for row in sheet:
            for cell in row:
                color = cell.font.color
                cells.append((cell.coordinate, cell.value, cached[name][cell.coordinate].value,
                              cell.font.name, cell.font.sz, cell.font.bold,
                              None if color is None else (color.type, str(color.value)),
                              cell.number_format, cell.alignment.wrap_text,
                              tuple(getattr(cell.border, edge).style for edge in EDGES),
                              cell.fill.patternType, cell.fill.fgColor.type, cell.fill.fgColor.index))
        doc = sheet_xml(path, name)
        result[name] = (cells, sorted(map(str, sheet.merged_cells.ranges)), str(sheet.print_area),
                        doc.find('s:dimension', NS).get('ref'),
                        [(int(col.get('min')), int(col.get('max')), float(col.get('width')))
                         for col in doc.findall('s:cols/s:col', NS)],
                        [(r, d.outlineLevel, d.hidden, d.collapsed, d.height)
                         for r, d in sorted(sheet.row_dimensions.items())])
    formulas.close()
    cached.close()
    return result

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
        assert_unframed(s)
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

def test_layout_migration():
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


def artificial_source(path, empty=False):
    """Use real register headers with entirely artificial, deterministic review rows."""
    book = load_workbook(ROOT / 'tests/fixtures/Анна.xlsx')
    main = book[MAIN]
    main.delete_rows(7, main.max_row)
    for sheet in list(book):
        if sheet.title.upper() == 'СВОД' or sheet.title in REPORTS or sheet.title == 'Справка по томам':
            book.remove(sheet)
    # A source calendar makes the first and last report dates explicit. On a repeat
    # merge these dates are also recoverable from dated register rows plus as-of.
    dates = [AS_OF] if empty else [datetime(2026, 10, day) for day in range(1, 6)]
    old = book.create_sheet('Свод')
    old['A7'], old['A8'] = 'Получено коробов', 'Томов по акту'
    for column, date in enumerate(dates, 2):
        old.cell(5, column, date)
    ref = book.create_sheet('Справка по томам')
    ref['A1'], ref['B1'] = 'Коробов получено', 20
    ref['A2'], ref['B2'] = 'Томов по проекту акта', 100
    ref['A8'] = 'ФИО проверяющего'
    ref['A42'], ref['A43'] = 'ИТОГО проверено коробов', 'ИТОГО проверено томов'
    if empty:
        records = [('Без проверки', None, None, 11)]
    else:
        ref['A12'] = 'Громов Григорий Григорьевич'  # Not present in any file row.
        records = [
            ('Совместный_Часть1', 'Анна', 1, 11),
            ('Совместный_Часть2', 'Борис', 2, 11),
            ('Совместный-ИУЛ', 'Анна', 1, None),
            ('Первый том', 'Анна', 1, 12),
            ('Второй том', 'Борис', 2, 11),  # Same physical box as the shared volume.
            ('Поздний том', 'Анна', 3, 13),
            ('Будущий том', 'Вера', 5, 14),
            ('Без даты', 'Зорина', None, 15),  # Reviewer exists, but has no dated work.
            ('Без проверяющего', None, None, 16),
            ('Отдельный-ИУЛ', 'Борис', 2, 17),  # A checked file/box, never a volume.
        ]
    for row, (filename, reviewer, day, box) in enumerate(records, 7):
        for column, value in {1: row - 6, 2: 'Искусственный раздел', 3: filename,
                              4: 'pdf', 5: '1 МБ', 6: '01.10.2026', 7: '10:00',
                              8: 'FEED%04d' % row, 9: reviewer,
                              10: None if day is None else datetime(2026, 10, day),
                              11: 'ПД', 12: box}.items():
            main.cell(row, column, value)
    # Deliberately hostile legacy caption formatting must not leak into outputs.
    heavy = Side(style='thick', color='FF000000')
    for row in old.iter_rows(min_row=1, max_row=3, max_col=19):
        for cell in row:
            cell.border = Border(left=heavy, right=heavy, top=heavy, bottom=heavy)
            cell.alignment = Alignment(wrap_text=True)
    book.save(path)
    book.close()
    return dates


def assert_matrix_layout(path, sheet, reviewers, dates):
    """Two columns per reviewer plus a real, final two-column ВСЕГО group."""
    end = 4 + 2 * len(reviewers)
    total = end - 1
    last = 6 + len(dates)
    headers = {sheet.cell(4, column).value: column for column in range(3, end + 1, 2)}
    assert set(headers) == set(reviewers) | {'ВСЕГО'}, headers
    assert headers['ВСЕГО'] == total, headers
    assert sheet['B6'].value == 'Всего'
    assert len(headers) == len(reviewers) + 1
    merges = set(map(str, sheet.merged_cells.ranges))
    assert 'A1:' + letter(end) + '1' in merges
    assert 'A4:A5' in merges and 'B4:B5' in merges
    for column in range(3, end + 1, 2):
        assert '%s4:%s4' % (letter(column), letter(column + 1)) in merges
        assert sheet.cell(5, column).value == 'коробов'
        assert sheet.cell(5, column + 1).value == 'томов'
        assert sheet.cell(4, column).font.bold
        for row in range(5, last + 1):
            assert sheet.cell(row, column).border.right.style == 'thin'
            assert sheet.cell(row, column + 1).border.left.style == 'thin'
    # Inspect actual OOXML col intervals: openpyxl does not expand grouped widths.
    doc = sheet_xml(path, sheet.title)
    for column in range(3, end + 1):
        assert column_width(doc, column) == 9, (column, column_width(doc, column))
    assert doc.find('s:dimension', NS).get('ref').split(':')[-1].rstrip('0123456789') == letter(max(8, end))
    areas = str(sheet.print_area)
    assert '$A$1:$%s$%d' % (letter(end), last) in areas, areas
    # The matrix must have no nonempty columns after ВСЕГО, even when a wider
    # alternative table below requires eight sheet columns (zero-reviewer case).
    for row in sheet.iter_rows(min_row=4, max_row=last, min_col=end + 1,
                               max_col=max(end + 1, sheet.max_column)):
        assert all(cell.value is None for cell in row), row
    for row in range(4, last + 1):
        assert sheet.cell(row, 1).border.left.style == 'thick'
        assert sheet.cell(row, end).border.right.style == 'thick'
        for column in range(3, end, 2):
            assert sheet.cell(row, column).border.left.style == 'medium'
            assert sheet.cell(row, column - 1).border.right.style == 'medium'
    for column in range(1, end + 1):
        assert sheet.cell(4, column).border.top.style == 'thick'
        assert sheet.cell(last, column).border.bottom.style == 'thick'
    assert_unframed(sheet)
    alternative = next(cell.row for row in sheet for cell in row
                       if cell.value == 'Альтернативное представление по проверяющим')
    assert alternative == last + 3
    assert_unframed(sheet, alternative, alternative, 8)
    assert [sheet.cell(7 + index, 2).value for index in range(len(dates))] == dates
    return headers


def test_reviewer_histories_and_global_totals():
    source, target = B / 'reviewer-history-source.xlsx', B / 'reviewer-history.xlsx'
    dates = artificial_source(source)
    digest = hashlib.sha256(source.read_bytes()).hexdigest()
    run(target, source)
    book, cached = load_workbook(target), load_workbook(target, data_only=True)
    reviewers = {'Анна', 'Борис', 'Вера', 'Громов', 'Зорина'}
    summary, summary_values = book['СВОД'], cached['СВОД']
    work, work_values = book['Итоги работы'], cached['Итоги работы']
    for name in REPORTS:
        assert_unframed(book[name])
    header = assert_matrix_layout(target, work, reviewers, dates)
    assert not any(cell.value == 'Громов Григорий Григорьевич' for row in work for cell in row), (
        'Reference-only reviewer must use the same surname label as other reports')
    # Expected unique completions: box 11 first counted for Anna on October 1;
    # its two-part/ИУЛ volume counted once for Boris on October 2. A standalone
    # ИУЛ is never a volume. Reviewer without a valid date contributes zero.
    daily = {
        1: {'Анна': (2, 1)},
        2: {'Борис': (1, 2)},
        3: {'Анна': (1, 1)},
        4: {},
    }
    global_daily = {1: (2, 1), 2: (1, 2), 3: (1, 1), 4: (0, 0)}
    cumulative = {name: [0, 0] for name in reviewers}
    global_cumulative = [0, 0]
    assert summary.max_row == 6 + len(dates) * (len(reviewers) + 1)
    assert summary.max_column == 19
    assert sheet_xml(target, 'СВОД').find('s:dimension', NS).get('ref') == 'A1:S36'
    seen = []
    for index, date in enumerate(dates):
        row = 7 + index * (len(reviewers) + 1)
        assert summary.cell(row, 2).value == date
        assert summary.cell(row, 2).font.bold
        names = [summary.cell(r, 2).value for r in range(row + 1, row + len(reviewers) + 1)]
        assert len(names) == len(set(names)) and set(names) == reviewers, (date, names)
        seen.append(names)
        future = date > AS_OF
        if not future:
            for name in reviewers:
                for kind in range(2):
                    cumulative[name][kind] += daily[date.day].get(name, (0, 0))[kind]
            for kind in range(2):
                global_cumulative[kind] += global_daily[date.day][kind]
        for offset, name in enumerate([None] + names):
            r = row + offset
            anchor = '$B%d' % row if name is None else '$B$%d' % row
            if name is not None:
                assert summary.row_dimensions[r].outlineLevel == 1
                for column in list(range(4, 8)) + list(range(12, 16)):
                    cell = summary.cell(r, column)
                    assert cell.value in (None, '') and cell.data_type != 'f', (cell.coordinate, cell.value)
                    assert summary_values.cell(r, column).value in (None, '')
            for kind, denominator in enumerate((20, 100)):
                done = global_cumulative[kind] if name is None else cumulative[name][kind]
                today = None if future else (global_daily[date.day][kind] if name is None
                                             else daily[date.day].get(name, (0, 0))[kind])
                formula = assert_cached(summary, summary_values, r, 8 + kind * 2, None if future else done)
                assert_day_formula(formula, anchor, cumulative=True)
                if name == 'Громов':
                    criterion = 'Громов Григорий Григорьевич' if kind == 0 else 'Громов'
                    assert '"' + criterion + '"' in formula, formula
                assert_cached(summary, summary_values, r, 9 + kind * 2,
                              None if future else done / denominator * 100)
                formula = assert_cached(summary, summary_values, r, 16 + kind * 2, today)
                assert_day_formula(formula, anchor, cumulative=False)
                assert_cached(summary, summary_values, r, 17 + kind * 2,
                              None if future else today / denominator * 100)
                if name is None:
                    assert_cached(summary, summary_values, r, 4 + kind * 2, None if future else denominator)
                    assert_cached(summary, summary_values, r, 5 + kind * 2, None if future else 100)
                    assert_cached(summary, summary_values, r, 12 + kind * 2,
                                  None if future else denominator - done)
                    assert_cached(summary, summary_values, r, 13 + kind * 2,
                                  None if future else (denominator - done) / denominator * 100)
        matrix_row = 7 + index
        for name, column in header.items():
            for kind in range(2):
                value = None if future else (global_daily[date.day][kind] if name == 'ВСЕГО'
                                              else daily[date.day].get(name, (0, 0))[kind])
                formula = assert_cached(work, work_values, matrix_row, column + kind, value)
                assert_day_formula(formula, '$B%d' % matrix_row, cumulative=False)
                if name == 'ВСЕГО':
                    assert all('"' + who + '"' not in formula for who in reviewers), formula
    assert all(names == seen[0] for names in seen), 'Reviewer order must be stable under every date'
    for name, column in header.items():
        expected = global_cumulative if name == 'ВСЕГО' else cumulative[name]
        for kind in range(2):
            formula = assert_cached(work, work_values, 6, column + kind, expected[kind])
            assert 'COUNTIFS(' in formula and '\"<=\"&' in formula, formula
            assert work.cell(6, column + kind).font.bold
            if name == 'ВСЕГО':
                assert all('"' + who + '"' not in formula for who in reviewers), formula
    assert tuple(global_cumulative) == (4, 4)
    assert tuple(cumulative['Анна']) == (3, 2)
    assert tuple(cumulative['Борис']) == (1, 2)
    # Source received totals remain independent of the reviewed global totals.
    assert cached['Справка по томам']['B1'].value == 20
    assert cached['Справка по томам']['B2'].value == 100
    book.close()
    cached.close()
    before = report_snapshot(target)
    run(target, source)
    assert report_snapshot(target) == before, 'Repeat merge changed formulas, caches, grouping, dimensions or styles'
    assert hashlib.sha256(source.read_bytes()).hexdigest() == digest


def test_no_reviewers():
    source, target = B / 'no-reviewers-source.xlsx', B / 'no-reviewers.xlsx'
    dates = artificial_source(source, empty=True)
    run(target, source)
    book, cached = load_workbook(target), load_workbook(target, data_only=True)
    headers = assert_matrix_layout(target, book['Итоги работы'], set(), dates)
    assert headers == {'ВСЕГО': 3}
    for row in (6, 7):
        for column in (3, 4):
            assert_cached(book['Итоги работы'], cached['Итоги работы'], row, column, 0)
    summary = book['СВОД']
    assert summary.max_row == 7 and summary.max_column == 19
    assert summary['B7'].value == AS_OF
    assert not summary.row_dimensions[7].outlineLevel
    assert not any(d.outlineLevel for d in summary.row_dimensions.values())
    for column in (8, 9, 10, 11, 16, 17, 18, 19):
        assert_cached(summary, cached['СВОД'], 7, column, 0)
    assert_cached(summary, cached['СВОД'], 7, 4, 20)
    assert_cached(summary, cached['СВОД'], 7, 6, 100)
    assert_cached(summary, cached['СВОД'], 7, 12, 20)
    assert_cached(summary, cached['СВОД'], 7, 14, 100)
    for name in REPORTS:
        assert_unframed(book[name])
    book.close()
    cached.close()
    before = report_snapshot(target)
    run(target, source)
    assert report_snapshot(target) == before, 'Zero-reviewer repeat merge changed reports'


def test_same_surname_reference_reviewer():
    """A reference-only person must not absorb an active namesake's statistics."""
    source, target = B / 'namesakes-source.xlsx', B / 'namesakes.xlsx'
    dates = artificial_source(source)
    book = load_workbook(source)
    active, idle = 'Иванов Иван Иванович', 'Иванов Игорь Игоревич'
    book['Справка по томам']['A12'] = idle
    for row in book[MAIN]:
        if row[8].value == 'Анна':
            row[8].value = active
    book.save(source)
    book.close()
    run(target, source)
    book, cached = load_workbook(target), load_workbook(target, data_only=True)
    reviewers = {active, idle, 'Борис', 'Вера', 'Зорина'}
    headers = assert_matrix_layout(target, book['Итоги работы'], reviewers, dates)
    assert 'Иванов' not in headers
    for name, expected in ((active, (3, 2)), (idle, (0, 0))):
        for kind in range(2):
            formula = assert_cached(book['Итоги работы'], cached['Итоги работы'],
                                    6, headers[name] + kind, expected[kind])
            assert '"' + name + '"' in formula, formula
    summary = book['СВОД']
    for index, date in enumerate(dates):
        start = 7 + index * (len(reviewers) + 1)
        rows = {summary.cell(row, 2).value: row for row in range(start + 1, start + len(reviewers) + 1)}
        assert set(rows) == reviewers, rows
        for column in (8, 10, 16, 18):
            formula = assert_cached(summary, cached['СВОД'], rows[idle], column,
                                    None if date > AS_OF else 0)
            assert '"' + idle + '"' in formula, formula
    book.close()
    cached.close()


if __name__ == '__main__':
    test_layout_migration()
    test_reviewer_histories_and_global_totals()
    test_no_reviewers()
    test_same_surname_reference_reviewer()
    print('Report recreation passed: old/absent/custom layouts; all-date cumulative reviewer rows; '
          'daily and unique global cached totals; future dates; reference-only/undated/zero reviewers; namesakes; '
          'widths, merged groups, borders, unwrapped captions and repeat merge')
    print(B)
