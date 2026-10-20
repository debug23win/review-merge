"""Independent PD regression cases: resolved marks, user columns, dates and grouping."""
import json, re, subprocess, uuid, shutil, os
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT / 'tests/.runs' / uuid.uuid4().hex
BASE.mkdir(parents=True)
MAIN = 'Все загруженные файлы'

def fixture(name, records, resolved=False):
    path = BASE / name
    w = load_workbook(ROOT / 'tests/fixtures/Анна.xlsx')
    m = w[MAIN]; m.delete_rows(7, m.max_row)
    if resolved: m['U5'] = 'Замечания устранены'; m['U6'] = 21
    for r, data in enumerate(records, 7):
        values = {1:r-6, 3:'Том '+str(r-6), 4:'pdf', 6:'01.10.2026', 7:'10:00', 8:f'{r:08X}', 9:'Анна', 10:datetime(2026,10,2), 11:'ПД', 12:1}
        values.update(data)
        for c, value in values.items(): m.cell(r,c,value)
    w.save(path)
    return path

def build(path, *sources):
    subprocess.run([str(ROOT/'Свод_проверки.exe'),'--batch',str(path),'2026-10-04',MAIN,str(path),*map(str,sources)],check=True,timeout=180)
    return json.loads(Path(str(path)+'.run.json').read_text(encoding='utf-8-sig'))

def summary(path):
    s=load_workbook(path,data_only=True)['СВОД']
    return tuple(int(re.search(r': (\d+)$',s.cell(4,c).value).group(1)) for c in (4,8,12))

def resolved_column(m):
    return next(c.column for c in m[5] if c.value=='Замечания устранены')

def test_marks():
    marks=['0.0','-0','0,00','0e3',False,'\u00a0\t','нет',1,datetime(2026,10,4),'да',True]
    path=fixture('marks.xlsx',[{13:1,21:mark} for mark in marks],True)
    build(path)
    assert summary(path)==(11,4,7), ('resolved zeros and booleans',summary(path))
    build(path)
    assert summary(path)==(11,4,7), 'repeat merge must preserve the same resolved count'
    return path

def test_user_columns():
    path=fixture('user-columns.xlsx',[{13:1},{13:1}])
    w=load_workbook(path);m=w[MAIN]
    m['U7']='Сохранить пользовательский текст'
    m['V8']='=1+1'
    w.save(path);build(path)
    w=load_workbook(path);m=w[MAIN]
    assert m['V7'].value=='Сохранить пользовательский текст' and m['W8'].value=='=1+1'
    assert resolved_column(m)==24, 'new header must follow data and formulas, including columns without a header'
    assert summary(path)==(2,0,2)
    build(path)
    w=load_workbook(path);m=w[MAIN]
    assert m['V7'].value=='Сохранить пользовательский текст' and m['W8'].value=='=1+1'
    assert resolved_column(m)==24 and summary(path)==(2,0,2)

def test_changed_date_format():
    path=fixture('mark-format.xlsx',[{13:1,21:datetime(2026,10,3)}],True)
    build(path);w=load_workbook(path);m=w[MAIN];c=m.cell(7,resolved_column(m));c.value=1;w.save(path)
    build(path);m=load_workbook(path)[MAIN];c=m.cell(7,resolved_column(m))
    assert c.value==1 and not c.is_date, (c.value,c.number_format)

def test_invalid_marks_and_whitespace():
    path=fixture('invalid-marks.xlsx',[{13:'+',21:1},{13:2,21:1},{19:'\u00a0\t ',21:1},{13:1,21:1}],True)
    report=build(path)
    assert report['WithIssues']==1 and summary(path)==(1,1,0), (report['WithIssues'],summary(path))

def test_rules():
    compiler=Path(os.environ['WINDIR'])/'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
    if not compiler.exists(): compiler=Path(os.environ['WINDIR'])/'Microsoft.NET/Framework/v4.0.30319/csc.exe'
    target=BASE/'PdRulesTest.exe'
    subprocess.run([str(compiler),'/nologo','/codepage:65001','/out:'+str(target),'/r:'+str(ROOT/'Свод_проверки.exe'),str(ROOT/'tests/PdRulesTest.cs')],check=True)
    shutil.copy2(ROOT/'Свод_проверки.exe',BASE/'Свод_проверки.exe')
    subprocess.run([str(target)],check=True,timeout=20)

def test_fragments():
    path=fixture('fragments.xlsx',[{3:'Том А фрагмент 1',13:1},{3:'Том А фрагмент 2',13:1,21:1},{3:'Том А-УЛ'},
                                   {3:'Том Б',13:1},{3:'Том В',21:1}],True)
    build(path)
    assert summary(path)==(2,1,1),summary(path)
    return path

def run(path, date, *args):
    """A batch build of path (also its first source) with more sources and switches; the run report or the error text."""
    for suffix in ('.run.json', '.error.txt'):
        if Path(str(path)+suffix).exists(): Path(str(path)+suffix).unlink()
    subprocess.run([str(ROOT/'Свод_проверки.exe'),'--batch',str(path),date,MAIN,str(path),*map(str,args)],timeout=180)
    error=Path(str(path)+'.error.txt')
    if error.exists(): return {'error':error.read_text(encoding='utf-8-sig')}
    return json.loads(Path(str(path)+'.run.json').read_text(encoding='utf-8-sig'))

def cached(path, sheet):
    ws=load_workbook(path,data_only=True)[sheet]
    return {c.coordinate:c.value for row in ws.iter_rows() for c in row if c.value is not None}

def main_values(path):
    m=load_workbook(path)[MAIN]
    return {(r,c):m.cell(r,c).value for r in range(7,m.max_row+1) for c in range(1,23)}

def test_reviewer_names_stable():
    # Without «Сопоставлять по фамилии» a bare surname joins the only full name, also after a short spelling got its own row.
    path=fixture('names.xlsx',[{9:'Яковлева Александра Алексеевна'},{9:None,10:None},{9:None,10:None},{9:None,10:None}])
    run(path,'2026-10-04')
    files=fixture('names-files.xlsx',[{9:'Яковлева Александра Алексеевна'},{9:'Яковлева',10:datetime(2026,10,3)},{9:'Яковлева',10:datetime(2026,10,3)},{9:'АА Яковлева',10:datetime(2026,10,3)}])
    run(path,'2026-10-04',files)
    first=cached(path,'Итоги работы')
    for _ in range(2):
        run(path,'2026-10-04',files)
        assert cached(path,'Итоги работы')==first, 'a repeated build moved volumes to another person'
    people=[load_workbook(path,data_only=True)['Справка по томам'].cell(r,1).value for r in range(12,42)]
    assert 'Яковлева' not in people and 'Яковлева Александра Алексеевна' in people, people

def test_prefilled_reviewer_gets_date():
    # A reviewer typed in the consolidated book without a date takes the date of the same person's check.
    path=fixture('prefilled.xlsx',[{9:'Иванов И.И.',10:None},{9:'Петров',10:None},{9:'Анна',10:None}])
    files=fixture('prefilled-files.xlsx',[{9:'Иванов Иван Иванович',10:datetime(2026,10,2)},{9:'Петров П.П.',10:datetime(2026,10,3)},{9:'Борис',10:datetime(2026,10,3)}])
    report=run(path,'2026-10-04',files)
    m=load_workbook(path)[MAIN]
    assert [m.cell(r,9).value for r in (7,8,9)]==['Иванов И.И.','Петров','Анна']
    assert m['J7'].value==datetime(2026,10,2) and m['J8'].value==datetime(2026,10,3) and m['J9'].value is None, (m['J7'].value,m['J8'].value,m['J9'].value)
    assert [p for p in report['ProblemList'] if 'отсутствует дата' in p]==['prefilled.xlsx, строка 9: Указан проверяющий, но отсутствует дата проверки.'], report['ProblemList']

def test_missing_crc_finds_document():
    path=fixture('crc.xlsx',[{8:'AAAA0001',9:None,10:None},{8:'AAAA0002',9:None,10:None}])
    files=fixture('crc-files.xlsx',[{8:None,13:1},{8:'AAAA0002'}])
    report=run(path,'2026-10-04',files)
    m=load_workbook(path)[MAIN]
    assert report['Documents']==2 and m.max_row==8 and m['I7'].value=='Анна' and m['M7'].value is None and m['N7'].value==1, (report['Documents'],m.max_row,m['I7'].value)
    assert any('документ найден по имени, формату, дате и времени выгрузки (CRC32 AAAA0001)' in p for p in report['ProblemList']), report['ProblemList']

def test_conflicts_shown_once():
    path=fixture('conflicts.xlsx',[{9:None,10:None},{9:None,10:None}])
    a=fixture('conflicts-a.xlsx',[{9:'Иванов',10:datetime(2026,10,2)},{9:'Иванов',10:datetime(2026,10,2)}])
    b=fixture('conflicts-b.xlsx',[{9:'Петров',10:datetime(2026,10,3)},{9:'Петров',10:datetime(2026,10,3)}])
    assert run(path,'2026-10-04',a,b)['Conflicts']==4
    report=run(path,'2026-10-04',a,b)
    assert report['Conflicts']==0, report['ConflictList']

def test_recalculation_keeps_register():
    path=fixture('recalc.xlsx',[{},{10:None},{9:None,10:None,13:'да'}])
    run(path,'2026-10-04')
    files=fixture('recalc-files.xlsx',[{},{},{9:'Борис',10:datetime(2026,10,3),13:1}])
    before=main_values(path)
    report=run(path,'2026-10-04',files,'--recalculate-history','--current-date-for-new','--fill-incomplete','--convert-marks')
    assert 'error' not in report, report
    assert main_values(path)==before, [(k,before[k],v) for k,v in main_values(path).items() if before.get(k)!=v]
    assert report['Files']==0 and report['DatesAssignedToday']==0 and report['FilledRows']==0 and report['ConvertedMarks']==[], report
    assert any('файлы проверяющих не загружаются (1): recalc-files.xlsx' in line for line in report['Log']), report['Log']

def test_new_rows_follow_data():
    path=fixture('rows.xlsx',[{},{}])
    w=load_workbook(path);m=w[MAIN]
    for r in range(9,20): m.cell(r,10).number_format='dd.mm.yyyy'
    w.save(path)
    files=fixture('rows-files.xlsx',[{},{},{3:'Новый том',8:'BBBB0001'}])
    run(path,'2026-10-04',files)
    m=load_workbook(path)[MAIN]
    assert [m.cell(9,c).value for c in (1,3,4,8,9)]==[3,'Новый том','pdf','BBBB0001','Анна'], [m.cell(9,c).value for c in (1,3,4,8,9)]
    assert all(m.cell(r,c).value is None for r in range(10,21) for c in range(1,23)), 'a new document left a gap of formatted rows'

def test_damaged_source():
    path=fixture('damaged.xlsx',[{}]);run(path,'2026-10-04')
    bad=BASE/'broken.xlsx';bad.write_bytes(b'PK\x03\x04broken')
    before=path.read_bytes()
    report=run(path,'2026-10-04',bad)
    assert 'broken.xlsx: не удалось прочитать книгу .xlsx' in report.get('error',''), report
    assert path.read_bytes()==before
    report=run(path,'2026-10-04',BASE/'none.xlsx')
    assert 'Файл не найден: ' in report.get('error',''), report

def test_kind_of_volume():
    path=fixture('kind.xlsx',[{3:'Том А фрагмент 1',12:5},{3:'Том А фрагмент 2',11:None,12:5},{3:'Том Б',12:6}])
    report=run(path,'2026-10-04')
    assert load_workbook(path)[MAIN]['K8'].value=='ПД' and report['Boxes']==2, report['Boxes']
    assert any('Не указан вид документации: взят вид тома «ПД»' in p for p in report['ProblemList']), report['ProblemList']

def test_check_dates():
    # A date that makes the period longer than two years is named with its file and row; the book is not changed.
    path=fixture('dates.xlsx',[{},{10:'06.10.2020'}])
    before=path.read_bytes()
    report=run(path,'2026-10-04')
    assert 'самая ранняя дата проверки 06.10.2020 — «Том 2» (dates.xlsx, строка 8)' in report.get('error',''), report
    assert path.read_bytes()==before
    # A check before the upload date or after the report date is a warning.
    path=fixture('dates-warn.xlsx',[{},{10:datetime(2026,9,15)},{10:datetime(2062,10,1)}])
    report=run(path,'2026-10-04')
    problems=report['ProblemList']
    assert 'dates-warn.xlsx, строка 8: Дата проверки 15.09.2026 раньше даты выгрузки файла 01.10.2026: проверьте столбец J.' in problems, problems
    assert any(p.startswith('dates-warn.xlsx, строка 9: Дата проверки 01.10.2062 позже даты отчёта 04.10.2026') for p in problems), problems

def test_volume_column_references():
    # Inserting «Номер тома» keeps filters, notes, frozen columns, charts and the numbers row on their columns,
    # writes no empty texts into M and puts calcPr before extLst.
    from openpyxl.comments import Comment
    from openpyxl.chart import BarChart, Reference
    import zipfile
    path=fixture('migrate.xlsx',[{14:1},{},{}],True)
    w=load_workbook(path);m=w[MAIN]
    m.auto_filter.ref='A5:U9';m.auto_filter.add_filter_column(13,['1'])
    m['N7'].comment=Comment('шифр','QA');m.freeze_panes='N7';m['W5']='Моё';m['W6']=23
    chart=BarChart();chart.add_data(Reference(m,min_col=14,min_row=5,max_row=9),titles_from_data=True);w.create_sheet('Диаграмма').add_chart(chart,'A1')
    w.save(path)
    raw=BASE/'migrate-raw.xlsx';path.rename(raw)
    with zipfile.ZipFile(raw) as zi, zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            data=zi.read(item.filename)
            if item.filename=='xl/workbook.xml': data=re.sub(r'<calcPr[^>]*/>','',data.decode()).replace('</workbook>','<extLst><ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}"/></extLst></workbook>').encode()
            zo.writestr(item,data)
    run(path,'2026-10-04');run(path,'2026-10-04')
    m=load_workbook(path)[MAIN]
    assert m['M5'].value=='Номер тома' and m['O5'].value=='несоотв шифр тома'
    assert [m.cell(6,c).value for c in range(12,25)]==[12,13,14,15,16,17,18,19,20,21,22,None,24], [m.cell(6,c).value for c in range(12,25)]
    assert m.auto_filter.filterColumn[0].colId==14 and m.freeze_panes=='O7' and m['O7'].comment is not None and m['N7'].comment is None
    with zipfile.ZipFile(path) as z:
        names=z.namelist();main=next(x for x in (z.read(n).decode() for n in names if n.startswith('xl/worksheets/sheet')) if '>Номер тома<' in x)
        charts=[z.read(n).decode() for n in names if n.startswith('xl/charts/chart')]
        workbook=z.read('xl/workbook.xml').decode()
        vml=[z.read(n).decode('latin-1') for n in names if n.endswith('.vml')]
    assert all('$O$6:$O$9' in c for c in charts) and charts, charts
    assert workbook.index('<calcPr')<workbook.index('<extLst'), workbook
    assert all(re.search(r'<(?:\w+:)?Column>14</',v) for v in vml) and vml, vml
    assert all(m.cell(r,13).value is None for r in (7,8,9)) and not re.search(r'<c r="M(?:[7-9]|[1-9]\d+)"[^>]*t="inlineStr"',main), [m.cell(r,13).value for r in (7,8,9)]

if __name__=='__main__':
    path=test_marks()
    test_user_columns()
    test_changed_date_format()
    test_invalid_marks_and_whitespace()
    test_rules()
    test_fragments()
    test_reviewer_names_stable()
    test_prefilled_reviewer_gets_date()
    test_missing_crc_finds_document()
    test_conflicts_shown_once()
    test_recalculation_keeps_register()
    test_new_rows_follow_data()
    test_damaged_source()
    test_kind_of_volume()
    test_check_dates()
    test_volume_column_references()
    print('PD recheck passed: resolved zero variants, boolean marks, preserved user columns, repeat merge, numeric format, invalid marks, stable names, dates, CRC, conflicts, recalculation, rows, kinds and column migration')
    print(path)
