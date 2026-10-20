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

if __name__=='__main__':
    path=test_marks()
    test_user_columns()
    test_changed_date_format()
    test_invalid_marks_and_whitespace()
    test_rules()
    test_fragments()
    print('PD recheck passed: resolved zero variants, boolean marks, preserved user columns, repeat merge, numeric format and invalid marks')
    print(path)
