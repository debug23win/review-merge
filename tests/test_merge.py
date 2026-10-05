"""Regression checks against two artificial XLSX review copies."""
import json
import hashlib
import shutil
from openpyxl import load_workbook
import subprocess
import uuid
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def cells(archive, member):
    doc = ET.fromstring(archive.read(member))
    result = {}
    for cell in doc.findall('.//s:sheetData/s:row/s:c', NS):
        value = ''.join(t.text or '' for t in cell.findall('.//s:t', NS)) if cell.get('t') == 'inlineStr' else cell.findtext('s:v', '', NS)
        result[cell.get('r')] = (value, cell.findtext('s:f', None, NS))
    return result

def run(output, sources):
    subprocess.run([str(ROOT/'Свод_проверки.exe'), '--batch', str(output), '2026-10-04', 'Все загруженные файлы', *map(str, sources)], check=True, timeout=90)
    return json.loads(Path(str(output)+'.run.json').read_text(encoding='utf-8-sig'))

run_root = ROOT/'tests/.runs'
run_root.mkdir(exist_ok=True)
def test_merge():
    temp = run_root/uuid.uuid4().hex
    temp.mkdir()
    result = temp/'result.xlsx'
    report = run(result, [ROOT/'tests/fixtures/Анна.xlsx', ROOT/'tests/fixtures/Борис.xlsx'])
    assert (report['Documents'], report['Reviewed'], report['WithIssues'], report['Boxes']) == (7, 6, 4, 3), report
    assert report['Conflicts'] == 11 and report['Problems'] == 3, report
    with zipfile.ZipFile(result) as z:
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        names = [s.get('name') for s in wb.findall('s:sheets/s:sheet', NS)]
        assert set(names) == {'Все загруженные файлы', 'Свод', 'Служебный лист'}, names
        main = cells(z, 'xl/worksheets/sheet1.xml')
        notes = [v for addr, (v, f) in main.items() if addr.startswith(('S','T'))]
        assert any('=1+1' in v for v in notes)
        assert all(f is None for addr, (v, f) in main.items() if addr.startswith(('S','T')))
        assert sum(f is not None for v, f in main.values()) >= 7*6
        for addr in ['M7','N7','Q9']:
            assert main[addr][0] == '1', (addr, main[addr])
    repeated = temp/'repeated.xlsx'
    rerun = run(repeated, [result, ROOT/'tests/fixtures/Анна.xlsx', ROOT/'tests/fixtures/Борис.xlsx'])
    assert rerun['Documents'] == 7 and rerun['WithIssues'] == 4
    with zipfile.ZipFile(result) as a, zipfile.ZipFile(repeated) as b:
        first = cells(a, 'xl/worksheets/sheet1.xml')
        second = cells(b, 'xl/worksheets/sheet1.xml')
        for addr in ['S7','T7','AA7']:
            assert first.get(addr) == second.get(addr), (addr, first.get(addr), second.get(addr))
def test_existing_target():
    temp=run_root/uuid.uuid4().hex;temp.mkdir()
    target=temp/'existing.xlsx';shutil.copyfile(ROOT/'tests/fixtures/Анна.xlsx',target)
    wb=load_workbook(target);ws=wb['Все загруженные файлы']
    ws['A7']='=ROW()-5';ws['E7']='=100+23';ws['AB7']='Мои данные'
    ws['I100']='Строка без документа';ws['T100']='Сохранить это замечание'
    wb.create_sheet('Ручные итоги')['A1']='=2+3'
    wb.create_sheet('История проверок')['A1']='Личная история — сохранить'
    wb.save(target)
    before=target.read_bytes()
    source=ROOT/'tests/fixtures/Борис.xlsx';source_hash=hashlib.sha256(source.read_bytes()).hexdigest()
    report=run(target,[source,target,source])
    assert report['Documents']==7 and report['WithIssues']==4,report
    backups=list(temp.glob('existing.xlsx.backup-*'));assert len(backups)==1
    assert backups[0].read_bytes()==before
    assert hashlib.sha256(source.read_bytes()).hexdigest()==source_hash
    wb=load_workbook(target);ws=wb['Все загруженные файлы']
    assert ws['A7'].value=='=ROW()-5' and ws['E7'].value=='=100+23'
    assert ws['AB7'].value=='Мои данные' and ws['I100'].value=='Строка без документа'
    assert ws['T100'].value=='Сохранить это замечание'
    assert ws['M7'].value==1 and ws['N7'].value==1
    assert wb['Ручные итоги']['A1'].value=='=2+3'
    assert wb['История проверок']['A1'].value=='Личная история — сохранить'
    assert 'Свод' in wb.sheetnames and 'Конфликты свода' not in wb.sheetnames
    assert sum(bool(ws.cell(r,3).value) for r in range(101,ws.max_row+1))==2
    notes=ws['S7'].value,ws['T7'].value
    run(target,[source]);again=load_workbook(target)['Все загруженные файлы']
    assert notes==(again['S7'].value,again['T7'].value)
    assert again['A7'].value=='=ROW()-5'

def test_occupied_helpers():
    temp=run_root/uuid.uuid4().hex;temp.mkdir();target=temp/'occupied.xlsx'
    shutil.copyfile(ROOT/'tests/fixtures/Анна.xlsx',target)
    wb=load_workbook(target);wb['Все загруженные файлы']['U7']='Сохранить';wb.save(target)
    before=target.read_bytes()
    args=[str(ROOT/'Свод_проверки.exe'),'--batch',str(target),'2026-10-04','Все загруженные файлы',str(ROOT/'tests/fixtures/Борис.xlsx')]
    assert subprocess.run(args,timeout=90).returncode!=0
    assert target.read_bytes()==before and not list(temp.glob('*.backup-*'))

test_merge()
test_existing_target()
test_occupied_helpers()
print('Merge tests passed: union, flags, comments, conflicts, invalid data, formulas and repeated merge')
