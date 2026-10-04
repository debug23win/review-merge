"""Regression checks against two artificial XLSX review copies."""
import json
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
        assert all(n in names for n in ['СВОД', 'История проверок', 'Конфликты свода', 'Проблемы данных', 'Служебный лист'])
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
test_merge()
print('Merge tests passed: union, flags, comments, conflicts, invalid data, formulas and repeated merge')
