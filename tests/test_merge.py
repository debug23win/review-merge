"""Regression checks against two artificial XLSX review copies."""
import json
import hashlib
import shutil
from openpyxl import load_workbook
from copy import copy
from datetime import datetime
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
        assert not any(addr.startswith(('U','V','W','X','Y','Z','AA')) for addr,(v,f) in main.items() if v or f)
        assert any('COUNTIFS' in f for member in z.namelist() if member.startswith('xl/worksheets/') and member.endswith('.xml') for v,f in cells(z,member).values() if f)
        assert not any('[Анна,' in v or '[Борис,' in v for v in notes)
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

def test_user_columns():
    temp=run_root/uuid.uuid4().hex;temp.mkdir();target=temp/'occupied.xlsx'
    shutil.copyfile(ROOT/'tests/fixtures/Анна.xlsx',target)
    wb=load_workbook(target);wb['Все загруженные файлы']['U7']='Сохранить';wb.save(target)
    before=target.read_bytes()
    args=[str(ROOT/'Свод_проверки.exe'),'--batch',str(target),'2026-10-04','Все загруженные файлы',str(ROOT/'tests/fixtures/Борис.xlsx')]
    assert subprocess.run(args,timeout=90).returncode==0
    assert load_workbook(target)['Все загруженные файлы']['U7'].value=='Сохранить'
    assert list(temp.glob('*.backup-*'))[0].read_bytes()==before

def test_template_and_fragments():
    temp=run_root/uuid.uuid4().hex;temp.mkdir();source=temp/'template.xlsx'
    shutil.copyfile(ROOT/'tests/fixtures/Анна.xlsx',source)
    wb=load_workbook(source);main=wb['Все загруженные файлы']
    main['C7']='Любой шифр A/7 Фрагмент 1';main['C8']='Любой шифр A/7 Фрагмент 2'
    for r in [7,8]:
        main.cell(r,9,'Проверяющий');main.cell(r,10,datetime(2026,10,2));main.cell(r,13,1)
    main['T7']='Первое исходное замечание';main['T8']='Второе исходное замечание'
    main['S7']='1. Ошибка в наименовании\n2. Ошибка в номере изменения';main['S8']=main['S7'].value
    main['U7']='Пользовательские данные U';main['AA7']='Пользовательские данные AA'
    for sheet in list(wb):
        if sheet.title.lower()=='свод':wb.remove(sheet)
    ref=wb.create_sheet('СВОД');ref.merge_cells('A1:D1');ref['A1']='СВОДНАЯ ТАБЛИЦА';ref['A2']='Мой объект';ref['A7']='Получено коробов';ref['A8']='Кол-во томов по акту'
    physical=wb.create_sheet('Справка по томам');physical['A1']='Коробов получено';physical['B1']=50;physical['B2']=800;physical['B8']=datetime(2026,10,2);physical['G42']=5;physical['G43']=120
    for c,date in [(2,datetime(2026,10,2)),(3,datetime(2026,10,4)),(4,datetime(2026,10,5))]:
        ref.cell(5,c,date);ref.cell(5,c).number_format='dd.mm.yyyy'
        ref.cell(7,c,"='Справка по томам'!B1");ref.cell(8,c,"='Справка по томам'!B2")
    ref['B9']="='Справка по томам'!G42";ref['B10']="='Справка по томам'!G43"
    ref['A17']='Сохранить исходное примечание';ref['A65']='Пользовательский текст ниже статистики'
    ref['A7'].font=copy(main['C7'].font);style=copy(ref['A7']._style)
    wb.save(source);output=temp/'result.xlsx';run(output,[source])
    w=load_workbook(output);s=w['Свод'];m=w['Все загруженные файлы']
    assert s['A1'].value=='СВОДНАЯ ТАБЛИЦА' and s['A2'].value=='Мой объект'
    assert s['A7']._style==style and s['A17'].value=='Сохранить исходное примечание'
    assert s['A65'].value=='Пользовательский текст ниже статистики'
    assert s['B10'].value=="='Справка по томам'!G43"
    assert m['U7'].value=='Пользовательские данные U' and m['AA7'].value=='Пользовательские данные AA'
    assert m['T7'].value=='Первое исходное замечание' and m['T8'].value=='Второе исходное замечание'
    marker=next(c for c in s[1] if c.value=='ReviewMerge.FirstDates.v2');vh=marker.column+12
    matching=[r for r in range(7,s.max_row+1) if s.cell(r,vh).value=='ЛЮБОЙ ШИФР A/7']
    assert len(matching)==1,matching
    vc=s.cell(matching[0],vh+3);assert vc.value.startswith('=MAX(')
    cached=load_workbook(output,data_only=True)['Свод'];assert cached.cell(matching[0],vh+9).value==2
    assert cached.cell(matching[0],vh+10).value==2
    before_notes=m['T7'].value,m['T8'].value;run(output,[source]);s2=load_workbook(output)['Свод'];m2=load_workbook(output)['Все загруженные файлы']
    assert before_notes==(m2['T7'].value,m2['T8'].value)
    assert s2['B10'].value=="='Справка по томам'!G43"

def test_legacy_annotations():
    temp=run_root/uuid.uuid4().hex;temp.mkdir();source=temp/'legacy.xlsx';shutil.copyfile(ROOT/'tests/fixtures/Анна.xlsx',source)
    wb=load_workbook(source);ws=wb['Все загруженные файлы']
    headers=['Ключ короба (формула)','Первая строка короба','Первая проверка короба','Документ проверен','Дата для статистики','Есть замечания','Первая дата из источников']
    for c,text in enumerate(headers,21):ws.cell(5,c,text)
    ws['AA7']=datetime(2026,10,1);ws['T7']='[Анна, 01.10.2026] Первый текст\n\n[Борис, 02.10.2026] Второй текст'
    wb.save(source);output=temp/'clean.xlsx';run(output,[source])
    ws=load_workbook(output)['Все загруженные файлы']
    assert ws['T7'].value=='Первый текст\n\nВторой текст'
    assert all(ws.cell(5,c).value is None and ws.cell(7,c).value is None for c in range(21,28))
    run(output,[source]);assert load_workbook(output)['Все загруженные файлы']['T7'].value=='Первый текст\n\nВторой текст'

test_merge()
test_existing_target()
test_user_columns()
test_template_and_fragments()
test_legacy_annotations()
print('Merge tests passed: union, flags, comments, conflicts, invalid data, formulas and repeated merge')
