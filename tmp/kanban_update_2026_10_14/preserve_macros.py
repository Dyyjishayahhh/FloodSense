from pathlib import Path
import copy
import hashlib
import json
import os
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET

work = Path(__file__).resolve().parent
manifest = json.loads((work / 'manifest.json').read_text(encoding='utf-8'))
target = Path(manifest['target'])
backup = work / 'FloodSense_Kanban_Board.before.xlsm'
shutil.copy2(target, backup)
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ET.register_namespace('', NS)
q = lambda name: '{' + NS + '}' + name

with zipfile.ZipFile(backup) as original, zipfile.ZipFile(work / 'staging.xlsx') as staged:
    original_sheet = original.read('xl/worksheets/sheet1.xml').decode('utf-8')
    original_styles = original.read('xl/styles.xml').decode('utf-8')
    old_sheet = ET.fromstring(original_sheet)
    old_styles = ET.fromstring(original_styles)
    new_sheet = ET.fromstring(staged.read('xl/worksheets/sheet1.xml'))
    new_styles = ET.fromstring(staged.read('xl/styles.xml'))
    # The artifact tool handles the cells and layout. Native package parts,
    # including VBA, remain from the original macro-enabled workbook.
    for group in ['fonts', 'fills', 'borders', 'cellStyleXfs']:
        old_group, new_group = old_styles.find(q(group)), new_styles.find(q(group))
        assert len(old_group) == len(new_group), group
        assert [ET.tostring(item) for item in old_group] == [ET.tostring(item) for item in new_group], group
    old_xfs = old_styles.find(q('cellXfs'))
    new_xfs = new_styles.find(q('cellXfs'))
    assert ET.tostring(old_xfs[0]) == ET.tostring(new_xfs[0])
    offset = len(old_xfs) - 1
    for style in list(new_xfs)[1:]:
        old_xfs.append(copy.deepcopy(style))
    old_xfs.set('count', str(len(old_xfs)))
    styles_xml = re.sub(r'<cellXfs\b[^>]*>.*?</cellXfs>', ET.tostring(old_xfs, encoding='unicode'), original_styles, flags=re.S)
    shared = []
    if 'xl/sharedStrings.xml' in staged.namelist():
        shared = [''.join(node.itertext()) for node in ET.fromstring(staged.read('xl/sharedStrings.xml'))]
    data = new_sheet.find(q('sheetData'))
    for cell in data.iter(q('c')):
        style = int(cell.get('s', '0'))
        if style:
            cell.set('s', str(style + offset))
        cell_type = cell.get('t')
        if cell_type in ('s', 'str'):
            value = cell.find(q('v'))
            content = shared[int(value.text)] if cell_type == 's' else value.text or ''
            cell.remove(value)
            cell.set('t', 'inlineStr')
            text = ET.SubElement(ET.SubElement(cell, q('is')), q('t'))
            text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            text.text = content
    data.remove(data[0])
    data.insert(0, copy.deepcopy(old_sheet.find(q('sheetData'))[0]))
    sheet_xml = re.sub(r'<sheetData\b[^>]*>.*?</sheetData>', ET.tostring(data, encoding='unicode'), original_sheet, flags=re.S)
    sheet_xml = re.sub(r'<cols\b[^>]*>.*?</cols>', ET.tostring(new_sheet.find(q('cols')), encoding='unicode'), sheet_xml, flags=re.S)
    sheet_xml = re.sub(r'<dimension\b[^>]*/>', f'<dimension ref="A1:D{manifest["noteRow"]}"/>', sheet_xml)
    merges = new_sheet.find(q('mergeCells'))
    if merges is not None:
        sheet_xml = sheet_xml.replace('</sheetData>', '</sheetData>' + ET.tostring(merges, encoding='unicode'), 1)
    candidate = work / 'FloodSense_Kanban_Board.updated.xlsm'
    changed = {'xl/worksheets/sheet1.xml': sheet_xml.encode('utf-8'), 'xl/styles.xml': styles_xml.encode('utf-8')}
    with zipfile.ZipFile(candidate, 'w') as result:
        for entry in original.infolist():
            result.writestr(copy.copy(entry), changed.get(entry.filename, original.read(entry.filename)))
    with zipfile.ZipFile(candidate) as result:
        assert result.testzip() is None
        assert set(result.namelist()) == set(original.namelist())
        for name in original.namelist():
            if name not in changed:
                assert result.read(name) == original.read(name), name
        saved = ET.fromstring(result.read('xl/worksheets/sheet1.xml'))
        cells = {cell.get('r'): cell for cell in saved.iter(q('c'))}
        for item in manifest['written']:
            value = ''.join(cells[item['cell']].find(q('is')).itertext())
            assert value == item['value'], item['cell']
        assert saved.find(q('sheetPr')).get('codeName') == 'Sheet1'
        assert saved.find(q('mergeCells')) is not None
        assert ET.fromstring(result.read('xl/styles.xml')).find(q('cellXfs')).get('count') == '4'
        print(json.dumps({'entries_verified': len(manifest['written']), 'unchanged_package_parts': len(original.namelist()) - len(changed), 'macros_preserved': result.read('xl/vbaProject.bin') == original.read('xl/vbaProject.bin')}))
    os.replace(candidate, target)
    print('Saved the updated original workbook.')
