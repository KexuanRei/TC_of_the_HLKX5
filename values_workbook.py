"""Freeze XLSX formula caches without reserializing worksheet formatting.

Worksheet XML is scanned with Expat, then only the byte spans occupied by cell
formulas are removed. All other cell bytes and every unrelated ZIP member are
retained, including WPS extensions that spreadsheet libraries may discard.
"""
from dataclasses import dataclass
import hashlib
import os
from pathlib import Path
import re
import tempfile
from xml.parsers import expat
from zipfile import ZipFile

from lxml import etree

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS = MAIN + '}'
TAG = re.compile(br'<(?:[^>\"\']|\"[^\"]*\"|\'[^\']*\')*>')
CHUNK = 1024 * 1024


@dataclass
class SheetScan:
    cells: int
    formulas: int
    errors: int
    spans: list


def _check(stop):
    if stop():
        from sync_runner import SyncStopped
        raise SyncStopped('已停止文件转换')


def scan_sheet(stream, name, stop=lambda: False):
    parser = expat.ParserCreate(namespace_separator='}')
    cells = formulas = errors = 0
    current = None
    spans = []
    active_formula = None
    in_value = False

    def start(tag, attrs):
        nonlocal cells, current, active_formula, in_value
        if tag == NS + 'c':
            cells += 1
            current = {'ref': attrs.get('r', '?'), 'type': attrs.get('t', 'n'),
                       'formula': False, 'value': False, 'text': False}
        elif current is not None and tag == NS + 'f':
            current['formula'] = True
            context = parser.GetInputContext()
            opening = TAG.match(context)
            if opening is None:
                raise ValueError(f'{name}: 无法定位公式 XML')
            active_formula = (parser.CurrentByteIndex, opening.group().endswith(b'/>'))
        elif current is not None and tag == NS + 'v':
            current['value'] = True
            in_value = True

    def text(value):
        if in_value and current is not None and value:
            current['text'] = True

    def end(tag):
        nonlocal current, active_formula, in_value, formulas, errors
        if tag == NS + 'f' and active_formula is not None:
            start_byte, self_closing = active_formula
            end_byte = parser.CurrentByteIndex
            if not self_closing:
                closing = TAG.match(parser.GetInputContext())
                if closing is None:
                    raise ValueError(f'{name}: 无法定位公式结束位置')
                end_byte += len(closing.group())
            spans.append((start_byte, end_byte))
            active_formula = None
        elif tag == NS + 'v':
            in_value = False
        elif tag == NS + 'c' and current is not None:
            if current['formula']:
                formulas += 1
                if not current['value'] or (not current['text'] and current['type'] != 'str'):
                    raise ValueError(f'{name}!{current["ref"]} 公式缺少计算结果；请在WPS完成重算并保存后重新下载。未生成上传文件。')
                if current['type'] == 'e':
                    errors += 1
            current = None

    def reject_dtd(*args):
        raise ValueError('不支持包含 DTD 的工作簿')

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.CharacterDataHandler = text
    parser.StartDoctypeDeclHandler = reject_dtd
    while True:
        _check(stop)
        data = stream.read(CHUNK)
        parser.Parse(data, not data)
        if not data:
            break
    return SheetScan(cells, formulas, errors, spans)


def _copy_without(source, target, spans, stop):
    position = 0
    for start, end in [*spans, (None, None)]:
        while start is None or position < start:
            _check(stop)
            data = source.read(CHUNK if start is None else min(CHUNK, start - position))
            if not data:
                if start is not None:
                    raise ValueError('工作表提前结束')
                return
            target.write(data)
            position += len(data)
        remaining = end - start
        while remaining:
            data = source.read(min(CHUNK, remaining))
            if not data:
                raise ValueError('公式范围不完整')
            remaining -= len(data)
            position += len(data)


def file_hash(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for data in iter(lambda: stream.read(CHUNK), b''):
            digest.update(data)
    return digest.hexdigest()


def _xml(data):
    return etree.fromstring(data, etree.XMLParser(resolve_entities=False, no_network=True))


def freeze_workbook(source, destination, stop=lambda: False, progress=lambda text: None):
    source, destination = Path(source).resolve(), Path(destination).resolve()
    if source == destination:
        raise ValueError('纯值输出不能覆盖原始表C')
    before = file_hash(source)
    destination.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary = tempfile.mkstemp(prefix='.values-', suffix='.xlsx', dir=destination.parent)
    os.close(handle)
    total = {'cells': 0, 'formulas': 0, 'formula_errors': 0, 'sheets': [], 'source_sha256': before}
    try:
        with ZipFile(source) as zin, ZipFile(temporary, 'w') as zout:
            if len(zin.namelist()) != len(set(zin.namelist())):
                raise ValueError('工作簿存在重复的 ZIP 成员')
            book = _xml(zin.read('xl/workbook.xml'))
            total['sheets'] = [el.get('name') for el in book.findall('{'+MAIN+'}sheets/{'+MAIN+'}sheet')]
            if not total['sheets']:
                raise ValueError('工作簿没有工作表')
            for info in zin.infolist():
                _check(stop)
                name = info.filename
                if name == 'xl/calcChain.xml':
                    continue
                if re.fullmatch(r'xl/worksheets/[^/]+\.xml', name):
                    progress('检查公式缓存：' + name)
                    with zin.open(info) as stream:
                        scan = scan_sheet(stream, name, stop)
                    total['cells'] += scan.cells
                    total['formulas'] += scan.formulas
                    total['formula_errors'] += scan.errors
                    with zin.open(info) as src, zout.open(info, 'w', force_zip64=True) as dst:
                        _copy_without(src, dst, scan.spans, stop)
                elif name in ('[Content_Types].xml', 'xl/_rels/workbook.xml.rels') or re.fullmatch(r'xl/tables/[^/]+\.xml', name):
                    original = zin.read(info)
                    root = _xml(original)
                    removed = False
                    for el in list(root.iter()):
                        local = etree.QName(el).localname
                        chain = (el.get('PartName') == '/xl/calcChain.xml' or
                                 el.get('Type', '').endswith('/calcChain'))
                        if chain or local in ('calculatedColumnFormula', 'totalsRowFormula'):
                            el.getparent().remove(el)
                            removed = True
                    zout.writestr(info, etree.tostring(root, xml_declaration=True, encoding='UTF-8') if removed else original)
                else:
                    with zin.open(info) as src, zout.open(info, 'w', force_zip64=True) as dst:
                        _copy_without(src, dst, [], stop)
        if file_hash(source) != before:
            raise ValueError('转换期间表C发生变化，本次输出已取消')
        progress('核验纯值文件：公式已移除，其他工作表字节保持不变')
        # A second parse verifies that the byte edits produced valid XML and no
        # cell formula survived (including shared/array and self-closing forms).
        with ZipFile(temporary) as checked:
            for name in checked.namelist():
                if re.fullmatch(r'xl/worksheets/[^/]+\.xml', name):
                    with checked.open(name) as stream:
                        if scan_sheet(stream, name, stop).formulas:
                            raise ValueError('输出仍含有单元格公式')
        _check(stop)
        os.replace(temporary, destination)
        total.update(path=str(destination), sha256=file_hash(destination))
        return total
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


