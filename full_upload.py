"""表C → 表A 整表同步：首轮先整表格式，再整表纯数据。

第一轮（with_formats=True）先在修改表A之前预检表C格式，再按配置允许的最大行数
逐批选择 WPS“粘贴格式”。全部格式批次完成后，清空旧数据（保留格式），再从同一起点
逐批选择“粘贴值”，每批同时回读数据和格式；全部通过后才结束首轮格式待办。
空白区域的剪贴板可能不含真实格式，格式阶段只标记待复核，绝不据此宣告格式成功。
后续轮次（with_formats=False）仅清空旧内容并粘贴值，不携带任何 CSS 格式。

两阶段共用 upload.batch_cells（默认 80000 格），最大行数为格数上限除以实际列数
向下取整，最后一批取剩余行数。不再额外把格式批次限制成 8000 格。载荷超过 64 MiB
或校验不一致时才缩小范围；补贴失败则明确报错，下一轮仍执行两阶段。

范围由 source.scan_range 或 sync.start_cell 和文件实际边界确定。复制静态单元格格式，
不复制行高、列宽、合并和条件格式规则。公式仅取缓存结果，不复制公式表达式。
"""
import datetime as dt
import time
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.datetime import to_excel

from copy_range import column_name
from update_excel import address
from sync_engine import data_start_row, start_cell
from source_file import render_value

# 单次粘贴格数上限：默认沿用实测稳定的 8 万格（约 30 MB CF_HTML）；超过会压垮 WPS 标签页
# （此前约 40 万格即崩溃）。可通过 upload.batch_cells 调大以提速，但请勿超过约 20 万格。
WRITE_MAX_CELLS = 80_000


def _upload_batch_cells(config):
    """读取每次粘贴的格数上限；未配置或非法时回退到默认值。"""
    value = (config or {}).get('upload', {}).get('batch_cells') or WRITE_MAX_CELLS
    if isinstance(value, bool) or not isinstance(value, int) or value < 1:
        return WRITE_MAX_CELLS
    return value


def _source_path(root, config):
    value = (config.get('source', {}).get('format_file') or '').strip() or 'DATA/表C.xlsx'
    return (Path(root) / value).resolve()


def _sheet_name(config, workbook):
    name = (config.get('source', {}).get('format_sheet') or '').strip() \
        or (config.get('source', {}).get('sheet_name') or '').strip()
    return name if name in workbook.sheetnames else workbook.sheetnames[0]


def _excel_number_text(value):
    """把数值/日期序号转成 x:num 属性的精确字符串（整数不带小数点）。"""
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return repr(value)
    return str(value)


def _formula_cells(sheet):
    """流式扫描工作表 XML，返回公式单元格坐标集合 {(row, col)}。

    openpyxl 只读模式 data_only=True 时，公式格返回缓存值且 data_type='n'（不再标记 'f'），
    与带格式轮（data_only=False，公式按文本粘贴、不写 x:num）无法通过单元格属性区分。
    这里直接扫描工作表 XML 的 <f> 元素识别公式格，保证只数据轮与带格式轮对公式格的
    处理完全一致：都按缓存结果文本粘贴、不写 x:num。
    """
    from xml.etree import ElementTree as ET
    from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
    cells = set()
    source = sheet.parent._archive.open(sheet._worksheet_path)
    try:
        current = None
        for _event, elem in ET.iterparse(source, events=('start',)):
            tag = elem.tag.rsplit('}', 1)[-1]
            if tag == 'c':
                current = elem.get('r')
            elif tag == 'f' and current:
                try:
                    col_letter, row = coordinate_from_string(current)
                    col = column_index_from_string(col_letter)
                except ValueError:
                    pass
                else:
                    cells.add((row, col))
            elem.clear()
    finally:
        source.close()
    return cells


def _num_for_cell(cell, rendered):
    """返回该格用于 x:num 的精确序号字符串（数值/日期格），否则返回 None。

    规则与 HtmlCellFormats._number_attribute 完全一致：公式格按缓存结果文本粘贴、
    不写 x:num；文本格式（@）不写 x:num；布尔/文本/空不写 x:num；序号还必须能被
    比较逻辑视作与渲染文本等价（避免“渲染文本 vs 序号”在回读校验里被误判为差异）。
    """
    value = getattr(cell, 'value', None)
    if value is None or isinstance(value, (bool, str)):
        return None
    if (getattr(cell, 'number_format', None) or 'General') == '@':
        return None
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        number = to_excel(value)
    elif isinstance(value, (int, float)):
        number = value
    else:
        return None
    from upload_verify import values_equal
    if not values_equal(rendered, number):
        return None
    return _excel_number_text(number)


def read_table_c_whole(root, config, with_nums=False):
    """读取 DATA/表C.xlsx 的数据区域（自起始单元格起），返回与整表复制同构的网格。

    数值/日期渲染为显示文本；公式取缓存的计算结果（data_only=True，不含公式表达式）。
    起点取 source.scan_range 的左上角（若填写），否则取 sync.start_cell（默认 A5）；
    结束列/结束行取 source.scan_range 的右下角（若填写），否则取文件实际最大列/行。

    with_nums=True（只粘贴数据轮）时同时返回 nums：与 rows 同形状的 x:num 序号
    （数值/日期格）或 None，供剪贴板保持数值/日期/文本类型；公式格一律 None，
    与带格式轮一致（都按缓存结果文本粘贴）。
    """
    path = _source_path(root, config)
    if not path.is_file():
        raise ValueError('表C 数据文件不存在：' + str(path) + '；请先下载表C')
    _fallback_name, start_col, start_row = start_cell(config)
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook[_sheet_name(config, workbook)]
        last_col = sheet.max_column
        last_row = sheet.max_row
        # scan_range 是数据区域的权威声明：起点取其左上角，结束列/行收敛到其右下角，
        # 避免表C下载文件里的“幽灵列/多余行”把范围撑大。
        scan_range = (config.get('source', {}).get('scan_range') or '').strip()
        if scan_range and ':' in scan_range:
            try:
                scan_left, scan_top = address(scan_range.split(':', 1)[0])
                scan_right, scan_bottom = address(scan_range.split(':', 1)[1])
                start_col, start_row = scan_left, scan_top
                if scan_right < last_col:
                    last_col = scan_right
                if scan_bottom < last_row:
                    last_row = scan_bottom
            except ValueError:
                pass
        if last_row < start_row or last_col < start_col:
            raise ValueError('表C 数据文件没有数据：起始单元格之后没有数据行（' + str(path) + '）')
        formula_cells = _formula_cells(sheet) if with_nums else set()
        width = last_col - start_col + 1
        total_rows = last_row - start_row + 1
        rows = []
        nums = []
        for row in sheet.iter_rows(min_row=start_row, max_row=last_row,
                                   min_col=start_col, max_col=last_col,
                                   values_only=False):
            values = []
            numbers = []
            for cell in row:
                value = getattr(cell, 'value', None)
                rendered = render_value(value, getattr(cell, 'number_format', 'General'))
                values.append(rendered)
                if with_nums:
                    if (getattr(cell, 'row', 0), getattr(cell, 'column', 0)) in formula_cells:
                        numbers.append(None)
                    else:
                        numbers.append(_num_for_cell(cell, rendered))
            rows.append(values)
            if with_nums:
                nums.append(numbers)
        return dict(rows=rows, nums=(nums if with_nums else None), first_col=start_col,
                    first_row=start_row, last_col=last_col, last_row=last_row,
                    width=width, row_count=total_rows)
    finally:
        workbook.close()


def plan_batches(row_count, width, config=None, with_formats=False):
    """按行分批，每批 cells = rows × width 不超过单次粘贴上限；返回 [(start, end), ...]。"""
    cells = _upload_batch_cells(config)
    rows_per = max(1, cells // max(1, width))
    batches = []
    start = 1
    while start <= row_count:
        end = min(row_count, start + rows_per - 1)
        batches.append((start, end))
        start = end + 1
    return batches


def _styles_for_rows(formats, rows, start_row, width, data_start, first_col=1):
    """为所选绝对行区间构建样式矩阵，包括数据起始行之前的行。

    格式全部来自表C 自身（HtmlCellFormats 逐格读取表C单元格样式），不再使用任何
    单独的格式文件；数字/日期格式同样取自表C 该格自身。
    """
    styles = []
    joined_cache = {}
    for i, value_row in enumerate(rows):
        source_row = start_row + i
        css_rows = []
        for j in range(width):
            css = formats.style(source_row, first_col + j)
            # 同一 style_id 返回同一 dict 对象，按对象身份缓存 join 结果，避免逐格重复拼接。
            key = id(css)
            css_text = joined_cache.get(key)
            if css_text is None:
                css_text = ';'.join(f'{k}:{v}' for k, v in css.items())
                joined_cache[key] = css_text
            entry = {'_css': css_text}
            num = formats.number_attribute(source_row, first_col + j, value_row[j])
            if num is not None:
                entry['_num'] = num
            css_rows.append(entry)
        styles.append(css_rows)
    return styles


def _data_only_nums(nums, first_row, data_start):
    """保留所选范围全部行的数值类型，不以旧主键扫描起始行裁掉类型。"""
    if nums is None:
        return None
    return [list(row) for row in nums]


def _destination_ref(copier, page_a, config, grid):
    """整表上传的目标清空范围：自起始单元格起到 upload.scan_range 右下角（或表A已用范围末尾）。"""
    upload_scan = (config.get('upload', {}).get('scan_range') or '').strip()
    end_col, end_row = grid['last_col'], grid['last_row']
    if upload_scan and ':' in upload_scan:
        try:
            scan_right, scan_bottom = address(upload_scan.split(':', 1)[1])
            end_col, end_row = max(end_col, scan_right), max(end_row, scan_bottom)
        except ValueError:
            pass
    else:
        used = copier._used_range(page_a)
        try:
            _last_col, _last_row = address(used.split(':')[1])
            end_col, end_row = max(end_col, _last_col), max(end_row, _last_row)
        except ValueError:
            pass
    return f'{column_name(grid["first_col"])}{grid["first_row"]}:{column_name(end_col)}{end_row}'


def upload_whole_table(copier, config, with_formats=True):
    """首轮两阶段，后续仅粘贴值；首轮数据和格式全部验证后才能完成。"""
    from sync_runner import _open_sheet, _stop_requested, SyncStopped
    from full_upload_formats import paste_formatted_batch, paste_data_batch

    if _stop_requested(copier):
        raise SyncStopped()

    activity = getattr(copier, 'activity', None)
    grid = read_table_c_whole(copier.root, config, with_nums=True)
    width = grid['width']
    first_col = grid['first_col']
    first_row = grid['first_row']
    last_col = grid['last_col']
    last_row = grid['last_row']
    data_start = data_start_row(config)
    rows = grid['rows']
    row_count = grid['row_count']

    formats = None
    styles = None
    plain_nums = _data_only_nums(grid.get('nums'), first_row, data_start)
    if with_formats:
        # 带格式轮：格式全部来自表C，预检 + 一次性生成全部样式（在改动表A之前验证格式
        # 可表达，同时避免逐批重复计算）。预检失败直接中止本轮，不落进无格式粘贴。
        from html_cell_formats import HtmlCellFormats
        formats = HtmlCellFormats(copier, config, keyed=False)
    try:
        if formats is not None:
            copier.update_upload(message='正在预检表C单元格格式并生成样式（格式来源：DATA/表C.xlsx）…')
            styles = _styles_for_rows(formats, rows, first_row, width, data_start, first_col)
        else:
            copier.update_upload(message='正在预检表C数据（本轮只粘贴数据、不携带格式）…')

        page_a = _open_sheet(copier, config, 'a')
        batches = plan_batches(row_count, width, config, with_formats=with_formats)
        total = len(batches)
        max_cells = config['source']['max_cells']
        # Complete the entire format pass before clearing contents or writing
        # any data batches. Paste-format mode preserves the target's values.
        if styles is not None:
            for idx, (r1, r2) in enumerate(batches):
                if _stop_requested(copier):
                    raise SyncStopped()
                abs_r1, abs_r2 = first_row+r1-1, first_row+r2-1
                target_ref = f'{column_name(first_col)}{abs_r1}:{column_name(last_col)}{abs_r2}'
                copier.update_upload(message=f'阶段1/2：粘贴格式 {target_ref}（{idx+1}/{total}）…',
                                     scan_row=abs_r1, scan_end=last_row, batch_end=abs_r2)
                verified = paste_formatted_batch(copier, page_a, rows[r1-1:r2], styles[r1-1:r2],
                                                 first_col, abs_r1, max_cells, format_only=True)
                if activity is not None:
                    activity.record('格式阶段批次完成', target=target_ref, rows=r2-r1+1,
                                    verified=bool(verified), final_check='数据写入后逐格复核')

        # Delete removes old contents (including excess old rows), preserving
        # the formats just applied. No second clear occurs between data batches.
        if _stop_requested(copier):
            raise SyncStopped()
        clear_ref = _destination_ref(copier, page_a, config, grid)
        copier.update_upload(message=f'格式阶段完成，正在清空旧数据 {clear_ref} …' if with_formats
                             else f'正在清空表A旧数据 {clear_ref} …')
        copier._select(page_a, clear_ref)
        page_a.keyboard.press('Delete')
        page_a.wait_for_timeout(800)

        for idx, (r1, r2) in enumerate(batches):
            if _stop_requested(copier):
                raise SyncStopped()
            start = time.monotonic()
            abs_r1, abs_r2 = first_row+r1-1, first_row+r2-1
            target_ref = f'{column_name(first_col)}{abs_r1}:{column_name(last_col)}{abs_r2}'
            phase = '阶段2/2：' if with_formats else ''
            copier.update_upload(message=f'{phase}粘贴纯数据 {target_ref}（{idx+1}/{total}）…',
                                 scan_row=abs_r1, scan_end=last_row, batch_end=abs_r2)
            batch_rows = rows[r1-1:r2]
            paste_data_batch(copier, page_a, batch_rows, plain_nums[r1-1:r2], first_col,
                             abs_r1, max_cells, styles=styles[r1-1:r2] if styles is not None else None)
            copier.update_upload(message=f'已完成 {target_ref}，数据'+('和格式' if with_formats else '')+
                                 f'回读校验通过（耗时 {time.monotonic()-start:.1f}s）')
            if activity is not None:
                activity.record('写回表A完成', target=target_ref, rows=len(batch_rows),
                                columns=f'{column_name(first_col)}～{column_name(last_col)}',
                                cells=len(batch_rows)*width, with_formats=bool(with_formats),
                                phase='纯数据及最终校验')
    finally:
        if formats is not None:
            formats.close()

    mode_text = '先格式后纯数据，最终逐格校验数据和格式' if with_formats else '仅数据（粘贴值）'
    result = dict(uploaded=row_count, rows=row_count, columns=width,
                  batches=total*(2 if with_formats else 1), format_batches=total if with_formats else 0, data_batches=total, batch_rows=max(r2-r1+1 for r1, r2 in batches),
                  cells=row_count * width, with_formats=bool(with_formats),
                  message=(f'已把表C整表上传到表A（{column_name(first_col)}{first_row} 起，'
                           f'共 {row_count} 行 × {width} 列，{total*(2 if with_formats else 1)} 个计划批次，{mode_text}，逐批回读校验通过）。'))
    copier.update_upload(message=result['message'])
    return result
