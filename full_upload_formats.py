"""Bounded rich-text pastes with real, per-cell content and format readback."""
from copy_range import column_name
from clipboard_html import CopyIntegrityError
from html_cell_formats import FormatReadbackError, canonical, declarations, read_styles
from local_clipboard import clipboard_rows, clipboard_data_rows
from paste_menu import select_paste_mode
from upload_verify import snapshot, differences

MAX_HTML_BYTES = 64 * 1024 * 1024
MAX_REPAIR_DEPTH = 3


def verify_formats(clip, styles, first_col, first_row):
    if clip.get('blank'):
        raise FormatReadbackError('表A空白占位剪贴板不能证明格式已复制')
    actual = read_styles(clip)
    if len(actual) != len(styles) or any(len(a) != len(s) for a, s in zip(actual, styles)):
        raise FormatReadbackError('表A格式回读行列不完整')
    # Compare all transmitted fields, including borders and number formats.
    # Missing fields must not count as a successful formatted paste.
    for i, (wanted_row, got_row) in enumerate(zip(styles, actual)):
        for j, (entry, got) in enumerate(zip(wanted_row, got_row)):
            for field, wanted in declarations(entry['_css']).items():
                value = got.get(field)
                if value is None or canonical(field, value) != canonical(field, wanted):
                    cell = f'{column_name(first_col+j)}{first_row+i}'
                    raise FormatReadbackError(
                        f'表A {cell} 格式未复制完整：{field} '
                        f'期望 {wanted}，实际 {value if value is not None else "缺失"}')


def _verify(copier, page, ref, rows, styles, first_col, first_row, max_cells,
            *, format_only=False):
    from sync_runner import _stop_requested, SyncStopped
    expected = dict(rows=rows, row_count=len(rows), width=len(rows[0]))
    for attempt in range(3):
        if _stop_requested(copier):
            raise SyncStopped()
        try:
            clip = copier._capture(page, ref)
            if format_only and clip.get('blank'):
                # WPS supplies no genuine HTML for an empty target. This is
                # pending evidence, not a successful format verification. The
                # data phase must verify both contents and styles before success.
                return False
            actual = snapshot(clip, ref, max_cells, strict_rows=True)
            if not format_only and differences(expected, actual, ref.split(':')[0], limit=1):
                raise FormatReadbackError(f'表A {ref} 内容回读不一致（可能发生文本自动转换）')
            if styles is not None:
                verify_formats(clip, styles, first_col, first_row)
            return True
        except CopyIntegrityError as exc:
            # A clipboard read limit does not imply a paste limit. Keep the
            # original maximum-size paste and split only the verification reads.
            if not exc.can_split or len(rows) == 1:
                raise
            mid = len(rows) // 2
            verified = True
            for start, end in ((0, mid), (mid, len(rows))):
                subref = (f'{column_name(first_col)}{first_row+start}:'
                          f'{column_name(first_col+len(rows[0])-1)}{first_row+end-1}')
                result = _verify(copier, page, subref, rows[start:end],
                                 styles[start:end] if styles is not None else None,
                                 first_col, first_row+start, max_cells, format_only=format_only)
                verified = verified and result
            return verified
        except FormatReadbackError:
            if attempt == 2:
                raise
            # Let asynchronous formatting settle before considering a repair.
            page.wait_for_timeout(500 * (attempt+1))


def paste_formatted_batch(copier, page, rows, styles, first_col, first_row,
                          max_cells, depth=0, *, format_only=False):
    from sync_runner import _stop_requested, SyncStopped
    if _stop_requested(copier):
        raise SyncStopped()
    width = len(rows[0])
    ref = (f'{column_name(first_col)}{first_row}:'
           f'{column_name(first_col+width-1)}{first_row+len(rows)-1}')
    clip = clipboard_rows(rows, styles)
    oversized = len(clip['html']) > MAX_HTML_BYTES
    if oversized and len(rows) == 1:
        raise ValueError(f'表C {ref} 单行富文本超过安全大小，未执行本批粘贴')
    if not oversized:
        # Selecting/focusing the web sheet can alter its clipboard state. Stage
        # only after selection and check the OS clipboard immediately before paste.
        copier._select(page, f'{column_name(first_col)}{first_row}')
        copier._set_local_clipboard(rows, clip)
        copier._verify_local_clipboard(clip)
        copier._paste(page, ref, values=rows)
        if format_only:
            select_paste_mode(page, 'formats')
        page.wait_for_timeout(500)
        page.keyboard.press('Escape')
        try:
            return _verify(copier, page, ref, rows, styles, first_col, first_row,
                           max_cells, format_only=format_only)
        except FormatReadbackError as exc:
            if depth >= MAX_REPAIR_DEPTH or len(rows) == 1:
                raise FormatReadbackError(f'{exc}；缩小批次后仍不一致，本轮未完成，保留带格式重试状态') from exc
            copier.update_upload(message=f'{ref} 内容或格式未完整写入，正在缩小批次补贴…')
            activity = getattr(copier, 'activity', None)
            if activity is not None:
                activity.record('格式补贴', target=ref, reason=str(exc), depth=depth,
                                html_bytes=len(clip['html']))
    mid = len(rows) // 2
    del clip
    verified = True
    for start, end in ((0, mid), (mid, len(rows))):
        part_verified = paste_formatted_batch(copier, page, rows[start:end], styles[start:end],
                              first_col, first_row+start, max_cells,
                              depth if oversized else depth+1, format_only=format_only)
        verified = verified and part_verified
    return verified


def paste_data_batch(copier, page, rows, nums, first_col, first_row, max_cells,
                     *, styles=None, depth=0, repair_formats=False):
    """Paste values explicitly; first-round completion verifies both phases."""
    from sync_runner import _stop_requested, SyncStopped
    if _stop_requested(copier):
        raise SyncStopped()
    width = len(rows[0])
    ref = (f'{column_name(first_col)}{first_row}:'
           f'{column_name(first_col+width-1)}{first_row+len(rows)-1}')
    clip = clipboard_data_rows(rows, nums, preserve_formats=True)
    oversized = len(clip['html']) > MAX_HTML_BYTES
    if oversized and len(rows) == 1:
        raise ValueError(f'表C {ref} 单行数据超过安全大小，未执行本批粘贴')
    if not oversized:
        if repair_formats and styles is not None:
            paste_formatted_batch(copier, page, rows, styles, first_col, first_row,
                                  max_cells, depth=depth, format_only=True)
        copier._select(page, f'{column_name(first_col)}{first_row}')
        copier._set_local_clipboard(rows, clip)
        copier._verify_local_clipboard(clip)
        copier._paste(page, ref, values=rows)
        select_paste_mode(page, 'values')
        page.wait_for_timeout(500)
        page.keyboard.press('Escape')
        try:
            _verify(copier, page, ref, rows, styles, first_col, first_row, max_cells)
            return
        except FormatReadbackError as exc:
            if depth >= MAX_REPAIR_DEPTH or len(rows) == 1:
                raise FormatReadbackError(f'{exc}；两阶段校验未通过，保留首轮格式重试状态') from exc
            copier.update_upload(message=f'{ref} 数据或格式不一致，正在缩小范围补贴并重新校验…')
            activity = getattr(copier, 'activity', None)
            if activity is not None:
                activity.record('两阶段补贴', target=ref, reason=str(exc), depth=depth)
    mid = len(rows) // 2
    del clip
    for start, end in ((0, mid), (mid, len(rows))):
        paste_data_batch(copier, page, rows[start:end], nums[start:end] if nums else None,
                         first_col, first_row+start, max_cells,
                         styles=styles[start:end] if styles is not None else None,
                         depth=depth if oversized else depth+1,
                         repair_formats=repair_formats if oversized else styles is not None)
